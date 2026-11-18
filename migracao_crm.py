"""Migrador em lote de documentos entre CRM de Origem (AdvBox) e CRM de Destino (Juridiq)."""
import argparse
import logging
import os
import re
import shutil
import unicodedata
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

from coleta_nuvem import coletar, esperar_lista, limpa
from destino_juridiq import enviar_cliente
from planilha_cadastro import ABAS_PADRAO, PLANILHA_PADRAO, atualizar_etapa_cliente, garantir_coluna_etapa

ORIGEM_URL_PADRAO = "https://app.advbox.com.br/"
DESTINO_URL_PADRAO = "https://dashboard.juridiq.com.br/"
RELATORIO_SAIDA = Path("status_migracao.xlsx")
PASTA_DOWNLOADS = Path("./downloads")


def configurar_logger() -> logging.Logger:
    """Configura logging simultâneo no console e em migracao_crm.log."""
    logger = logging.getLogger("migracao_crm")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
        fh = logging.FileHandler("migracao_crm.log", encoding="utf-8")
        fh.setFormatter(fmt)
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        logger.addHandler(fh)
        logger.addHandler(sh)
    return logger


def resolver_url(env_key: str, padrao: str) -> str:
    val = (os.getenv(env_key) or "").strip()
    if not val or "exemplo.com" in val:
        return padrao
    return val


def normalizar_texto(texto: str) -> str:
    """Remove acentos, normaliza espaços e converte para minúsculas."""
    sem_acentos = "".join(
        ch for ch in unicodedata.normalize("NFKD", texto or "") if not unicodedata.combining(ch)
    )
    return re.sub(r"\s+", " ", sem_acentos).strip().casefold()


def selecionar_resultado_homonimos(
    itens_locator,
    nome_cliente: str,
    extrair_texto_fn,
    verificar_chips_advbox: bool = False,
) -> tuple[str, int, str]:
    """Aplica as 3 regras de verificação autônoma de homônimos usando .count().

    Retorna (status, total_encontrados, observacao):
    - "OK" quando um único item válido foi selecionado e clicado;
    - "Não encontrado" quando 0 itens foram retornados;
    - "Ambiguidade - Revisão Manual" quando 2+ resultados persistem sem correspondência única exata.
    """
    total = itens_locator.count()
    if total == 0:
        return "Não encontrado", 0, "0 resultados na busca"

    if total == 1:
        itens_locator.first.click()
        return "OK", 1, ""

    alvo_norm = normalizar_texto(nome_cliente)
    padrao_bruto = re.compile(rf"^\s*{re.escape(nome_cliente.strip())}\s*$", re.I)
    padrao_norm = re.compile(rf"^\s*{re.escape(alvo_norm)}\s*$", re.I)
    palavras_alvo = set(alvo_norm.split())

    correspondencias: list[int] = []
    for idx in range(total):
        item = itens_locator.nth(idx)
        texto_item = extrair_texto_fn(item).strip()
        texto_norm = normalizar_texto(texto_item)
        bate = bool(padrao_bruto.match(texto_item) or padrao_norm.match(texto_norm))
        if bate and verificar_chips_advbox:
            chips = [normalizar_texto(c) for c in item.locator("p.title span.mark").all_inner_texts()]
            if chips and not palavras_alvo.issubset(set(" ".join(chips).split())):
                bate = False
        if bate:
            correspondencias.append(idx)

    if len(correspondencias) == 1:
        itens_locator.nth(correspondencias[0]).click()
        return "OK", total, f"Desempatado por correspondência exata entre {total} resultados"

    return (
        "Ambiguidade - Revisão Manual",
        total,
        f"{total} resultados retornados ({len(correspondencias)} correspondência(s) exata(s))",
    )


def carregar_clientes(
    caminho_planilha: Path,
    abas: tuple[str, ...],
    reprocessar_concluidos: bool = False,
) -> tuple[list[dict], bool]:
    """Lê os clientes via pandas mantendo aba, linha no Excel e status de Etapa."""
    eh_cadastro = caminho_planilha.name.lower() == PLANILHA_PADRAO.name.lower()
    if eh_cadastro:
        garantir_coluna_etapa(caminho_planilha, abas=abas)
        xl = pd.ExcelFile(caminho_planilha, engine="openpyxl")
        abas_alvo = [a for a in abas if a in xl.sheet_names]
    else:
        xl = pd.ExcelFile(caminho_planilha, engine="openpyxl")
        abas_alvo = list(xl.sheet_names[:1])

    clientes: list[dict] = []
    for aba in abas_alvo:
        df = pd.read_excel(caminho_planilha, sheet_name=aba, engine="openpyxl")
        col_nome = "Nome do cliente" if "Nome do cliente" in df.columns else "nome_cliente"
        if col_nome not in df.columns:
            raise ValueError(f"Coluna de nome do cliente não encontrada na aba {aba!r}: {list(df.columns)}")

        for idx, row in df.iterrows():
            bruto = row.get(col_nome)
            if pd.isna(bruto):
                continue
            nome = str(bruto).strip()
            if not nome:
                continue
            etapa_atual = ""
            if "Etapa" in df.columns and not pd.isna(row.get("Etapa")):
                etapa_atual = str(row.get("Etapa")).strip().lower()
            if etapa_atual == "concluido" and not reprocessar_concluidos:
                continue
            clientes.append(
                {
                    "aba": aba,
                    "linha_excel": int(idx) + 2,
                    "nome_cliente": nome,
                    "etapa_atual": etapa_atual,
                }
            )
    return clientes, eh_cadastro


def obter_ou_abrir_pagina(ctx, trecho_url: str, url_inicial: str):
    """Reaproveita a aba existente que contém trecho_url ou abre uma nova."""
    for pg in ctx.pages:
        if trecho_url in pg.url:
            return pg
    pg = ctx.new_page()
    pg.goto(url_inicial)
    pg.wait_for_load_state("domcontentloaded")
    return pg


def aguardar_resultados_advbox(pg, timeout_s: int = 15) -> None:
    """Aguarda o dropdown do AdvBox concluir o carregamento da busca."""
    pg.locator(".form-search-results").first.wait_for(state="visible", timeout=timeout_s * 1000)
    for _ in range(timeout_s * 4):
        carregando = pg.locator(".form-search-results li.loading:visible").count()
        tem_pessoas = pg.locator(".form-search-results li.customers").count()
        sem_linhas = pg.locator(".form-search-results li.no-rows:visible").count()
        tem_processos = pg.locator(".form-search-results li.lawsuits").count()
        if carregando == 0 and (tem_pessoas > 0 or sem_linhas > 0 or tem_processos > 0):
            pg.wait_for_timeout(500)
            return
        pg.wait_for_timeout(250)


def buscar_e_baixar_origem(pg, ctx, url_origem: str, nome_cliente: str, pasta_cliente: Path, log) -> dict:
    """Executa busca com regra de homônimos e baixa arquivos recursivamente na Origem (AdvBox)."""
    if pasta_cliente.exists():
        shutil.rmtree(pasta_cliente)
    pasta_cliente.mkdir(parents=True, exist_ok=True)

    pg.bring_to_front()
    fechar_antigo = pg.locator("a.button-close-popup:visible")
    if fechar_antigo.count():
        fechar_antigo.first.click()
        pg.wait_for_timeout(400)

    pg.goto(url_origem)
    pg.wait_for_load_state("networkidle")

    campo = pg.locator('input[placeholder*="pesquisar" i]').first
    campo.click()
    campo.fill("")
    campo.fill(nome_cliente)
    aguardar_resultados_advbox(pg)

    itens_pessoas = pg.locator(".form-search-results li.customers a.row")
    st, total, obs = selecionar_resultado_homonimos(
        itens_pessoas,
        nome_cliente,
        lambda el: el.locator("p.title").inner_text(),
        verificar_chips_advbox=True,
    )
    if st != "OK":
        return {"ok": False, "status": st, "arquivos": 0, "observacao": f"Origem: {obs}"}

    pg.locator(".popup-content").first.wait_for(timeout=15000)
    pg.wait_for_timeout(1200)
    pg.locator("a.openFolder").first.click()
    esperar_lista(pg)

    listados, baixados, erros = coletar(pg, ctx, pasta_cliente, log=log)

    fechar = pg.locator("a.button-close-popup:visible")
    if fechar.count():
        fechar.first.click()
        pg.wait_for_timeout(400)

    if listados == 0:
        return {
            "ok": False,
            "status": "Não encontrado",
            "arquivos": 0,
            "observacao": "Origem: nenhum documento na nuvem do cliente",
        }
    if baixados != listados or erros:
        return {
            "ok": False,
            "status": "Erro no download",
            "arquivos": baixados,
            "observacao": f"Origem: baixados {baixados}/{listados} ({'; '.join(erros[:3])})",
        }
    return {"ok": True, "status": "Baixado", "arquivos": baixados, "observacao": obs}


def aguardar_resultados_juridiq(dlg, timeout_s: int = 12) -> None:
    """Aguarda a lista do comando de busca no Juridiq estabilizar após digitar."""
    pg = dlg.page
    pg.wait_for_timeout(1500)
    anterior, estavel = -1, 0
    for _ in range(timeout_s * 4):
        atual = dlg.locator("[cmdk-item]").count()
        vazio = dlg.get_by_text(re.compile(r"nenhum resultado", re.I)).count()
        if atual == anterior and (atual > 0 or vazio > 0):
            estavel += 1
            if estavel >= 3:
                return
        else:
            estavel = 0
        anterior = atual
        pg.wait_for_timeout(250)


def buscar_e_enviar_destino(pg, url_destino: str, nome_cliente: str, pasta_cliente: Path, log) -> dict:
    """Localiza o cliente no Juridiq com regra de homônimos e envia os arquivos sem duplicar."""
    pg.bring_to_front()
    pg.goto(url_destino)
    pg.wait_for_load_state("networkidle")

    pg.locator("#search-bar").click()
    dlg = pg.locator('[role="dialog"]').first
    dlg.wait_for(timeout=10000)
    dlg.locator("input[cmdk-input]").fill(nome_cliente)
    aguardar_resultados_juridiq(dlg)

    extrair_juridiq = lambda el: (el.get_attribute("data-value") or el.inner_text()).strip()
    st1, _, obs1 = selecionar_resultado_homonimos(dlg.locator("[cmdk-item]"), nome_cliente, extrair_juridiq)
    if st1 != "OK":
        pg.keyboard.press("Escape")
        return {"ok": False, "status": st1, "arquivos": 0, "observacao": f"Destino (1ª busca): {obs1}"}

    pg.wait_for_timeout(2000)
    if "/people/" not in pg.url:
        cli = pg.locator('[role="dialog"] [cmdk-group][data-value="Cliente"] [cmdk-item]')
        st2, _, obs2 = selecionar_resultado_homonimos(cli, nome_cliente, extrair_juridiq)
        if st2 != "OK":
            pg.keyboard.press("Escape")
            return {"ok": False, "status": st2, "arquivos": 0, "observacao": f"Destino (2ª busca): {obs2}"}

    pg.wait_for_url(re.compile(r"/people/"), timeout=15000)
    aba = pg.get_by_role("tab", name="Arquivos")
    if aba.count() == 0:
        aba = pg.get_by_text("Arquivos", exact=True)
    aba.first.click()
    pg.wait_for_url(re.compile(r"tab=arquivos"), timeout=15000)
    pg.locator("#files-list").first.wait_for(timeout=15000)
    pg.wait_for_timeout(2000)

    esperados, enviados, erros = enviar_cliente(pg, pasta_cliente, log=log)
    if enviados == esperados and not erros:
        return {
            "ok": True,
            "status": "Concluído",
            "arquivos": enviados,
            "observacao": f"Migrados/verificados {enviados}/{esperados} arquivo(s) sem duplicidade",
        }
    return {
        "ok": False,
        "status": "Erro no upload",
        "arquivos": enviados,
        "observacao": f"Destino: enviados {enviados}/{esperados} ({'; '.join(erros[:3])})",
    }


def salvar_relatorio(registros: list[dict], destino: Path = RELATORIO_SAIDA) -> None:
    """Atualiza status_migracao.xlsx com as colunas exigidas."""
    colunas = ["nome_cliente", "status", "arquivos_transferidos", "observacao"]
    df = pd.DataFrame(registros, columns=colunas)
    df.to_excel(destino, index=False, engine="openpyxl")


def migrar_cliente(
    ctx,
    pg_origem,
    pg_destino,
    url_origem: str,
    url_destino: str,
    nome_cliente: str,
    logger: logging.Logger,
) -> tuple[str, dict]:
    """Executa o fluxo Origem -> Destino para um cliente. Retorna (etapa, registro_relatorio)."""
    pasta_cliente = PASTA_DOWNLOADS / limpa(nome_cliente)

    res_origem = buscar_e_baixar_origem(pg_origem, ctx, url_origem, nome_cliente, pasta_cliente, logger.info)
    if not res_origem["ok"]:
        return "pendente", {
            "nome_cliente": nome_cliente,
            "status": res_origem["status"],
            "arquivos_transferidos": res_origem["arquivos"],
            "observacao": res_origem["observacao"],
        }

    res_destino = buscar_e_enviar_destino(pg_destino, url_destino, nome_cliente, pasta_cliente, logger.info)
    etapa = "concluido" if res_destino["ok"] else "pendente"
    return etapa, {
        "nome_cliente": nome_cliente,
        "status": res_destino["status"],
        "arquivos_transferidos": res_destino["arquivos"],
        "observacao": res_destino["observacao"],
    }


def parse_args():
    parser = argparse.ArgumentParser(description="Migração em lote de documentos entre AdvBox e Juridiq.")
    parser.add_argument("--arquivo", default=str(PLANILHA_PADRAO), help="Planilha de entrada (.xlsx)")
    parser.add_argument("--abas", nargs="+", default=list(ABAS_PADRAO), help="Abas a processar em sequência")
    parser.add_argument("--cliente", default=None, help="Processa apenas o cliente informado (opcional)")
    parser.add_argument("--limite", type=int, default=None, help="Número máximo de clientes a processar")
    parser.add_argument(
        "--reprocessar-concluidos",
        action="store_true",
        help="Reprocessa também clientes já marcados como 'concluido' na coluna Etapa",
    )
    parser.add_argument(
        "--novo-navegador",
        action="store_true",
        help="Abre uma nova janela do Chromium (headless=False) para login manual em vez de usar porta 9222",
    )
    return parser.parse_args()


def main() -> None:
    load_dotenv()
    args = parse_args()
    logger = configurar_logger()
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)

    caminho_planilha = Path(args.arquivo)
    url_origem = resolver_url("ORIGEM_URL", ORIGEM_URL_PADRAO)
    url_destino = resolver_url("DESTINO_URL", DESTINO_URL_PADRAO)

    clientes, eh_cadastro = carregar_clientes(
        caminho_planilha,
        tuple(args.abas),
        reprocessar_concluidos=args.reprocessar_concluidos,
    )
    if args.cliente:
        alvo = normalizar_texto(args.cliente)
        clientes = [c for c in clientes if normalizar_texto(c["nome_cliente"]) == alvo]
    if args.limite is not None:
        clientes = clientes[: args.limite]

    logger.info("Total de clientes na fila de processamento: %d", len(clientes))
    if not clientes:
        salvar_relatorio([])
        logger.info("Nenhum cliente pendente para processar.")
        return

    registros: list[dict] = []
    with sync_playwright() as p:
        if args.novo_navegador:
            browser = p.chromium.launch(headless=False)
            ctx = browser.new_context(accept_downloads=True)
            pg_origem = obter_ou_abrir_pagina(ctx, "advbox", url_origem)
            pg_destino = obter_ou_abrir_pagina(ctx, "juridiq", url_destino)
            input("Faça login manual nos dois CRMs (AdvBox e Juridiq) e pressione ENTER para iniciar... ")
        else:
            cdp_url = os.getenv("CHROME_CDP_URL", "http://localhost:9222")
            browser = p.chromium.connect_over_cdp(cdp_url)
            ctx = browser.contexts[0]
            pg_origem = obter_ou_abrir_pagina(ctx, "advbox", url_origem)
            pg_destino = obter_ou_abrir_pagina(ctx, "juridiq", url_destino)

        try:
            for pos, cli in enumerate(clientes, start=1):
                nome = cli["nome_cliente"]
                logger.info("[%d/%d] Processando (%s linha %d): %s", pos, len(clientes), cli["aba"], cli["linha_excel"], nome)
                try:
                    etapa, reg = migrar_cliente(ctx, pg_origem, pg_destino, url_origem, url_destino, nome, logger)
                except Exception as ex:
                    logger.exception("Erro inesperado ao processar %s", nome)
                    etapa = "pendente"
                    reg = {
                        "nome_cliente": nome,
                        "status": "Erro",
                        "arquivos_transferidos": 0,
                        "observacao": f"Exceção: {ex}",
                    }

                registros.append(reg)
                salvar_relatorio(registros)
                if eh_cadastro:
                    try:
                        atualizar_etapa_cliente(caminho_planilha, cli["aba"], cli["linha_excel"], etapa)
                    except PermissionError:
                        logger.error(
                            "Não foi possível salvar '%s' (feche a planilha no Excel!). Etapa do cliente %s seria '%s'.",
                            caminho_planilha,
                            nome,
                            etapa,
                        )
                logger.info(
                    "[%d/%d] Finalizado %s -> Etapa=%s | Status=%s | Arquivos=%s | %s",
                    pos,
                    len(clientes),
                    nome,
                    etapa,
                    reg["status"],
                    reg["arquivos_transferidos"],
                    reg["observacao"],
                )
        finally:
            browser.close()

    logger.info("Lote encerrado. Relatório salvo em %s", RELATORIO_SAIDA.resolve())


if __name__ == "__main__":
    main()


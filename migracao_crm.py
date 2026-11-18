"""Migrador em lote de documentos entre CRM de Origem (AdvBox) e CRM de Destino (Juridiq)."""
import argparse
import logging
import os
import re
import shutil
import sys
import unicodedata
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

from coleta_nuvem import (
    coletar,
    contar_estrutura_local,
    esperar_lista,
    formatar_resumo_docs,
    limpa,
)
from destino_juridiq import enviar_e_verificar_cliente
from planilha_cadastro import (
    ABAS_PADRAO,
    PLANILHA_PADRAO,
    atualizar_etapa_cliente,
    garantir_colunas_controle,
)

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
    reprocessar_pendentes: bool = False,
) -> tuple[list[dict], bool]:
    """Lê os clientes via pandas mantendo aba, linha no Excel e status de Etapa."""
    eh_cadastro = caminho_planilha.name.lower() == PLANILHA_PADRAO.name.lower()
    if eh_cadastro:
        garantir_colunas_controle(caminho_planilha, abas=abas)
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
            if etapa_atual == "pendente" and not reprocessar_pendentes:
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


def limpar_pasta_cliente(pasta: Path) -> None:
    """Garante pasta isolada limpa removendo apenas itens internos (evita WinError 5 no rmdir)."""
    if pasta.exists():
        for item in list(pasta.iterdir()):
            if item.is_dir():
                shutil.rmtree(item, ignore_errors=True)
            else:
                try:
                    item.unlink(missing_ok=True)
                except OSError:
                    pass
    pasta.mkdir(parents=True, exist_ok=True)


def buscar_e_baixar_origem(pg, ctx, url_origem: str, nome_cliente: str, pasta_cliente: Path, log) -> dict:
    """Executa busca com regra de homônimos e baixa arquivos recursivamente na Origem (AdvBox)."""
    limpar_pasta_cliente(pasta_cliente)

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
        return {
            "ok": False,
            "status": st,
            "arquivos": 0,
            "baixados": 0,
            "pastas": 0,
            "resumo_origem": "",
            "observacao": f"Origem: {obs}",
        }

    pg.locator(".popup-content").first.wait_for(timeout=15000)
    pg.wait_for_timeout(1200)
    pg.locator("a.openFolder").first.click()
    esperar_lista(pg)

    listados, baixados, erros = coletar(pg, ctx, pasta_cliente, log=log)
    _, pastas_locais = contar_estrutura_local(pasta_cliente)
    resumo_origem = formatar_resumo_docs(listados, pastas_locais)

    fechar = pg.locator("a.button-close-popup:visible")
    if fechar.count():
        fechar.first.click()
        pg.wait_for_timeout(400)

    if listados == 0:
        return {
            "ok": False,
            "status": "Não encontrado",
            "arquivos": 0,
            "baixados": 0,
            "pastas": pastas_locais,
            "resumo_origem": resumo_origem,
            "observacao": "Origem: nenhum documento na nuvem do cliente",
        }
    if baixados != listados or erros:
        return {
            "ok": False,
            "status": "Erro no download",
            "arquivos": listados,
            "baixados": baixados,
            "pastas": pastas_locais,
            "resumo_origem": resumo_origem,
            "observacao": f"Origem: baixados {baixados}/{listados} ({'; '.join(erros[:3])})",
        }
    return {
        "ok": True,
        "status": "Baixado",
        "arquivos": listados,
        "baixados": baixados,
        "pastas": pastas_locais,
        "resumo_origem": resumo_origem,
        "observacao": obs,
    }


def aguardar_resultados_juridiq(dlg, timeout_s: int = 12) -> None:
    """Aguarda a lista do comando de busca no Juridiq estabilizar após digitar."""
    pg = dlg.page
    pg.wait_for_timeout(1500)
    anterior, estavel = -1, 0
    for _ in range(timeout_s * 4):
        atual = dlg.locator("[cmdk-item]").count()
        vazio = dlg.get_by_text(
            re.compile(r"nenhum resultado|sem resultado|nada encontrado|não encontrado", re.I)
        ).count()
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
        return {
            "ok": False,
            "status": st1,
            "arquivos": 0,
            "pastas": 0,
            "resumo_destino": "",
            "erros": [obs1],
            "observacao": f"Destino (1ª busca): {obs1}",
        }

    pg.wait_for_timeout(2000)
    if "/people/" not in pg.url:
        cli = pg.locator('[role="dialog"] [cmdk-group][data-value="Cliente"] [cmdk-item]')
        st2, _, obs2 = selecionar_resultado_homonimos(cli, nome_cliente, extrair_juridiq)
        if st2 != "OK":
            pg.keyboard.press("Escape")
            return {
                "ok": False,
                "status": st2,
                "arquivos": 0,
                "pastas": 0,
                "resumo_destino": "",
                "erros": [obs2],
                "observacao": f"Destino (2ª busca): {obs2}",
            }

    pg.wait_for_url(re.compile(r"/people/"), timeout=15000, wait_until="commit")
    aba = pg.get_by_role("tab", name="Arquivos").or_(pg.locator("main").get_by_text("Arquivos", exact=True)).first
    aba.wait_for(state="visible", timeout=15000)
    aba.click()
    pg.locator("#files-list").first.wait_for(state="visible", timeout=15000)
    pg.wait_for_timeout(2000)

    docs_esp, docs_conf, pastas_esp, pastas_conf, erros = enviar_e_verificar_cliente(pg, pasta_cliente, log=log)
    resumo_destino = formatar_resumo_docs(docs_conf, pastas_conf)
    if docs_conf == docs_esp and pastas_conf == pastas_esp and not erros:
        return {
            "ok": True,
            "status": "Concluído",
            "arquivos": docs_conf,
            "pastas": pastas_conf,
            "resumo_destino": resumo_destino,
            "erros": [],
            "observacao": f"Migrados/verificados {resumo_destino} sem duplicidade",
        }
    return {
        "ok": False,
        "status": "Erro no upload",
        "arquivos": docs_conf,
        "pastas": pastas_conf,
        "resumo_destino": resumo_destino,
        "erros": erros,
        "observacao": (
            f"Destino: {docs_conf}/{docs_esp} docs e {pastas_conf}/{pastas_esp} pastas "
            f"({'; '.join(erros[:3])})"
        ),
    }


def salvar_relatorio(registros: list[dict], destino: Path = RELATORIO_SAIDA) -> None:
    """Atualiza status_migracao.xlsx com as colunas exigidas."""
    colunas = ["nome_cliente", "status", "arquivos_transferidos", "observacao"]
    df = pd.DataFrame(registros, columns=colunas)
    df.to_excel(destino, index=False, engine="openpyxl")


def diagnosticar_problema(contexto: str, erro_ou_status: str, detalhe: str = "") -> dict:
    """Traduz erros técnicos e falhas de fluxo para explicação simples e alternativas."""
    texto_completo = f"{contexto} {erro_ou_status} {detalhe}".lower()

    if contexto == "conexao_chrome" or "connect_over_cdp" in texto_completo or "9222" in texto_completo:
        return {
            "titulo": "Não foi possível conectar ao Chrome de automação",
            "o_que_ocorreu": (
                "O script tentou se conectar ao Chrome já logado na porta 9222, "
                "mas não encontrou o navegador aberto nesse modo."
            ),
            "como_identificar": (
                "Verifique se você abriu o navegador pelo arquivo 'abrir_chrome_9222.bat' "
                "e se a janela do Chrome continua aberta (um Chrome comum aberto fora do .bat não funciona)."
            ),
            "alternativas": [
                "Dê dois cliques em 'abrir_chrome_9222.bat', faça login no AdvBox e no Juridiq e rode 'python migracao_crm.py' novamente.",
                "Ou rode 'python migracao_crm.py --novo-navegador' para o script abrir um navegador novo e esperar seu login manual.",
            ],
        }

    if contexto == "planilha" or (
        ".xlsx" in texto_completo and ("permission" in texto_completo or "acesso negado" in texto_completo)
    ):
        return {
            "titulo": "Planilha bloqueada ou inacessível",
            "o_que_ocorreu": (
                "O Windows bloqueou a leitura ou gravação da planilha, provavelmente porque ela "
                "está aberta em outro programa (como o Excel)."
            ),
            "como_identificar": (
                "Olhe na barra de tarefas do Windows se 'PLANILHA DE CADASTRO.xlsx' ou "
                "'status_migracao.xlsx' está aberta no Excel."
            ),
            "alternativas": [
                "Feche a planilha no Excel (salvando antes se precisar) e tente novamente.",
                "Confira também se o arquivo existe na pasta correta e não está marcado como somente leitura.",
            ],
        }

    if "nenhum documento na nuvem" in texto_completo:
        return {
            "titulo": "Cliente sem arquivos na nuvem do AdvBox",
            "o_que_ocorreu": (
                "O cliente foi encontrado no AdvBox, mas a aba de arquivos (nuvem) dele estava vazia."
            ),
            "como_identificar": (
                "Abra o cadastro desse cliente no AdvBox e clique no ícone da nuvem para confirmar "
                "se não existe mesmo nenhum documento anexado (ou se está em outra área)."
            ),
            "alternativas": [
                "Se esse cliente realmente não tiver documentos, escolha [2] para deixar 'pendente' (ou [3] se já considerar concluído) e seguir para o próximo.",
                "Se os documentos estiverem em outra tela, envie manualmente no Juridiq e escolha [3].",
            ],
        }

    if erro_ou_status == "Não encontrado":
        crm_onde = "CRM de Origem (AdvBox)" if "origem" in texto_completo else "CRM de Destino (Juridiq)"
        return {
            "titulo": f"Cliente não encontrado no {crm_onde}",
            "o_que_ocorreu": (
                f"A pesquisa pelo nome exato do cliente não retornou nenhum cadastro no {crm_onde}."
            ),
            "como_identificar": (
                f"Busque esse cliente manualmente no {crm_onde} para ver se há erro de digitação, "
                "abreviação diferente na planilha ou se o cliente ainda não foi cadastrado lá."
            ),
            "alternativas": [
                "Corrija o nome na planilha ou crie/ajuste o cadastro no CRM e escolha [1] para tentar novamente.",
                "Ou escolha [2] para deixar este cliente como 'pendente' e continuar o lote pelos próximos clientes.",
            ],
        }

    if "ambigu" in erro_ou_status.lower():
        crm_onde = "CRM de Origem (AdvBox)" if "origem" in texto_completo else "CRM de Destino (Juridiq)"
        return {
            "titulo": f"Mais de um cliente com nome parecido no {crm_onde}",
            "o_que_ocorreu": (
                f"Apareceram vários resultados no {crm_onde} e não foi possível garantir qual é o cadastro certo. "
                "O script parou por segurança para não misturar documentos entre clientes."
            ),
            "como_identificar": (
                f"Pesquise esse nome no {crm_onde} e confira pelo CPF qual cadastro pertence ao cliente correto."
            ),
            "alternativas": [
                "Faça a migração deste cliente manualmente no navegador conferindo o CPF e escolha [3] para marcar como concluído e seguir.",
                "Ou escolha [2] para deixar como 'pendente' na planilha e continuar o lote pelos demais clientes.",
            ],
        }

    if erro_ou_status in ("Erro no download", "Erro no upload"):
        etapa_txt = "baixar da Origem (AdvBox)" if erro_ou_status == "Erro no download" else "enviar para o Destino (Juridiq)"
        return {
            "titulo": f"Falha parcial ao {etapa_txt}",
            "o_que_ocorreu": (
                f"Nem todos os documentos do cliente puderam ser confirmados ao {etapa_txt}. Detalhe: {detalhe}"
            ),
            "como_identificar": (
                "Confira a pasta './downloads/<nome_cliente>' e a aba 'Arquivos' do cliente no Juridiq "
                "para ver qual arquivo ficou faltando."
            ),
            "alternativas": [
                "Verifique a conexão/login no Chrome e escolha [1] para tentar este cliente novamente (arquivos já anexos com o mesmo nome são ignorados sem duplicar).",
                "Ou envie manualmente os arquivos que faltaram pela aba do navegador e escolha [3].",
            ],
        }

    if "timeout" in texto_completo or "waiting for" in texto_completo:
        return {
            "titulo": "A página demorou demais para responder",
            "o_que_ocorreu": (
                "Um botão, lista ou tela do CRM não carregou dentro do tempo limite esperado."
            ),
            "como_identificar": (
                "Olhe para a janela do Chrome neste momento: veja se o sistema deslogou (pedindo senha), "
                "se a internet caiu ou se algum aviso/pop-up ficou na frente bloqueando a tela."
            ),
            "alternativas": [
                "Se tiver deslogado ou travado num aviso, resolva direto na janela do Chrome e escolha [1] para tentar este cliente novamente.",
                "Ou escolha [4] para parar, revise a janela do Chrome e rode 'python migracao_crm.py' depois (clientes concluídos serão pulados).",
            ],
        }

    return {
        "titulo": "Erro inesperado durante a migração",
        "o_que_ocorreu": f"Aconteceu uma falha inesperada: {detalhe or erro_ou_status}",
        "como_identificar": (
            "Olhe na janela do Chrome em qual tela parou e confira os detalhes completos no arquivo 'migracao_crm.log'."
        ),
        "alternativas": [
            "Verifique se as duas abas (AdvBox e Juridiq) continuam abertas e logadas no Chrome e escolha [1] para tentar novamente.",
            "Ou escolha [2] para deixar este cliente como 'pendente' e seguir, ou [4] para parar o script agora.",
        ],
    }


def exibir_painel_erro(titulo_contexto: str, diag: dict, logger: logging.Logger) -> None:
    """Exibe e registra no log um resumo simples do erro com orientações e alternativas."""
    linhas = [
        "",
        "=" * 72,
        f"[PARADA POR ERRO] {titulo_contexto} -> {diag['titulo']}",
        "=" * 72,
        "O QUE OCORREU DE INESPERADO:",
        f"  {diag['o_que_ocorreu']}",
        "",
        "COMO IDENTIFICAR O PROBLEMA:",
        f"  {diag['como_identificar']}",
        "",
        "ALTERNATIVAS PARA RESOLVER:",
    ]
    for idx, alt in enumerate(diag["alternativas"], start=1):
        linhas.append(f"  {idx}. {alt}")
    linhas.append("=" * 72)
    bloco = "\n".join(linhas)
    print(bloco)
    logger.error("%s | %s | %s", titulo_contexto, diag["titulo"], diag["o_que_ocorreu"])


def tratar_parada_erro(
    cli: dict,
    reg: dict,
    dados_planilha: dict,
    caminho_planilha: Path,
    eh_cadastro: bool,
    registros: list[dict],
    logger: logging.Logger,
    nao_interativo: bool,
) -> str:
    """Salva o progresso, explica o erro ao usuário e pergunta como prosseguir."""
    nome = cli["nome_cliente"]
    diag = diagnosticar_problema("migracao", reg["status"], reg["observacao"])
    exibir_painel_erro(f"{nome} ({cli['aba']} linha {cli['linha_excel']})", diag, logger)

    if nao_interativo or not sys.stdin.isatty():
        print(
            "\nScript interrompido para sua verificação. "
            "Após resolver, rode novamente: python migracao_crm.py"
        )
        return "parar"

    print("\nO que deseja fazer agora?")
    print("  [1] Já resolvi — tentar novamente este mesmo cliente")
    print("  [2] Deixar como 'pendente' e continuar para o próximo cliente")
    print("  [3] Já resolvi manualmente no CRM — marcar como 'concluido' e continuar")
    print("  [4] Parar e encerrar o script agora (padrão)")
    escolha = input("Escolha [1/2/3/4]: ").strip()

    if escolha == "1":
        return "tentar_novamente"
    if escolha == "2":
        return "proximo"
    if escolha == "3":
        reg["status"] = "Concluído"
        reg["observacao"] = "Concluído manualmente pelo usuário após parada assistida"
        salvar_relatorio(registros)
        if eh_cadastro:
            try:
                atualizar_etapa_cliente(
                    caminho_planilha,
                    cli["aba"],
                    cli["linha_excel"],
                    "concluido",
                    docs_origem=dados_planilha.get("docs_origem"),
                    docs_destino=dados_planilha.get("docs_destino"),
                    obs="Concluído manualmente pelo usuário",
                )
            except PermissionError:
                diag_perm = diagnosticar_problema("planilha", "PermissionError", str(caminho_planilha))
                exibir_painel_erro("Gravação na planilha", diag_perm, logger)
                return "parar"
        return "proximo"

    print("\nExecução encerrada. Quando estiver pronto para retomar, execute: python migracao_crm.py")
    return "parar"


def avaliar_resultado_migracao(res_origem: dict, res_destino: dict | None, nome_cliente: str) -> tuple[str, dict, dict]:
    """Confere igualdade Origem vs Destino e monta relatório + colunas da planilha."""
    if not res_origem["ok"]:
        st = res_origem["status"]
        obs_bruta = res_origem["observacao"]
        if "nenhum documento na nuvem" in obs_bruta.lower():
            obs_planilha = "Cliente sem documentos no CRM de Origem"
        elif st == "Não encontrado":
            obs_planilha = "Cliente não encontrado no CRM de Origem"
        elif "ambigu" in st.lower():
            obs_planilha = obs_bruta.replace("Origem: ", "Ambiguidade na Origem: ")
        elif st == "Erro no download":
            falt = max(0, res_origem["arquivos"] - res_origem.get("baixados", 0))
            if falt > 0:
                doc_txt = "documento" if falt == 1 else "documentos"
                obs_planilha = f"{falt} {doc_txt} com erro ao baixar da origem"
            else:
                obs_planilha = obs_bruta
        else:
            obs_planilha = obs_bruta

        return (
            "pendente",
            {
                "nome_cliente": nome_cliente,
                "status": st,
                "arquivos_transferidos": res_origem.get("baixados", 0),
                "observacao": obs_bruta,
            },
            {
                "docs_origem": res_origem.get("resumo_origem", ""),
                "docs_destino": "",
                "obs": obs_planilha,
            },
        )

    assert res_destino is not None
    docs_orig = res_origem["arquivos"]
    pastas_orig = res_origem["pastas"]
    docs_dest = res_destino["arquivos"]
    pastas_dest = res_destino["pastas"]
    erros_dest = res_destino.get("erros") or []

    tudo_igual = (
        res_destino["ok"]
        and docs_orig == docs_dest
        and pastas_orig == pastas_dest
        and len(erros_dest) == 0
    )

    if tudo_igual:
        return (
            "concluido",
            {
                "nome_cliente": nome_cliente,
                "status": "Concluído",
                "arquivos_transferidos": docs_dest,
                "observacao": res_destino["observacao"],
            },
            {
                "docs_origem": res_origem["resumo_origem"],
                "docs_destino": res_destino["resumo_destino"],
                "obs": "",
            },
        )

    st_dest = res_destino["status"]
    if st_dest == "Não encontrado":
        obs_planilha = "Cliente não encontrado no CRM de Destino"
    elif "ambigu" in st_dest.lower():
        obs_planilha = res_destino["observacao"].replace("Destino", "Ambiguidade no Destino")
    else:
        partes: list[str] = []
        diff_docs = max(0, docs_orig - docs_dest)
        diff_pastas = max(0, pastas_orig - pastas_dest)
        if diff_docs > 0:
            doc_txt = "documento" if diff_docs == 1 else "documentos"
            partes.append(f"{diff_docs} {doc_txt} com erro de migração")
        if diff_pastas > 0:
            pasta_txt = "pasta" if diff_pastas == 1 else "pastas"
            partes.append(f"{diff_pastas} {pasta_txt} com erro de migração")
        if not partes:
            partes.append("; ".join(erros_dest[:2]) if erros_dest else res_destino["observacao"])
        obs_planilha = "; ".join(partes)

    return (
        "pendente",
        {
            "nome_cliente": nome_cliente,
            "status": "Erro no upload" if st_dest == "Concluído" else st_dest,
            "arquivos_transferidos": docs_dest,
            "observacao": f"Origem ({res_origem['resumo_origem']}) != Destino ({res_destino['resumo_destino']}): {obs_planilha}"
            if st_dest == "Concluído"
            else res_destino["observacao"],
        },
        {
            "docs_origem": res_origem["resumo_origem"],
            "docs_destino": res_destino.get("resumo_destino", ""),
            "obs": obs_planilha,
        },
    )


def migrar_cliente(
    ctx,
    pg_origem,
    pg_destino,
    url_origem: str,
    url_destino: str,
    nome_cliente: str,
    logger: logging.Logger,
) -> tuple[str, dict, dict]:
    """Executa o fluxo Origem -> Destino e retorna (etapa, registro_relatorio, dados_planilha)."""
    pasta_cliente = PASTA_DOWNLOADS / limpa(nome_cliente)

    res_origem = buscar_e_baixar_origem(pg_origem, ctx, url_origem, nome_cliente, pasta_cliente, logger.info)
    if not res_origem["ok"]:
        return avaliar_resultado_migracao(res_origem, None, nome_cliente)

    try:
        res_destino = buscar_e_enviar_destino(pg_destino, url_destino, nome_cliente, pasta_cliente, logger.info)
    except Exception as ex:
        logger.exception("Erro inesperado no destino ao processar %s", nome_cliente)
        return (
            "pendente",
            {
                "nome_cliente": nome_cliente,
                "status": "Erro",
                "arquivos_transferidos": 0,
                "observacao": f"{type(ex).__name__}: {ex}",
            },
            {
                "docs_origem": res_origem["resumo_origem"],
                "docs_destino": "",
                "obs": f"Erro no destino: {type(ex).__name__}",
            },
        )
    return avaliar_resultado_migracao(res_origem, res_destino, nome_cliente)


def eh_condicao_dados(status: str) -> bool:
    """Identifica condições de cadastro/dados que devem marcar pendente e seguir o lote."""
    return status == "Não encontrado" or "ambigu" in status.lower()


def eh_erro_critico_global(pg_origem, pg_destino, detalhe: str) -> bool:
    """Verifica se o navegador desconectou ou se alguma sessão de login expirou."""
    txt = (detalhe or "").lower()
    sinais_fatal = (
        "target page, context or browser has been closed",
        "browser has been closed",
        "connection closed",
        "connect_over_cdp",
        "econnrefused",
    )
    if any(s in txt for s in sinais_fatal):
        return True
    try:
        if pg_origem.is_closed() or pg_destino.is_closed():
            return True
        if "/login" in pg_origem.url.lower() or "/login" in pg_destino.url.lower():
            return True
    except Exception:
        return True
    return False


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
        "--reprocessar-pendentes",
        action="store_true",
        help="Reprocessa também clientes já marcados como 'pendente' na coluna Etapa",
    )
    parser.add_argument(
        "--novo-navegador",
        action="store_true",
        help="Abre uma nova janela do Chromium (headless=False) para login manual em vez de usar porta 9222",
    )
    parser.add_argument(
        "--nao-interativo",
        action="store_true",
        help="Em caso de erro crítico ou após retry em lote, não aguarda confirmação no terminal",
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

    try:
        clientes, eh_cadastro = carregar_clientes(
            caminho_planilha,
            tuple(args.abas),
            reprocessar_concluidos=args.reprocessar_concluidos,
            reprocessar_pendentes=args.reprocessar_pendentes or bool(args.cliente),
        )
    except Exception as ex:
        diag = diagnosticar_problema("planilha", type(ex).__name__, str(ex))
        exibir_painel_erro("Leitura da planilha de entrada", diag, logger)
        sys.exit(1)

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
    parar_lote = False
    with sync_playwright() as p:
        try:
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
        except Exception as ex:
            diag = diagnosticar_problema("conexao_chrome", type(ex).__name__, str(ex))
            exibir_painel_erro("Conexão com o navegador Chrome", diag, logger)
            sys.exit(1)

        try:
            for pos, cli in enumerate(clientes, start=1):
                nome = cli["nome_cliente"]
                indice_reg = None
                tentativa_auto = 1
                while True:
                    logger.info(
                        "[%d/%d] Processando (%s linha %d, tentativa %d): %s",
                        pos,
                        len(clientes),
                        cli["aba"],
                        cli["linha_excel"],
                        tentativa_auto,
                        nome,
                    )
                    try:
                        etapa, reg, dados_planilha = migrar_cliente(
                            ctx, pg_origem, pg_destino, url_origem, url_destino, nome, logger
                        )
                    except Exception as ex:
                        logger.exception("Erro inesperado ao processar %s", nome)
                        etapa = "pendente"
                        reg = {
                            "nome_cliente": nome,
                            "status": "Erro",
                            "arquivos_transferidos": 0,
                            "observacao": f"{type(ex).__name__}: {ex}",
                        }
                        dados_planilha = {
                            "docs_origem": "",
                            "docs_destino": "",
                            "obs": f"Erro inesperado na migração: {ex}",
                        }

                    if indice_reg is None:
                        registros.append(reg)
                        indice_reg = len(registros) - 1
                    else:
                        registros[indice_reg] = reg

                    try:
                        salvar_relatorio(registros)
                        if eh_cadastro:
                            atualizar_etapa_cliente(
                                caminho_planilha,
                                cli["aba"],
                                cli["linha_excel"],
                                etapa,
                                docs_origem=dados_planilha["docs_origem"],
                                docs_destino=dados_planilha["docs_destino"],
                                obs=dados_planilha["obs"],
                            )
                    except PermissionError as ex:
                        diag = diagnosticar_problema("planilha", "PermissionError", str(ex))
                        exibir_painel_erro(f"Gravação da planilha ({nome})", diag, logger)
                        parar_lote = True
                        break

                    logger.info(
                        "[%d/%d] Finalizado %s -> Etapa=%s | Status=%s | Origem=%s | Destino=%s | Obs=%s",
                        pos,
                        len(clientes),
                        nome,
                        etapa,
                        reg["status"],
                        dados_planilha["docs_origem"] or "-",
                        dados_planilha["docs_destino"] or "-",
                        dados_planilha["obs"] or "-",
                    )

                    if etapa == "concluido":
                        break

                    # Condições de dados (não encontrado, ambiguidade, nuvem vazia) avançam o lote sem travar
                    if eh_condicao_dados(reg["status"]):
                        diag = diagnosticar_problema("migracao", reg["status"], reg["observacao"])
                        logger.warning(
                            "[%d/%d] %s marcado como 'pendente' (%s). Avançando para o próximo cliente.",
                            pos,
                            len(clientes),
                            nome,
                            diag["titulo"],
                        )
                        break

                    # Erro crítico/global (navegador fechado ou sessão deslogada) para o script imediatamente
                    if eh_erro_critico_global(pg_origem, pg_destino, reg["observacao"]):
                        diag = diagnosticar_problema("migracao", reg["status"], reg["observacao"])
                        exibir_painel_erro(f"{nome} ({cli['aba']} linha {cli['linha_excel']})", diag, logger)
                        parar_lote = True
                        break

                    # Erro técnico transitório: tenta automaticamente mais 1 vez antes de marcar pendente e seguir
                    if tentativa_auto < 2:
                        tentativa_auto += 1
                        logger.warning(
                            "[%d/%d] Falha transitória em %s (%s). Executando tentativa 2/2 automaticamente...",
                            pos,
                            len(clientes),
                            nome,
                            reg["observacao"],
                        )
                        continue

                    if args.nao_interativo or not sys.stdin.isatty():
                        diag = diagnosticar_problema("migracao", reg["status"], reg["observacao"])
                        exibir_painel_erro(f"{nome} ({cli['aba']} linha {cli['linha_excel']})", diag, logger)
                        logger.warning(
                            "[%d/%d] %s manteve falha após tentativa 2/2; registrado como 'pendente' e seguindo lote.",
                            pos,
                            len(clientes),
                            nome,
                        )
                        break

                    acao = tratar_parada_erro(
                        cli,
                        reg,
                        dados_planilha,
                        caminho_planilha,
                        eh_cadastro,
                        registros,
                        logger,
                        args.nao_interativo,
                    )
                    if acao == "tentar_novamente":
                        continue
                    if acao == "proximo":
                        break
                    parar_lote = True
                    break

                if parar_lote:
                    break
        finally:
            browser.close()

    if parar_lote:
        logger.warning("Execução interrompida após erro. Relatório parcial salvo em %s", RELATORIO_SAIDA.resolve())
        sys.exit(1)

    logger.info("Lote encerrado. Relatório salvo em %s", RELATORIO_SAIDA.resolve())


if __name__ == "__main__":
    main()


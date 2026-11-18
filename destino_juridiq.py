"""Helpers do CRM de Destino (Juridiq)."""
import re
import time
from pathlib import Path


def abrir_arquivos_cliente(pg, nome):
    """Busca o cliente e abre a aba Arquivos. Retorna ('ok', n) ou (status, n_resultados)."""
    pg.goto("https://dashboard.juridiq.com.br/")
    pg.wait_for_load_state("networkidle")
    pg.locator("#search-bar").click()
    dlg = pg.locator('[role="dialog"]').first
    dlg.locator("input[cmdk-input]").fill(nome)
    pg.wait_for_timeout(2500)
    itens = dlg.locator("[cmdk-item]")
    n = itens.count()
    if n == 0:
        return "nao_encontrado", 0
    if n > 1:
        exatos = itens.filter(has_text=re.compile(rf"^\s*{re.escape(nome)}\s*$", re.I))
        if exatos.count() != 1:
            return "ambiguo", n
        exatos.first.click()
    else:
        itens.first.click()
    pg.wait_for_timeout(2500)
    cli = pg.locator('[role="dialog"] [cmdk-group][data-value="Cliente"] [cmdk-item]')
    nc = cli.count()
    if nc != 1:
        exatos = cli.filter(has_text=re.compile(rf"^\s*{re.escape(nome)}\s*$", re.I))
        if exatos.count() != 1:
            return ("nao_encontrado" if nc == 0 else "ambiguo"), nc
        exatos.first.click()
    else:
        cli.first.click()
    pg.wait_for_url(re.compile(r"/people/"), timeout=15000, wait_until="commit")
    aba = pg.get_by_role("tab", name="Arquivos").or_(pg.locator("main").get_by_text("Arquivos", exact=True)).first
    aba.wait_for(state="visible", timeout=15000)
    aba.click()
    pg.locator("#files-list").first.wait_for(state="visible", timeout=15000)
    pg.wait_for_timeout(2000)
    return "ok", 1


def nomes_na_lista(pg):
    return pg.locator("#files-list").inner_text()


def arquivos_na_lista(pg) -> set[str]:
    """Extrai de #files-list os nomes completos dos arquivos já anexados na pasta atual."""
    nomes = pg.evaluate(
        """() => {
          const root = document.querySelector('#files-list');
          if (!root) return [];
          const out = new Set();
          root.querySelectorAll('[aria-label^="Pré-visualizar "]').forEach(el => {
            const lbl = (el.getAttribute('aria-label') || '').replace(/^Pré-visualizar\\s+/, '').trim();
            if (lbl) out.add(lbl);
          });
          root.querySelectorAll('p').forEach(el => {
            const txt = (el.innerText || el.textContent || '').trim();
            if (txt) out.add(txt);
          });
          return Array.from(out);
        }"""
    )
    return {n.strip() for n in nomes if n and n.strip()}


def _nome_ja_anexado(nome_arquivo: str, existentes: set[str]) -> bool:
    """Compara o nome completo do arquivo (tratando truncamento de 80 chars do Juridiq)."""
    alvo = nome_arquivo.strip()
    if alvo in existentes:
        return True
    alvo_cf = alvo.casefold()
    if any(ex.strip().casefold() == alvo_cf for ex in existentes):
        return True
    if len(alvo) > 80:
        alvo_p = Path(alvo)
        alvo_stem = alvo_p.stem.casefold()
        alvo_suf = alvo_p.suffix.casefold()
        for ex in existentes:
            ex_p = Path(ex.strip())
            ex_stem = ex_p.stem.casefold()
            if ex_p.suffix.casefold() == alvo_suf and len(ex_stem) >= 70 and alvo_stem.startswith(ex_stem):
                return True
    return False


def _agrupar_lotes(arquivos, max_arquivos=5, max_bytes=35 * 1024 * 1024):
    """Agrupa arquivos respeitando limite de quantidade e tamanho total por lote."""
    lotes: list[list[Path]] = []
    lote_atual: list[Path] = []
    bytes_atual = 0
    for arq in arquivos:
        p = Path(arq)
        tam = p.stat().st_size if p.exists() else 0
        if lote_atual and (len(lote_atual) >= max_arquivos or bytes_atual + tam > max_bytes):
            lotes.append(lote_atual)
            lote_atual = []
            bytes_atual = 0
        lote_atual.append(p)
        bytes_atual += tam
    if lote_atual:
        lotes.append(lote_atual)
    return lotes


def _set_input_files_cdp(pg, caminhos: list[str]) -> None:
    """Passa os caminhos locais diretamente via CDP (sem transferir bytes nem limite de 50 MB)."""
    cdp = pg.context.new_cdp_session(pg)
    try:
        doc = cdp.send("DOM.getDocument", {"depth": 0})
        res = cdp.send(
            "DOM.querySelector",
            {"nodeId": doc["root"]["nodeId"], "selector": 'input[type="file"][multiple]'},
        )
        node_id = res.get("nodeId")
        if not node_id:
            raise RuntimeError("Input de upload não encontrado no DOM via CDP")
        cdp.send("DOM.setFileInputFiles", {"nodeId": node_id, "files": caminhos})
    finally:
        try:
            cdp.detach()
        except Exception:
            pass


def anexar(pg, arquivos, timeout_s=180, tamanho_lote=5):
    """Anexa arquivos em sub-lotes pelo input da dropzone e confirma o aparecimento na lista."""
    faltando_total: list[str] = []
    for lote in _agrupar_lotes(arquivos, max_arquivos=tamanho_lote):
        caminhos = [str(a.resolve()) for a in lote]
        try:
            _set_input_files_cdp(pg, caminhos)
        except Exception:
            inp = pg.locator('input[type="file"][multiple]').first
            inp.set_input_files(caminhos, timeout=120000)
        pg.wait_for_timeout(500)
        bytes_lote = sum(a.stat().st_size for a in lote if a.exists())
        espera_s = max(timeout_s, 180 + int(bytes_lote / (1024 * 1024) * 4))
        fim = time.time() + espera_s
        faltando = list(lote)
        while time.time() < fim:
            existentes = arquivos_na_lista(pg)
            faltando = [a for a in lote if not _nome_ja_anexado(a.name, existentes)]
            if not faltando:
                break
            pg.wait_for_timeout(1500)
        faltando_total.extend(a.name for a in faltando)
    return faltando_total


def _locator_pasta(pg, nome_pasta: str):
    padrao = re.compile(rf"^\s*{re.escape(nome_pasta.strip())}\s*$", re.I)
    return pg.locator("#files-list").get_by_text(padrao)


def _existe_na_lista(pg, nome: str) -> bool:
    return _locator_pasta(pg, nome).count() > 0


def _entrar_pasta(pg, nome_pasta: str):
    if not _existe_na_lista(pg, nome_pasta):
        pg.get_by_role("button", name="Nova pasta").click()
        pg.get_by_placeholder("Nome da pasta").fill(nome_pasta)
        pg.get_by_role("button", name="Confirmar").click()
        _locator_pasta(pg, nome_pasta).first.wait_for(state="visible", timeout=12000)
        pg.wait_for_timeout(600)
    _locator_pasta(pg, nome_pasta).first.click()
    padrao = re.compile(rf"^\s*{re.escape(nome_pasta.strip())}\s*$", re.I)
    try:
        _locator_pasta(pg, nome_pasta).first.wait_for(state="hidden", timeout=8000)
    except Exception:
        pass
    pg.get_by_text(padrao).first.wait_for(state="visible", timeout=10000)
    pg.wait_for_timeout(1200)


def _voltar(pg, nome_pasta: str, pai_pasta: str | None = None):
    upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÈÊÍÌÎÓÒÔÕÚÙÛÇ"
    lower = "abcdefghijklmnopqrstuvwxyzáàâãéèêíìîóòôõúùûç"
    atual_cf = nome_pasta.strip().lower()
    alvo_cf = (pai_pasta or "Arquivos").strip().lower()
    pg.locator(
        f"xpath=(//*[translate(normalize-space(text()), '{upper}', '{lower}')='{atual_cf}'])[1]"
        f"/preceding::*[translate(normalize-space(text()), '{upper}', '{lower}')='{alvo_cf}'][1]"
    ).click()
    _locator_pasta(pg, nome_pasta).first.wait_for(state="visible", timeout=12000)
    pg.wait_for_timeout(1200)


def enviar_e_verificar_cliente(pg, pasta_local, log=print, nivel=0, pai_pasta: str | None = None):
    """Espelha todas as pastas e arquivos de pasta_local no destino sem duplicar e confirma tudo.

    Retorna (docs_esperados, docs_confirmados, pastas_esperadas, pastas_confirmadas, erros).
    """
    pasta_local = Path(pasta_local)
    docs_esperados = docs_confirmados = 0
    pastas_esperadas = pastas_confirmadas = 0
    erros: list[str] = []
    ind = "  " * nivel
    subpastas = sorted(d for d in pasta_local.iterdir() if d.is_dir())
    arquivos = sorted(f for f in pasta_local.iterdir() if f.is_file())

    pastas_esperadas += len(subpastas)
    for sp in subpastas:
        log(f"{ind}pasta: {sp.name}")
        try:
            _entrar_pasta(pg, sp.name)
            de, dc, pe, pc, er = enviar_e_verificar_cliente(
                pg,
                sp,
                log=log,
                nivel=nivel + 1,
                pai_pasta=sp.name,
            )
            docs_esperados += de
            docs_confirmados += dc
            pastas_esperadas += pe
            pastas_confirmadas += pc
            erros.extend(er)
            _voltar(pg, sp.name, pai_pasta=pai_pasta)
            if _existe_na_lista(pg, sp.name):
                pastas_confirmadas += 1
            else:
                erros.append(f"pasta não confirmada no destino: {sp.name}")
        except Exception as ex:
            erros.append(f"pasta {sp.name}: {ex}")
            log(f"{ind}ERRO pasta {sp.name}: {ex}")
            return docs_esperados, docs_confirmados, pastas_esperadas, pastas_confirmadas, erros

    docs_esperados += len(arquivos)
    existentes = arquivos_na_lista(pg)
    ja_presentes = [a for a in arquivos if _nome_ja_anexado(a.name, existentes)]
    novos = [a for a in arquivos if not _nome_ja_anexado(a.name, existentes)]
    for arq in ja_presentes:
        log(f"{ind}ignorado (já anexo): {arq.name}")

    if novos:
        faltando = anexar(pg, novos)
        if faltando:
            erros.append(f"{pasta_local.name}: nao apareceram {faltando}")
        log(f"{ind}enviados {len(novos) - len(faltando)}/{len(novos)} em '{pasta_local.name}'")

    # Conferência final rigorosa na lista visível da pasta atual
    existentes_pos = arquivos_na_lista(pg)
    confirmados_aqui = [a for a in arquivos if _nome_ja_anexado(a.name, existentes_pos)]
    docs_confirmados += len(confirmados_aqui)
    if len(confirmados_aqui) != len(arquivos):
        nao_conf = [a.name for a in arquivos if a not in confirmados_aqui]
        msg = f"{pasta_local.name}: falta confirmar {nao_conf}"
        if msg not in erros:
            erros.append(msg)

    return docs_esperados, docs_confirmados, pastas_esperadas, pastas_confirmadas, erros


def enviar_cliente(pg, pasta_local, log=print, nivel=0):
    """Compatibilidade: retorna (esperados, enviados, erros)."""
    de, dc, _, _, erros = enviar_e_verificar_cliente(pg, pasta_local, log=log, nivel=nivel)
    return de, dc, erros

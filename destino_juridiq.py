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
    pg.wait_for_url(re.compile(r"/people/"), timeout=15000)
    pg.get_by_role("tab", name="Arquivos").first.click()
    pg.wait_for_url(re.compile(r"tab=arquivos"), timeout=15000)
    pg.wait_for_timeout(2500)
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
    """Compara o nome completo do arquivo (incluindo extensão) de forma exata e case-insensitive."""
    alvo = nome_arquivo.strip()
    if alvo in existentes:
        return True
    alvo_cf = alvo.casefold()
    return any(ex.strip().casefold() == alvo_cf for ex in existentes)


def anexar(pg, arquivos, timeout_s=180):
    """Anexa vários arquivos pelo input da dropzone e espera todos aparecerem com nome exato na lista."""
    pg.locator('input[type="file"][multiple]').first.set_input_files([str(a) for a in arquivos])
    fim = time.time() + timeout_s
    faltando = list(arquivos)
    while time.time() < fim:
        existentes = arquivos_na_lista(pg)
        faltando = [a for a in arquivos if not _nome_ja_anexado(Path(a).name, existentes)]
        if not faltando:
            return []
        pg.wait_for_timeout(1500)
    return [Path(a).name for a in faltando]


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
        pg.wait_for_timeout(2500)
    _locator_pasta(pg, nome_pasta).first.click()
    pg.wait_for_timeout(2000)
    padrao = re.compile(rf"^\s*{re.escape(nome_pasta.strip())}\s*$", re.I)
    pg.get_by_text(padrao).first.wait_for(timeout=10000)  # breadcrumb


def _voltar(pg, nome_pasta: str):
    padrao = nome_pasta.strip().lower()
    upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÈÊÍÌÎÓÒÔÕÚÙÛÇ"
    lower = "abcdefghijklmnopqrstuvwxyzáàâãéèêíìîóòôõúùûç"
    pg.locator(
        f"xpath=(//*[translate(normalize-space(text()), '{upper}', '{lower}')='{padrao}'])[1]"
        "/preceding::*[normalize-space(text())='Arquivos'][1]"
    ).click()
    pg.wait_for_timeout(2000)


def enviar_cliente(pg, pasta_local, log=print, nivel=0):
    """Espelha pasta_local no cliente aberto, ignorando arquivos já existentes.

    Retorna (esperados, enviados, erros).
    """
    pasta_local = Path(pasta_local)
    esperados = enviados = 0
    erros = []
    ind = "  " * nivel
    subpastas = sorted(d for d in pasta_local.iterdir() if d.is_dir())
    arquivos = sorted(f for f in pasta_local.iterdir() if f.is_file())

    for sp in subpastas:
        log(f"{ind}pasta: {sp.name}")
        try:
            _entrar_pasta(pg, sp.name)
            e, s, er = enviar_cliente(pg, sp, log, nivel + 1)
            esperados += e
            enviados += s
            erros += er
            _voltar(pg, sp.name)
        except Exception as ex:
            erros.append(f"pasta {sp.name}: {ex}")
            log(f"{ind}ERRO pasta {sp.name}: {ex}")
            return esperados, enviados, erros

    esperados += len(arquivos)
    existentes = arquivos_na_lista(pg)
    ja_presentes = [a for a in arquivos if _nome_ja_anexado(a.name, existentes)]
    novos = [a for a in arquivos if not _nome_ja_anexado(a.name, existentes)]
    for arq in ja_presentes:
        log(f"{ind}ignorado (já anexo): {arq.name}")
    if novos:
        faltando = anexar(pg, novos)
        if faltando:
            erros.append(f"{pasta_local.name}: nao apareceram {faltando}")
        enviados += len(novos) - len(faltando)
        log(f"{ind}enviados {len(novos) - len(faltando)}/{len(novos)} em '{pasta_local.name}'")
    enviados += len(ja_presentes)
    return esperados, enviados, erros

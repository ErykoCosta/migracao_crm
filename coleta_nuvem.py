"""Coleta recursiva de arquivos da nuvem (AdvBox) -> pasta local espelhando as pastas do CRM."""
import re
from pathlib import Path
from urllib.parse import urljoin

LI = "ul.files > li:not(.new-folder):not(.item-load)"


def limpa(nome: str) -> str:
    nome = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nome.strip()).rstrip(". ")
    return nome or "sem_nome"


def unico(pasta: Path, nome: str) -> Path:
    p = pasta / nome
    i = 1
    while p.exists():
        p = pasta / f"{p.stem} ({i}){p.suffix}"
        i += 1
    return p


def esperar_lista(pg, timeout_s=12):
    """Espera a lista parar de mudar (contagem + nomes estáveis)."""
    pg.wait_for_timeout(800)
    anterior, estavel = None, 0
    for _ in range(timeout_s * 4):
        atual = tuple(pg.locator(f"{LI} p.name").all_inner_texts())
        if atual == anterior and atual:
            estavel += 1
            if estavel >= 3:
                return atual
        else:
            estavel = 0
        anterior = atual
        pg.wait_for_timeout(250)
    return anterior or ()


def coletar(pg, ctx, destino: Path, log=print, nivel=0):
    """Retorna (listados, baixados, erros[list]). Pastas primeiro (recursivo), depois arquivos soltos."""
    destino.mkdir(parents=True, exist_ok=True)
    listados = baixados = 0
    erros = []
    ind = "  " * nivel

    itens = pg.locator(LI)
    n = itens.count()
    pastas, arquivos = [], []
    for i in range(n):
        li = itens.nth(i)
        nome = li.locator("p.name").first.inner_text().strip()
        if li.locator("a.item.folder").count():
            pastas.append(nome)
        elif li.locator("a.download-file").count():
            href = li.locator("a.download-file").first.get_attribute("href")
            arquivos.append((nome, urljoin(pg.url, href)))
    log(f"{ind}[{destino.name}] {len(pastas)} pasta(s), {len(arquivos)} arquivo(s)")

    # 1) pastas
    for nome_pasta in pastas:
        link = pg.locator(f'{LI} a.item.folder[data-path="{nome_pasta}"]')
        if link.count() == 0:
            link = pg.locator(f"{LI} a.item.folder").filter(has_text=nome_pasta)
        link.first.click()
        esperar_lista(pg)
        l, b, e = coletar(pg, ctx, destino / limpa(nome_pasta), log, nivel + 1)
        listados += l; baixados += b; erros += e
        pg.locator("a.folder-open-back").first.click()
        esperar_lista(pg)

    # 2) arquivos soltos
    for nome, url in arquivos:
        listados += 1
        try:
            r = ctx.request.get(url)
            if not r.ok:
                raise RuntimeError(f"HTTP {r.status}")
            unico(destino, limpa(nome)).write_bytes(r.body())
            baixados += 1
            log(f"{ind}  OK  {nome}")
        except Exception as ex:
            erros.append(f"{destino.name}/{nome}: {ex}")
            log(f"{ind}  ERRO {nome}: {ex}")
    return listados, baixados, erros


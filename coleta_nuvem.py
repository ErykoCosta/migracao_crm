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


def contar_estrutura_local(pasta_cliente: Path) -> tuple[int, int]:
    """Retorna (total_arquivos, total_subpastas) existentes dentro de pasta_cliente."""
    pasta = Path(pasta_cliente)
    if not pasta.exists():
        return 0, 0
    docs = sum(1 for item in pasta.rglob("*") if item.is_file())
    pastas = sum(1 for item in pasta.rglob("*") if item.is_dir())
    return docs, pastas


def formatar_resumo_docs(qtd_docs: int, qtd_pastas: int) -> str:
    """Formata a contagem no padrão '17 docs e 1 pasta'."""
    doc_lbl = "doc" if qtd_docs == 1 else "docs"
    pasta_lbl = "pasta" if qtd_pastas == 1 else "pastas"
    return f"{qtd_docs} {doc_lbl} e {qtd_pastas} {pasta_lbl}"


def esperar_lista(pg, timeout_s=12):
    """Espera a lista parar de mudar (contagem + nomes estáveis)."""
    pg.wait_for_timeout(800)
    anterior, estavel_itens, estavel_vazio = None, 0, 0
    for _ in range(timeout_s * 4):
        atual = tuple(pg.locator(f"{LI} p.name").all_inner_texts())
        if atual == anterior:
            if atual:
                estavel_itens += 1
                estavel_vazio = 0
                if estavel_itens >= 3:
                    return atual
            else:
                estavel_vazio += 1
                estavel_itens = 0
                if estavel_vazio >= 6:
                    return ()
        else:
            estavel_itens = 0
            estavel_vazio = 0
        anterior = atual
        pg.wait_for_timeout(250)
    return anterior or ()


def _abrir_subpasta(pg, nome_pasta: str) -> None:
    padrao = re.compile(rf"^\s*{re.escape(nome_pasta.strip())}\s*$")
    item = pg.locator(f"{LI}:has(a.item.folder)").filter(
        has=pg.locator("p.name", has_text=padrao)
    )
    if item.count() == 0:
        link = pg.locator(f"{LI} a.item.folder").filter(has_text=nome_pasta).first
    else:
        link = item.first.locator("a.item.folder").first
    link.scroll_into_view_if_needed()
    link.click()
    try:
        pg.locator(".folder-open.show").first.wait_for(state="visible", timeout=10000)
    except Exception:
        pass
    esperar_lista(pg)


def _voltar_pasta(pg, nivel: int, nome_pasta: str) -> None:
    pg.locator("a.folder-open-back").first.click()
    if nivel == 0:
        try:
            pg.locator(".folder-open.show").first.wait_for(state="hidden", timeout=10000)
        except Exception:
            pass
    else:
        for _ in range(40):
            titulo = pg.locator(".folder-open.show .folder-open-content p.name")
            if titulo.count() == 0 or titulo.first.inner_text().strip() != nome_pasta:
                break
            pg.wait_for_timeout(250)
    esperar_lista(pg)


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

    # 1) pastas (todas, em sequência, aguardando estabilizar na volta)
    for nome_pasta in pastas:
        try:
            _abrir_subpasta(pg, nome_pasta)
            l, b, e = coletar(pg, ctx, destino / limpa(nome_pasta), log, nivel + 1)
            listados += l
            baixados += b
            erros += e
            _voltar_pasta(pg, nivel, nome_pasta)
        except Exception as ex:
            erros.append(f"pasta {nome_pasta}: {ex}")
            log(f"{ind}  ERRO pasta {nome_pasta}: {ex}")

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

"""Gravador guiado: mapeia seus passos (Origem -> Destino) em passos_mapeados.json.

Uso:  python gravador_guiado.py
Fase 1 apenas. NÃO gera migracao_crm.py até você digitar: todos os passos feitos

Instruções reconhecidas (texto livre, palavras-chave):
  "ir para origem" / "trocar para destino"   -> muda a aba ativa
  "navegar https://..."                       -> abre URL na aba ativa
  "clicar ..."                                -> VOCÊ clica no elemento no navegador; o seletor é capturado
  "colar {nome_cliente} e pressionar Enter"   -> preenche o último campo clicado (Enter se mencionado)
  "digitar 'texto'"                           -> idem, com texto fixo
  "pressionar Enter"                          -> tecla Enter
  "baixar" / "clicar no ícone de download"    -> você clica; registra download (diga "todos" p/ baixar vários)
  "upload" / "anexar arquivo"                 -> você clica no botão de upload; registra envio dos arquivos do cliente
  "esperar 3"                                 -> aguarda N segundos
  "desfazer"                                  -> remove o último passo
  "todos os passos feitos"                    -> encerra o mapeamento
Após um passo de busca com {nome_cliente}, o próximo "clicar" (resultado) é detectado
como LISTA de resultados e recebe automaticamente a regra de homônimos (0 / 1 / 2+).
"""
import json
import os
import re
import unicodedata
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()
ARQUIVO = Path("passos_mapeados.json")
DOWNLOADS = Path("./downloads").resolve()
DOWNLOADS.mkdir(exist_ok=True)

JS_CAPTURA = r"""
(() => {
  if (window.__gravadorInstalado) return;
  window.__gravadorInstalado = true;
  const esc = s => CSS.escape(s);
  const uniq = (sel) => { try { return document.querySelectorAll(sel).length === 1; } catch(e){ return false; } };
  function seletor(el) {
    // sobe até um elemento "clicável" semântico
    const clic = el.closest('button,a,[role=button],[role=link],[role=option],[role=row],input,select,textarea,label,li,tr') || el;
    el = clic;
    const tag = el.tagName.toLowerCase();
    for (const a of ['data-testid','data-test','data-cy','id','name','aria-label','placeholder','title']) {
      const v = el.getAttribute(a);
      if (v && !/\d{4,}/.test(v)) {
        const s = a === 'id' ? '#' + esc(v) : `${tag}[${a}="${v.replace(/"/g,'\\"')}"]`;
        if (uniq(s)) return s;
      }
    }
    const txt = (el.innerText || '').trim().split('\n')[0].slice(0, 40);
    if (txt && ['button','a','label'].includes(tag)) return `${tag}:has-text("${txt.replace(/"/g,'\\"')}")`;
    // caminho estrutural
    const partes = [];
    let n = el;
    while (n && n.nodeType === 1 && n !== document.body) {
      let p = n.tagName.toLowerCase();
      if (n.id && uniq('#' + esc(n.id))) { partes.unshift('#' + esc(n.id)); break; }
      const irm = [...n.parentElement.children].filter(c => c.tagName === n.tagName);
      if (irm.length > 1) p += `:nth-of-type(${irm.indexOf(n) + 1})`;
      partes.unshift(p);
      n = n.parentElement;
    }
    return partes.join(' > ');
  }
  function itemLista(el) {
    // encontra o ancestral que se repete entre irmãos (linha/card de resultado)
    let n = el;
    while (n && n !== document.body) {
      const pai = n.parentElement;
      if (!pai) break;
      const irm = [...pai.children].filter(c => c.tagName === n.tagName);
      if (irm.length >= 2 || (n.tagName === 'TR' )) {
        let cls = [...n.classList].filter(c => !/\d{3,}|active|selected|hover|focus/i.test(c)).slice(0, 2);
        let sel = n.tagName.toLowerCase() + cls.map(c => '.' + esc(c)).join('');
        const role = n.getAttribute('role');
        if (!cls.length && role) sel += `[role="${role}"]`;
        const base = seletor(pai);
        // se for tabela, restringe ao tbody
        return { item: sel, container: base };
      }
      n = pai;
    }
    return null;
  }
  document.addEventListener('click', (e) => {
    const el = e.target;
    let info = { selector: seletor(el), texto: (el.innerText || el.value || '').trim().slice(0, 80), url: location.href };
    info.lista = itemLista(el);
    window.__gravar_clique(JSON.stringify(info));
  }, true);
})();
"""


def norm(t):
    return unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()


class Gravador:
    def __init__(self):
        # retoma passos já mapeados (ex.: passo 1 = fonte_planilha)
        self.passos = json.loads(ARQUIVO.read_text(encoding="utf-8")) if ARQUIVO.exists() else []
        self.ultimo_clique = None  # dict capturado
        self.clique_pendente = None
        self.crm = "origem"
        self.paginas = {}
        self.busca_pendente = False
        self.ultimo_campo = None

    # ---- captura de clique vindo do navegador
    def _on_clique(self, source, payload):
        d = json.loads(payload)
        d["crm"] = next((k for k, p in self.paginas.items() if p == source["page"]), self.crm)
        self.clique_pendente = d

    def esperar_clique(self, ctx, timeout_s=180):
        self.clique_pendente = None
        print("   👉 Faça o clique no navegador agora (estou aguardando)...")
        pg = self.paginas[self.crm]
        for _ in range(timeout_s * 4):
            if self.clique_pendente:
                pg.wait_for_timeout(1200)  # deixa a página reagir
                return self.clique_pendente
            pg.wait_for_timeout(250)
        print("   ⚠️ Tempo esgotado sem clique.")
        return None

    def salvar(self):
        ARQUIVO.write_text(json.dumps(self.passos, ensure_ascii=False, indent=2), encoding="utf-8")

    def add(self, passo):
        passo["crm"] = passo.get("crm", self.crm)
        self.passos.append(passo)
        self.salvar()
        print(f"   ✅ Registrado #{len(self.passos)}: {json.dumps(passo, ensure_ascii=False)}")

    # ---- interpretação
    def tratar(self, ctx, instr):
        n = norm(instr)
        pg = self.paginas[self.crm]

        if n.strip() == "desfazer":
            if self.passos:
                self.passos.pop(); self.salvar(); print("   ↩️ Último passo removido.")
            return

        m = re.search(r"(trocar|ir|mudar|voltar|abrir).*\b(origem|destino)\b", n)
        if m and "http" not in n:
            self.crm = m.group(2)
            self.paginas[self.crm].bring_to_front()
            self.add({"acao": "trocar_aba"})
            return

        m = re.search(r"(https?://\S+)", instr)
        if m:
            self.add({"acao": "goto", "url": m.group(1)})
            pg.goto(m.group(1))
            return

        m = re.search(r"esper\w*\s+(\d+)", n)
        if m:
            s = int(m.group(1)); pg.wait_for_timeout(s * 1000)
            self.add({"acao": "esperar", "segundos": s})
            return

        if re.search(r"\b(colar|digitar|preencher|escrever|pesquisar|buscar)\b", n):
            if "{nome_cliente}" in instr or "nome_cliente" in n:
                valor = "{nome_cliente}"
            else:
                q = re.search(r"[\"'“‘](.+?)[\"'”’]", instr)
                valor = q.group(1) if q else input("   Texto a digitar: ").strip()
            enter = "enter" in n
            if not self.ultimo_campo:
                print("   ℹ️ Nenhum campo clicado antes. Clique no campo agora.")
                c = self.esperar_clique(ctx)
                if not c: return
                self.add({"acao": "click", "seletor": c["selector"], "descricao": "campo de texto"})
                self.ultimo_campo = c["selector"]
            passo = {"acao": "fill", "seletor": self.ultimo_campo, "valor": valor, "enter": enter}
            try:
                pg.fill(self.ultimo_campo, valor.replace("{nome_cliente}", "teste"))
                if enter: pg.keyboard.press("Enter")
                pg.wait_for_timeout(1500)
            except Exception as e:
                print(f"   (aviso ao reproduzir: {e})")
            self.add(passo)
            if valor == "{nome_cliente}":
                self.busca_pendente = True
                print("   🔎 Busca com {nome_cliente} detectada: o próximo clique de resultado ativará a regra de homônimos.")
            return

        if re.fullmatch(r"\s*(pressionar|apertar)?\s*enter\s*", n):
            pg.keyboard.press("Enter")
            self.add({"acao": "press", "tecla": "Enter"})
            return

        if re.search(r"\b(upload|anexar|enviar arquivo|subir)\b", n):
            c = self.esperar_clique(ctx)
            if c:
                self.add({"acao": "upload", "seletor": c["selector"], "pasta": "downloads/{nome_cliente}"})
            return

        if re.search(r"\b(baix\w+|download)\b", n):
            todos = "todos" in n or "todas" in n
            c = self.esperar_clique(ctx)
            if c:
                self.add({"acao": "download", "seletor": c["selector"], "todos": todos,
                          "lista": c["lista"] if todos else None})
            return

        if re.search(r"\b(clicar|clique|selecionar|abrir)\b", n):
            c = self.esperar_clique(ctx)
            if not c: return
            if self.busca_pendente:
                lista = c.get("lista")
                if lista:
                    self.add({"acao": "resultado_busca", "seletor_item": lista["item"],
                              "seletor_container": lista["container"], "seletor_clicado": c["selector"],
                              "texto_exemplo": c["texto"], "regra_homonimos": True})
                else:
                    print("   ⚠️ Não identifiquei lista repetida; registrando clique simples com regra de homônimos "
                          "(ajustável no JSON).")
                    self.add({"acao": "resultado_busca", "seletor_item": c["selector"],
                              "seletor_container": None, "seletor_clicado": c["selector"],
                              "texto_exemplo": c["texto"], "regra_homonimos": True})
                self.busca_pendente = False
            else:
                self.add({"acao": "click", "seletor": c["selector"], "texto": c["texto"]})
                self.ultimo_campo = c["selector"]
            return

        print("   ❓ Não entendi. Use: clicar..., colar {nome_cliente}..., baixar, upload, esperar N, "
              "ir para origem/destino, navegar URL, desfazer.")


def main():
    g = Gravador()
    with sync_playwright() as p:
        cdp = os.getenv("CHROME_CDP_URL")
        if cdp:
            # Anexa ao Chrome já aberto/logado (iniciado com --remote-debugging-port=9222)
            browser = p.chromium.connect_over_cdp(cdp)
            ctx = browser.contexts[0]
        else:
            browser = p.chromium.launch(headless=False)
            ctx = browser.new_context(accept_downloads=True)
        ctx.expose_binding("__gravar_clique", g._on_clique)
        ctx.add_init_script(JS_CAPTURA)
        for nome in ("origem", "destino"):
            pg = ctx.new_page()
            url = os.getenv(f"{nome.upper()}_URL")
            if url:
                try: pg.goto(url)
                except Exception as e: print(f"Aviso ao abrir {nome}: {e}")
            g.paginas[nome] = pg
        print(f"Downloads configurados em: {DOWNLOADS}")
        print("Aba 1 = ORIGEM, Aba 2 = DESTINO. Faça login manualmente (CAPTCHA/2FA) nas duas.")
        input("Pressione ENTER aqui quando estiver logado nos dois CRMs... ")
        g.crm = "origem"; g.paginas["origem"].bring_to_front()

        while True:
            instr = input("\nQual o próximo passo? ").strip()
            if not instr:
                continue
            if instr.lower() == "todos os passos feitos":
                g.salvar()
                print(f"Mapeamento encerrado. {len(g.passos)} passos salvos em {ARQUIVO}.")
                print("Peça ao agente para compilar o migracao_crm.py.")
                break
            try:
                g.tratar(ctx, instr)
            except Exception as e:
                print(f"   ❌ Erro: {e}")
        browser.close()


if __name__ == "__main__":
    main()


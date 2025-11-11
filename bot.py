# bot_brave_attach.py
import re, time
import pandas as pd
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ===== Planilha: B = "1" ; M = "VALOR UNIT. COM BDI" (2ª linha é o header) ====
ARQ_XLSX = "PLANILHA VENCEDORA_094023.xlsx"
ABA      = "PO-PLE"
COL_ID   = "1"                    # coluna B
COL_VAL  = "VALOR UNIT. COM BDI"  # coluna M

ARQ_LOG   = "resultado_atualizacao.csv"
DIR_ERROS = Path("errors"); DIR_ERROS.mkdir(exist_ok=True)

ROW_CANDIDATES = "tr, [role='row'], .row, .MuiTableRow-root, .ant-table-row"
EDIT_QUERY_CHAIN = [
    "[title*='Editar' i], [aria-label*='Editar' i]",
    ".fa-pencil, .icon-pencil, [data-icon='edit'], svg[aria-label*='Editar' i]",
    "button:has(svg), a:has(svg)"
]
SUCCESS_TEXTS = ["Salvo com sucesso", "Atualizado com sucesso"]

# -------------------- Planilha --------------------
def read_pairs(xlsx, sheet, col_id, col_val):
    def pick_col(df, keys, idx_fallback=None):
        norm = {str(c).strip().lower(): c for c in df.columns}
        for k in keys:
            if k in df.columns:
                return df[k]
            if isinstance(k, str) and k.isdigit() and int(k) in df.columns:
                return df[int(k)]
            lk = str(k).strip().lower()
            if lk in norm:
                return df[norm[lk]]
        if idx_fallback is not None and idx_fallback < df.shape[1]:
            return df.iloc[:, idx_fallback]
        raise KeyError(f"Coluna(s) não encontrada(s): {keys}")

    # tenta com a 2ª linha como header
    try:
        df = pd.read_excel(xlsx, sheet_name=sheet, header=1)
        s_id  = pick_col(df,  [col_id, "N° Macrosserviço / Serviço", "Nº Macrosserviço / Serviço", "1"], idx_fallback=1)
        s_val = pick_col(df,  [col_val, "VALOR UNIT. COM BDI", "Preço Unitário (valor calculado).1"], idx_fallback=12)
    except Exception:
        df = pd.read_excel(xlsx, sheet_name=sheet, header=None)
        df = df.iloc[1:].reset_index(drop=True)
        s_id  = df.iloc[:, 1]   # B
        s_val = df.iloc[:, 12]  # M

    def norm_id(v):
        if pd.isna(v): return None
        s = str(v).strip()
        if s.endswith(".0"): s = s[:-2]
        return s or None

    def parse_val_to_float(x):
        if isinstance(x, (int, float)) and pd.notna(x):
            return float(x)
        if pd.isna(x): return None
        s = str(x).strip().replace("R$", "").replace(" ", "")
        s = s.replace(".", "").replace(",", ".")
        try: return float(s)
        except: return None

    def float_to_br(val):
        return f"{val:.2f}".replace(".", ",")

    df2 = pd.DataFrame({"__id": s_id, "__val": s_val})
    df2["_id"]   = df2["__id"].apply(norm_id)
    df2["_fval"] = df2["__val"].apply(parse_val_to_float)
    df2 = df2[df2["_id"].notna() & df2["_fval"].notna()].drop_duplicates("_id", keep="last")

    def nat_key(s):
        parts = re.split(r"(\d+)", s)
        return [int(p) if p.isdigit() else p.lower() for p in parts]

    pares = [(r["_id"], float_to_br(r["_fval"])) for _, r in df2.iterrows()]
    pares.sort(key=lambda kv: nat_key(kv[0]))
    return pares

# -------------------- Navegacao/UI --------------------
def find_first(scope, selectors, timeout_each=1800):
    for sel in selectors:
        loc = scope.locator(sel).first
        if loc.count() > 0:
            try:
                loc.wait_for(state="visible", timeout=timeout_each)
                return loc
            except:
                pass
    return None

def wait_list_ready(page):
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_selector(
        "[title*='Editar' i], [aria-label*='Editar' i], .fa-pencil, .icon-pencil",
        timeout=15000
    )

def locate_row(page, codigo, max_scrolls=60):
    pattern = re.compile(rf"^\s*{re.escape(codigo)}\s*$")
    row = page.locator(ROW_CANDIDATES, has=page.get_by_text(pattern)).first
    if row.count() > 0:
        try:
            row.wait_for(state="visible", timeout=1000)
            return row
        except:
            pass

    # tenta rolar containers
    containers = page.locator("div,section,main,table,tbody")
    for ci in range(min(containers.count(), 25)):
        cont = containers.nth(ci)
        try:
            for _ in range(max_scrolls):
                row = cont.locator(ROW_CANDIDATES, has=page.get_by_text(pattern)).first
                if row.count() > 0:
                    try:
                        row.wait_for(state="visible", timeout=500)
                        return row
                    except:
                        pass
                try:
                    cont.evaluate("(el) => el.scrollBy(0, 800)")
                except:
                    pass
                time.sleep(0.08)
        except:
            continue

    # rola janela
    for _ in range(max_scrolls):
        try: page.mouse.wheel(0, 1000)
        except: page.evaluate("window.scrollBy(0, 1000)")
        time.sleep(0.08)
        row = page.locator(ROW_CANDIDATES, has=page.get_by_text(pattern)).first
        if row.count() > 0:
            try:
                row.wait_for(state="visible", timeout=500)
                return row
            except:
                pass
    return None

def click_edit_on_row(page, codigo):
    row = locate_row(page, codigo)
    if not row:
        raise TimeoutError(f"Linha '{codigo}' não localizada após rolagem.")
    try: row.scroll_into_view_if_needed()
    except: pass
    edit = find_first(row, EDIT_QUERY_CHAIN)
    if not edit:
        edit = page.locator("[title*='Editar' i], [aria-label*='Editar' i], .fa-pencil, .icon-pencil").first
    edit.click()

def find_preco_licitado_input(scope):
    xp = "xpath=//*[contains(normalize-space(.), 'Preço Unitário Licitado')]/following::input[1]"
    campo = scope.locator(xp).first
    if campo.count() > 0:
        try:
            campo.wait_for(state="visible", timeout=1500)
            return campo
        except: pass
    cand = [
        "input[aria-label*='Preço' i][aria-label*='Licitado' i]",
        "[data-testid='campo-preco-unitario'], [data-testid='campo-valor-unitario'], "
        "input[name*='preco' i][name*='licit' i], input[name*='valor' i][name*='licit' i]"
    ]
    return find_first(scope, cand)

def extract_digits(text):
    m = re.search(r"(\d+)", text or "")
    return m.group(1) if m else None

def verify_item_matches(page, codigo):
    """Confere se estamos editando o item certo: '1.2' -> macro=1 e numero=2."""
    try:
        macro_txt  = page.locator("xpath=//*[contains(normalize-space(.),'Macrosserviço Associado')]/following::*[self::div or self::span or self::p][1]").first.text_content(timeout=1200) or ""
        numserv_txt= page.locator("xpath=//*[contains(normalize-space(.),'Número do Serviço')]/following::*[self::div or self::span or self::p][1]").first.text_content(timeout=1200) or ""
    except Exception:
        return False
    parts = codigo.split(".")
    if len(parts) != 2: return False
    macro_expected, num_expected = parts[0], parts[1]
    macro_seen = extract_digits(macro_txt)
    num_seen   = extract_digits(numserv_txt)
    return (macro_seen == macro_expected) and (num_seen == num_expected)

def open_edit_form(page, list_url, codigo, retries=2):
    """Abre a edição do código e garante que é o item correto; se não, volta à lista e tenta novamente."""
    for attempt in range(retries + 1):
        click_edit_on_row(page, codigo)
        page.wait_for_load_state("domcontentloaded")
        if verify_item_matches(page, codigo):
            return page
        # item errado? volta forçando a lista e tenta de novo
        page.goto(list_url, wait_until="domcontentloaded")
        wait_list_ready(page)
    raise RuntimeError(f"A página de edição aberta não corresponde ao item {codigo}.")

def type_exact_money(input_loc, valor_str):
    input_loc.click()
    input_loc.press("Control+A")
    input_loc.press("Delete")
    input_loc.type(valor_str)     # ex.: "87,04"
    time.sleep(0.12)
    try:
        got = input_loc.input_value(timeout=800) or ""
    except Exception:
        got = ""
    # se a máscara alterou, tenta mais uma vez
    if got.strip() != valor_str.strip():
        input_loc.press("Control+A")
        input_loc.press("Delete")
        input_loc.type(valor_str)
        time.sleep(0.12)

def save_and_back_to_list(page, list_url):
    # clica no botão Salvar (qualquer um visível)
    salvar = find_first(page, ["button:has-text('Salvar')", "[data-testid='btn-salvar']", "button[title*='Salvar' i]"])
    if not salvar:
        raise RuntimeError("Botão 'Salvar' não encontrado.")
    salvar.click()
    page.wait_for_load_state("networkidle")
    try:
        page.wait_for_selector("|".join([f"text={t}" for t in SUCCESS_TEXTS]), timeout=7000)
    except PWTimeout:
        pass
    # volta determinístico: força a URL da lista
    page.goto(list_url, wait_until="domcontentloaded")
    wait_list_ready(page)

# -------------------- Logging --------------------
def ensure_log():
    if not Path(ARQ_LOG).exists():
        with open(ARQ_LOG, "w", encoding="utf-8") as f:
            f.write("id;valor;status;mensagem\n")

def log_result(codigo, valor, status, msg):
    with open(ARQ_LOG, "a", encoding="utf-8") as f:
        f.write(f"{codigo};{valor};{status};{msg}\n")

# -------------------- MAIN --------------------
def main():
    pares = read_pairs(ARQ_XLSX, ABA, COL_ID, COL_VAL)
    print(f"Total a atualizar: {len(pares)} itens (começando por 1.1).")
    ensure_log()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        # pega a aba da lista e memoriza a URL da lista
        # (deixe a lista aberta antes de rodar o script)
        # se pegar aba errada, você pode colar diretamente a URL da lista abaixo.
        page = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                page = pg
        page.bring_to_front()
        page.wait_for_load_state("domcontentloaded")
        list_url = page.url  # <- VOLTAMOS SEMPRE PARA ESTA URL

        for codigo, valor in pares:
            status, msg = "OK", "Atualizado"
            try:
                # abre a edição do item correto
                scope = open_edit_form(page, list_url, codigo)
                # campo "Preço Unitário Licitado (R$)"
                campo = find_preco_licitado_input(scope)
                if not campo:
                    raise RuntimeError("Campo 'Preço Unitário Licitado (R$)' não encontrado.")
                # digita exatamente o valor da planilha
                type_exact_money(campo, valor)
                # salva e retorna para a lista correta (sempre pela mesma URL)
                save_and_back_to_list(page, list_url)
                time.sleep(0.12)
            except Exception as e:
                status = "ERRO"
                msg = str(e).replace("\n", " ")[:500]
                ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                png = DIR_ERROS / f"{codigo}-{ts}.png"
                try:
                    page.screenshot(path=str(png), full_page=True)
                    msg += f" (screenshot: {png.name})"
                except:
                    msg += " (screenshot falhou)"
            finally:
                log_result(codigo, valor, status, msg)

        print(f"Fim! Veja '{ARQ_LOG}' e, se houve falhas, a pasta '{DIR_ERROS}'.")
        # não fechamos o Brave

if __name__ == "__main__":
    main()

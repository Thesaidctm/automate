# bot_brave_attach.py
import re, time
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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
        df = pd.read_excel(
            xlsx,
            sheet_name=sheet,
            header=1,
            dtype=str,
            keep_default_na=False,
        )
        s_id  = pick_col(df,  [col_id, "N° Macrosserviço / Serviço", "Nº Macrosserviço / Serviço", "1"], idx_fallback=1)
        s_val = pick_col(df,  [col_val, "VALOR UNIT. COM BDI", "Preço Unitário (valor calculado).1"], idx_fallback=12)
    except Exception:
        df = pd.read_excel(
            xlsx,
            sheet_name=sheet,
            header=None,
            dtype=str,
            keep_default_na=False,
        )
        df = df.iloc[1:].reset_index(drop=True)
        s_id  = df.iloc[:, 1]   # B
        s_val = df.iloc[:, 12]  # M

    def norm_id(v):
        if pd.isna(v): return None
        s = str(v).strip()
        if s.endswith(".0"): s = s[:-2]
        return s or None

    def parse_val_to_decimal(x):
        if pd.isna(x):
            return None

        if isinstance(x, Decimal):
            dec = x
        elif isinstance(x, (int, float)):
            # trata valores numéricos preservando as casas decimais originais
            dec = Decimal(str(x))
        else:
            s = str(x).strip()
            if not s:
                return None

            s = (
                s.replace("R$", "")
                .replace("\u00a0", "")
                .replace("\u202f", "")
                .replace(" ", "")
            )
            s = re.sub(r"[^0-9,.-]", "", s)

            if s.count(",") > 1:
                s = s.replace(",", "")

            if "," in s and "." in s:
                # decide qual é o separador decimal pela última ocorrência
                if s.rfind(",") > s.rfind("."):
                    s = s.replace(".", "")
                    s = s.replace(",", ".")
                else:
                    s = s.replace(",", "")
            elif "," in s:
                s = s.replace(".", "")
                s = s.replace(",", ".")
            elif s.count(".") > 1:
                s = s.replace(".", "")

            try:
                dec = Decimal(s)
            except InvalidOperation:
                return None

        return dec.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    def decimal_to_br(val):
        if val is None:
            return None
        return f"{val:.2f}".replace(".", ",")

    df2 = pd.DataFrame({"__id": s_id, "__val": s_val})
    df2["_id"]   = df2["__id"].apply(norm_id)
    df2["_dval"] = df2["__val"].apply(parse_val_to_decimal)
    df2 = df2[df2["_id"].notna() & df2["_dval"].notna()].drop_duplicates("_id", keep="last")

    def codigo_sort_key(codigo):
        """Ordena códigos hierárquicos como "1.1", "1.10", "1.100" corretamente."""
        if codigo is None:
            return ()

        s = str(codigo).strip()
        if not s:
            return ()

        # Se todos os trechos separados por ponto forem numéricos, tratamos como hierarquia.
        # Assim, "1.100" vem após "1.99" (serviço 100 do macro 1).
        dot_parts = [p for p in s.split(".") if p != ""]
        if dot_parts and all(part.isdigit() for part in dot_parts):
            return tuple(int(part) for part in dot_parts)

        # Fallback: ordenação natural geral (mistura de letras e números, outros separadores).
        parts = re.split(r"(\d+)", s)
        return tuple(int(p) if p.isdigit() else p.lower() for p in parts if p)

    pares = [(r["_id"], decimal_to_br(r["_dval"])) for _, r in df2.iterrows()]
    pares.sort(key=lambda kv: codigo_sort_key(kv[0]))
    return pares

# -------------------- Utilidades de valor monetário --------------------
def parse_money_decimal(texto):
    if texto is None:
        return None

    if isinstance(texto, Decimal):
        try:
            return texto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except Exception:
            return None

    if isinstance(texto, (int, float)):
        texto = str(texto)

    s = str(texto).strip()
    if not s:
        return None

    s = re.sub(r"(?i)r\$", "", s)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^0-9,.-]", "", s)

    if not s:
        return None

    if s.count(",") > 1:
        s = s.replace(",", "")

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")

    try:
        return Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return None

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
    try:
        page.wait_for_load_state("domcontentloaded")
    except Exception:
        pass

    try:
        page.wait_for_load_state("networkidle")
    except Exception:
        pass

    page.wait_for_selector(
        "[title*='Editar' i], [aria-label*='Editar' i], .fa-pencil, .icon-pencil",
        timeout=20000
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

def read_row_text(row):
    try:
        return row.inner_text(timeout=1200) or ""
    except Exception:
        try:
            return row.text_content(timeout=1200) or ""
        except Exception:
            return ""


def row_contains_price(row_text, alvo_decimal):
    if not row_text:
        return False

    partes = [p.strip() for p in re.split(r"[\n\t]+", row_text) if p.strip()]
    for trecho in partes:
        visto = parse_money_decimal(trecho)
        if visto is not None and visto == alvo_decimal:
            return True
    return False


def row_needs_update(page, codigo, valor):
    row = locate_row(page, codigo)
    if not row:
        raise TimeoutError(f"Linha '{codigo}' não localizada após rolagem.")
    try:
        row.scroll_into_view_if_needed()
    except Exception:
        pass

    alvo_decimal = parse_money_decimal(valor)
    if alvo_decimal is None:
        raise ValueError(f"Valor '{valor}' não pôde ser interpretado como monetário.")

    texto = read_row_text(row)
    if row_contains_price(texto, alvo_decimal):
        return False, alvo_decimal
    return True, alvo_decimal


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
    xp = (
        "xpath="
        "(//*[contains(normalize-space(.), 'Preço Unitário Licitado') or "
        "contains(normalize-space(.), 'Preço unitário licitado')]/following::input)[1]"
    )
    campo = scope.locator(xp).first
    if campo.count() > 0:
        try:
            campo.wait_for(state="visible", timeout=1500)
            return campo
        except:
            pass
    cand = [
        "input[aria-label*='Preço' i][aria-label*='Licitado' i]",
        "input[aria-label*='Preço Unitário Licitado' i]",
        "[data-testid='campo-preco-unitario-licitado'], [data-testid='campo-valor-unitario-licitado']",
        "input[name*='preco' i][name*='licitado' i]",
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

def type_exact_money(input_loc, valor_str, alvo_decimal=None):
    bruto = (valor_str or "").strip()
    alvo = re.sub(r"(?i)r\$", "", bruto)
    alvo = re.sub(r"\s+", "", alvo)
    alvo = re.sub(r"[^0-9,.-]", "", alvo)
    if not alvo:
        raise ValueError("Valor vazio fornecido para preenchimento do preço.")

    if alvo_decimal is None:
        alvo_decimal = parse_money_decimal(alvo)
    if alvo_decimal is None:
        raise ValueError(f"Valor '{valor_str}' não pôde ser interpretado como monetário.")

    alvo_digitos = re.sub(r"\D", "", alvo)
    if not alvo_digitos and alvo_decimal is not None:
        alvo_digitos = f"{int((alvo_decimal * 100).quantize(Decimal('1')))}"

    def ler_input():
        try:
            return input_loc.input_value(timeout=800) or ""
        except Exception:
            return ""

    def confere(obtido):
        visto = parse_money_decimal(obtido)
        return visto is not None and visto == alvo_decimal

    tentativas = [
        ("type_digits", alvo_digitos or re.sub(r"\D", "", alvo) or alvo),
        ("fill", alvo),
    ]

    for modo, texto in tentativas:
        input_loc.click()
        input_loc.press("Control+A")
        input_loc.press("Delete")
        if modo == "fill":
            input_loc.fill(texto)
        else:
            input_loc.type(texto, delay=40)
        time.sleep(0.2)
        obtido = ler_input().strip()
        if confere(obtido):
            return

    obtido = ler_input().strip()
    raise RuntimeError(
        f"Não foi possível definir o valor '{valor_str}' (campo ficou '{obtido or '[vazio]'}')."
    )


def field_needs_update(input_loc, valor_str):
    alvo_decimal = parse_money_decimal(valor_str)
    if alvo_decimal is None:
        raise ValueError(f"Valor '{valor_str}' não pôde ser interpretado como monetário.")

    try:
        atual = input_loc.input_value(timeout=800) or ""
    except Exception:
        try:
            atual = input_loc.evaluate("el => el.value") or ""
        except Exception:
            atual = ""

    atual_decimal = parse_money_decimal(atual)
    if atual_decimal is not None and atual_decimal == alvo_decimal:
        return False, alvo_decimal
    return True, alvo_decimal


def wait_success_feedback(page, timeout=7000):
    deadline = time.time() + (timeout / 1000)
    for txt in SUCCESS_TEXTS:
        restante = max(0, int((deadline - time.time()) * 1000))
        if restante <= 0:
            break
        loc = page.locator(f"text={txt}")
        try:
            loc.wait_for(state="visible", timeout=restante)
            return
        except Exception:
            continue


def save_and_return_to_list(page, list_url, edit_url):
    salvar = find_first(page, ["button:has-text('Salvar')", "[data-testid='btn-salvar']", "button[title*='Salvar' i]"])
    if not salvar:
        raise RuntimeError("Botão 'Salvar' não encontrado.")
    salvar.click()
    page.wait_for_load_state("networkidle")
    wait_success_feedback(page)

    # aguarda a navegação natural do sistema; se não sair da página de edição,
    # voltamos para a URL da lista como último recurso.
    limite = time.time() + 15
    while time.time() < limite:
        if page.url != edit_url:
            try:
                wait_list_ready(page)
                return
            except Exception:
                pass
        time.sleep(0.25)

    if page.url != list_url:
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
        wait_list_ready(page)
        list_url = page.url  # URL atual da lista de itens

        for codigo, valor in pares:
            status, msg = "OK", "Atualizado"
            try:
                precisa_editar, alvo_decimal = row_needs_update(page, codigo, valor)
                if not precisa_editar:
                    msg = "Sem alteração (valor já corresponde à planilha na lista)"
                    continue

                # abre a edição do item correto
                scope = open_edit_form(page, list_url, codigo)
                edit_url = page.url
                # campo "Preço Unitário Licitado (R$)"
                campo = find_preco_licitado_input(scope)
                if not campo:
                    raise RuntimeError("Campo 'Preço Unitário Licitado (R$)' não encontrado.")
                precisa_editar_form, alvo_decimal_form = field_needs_update(campo, valor)
                if not precisa_editar_form:
                    msg = "Sem alteração (valor já corresponde à planilha na edição)"
                    save_and_return_to_list(page, list_url, edit_url)
                    list_url = page.url
                    continue

                alvo_decimal = alvo_decimal_form
                # digita exatamente o valor da planilha
                type_exact_money(campo, valor, alvo_decimal=alvo_decimal)
                # salva e aguarda a volta para a lista dos itens
                save_and_return_to_list(page, list_url, edit_url)
                list_url = page.url
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

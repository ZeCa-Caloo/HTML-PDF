import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile, os, sys, io, re
from html import unescape

st.set_page_config(page_title="Converter HTML/XLS(X) ➜ PDF", page_icon="🧾", layout="centered")
st.title("🧾 Converter HTML/XLS(X) ➜ PDF")
st.caption("Envie arquivos .html, .htm, .xls ou .xlsx e receba um PDF.")

# -----------------------
# Info de ambiente
# -----------------------
st.caption(f"Python em uso: {sys.executable}")
try:
    import xhtml2pdf  # apenas para exibir versão
    st.caption(f"xhtml2pdf: OK ({getattr(xhtml2pdf, '__version__', 'versão desconhecida')})")
except Exception:
    st.warning("xhtml2pdf não encontrado neste Python. Instale com:  python -m pip install xhtml2pdf")

# -----------------------
# Configurações (Sidebar)
# -----------------------
engine = st.sidebar.selectbox("Motor de PDF", ["WeasyPrint (preservar layout)", "xhtml2pdf (compat)"], index=0)
preserve_layout = st.sidebar.checkbox("Preservar layout do HTML (usar CSS do documento)", True)
page_size = st.sidebar.selectbox("Tamanho da página (se NÃO preservar layout)", ["A4", "Letter"], index=0)
orientation = st.sidebar.selectbox("Orientação (se NÃO preservar layout)", ["portrait", "landscape"], index=0)

# Margens laterais simétricas
margin_mm = st.sidebar.slider("Margem lateral (mm) – esquerda = direita", 5, 25, 10)

paginate_sheets = st.sidebar.checkbox("Quebrar página entre planilhas (Excel)", True)

# Somente para xhtml2pdf (não preserva CSS moderno)
sanitize = st.sidebar.checkbox("Sanitizar CSS (apenas xhtml2pdf)", True)

uploaded = st.file_uploader("Envie um arquivo .html, .htm, .xls ou .xlsx",
                            type=["html", "htm", "xls", "xlsx"])

# -----------------------
# Sanitização (para xhtml2pdf)
# -----------------------
UNSUPPORTED_AT_RULES = ("@media", "@supports", "@keyframes", "@-webkit-", "@-moz-", "@-ms-")

def strip_unsupported_at_rules(css: str) -> str:
    out, i = [], 0
    while i < len(css):
        if css[i] == "@" and any(css.startswith(x, i) for x in UNSUPPORTED_AT_RULES):
            depth, j = 0, i
            while j < len(css):
                if css[j] == "{":
                    depth += 1
                elif css[j] == "}":
                    depth -= 1
                    if depth <= 0:
                        j += 1
                        break
                j += 1
            i = j
            continue
        out.append(css[i]); i += 1
    return "".join(out)

def sanitize_selectors(css: str) -> str:
    css = re.sub(r"::[a-zA-Z0-9_-]+", "", css)
    css = re.sub(r":[a-zA-Z-]+\([^)]*\)", "", css)
    css = re.sub(r":[a-zA-Z-]+", "", css)
    css = css.replace("~", " ").replace(">", " ").replace("+", " ")
    return css

def normalize_display_props(css: str) -> str:
    patterns = [
        r"display\s*:\s*inline-flex[^;]*;", r"display\s*:\s*inline-grid[^;]*;",
        r"display\s*:\s*flex[^;]*;", r"display\s*:\s*grid[^;]*;", r"display\s*:\s*contents[^;]*;",
    ]
    for p in patterns:
        css = re.sub(p, "display:block;", css, flags=re.IGNORECASE)
    return css

def neutralize_css_functions(css: str) -> str:
    css = re.sub(r"var\(\s*--[^)]+\)", "", css, flags=re.IGNORECASE)
    css = re.sub(r"calc\([^)]+\)", "1", css, flags=re.IGNORECASE)
    return css

def strip_unsupported_props(css: str) -> str:
    props = [
        r"position\s*:\s*fixed[^;]*;", r"position\s*:\s*absolute[^;]*;",
        r"backdrop-filter\s*:[^;]*;", r"filter\s*:[^;]*;", r"box-shadow\s*:[^;]*;",
        r"transform\s*:[^;]*;", r"transition\s*:[^;]*;", r"animation\s*:[^;]*;", r"@font-face\s*{[^}]*}",
    ]
    for p in props:
        css = re.sub(p, "", css, flags=re.IGNORECASE)
    return css

def sanitize_css(css: str) -> str:
    css = unescape(css)
    css = strip_unsupported_at_rules(css)
    css = sanitize_selectors(css)
    css = normalize_display_props(css)
    css = strip_unsupported_props(css)
    css = neutralize_css_functions(css)
    return css

def sanitize_html_for_xhtml2pdf(html: str, page_css: str) -> str:
    html = unescape(html)
    # injeta @page (xhtml2pdf precisa)
    page_block = f"<style>{page_css}</style>"
    lower = html.lower()
    if "</head>" in lower:
        idx = lower.rfind("</head>")
        html = html[:idx] + page_block + html[idx:]
    else:
        html = f"<html><head><meta charset='utf-8'>{page_block}</head><body>{html}</body></html>"

    # blocos <style>
    def _clean_style_block(m):
        raw = m.group(1)
        return "<style>" + sanitize_css(raw) + "</style>"
    html = re.sub(r"<style[^>]*>(.*?)</style>", _clean_style_block, html, flags=re.IGNORECASE|re.DOTALL)

    # inline style básico
    def _clean_inline_style(m):
        raw = neutralize_css_functions(m.group(1))
        allowed = []
        for decl in raw.split(";"):
            d = decl.strip()
            if not d:
                continue
            if any(d.lower().startswith(x) for x in [
                "color:", "background-color:", "font-size:", "font-family:",
                "border", "padding", "margin", "text-align:", "width:", "height:"
            ]):
                d = neutralize_css_functions(d)
                allowed.append(d)
        return ' style="' + "; ".join(allowed) + ('"' if allowed else '"')
    html = re.sub(r'\sstyle="(.*?)"', _clean_inline_style, html, flags=re.IGNORECASE|re.DOTALL)
    return html

def _inject_page_css(html_str: str, page_css: str) -> str:
    lower = html_str.lower()
    block = f"<style>{page_css}</style>"
    if "</head>" in lower:
        idx = lower.rfind("</head>")
        return html_str[:idx] + block + html_str[idx:]
    return f"<html><head><meta charset='utf-8'>{block}</head><body>{html_str}</body></html>"

# -----------------------
# Monkey-patch xhtml2pdf.parser.lower() para evitar crash
# -----------------------
def _patch_xhtml2pdf_lower():
    try:
        import xhtml2pdf.parser as _p
    except Exception:
        return
    def _safe_lower(seq):
        if isinstance(seq, (list, tuple)) and seq:
            seq = seq[0]
        if seq is None or seq is NotImplemented:
            return ""
        try:
            return str(seq).lower()
        except Exception:
            return ""
    _p.lower = _safe_lower

# -----------------------
# Leitura do HTML + base_url (para preservar caminhos relativos)
# -----------------------
def read_html_and_base(uploaded_file):
    raw = uploaded_file.read()
    try:
        html_str = raw.decode("utf-8")
    except UnicodeDecodeError:
        html_str = raw.decode("latin-1", errors="ignore")

    tmpdir = tempfile.mkdtemp(prefix="html2pdf_")
    fname = Path(uploaded_file.name).name
    fpath = os.path.join(tmpdir, fname)
    with open(fpath, "wb") as f:
        f.write(raw)
    base_url = tmpdir  # permite que <img src="..."> relativos funcionem
    return html_str, base_url

# -----------------------
# Helpers para fallback de fontes/emoji (WeasyPrint)
# -----------------------
def _strip_emojis(text: str) -> str:
    # Remove emojis / pictogramas que quebram o HarfBuzz em alguns setups Windows
    emoji_ranges = [
        (0x1F600, 0x1F64F),  # Emoticons
        (0x1F300, 0x1F5FF),  # Símbolos & pictogramas
        (0x1F680, 0x1F6FF),  # Transporte & mapas
        (0x2600,  0x26FF),   # Diversos
        (0x2700,  0x27BF),   # Dingbats
        (0xFE00,  0xFE0F),   # Variation Selectors
        (0x1F900, 0x1F9FF),  # Suplemento de pictogramas
        (0x1FA70, 0x1FAFF),  # Símbolos adicionais
        (0x1F1E6, 0x1F1FF),  # Bandeiras (pares region)
    ]
    out_chars = []
    for ch in text:
        cp = ord(ch)
        if any(start <= cp <= end for start, end in emoji_ranges):
            continue
        out_chars.append(ch)
    return "".join(out_chars)

# -----------------------
# Builders
# -----------------------
def build_pdf_weasy(html_str: str, base_url: str) -> bytes:
    try:
        from weasyprint import HTML, CSS
        try:
            # WeasyPrint ≥ 60
            from weasyprint.fonts import FontConfiguration
        except Exception:
            # WeasyPrint 53.x
            from weasyprint.text.fonts import FontConfiguration
    except Exception:
        st.error("WeasyPrint não está instalado.\nTente:  pip install weasyprint  (ou em conda:  conda install -c conda-forge weasyprint)")
        st.stop()

    # Garante charset
    if "<meta charset" not in html_str.lower():
        if "<head>" in html_str.lower():
            html_str = html_str.replace("<head>", "<head><meta charset='utf-8'>", 1)
        else:
            html_str = f"<html><head><meta charset='utf-8'></head><body>{html_str}</body></html>"

    font_config = FontConfiguration()

    # CSS de margens laterais iguais (não altera o restante do layout)
    if preserve_layout:
        page_css = CSS(string=f"""
            @page {{
                margin-left: {margin_mm}mm;
                margin-right: {margin_mm}mm;
            }}
        """, font_config=font_config)
    else:
        page_css = CSS(string=f"""
            @page {{
                size: {page_size} {orientation};
                margin-left: {margin_mm}mm;
                margin-right: {margin_mm}mm;
            }}
        """, font_config=font_config)

    # “Cinto de segurança” contra vazamento lateral, sem mexer no design
    safety_css = CSS(string="""
        html, body { overflow: visible !important; }
        * { box-sizing: border-box; min-width: 0 !important; }
        img, svg, canvas, video { max-width: 100% !important; height: auto !important; }
        table { width: 100% !important; table-layout: fixed !important; border-collapse: collapse; }
        td, th { word-break: break-word; }
        pre, code { white-space: pre-wrap; word-break: break-word; }
    """, font_config=font_config)

    styles = [page_css, safety_css]

    # 1ª tentativa: render normal (preserva layout + margens simétricas)
    try:
        pdf_bytes = HTML(string=html_str, base_url=base_url or ".").write_pdf(
            stylesheets=styles, font_config=font_config
        )
        if pdf_bytes is None:
            raise RuntimeError("WeasyPrint não retornou bytes do PDF.")
        return pdf_bytes
    except Exception as e1:
        # 2ª tentativa: força fonte segura e remove emojis problemáticos
        # (instale fontes pelo conda, se precisar: dejavu-fonts-ttf, liberation-fonts)
        fallback_font_css = CSS(string="""
            html, body, * {
                font-family: "DejaVu Sans", "Liberation Sans", Arial, sans-serif !important;
                font-variant-ligatures: none;
            }
        """, font_config=font_config)
        styles2 = [page_css, safety_css, fallback_font_css]

        safe_html = _strip_emojis(html_str)

        try:
            pdf_bytes = HTML(string=safe_html, base_url=base_url or ".").write_pdf(
                stylesheets=styles2, font_config=font_config
            )
            if pdf_bytes is None:
                raise RuntimeError("WeasyPrint não retornou bytes do PDF (fallback).")
            return pdf_bytes
        except Exception as e2:
            st.error("Falha ao gerar PDF no WeasyPrint (mesmo no modo de fallback de fontes).")
            st.exception(e1)  # erro original
            st.exception(e2)  # erro do fallback
            st.stop()

def build_pdf_xhtml2pdf(html_str: str) -> bytes:
    try:
        from xhtml2pdf import pisa
    except Exception:
        st.error("xhtml2pdf não está instalado. Rode:  python -m pip install xhtml2pdf")
        st.stop()

    _patch_xhtml2pdf_lower()

    # @page com margens laterais iguais também para xhtml2pdf
    page_css = f"@page {{ margin-left: {margin_mm}mm; margin-right: {margin_mm}mm; }}"
    html2 = sanitize_html_for_xhtml2pdf(html_str, page_css) if sanitize or preserve_layout else _inject_page_css(html_str, page_css)

    if "<meta charset" not in html2.lower():
        if "<head>" in html2.lower():
            html2 = html2.replace("<head>", "<head><meta charset='utf-8'>", 1)
        else:
            html2 = f"<html><head><meta charset='utf-8'></head><body>{html2}</body></html>"

    out = io.BytesIO()
    res = pisa.CreatePDF(src=html2, dest=out, encoding="utf-8")
    if res.err:
        raise RuntimeError("xhtml2pdf falhou após sanitização.")
    return out.getvalue()

def excel_to_html(uploaded_file, break_between=True) -> str:
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.exception(e); st.stop()

    parts = []
    for i, sheet in enumerate(xls.sheet_names):
        df = xls.parse(sheet)
        styled = (df.style
                  .set_table_attributes('border="1" cellspacing="0" cellpadding="6"')
                  .set_properties(**{"font-family": "Arial", "font-size": "12px"}))
        br = 'style="page-break-before: always;"' if (break_between and i > 0) else ""
        parts.append(f'<h2 {br}>Planilha: {sheet}</h2>' + styled.to_html())

    return f"""
    <html><head><meta charset="utf-8">
    <style>
      body {{ margin:24px; }}
      h2 {{ font-family: Arial, sans-serif; }}
      table {{ border-collapse: collapse; width: 100%; table-layout: fixed; }}
      th, td {{ border: 1px solid #999; padding: 6px; word-wrap: break-word; }}
      pre, code {{ white-space: pre-wrap; word-wrap: break-word; }}
    </style></head>
    <body>{''.join(parts)}</body></html>
    """

def html_file_to_str(uploaded_file) -> str:
    raw = uploaded_file.read()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1", errors="ignore")

# -----------------------
# Fluxo principal
# -----------------------
if uploaded:
    ext = Path(uploaded.name).suffix.lower()

    if ext in [".html", ".htm"]:
        html_str, base_url = read_html_and_base(uploaded)
        if engine.startswith("WeasyPrint"):
            pdf_bytes = build_pdf_weasy(html_str, base_url)
        else:
            if preserve_layout:
                st.info("Para preservar 100% do layout/CSS do documento, prefira o motor WeasyPrint.")
            pdf_bytes = build_pdf_xhtml2pdf(html_str)

        st.success("HTML convertido com sucesso.")
        st.download_button("⬇️ Baixar PDF", data=pdf_bytes,
                           file_name=f"{Path(uploaded.name).stem}.pdf", mime="application/pdf")

    elif ext in [".xls", ".xlsx"]:
        html_doc = excel_to_html(uploaded, break_between=paginate_sheets)
        # Excel vira HTML simples; WeasyPrint preserva, xhtml2pdf usa @page
        if engine.startswith("WeasyPrint"):
            pdf_bytes = build_pdf_weasy(html_doc, base_url=".")
        else:
            pdf_bytes = build_pdf_xhtml2pdf(html_doc)

        st.success("Excel convertido com sucesso.")
        st.download_button("⬇️ Baixar PDF", data=pdf_bytes,
                           file_name=f"{Path(uploaded.name).stem}.pdf", mime="application/pdf")

    else:
        st.warning("Formato não suportado. Envie .html, .htm, .xls ou .xlsx.")
else:
    st.info("Envie um arquivo para iniciar a conversão.")



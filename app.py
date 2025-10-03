import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile, os, sys, io, re, base64
from html import unescape
from PIL import Image  # para imagens -> PDF

st.set_page_config(page_title="Converter para PDF", page_icon="🧾", layout="centered")
st.title("🧾 Converter HTML/XLS(X)/DOCX/Imagens ➜ PDF")
st.caption("Envie .html, .htm, .xls, .xlsx, .docx ou imagens (jpg/png/gif/bmp/tiff/webp/svg) e receba um PDF. "
           "Agora também é possível enviar vários arquivos e gerar um único PDF.")

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
# ✅ Ajuste: xhtml2pdf como padrão (index=1)
engine = st.sidebar.selectbox("Motor de PDF", ["WeasyPrint (preservar layout)", "xhtml2pdf (compat)"], index=1)
preserve_layout = st.sidebar.checkbox("Preservar layout do HTML (usar CSS do documento)", True)
page_size = st.sidebar.selectbox("Tamanho da página (se NÃO preservar layout)", ["A4", "Letter"], index=0)
orientation = st.sidebar.selectbox("Orientação (se NÃO preservar layout)", ["portrait", "landscape"], index=0)
margin_mm = st.sidebar.slider("Margem lateral (mm) – esquerda = direita", 5, 25, 10)
paginate_sheets = st.sidebar.checkbox("Quebrar página entre planilhas (Excel)", True)

# Unir múltiplos
combine_all = st.sidebar.checkbox("Unir todos os arquivos em um único PDF", True)

# Somente para xhtml2pdf (não preserva CSS moderno)
sanitize = st.sidebar.checkbox("Sanitizar CSS (apenas xhtml2pdf)", True)

uploaded_files = st.file_uploader(
    "Envie um ou mais arquivos .html, .htm, .xls, .xlsx, .docx ou imagem (jpg/png/gif/bmp/tiff/webp/svg)",
    type=["html", "htm", "xls", "xlsx", "docx", "jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp", "svg"],
    accept_multiple_files=True
)

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
    page_block = f"<style>{page_css}</style>"
    lower = html.lower()
    if "</head>" in lower:
        idx = lower.rfind("</head>")
        html = html[:idx] + page_block + html[idx:]
    else:
        html = f"<html><head><meta charset='utf-8'>{page_block}</head><body>{html}</body></html>"

    def _clean_style_block(m):
        raw = m.group(1)
        return "<style>" + sanitize_css(raw) + "</style>"
    html = re.sub(r"<style[^>]*>(.*?)</style>", _clean_style_block, html, flags=re.IGNORECASE|re.DOTALL)

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
    uploaded_file.seek(0)
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
    base_url = tmpdir
    return html_str, base_url

# -----------------------
# Helpers (DOCX/Imagens)
# -----------------------
def _img_to_data_uri(image):
    """Usado pelo Mammoth (Python) para embutir imagens do DOCX como data URI."""
    with image.open() as img_bytes:
        encoded = base64.b64encode(img_bytes.read()).decode("ascii")
    return {"src": f"data:{image.content_type};base64,{encoded}"}

def docx_to_html(uploaded_file) -> str:
    """Converte .docx em HTML com imagens embutidas (API correta do Mammoth em Python)."""
    try:
        import mammoth
    except Exception:
        st.error("Pacote 'mammoth' não está instalado. Adicione 'mammoth' ao requirements.txt.")
        st.stop()

    uploaded_file.seek(0)
    raw = uploaded_file.read()
    with io.BytesIO(raw) as f:
        result = mammoth.convert_to_html(
            f,
            convert_image=mammoth.images.img_element(_img_to_data_uri)
        )
    html = result.value
    return f"<html><head><meta charset='utf-8'></head><body>{html}</body></html>"

def image_file_to_html(uploaded_file) -> str:
    """Gera HTML contendo a imagem embutida (data URI), max-width 100%."""
    uploaded_file.seek(0)
    raw = uploaded_file.read()
    try:
        im = Image.open(io.BytesIO(raw))
        if im.mode in ("P", "LA", "RGBA", "CMYK"):
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        raw = buf.getvalue()
        mime = "image/png"
    except Exception:
        ext = Path(uploaded_file.name).suffix.lower().lstrip(".")
        mime = f"image/{'jpeg' if ext in ['jpg','jpeg'] else ext}"
    b64 = base64.b64encode(raw).decode("ascii")
    data_uri = f"data:{mime};base64,{b64}"
    html = f"""
    <html><head><meta charset="utf-8">
      <style>
        html,body{{margin:0;padding:0}}
        .wrap{{padding:0; margin:0 auto;}}
        img{{display:block; max-width:100%; height:auto; margin:0 auto;}}
      </style>
    </head>
    <body><div class="wrap"><img src="{data_uri}"/></div></body></html>
    """
    return html

# -----------------------
# Builders (PDF)
# -----------------------
def build_pdf_weasy(html_str: str, base_url: str) -> bytes:
    try:
        from weasyprint import HTML, CSS
        try:
            from weasyprint.fonts import FontConfiguration  # >=60
        except Exception:
            from weasyprint.text.fonts import FontConfiguration  # 53.x
    except Exception:
        st.error("WeasyPrint não está instalado.\nTente: pip install weasyprint (ou conda-forge).")
        st.stop()

    if "<meta charset" not in html_str.lower():
        if "<head>" in html_str.lower():
            html_str = html_str.replace("<head>", "<head><meta charset='utf-8'>", 1)
        else:
            html_str = f"<html><head><meta charset='utf-8'></head><body>{html_str}</body></html>"

    font_config = FontConfiguration()

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

    safety_css = CSS(string="""
        html, body { overflow: visible !important; }
        * { box-sizing: border-box; min-width: 0 !important; }
        img, svg, canvas, video { max-width: 100% !important; height: auto !important; }
        table { width: 100% !important; table-layout: fixed !important; border-collapse: collapse; }
        td, th { word-break: break-word; }
        pre, code { white-space: pre-wrap; word-break: break-word; }
    """, font_config=font_config)

    styles = [page_css, safety_css]

    # Primeira tentativa normal
    try:
        pdf_bytes = HTML(string=html_str, base_url=base_url or ".").write_pdf(
            stylesheets=styles, font_config=font_config
        )
        if pdf_bytes is None:
            raise RuntimeError("WeasyPrint não retornou bytes do PDF.")
        return pdf_bytes
    except Exception:
        # Fallback de fontes/emoji
        def _strip_emojis(text: str) -> str:
            ranges = [(0x1F600,0x1F64F),(0x1F300,0x1F5FF),(0x1F680,0x1F6FF),(0x2600,0x26FF),
                      (0x2700,0x27BF),(0xFE00,0xFE0F),(0x1F900,0x1F9FF),(0x1FA70,0x1FAFF),(0x1F1E6,0x1F1FF)]
            out=[]
            for ch in text:
                cp=ord(ch)
                if any(a<=cp<=b for a,b in ranges): continue
                out.append(ch)
            return "".join(out)

        fallback_font_css = CSS(string="""
            html, body, * { font-family: "DejaVu Sans", "Liberation Sans", Arial, sans-serif !important; font-variant-ligatures: none; }
        """, font_config=font_config)
        safe_html = _strip_emojis(html_str)
        pdf_bytes = HTML(string=safe_html, base_url=base_url or ".").write_pdf(
            stylesheets=[page_css, safety_css, fallback_font_css], font_config=font_config
        )
        if pdf_bytes is None:
            raise RuntimeError("WeasyPrint não retornou bytes do PDF (fallback).")
        return pdf_bytes

def build_pdf_xhtml2pdf(html_str: str) -> bytes:
    try:
        from xhtml2pdf import pisa
    except Exception:
        st.error("xhtml2pdf não está instalado. Rode:  python -m pip install xhtml2pdf")
        st.stop()

    _patch_xhtml2pdf_lower()

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

def convert_html_to_pdf(html_str: str, base_url: str = ".") -> bytes:
    if engine.startswith("WeasyPrint"):
        return build_pdf_weasy(html_str, base_url)
    else:
        return build_pdf_xhtml2pdf(html_str)

# -----------------------
# Excel -> HTML (com engines corretos)
# -----------------------
def excel_to_html(uploaded_file, break_between=True) -> str:
    """Converte .xls/.xlsx para HTML usando o engine correto e checando dependências."""
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    data = uploaded_file.read()
    bio = io.BytesIO(data)

    xls_engine = None
    if name.endswith(".xlsx"):
        xls_engine = "openpyxl"
        try:
            import openpyxl  # noqa
        except Exception:
            st.error("Falta a dependência 'openpyxl' para ler .xlsx. Adicione 'openpyxl' ao requirements.txt e reinstale.")
            st.stop()
    elif name.endswith(".xls"):
        xls_engine = "xlrd"
        try:
            import xlrd  # noqa
        except Exception:
            st.error("Falta a dependência 'xlrd' para ler .xls. Adicione 'xlrd' ao requirements.txt (>=2.0) e reinstale.")
            st.stop()

    try:
        xls = pd.ExcelFile(bio, engine=xls_engine) if xls_engine else pd.ExcelFile(bio)
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
      th, td {{ border: 1px solid {"#"}999; padding: 6px; word-wrap: break-word; }}
      pre, code {{ white-space: pre-wrap; word-wrap: break-word; }}
    </style></head>
    <body>{''.join(parts)}</body></html>
    """

def html_file_to_str(uploaded_file) -> str:
    uploaded_file.seek(0)
    raw = uploaded_file.read()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1", errors="ignore")

# -----------------------
# Merge de múltiplos PDFs
# -----------------------
def merge_pdfs(pdf_bytes_list: list[bytes]) -> bytes:
    """Une vários PDFs (bytes) em um único PDF."""
    writer = None
    # Tenta pypdf primeiro
    try:
        from pypdf import PdfReader, PdfWriter
        writer = PdfWriter()
    except Exception:
        try:
            from PyPDF2 import PdfReader, PdfWriter
            writer = PdfWriter()
        except Exception:
            st.error("Para unir PDFs, instale 'pypdf' ou 'PyPDF2' no requirements.txt.")
            st.stop()

    for pdf_bytes in pdf_bytes_list:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.getvalue()

# -----------------------
# Processamento por arquivo
# -----------------------
def convert_uploaded_file_to_pdf_bytes(file) -> bytes:
    ext = Path(file.name).suffix.lower()

    if ext in [".html", ".htm"]:
        html_str, base_url = read_html_and_base(file)
        return convert_html_to_pdf(html_str, base_url)

    elif ext in [".xls", ".xlsx"]:
        html_doc = excel_to_html(file, break_between=paginate_sheets)
        return convert_html_to_pdf(html_doc, base_url=".")

    elif ext == ".docx":
        html_doc = docx_to_html(file)
        return convert_html_to_pdf(html_doc, base_url=".")  # imagens embutidas

    elif ext in [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg"]:
        html_doc = image_file_to_html(file)
        return convert_html_to_pdf(html_doc, base_url=".")

    else:
        raise ValueError(f"Formato não suportado: {ext}")

# -----------------------
# Fluxo principal
# -----------------------
if uploaded_files:
    pdfs = []
    errors = []

    for f in uploaded_files:
        try:
            pdf_bytes = convert_uploaded_file_to_pdf_bytes(f)
            pdfs.append((f.name, pdf_bytes))
        except Exception as e:
            errors.append((f.name, e))

    # mostra erros (se houver) mas segue com os válidos
    if errors:
        for name, e in errors:
            with st.expander(f"⚠️ Falha ao converter: {name}"):
                st.exception(e)

    if not pdfs:
        st.warning("Nenhum arquivo pôde ser convertido.")
        st.stop()

    if combine_all and len(pdfs) > 1:
        merged_bytes = merge_pdfs([b for _, b in pdfs])
        st.success(f"Convertidos {len(pdfs)} arquivos e unidos em um único PDF.")
        st.download_button("⬇️ Baixar PDF único", data=merged_bytes,
                           file_name="documentos_unificados.pdf", mime="application/pdf")
    else:
        st.success(f"Convertidos {len(pdfs)} arquivo(s). Baixe individualmente abaixo.")
        for idx, (name, b) in enumerate(pdfs, start=1):
            st.download_button(f"⬇️ Baixar {idx}: {name}.pdf", data=b,
                               file_name=f"{Path(name).stem}.pdf", mime="application/pdf", key=f"dl_{idx}")

else:
    st.info("Envie um ou mais arquivos para iniciar a conversão.")


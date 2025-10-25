import sys, subprocess, streamlit as st
st.write("Python:", sys.version, "Exec:", sys.executable)
st.code(subprocess.check_output([sys.executable,"-m","pip","freeze"]).decode())

import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile, os, sys, io, re, base64
from html import unescape
from PIL import Image  # para imagens -> PDF

st.set_page_config(page_title="Converter para PDF", page_icon="🧾", layout="centered")
st.title("🧾 Converter HTML/XLS(X)/DOCX/Imagens ➜ PDF")
st.caption("Envie .html, .htm, .xls, .xlsx, .docx, imagens (jpg/png/gif/bmp/tiff/webp/svg) e/ou PDFs (.pdf). "
           "Você pode enviar vários arquivos e gerar um único PDF unificado.")

# -----------------------
# Diagnóstico e seleção automática de backend PDF
# -----------------------
import importlib.util as _iu
import importlib.metadata
import platform

def _has_module(mod: str) -> bool:
    return _iu.find_spec(mod) is not None

def _mod_ver(dist: str) -> str:
    try:
        return importlib.metadata.version(dist)
    except importlib.metadata.PackageNotFoundError:
        return "não encontrado"

def _test_weasyprint() -> tuple[bool, str]:
    try:
        if not _has_module("weasyprint"):
            return False, "WeasyPrint não instalado"
        from weasyprint import HTML
        html = HTML(string="<html><body><p>ok</p></body></html>")
        with tempfile.NamedTemporaryFile(delete=True, suffix=".pdf") as tmp:
            html.write_pdf(tmp.name)
        return True, "WeasyPrint ok"
    except Exception as e:
        return False, f"Falha ao usar WeasyPrint: {e}"

def _test_xhtml2pdf() -> tuple[bool, str]:
    try:
        if not _has_module("xhtml2pdf"):
            return False, "xhtml2pdf não instalado"
        from xhtml2pdf import pisa
        src = "<html><body><p>ok</p></body></html>"
        with tempfile.NamedTemporaryFile(delete=True, suffix=".pdf") as tmp:
            with open(tmp.name, "wb") as out:
                result = pisa.CreatePDF(src, dest=out)
            if result.err:
                return False, f"xhtml2pdf falhou na renderização: {result.err}"
        return True, "xhtml2pdf ok"
    except Exception as e:
        return False, f"Falha ao importar/usar xhtml2pdf: {e}"

def pick_pdf_backend() -> tuple[str, str]:
    if platform.system() == "Windows":
        if _has_module("weasyprint"):
            ok, msg = _test_weasyprint()
            if ok:
                return "weasyprint", msg
            ok2, msg2 = _test_xhtml2pdf()
            if ok2:
                return "xhtml2pdf", f"WeasyPrint indisponível ({msg}); usando xhtml2pdf ({msg2})"
            return "none", f"Sem backends funcionais: WeasyPrint=({msg}); xhtml2pdf=({msg2})"
        else:
            ok2, msg2 = _test_xhtml2pdf()
            if ok2:
                return "xhtml2pdf", msg2
            return "none", f"xhtml2pdf=({msg2})"
    ok, msg = _test_weasyprint()
    if ok:
        return "weasyprint", msg
    ok2, msg2 = _test_xhtml2pdf()
    if ok2:
        return "xhtml2pdf", f"WeasyPrint indisponível ({msg}); usando xhtml2pdf ({msg2})"
    return "none", f"Sem backends funcionais: WeasyPrint=({msg}); xhtml2pdf=({msg2})"

PDF_BACKEND, PDF_BACKEND_MSG = pick_pdf_backend()

# -----------------------
# Configurações (Sidebar) + Diagnóstico
# -----------------------
with st.sidebar:
    st.subheader("Diagnóstico PDF")
    st.write(f"**Python:** {sys.executable}")
    st.write(f"**SO:** {platform.system()} {platform.release()}")
    st.write(f"**Backend detectado:** `{PDF_BACKEND}`")
    st.write(f"**WeasyPrint:** {_mod_ver('weasyprint')}")
    st.write(f"**xhtml2pdf:** {_mod_ver('xhtml2pdf')}")
    st.caption(PDF_BACKEND_MSG)
    if platform.system() == "Windows" and PDF_BACKEND != "weasyprint":
        st.info(
            "WeasyPrint requer GTK3/Pango (Cairo) no Windows. Mesmo com o pacote Python instalado, "
            "faltam DLLs do runtime. O app está usando **xhtml2pdf** como fallback."
        )

_default_index = 0 if PDF_BACKEND == "weasyprint" else 1
if PDF_BACKEND == "none":
    _default_index = 1

engine = st.sidebar.selectbox("Motor de PDF", ["WeasyPrint (preservar layout)", "xhtml2pdf (compat)"], index=_default_index)
preserve_layout = st.sidebar.checkbox("Preservar layout do HTML (usar CSS do documento)", True)
page_size = st.sidebar.selectbox("Tamanho da página (se NÃO preservar layout)", ["A4", "Letter"], index=0)
orientation = st.sidebar.selectbox("Orientação (se NÃO preservar layout)", ["portrait", "landscape"], index=0)
margin_mm = st.sidebar.slider("Margem lateral (mm) – esquerda = direita", 5, 25, 10)
paginate_sheets = st.sidebar.checkbox("Quebrar página entre planilhas (Excel)", True)

combine_all = st.sidebar.checkbox("Unir todos os arquivos em um único PDF", True)
sanitize = st.sidebar.checkbox("Sanitizar CSS (apenas xhtml2pdf)", True)

uploaded_files = st.file_uploader(
    "Envie um ou mais arquivos .html, .htm, .xls, .xlsx, .docx, imagem (jpg/png/gif/bmp/tiff/webp/svg) **ou .pdf**",
    type=["html", "htm", "xls", "xlsx", "docx", "jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp", "svg", "pdf"],
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
                if j < len(css) and css[j] == "{":
                    depth += 1
                elif j < len(css) and css[j] == "}":
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
# Diagnóstico fino das dependências do xhtml2pdf
# -----------------------
def _probe_xhtml2pdf_deps():
    checks = {
        "reportlab": "reportlab",
        "Pillow (PIL)": "PIL",
        "html5lib": "html5lib",
        "pypdf": "pypdf",
        "svglib": "svglib",
        "cssselect2": "cssselect2",
        "tinycss2": "tinycss2",
        "lxml": "lxml",
        "python-bidi": "bidi",
        "arabic-reshaper": "arabic_reshaper",
    }
    missing = []
    present = []
    for label, mod in checks.items():
        try:
            __import__(mod)
            present.append(label)
        except Exception:
            missing.append(label)

    has_pypdf2 = False
    try:
        __import__("PyPDF2")
        has_pypdf2 = True
    except Exception:
        pass

    return missing, present, has_pypdf2

# -----------------------
# Helpers extras para fallback forte (xhtml2pdf)
# -----------------------
def _strip_external_fonts(html: str) -> str:
    html = re.sub(r"@font-face\s*{[^}]*}", "", html, flags=re.IGNORECASE|re.DOTALL)
    html = re.sub(r'<link[^>]+href="[^"]*fonts[^"]*"[^>]*>', "", html, flags=re.IGNORECASE)
    html = re.sub(r'url\([^)]+\.woff2?\)', "", html, flags=re.IGNORECASE)
    return html

def _very_simple_html(html: str) -> str:
    body = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.IGNORECASE|re.DOTALL)
    body = re.sub(r'\sstyle="[^"]*"', "", body, flags=re.IGNORECASE)
    return f"""<html><head><meta charset="utf-8">
    <style>
      body {{ font-family: Arial, sans-serif; font-size: 12pt; margin: 10mm; }}
      img {{ max-width: 100%; height: auto; display: block; margin: 6px 0; }}
      table {{ width:100%; border-collapse: collapse; table-layout: fixed; }}
      th, td {{ border: 1px solid #999; padding: 6px; word-wrap: break-word; }}
      h1,h2,h3 {{ margin: 8px 0; }}
      pre,code {{ white-space: pre-wrap; word-break: break-word; }}
    </style>
    </head><body>{body}</body></html>"""

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
    with image.open() as img_bytes:
        encoded = base64.b64encode(img_bytes.read()).decode("ascii")
    return {"src": f"data:{image.content_type};base64,{encoded}"}

def docx_to_html(uploaded_file) -> str:
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

    try:
        pdf_bytes = HTML(string=html_str, base_url=base_url or ".").write_pdf(
            stylesheets=styles, font_config=font_config
        )
        if pdf_bytes is None:
            raise RuntimeError("WeasyPrint não retornou bytes do PDF.")
        return pdf_bytes
    except Exception:
        def _strip_emojis(text: str) -> str:
            ranges = [(0x1F600,0x1F64F),(0x1F300,0x1F5FF),(0x1F680,0x1F6FF),(0x2600,0x26FF),
                      (0x2700,0x27BF),(0xFE00,0xFE0F),(0x1F900,0x1F9FF),(0x1FA70,0x1FAFF),(0x1F1E6,0x1F1FF)]
            out=[]
            for ch in text:
                cp=ord(ch)
                if any(a<=cp<=b for a,b in ranges): continue
                out.append(ch)
            return "".join(out)

        from weasyprint import CSS
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
    import importlib.util as _ius
    spec = _ius.find_spec("xhtml2pdf")
    if spec is None:
        st.error("xhtml2pdf não está instalado neste Python. Rode:  python -m pip install xhtml2pdf")
        st.stop()

    try:
        from xhtml2pdf import pisa
    except ImportError:
        st.error("xhtml2pdf não está instalado neste Python. Rode:  python -m pip install xhtml2pdf")
        st.stop()
    except Exception as e:
        st.error(f"Falha ao importar xhtml2pdf: {e.__class__.__name__}: {e}")
        missing, present, has_pypdf2 = _probe_xhtml2pdf_deps()
        with st.expander("🧩 Diagnóstico de dependências do xhtml2pdf"):
            if present: st.write("**Pacotes encontrados:** " + ", ".join(present))
            if missing: st.write("**Pacotes ausentes:** " + ", ".join(missing))
            if has_pypdf2:
                st.warning("Conflito: **PyPDF2** instalado. Prefira **pypdf**.")
                st.code("python -m pip uninstall -y PyPDF2", language="bash")
            st.code("python -m pip install --upgrade reportlab 'Pillow<12' pypdf svglib html5lib cssselect2 tinycss2 lxml python-bidi arabic-reshaper", language="bash")
        st.stop()

    _patch_xhtml2pdf_lower()

    page_css = f"@page {{ margin-left: {margin_mm}mm; margin-right: {margin_mm}mm; }}"
    if sanitize or preserve_layout:
        candidate1 = sanitize_html_for_xhtml2pdf(html_str, page_css)
    else:
        candidate1 = _inject_page_css(html_str, page_css)

    if "<meta charset" not in candidate1.lower():
        if "<head>" in candidate1.lower():
            candidate1 = candidate1.replace("<head>", "<head><meta charset='utf-8'>", 1)
        else:
            candidate1 = f"<html><head><meta charset='utf-8'></head><body>{candidate1}</body></html>"

    attempts = []
    attempts.append(("HTML atual", candidate1))
    cand2 = sanitize_html_for_xhtml2pdf(_strip_external_fonts(candidate1), page_css)
    attempts.append(("HTML sanitizado (forte)", cand2))
    cand3 = _very_simple_html(candidate1)
    attempts.append(("Modo simples (HTML básico)", cand3))

    last_error = None
    last_log = None

    for label, html_try in attempts:
        out = io.BytesIO()
        pisa_log = io.StringIO()
        try:
            res = pisa.CreatePDF(src=html_try, dest=out, encoding="utf-8", log=pisa_log)
            if res.err:
                last_error = RuntimeError(f"xhtml2pdf retornou erro (tentativa: {label})")
                last_log = pisa_log.getvalue()
            else:
                return out.getvalue()
        except Exception as e:
            last_error = e
            last_log = pisa_log.getvalue()

    with st.expander("📄 Log detalhado do xhtml2pdf (pisa)"):
        if last_log:
            st.code(last_log)
        else:
            st.write("Sem log do pisa. Veja o traceback abaixo.")
    with st.expander("⚠️ Traceback da última tentativa"):
        if last_error:
            st.exception(last_error)

    try:
        html_bytes = attempts[1][1].encode("utf-8", errors="ignore")
        st.download_button("⬇️ Baixar HTML sanitizado (para depurar)",
                           data=html_bytes, file_name="html_sanitizado_para_debug.html", mime="text/html")
    except Exception:
        pass

    st.error("xhtml2pdf encontrou um erro ao gerar o PDF. Revise o **Log do pisa** acima.")
    st.stop()

# -----------------------
# Fallback automático
# -----------------------
def convert_html_to_pdf(html_str: str, base_url: str = ".") -> bytes:
    if engine.startswith("WeasyPrint"):
        try:
            return build_pdf_weasy(html_str, base_url)
        except Exception as e:
            if _iu.find_spec("xhtml2pdf"):
                st.warning("WeasyPrint indisponível. Usando xhtml2pdf como fallback.")
                return build_pdf_xhtml2pdf(html_str)
            st.error("WeasyPrint falhou e xhtml2pdf não está instalado.")
            raise e
    else:
        return build_pdf_xhtml2pdf(html_str)

# -----------------------
# Excel -> HTML
# -----------------------
def excel_to_html(uploaded_file, break_between=True) -> str:
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
            st.error("Falta 'openpyxl' para ler .xlsx.")
            st.stop()
    elif name.endswith(".xls"):
        xls_engine = "xlrd"
        try:
            import xlrd  # noqa
        except Exception:
            st.error("Falta 'xlrd' (>=2.0) para ler .xls.")
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
    writer = None
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

    if ext == ".pdf":
        file.seek(0)
        return file.read()

    if ext in [".html", ".htm"]:
        html_str, base_url = read_html_and_base(file)
        return convert_html_to_pdf(html_str, base_url)

    elif ext in [".xls", ".xlsx"]:
        html_doc = excel_to_html(file, break_between=paginate_sheets)
        return convert_html_to_pdf(html_doc, base_url=".")

    elif ext == ".docx":
        html_doc = docx_to_html(file)
        return convert_html_to_pdf(html_doc, base_url=".")

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

    if errors:
        for name, e in errors:
            with st.expander(f"⚠️ Falha ao converter: {name}"):
                st.exception(e)

    if not pdfs:
        st.warning("Nenhum arquivo pôde ser convertido.")
        st.stop()

    # =======================
    # NOVO: Seleção & Ordem
    # =======================
    st.subheader("Seleção e ordem dos documentos para unificação")
    df_sel = pd.DataFrame({
        "Incluir": [True] * len(pdfs),
        "Nome": [name for name, _ in pdfs],
        "Ordem": list(range(1, len(pdfs) + 1)),
    })

    # Persistência simples entre reruns
    if "df_ordem_cache" not in st.session_state or \
       sorted(st.session_state.get("df_ordem_cache", {}).get("Nome", [])) != sorted(df_sel["Nome"].tolist()):
        st.session_state["df_ordem_cache"] = df_sel.copy()
    else:
        # Restaura preferências anteriores (se nomes baterem)
        cache = st.session_state["df_ordem_cache"]
        df_sel = cache.reindex(columns=df_sel.columns).fillna(df_sel)

    edited = st.data_editor(
        df_sel,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Incluir": st.column_config.CheckboxColumn(help="Marque para incluir no PDF final."),
            "Nome": st.column_config.TextColumn(disabled=True),
            "Ordem": st.column_config.NumberColumn(
                min_value=1, max_value=max(1, len(pdfs)), step=1,
                help="Defina a ordem desejada (1 = primeiro)."
            ),
        },
        key="editor_ordem"
    )
    # Guarda no cache para o próximo rerun
    st.session_state["df_ordem_cache"] = edited.copy()

    # Normaliza a ordem (resolve empates mantendo ordem original)
    edited["_idx_orig"] = range(len(edited))
    edited_sorted = (
        edited[edited["Incluir"] == True]
        .sort_values(by=["Ordem", "_idx_orig"], kind="mergesort")
        .reset_index(drop=True)
    )

    if edited_sorted.empty:
        st.warning("Nenhum documento selecionado para unificação. Marque pelo menos um em 'Incluir'.")
    else:
        # Nome do arquivo final
        out_name = st.text_input("Nome do PDF unificado", value="documentos_unificados.pdf")
        if not out_name.lower().endswith(".pdf"):
            out_name += ".pdf"

        # Botão para unir conforme seleção/ordem
        if st.button("🔗 Unir conforme seleção e ordem"):
            bytes_na_ordem = []
            nomes_na_ordem = edited_sorted["Nome"].tolist()
            # mapeia nome -> bytes (da lista pdfs)
            mapa = {n: b for n, b in pdfs}
            for n in nomes_na_ordem:
                if n in mapa:
                    bytes_na_ordem.append(mapa[n])

            if len(bytes_na_ordem) == 1:
                st.info("Apenas um documento selecionado. Baixe-o diretamente abaixo.")
                st.download_button("⬇️ Baixar PDF", data=bytes_na_ordem[0],
                                   file_name=Path(nomes_na_ordem[0]).with_suffix(".pdf").name,
                                   mime="application/pdf", key="dl_single_selected")
            else:
                merged_bytes = merge_pdfs(bytes_na_ordem)
                st.success(f"Unificados {len(bytes_na_ordem)} documentos na ordem definida.")
                st.download_button("⬇️ Baixar PDF unificado", data=merged_bytes,
                                   file_name=out_name, mime="application/pdf", key="dl_merged_custom")

    # Também oferece os downloads individuais abaixo
    st.divider()
    st.subheader("Downloads individuais")
    for idx, (name, b) in enumerate(pdfs, start=1):
        st.download_button(f"⬇️ Baixar {idx}: {name}.pdf", data=b,
                           file_name=f"{Path(name).stem}.pdf", mime="application/pdf", key=f"dl_{idx}")

else:
    st.info("Envie um ou mais arquivos para iniciar a conversão.")

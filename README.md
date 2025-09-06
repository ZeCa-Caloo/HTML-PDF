
# 🧾 Streamlit – Conversor de HTML/XLS(X) para PDF

App Streamlit simples que converte **HTML** e **Excel (.xls/.xlsx)** em **PDF**.

## 🚀 Como funciona
- **HTML → PDF**: renderiza o HTML com **wkhtmltopdf** (via `pdfkit`) *ou* **WeasyPrint** (opcional).
- **Excel → PDF**: lê as planilhas com `pandas`, estiliza como HTML e converte usando o mesmo motor acima.
  - Quebra de página opcional entre planilhas.
  - Tamanho de papel (A4/Letter), orientação e margens configuráveis.

## 📦 Rodando Localmente (wkhtmltopdf + pdfkit)
1. Instale o Python 3.10+
2. Instale o **wkhtmltopdf** (Linux/macOS/Windows)
3. Crie um venv e instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
4. Rode o app:
   ```bash
   streamlit run app.py
   ```

> Dica: No Linux, você também pode usar Docker (ver abaixo).

## ☁️ Deploy no Streamlit Community Cloud
- Faça o push deste repositório no GitHub.
- Em **Advanced settings**, aponte `packages.txt` (instala o `wkhtmltopdf`).
- O app usa `app.py` como entrypoint.

## 🪟 Modo Windows com Excel (100% da formatação)
Se você precisa **preservar fielmente** a impressão do Excel (larguras de coluna, cabeçalho/rodapé, quebras):
- Use o **Excel COM** no Windows com MS Office instalado.
- Veja `excel_to_pdf_win.py` para um exemplo de exportação nativa do Excel para PDF.

## 🧪 Alternativa sem wkhtmltopdf: WeasyPrint
- Descomente `weasyprint` no `requirements.txt` e as libs em `packages.txt`.
- No app, selecione “WeasyPrint” no seletor lateral.

## 🐳 Docker (opcional, Linux)
Crie uma imagem com o `Dockerfile` e execute o container com a porta 8501.

```bash
docker build -t streamlit-converter .
docker run -p 8501:8501 streamlit-converter
```

## ⚠️ Limitações e Notas
- Arquivos **.xls** antigos são suportados via `xlrd`.
- HTMLs que dependem de **recursos externos** (CSS/JS/imagens remotas) podem precisar de `enable-local-file-access` e caminhos absolutos.
- WeasyPrint exige libs do SO (cairo, pango, gdk-pixbuf).

---

### 📜 Licença
MIT

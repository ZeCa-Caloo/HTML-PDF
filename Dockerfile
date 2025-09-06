
FROM python:3.11-slim

# Instala wkhtmltopdf e dependências úteis
RUN apt-get update && apt-get install -y --no-install-recommends \
    wkhtmltopdf \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY . /app

EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]

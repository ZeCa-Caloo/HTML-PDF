@echo off
echo === Criando ambiente virtual (.venv) ===
if not exist .venv (
  py -m venv .venv
)
echo === Ativando .venv ===
call .venv\Scripts\activate.bat
echo === Atualizando pip ===
python -m pip install --upgrade pip
echo === Instalando dependencias ===
pip install -r requirements.txt
echo === Iniciando o Streamlit ===
streamlit run app\app.py

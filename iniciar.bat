@echo off
echo Iniciando Gerador de Plano Alimentar...
cd /d "%~dp0"
streamlit run app.py --server.headless false
pause

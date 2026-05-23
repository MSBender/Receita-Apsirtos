@echo off
echo ============================================
echo  Instalacao - Gerador de Plano Alimentar
echo ============================================
echo.

REM Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado.
    echo Baixe em: https://www.python.org/downloads/
    echo Marque "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo [OK] Python encontrado.
echo.

REM Instalar dependencias Python
echo Instalando bibliotecas Python...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Bibliotecas instaladas.
echo.

REM Verificar LibreOffice
if exist "C:\Program Files\LibreOffice\program\soffice.exe" (
    echo [OK] LibreOffice encontrado.
) else (
    echo [ATENCAO] LibreOffice NAO encontrado.
    echo Baixe e instale em: https://www.libreoffice.org/download/
    echo Necessario para converter PPTX em PDF.
)
echo.

REM Verificar Tesseract
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo [OK] Tesseract OCR encontrado.
) else (
    echo [ATENCAO] Tesseract NAO encontrado.
    echo Baixe em: https://github.com/UB-Mannheim/tesseract/wiki
    echo Durante a instalacao, selecione "Portuguese" nos idiomas adicionais.
)
echo.

echo ============================================
echo  Instalacao concluida!
echo  Execute "iniciar.bat" para abrir o sistema.
echo ============================================
pause

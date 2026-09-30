@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERRO] Python nao foi encontrado no PATH.
    echo Instale o Python 3.11 ou superior e tente novamente.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Criando ambiente virtual...
    python -m venv .venv
    if errorlevel 1 goto :erro
)

echo Instalando ou atualizando dependencias...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :erro

echo Iniciando o Nexus no navegador...
".venv\Scripts\python.exe" -m streamlit run app.py
exit /b 0

:erro
echo.
echo [ERRO] Nao foi possivel preparar ou iniciar o aplicativo.
pause
exit /b 1


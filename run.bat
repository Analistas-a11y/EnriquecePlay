@echo off
setlocal
cd /d "%~dp0"
echo Iniciando o servidor...
if not exist ".venv\Scripts\python.exe" (
    echo Ambiente Python nao encontrado. Crie .venv e instale requirements.txt.
    exit /b 1
)
"%~dp0.venv\Scripts\python.exe" app.py
pause

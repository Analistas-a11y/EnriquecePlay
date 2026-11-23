@echo off
chcp 65001 >nul
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" logs_cli.py %*
) else if exist "..\.venv\Scripts\python.exe" (
    "..\.venv\Scripts\python.exe" logs_cli.py %*
) else (
    py logs_cli.py %*
)

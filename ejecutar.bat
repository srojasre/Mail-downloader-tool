@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Creando entorno virtual en .venv...
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if not exist ".venv\Scripts\python.exe" (
        echo [ERROR] No se encontro Python 3 instalado en el sistema.
        pause
        exit /b 1
    )
    echo [INFO] Instalando dependencias desde requirements.txt...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [ERROR] Fallo la instalacion de dependencias.
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" main.py %*

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
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)

echo [INFO] Generando ejecutable portable dist\CorreoPDFExporter.exe...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile ^
    --name CorreoPDFExporter ^
    --hidden-import=lxml._elementpath ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    main.py

if %errorlevel% neq 0 (
    echo [ERROR] No se pudo compilar el ejecutable.
    pause
    exit /b 1
)

echo.
echo [OK] Ejecutable generado exitosamente en: dist\CorreoPDFExporter.exe
echo Puedes copiar CorreoPDFExporter.exe a cualquier PC con Windows sin instalar Python.
pause

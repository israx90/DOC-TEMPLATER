@echo off
REM ============================================
REM  EDTECH DOC TEMPLATER — Setup Windows
REM  Ejecutar: click derecho → Ejecutar como Administrador
REM ============================================
title EDTECH DOC TEMPLATER — Instalador Windows
echo.
echo ==========================================
echo   EDTECH DOC TEMPLATER — Instalador Windows
echo ==========================================
echo.

REM 1. Check Python
echo [1/5] Verificando Python...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo      Python no encontrado. Descargando...
    echo.
    echo      IMPORTANTE: Durante la instalacion, marca la casilla
    echo      "Add Python to PATH"
    echo.
    curl -L -o python_installer.exe https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
    start /wait python_installer.exe /passive InstallAllUsers=1 PrependPath=1
    del python_installer.exe
    echo      ✅ Python instalado
) else (
    for /f "tokens=*" %%v in ('python --version') do echo      ✅ %%v ya instalado
)

REM 2. Check/Install Chocolatey (for Tesseract)
echo [2/5] Verificando Chocolatey...
choco --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo      Instalando Chocolatey...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Set-ExecutionPolicy Bypass -Scope Process -Force; ^
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; ^
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))"
    echo      ✅ Chocolatey instalado
) else (
    echo      ✅ Chocolatey ya instalado
)

REM 3. Install Tesseract
echo [3/5] Verificando Tesseract OCR...
tesseract --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo      Instalando Tesseract OCR...
    choco install tesseract --no-progress -y
    echo      ✅ Tesseract instalado
) else (
    echo      ✅ Tesseract ya instalado
)

REM 4. Virtual environment + dependencies
echo [4/5] Configurando entorno Python...
cd /d "%~dp0"

if not exist "venv" (
    python -m venv venv
    echo      Entorno virtual creado
)

call venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
pip install pyinstaller --quiet
echo      ✅ Dependencias Python instaladas

REM 5. Create working directories
echo [5/5] Creando carpetas de trabajo...
if not exist "uploads" mkdir uploads
if not exist "outputs" mkdir outputs
echo      ✅ Listo

echo.
echo ==========================================
echo   ✅ INSTALACION COMPLETADA
echo ==========================================
echo.
echo   Para ejecutar la aplicacion:
echo     venv\Scripts\activate
echo     python test_app.py
echo.
echo   Para compilar el instalador .exe:
echo     venv\Scripts\activate
echo     python build.py
echo.
echo ==========================================
echo.
pause

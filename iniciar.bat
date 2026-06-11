@echo off
:: ============================================================
::  Descargador de Cartas Glosa - Seguros Mundial
:: ============================================================

setlocal

echo.
echo ============================================================
echo   Descargador de Cartas Glosa -- Seguros Mundial
echo ============================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%venv"

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado.
    pause
    exit /b 1
)

:: Crear entorno virtual si no existe
if not exist "%VENV_DIR%" (
    echo Creando entorno virtual Python...
    python -m venv "%VENV_DIR%"
)

:: Activar entorno virtual
call "%VENV_DIR%\Scripts\activate.bat"

:: Instalar/actualizar dependencias siempre (detecta nuevos paquetes automaticamente)
echo Verificando dependencias...
pip install --upgrade -r "%SCRIPT_DIR%requirements.txt"
if errorlevel 1 (
    echo [ERROR] Fallo la instalacion de dependencias.
    pause
    exit /b 1
)

:: Instalar Chromium (solo descarga si ya esta instalado, es rapido)
echo Verificando Playwright Chromium...
python -m playwright install chromium

:: Crear carpeta de descargas
if not exist "%SCRIPT_DIR%downloads" mkdir "%SCRIPT_DIR%downloads"

:: Abrir navegador
start "" "http://localhost:8080"

:: Ejecutar aplicacion
python "%SCRIPT_DIR%app.py"

pause

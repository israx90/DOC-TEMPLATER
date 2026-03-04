@echo off
title EDTECH DOC TEMPLATER — Instalador
echo.
echo  Iniciando instalador...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_windows.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Hubo un error. Revisa los mensajes arriba.
    echo.
)
pause

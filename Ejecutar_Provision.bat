@echo off
title Ejecutando Provision 2026
echo ======================================================
echo Iniciando Proceso de Provision...
echo ======================================================
cd /d %~dp0

:: Verificar si Python está instalado
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python no esta instalado o no se encuentra en el PATH.
    pause
    exit /b
)

python 2026_Provision.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] El proceso termino con errores. Ver mensaje arriba.
    pause
) else (
    echo.
    echo [OK] Proceso finalizado con exito.
    timeout /t 5
)
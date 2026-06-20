@echo off
title Ayla Add-Exc - Instalacion
cd /d "%~dp0"

:: Solicitar elevacion
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    echo Solicitando permisos de Administrador...
    powershell start-process cmd -ArgumentList "/c `"%~f0`"" -verb runas
    exit /b
)

echo ============================================
echo   Ayla Add-Exc - Instalacion en Excel
echo ============================================
echo.
echo Este asistente instalara el complemento en Excel.
echo.
echo Paso 1: Creando carpeta de confianza...
set WEFDIR=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef
if not exist "%WEFDIR%" mkdir "%WEFDIR%"

echo Paso 2: Copiando manifest.xml a Excel...
copy /Y "%~dp0manifest.xml" "%WEFDIR%\manifest.xml"

echo Paso 3: Configurando Excel para aceptar complementos...
reg add HKCU\Software\Microsoft\Office\16.0\Wef /f >nul 2>&1
reg add HKCU\Software\Microsoft\Office\16.0\Excel\Options /f >nul 2>&1

echo.
echo ============================================
echo   INSTALACION COMPLETADA!
echo ============================================
echo.
echo   CIERRA Excel completamente y vuelve a abrirlo.
echo   Busca la pestana "Ayla Add-Exc" en la cinta.
echo.
echo   Tu API Key de Gemini (GRATIS):
echo   https://aistudio.google.com/apikey
echo.
echo   Endpoint: https://generativelanguage.googleapis.com/v1beta/openai/
echo   Modelo:   gemini-2.5-flash
echo.
pause

@echo off
setlocal
title Ayla - Puente FreeLLMAPI
set "AYLA_DIR=%LOCALAPPDATA%\Ayla"
set "PROXY=%AYLA_DIR%\freellmapi-proxy.js"
set "PROXY_URL=https://raw.githubusercontent.com/mhckrown/ayla-add-exc/main/tools/freellmapi-proxy.js"

echo ============================================
echo   Ayla - Conexion con FreeLLMAPI local
echo ============================================
echo.

rem 1) Node.js
set "NODE=node"
where node >nul 2>&1
if %errorlevel% equ 0 goto node_ok
if exist "%ProgramFiles%\nodejs\node.exe" (
  set "NODE=%ProgramFiles%\nodejs\node.exe"
  goto node_ok
)
echo [1/3] Node.js no esta instalado. Instalando con winget...
winget install -e --id OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
if exist "%ProgramFiles%\nodejs\node.exe" (
  set "NODE=%ProgramFiles%\nodejs\node.exe"
  goto node_ok
)
echo.
echo No se pudo instalar Node.js automaticamente.
echo Instalalo desde https://nodejs.org (version LTS) y vuelve a ejecutar este archivo.
pause
exit /b 1
:node_ok
echo [1/3] Node.js OK

rem 2) Descargar el puente
if not exist "%AYLA_DIR%" mkdir "%AYLA_DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -UseBasicParsing '%PROXY_URL%' -OutFile '%PROXY%'; exit 0 } catch { exit 1 }"
if %errorlevel% neq 0 (
  if not exist "%PROXY%" (
    echo No se pudo descargar el puente. Revisa tu conexion a internet.
    pause
    exit /b 1
  )
  echo Sin internet: usando la copia guardada del puente.
)
echo [2/3] Puente descargado en %PROXY%

rem 3) Comprobar FreeLLMAPI
powershell -NoProfile -Command "foreach ($p in 3001,31415) { try { $null = Invoke-WebRequest -UseBasicParsing ('http://127.0.0.1:' + $p + '/v1/models') -TimeoutSec 3; exit 0 } catch { if ($_.Exception.Response) { exit 0 } } }; exit 1"
if %errorlevel% neq 0 (
  echo.
  echo AVISO: FreeLLMAPI no responde en 127.0.0.1:3001 ni :31415.
  echo Arrancalo y deja esta ventana abierta; el puente lo usara en cuanto este listo.
  echo.
) else (
  echo [3/3] FreeLLMAPI responde
)

echo.
echo ============================================
echo   En Excel - Ayla - Configuracion:
echo     Endpoint: http://127.0.0.1:3939/v1
echo     API Key : tu clave unificada de FreeLLMAPI
echo     Modelo  : auto
echo.
echo   NO cierres esta ventana mientras uses Ayla.
echo ============================================
echo.
"%NODE%" "%PROXY%"
echo.
echo El puente se detuvo.
pause

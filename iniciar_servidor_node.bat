@echo off
setlocal

cd /d "%~dp0"

echo [Geo Rural] Verificando Node.js...
where node >nul 2>&1
if errorlevel 1 (
  echo ERROR: Node.js no esta instalado o no esta en PATH.
  echo Instala Node.js desde https://nodejs.org/ y vuelve a intentar.
  pause
  exit /b 1
)

set "NPM_CMD="
for /f "delims=" %%I in ('where npm.cmd 2^>nul') do if not defined NPM_CMD set "NPM_CMD=%%~fI"

if not defined NPM_CMD if exist "C:\Program Files\nodejs\npm.cmd" set "NPM_CMD=C:\Program Files\nodejs\npm.cmd"

if not defined NPM_CMD (
  echo ERROR: No se encontro npm.cmd.
  echo Reinstala Node.js y asegurate de incluir npm.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo [Geo Rural] Instalando dependencias por primera vez...
  call "%NPM_CMD%" install
  if errorlevel 1 (
    echo ERROR: Fallo la instalacion de dependencias.
    pause
    exit /b 1
  )
)

set "DB_HOST=127.0.0.1"
set "DB_PORT=3306"

if exist ".env" (
  for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
    if /I "%%~A"=="DB_HOST" set "DB_HOST=%%~B"
    if /I "%%~A"=="DB_PORT" set "DB_PORT=%%~B"
  )
)

if "%DB_HOST%"=="" set "DB_HOST=127.0.0.1"
if "%DB_PORT%"=="" set "DB_PORT=3306"

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $hostName='%DB_HOST%'; $port=[int]%DB_PORT%; $client=New-Object System.Net.Sockets.TcpClient; try { $async=$client.BeginConnect($hostName,$port,$null,$null); if(-not $async.AsyncWaitHandle.WaitOne(1500)){ throw 'timeout' }; $client.EndConnect($async) | Out-Null; exit 0 } catch { exit 1 } finally { $client.Close() }" >nul 2>&1
if errorlevel 1 (
  echo [Geo Rural] No hay conexion a MySQL en %DB_HOST%:%DB_PORT%.
  echo [Geo Rural] Inicia MySQL desde XAMPP y vuelve a intentar.
  echo [Geo Rural] Si MySQL usa otro puerto, actualiza DB_PORT en .env.
  pause
  exit /b 1
)

echo [Geo Rural] Iniciando servidor Node...
echo [Geo Rural] URL: http://127.0.0.1:3000
call "%NPM_CMD%" start

echo.
echo [Geo Rural] El servidor se detuvo o hubo un error.
pause
exit /b 0

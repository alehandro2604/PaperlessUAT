@echo off
REM Retries the Electron download and saves everything to electron-install.log
setlocal
cd /d "%~dp0"
set LOG=%~dp0electron-install.log
echo Retrying Electron download, please wait (up to a few minutes)...
echo === node/npm === > "%LOG%"
node -v >> "%LOG%" 2>&1
npm -v >> "%LOG%" 2>&1
echo === proxy/env === >> "%LOG%"
set | findstr /i "proxy electron npm_config" >> "%LOG%" 2>&1
echo === npm config registry/proxy === >> "%LOG%"
call npm config get proxy >> "%LOG%" 2>&1
call npm config get https-proxy >> "%LOG%" 2>&1
echo === install.js === >> "%LOG%"
set ELECTRON_ENABLE_LOGGING=1
set DEBUG=@electron/get:*
if exist node_modules\electron\dist rmdir /s /q node_modules\electron\dist
if exist node_modules\electron\path.txt del node_modules\electron\path.txt
call node node_modules\electron\install.js >> "%LOG%" 2>&1
echo exit code %errorlevel% >> "%LOG%"
echo === result === >> "%LOG%"
dir node_modules\electron\dist >> "%LOG%" 2>&1
echo.
echo Finished. Log saved to electron-install.log - tell Claude it is done.
pause

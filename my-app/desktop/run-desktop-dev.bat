@echo off
REM Quick test: build the site and open it in Electron (no installer).
setlocal
cd /d "%~dp0.."
call npx ng build || goto :fail
cd desktop
if not exist node_modules call npm install || goto :fail
call npm start
exit /b 0
:fail
echo FAILED
exit /b 1

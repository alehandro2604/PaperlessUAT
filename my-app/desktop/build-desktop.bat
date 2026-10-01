@echo off
REM Builds the Angular site, then packages the Windows desktop app.
REM Run from anywhere; output goes to my-app\desktop\release\
setlocal
cd /d "%~dp0.."
call npx ng build || goto :fail
cd desktop
if not exist node_modules call npm install || goto :fail
call npm run dist || goto :fail
echo.
echo Done. Installer + portable exe are in: %cd%\release
exit /b 0
:fail
echo BUILD FAILED
exit /b 1

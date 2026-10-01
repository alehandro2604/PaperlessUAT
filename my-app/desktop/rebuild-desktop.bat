@echo off
REM Cleans the old half-built output and packages the desktop app again, saving a log.
setlocal
cd /d "%~dp0"
set LOG=%~dp0build-log.txt
echo Cleaning old build...
if exist release rmdir /s /q release
echo Packaging (a few minutes)... > "%LOG%"
call npm run dist >> "%LOG%" 2>&1
echo exit code %errorlevel% >> "%LOG%"
echo.
type "%LOG%"
echo.
echo ---- Finished. Files in the release folder: ----
dir /b release
echo.
echo Tell Claude it is finished.
pause

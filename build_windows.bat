@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_windows.ps1"
if errorlevel 1 goto :error
pause
exit /b 0
:error
echo.
echo Windows build failed. See the message above.
pause
exit /b 1

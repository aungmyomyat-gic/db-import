@echo off
REM Pull the latest release from main and rebuild the container.
REM Your saved DB connection in .\data is kept.
cd /d "%~dp0"
echo ==^> Pulling latest version from GitHub (main)...
git pull --ff-only origin main || goto :error
echo ==^> Rebuilding and restarting the app...
docker compose -f docker-compose.share.yml up -d --build || goto :error
echo ==^> Done. Open http://localhost:5000 and reload the page.
pause
exit /b 0
:error
echo.
echo Update failed. See the message above.
pause
exit /b 1

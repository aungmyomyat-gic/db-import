@echo off
REM Update the current branch; keep the saved connection in .\data.
cd /d "%~dp0"
set "app_branch="
for /f "delims=" %%B in ('git branch --show-current') do set "app_branch=%%B"
if not defined app_branch goto :no_branch
echo ==^> Pulling latest version from GitHub (%app_branch%)...
git pull --ff-only origin "%app_branch%" || goto :error
echo ==^> Rebuilding and restarting the app...
docker compose up -d --build || goto :error
echo ==^> Done. Open http://localhost:5000 and reload the page.
pause
exit /b 0
:no_branch
echo Check out the CSV/SQL branch before updating.
:error
echo.
echo Update failed. See the message above.
pause
exit /b 1

#!/usr/bin/env sh
# Update the current branch and rebuild; keep the saved connection in ./data.
set -e
cd "$(dirname "$0")"
app_branch=$(git branch --show-current)
if [ -z "$app_branch" ]; then
  echo "Check out the CSV/SQL branch before updating." >&2
  exit 1
fi
echo "==> Pulling latest version from GitHub ($app_branch)..."
git pull --ff-only origin "$app_branch"
echo "==> Rebuilding and restarting the app..."
docker compose up -d --build
echo "==> Done. Open http://localhost:5000 and reload the page."

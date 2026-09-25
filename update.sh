#!/usr/bin/env sh
# Pull the latest release from main and rebuild the container.
# Your saved DB connection in ./data is kept.
set -e
cd "$(dirname "$0")"
echo "==> Pulling latest version from GitHub (main)..."
git pull --ff-only origin main
echo "==> Rebuilding and restarting the app..."
docker compose -f docker-compose.share.yml up -d --build
echo "==> Done. Open http://localhost:5000 and reload the page."

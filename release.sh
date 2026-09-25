#!/usr/bin/env bash
# Release a new version: bump version.json on dev, merge dev → main, tag, push.
#
#   ./release.sh patch       1.1.0 → 1.1.1   (bug fixes)
#   ./release.sh minor       1.1.0 → 1.2.0   (new features)
#   ./release.sh major       1.1.0 → 2.0.0   (big / breaking changes)
#   ./release.sh minor --dry-run   show what would happen, change nothing
#
# Release notes are generated from commit subjects since the last v* tag
# ("feat: add paste rows" → "Add paste rows"). You can edit them before
# anything is committed.
set -euo pipefail
cd "$(dirname "$0")"

BUMP="${1:-}"
DRY_RUN=0
[[ "${2:-}" == "--dry-run" || "${1:-}" == "--dry-run" ]] && DRY_RUN=1
[[ "$BUMP" == "--dry-run" ]] && BUMP="${2:-}"
if [[ "$BUMP" != "patch" && "$BUMP" != "minor" && "$BUMP" != "major" ]]; then
  echo "Usage: ./release.sh patch|minor|major [--dry-run]" >&2
  exit 1
fi

die() { echo "✗ $*" >&2; exit 1; }

# ── Preconditions ────────────────────────────────────────────────────────────
[[ "$(git rev-parse --abbrev-ref HEAD)" == "dev" ]] || die "Run this on the dev branch (git checkout dev)."
[[ -z "$(git status --porcelain)" ]] || die "Working tree has uncommitted changes. Commit or stash them first."

echo "==> Fetching origin..."
git fetch --quiet --tags origin
if git rev-parse --verify --quiet origin/dev >/dev/null; then
  git merge --ff-only --quiet origin/dev || die "Local dev has diverged from origin/dev. Sort that out first."
fi
if git rev-parse --verify --quiet origin/main >/dev/null; then
  git merge-base --is-ancestor origin/main HEAD \
    || die "main has commits that dev doesn't. Run: git merge origin/main (on dev), then retry."
fi

# ── Work out the new version and notes ───────────────────────────────────────
CURRENT="$(python3 -c 'import json; print(json.load(open("version.json"))["version"])')"
NEW="$(python3 - "$CURRENT" "$BUMP" <<'PY'
import re, sys
parts = [int(n) for n in re.findall(r"\d+", sys.argv[1])[:3]] + [0, 0, 0]
major, minor, patch = parts[:3]
bump = sys.argv[2]
if bump == "major": major, minor, patch = major + 1, 0, 0
elif bump == "minor": minor, patch = minor + 1, 0
else: patch += 1
print(f"{major}.{minor}.{patch}")
PY
)"
git rev-parse --verify --quiet "refs/tags/v$NEW" >/dev/null && die "Tag v$NEW already exists."

LAST_TAG="$(git describe --tags --abbrev=0 --match 'v*' 2>/dev/null || true)"
RANGE="${LAST_TAG:+$LAST_TAG..}HEAD"
NOTES_FILE="$(mktemp)"
trap 'rm -f "$NOTES_FILE"' EXIT
# One note per commit subject; drop merges and earlier release commits,
# strip "feat:" / "fix(scope):" / "【fix】" prefixes, capitalise.
git log "$RANGE" --no-merges --format='%s' \
  | grep -viE '^release: ' \
  | sed -E 's/^[a-zA-Z]+(\([^)]*\))?!?:[[:space:]]*//; s/^【[^】]*】[[:space:]]*//' \
  | awk 'NF && !seen[$0]++ { print toupper(substr($0,1,1)) substr($0,2) }' \
  > "$NOTES_FILE" || true
[[ -s "$NOTES_FILE" ]] || echo "Maintenance release" > "$NOTES_FILE"

show_plan() {
  echo
  echo "  Version : $CURRENT → $NEW"
  echo "  Since   : ${LAST_TAG:-(first release)}"
  echo "  Notes   :"
  sed 's/^/    - /' "$NOTES_FILE"
  echo
}
show_plan

if (( DRY_RUN )); then
  echo "(dry run — nothing changed)"
  exit 0
fi

while true; do
  read -r -p "Release v$NEW? [y]es / [e]dit notes / [N]o: " answer
  case "${answer,,}" in
    y|yes) break ;;
    e|edit) "${EDITOR:-nano}" "$NOTES_FILE"; show_plan ;;
    *) echo "Cancelled."; exit 1 ;;
  esac
done

# ── Bump, merge, tag, push ───────────────────────────────────────────────────
python3 - "$NEW" "$NOTES_FILE" <<'PY'
import json, sys
from datetime import date
path = "version.json"
data = json.load(open(path, encoding="utf-8"))
data["version"] = sys.argv[1]
data["date"] = date.today().isoformat()
data["notes"] = [line.strip() for line in open(sys.argv[2], encoding="utf-8") if line.strip()]
with open(path, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
    f.write("\n")
PY

git add version.json
git commit --quiet -m "release: v$NEW"
echo "==> Committed version.json on dev"

git checkout --quiet main
git merge --ff-only --quiet origin/main 2>/dev/null || true
git merge --no-ff --quiet dev -m "Merge dev into main for v$NEW"
git tag -a "v$NEW" -m "v$NEW"
echo "==> Merged dev → main and tagged v$NEW"

git push --quiet origin main dev "v$NEW"
echo "==> Pushed main, dev and tag v$NEW"

# Keep dev level with main so the next merge is clean
git checkout --quiet dev
git merge --ff-only --quiet main

UPDATE_URL="$(python3 -c 'import json; print(json.load(open("version.json")).get("update_url", ""))')"
echo
echo "✓ Released v$NEW"
echo
echo "Last step — paste this into your Gist's version.json so friends see the update:"
if [[ -n "$UPDATE_URL" ]]; then
  echo "  $(echo "${UPDATE_URL%/raw/*}" | sed 's#gist.githubusercontent.com#gist.github.com#')"
fi
echo "────────────────────────────────────────"
cat version.json
echo "────────────────────────────────────────"

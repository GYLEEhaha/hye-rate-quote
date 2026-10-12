#!/usr/bin/env bash
# Verify static site is ready for GitHub Pages (repo root as site root).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

need=(
  index.html
  static/style.css
  static/app.js
  data/parsed/index.json
  data/parsed/meta.json
  data/parsed/fba_codes.json
)

missing=0
for f in "${need[@]}"; do
  if [[ ! -f "$f" ]]; then
    echo "MISSING: $f"
    missing=1
  fi
done

# At least one rates_*.json
shopt -s nullglob
rates=(data/parsed/rates_*.json)
if [[ ${#rates[@]} -eq 0 ]]; then
  echo "MISSING: data/parsed/rates_*.json — run: python3 scripts/parse_rates.py"
  missing=1
else
  echo "OK: ${#rates[@]} rate JSON file(s)"
fi

if [[ -f data/rates.xlsx ]]; then
  echo "NOTE: data/rates.xlsx present locally (gitignored) — do not force-add to public repo."
fi

if [[ "$missing" -ne 0 ]]; then
  echo "prepare_pages: FAILED — fix missing files above."
  exit 1
fi

echo "prepare_pages: ready for GitHub Pages from repo root."
echo "  Publish: index.html + static/ + data/parsed/"
echo "  Pages setting: branch main, folder / (root)"
echo "  Relative paths used — works at https://<user>.github.io/<repo>/"

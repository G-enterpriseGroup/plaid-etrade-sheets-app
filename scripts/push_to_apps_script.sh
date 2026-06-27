#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."

if ! command -v clasp >/dev/null 2>&1; then
  echo "clasp is not installed. Run: npm install -g @google/clasp"
  exit 1
fi

if [ ! -f "apps-script/.clasp.json" ]; then
  echo "Missing apps-script/.clasp.json. Run: clasp clone-script YOUR_SCRIPT_ID --rootDir apps-script"
  exit 1
fi

cd apps-script
clasp push

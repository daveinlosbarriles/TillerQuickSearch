#!/usr/bin/env bash
# Sync Amazon Orders script + dialog from tiller-tools (clasp root) into the
# separate TillerAmazonOrdersCSVImport GitHub repo clone.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
MIRROR="${AMAZON_MIRROR:-$ROOT/TillerAmazonOrdersCSVImport}"
if [[ ! -d "$MIRROR" ]]; then
  echo "Mirror folder not found: $MIRROR" >&2
  echo "Clone: git clone https://github.com/daveinlosbarriles/TillerAmazonOrdersCSVImport.git TillerAmazonOrdersCSVImport" >&2
  exit 1
fi
cp "$ROOT/amazonorders.gs" "$MIRROR/AmazonOrders.js"
cp "$ROOT/AmazonOrdersDialog.html" "$MIRROR/AmazonOrdersDialog.html"
echo "Synced AmazonOrders.js and AmazonOrdersDialog.html -> $MIRROR"
echo "Next: cd \"$MIRROR\" && git status"

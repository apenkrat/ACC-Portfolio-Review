#!/usr/bin/env bash
# ACC Portfolio Review — setup checker and launcher
# Usage:
#   ./run.sh                          # interactive prompts
#   ./run.sh --region tmt             # AMER TMT, interactive format/output
#   ./run.sh --region all --format html --output ~/Desktop
#   ./run.sh --region cbs --format all --sf-alias myorg

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON=$(command -v python3 2>/dev/null || true)

# ── 1. Python check ────────────────────────────────────────────────────────────
if [[ -z "$PYTHON" ]]; then
  echo "ERROR: python3 not found. Install from https://www.python.org or via Homebrew:"
  echo "  brew install python"
  exit 1
fi

PYTHON_VER=$("$PYTHON" -c 'import sys; print(sys.version_info[:2] >= (3,9))')
if [[ "$PYTHON_VER" != "True" ]]; then
  echo "ERROR: Python 3.9+ required. Found: $($PYTHON --version)"
  exit 1
fi

# ── 2. Dependencies check ──────────────────────────────────────────────────────
MISSING=()
for pkg in simple_salesforce docx pptx lxml; do
  mod="${pkg/simple_salesforce/simple_salesforce}"
  mod="${mod/docx/docx}"
  mod="${mod/pptx/pptx}"
  "$PYTHON" -c "import $mod" 2>/dev/null || MISSING+=("$pkg")
done

if [[ ${#MISSING[@]} -gt 0 ]]; then
  echo "Installing missing Python packages..."
  "$PYTHON" -m pip install -r "$SCRIPT_DIR/requirements.txt" --quiet
fi

# ── 3. sf CLI check ────────────────────────────────────────────────────────────
if ! command -v sf &>/dev/null; then
  echo "ERROR: Salesforce CLI (sf) not found."
  echo "  Install: npm install -g @salesforce/cli"
  echo "  Then authenticate: sf org login web --alias org62"
  exit 1
fi

# Determine alias (--sf-alias arg or default org62)
SF_ALIAS="org62"
for i in "$@"; do
  if [[ "$i" == --sf-alias=* ]]; then SF_ALIAS="${i#--sf-alias=}"; fi
done
for ((i=1; i<=$#; i++)); do
  if [[ "${!i}" == "--sf-alias" ]]; then j=$((i+1)); SF_ALIAS="${!j}"; fi
done

if ! sf org display --target-org "$SF_ALIAS" --json &>/dev/null; then
  echo "ERROR: Not authenticated to Salesforce org '$SF_ALIAS'."
  echo "  Run: sf org login web --alias $SF_ALIAS"
  exit 1
fi

# ── 4. Launch ──────────────────────────────────────────────────────────────────
echo "All checks passed. Starting ACC Portfolio Review..."
echo ""
exec "$PYTHON" "$SCRIPT_DIR/run_acc_audit.py" "$@"

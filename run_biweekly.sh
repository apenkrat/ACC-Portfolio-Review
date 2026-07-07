#!/bin/bash
# TMT Bi-Weekly Report — audit (TXT) + template-filled PPTX
# Runs unattended (cron) or manually: ./run_biweekly.sh
set -euo pipefail
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
LOG="$SCRIPT_DIR/export.log"

echo "──────────────────────────────────────────────────────" >> "$LOG"
echo "$(date)  run_biweekly.sh START" >> "$LOG"

echo "=== $(date): Starting TMT Bi-Weekly Report ===" 2>&1 | tee -a "$LOG"

# Step 1: Audit (TXT)
echo "--- Step 1: Audit ---" | tee -a "$LOG"
python3 "$SCRIPT_DIR/run_acc_audit.py" 1 2>&1 | tee -a "$LOG"

# Step 2: Template PPTX
echo "--- Step 2: Template PPTX ---" | tee -a "$LOG"
python3 "$SCRIPT_DIR/generate_template_pptx.py" 2>&1 | tee -a "$LOG"

echo "=== $(date): Done ===" 2>&1 | tee -a "$LOG"

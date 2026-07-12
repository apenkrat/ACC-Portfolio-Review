#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENV_FILE="$SCRIPT_DIR/.env"
if [[ -f "$ENV_FILE" ]]; then
  set -a
  source "$ENV_FILE"
  set +a
fi

TILE_ID="${TILE_ID:-817}"
PAGE_HOST_URL="${PAGE_HOST_URL:-https://single-html-page-app-host-07cda8a7041b.herokuapp.com}"
PAGE_HOST_TOKEN="${PAGE_HOST_TOKEN:-}"
OUTPUT_DIR="${OUTPUT_DIR:-$SCRIPT_DIR/generated}"
VERSION_KIND="${VERSION_KIND:-minor}"
LOG_FILE="${LOG_FILE:-$SCRIPT_DIR/export.log}"
DRY_RUN=0

for arg in "$@"; do
  case "$arg" in
    --dry-run)
      DRY_RUN=1
      ;;
    --tile-id=*)
      TILE_ID="${arg#*=}"
      ;;
    --output-dir=*)
      OUTPUT_DIR="${arg#*=}"
      ;;
    --version-kind=*)
      VERSION_KIND="${arg#*=}"
      ;;
  esac
done

mkdir -p "$OUTPUT_DIR" "$(dirname "$LOG_FILE")"

# Truncate log to last 500 lines to prevent unbounded growth
if [[ -f "$LOG_FILE" ]]; then
  tail -500 "$LOG_FILE" > "${LOG_FILE}.tmp" && mv "${LOG_FILE}.tmp" "$LOG_FILE" || rm -f "${LOG_FILE}.tmp"
fi

{
  echo "=== $(date '+%Y-%m-%d %H:%M:%S') ==="
  echo "Starting publish run..."
} >> "$LOG_FILE"

# ── Push a single data file to Page Host Data Push API ───────────────────────
push_data_file() {
  local filepath="$1"
  local filename
  filename="$(basename "$filepath")"
  local hash_file="$SCRIPT_DIR/.last_$(echo "$filename" | tr '.' '_')_hash"
  local current_hash
  current_hash=$(shasum -a 256 "$filepath" | awk '{print $1}')
  local previous_hash=""
  [[ -f "$hash_file" ]] && previous_hash=$(cat "$hash_file")

  if [[ "$previous_hash" == "$current_hash" ]]; then
    echo "No change in $filename; skipping data push." >> "$LOG_FILE"
    return 0
  fi

  echo "Pushing data file: $filename" >> "$LOG_FILE"
  if [[ "$DRY_RUN" == "1" ]]; then
    echo "Dry run: would PUT $filename to tile $TILE_ID" >> "$LOG_FILE"
    return 0
  fi

  if curl -sS --fail -X PUT "$PAGE_HOST_URL/api/uploads/$TILE_ID/data/$filename" \
    -H "Authorization: Bearer $PAGE_HOST_TOKEN" \
    -F "file=@$filepath;type=application/json" >> "$LOG_FILE" 2>&1; then
    printf '%s\n' "$current_hash" > "$hash_file"
    echo "Data push complete: $filename" >> "$LOG_FILE"
  else
    echo "ERROR: Data push failed for $filename (exit $?)" >> "$LOG_FILE"
  fi
}

# ── Build zip bundle (index.html → acc_portfolio_bundle.zip) ─────────────────
build_bundle() {
  local html_file="$1"
  local zip_path="$OUTPUT_DIR/acc_portfolio_bundle.zip"
  # Copy to index.html for bundle entry point
  cp "$html_file" "$OUTPUT_DIR/index.html"
  (cd "$OUTPUT_DIR" && zip -jq "$zip_path" index.html)
  echo "$zip_path"
}

# ── Fetch current version for embedding in HTML ───────────────────────────────
if [[ -n "$PAGE_HOST_TOKEN" ]]; then
  _ver_json=$(curl -s "$PAGE_HOST_URL/api/uploads/$TILE_ID" -H "Authorization: Bearer $PAGE_HOST_TOKEN" 2>/dev/null || true)
  _major=$(echo "$_ver_json" | python3 -c "import sys,re; raw=sys.stdin.read(); m=re.search(r'\"version_major\":\s*(\d+)',raw); print(m.group(1) if m else '1')" 2>/dev/null || echo "1")
  _minor=$(echo "$_ver_json" | python3 -c "import sys,re; raw=sys.stdin.read(); m=re.search(r'\"version_minor\":\s*(\d+)',raw); print(m.group(1) if m else '0')" 2>/dev/null || echo "0")
  export NEXT_HTML_VERSION="${_major}.$((${_minor}+1))"
  echo "Next HTML version: $NEXT_HTML_VERSION" >> "$LOG_FILE"
fi

# ── Generate data ─────────────────────────────────────────────────────────────
echo "Generating TMT HTML + data..." >> "$LOG_FILE"
"$SCRIPT_DIR/run.sh" --region tmt --format html --output "$OUTPUT_DIR" --sf-alias "${SF_ALIAS:-org62}" >> "$LOG_FILE" 2>&1
echo "Generating CBS HTML + data..." >> "$LOG_FILE"
"$SCRIPT_DIR/run.sh" --region cbs --format html --output "$OUTPUT_DIR" --sf-alias "${SF_ALIAS:-org62}" >> "$LOG_FILE" 2>&1
echo "Combining TMT + CBS into ACC HTML..." >> "$LOG_FILE"
python3 "$SCRIPT_DIR/combine_html.py" --output "$OUTPUT_DIR" >> "$LOG_FILE" 2>&1

# ── Push data files ───────────────────────────────────────────────────────────
if [[ -z "$PAGE_HOST_TOKEN" ]]; then
  echo "PAGE_HOST_TOKEN is not set. Set it in .env or export it." >&2
  exit 1
fi

for json_file in "$OUTPUT_DIR"/acc_*_data.json; do
  [[ -f "$json_file" ]] && push_data_file "$json_file"
done

# ── HTML bundle upload ────────────────────────────────────────────────────────
latest_html=$(find "$OUTPUT_DIR" -type f -name 'ACC_Audit_*.html' | sort | tail -1)
if [[ -z "$latest_html" ]]; then
  echo "No HTML file generated." >&2
  exit 1
fi

latest_html_hash=$(shasum -a 256 "$latest_html" | awk '{print $1}')
state_file="$SCRIPT_DIR/.last_uploaded_html"
previous_hash=""
[[ -f "$state_file" ]] && previous_hash=$(cat "$state_file")

echo "Using HTML file: $latest_html" >> "$LOG_FILE"
echo "HTML hash: $latest_html_hash" >> "$LOG_FILE"

if [[ "$previous_hash" == "$latest_html_hash" ]]; then
  echo "No HTML change detected; skipping bundle upload." >> "$LOG_FILE"
  exit 0
fi

if [[ "$DRY_RUN" == "1" ]]; then
  echo "Dry run: would upload bundle for $latest_html to tile $TILE_ID" | tee -a "$LOG_FILE"
  echo "Dry run complete." | tee -a "$LOG_FILE"
  exit 0
fi

bundle_zip=$(build_bundle "$latest_html")
echo "Uploading bundle: $bundle_zip" >> "$LOG_FILE"

if curl -sS --fail -X POST "$PAGE_HOST_URL/api/uploads/$TILE_ID/version" \
  -H "Authorization: Bearer $PAGE_HOST_TOKEN" \
  -F "file=@$bundle_zip" \
  -F "kind=$VERSION_KIND" >> "$LOG_FILE" 2>&1; then
  printf '%s\n' "$latest_html_hash" > "$state_file"
  echo "Bundle upload complete." >> "$LOG_FILE"
else
  echo "ERROR: Bundle upload failed (exit $?)" >> "$LOG_FILE"
  exit 1
fi

# Prune old timestamped HTML files, keep newest 5 per prefix
for prefix in ACC_Audit AMER_TMT_Audit AMER_CBS_Audit; do
  ls -t "$OUTPUT_DIR/${prefix}_"20[0-9][0-9]*.html 2>/dev/null | tail -n +6 | while IFS= read -r f; do
    rm -- "$f" && echo "Pruned: $(basename "$f")" >> "$LOG_FILE"
  done
done

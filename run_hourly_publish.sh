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
  tail -500 "$LOG_FILE" > "${LOG_FILE}.tmp" && mv "${LOG_FILE}.tmp" "$LOG_FILE"
fi

{
  echo "=== $(date '+%Y-%m-%d %H:%M:%S') ==="
  echo "Generating TMT HTML report..."
} >> "$LOG_FILE"

# Fetch current version from page host and compute next version for embedding in HTML
if [[ -n "$PAGE_HOST_TOKEN" ]]; then
  _ver_json=$(curl -s "$PAGE_HOST_URL/api/uploads/$TILE_ID" -H "Authorization: Bearer $PAGE_HOST_TOKEN" 2>/dev/null || true)
  _major=$(echo "$_ver_json" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('upload',{}).get('version_major',1))" 2>/dev/null || echo "1")
  _minor=$(echo "$_ver_json" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('upload',{}).get('version_minor',0))" 2>/dev/null || echo "0")
  export NEXT_HTML_VERSION="${_major}.$((${_minor}+1))"
  echo "Next HTML version: $NEXT_HTML_VERSION" >> "$LOG_FILE"
fi

"$SCRIPT_DIR/run.sh" --region tmt --format html --output "$OUTPUT_DIR" --sf-alias "${SF_ALIAS:-org62}" >> "$LOG_FILE" 2>&1

latest_html=$(find "$OUTPUT_DIR" -type f -name 'AMER_TMT_Audit_*.html' | sort | tail -1)
if [[ -z "$latest_html" ]]; then
  echo "No HTML file generated." >&2
  exit 1
fi

latest_html_name=$(basename "$latest_html")
latest_html_hash=$(shasum -a 256 "$latest_html" | awk '{print $1}')
state_file="$SCRIPT_DIR/.last_uploaded_html"
previous_hash=""
if [[ -f "$state_file" ]]; then
  previous_hash=$(cat "$state_file")
fi

echo "Using HTML file: $latest_html" >> "$LOG_FILE"

echo "HTML hash: $latest_html_hash" >> "$LOG_FILE"

if [[ "$previous_hash" == "$latest_html_hash" ]]; then
  echo "No content change detected; skipping upload." >> "$LOG_FILE"
  exit 0
fi

if [[ "$DRY_RUN" == "1" ]]; then
  echo "Dry run: would upload $latest_html to tile $TILE_ID" | tee -a "$LOG_FILE"
  echo "Dry run complete." | tee -a "$LOG_FILE"
  exit 0
fi

if [[ -z "$PAGE_HOST_TOKEN" ]]; then
  echo "PAGE_HOST_TOKEN is not set. Set it in .env or export it." >&2
  exit 1
fi

curl -sS -X POST "$PAGE_HOST_URL/api/uploads/$TILE_ID/version" \
  -H "Authorization: Bearer $PAGE_HOST_TOKEN" \
  -F "file=@$latest_html" \
  -F "kind=$VERSION_KIND" >> "$LOG_FILE" 2>&1

printf '%s\n' "$latest_html_hash" > "$state_file"
echo "Upload complete." >> "$LOG_FILE"

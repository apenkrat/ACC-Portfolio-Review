# ACC-Portfolio-Review

## Setup Guide

### What this does

| Script | Purpose | Output |
|---|---|---|
| `run_acc_audit.py` | Full portfolio audit — pulls live Org62 data, applies all SPSM/DAF rules | `AMER_TMT_Audit_YYYY-MM-DD.{txt,docx,pptx,html}` to Google Drive |
| `run.sh` | Setup checker and launcher — verifies dependencies, then runs the audit | Same as above |
| `generate_template_pptx.py` | Fills the bi-weekly PPTX template with current account cards | `TMT_Bi-Weekly_YYYY-MM-DD.pptx` to Google Drive |
| `run_biweekly.sh` | Wrapper: runs audit (TXT) + template PPTX together | — |
| `export_acc_pulses.sh` | Daily pulse CSV export | `TMT_Project_Pulses_YYYY-MM-DD.csv` to Google Drive |
| `export_acc_resources.sh` | Daily resource assignment CSV export | `TMT_Project_Resources_YYYY-MM-DD.csv` + `_LATEST.csv` to Google Drive |

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                        Data Sources                                 │
│  Salesforce Org62 (SOQL via sf CLI)  ·  Pulse records  ·  Resources│
└──────────────────────────────┬──────────────────────────────────────┘
                               │ sf apex run / sf data query
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    run_acc_audit.py  (Python)                       │
│                                                                     │
│  ┌─────────────┐   ┌──────────────┐   ┌──────────────────────────┐ │
│  │ SOQL queries│   │ Rule engine  │   │   write_html()           │ │
│  │ (TMT / CBS) │──▶│ (SPSM / DAF) │──▶│   Generates self-       │ │
│  └─────────────┘   └──────────────┘   │   contained HTML + JS   │ │
│                                        └────────────┬─────────────┘ │
│  Outputs: .txt  .docx  .pptx  .html                │               │
└────────────────────────────────────────────────────┼───────────────┘
                                                      │ HTML + JSON
                               ┌──────────────────────▼──────────────┐
                               │        combine_html.py               │
                               │  Merges TMT + CBS JSON into one      │
                               │  ACC HTML  (_INLINE embedded data)   │
                               └──────────────────────┬──────────────┘
                                                       │
                               ┌───────────────────────▼─────────────┐
                               │      run_hourly_publish.sh           │
                               │                                      │
                               │  1. Generate TMT + CBS HTML/JSON     │
                               │  2. combine_html.py → ACC HTML       │
                               │  3. Push JSON data files             │
                               │     PUT /api/uploads/817/data/…      │
                               │  4. Upload HTML bundle (zip)         │
                               │     POST /api/uploads/817/version    │
                               │  5. Prune old files (keep 5)         │
                               └───────────────────────┬─────────────┘
                                                        │
                               ┌────────────────────────▼────────────┐
│                   Page Host Platform (Heroku)                      │
│                   https://…herokuapp.com  ·  Tile 817              │
│                                                                     │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │  Served HTML  (ACC Delivery Portfolio — Interactive Dashboard)│  │
│  │                                                              │  │
│  │  ┌─────────────┐  ┌──────────────┐  ┌────────────────────┐  │  │
│  │  │ _INLINE JSON│  │ /_data/ JSON │  │  SQLite DB         │  │  │
│  │  │ (embedded   │  │ (pushed data │  │  (tier/PO assign-  │  │  │
│  │  │  at build)  │  │  files)      │  │   ments, live RW)  │  │  │
│  │  └─────────────┘  └──────────────┘  └────────────────────┘  │  │
│  │                                                              │  │
│  │  ┌─────────────────────────────────────────────────────┐    │  │
│  │  │  LLM Proxy  /api/proxy/llm/817                      │    │  │
│  │  │  Business Overview button → powerful / balanced /   │    │  │
│  │  │  fast tier cascade (Salesforce Bedrock + Heroku)    │    │  │
│  │  └─────────────────────────────────────────────────────┘    │  │
│  └──────────────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────────────┘
```

### Key Components

| Component | File | Role |
|---|---|---|
| Audit engine | `run_acc_audit.py` | Pulls Org62 data, applies rules, generates HTML/JSON/DOCX/PPTX |
| Dashboard HTML/JS | `run_acc_audit.py → write_html()` | ~4,000-line self-contained SPA: filters, scorecard, group-by, assignments, AI briefing |
| Combine step | `combine_html.py` | Merges TMT + CBS JSON into a single ACC HTML with patched PO dropdowns |
| Publish automation | `run_hourly_publish.sh` | End-to-end: generate → combine → push data → upload bundle; hash-gated to skip unchanged files |
| Assignment store | Page Host SQLite DB (tile 817) | Server-side `assignments` table with `pid TEXT PRIMARY KEY`; written via `db/write`, read via `TileDB.serverQuery()` |
| LLM proxy | Page Host `/api/proxy/llm/817` | AI Business Overview; cascades `powerful → balanced → fast` tiers; auth via `window.__PROXY_TOKEN__` injected at page load |
| Data files | `generated/acc_amer_*_data.json` | Pushed to `/_data/` on Page Host; loaded at runtime if `_INLINE` is stale |

### Data Flow (runtime, in the browser)

1. Page loads → `_INLINE` JSON (baked into HTML at build time) populates `RAW[]`
2. `loadAssignments()` fetches live tier/PO overrides from the SQLite DB via `TileDB.serverQuery()`
3. `applyFilters()` runs on every user interaction — filters `RAW[]`, updates scorecard KPIs, re-renders the table
4. Edit Assignment dialog → `saveAssignment()` upserts to SQLite → updates `RAW[]` in-place → re-renders
5. Business Overview button → `openGMOverview()` → LLM proxy call → markdown rendered in modal

---

## Data Lineage

### Source fields (Salesforce Org62 → `run_acc_audit.py`)

| Dashboard Field | Salesforce Object / Field | Notes |
|---|---|---|
| Project Name | `Opportunity.Name` | |
| Account | `Opportunity.Account.Name` | |
| Region | `Opportunity.Owner.Region__c` | Used to split TMT vs CBS |
| Status | `PS_Project__c.Project_Status__c` | Red / Yellow / Green / Watermelon / No Pulse / On Hold |
| Bookings | `Opportunity.Amount` | Total contract value |
| Billings | `PS_Project__c.Total_Billed__c` | Cumulative invoiced to date |
| FAR (Financial At Risk) | Derived: `Bookings - Billings - Scheduled_Hours * Bill_Rate` | Budget remaining vs scheduled work |
| Overdue Invoices | `Invoice__c` records where `Due_Date__c < today` | Aggregated per project |
| Revenue @ Risk (RR) | `PS_Project__c.Revenue_At_Risk__c` | PM-entered field |
| Unscheduled Backlog | Derived: `FAR - (Remaining_Hours * Bill_Rate)` | Budget with no resource scheduled |
| Bid Margin % | `Opportunity.Bid_Margin__c` | Original sold margin |
| Margin @ Close % | `PS_Project__c.Margin_At_Close__c` | Current delivery margin estimate |
| GDC % | `Resource_Assignment__c` aggregated by offshore flag | Offshore hours / total hours |
| Data Quality Score | Derived rule engine score (0–100) | Based on pulse staleness, missing fields, rule violations |
| CSAT | `CSAT_Survey__c.Score__c` | Latest survey per project |
| High Watch | `PS_Project__c.High_Watch__c` | Boolean flag set by PM |
| Tier | `assignments` SQLite table (Page Host DB) | Manually assigned via dashboard; not in Salesforce |
| Portfolio Owner | `assignments` SQLite table (Page Host DB) | Manually assigned via dashboard; not in Salesforce |
| Pulse / Leadership Notes | `PS_Project_Pulse__c.Notes__c` | Latest pulse per project, used for AI context |

### Transformation steps

```
Salesforce Org62 (SOQL)
    │
    │  run_acc_audit.py
    ├─ SOQL queries pull raw Opportunity + PS_Project + Resource + Invoice + CSAT records
    ├─ Rule engine evaluates SPSM / DAF rules → assigns rules[] flags per project
    ├─ Derived fields computed (FAR, Unscheduled Backlog, GDC %, Data Quality Score)
    ├─ Output: acc_amer_tmt_data.json  /  acc_amer_cbs_data.json
    │          (rows[]: one object per project, all fields flattened)
    │
    │  combine_html.py
    ├─ Merges TMT + CBS rows[] arrays → combined ACC dataset
    ├─ Embeds as _INLINE JSON constant inside the HTML file
    ├─ Also rebuilds PO dropdown from merged po values
    │
    │  Page Host (runtime, browser)
    ├─ _INLINE loaded into RAW[] array on page load
    ├─ loadAssignments() overlays tier/po from SQLite DB (server-side, manual overrides)
    ├─ applyFilters() derives all KPI scorecard values from RAW[] on every user interaction
    └─ buildGMPrompt() aggregates RAW[] into a structured text block for LLM calls
```

### What is NOT in Salesforce

| Field | Where it lives | How it's set |
|---|---|---|
| Tier (T1 / T2 / T3) | Page Host SQLite DB, tile 817, `assignments` table | Edit Assignment dialog in the dashboard |
| Portfolio Owner | Page Host SQLite DB, tile 817, `assignments` table | Edit Assignment dialog in the dashboard |

These two fields are the only ones that write back — everything else is read-only from Salesforce.

---

## Quickstart

### 1. Install the Salesforce CLI
```bash
brew install salesforce-cli
sf --version
```

### 2. Authenticate to Org62
```bash
sf org login web --alias org62 --instance-url https://org62.my.salesforce.com
```
Verify:
```bash
sf org display --target-org org62
```

### 3. Python 3.9+
```bash
python3 --version
# install if needed:
brew install python
```

### 4. Install Python dependencies
```bash
pip3 install -r requirements.txt
```
Or skip this step — `run.sh` installs missing packages automatically.

### 5. Google Drive for Desktop
Outputs write to your Google Drive by default. Make sure Google Drive for Desktop is installed, running, and signed in. You will be prompted to confirm or change the output directory on first run.

---

## Running the audit

### Simplest — interactive prompts
```bash
./run.sh
```
`run.sh` checks all dependencies before launching. You'll be prompted for region, format, and output directory.

### With arguments — skip the prompts
```bash
./run.sh --region tmt --format html
./run.sh --region cbs --format all --output ~/Desktop
./run.sh --region all --format all
```

### Arguments reference

| Argument | Short | Values | Default |
|---|---|---|---|
| `--region` | `-r` | `tmt`, `cbs`, `all` | interactive prompt |
| `--format` | `-f` | `txt`, `docx`, `pptx`, `html`, `all` | interactive prompt |
| `--output` | `-o` | any directory path | last-used or Google Drive |
| `--sf-alias` | — | any `sf` org alias | `org62` |

### Using a different Salesforce org alias
If your org is authenticated under a different alias:
```bash
./run.sh --region tmt --format all --sf-alias myorg
```

### Running the script directly (without run.sh)
```bash
python3 run_acc_audit.py --region tmt --format html
```

---

## Running the bi-weekly deck

```bash
python3 ~/Documents/claude/ACC-Portfolio-Review/generate_template_pptx.py
```

This copies the latest bi-weekly PPTX template from `~/Downloads/`, fills every account card with current data, and saves `TMT_Bi-Weekly_YYYY-MM-DD.pptx` to Google Drive.

To run both the audit and the deck in one step:
```bash
~/Documents/claude/ACC-Portfolio-Review/run_biweekly.sh
```

---

## Scheduled runs (crontab)

```bash
# Daily 6 AM — pulse + resource CSV exports
0 6 * * *  .../export_acc_pulses.sh    >> .../export.log 2>&1
0 6 * * *  .../export_acc_resources.sh >> .../export.log 2>&1

# Every Wednesday 7 AM — audit TXT + template PPTX
0 7 * * 3  .../run_biweekly.sh         >> .../export.log 2>&1
```

All scripts are in `/Users/apenkrat/Documents/claude/ACC-Portfolio-Review/`. Logs: `export.log` in the same directory.

To view/edit the schedule:
```bash
crontab -l       # view
crontab -e       # edit
```

---

## Output files

All outputs land in your configured output directory (Google Drive by default).

| File | Generated by |
|---|---|
| `AMER_TMT_Audit_YYYY-MM-DD.txt` | `run_acc_audit.py` |
| `AMER_TMT_Audit_YYYY-MM-DD.docx` | `run_acc_audit.py` |
| `AMER_TMT_Audit_YYYY-MM-DD.pptx` | `run_acc_audit.py` |
| `AMER_TMT_Audit_YYYY-MM-DD.html` | `run_acc_audit.py` |
| `TMT_Bi-Weekly_YYYY-MM-DD.pptx` | `generate_template_pptx.py` |
| `TMT_Project_Pulses_YYYY-MM-DD.csv` | `export_acc_pulses.sh` |
| `TMT_Project_Resources_YYYY-MM-DD.csv` | `export_acc_resources.sh` |
| `TMT_Project_Resources_LATEST.csv` | `export_acc_resources.sh` |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `sf: command not found` | Add Homebrew to PATH: `export PATH="/opt/homebrew/bin:$PATH"` |
| `Auth failed` / `INVALID_SESSION_ID` | Re-authenticate: `sf org login web --alias org62` |
| `ModuleNotFoundError: simple_salesforce` | `pip3 install simple-salesforce` |
| `ModuleNotFoundError: docx` | `pip3 install python-docx` |
| `ModuleNotFoundError: pptx` | `pip3 install python-pptx` |
| Output dir not found | Confirm Google Drive for Desktop is running and synced |
| 0 rows returned | Check session: `sf org display --target-org org62` |
| Cron job not running | Check log: `tail -50 ~/Documents/claude/ACC-Portfolio-Review/export.log` |

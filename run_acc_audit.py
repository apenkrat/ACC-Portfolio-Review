5#!/usr/bin/env python3
"""
{REGION_LABEL} Portfolio Audit — standalone, no Excel dependency.
Pulls live data from Org62 via sf CLI, applies all SPSM/DAF rules,
writes a dated audit report to Google Drive TMT Reports folder.

Usage:  python3 ~/Documents/claude/tmt-reports/run_acc_audit.py
"""
import subprocess, json, sys, os, re
from datetime import date, datetime, timedelta
from collections import defaultdict

TODAY = date.today()
REPORT_DATE = datetime.now().strftime('%Y-%m-%d %H:%M')

# ── Output directory — always prompted, last value saved as default ───────────
_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.tmt_config')
_DEFAULT_DIR = "/Users/apenkrat/Library/CloudStorage/GoogleDrive-apenkrat@salesforce.com/My Drive/TMT Reports"

def _load_output_dir():
    if os.path.exists(_CONFIG_FILE):
        with open(_CONFIG_FILE) as _f:
            for _line in _f:
                if _line.startswith('output_dir='):
                    return _line.strip().split('=', 1)[1].strip().strip("'\"")
    return None

def _save_output_dir(path):
    lines = []
    if os.path.exists(_CONFIG_FILE):
        with open(_CONFIG_FILE) as _f:
            lines = [l for l in _f if not l.startswith('output_dir=')]
    lines.append(f'output_dir={path}\n')
    with open(_CONFIG_FILE, 'w') as _f:
        _f.writelines(lines)

# ── Slack intelligence cache (written by Claude Code, read by the script) ─────
# Run `/slack-intel` in Claude Code to refresh before generating reports.
_SLACK_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), f'slack_intel_{date.today().isoformat()}.json')

import argparse as _argparse

# ── CLI argument parsing ──────────────────────────────────────────────────────
_parser = _argparse.ArgumentParser(
    prog='run_acc_audit.py',
    description='ACC Portfolio Review — pull live Org62 data and generate audit reports.',
    formatter_class=_argparse.RawTextHelpFormatter,
)
_parser.add_argument('--region', '-r', metavar='REGION',
    help='Region to audit: "tmt", "cbs", or "all" (default: interactive prompt)')
_parser.add_argument('--format', '-f', metavar='FORMAT', dest='fmt',
    help='Output format: txt, docx, pptx, html, or all (default: interactive prompt)')
_parser.add_argument('--output', '-o', metavar='DIR',
    help='Output directory path (default: last-used or Google Drive)')
_parser.add_argument('--sf-alias', metavar='ALIAS', default=None,
    help='Salesforce CLI alias to use (default: org62)')
_parser.add_argument('--data-only', action='store_true', default=False,
    help='Write JSON data file(s) only; skip HTML/TXT/DOCX/PPTX generation')
_CLI = _parser.parse_args()

# ── Output directory ──────────────────────────────────────────────────────────
_last_dir = _load_output_dir() or _DEFAULT_DIR
if _CLI.output:
    OUTPUT_DIR = _CLI.output
    _save_output_dir(OUTPUT_DIR)
elif not sys.stdin.isatty():
    OUTPUT_DIR = _last_dir
else:
    print(f"\n📁  Output directory [{_last_dir}]: ", end='', flush=True)
    _entered = input().strip()
    OUTPUT_DIR = _entered if _entered else _last_dir
    _save_output_dir(OUTPUT_DIR)
    print(f"    → Saving to: {OUTPUT_DIR}\n")

# ── Output format ─────────────────────────────────────────────────────────────
_FMT_MAP = {'txt': ['txt'], 'docx': ['docx'], 'pptx': ['pptx'], 'html': ['html'],
            'all': ['txt','docx','pptx','html']}
if _CLI.data_only:
    OUTPUT_FORMATS = ['json']
elif _CLI.fmt:
    OUTPUT_FORMATS = _FMT_MAP.get(_CLI.fmt.lower(), ['txt'])
elif not sys.stdin.isatty():
    OUTPUT_FORMATS = ['txt']
else:
    print("Select output format:")
    print("  1 — TXT  (plain text, default)")
    print("  2 — DOCX (Word document)")
    print("  3 — PPTX (PowerPoint)")
    print("  4 — HTML (interactive dashboard)")
    print("  5 — ALL  (all four formats)")
    _choice = input("Enter choice [1]: ").strip() or '1'
    OUTPUT_FORMATS = {'1':['txt'],'2':['docx'],'3':['pptx'],'4':['html'],'5':['txt','docx','pptx','html']}.get(_choice, ['txt'])
print(f"  → Generating: {', '.join(f.upper() for f in OUTPUT_FORMATS)}\n")

# ── Region registry — add new regions here ────────────────────────────────────
# Each entry covers all subregions for that region. To add a new region,
# append a dict with: key, label, subregions, pipe_like, slug, meta_file.
REGION_REGISTRY = [
    {'key': 'AMER TMT', 'label': 'AMER TMT', 'subregions': ['AMER TMT - 1', 'AMER TMT - 2'], 'pipe_like': 'AMER TMT%', 'slug': 'AMER_TMT', 'meta_file': 'acc_portfolio_metadata.md'},
    {'key': 'AMER CBS', 'label': 'AMER CBS', 'subregions': ['AMER CBS - 1', 'AMER CBS - 2'], 'pipe_like': 'AMER CBS%', 'slug': 'AMER_CBS', 'meta_file': 'acc_portfolio_metadata.md'},
]
_REGION_SLUG_MAP = {'tmt': 'AMER TMT', 'cbs': 'AMER CBS', 'all': None}

# ── Region selection ──────────────────────────────────────────────────────────
if _CLI.region:
    _rkey = _CLI.region.lower()
    if _rkey == 'all':
        SELECTED_REGION_KEYS = [_r['key'] for _r in REGION_REGISTRY]
    elif _rkey in _REGION_SLUG_MAP:
        SELECTED_REGION_KEYS = [_REGION_SLUG_MAP[_rkey]]
    else:
        # allow passing the full key directly e.g. "AMER TMT"
        SELECTED_REGION_KEYS = [_r['key'] for _r in REGION_REGISTRY if _r['key'].lower() == _rkey] or [REGION_REGISTRY[0]['key']]
    print(f"  → Region(s): {', '.join(SELECTED_REGION_KEYS)}\n")
elif not sys.stdin.isatty():
    SELECTED_REGION_KEYS = [REGION_REGISTRY[0]['key']]
else:
    print("Select region to audit:")
    for _i, _r in enumerate(REGION_REGISTRY, 1):
        print(f"  {_i} — {_r['label']}")
    print(f"  {len(REGION_REGISTRY)+1} — ALL regions")
    _rchoice = input("Enter choice [1]: ").strip() or '1'
    if _rchoice == str(len(REGION_REGISTRY)+1):
        SELECTED_REGION_KEYS = [_r['key'] for _r in REGION_REGISTRY]
    elif _rchoice.isdigit() and 1 <= int(_rchoice) <= len(REGION_REGISTRY):
        SELECTED_REGION_KEYS = [REGION_REGISTRY[int(_rchoice)-1]['key']]
    else:
        SELECTED_REGION_KEYS = [REGION_REGISTRY[0]['key']]
    print(f"  → Region(s): {', '.join(SELECTED_REGION_KEYS)}\n")

# ── PM-asserted SWE overrides (project name fragments, lowercase) ─────────────
SWE_PM_OVERRIDES = [
    'aws - afx salesforce rca - csg',
    'disney parks - ccn - r2',
    'nvidia <> slack',
    'nvidia — slack',
]

# ── Auth ───────────────────────────────────────────────────────────────────────
_SF_ALIAS = _CLI.sf_alias or 'org62'

def get_sf():
    try:
        from simple_salesforce import Salesforce
        import re as _re

        # Get instance URL
        display = subprocess.run(
            ['sf', 'org', 'display', '--target-org', _SF_ALIAS, '--json'],
            capture_output=True, text=True, check=True
        )
        instance_url = json.loads(display.stdout)['result']['instanceUrl']

        # sf CLI ≥2.x redacts accessToken in org display — use auth show-access-token
        token_out = subprocess.run(
            ['sf', 'org', 'auth', 'show-access-token', '--target-org', _SF_ALIAS],
            capture_output=True, text=True, input='y\n'
        )
        m = _re.search(r'Access Token\s*│\s*(\S+)', token_out.stdout)
        if not m:
            raise ValueError(f"Could not parse access token. Output: {token_out.stdout[:200]}")
        access_token = m.group(1)

        return Salesforce(instance_url=instance_url, session_id=access_token)
    except Exception as e:
        print(f"Auth failed (alias: {_SF_ALIAS}): {e}")
        sys.exit(1)

sf = get_sf()

def soql(q, label=''):
    print(f'  {label}...')
    recs, result = [], sf.query(q)
    recs.extend(result['records'])
    while not result['done']:
        result = sf.query_more(result['nextRecordsUrl'], True)
        recs.extend(result['records'])
    def flatten(obj, prefix=''):
        row = {}
        for k, v in obj.items():
            if k == 'attributes': continue
            key = f'{prefix}.{k}' if prefix else k
            row.update(flatten(v, key) if isinstance(v, dict) else {key: v})
        return row
    rows = [flatten(r) for r in recs]
    print(f'    -> {len(rows)} rows')
    return rows

def to_f(v):
    if v is None: return None
    try: return float(str(v).replace(',', '').replace('$', '').strip())
    except: return None

def to_d(v):
    if not v: return None
    if isinstance(v, (date, datetime)): return v.date() if isinstance(v, datetime) else v
    for fmt in ['%Y-%m-%d', '%m/%d/%Y']:
        try: return datetime.strptime(str(v)[:10], fmt).date()
        except: pass
    return None

# ── Pull data ─────────────────────────────────────────────────────────────────
_all_results = []  # accumulated across all region iterations for combined HTML
for _RC in [_r for _r in REGION_REGISTRY if _r['key'] in SELECTED_REGION_KEYS]:
    REGION_LABEL       = _RC['label']
    REGION_SLUG        = _RC['slug']
    REGION_SUBREGIONS  = _RC['subregions']
    REGION_PIPE_LIKE   = _RC['pipe_like']
    METADATA_FILE_NAME = _RC['meta_file']
    _SR_LIST = "'" + "','".join(REGION_SUBREGIONS) + "'"

    if len(SELECTED_REGION_KEYS) > 1:
        print(f"\n{'='*60}")
        print(f"  Running region: {REGION_LABEL}")
        print(f"{'='*60}\n")

    print(f"Pulling live data from Org62 ({REGION_LABEL})...")

# ── Parallel SOQL fetch ───────────────────────────────────────────────────────
import concurrent.futures as _cf

_last_sat = (TODAY - timedelta(days=(TODAY.weekday() + 2) % 7)).isoformat()

_queries = {
    'q1': (f"""
SELECT Id, Name, pse__Stage__c, pse__Account__c, pse__Account__r.Name,
  pse__Project_Manager__r.Name, ProjectManager2Contact__r.Name,
  pse__Practice__r.Name, pse__Billing_Type__c, pse__Start_Date__c, pse__End_Date__c,
  pse__Bookings__c, pse__Billings__c, Total_Amount_Remaining__c,
  Revenue_Recognized_Comments__c, Forecasted_Amount_Remaining_Subreason__c,
  Forecasted_Amount_Remaining_Details__c,
  Overall_Bid_Margin_new__c, Margin_at_Close_Percent__c,
  Percent_Complete__c, Revenue_Treatment__c,
  Health_Risk_Score__c, Data_Quality_Score__c,
  pse__Opportunity__r.Owner.Name,
  pse__Opportunity__r.Salesforce_Exec_Sponsor__c,
  pse__Account__r.Owner.Name,
  pse__Unscheduled_Backlog__c, T_M_Amount_Remaining__c,
  Do_Not_Survey__c, Do_Not_Survey_Reason__c, pse_survey_send_date__c
FROM pse__Proj__c
WHERE pse__Stage__c IN ('In Progress', 'In Progress - SWE', 'On Hold')
  AND Subregion_new__c IN ({_SR_LIST})
  AND pse__Practice__r.Name != 'FDE'
  AND pse__Account__r.Name NOT IN ('Salesforce', 'Salesforce.com')
ORDER BY pse__Account__r.Name, Name
""", 'Q1 Projects'),
    'q2': (f"""
SELECT Project__c, Overall_Health__c, Trend_new__c, High_Watch_Visibility__c,
  Pulse_Update_Frequency_Required__c, LastModifiedDate,
  Scope_Status__c, Schedule_Status__c, Budget_Status__c,
  Resource_Status__c, Customer_Status__c, Governance_Status__c,
  Reason_for_RY_Path_to_Green_Scope__c, Reason_for_RY_Path_to_Green_Schedule__c,
  Reason_for_RY_Path_to_Green_Budget__c, Reason_for_RY_Path_to_Green_Resource__c,
  Reason_for_RY_Path_to_Green_Customer__c,
  Summary_Status__c, Leadership_Notes__c, Action_Needed_from_Leadership__c,
  Next_Steering_Committee_Date__c, Next_Go_Live_Date__c, SWE_or_CO_anticipated__c,
  Overall_Pulse_Status__c
FROM Project_Health_Check__c
WHERE Not_Primary_Pulse_Record__c = false
  AND Project__r.pse__Stage__c IN ('In Progress', 'In Progress - SWE', 'On Hold')
  AND Project__r.Subregion_new__c IN ({_SR_LIST})
  AND Project__r.pse__Practice__r.Name != 'FDE'
  AND Project__r.pse__Account__r.Name NOT IN ('Salesforce', 'Salesforce.com')
""", 'Q2 Pulse'),
    'q3': (f"""
SELECT Project__c
FROM Project_Health_Check__c
WHERE Project__r.pse__Stage__c IN ('In Progress', 'In Progress - SWE', 'On Hold')
  AND Project__r.Subregion_new__c IN ({_SR_LIST})
  AND Project__r.pse__Practice__r.Name != 'FDE'
  AND Project__r.pse__Account__r.Name NOT IN ('Salesforce', 'Salesforce.com')
GROUP BY Project__c
LIMIT 2000
""", 'Q3 Any Pulse'),
    'q4': (f"""
SELECT pse__Project__c, COUNT(Id), SUM(pse__Request_Billable_Amount__c), MIN(pse__Start_Date__c)
FROM pse__Resource_Request__c
WHERE pse__Status__c IN ('Draft', 'Ready to Staff', 'Tentative', 'Hold')
  AND pse__Project__r.pse__Stage__c IN ('In Progress', 'In Progress - SWE', 'On Hold')
  AND pse__Project__r.Subregion_new__c IN ({_SR_LIST})
  AND pse__Project__r.pse__Practice__r.Name != 'FDE'
GROUP BY pse__Project__c
LIMIT 2000
""", 'Q4 Open RRs'),
    'q5': (f"""
SELECT AccountId, SUM(Amount) totalAmt
FROM Opportunity
WHERE pse__Is_Services_Opportunity__c = true
  AND Sub_region__c LIKE '{REGION_PIPE_LIKE}'
  AND StageName NOT IN ('Closed Won', 'Closed Lost', '10 - Closed Won', '10 - Closed Lost')
  AND Amount != null
GROUP BY AccountId
LIMIT 2000
""", 'Q5 Pipeline'),
    'q6': (f"""
SELECT pse__Project__c,
  SUM(pse__Actual_Billable_Amount__c) actualAmt,
  SUM(Estimated_Amount__c) estimatedAmt,
  SUM(pse__Actual_Hours__c) actualHrs,
  SUM(pse__Estimated_Hours__c) estimatedHrs
FROM pse__Est_vs_Actuals__c
WHERE pse__Time_Period_Type__c = 'Week'
  AND pse__End_Date__c = {_last_sat}
  AND pse__Project__r.Subregion_new__c IN ({_SR_LIST})
  AND pse__Project__r.pse__Practice__r.Name != 'FDE'
  AND pse__Project__r.pse__Account__r.Name NOT IN ('Salesforce', 'Salesforce.com')
GROUP BY pse__Project__c
LIMIT 2000
""", 'Q6 EvA'),
    'q7a': (f"""
SELECT pse__Account__c FROM pse__Proj__c
WHERE Subregion_new__c IN ({_SR_LIST})
  AND pse__Practice__r.Name != 'FDE'
GROUP BY pse__Account__c LIMIT 2000
""", 'Q7a Account IDs'),
    'q8': (f"""
SELECT PSE_Project__c, US_Overall_Satisfaction__c, COMPLETIONTIME__c
FROM Clicktools_Survey_Results__c
WHERE PSE_Project__c != null
  AND PSE_Project__r.Subregion_new__c IN ({_SR_LIST})
  AND Survey_Status__c = 'Complete'
ORDER BY COMPLETIONTIME__c DESC
LIMIT 2000
""", 'Q8 CSAT'),
    'q9': (f"""
SELECT pse__Project__c, Resource_Region__c,
  pse__Resource__r.Name, pse__Role__c,
  pse__Planned_Hours__c, pse__Scheduled_Hours__c,
  Total_Billable_and_Credited_Hours__c, Actual_Hours_Remaining__c,
  pse__Start_Date__c, pse__End_Date__c
FROM pse__Assignment__c
WHERE pse__Status__c IN ('Tentative', 'Scheduled')
  AND pse__Project__r.pse__Stage__c IN ('In Progress', 'In Progress - SWE', 'On Hold')
  AND pse__Project__r.Subregion_new__c IN ({_SR_LIST})
  AND pse__Project__r.pse__Practice__r.Name != 'FDE'
  AND pse__Project__r.pse__Account__r.Name NOT IN ('Salesforce', 'Salesforce.com')
LIMIT 5000
""", 'Q9 Assignments'),
}

with _cf.ThreadPoolExecutor(max_workers=9) as _pool:
    _futures = {k: _pool.submit(soql, q, lbl) for k, (q, lbl) in _queries.items()}
    _results = {k: f.result() for k, f in _futures.items()}

q1_rows  = _results['q1']
q2_rows  = _results['q2']
q3_rows  = _results['q3']
q4_rows  = _results['q4']
q5_pipe  = _results['q5']
q6_eva   = _results['q6']
q7_accts  = _results['q7a']
q9_rows   = _results['q9']

# Build CSAT score map: pid → most recent US_Overall_Satisfaction__c score
q8_csat_rows = _results['q8']
q8_csat = {}
for _row in q8_csat_rows:
    _pid = _row.get('PSE_Project__c')
    if _pid and _pid not in q8_csat:  # first = most recent (query is ORDER BY DESC)
        q8_csat[_pid] = _row.get('US_Overall_Satisfaction__c')

# ── Q7b/c: Overdue invoices (depends on q7a account IDs) ─────────────────────
overdue_map = {}
if q7_accts:
    acct_ids = list({r.get('pse__Account__c') for r in q7_accts if r.get('pse__Account__c')})
    chunk = 200
    inv_rows = []
    for i in range(0, len(acct_ids), chunk):
        ids_str = "'" + "','".join(acct_ids[i:i+chunk]) + "'"
        inv_rows += soql(f"""
SELECT Name, sfbill__BalanceDue__c
FROM sfbill__Transaction__c
WHERE sfbill__TransactionType__c = 'INV'
  AND sfbill__BalanceDue__c > 0
  AND sfbill__DueDate__c < TODAY
  AND sfbill__InvoiceAge__c > 30
  AND CurrencyIsoCode = 'USD'
  AND sfbill__AccountName__c IN ({ids_str})
LIMIT 2000
""", f'Q7b Invoices chunk {i//chunk+1}')
    if inv_rows:
        be_rows = soql(f"""
SELECT pse__Invoice_Number__c, pse__Project__c
FROM pse__Billing_Event__c
WHERE pse__Project__r.Subregion_new__c IN ({_SR_LIST})
  AND pse__Invoice_Number__c != null
LIMIT 5000
""", 'Q7c Billing Events')
        inv_to_proj = {r['pse__Invoice_Number__c']: r['pse__Project__c'] for r in be_rows if r.get('pse__Invoice_Number__c')}
        for r in inv_rows:
            pid = inv_to_proj.get(r.get('Name', ''))
            if pid:
                overdue_map.setdefault(pid, {'amount': 0, 'count': 0})
                overdue_map[pid]['amount'] += to_f(r.get('sfbill__BalanceDue__c')) or 0
                overdue_map[pid]['count'] += 1

acct_pipe_map = {}
for r in q5_pipe:
    aid = r.get('AccountId') or r.get('accountid') or ''
    amt = to_f(r.get('totalAmt') or r.get('totalamt')) or 0
    if aid:
        acct_pipe_map[aid] = amt
eva_map = {}
for r in q6_eva:
    pid = r.get('pse__Project__c','')
    if not pid: continue
    act_amt  = to_f(r.get('actualAmt')    or r.get('actualamt'))    or 0
    est_amt  = to_f(r.get('estimatedAmt') or r.get('estimatedamt')) or 0
    act_hrs  = to_f(r.get('actualHrs')    or r.get('actualhrs'))    or 0
    est_hrs  = to_f(r.get('estimatedHrs') or r.get('estimatedhrs')) or 0
    eva_amt  = act_amt - est_amt
    eva_pct  = ((act_hrs - est_hrs) / est_hrs * 100) if est_hrs else None
    eva_map[pid] = {'eva_amt': eva_amt, 'eva_pct': eva_pct,
                    'act_amt': act_amt, 'est_amt': est_amt}

# ── Build lookup maps ─────────────────────────────────────────────────────────
pulse_map = {}
for r in q2_rows:
    pid = r.get('Project__c', '')
    if pid:
        pulse_map[pid] = r

any_pulse_pids = {r.get('Project__c') for r in q3_rows if r.get('Project__c')}

rr_map = {}
for r in q4_rows:
    pid = r.get('pse__Project__c', '')
    if not pid: continue
    cnt = to_f(r.get('expr0')) or 0
    rev = to_f(r.get('expr1')) or 0
    earliest = r.get('expr2') or ''
    if earliest: earliest = str(earliest)[:10]
    rr_map[pid] = {'count': int(cnt), 'revenue': rev, 'earliest_start': earliest}

gdc_map = {}
for r in q9_rows:
    pid = r.get('pse__Project__c', '')
    if not pid: continue
    region   = (r.get('Resource_Region__c') or '').strip()
    res_name = (r.get('pse__Resource__r.Name') or '').strip()
    role     = (r.get('pse__Role__c') or '').strip()
    est_hrs  = to_f(r.get('pse__Planned_Hours__c'))
    sch_hrs  = to_f(r.get('pse__Scheduled_Hours__c'))
    act_hrs  = to_f(r.get('Total_Billable_and_Credited_Hours__c'))
    rem_hrs  = to_f(r.get('Actual_Hours_Remaining__c'))
    start    = (r.get('pse__Start_Date__c') or '')[:10]
    end      = (r.get('pse__End_Date__c') or '')[:10]
    if pid not in gdc_map:
        gdc_map[pid] = {'total': 0, 'india': 0, 'resources': []}
    gdc_map[pid]['total'] += 1
    if region == 'GDC India':
        gdc_map[pid]['india'] += 1
    gdc_map[pid]['resources'].append({
        'name': res_name, 'role': role, 'region': region,
        'est_hrs': est_hrs, 'sch_hrs': sch_hrs, 'act_hrs': act_hrs, 'rem_hrs': rem_hrs,
        'start': start, 'end': end,
    })

# ── Helpers ───────────────────────────────────────────────────────────────────
def is_swe_co(name, stage, swe_field):
    nl = name.lower()
    if any(o in nl for o in SWE_PM_OVERRIDES): return True
    if stage == 'In Progress - SWE': return True
    if swe_field and swe_field not in ('', 'No', None): return True
    if 'swe' in nl: return True
    if 'ari' in nl: return True  # ARI projects always treated as SWE — margin excluded
    return False

def is_multiyr_cto(name, end_dt, rev_treat):
    if not end_dt: return False
    keywords = ['cto', 'sela', 'seh', 'advisory', 'term', 'rta']
    return end_dt > date(2027, 1, 31) and (
        any(k in name.lower() for k in keywords) or
        (rev_treat or '') in ('Term', 'Percent Complete')
    )

def rag(v):
    if not v: return ''
    return str(v).strip().lower().capitalize()

# ── Rule group labels ─────────────────────────────────────────────────────────
RULE_GROUPS = {
    'MARGIN_RED':        'Margin',
    'MARGIN_YELLOW':     'Margin',
    'FAR_RED_NEG':       'FAR',
    'FAR_RED_UNDERUTIL': 'FAR',
    'FAR_YELLOW':        'FAR',
    'RR_RISK':           'Resource',
    'GDC_LOW':           'Resource',
    'OVERDUE_INV':       'Invoice',
    'SWE_BURNING_HOT':   'SWE Burn',
    'NO_PULSE':          'Governance',
    'NO_STEERCO':        'Governance',
    'MISSING_PTG':       'Governance',
    'END_DATE_PAST':     'End Date',
    'END_DATE_UPCOMING': 'End Date',
    'CSAT_OVERDUE':      'CSAT',
    'CSAT_EXEMPT':    'CSAT',
}

def fmt_violation(v):
    code, msg = v[0], v[1]
    grp = RULE_GROUPS.get(code, '')
    prefix = f"[{grp}] " if grp else ''
    return f"{prefix}{code}: {msg}"

# ── Evaluate ──────────────────────────────────────────────────────────────────
print("Evaluating projects...")
results = []

for p in q1_rows:
    pid      = p.get('Id', '')
    name     = p.get('Name', '') or ''
    acct_id  = p.get('pse__Account__c', '') or ''
    acct     = p.get('pse__Account__r.Name', '') or ''
    pm       = p.get('pse__Project_Manager__r.Name', '') or ''
    pm2      = p.get('ProjectManager2Contact__r.Name', '') or ''
    opp_owner   = p.get('pse__Opportunity__r.Owner.Name', '') or ''
    exec_sponsor= p.get('pse__Opportunity__r.Salesforce_Exec_Sponsor__c', '') or ''
    acct_owner  = p.get('pse__Account__r.Owner.Name', '') or ''
    stage    = p.get('pse__Stage__c', '') or ''
    practice = p.get('pse__Practice__r.Name', '') or ''
    rev_treat    = p.get('Revenue_Treatment__c', '') or ''
    billing_type = p.get('pse__Billing_Type__c', '') or ''
    bookings = to_f(p.get('pse__Bookings__c'))
    billings = to_f(p.get('pse__Billings__c'))
    far      = to_f(p.get('Total_Amount_Remaining__c'))
    unsch_backlog  = to_f(p.get('pse__Unscheduled_Backlog__c'))
    actuals_rem    = to_f(p.get('T_M_Amount_Remaining__c'))
    far_reason   = p.get('Revenue_Recognized_Comments__c') or ''
    far_details  = p.get('Forecasted_Amount_Remaining_Details__c') or ''
    far_subreason= p.get('Forecasted_Amount_Remaining_Subreason__c') or ''
    bid_margin   = to_f(p.get('Overall_Bid_Margin_new__c'))
    close_margin = to_f(p.get('Margin_at_Close_Percent__c'))
    work_pct      = to_f(p.get('Percent_Complete__c'))
    health_risk_score  = to_f(p.get('Health_Risk_Score__c'))
    data_quality_score = to_f(p.get('Data_Quality_Score__c'))
    start_dt     = to_d(p.get('pse__Start_Date__c'))
    end_dt       = to_d(p.get('pse__End_Date__c'))
    do_not_survey        = bool(p.get('Do_Not_Survey__c'))
    do_not_survey_reason = p.get('Do_Not_Survey_Reason__c') or ''
    survey_send_date     = p.get('pse_survey_send_date__c') or ''
    csat_score           = q8_csat.get(pid)

    pulse = pulse_map.get(pid, {})
    health   = pulse.get('Overall_Health__c') or 'Null'
    swe_field= pulse.get('SWE_or_CO_anticipated__c') or ''
    scope_s  = rag(pulse.get('Scope_Status__c'))
    sched_s  = rag(pulse.get('Schedule_Status__c'))
    budget_s = rag(pulse.get('Budget_Status__c'))
    resource_s = rag(pulse.get('Resource_Status__c'))
    customer_s = rag(pulse.get('Customer_Status__c'))
    ptg_scope    = pulse.get('Reason_for_RY_Path_to_Green_Scope__c') or ''
    ptg_sched    = pulse.get('Reason_for_RY_Path_to_Green_Schedule__c') or ''
    ptg_budget   = pulse.get('Reason_for_RY_Path_to_Green_Budget__c') or ''
    ptg_resource = pulse.get('Reason_for_RY_Path_to_Green_Resource__c') or ''
    ptg_customer = pulse.get('Reason_for_RY_Path_to_Green_Customer__c') or ''
    action_needed     = pulse.get('Action_Needed_from_Leadership__c') or ''
    leadership_notes  = pulse.get('Leadership_Notes__c') or ''
    overall_summary   = pulse.get('Summary_Status__c') or ''
    trend             = pulse.get('Trend_new__c') or ''
    high_watch        = (pulse.get('High_Watch_Visibility__c') or '').strip().lower() == 'include'
    last_updated   = pulse.get('LastModifiedDate') or ''
    steerco_date   = to_d(pulse.get('Next_Steering_Committee_Date__c'))
    next_golive    = to_d(pulse.get('Next_Go_Live_Date__c'))

    rr = rr_map.get(pid, {})
    rr_count   = rr.get('count', 0)
    rr_revenue = rr.get('revenue', 0)
    rr_earliest= rr.get('earliest_start', '')

    gdc_data  = gdc_map.get(pid, {})
    gdc_total = gdc_data.get('total', 0)
    gdc_india = gdc_data.get('india', 0)
    gdc_other = gdc_total - gdc_india
    gdc_pct       = (gdc_india / gdc_total) if gdc_total > 0 else None
    gdc_resources = sorted(gdc_data.get('resources', []), key=lambda x: (x['region'] != 'GDC India', x['name']))

    overdue_inv = overdue_map.get(pid, {}).get('amount', 0)
    overdue_cnt = overdue_map.get(pid, {}).get('count', 0)

    open_pipe = acct_pipe_map.get(acct_id, 0) or 0
    _eva      = eva_map.get(pid, {})
    eva_amt   = _eva.get('eva_amt')
    eva_pct   = _eva.get('eva_pct')
    swe_co    = is_swe_co(name, stage, swe_field)
    cto_exempt= is_multiyr_cto(name, end_dt, rev_treat)
    has_pulse = bool(pulse) or pid in any_pulse_pids

    baseline_ry = any(s in ('Red', 'Yellow') for s in [scope_s, sched_s, budget_s, resource_s, customer_s])

    violations = []

    # 1A: Margin degradation (SWE/CO excluded)
    if has_pulse and not swe_co and bid_margin is not None and close_margin is not None:
        delta = close_margin - bid_margin
        if delta < -5:
            violations.append(('MARGIN_RED', f'Margin delta {delta:+.1f}% (bid {bid_margin:.1f}% → close {close_margin:.1f}%)'))
        elif delta < 0:
            violations.append(('MARGIN_YELLOW', f'Margin delta {delta:+.1f}% (bid {bid_margin:.1f}% → close {close_margin:.1f}%)'))

    # 1C: FAR
    def far_context():
        lns = []
        if far_reason:    lns.append(f'  FAR Reason: "{far_reason}"')
        if far_details:   lns.append(f'  FAR Details: "{far_details}"')
        if far_subreason: lns.append(f'  FAR Subreason: "{far_subreason}"')
        return '\n'.join(lns) if lns else '  (no FAR context fields populated)'

    if has_pulse and far is not None:
        bk_val = bookings or 0
        if far < -1.0:
            violations.append(('FAR_RED_NEG', f'FAR overrun: ${far:,.0f}\n{far_context()}'))
        elif far > 0 and bk_val > 0:
            pct = (far / bk_val) * 100
            if pct > 15 and not cto_exempt:
                violations.append(('FAR_RED_UNDERUTIL', f'FAR ${far:,.0f} = {pct:.1f}% of bookings\n{far_context()}'))
            elif pct > 5:
                violations.append(('FAR_YELLOW', f'FAR ${far:,.0f} = {pct:.1f}% of bookings\n{far_context()}'))

    # 1C-RR: Pending RR revenue risk
    if rr_revenue and rr_revenue > 0:
        urgency = ''
        if far and far > 0 and rr_revenue / far > 0.10:
            urgency = ' ⚠️ ELEVATED — exceeds 10% of remaining FAR'
        violations.append(('RR_RISK', f'Pending RR Revenue: ${rr_revenue:,.0f} ({rr_count} open RRs){urgency}'))

    # 1D: GDC India resourcing threshold
    if gdc_total > 0 and gdc_pct is not None and gdc_pct <= 0.65:
        violations.append(('GDC_LOW',
            f'GDC India share {gdc_pct*100:.0f}% ({gdc_india}/{gdc_total} assigned) — below 65% threshold'))

    # 1E: Overdue invoices
    if overdue_inv > 0:
        violations.append(('OVERDUE_INV', f'Overdue invoices: ${overdue_inv:,.0f} ({overdue_cnt} invoice(s))'))

    # 1F: SWE burn rate
    swe_burn_str = None
    if swe_co and start_dt and end_dt and work_pct is not None:
        total_days = (end_dt - start_dt).days
        if total_days > 0:
            elapsed = (TODAY - start_dt).days
            time_pct = (elapsed / total_days) * 100
            if time_pct > 0:
                br = work_pct / time_pct
                tag = '🟢 NORMAL'
                if br > 1.20:   tag = '🔴 HOT'
                elif br > 1.10: tag = '🟡 WATCH'
                elif br < 0.80: tag = '🔵 SLOW'
                swe_burn_str = f'{work_pct:.0f}% complete | {time_pct:.0f}% elapsed | Ratio {br:.2f} {tag}'
                if br > 1.20:
                    violations.append(('SWE_BURNING_HOT', f'SWE burn ratio {br:.2f} — HOT'))

    # 2A: No pulse (≥$150K)
    if not has_pulse and (bookings or 0) >= 150000:
        violations.append(('NO_PULSE', f'No pulse — ${(bookings or 0):,.0f} project'))

    # 2F: SteerCo Date required for projects ≥$750K
    # Exempt projects must set date to 01/01/2100; missing or null = violation
    STEERCO_EXEMPT_DATE = date(2100, 1, 1)
    if (bookings or 0) >= 750000:
        if steerco_date is None:
            violations.append(('NO_STEERCO', f'Next Steering Committee Date is required for this ${(bookings or 0):,.0f} project. ACTION: If exempt (SEH, Advisory, etc.) set date to 01/01/2100.'))
        elif steerco_date == STEERCO_EXEMPT_DATE:
            pass  # Compliant exempt project — no violation
        # else: valid future date set — compliant

    # 3A: Watermelon
    fin_reds = {'MARGIN_RED', 'FAR_RED_NEG', 'FAR_RED_UNDERUTIL', 'OVERDUE_INV', 'SWE_BURNING_HOT'}
    has_fin_red = bool(fin_reds & {v[0] for v in violations})
    is_green = health.lower() == 'green'
    is_watermelon = is_green and (baseline_ry or has_fin_red)

    # 3B: Missing PTG
    if baseline_ry and has_pulse:
        for dim, (stat, ptg) in [
            ('Scope',    (scope_s,    ptg_scope)),
            ('Schedule', (sched_s,    ptg_sched)),
            ('Budget',   (budget_s,   ptg_budget)),
            ('Resource', (resource_s, ptg_resource)),
            ('Customer', (customer_s, ptg_customer)),
        ]:
            if stat in ('Red', 'Yellow') and not ptg:
                violations.append(('MISSING_PTG', f'Missing PTG for {dim} ({stat})'))

    # End date rules
    if end_dt:
        days_to_end = (end_dt - TODAY).days
        if days_to_end < 0:
            violations.append(('END_DATE_PAST', f'End date past due ({end_dt.isoformat()})'))
        elif days_to_end <= 45:
            violations.append(('END_DATE_UPCOMING', f'End date in {days_to_end} day(s) ({end_dt.isoformat()})'))

    # 4B: CSAT rules
    if do_not_survey:
        reason_txt = do_not_survey_reason or 'No reason provided'
        violations.append(('CSAT_EXEMPT', f'CSAT exempt: {reason_txt}'))
    elif (
        (bookings or 0) > 150_000
        and stage == 'In Progress'
        and start_dt
        and (TODAY - start_dt).days >= 90
        and not survey_send_date
    ):
        violations.append(('CSAT_OVERDUE',
            f'No CSAT survey sent; project active {(TODAY - start_dt).days} days, ${(bookings or 0):,.0f} bookings'))

    # Resource concern flag (resource baseline R/Y or open RRs)
    has_resource_concern = (resource_s in ('Red', 'Yellow')) or (rr_count > 0)

    results.append({
        'pid': pid, 'name': name, 'acct': acct, 'pm': pm, 'pm2': pm2, 'opp_owner': opp_owner, 'exec_sponsor': exec_sponsor, 'acct_owner': acct_owner,
        'stage': stage, 'health': health, 'trend': trend, 'high_watch': high_watch,
        'bookings': bookings, 'billings': billings, 'far': far,
        'far_reason': far_reason, 'far_details': far_details, 'far_subreason': far_subreason,
        'bid_margin': bid_margin, 'close_margin': close_margin,
        'scope_s': scope_s, 'sched_s': sched_s, 'budget_s': budget_s,
        'resource_s': resource_s, 'customer_s': customer_s,
        'swe_co': swe_co, 'cto_exempt': cto_exempt,
        'has_pulse': has_pulse, 'is_watermelon': is_watermelon,
        'violations': violations, 'baseline_ry': baseline_ry,
        'swe_burn_str': swe_burn_str, 'overall_summary': overall_summary,
        'leadership_notes': leadership_notes,
        'action_needed': action_needed, 'last_updated': last_updated,
        'steerco_date': steerco_date, 'next_golive': next_golive,
        'overdue_inv': overdue_inv, 'overdue_cnt': overdue_cnt,
        'rr_count': rr_count, 'rr_revenue': rr_revenue, 'rr_earliest': rr_earliest,
        'gdc_total': gdc_total, 'gdc_india': gdc_india, 'gdc_other': gdc_other, 'gdc_pct': gdc_pct,
        'gdc_resources': gdc_resources,
        'ptg_scope': ptg_scope, 'ptg_sched': ptg_sched, 'ptg_budget': ptg_budget,
        'ptg_resource': ptg_resource, 'ptg_customer': ptg_customer,
        'work_pct': work_pct, 'start_dt': start_dt, 'end_dt': end_dt,
        'rev_treat': rev_treat, 'billing_type': billing_type, 'practice': practice,
        'open_pipe': open_pipe,
        'unsch_backlog': unsch_backlog, 'actuals_rem': actuals_rem,
        'eva_amt': eva_amt, 'eva_pct': eva_pct,
        'health_risk_score': health_risk_score,
        'data_quality_score': data_quality_score,
        'has_resource_concern': has_resource_concern,
        'csat_score': csat_score,
        'do_not_survey': do_not_survey,
        'survey_send_date': survey_send_date,
    })

    # ── Load tier + portfolio owner from metadata file ──────────────────────────
    METADATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), METADATA_FILE_NAME)

meta_by_pid  = {}   # project record id → {tier, owner}
meta_by_name = {}   # project name      → {tier, owner}

if os.path.exists(METADATA_FILE):
    import re as _re
    _current_tier = None
    for _line in open(METADATA_FILE, encoding='utf-8'):
        _m = _re.match(r'^## (Tier \d+|Unassigned)', _line)
        if _m:
            _current_tier = _m.group(1)
            continue
        _m = _re.match(r'^\|\s*([A-Za-z0-9]{15,18})\s*\|\s*(.+?)\s*\|\s*(.+?)\s*\|\s*(.+?)\s*\|', _line)
        if _m and _current_tier:
            _pid, _name, _acct, _owner = _m.group(1), _m.group(2), _m.group(3), _m.group(4)
            _entry = {'tier': int(_current_tier.split()[-1]) if _current_tier.startswith('Tier') else 0, 'owner': _owner.strip()}
            meta_by_pid[_pid.strip()]  = _entry
            meta_by_name[_name.strip()] = _entry

def _tier_lookup(r):
    entry = meta_by_pid.get(r['pid']) or meta_by_name.get(r['name'])
    return entry['tier'] if entry else None


_unmatched = [r for r in results if not (meta_by_pid.get(r['pid']) or meta_by_name.get(r['name']))]
if _unmatched and sys.stdin.isatty():
    _meta_basename = os.path.basename(METADATA_FILE)
    print(f"\n⚠️  {len(_unmatched)} project(s) not in {_meta_basename}.")
    print(f"    Tip: add rows directly to {_meta_basename} to skip this prompt.")
    print(f"    For Portfolio Owner — press Enter to use PM2 as default.\n")
    for r in _unmatched:
        _pm2_default = r.get('pm2') or r.get('pm') or 'Unassigned'
        print(f"  [{r['pid'][:18]}] {r['name'][:60]}  (PM2: {_pm2_default})")
        _t = input("    Assign Tier (1/2/3) [3]: ").strip() or '3'
        _o = input(f"    Portfolio Owner [{_pm2_default}]: ").strip() or _pm2_default
        meta_by_pid[r['pid']] = {'tier': int(_t) if _t in ('1','2','3') else 3, 'owner': _o}
        with open(METADATA_FILE, 'a', encoding='utf-8') as _mf:
            _mf.write(f"| {r['pid']} | {r['name']} | {r['acct']} | {_o} |\n")
        print(f"    → Saved to {_meta_basename} (Tier {_t}, Owner: {_o})\n")
elif _unmatched:
    for r in _unmatched:
        _pm2_default = r.get('pm2') or r.get('pm') or 'Unassigned'
        meta_by_pid[r['pid']] = {'tier': 3, 'owner': _pm2_default}

for r in results:
    entry = meta_by_pid.get(r['pid']) or meta_by_name.get(r['name']) or {'tier': 3, 'owner': 'Unassigned'}
    r['tier']  = entry['tier']
    r['owner'] = entry['owner']

# ── Sync assignments with Page Host SQLite DB ─────────────────────────────────
# DB is the source of truth after first seed. On first run (empty table) the DB
# is seeded from metadata.md. On subsequent runs the DB overrides metadata.md.
_DB_CLIENT_ID     = os.environ.get('DB_CLIENT_ID', '')
_DB_CLIENT_SECRET = os.environ.get('DB_CLIENT_SECRET', '')
_PAGE_HOST_URL    = os.environ.get('PAGE_HOST_URL', 'https://single-html-page-app-host-07cda8a7041b.herokuapp.com')
_TILE_ID_ENV      = os.environ.get('TILE_ID', '817')

if _DB_CLIENT_ID and _DB_CLIENT_SECRET:
    import base64 as _b64, urllib.request as _ureq, urllib.error as _uerr
    _db_basic = _b64.b64encode(f"{_DB_CLIENT_ID}:{_DB_CLIENT_SECRET}".encode()).decode()
    _db_write_url = f"{_PAGE_HOST_URL}/api/tiles/{_TILE_ID_ENV}/db/write"
    _db_query_url = f"{_PAGE_HOST_URL}/api/tiles/{_TILE_ID_ENV}/db/query"

    def _db_request(url, method, body_dict, auth):
        body = json.dumps(body_dict).encode()
        req = _ureq.Request(url, data=body, method=method,
                            headers={'Content-Type': 'application/json', 'Authorization': auth})
        try:
            with _ureq.urlopen(req, timeout=15) as resp:
                return json.loads(resp.read())
        except _uerr.HTTPError as e:
            print(f"⚠️  DB request failed {e.code}: {e.read().decode()[:200]}")
            return None

    # Query current DB assignments (uses session-auth REST endpoint with Basic auth as fallback)
    _bearer_token = os.environ.get('PAGE_HOST_TOKEN', '')
    _db_rows = None
    if _bearer_token:
        _qresp = _db_request(_db_query_url, 'POST',
                             {'sql': 'SELECT pid, tier, po FROM assignments'},
                             f'Bearer {_bearer_token}')
        if _qresp and 'rows' in _qresp:
            _db_rows = _qresp['rows']

    if _db_rows:
        # DB has data — apply DB assignments, overriding metadata.md
        _db_map = {row['pid']: row for row in _db_rows}
        for r in results:
            if r['pid'] in _db_map:
                db_entry = _db_map[r['pid']]
                r['tier']  = int(db_entry.get('tier') or r['tier'])
                r['owner'] = db_entry.get('po') or r['owner']
        print(f"✅  DB assignments loaded: {len(_db_rows)} rows (overriding metadata.md)")
    else:
        # DB empty or unreachable — seed from current results
        _seed_rows = [{'pid': r['pid'], 'tier': r['tier'], 'po': r['owner']} for r in results]
        _seed_resp = _db_request(_db_write_url, 'POST', {
            'table': 'assignments',
            'mode': 'upsert',
            'key_column': 'pid',
            'columns': ['pid', 'tier', 'po'],
            'rows': _seed_rows,
        }, f'Basic {_db_basic}')
        if _seed_resp and _seed_resp.get('ok'):
            print(f"✅  DB seeded with {len(_seed_rows)} assignments from metadata.md")
        else:
            print("⚠️  DB seed failed — using metadata.md assignments")

# ── Slack Intelligence — load from cache written by Claude Code ───────────────
# To refresh: ask Claude Code to run the Slack enrichment before generating reports.
# Cache file: slack_intel_YYYY-MM-DD.json (auto-dated, one per day)
_slack_cache_data = {}
if os.path.exists(_SLACK_CACHE):
    try:
        with open(_SLACK_CACHE) as _f:
            _slack_cache_data = json.load(_f)
        print(f"💬  Slack intel loaded: {sum(1 for v in _slack_cache_data.values() if v)} projects with intel.")
    except Exception:
        pass

for r in results:
    r['slack_intel'] = _slack_cache_data.get(r['pid'], '')

# ── Segment ───────────────────────────────────────────────────────────────────
watermelons = [r for r in results if r['is_watermelon']]
reds        = [r for r in results if r['health'].lower() == 'red' and not r['is_watermelon']]
yellows     = [r for r in results if r['health'].lower() == 'yellow']
clean_green = [r for r in results if r['health'].lower() == 'green' and not r['is_watermelon']]
no_pulse    = [r for r in results if not r['has_pulse']]
no_pulse_flagged = [r for r in results if not r['has_pulse'] and (r['bookings'] or 0) >= 150000]
on_hold     = [r for r in results if r['stage'] == 'On Hold']

_all_results.extend(results)

total_bk  = sum(r['bookings'] or 0 for r in results)
_bk_sorted = sorted(r['bookings'] or 0 for r in results)
_n = len(_bk_sorted)
median_bk = (_bk_sorted[_n//2] if _n % 2 else (_bk_sorted[_n//2-1] + _bk_sorted[_n//2]) / 2) if _n else 0
total_bil = sum(r['billings'] or 0 for r in results)
total_backlog = total_bk - total_bil
total_far = sum(r['far'] or 0 for r in results if (r['far'] or 0) != 0)
total_far_overrun = sum(r['far'] for r in results if (r['far'] or 0) < 0)
total_rr_rev = sum(r['rr_revenue'] for r in results if r['rr_revenue'])
total_overdue = sum(r['overdue_inv'] for r in results)

margin_pool = [r for r in results if not r['swe_co'] and r['bookings'] and r['bid_margin'] is not None and r['close_margin'] is not None]
tbk = sum(r['bookings'] for r in margin_pool)
w_bid   = sum(r['bid_margin']   * r['bookings'] for r in margin_pool) / tbk if tbk else 0
w_close = sum(r['close_margin'] * r['bookings'] for r in margin_pool) / tbk if tbk else 0

hr_scores  = [r['health_risk_score']  for r in results if r.get('health_risk_score')  is not None]
dq_scores  = [r['data_quality_score'] for r in results if r.get('data_quality_score') is not None]
avg_hr = sum(hr_scores) / len(hr_scores) if hr_scores else None
avg_dq = sum(dq_scores) / len(dq_scores) if dq_scores else None

csat_scores_list = [r['csat_score'] for r in results if r.get('csat_score') is not None]
avg_csat = sum(csat_scores_list) / len(csat_scores_list) if csat_scores_list else None

# ── Format helpers ────────────────────────────────────────────────────────────
def bk(r):  return r['bookings'] or 0
def bks(r): return f"${bk(r)/1e3:,.0f}K"

def snip(text, n=300):
    t = ' '.join((text or '').split())
    return t[:n] + ('…' if len(t) > n else '')

def baselines_list(r):
    dims = [('Scope', r['scope_s']), ('Schedule', r['sched_s']), ('Budget', r['budget_s']),
            ('Resource', r['resource_s']), ('Customer', r['customer_s'])]
    return [(lbl, v) for lbl, v in dims if v] or [('All', 'Green')]

def baselines(r):
    items = baselines_list(r)
    return ' | '.join(f'{lbl}={v}' for lbl, v in items)

# ── Build report ──────────────────────────────────────────────────────────────
lines = []
L = lines.append

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — EXECUTIVE SCORECARD
# ══════════════════════════════════════════════════════════════════════════════
L("=" * 80)
L(f"{REGION_LABEL} PORTFOLIO AUDIT")
L(f"Audit Date: {REPORT_DATE}  |  Data: Live Org62  |  Auditor: SPSM/DAF")
L("=" * 80)
L("")
L("RULE KEY")
L("-" * 60)
L("  ── MARGIN RULES ──────────────────────────────────────────")
L("  MARGIN_RED          Margin delta < -5%  (bid vs. close margin)")
L("  MARGIN_YELLOW       Margin delta between 0% and -5%")
L("")
L("  ── FAR / FINANCIAL UTILIZATION RULES ─────────────────────")
L("  FAR_RED_NEG      FAR is negative — active contract overrun")
L("  FAR_RED_UNDERUTIL FAR > 15% of bookings — severe under-utilization")
L("  FAR_YELLOW       FAR between 5%-14% of bookings — under-utilization watch")
L("")
L("  ── RESOURCE RULES ─────────────────────────────────────────")
L("  RR_RISK      Pending RR revenue > $0 — staffing gap risk")
L("                  ELEVATED if pending RR > 10% of remaining FAR")
L("  SWE_BURNING_HOT          SWE burn ratio > 1.20 — scope exhaustion risk")
L("")
L("  ── INVOICE / BILLING RULES ────────────────────────────────")
L("  OVERDUE_INV  Overdue invoice balance outstanding > 30 days")
L("")
L("  ── PULSE / GOVERNANCE RULES ───────────────────────────────")
L("  NO_PULSE     No pulse record on project with bookings ≥ $150K")
L("  MISSING_PTG  Baseline R/Y with no Path to Green explanation")
L("")
L("  ── DATA HYGIENE RULES ─────────────────────────────────────")
L("  NO_STEERCO   Next Steering Committee Date is required for projects ≥ $750K")
L("                  ACTION: If exempt (SEH, Advisory, etc.) set date to 01/01/2100.")
L("")
L("  ── END DATE RULES ─────────────────────────────────────────")
L("  END_DATE_PAST     End date has passed — project may need extension or closure")
L("  END_DATE_UPCOMING End date within 45 days — renewal/extension decision needed")
L("-" * 60)
L("")
L("SECTION 1 — EXECUTIVE SCORECARD")
L("-" * 60)
L(f"Total Active Projects : {len(results)}")
L(f"Portfolio Bookings    : ${total_bk/1e6:.1f}M")
L(f"Portfolio Billings    : ${total_bil/1e6:.1f}M")
L(f"Outstanding Backlog   : ${total_backlog/1e6:.1f}M")
L("")
self_reported_green = len(clean_green) + len(watermelons)
wm_pct = len(watermelons) * 100 // max(self_reported_green, 1)
red_accts = ', '.join(sorted({r['acct'][:12].upper() for r in reds}))
L(f"  🟢 Green                 : {self_reported_green}")
L(f"     ├─ Genuine            : {len(clean_green)}")
L(f"     └─ 🍉 Watermelon      : {len(watermelons)}  ({wm_pct}% misrepresented)")
L(f"  🟡 Yellow                : {len(yellows)}")
L(f"  🔴 Red                   : {len(reds)}  {red_accts}")
L(f"  ⚫ No Pulse (Null)       : {len(no_pulse)}")
L(f"  ⏸  On Hold               : {len(on_hold)}")
L("")
L(f"Weighted Bid Margin   : {w_bid:.1f}%")
L(f"Delivered Margin      : {w_close-w_bid:+.1f}%")
L(f"Margin at Close       : {w_close:.1f}%  ({len(margin_pool)} projects / ${tbk/1e6:.1f}M bookings)")
L(f"SWE/CO Projects       : {sum(1 for r in results if r['swe_co'])} projects (excluded from margin calc)")
L("")
L("  ── FINANCIAL EXPOSURE ─────────────────────────────────────")
L(f"  Total FAR (net)       : ${total_far:,.0f}")
L(f"  Total FAR Overruns    : ${total_far_overrun:,.0f}  (negative FAR = active contract overrun)")
if total_overdue > 0:
    L(f"  Total Overdue Invoices: ${total_overdue:,.0f}")
if total_rr_rev > 0:
    L(f"  Total Revenue @ Risk  : ${total_rr_rev:,.0f}  ({sum(r['rr_count'] for r in results if r['rr_revenue'])} open RRs — staffing gaps)")
L("")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — PORTFOLIO LEADER ACTION PLAN
# ══════════════════════════════════════════════════════════════════════════════
L("=" * 80)
L("SECTION 2 — PORTFOLIO LEADER ACTION PLAN")
L("=" * 80)
L("  Top 15 priority escalations — all watermelons + reds, ranked by bookings")
L("")
L("  Top 15 Priority Escalations:")
top = sorted([r for r in results if r['is_watermelon'] or r['health'].lower() == 'red'],
             key=lambda x: -bk(x))[:15]
for r in top:
    icon  = '🍉' if r['is_watermelon'] else '🔴'
    codes = ', '.join(v[0] for v in r['violations'][:3])
    L(f"  {icon} {r['name'][:55]}  [{r['pm']}]  {bks(r)}  {codes}")

# RR Revenue at Risk table
rr_at_risk = sorted([r for r in results if r['rr_revenue'] > 0], key=lambda x: -x['rr_revenue'])
if rr_at_risk:
    L("")
    L("  Revenue at Risk — Unassigned Resource Requests")
    L("  " + "-" * 100)
    L(f"  {'Account':<18} {'Project':<42} {'PM':<22} {'RRs':>4} {'Pending Rev':>12} {'Earliest Start':<18} {'Flag'}")
    L("  " + "-" * 100)
    for r in rr_at_risk:
        today_str = TODAY.isoformat()
        elevated = '⚠️ Elevated' if r['rr_revenue'] > (r['far'] or 0) * 0.10 and (r['far'] or 0) > 0 else ''
        start = r['rr_earliest'] or 'N/A'
        overdue = start < today_str if (start and start != 'N/A') else False
        start_label = f"{start} (overdue)" if overdue else start
        acct_short = (r['acct'] or '')[:18]
        L(f"  {acct_short:<18} {r['name'][:42]:<42} {r['pm'][:22]:<22} {r['rr_count']:>4} ${r['rr_revenue']:>11,.0f} {start_label:<18} {elevated}")
    L("  " + "-" * 100)
    L(f"  {'TOTAL':<18} {'':42} {'':22} {sum(r['rr_count'] for r in rr_at_risk):>4} ${sum(r['rr_revenue'] for r in rr_at_risk):>11,.0f}")

# Negative FAR table (overruns > $1 to exclude rounding)
neg_far_list = sorted([r for r in results if (r['far'] or 0) < -1], key=lambda x: x['far'] or 0)
if neg_far_list:
    L("")
    L("  Contract Overruns — Projects with Negative FAR")
    L("  " + "-" * 110)
    L(f"  {'Account':<20} {'Project':<42} {'PM':<22} {'Tier':>4} {'FAR':>12}  {'FAR Reason'}")
    L("  " + "-" * 110)
    for r in neg_far_list:
        tier_lbl = f"T{r['tier']}"
        reason = (r['far_reason'] or r['far_details'] or '—')[:40]
        L(f"  {r['acct'][:20]:<20} {r['name'][:42]:<42} {r['pm'][:22]:<22} {tier_lbl:>4} ${r['far']:>11,.0f}  {reason}")
    L("  " + "-" * 110)
    L(f"  Total overrun: ${sum(r['far'] for r in neg_far_list):,.0f}  ({len(neg_far_list)} projects)")

# Resource concerns
resource_concern_list = sorted(
    [r for r in results if r['has_resource_concern']],
    key=lambda x: (-x['rr_revenue'], -(x['bookings'] or 0))
)
if resource_concern_list:
    L("")
    L("  Resource Concerns — Open RRs and/or Red/Yellow Resource Baseline")
    L("  " + "-" * 80)
    L(f"  {'Project':<45} {'PM':<20} {'Res Status':<12} {'Open RRs':>9} {'Pending Rev':>12}")
    L("  " + "-" * 80)
    for r in resource_concern_list:
        res_status = r['resource_s'] or 'Green'
        flag = ' ⚠️' if res_status in ('Red', 'Yellow') else ''
        L(f"  {r['name'][:45]:<45} {r['pm'][:20]:<20} {res_status+flag:<12} {r['rr_count']:>9}  ${r['rr_revenue']:>11,.0f}")
    L("  " + "-" * 80)
    L(f"  Projects with resource concerns: {len(resource_concern_list)}  |  "
      f"Red/Yellow resource baseline: {sum(1 for r in resource_concern_list if r['resource_s'] in ('Red','Yellow'))}  |  "
      f"Total open RRs: {sum(r['rr_count'] for r in resource_concern_list)}")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — PROJECT DETAIL BY TIER
# ══════════════════════════════════════════════════════════════════════════════
L("")
L("=" * 80)
L("SECTION 3 — PROJECT DETAIL BY TIER")
L("=" * 80)

def project_snapshot(r):
    parts = []
    ov = snip(r.get('overall_summary') or '', 400)
    if ov: parts.append(ov)
    ln = snip(r.get('leadership_notes') or '', 300)
    if ln and ln != ov: parts.append(f"Leadership: {ln}")
    if r['action_needed']:
        parts.append(f"Action: {snip(r['action_needed'], 200)}")
    # Path to green for R/Y
    ptg_parts = [r.get(f'ptg_{d}','') or '' for d in ['scope','sched','budget','resource','customer']]
    ptg = ' '.join(p for p in ptg_parts if p)
    if ptg: parts.append(f"Path to Green: {snip(ptg, 300)}")
    if not parts: parts.append('No status summary on file.')
    return '  |  '.join(parts)

def health_icon(r):
    if r.get('high_watch') and r['health'].lower() not in ('red','yellow'):
        return '⚑'
    return {'red':'🔴','yellow':'🟡','green':'🟢'}.get(r['health'].lower(), '⚫')

def pulse_indicator(r):
    """Single-char pulse presence marker for TXT/DOCX/PPTX."""
    return '📋' if r.get('has_pulse') else '○'

for tier_num in [1, 2, 3]:
    tier_label = {1:'TIER 1 — Strategic Accounts (≥$7M)',
                  2:'TIER 2 — Growth Accounts ($750K–$7M)',
                  3:'TIER 3 — Volume Accounts (<$750K)'}[tier_num]
    tier_projs = [r for r in results if r['tier'] == tier_num]
    t_hw  = [r for r in tier_projs if r.get('high_watch')]
    t_red = [r for r in tier_projs if r['health'].lower()=='red' and not r['is_watermelon']]
    t_yel = [r for r in tier_projs if r['health'].lower()=='yellow']
    t_wm  = [r for r in tier_projs if r['is_watermelon']]
    t_grn = [r for r in tier_projs if r['health'].lower()=='green' and not r['is_watermelon']]
    t_np  = [r for r in tier_projs if not r['has_pulse']]
    t_bk  = sum(bk(r) for r in tier_projs)
    t_bil = sum(r['billings'] or 0 for r in tier_projs)

    L("")
    L("=" * 80)
    L(f"  {tier_label}")
    L("=" * 80)
    L(f"  Projects: {len(tier_projs)}  |  Bookings: ${t_bk/1e6:.1f}M  |  Billings: ${t_bil/1e6:.1f}M")
    L(f"  ⚑  High Watch: {len(t_hw)}  |  🔴 Red: {len(t_red)}  |  🟡 Yellow: {len(t_yel)}  "
      f"|  🍉 Watermelon: {len(t_wm)}  |  🟢 Green: {len(t_grn)}  |  ⚫ No Status: {len(t_np)}")
    L("")

    # Order: High Watch → Red → Yellow → Watermelon → Green → No Status
    groups = [
        ('⚑  HIGH WATCH', t_hw),
        ('🔴 RED',         t_red),
        ('🟡 YELLOW',      t_yel),
        ('🍉 WATERMELON',  t_wm),
        ('🟢 GREEN',       t_grn),
        ('⚫ NO STATUS',   t_np),
    ]
    for group_label, group_projs in groups:
        if not group_projs: continue
        L(f"  ── {group_label} ({len(group_projs)}) ──────────────────────────────────────")
        L(f"  {'H':<4} {'P':<3} {'Project':<45} {'Account':<22} {'PM':<22} {'PM2':<18} {'Bookings':>10}  {'Rules':<35}  {'Baselines'}")
        L("  " + "-" * 200)
        for r in sorted(group_projs, key=lambda x: -bk(x)):
            icon  = health_icon(r)
            pulse_ind = pulse_indicator(r)
            pm2   = r.get('pm2','') or ''
            codes = ', '.join(fmt_violation(v) for v in r['violations'][:4])
            bl    = baselines(r)
            def _score_icon(v): return '🟢' if v>=70 else ('🟡' if v>=30 else '🔴') if v is not None else ''
            hr_dq = ''
            if r.get('health_risk_score') is not None: hr_dq += f"  {_score_icon(r['health_risk_score'])}H&R:{r['health_risk_score']:.0f}"
            if r.get('data_quality_score') is not None: hr_dq += f"  {_score_icon(r['data_quality_score'])}DQ:{r['data_quality_score']:.0f}"
            proj_url = f"https://org62.lightning.force.com/lightning/r/pse__Proj__c/{r['pid']}/view"
            _dates = f"{r['start_dt']} → {r['end_dt']}" if r.get('start_dt') or r.get('end_dt') else ''
            _stage_prac = '  '.join(x for x in [r.get('stage',''), r.get('practice','')] if x)
            L(f"  {icon:<4} {pulse_ind:<3} {r['name'][:45]:<45} {r['acct'][:22]:<22} {r['pm'][:22]:<22} {pm2[:18]:<18} {bks(r):>10}  {codes:<35}  {bl}{hr_dq}")
            L(f"         Link: {proj_url}  {_dates}  {_stage_prac}")
            # Financials sub-line
            def _m(v): return f'${v/1e6:.2f}M' if v and abs(v)>=1e6 else (f'${v:,.0f}' if v else '—')
            fin_parts = [f"Type:{r['rev_treat']}" if r.get('rev_treat') else None,
                         f"FAR:{_m(r['far'])}", f"BidM:{r['bid_margin']:.1f}%" if r['bid_margin'] is not None else None,
                         f"Margin@Close:{r['close_margin']:.1f}%" if r['close_margin'] is not None else None,
                         f"UnschBL:{_m(r['unsch_backlog'])}" if r.get('unsch_backlog') else None,
                         f"ActRem:{_m(r['actuals_rem'])}" if r.get('actuals_rem') else None,
                         f"EvA$:{r['eva_amt']:+,.0f}" if r.get('eva_amt') is not None else None,
                         f"EvA%:{r['eva_pct']:+.1f}%" if r.get('eva_pct') is not None else None,
                         f"Pipe:{_m(r['open_pipe'])}" if r.get('open_pipe') else None]
            fin_str = '  '.join(p for p in fin_parts if p)
            if fin_str: L(f"         Financials: {fin_str}")
            snap = project_snapshot(r)
            L(f"         Snapshot: {snap}")
            if r.get('slack_intel'):
                L(f"         💬 Slack Intel: {r['slack_intel']}")
            L("")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — DATA GAPS
# ══════════════════════════════════════════════════════════════════════════════
L("")
L("=" * 80)
L("SECTION 4 — DATA GAPS")
L("=" * 80)
L("  ARMV (Rule 1D): Milestone at-risk query not in this audit run — add Q10 to next version")
L(f"  CSAT (Rule 4B): {len(csat_scores_list)} projects with scores | CSAT_OVERDUE: {sum(1 for r in results if any(v[0]=='CSAT_OVERDUE' for v in r['violations']))} | CSAT_EXEMPT: {sum(1 for r in results if any(v[0]=='CSAT_EXEMPT' for v in r['violations']))}")
L("  SteerCo (Rule 4B): Queried from pulse — check Next_Steering_Committee_Date__c field coverage")
L("  Pulse Staleness (2B): LastModifiedDate captured in pulse; calculate staleness from REPORT_DATE")
L(f"  No-Pulse total  : {len(no_pulse)} projects ({len(no_pulse_flagged)} flagged ≥$150K)")

L("")
L("=" * 80)
L(f"Audit complete — {len(results)} projects evaluated — {REPORT_DATE}")
L("=" * 80)

# ── Write report ──────────────────────────────────────────────────────────────
report_text = '\n'.join(lines)
print(report_text)

os.makedirs(OUTPUT_DIR, exist_ok=True)
FILE_STAMP = datetime.now().strftime('%Y-%m-%d_%H%M')
base_path = f"{OUTPUT_DIR}/{REGION_SLUG}_Audit_{FILE_STAMP}"

def write_txt():
    path = base_path + '.txt'
    with open(path, 'w', encoding='utf-8') as f:
        f.write(report_text)
    print(f"✅  TXT saved:  {path}")

def write_docx():
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    C_BLUE   = RGBColor(0x1F, 0x49, 0x7D)
    C_RED    = RGBColor(0xC0, 0x00, 0x00)
    C_YELLOW = RGBColor(0x7F, 0x60, 0x00)
    C_GREEN  = RGBColor(0x37, 0x86, 0x30)
    C_WM     = RGBColor(0x2E, 0x75, 0xB6)
    C_GREY   = RGBColor(0x40, 0x40, 0x40)

    HEALTH_COLOR = {
        'red':       C_RED,
        'yellow':    C_YELLOW,
        'green':     C_GREEN,
        'watermelon':C_WM,
    }

    doc = Document()
    doc.core_properties.title = f"{REGION_LABEL} Portfolio Audit {REPORT_DATE}"

    # Page layout — landscape
    for sec in doc.sections:
        sec.orientation   = 1  # WD_ORIENT.LANDSCAPE
        sec.page_width    = Inches(11)
        sec.page_height   = Inches(8.5)
        sec.left_margin   = Inches(0.75)
        sec.right_margin  = Inches(0.75)
        sec.top_margin    = Inches(0.75)
        sec.bottom_margin = Inches(0.75)

    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(9)

    # ── helpers ──────────────────────────────────────────────────────────────
    def h1(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = C_BLUE
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after  = Pt(2)

    def h2(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = C_GREY
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(1)

    def body(text, bold=False, color=None, indent=0):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent  = Inches(indent * 0.2)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(1)
        run = p.add_run(text)
        run.bold = bold
        if color: run.font.color.rgb = color

    def spacer():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def shade_row(row, hex_fill):
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), hex_fill)
            tcPr.append(shd)

    def add_hyperlink_run(paragraph, url, text, bold=False, color=None, size=9):
        part = paragraph.part
        r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        run_el = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        # Underline + blue hyperlink style
        u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
        c = OxmlElement('w:color')
        hex_col = '{:02X}{:02X}{:02X}'.format(*(tuple(color) if color else (0x44, 0x72, 0xC4)))
        c.set(qn('w:val'), hex_col); rPr.append(c)
        if bold:
            b = OxmlElement('w:b'); rPr.append(b)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2))); rPr.append(sz)
        fnt = OxmlElement('w:rFonts'); fnt.set(qn('w:ascii'), 'Calibri'); fnt.set(qn('w:hAnsi'), 'Calibri'); rPr.append(fnt)
        run_el.append(rPr)
        t = OxmlElement('w:t'); t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        run_el.append(t)
        hyperlink.append(run_el)
        paragraph._p.append(hyperlink)
        return hyperlink

    def score_rag_color(v):
        if v is None: return C_GREY
        if v >= 70: return C_GREEN
        if v >= 30: return C_YELLOW
        return C_RED

    def cell_text(cell, text, bold=False, color=None, size=9, align=None):
        p = cell.paragraphs[0]
        p.clear()
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = 'Calibri'
        if color: run.font.color.rgb = color
        if align:
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            p.alignment = {'right': WD_ALIGN_PARAGRAPH.RIGHT,
                           'center': WD_ALIGN_PARAGRAPH.CENTER}.get(align)

    RAG_COLOR = {'Red': C_RED, 'Yellow': RGBColor(0xBF,0x8F,0x00), 'Green': C_GREEN}

    def cell_plain_bullets(cell, items, size=8):
        # items: list of strings — each rendered as a plain bullet line
        first = True
        for line in items:
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            p.clear()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(0)
            p.paragraph_format.left_indent  = Inches(0.05)
            run = p.add_run(f'• {line}')
            run.font.size = Pt(size)
            run.font.name = 'Calibri'

    def fmt_financials(r, size=8):
        # Write financials as bulleted lines directly into a cell
        def _fmt_m(v):
            if v is None: return '—'
            return f'${v/1e6:.2f}M' if abs(v) >= 1_000_000 else f'${v:,.0f}'
        def _fmt_pct(v):
            return f'{v:.1f}%' if v is not None else '—'
        delivered = (r['close_margin'] or 0) - (r['bid_margin'] or 0) if (r['bid_margin'] is not None and r['close_margin'] is not None) else None
        lines = []
        if r.get('rev_treat'): lines.append(f"Type: {r['rev_treat']}")
        lines += [
            f"Bookings: {_fmt_m(r['bookings'])}",
            f"Billings: {_fmt_m(r['billings'])}",
            f"FAR: {_fmt_m(r['far'])}",
            f"Bid Margin: {_fmt_pct(r['bid_margin'])}",
            f"Delivered Margin: {(f'{delivered:+.1f}%') if delivered is not None else '—'}",
            f"Margin at Close: {_fmt_pct(r['close_margin'])}",
        ]
        if r.get('unsch_backlog'): lines.append(f"Unsch Backlog: {_fmt_m(r['unsch_backlog'])}")
        if r.get('actuals_rem'):   lines.append(f"Actuals Rem: {_fmt_m(r['actuals_rem'])}")
        if r.get('eva_amt') is not None: lines.append(f"EvA $: {r['eva_amt']:+,.0f}")
        if r.get('eva_pct') is not None: lines.append(f"EvA %: {r['eva_pct']:+.1f}%")
        if r.get('open_pipe'): lines.append(f"Open Pipe: {_fmt_m(r['open_pipe'])}")
        return lines

    def cell_bullets(cell, items, size=8):
        # items: list of (label, value) tuples — each rendered as a bullet line
        first = True
        for lbl, val in items:
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            p.clear()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(0)
            p.paragraph_format.left_indent  = Inches(0.05)
            run = p.add_run(f'• {lbl}=')
            run.font.size = Pt(size)
            run.font.name = 'Calibri'
            run.bold = False
            val_run = p.add_run(val)
            val_run.font.size = Pt(size)
            val_run.font.name = 'Calibri'
            val_run.bold = True
            val_run.font.color.rgb = RAG_COLOR.get(val, C_GREY)

    def add_table(headers, col_widths, rows_data, header_fill='1F497D', stripe='EBF3FB', repeat_header=False):
        n_cols = len(headers)
        tbl = doc.add_table(rows=1, cols=n_cols)
        tbl.style = 'Table Grid'

        # Header row
        hdr = tbl.rows[0]
        shade_row(hdr, header_fill)
        for i, (h, w) in enumerate(zip(headers, col_widths)):
            hdr.cells[i].width = Inches(w)
            cell_text(hdr.cells[i], h, bold=True,
                      color=RGBColor(0xFF,0xFF,0xFF), size=9)

        # Repeat header on every page
        if repeat_header:
            trPr = hdr._tr.get_or_add_trPr()
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)

        # Data rows
        for idx, row_vals in enumerate(rows_data):
            row = tbl.add_row()
            if idx % 2 == 1 and stripe:
                shade_row(row, stripe)
            for i, val in enumerate(row_vals):
                if isinstance(val, dict) and 'plain_bullets' in val:
                    cell_plain_bullets(row.cells[i], val['plain_bullets'])
                elif isinstance(val, dict) and 'bullets' in val:
                    cell_bullets(row.cells[i], val['bullets'])
                elif isinstance(val, dict) and 'scores' in val:
                    # {'text':..., 'color':..., 'bold':..., 'scores': [...], 'url': optional}
                    p = row.cells[i].paragraphs[0]; p.clear()
                    if val.get('url'):
                        add_hyperlink_run(p, val['url'], val.get('text',''),
                                          bold=val.get('bold', False),
                                          color=val.get('color'), size=9)
                    else:
                        r0 = p.add_run(val.get('text',''))
                        r0.bold = val.get('bold', False)
                        r0.font.size = Pt(9); r0.font.name = 'Calibri'
                        if val.get('color'): r0.font.color.rgb = val['color']
                    for slbl, sval in val['scores']:
                        if sval is None: continue
                        sep = p.add_run('  ')
                        sep.font.size = Pt(7); sep.font.name = 'Calibri'
                        chip = p.add_run(f'{slbl}:{sval:.0f}')
                        chip.bold = True; chip.font.size = Pt(7); chip.font.name = 'Calibri'
                        chip.font.color.rgb = score_rag_color(sval)
                    if val.get('subtitle'):
                        sub_p = row.cells[i].add_paragraph()
                        sub_r = sub_p.add_run(val['subtitle'])
                        sub_r.font.size = Pt(7); sub_r.font.name = 'Calibri'
                        sub_r.font.color.rgb = C_GREY
                elif isinstance(val, dict):
                    cell_text(row.cells[i], val.get('text',''),
                              bold=val.get('bold', False),
                              color=val.get('color'),
                              align=val.get('align'))
                else:
                    cell_text(row.cells[i], str(val))
        spacer()
        return tbl

    def health_color(r):
        if r['is_watermelon']: return C_WM
        h = r['health'].lower()
        return HEALTH_COLOR.get(h, C_GREY)

    def health_label(r):
        if r['is_watermelon']: return '🍉 Watermelon'
        return {'red':'🔴 Red','yellow':'🟡 Yellow','green':'🟢 Green'}.get(r['health'].lower(), r['health'])

    # ── Cover / title ─────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    run = p.add_run(f"{REGION_LABEL} PORTFOLIO AUDIT")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = C_BLUE
    p.paragraph_format.space_after = Pt(2)

    body(f"Audit Date: {REPORT_DATE}  |  Data: Live Org62  |  Auditor: SPSM/DAF", color=C_GREY)
    spacer()

    # ── Section 1: Executive Scorecard ───────────────────────────────────────
    h1("SECTION 1 — EXECUTIVE SCORECARD")

    h2("Portfolio Summary")
    add_table(
        headers    = ['Metric', 'Value'],
        col_widths = [3.0, 2.0],
        rows_data  = [
            ['Total Active Projects', str(len(results))],
            ['Portfolio Bookings',    f'${total_bk/1e6:.1f}M'],
            ['Portfolio Billings',    f'${total_bil/1e6:.1f}M'],
            ['Outstanding Backlog',   f'${total_backlog/1e6:.1f}M'],
            ['SWE/CO Projects',       f'{sum(1 for r in results if r["swe_co"])} (excluded from margin calc)'],
            ['Weighted Bid Margin',   f'{w_bid:.1f}%'],
            ['Delivered Margin',      f'{w_close-w_bid:+.1f}%'],
            ['Margin at Close',       f'{w_close:.1f}%'],
        ],
        stripe=None,
    )

    h2("Health Breakdown")
    add_table(
        headers    = ['Status', 'Count', 'Notes'],
        col_widths = [1.8, 0.7, 3.5],
        rows_data  = [
            [{'text':'🟢 Green',            'color':C_GREEN,  'bold':True}, str(self_reported_green), ''],
            [{'text':'🟢 Green (genuine)',  'color':C_GREEN,  'bold':False}, str(len(clean_green)),   ''],
            [{'text':'🍉 Watermelon',       'color':C_WM,     'bold':True}, str(len(watermelons)),   f'{wm_pct}% of self-reported greens'],
            [{'text':'🟡 Yellow',           'color':C_YELLOW, 'bold':True}, str(len(yellows)),       ''],
            [{'text':'🔴 Red',              'color':C_RED,    'bold':True}, str(len(reds)),          red_accts],
            [{'text':'⚫ No Pulse (Null)',  'bold':True},                   str(len(no_pulse)),      f'{len(no_pulse_flagged)} flagged ≥$150K'],
            [{'text':'⏸ On Hold',          'bold':True},                   str(len(on_hold)),       ', '.join(r['acct'][:20] for r in on_hold)],
        ],
        stripe=None,
    )

    h2("Financial Exposure")
    expo_rows = [
        ['Total FAR (net)',       f'${total_far:,.0f}'],
        ['Total FAR Overruns',    f'${total_far_overrun:,.0f}',],
    ]
    if total_overdue > 0:
        expo_rows.append(['Total Overdue Invoices', f'${total_overdue:,.0f}'])
    if total_rr_rev > 0:
        n_rr = sum(r['rr_count'] for r in results if r['rr_revenue'])
        expo_rows.append(['Total Revenue @ Risk', f'${total_rr_rev:,.0f}  ({n_rr} open RRs)'])
    add_table(
        headers    = ['Exposure Item', 'Amount'],
        col_widths = [3.0, 2.5],
        rows_data  = expo_rows,
        stripe=None,
    )

    # ── Section 2: Portfolio Leader Action Plan ───────────────────────────────
    h1("SECTION 2 — PORTFOLIO LEADER ACTION PLAN")

    h2("Top Priority Escalations (Reds → Yellows → Watermelons, by Bookings)")
    def esc_sort(x):
        h = x['health'].lower()
        if h == 'red' and not x['is_watermelon']:    return (0, -bk(x))
        if h == 'yellow' and not x['is_watermelon']: return (1, -bk(x))
        return (2, -bk(x))
    top10 = sorted([r for r in results if r['is_watermelon'] or r['health'].lower() in ('red', 'yellow')],
                   key=esc_sort)[:15]
    def esc_people(r):
        p = [f"PM: {r['pm']}"]
        if r.get('pm2'): p.append(f"PM2: {r['pm2']}")
        if r.get('opp_owner'): p.append(f"AP: {r['opp_owner']}")
        if r.get('owner') and r.get('owner') != 'Unassigned': p.append(f"PO: {r['owner']}")
        return p
    add_table(
        headers    = ['Project', 'Team Leadership', 'Financials', 'Status', 'Rule Codes'],
        col_widths = [2.2, 1.6, 1.5, 0.9, 1.5],
        rows_data  = [
            [
                {'text': r['name'][:55], 'color': health_color(r), 'bold': True},
                {'plain_bullets': esc_people(r)},
                {'plain_bullets': fmt_financials(r)},
                health_label(r),
                '\n'.join(fmt_violation(v) for v in r['violations'][:3]),
            ]
            for r in top10
        ],
        repeat_header=True,
    )

    if rr_at_risk:
        h2("Revenue at Risk — Unassigned Resource Requests")
        rr_rows = [
            [
                r['name'][:50],
                r['pm'],
                {'text': f"${r['rr_revenue']:,.0f}", 'align': 'right'},
                {'text': str(r['rr_count']),          'align': 'right'},
                r['rr_earliest'] or 'N/A',
                '⚠️' if r['rr_revenue'] > (r['far'] or 0) * 0.10 and (r['far'] or 0) > 0 else '',
            ]
            for r in rr_at_risk
        ]
        rr_rows.append([
            {'text': 'TOTAL', 'bold': True},
            '',
            {'text': f"${sum(r['rr_revenue'] for r in rr_at_risk):,.0f}", 'align': 'right', 'bold': True},
            {'text': str(sum(r['rr_count'] for r in rr_at_risk)),         'align': 'right', 'bold': True},
            '', '',
        ])
        add_table(
            headers    = ['Project', 'PM', 'Pending Rev', 'RRs', 'Earliest Start', 'Flag'],
            col_widths = [2.4, 1.5, 1.0, 0.4, 1.0, 0.4],
        rows_data  = rr_rows,
        )

    if neg_far_list:
        h2("Contract Overruns — Projects with Negative FAR")
        neg_far_rows = [
            [
                r['acct'][:25],
                r['name'][:50],
                r['pm'],
                f"T{r['tier']}",
                {'text': f"${r['far']:,.0f}", 'align': 'right', 'color': C_RED, 'bold': True},
                (r['far_reason'] or r['far_details'] or '—')[:50],
            ]
            for r in neg_far_list
        ]
        neg_far_rows.append([
            {'text': 'TOTAL', 'bold': True}, '', '', '',
            {'text': f"${sum(r['far'] for r in neg_far_list):,.0f}", 'align': 'right', 'bold': True, 'color': C_RED},
            '',
        ])
        add_table(
            headers    = ['Account', 'Project', 'PM', 'Tier', 'FAR', 'FAR Reason'],
            col_widths = [1.4, 2.3, 1.4, 0.4, 1.0, 2.2],
            rows_data  = neg_far_rows,
        )


    # ── Section 3: Project Detail by Tier ────────────────────────────────────
    h1("SECTION 3 — PROJECT DETAIL BY TIER")

    GROUP_COLOR = {
        'HIGH WATCH': C_BLUE,
        'RED':        C_RED,
        'YELLOW':     C_YELLOW,
        'WATERMELON': C_WM,
        'GREEN':      C_GREEN,
        'NO STATUS':  C_GREY,
    }

    for tier_num in [1, 2, 3]:
        tier_label = {1:'TIER 1 — Strategic Accounts (≥$7M)',
                      2:'TIER 2 — Growth Accounts ($1.5M–$7M)',
                      3:'TIER 3 — Volume Accounts (<$1.5M)'}[tier_num]
        tier_projs = [r for r in results if r['tier'] == tier_num]
        t_hw  = [r for r in tier_projs if r.get('high_watch')]
        t_red = [r for r in tier_projs if r['health'].lower()=='red' and not r['is_watermelon']]
        t_yel = [r for r in tier_projs if r['health'].lower()=='yellow']
        t_wm  = [r for r in tier_projs if r['is_watermelon']]
        t_grn = [r for r in tier_projs if r['health'].lower()=='green' and not r['is_watermelon']]
        t_np  = [r for r in tier_projs if not r['has_pulse']]
        t_bk  = sum(bk(r) for r in tier_projs)
        t_bil = sum(r['billings'] or 0 for r in tier_projs)

        h2(tier_label)
        add_table(
            headers    = ['Metric', 'Value'],
            col_widths = [2.5, 1.5],
            rows_data  = [
                ['Projects',    str(len(tier_projs))],
                ['Bookings',    f'${t_bk/1e6:.1f}M'],
                ['Billings',    f'${t_bil/1e6:.1f}M'],
                ['⚑ High Watch', str(len(t_hw))],
                ['🔴 Red',      str(len(t_red))],
                ['🟡 Yellow',   str(len(t_yel))],
                ['🍉 Watermelon',str(len(t_wm))],
                ['🟢 Green',    str(len(t_grn))],
                ['⚫ No Status',str(len(t_np))],
            ],
            stripe=None,
        )

        for group_key, group_label, group_projs in [
            ('RED',        '🔴 Red',         t_red),
            ('YELLOW',     '🟡 Yellow',      t_yel),
            ('WATERMELON', '🍉 Watermelon',  t_wm),
            ('GREEN',      '🟢 Green',       t_grn),
            ('NO STATUS',  '⚫ No Status',   t_np),
        ]:
            if not group_projs: continue
            gc = GROUP_COLOR[group_key]
            h2(f"{group_label} ({len(group_projs)})")
            rows = []
            for r in sorted(group_projs, key=lambda x: (-bool(x.get('high_watch')), -bk(x))):
                codes = '\n'.join(fmt_violation(v) for v in r['violations'][:4])
                snap  = project_snapshot(r)
                hw_prefix = '⚑ ' if r.get('high_watch') else ''
                people = [f"PM: {r['pm']}"]
                if r.get('pm2'): people.append(f"PM2: {r['pm2']}")
                if r.get('opp_owner'): people.append(f"AP: {r['opp_owner']}")
                if r.get('owner') and r.get('owner') != 'Unassigned': people.append(f"PO: {r['owner']}")
                rows.append([
                    {'text': f"{health_icon(r)} {pulse_indicator(r)}", 'color': gc, 'bold': True},
                    {'text': f"{hw_prefix}{r['name'][:48]}", 'color': gc, 'bold': True,
                     'url': f"https://org62.lightning.force.com/lightning/r/pse__Proj__c/{r['pid']}/view",
                     'scores': [('H&R', r.get('health_risk_score')), ('DQ', r.get('data_quality_score'))],
                     'subtitle': '  '.join(x for x in [r['acct'][:25], f"{r.get('start_dt','?')} → {r.get('end_dt','?')}", r.get('stage',''), r.get('practice','')] if x)},
                    {'plain_bullets': people},
                    {'plain_bullets': fmt_financials(r)},
                    codes,
                    {'bullets': baselines_list(r)},
                    snap + (f'\n💬 {r["slack_intel"]}' if r.get('slack_intel') else ''),
                ])
            add_table(
                headers    = ['H', 'Project', 'Team Leadership', 'Financials', 'Rules', 'Baselines', 'Summary'],
                col_widths = [0.2, 2.5, 1.4, 1.4, 1.0, 1.0, 3.0],
                rows_data  = rows,
                repeat_header = True,
            )

    # ── Section 3: Data Gaps (end) ────────────────────────────────────────────
    h1("SECTION 4 — DATA GAPS")
    add_table(
        headers    = ['Item', 'Status'],
        col_widths = [2.5, 5.0],
        rows_data  = [
            ['ARMV (Rule 1D)',         'Milestone at-risk query not in this audit run'],
            ['CSAT (Rule 4B)',         f"{len(csat_scores_list)} scores | CSAT_OVERDUE: {sum(1 for r in results if any(v[0]=='CSAT_OVERDUE' for v in r['violations']))} | No-Exempt: {sum(1 for r in results if any(v[0]=='CSAT_EXEMPT' for v in r['violations']))}"],
            ['SteerCo (Rule 4B)',      'Queried from pulse — check Next_Steering_Committee_Date__c coverage'],
            ['Pulse Staleness (2B)',   'LastModifiedDate captured; staleness calc not yet implemented'],
            ['No-Pulse total',         f'{len(no_pulse)} projects  ({len(no_pulse_flagged)} flagged ≥$150K)'],
        ],
        stripe=None,
    )

    path = base_path + '.docx'
    doc.save(path)
    print(f"✅  DOCX saved: {path}")

def write_pptx():
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from lxml import etree

    # ── Salesforce Brand Colors ───────────────────────────────────────────────
    SF_BLUE    = RGBColor(0x00, 0x70, 0xD2)   # Salesforce primary blue
    SF_NAVY    = RGBColor(0x03, 0x2D, 0x60)   # Salesforce dark navy
    SF_CLOUD   = RGBColor(0xF4, 0xF6, 0xF9)   # Salesforce cloud grey bg
    SF_RAIN    = RGBColor(0x16, 0x32, 0x5B)   # Salesforce midnight
    C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
    C_BODY     = RGBColor(0x18, 0x18, 0x18)
    C_SUBTEXT  = RGBColor(0x54, 0x65, 0x7E)
    C_RED      = RGBColor(0xBA, 0x00, 0x21)
    C_YELLOW   = RGBColor(0xB8, 0x76, 0x00)
    C_GREEN    = RGBColor(0x2E, 0x7D, 0x32)
    C_WM       = RGBColor(0x00, 0x70, 0xD2)
    C_GREY     = RGBColor(0x54, 0x65, 0x7E)
    C_LINK     = RGBColor(0x00, 0x70, 0xD2)
    FONT       = 'Salesforce Sans'   # falls back to Calibri if not installed

    HEALTH_COL = {
        'red':       C_RED,
        'yellow':    C_YELLOW,
        'green':     C_GREEN,
        'watermelon':C_WM,
    }
    STATUS_ICONS = {'red':'🔴','yellow':'🟡','green':'🟢','watermelon':'🍉','no pulse':'⚫','on hold':'⏸'}

    SLIDE_W, SLIDE_H = Inches(13.33), Inches(7.5)
    HEADER_H = Inches(0.58)
    FOOTER_H = Inches(0.28)

    prs2 = Presentation()
    prs2.slide_width  = SLIDE_W
    prs2.slide_height = SLIDE_H
    blank = prs2.slide_layouts[6]

    def _run(p, text, size=9, bold=False, color=None, italic=False, font=FONT):
        run = p.add_run()
        run.text = str(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font
        if color: run.font.color.rgb = color
        return run

    def _para(tf, first=False):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.space_before = Pt(0); p.space_after = Pt(0)
        return p

    def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
        from pptx.util import Pt as _Pt
        shape = slide.shapes.add_shape(1, left, top, width, height)
        shape.line.fill.background()
        if fill_color:
            shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
        else:
            shape.fill.background()
        if line_color:
            shape.line.color.rgb = line_color
            if line_width: shape.line.width = _Pt(line_width)
        else:
            shape.line.fill.background()
        return shape

    def add_textbox(slide, left, top, width, height, word_wrap=True):
        tb = slide.shapes.add_textbox(left, top, width, height)
        tb.text_frame.word_wrap = word_wrap
        return tb

    def add_slide_header(title_text, subtitle_text='', accent_color=SF_BLUE):
        slide = prs2.slides.add_slide(blank)
        # Background
        bg = slide.background.fill; bg.solid(); bg.fore_color.rgb = SF_CLOUD
        # Top accent bar
        add_rect(slide, 0, 0, SLIDE_W, HEADER_H, fill_color=accent_color)
        # SF logo dot (circle) — left of title
        add_rect(slide, Inches(0.18), Inches(0.13), Inches(0.32), Inches(0.32), fill_color=C_WHITE)
        # Title text
        tb = add_textbox(slide, Inches(0.6), Inches(0.08), SLIDE_W - Inches(0.8), HEADER_H, word_wrap=False)
        tf = tb.text_frame
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        _run(p, title_text, size=16, bold=True, color=C_WHITE)
        if subtitle_text:
            _run(p, f'  ·  {subtitle_text}', size=11, bold=False, color=RGBColor(0xCC, 0xE0, 0xFF))
        # Bottom accent line
        add_rect(slide, 0, SLIDE_H - FOOTER_H, SLIDE_W, FOOTER_H, fill_color=SF_NAVY)
        # Footer text
        tb2 = add_textbox(slide, Inches(0.3), SLIDE_H - FOOTER_H + Inches(0.04), SLIDE_W - Inches(0.6), FOOTER_H)
        tf2 = tb2.text_frame; p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.LEFT
        _run(p2, f'Salesforce Professional Services  ·  {REGION_LABEL} Portfolio Review  ·  {REPORT_DATE}', size=7, color=RGBColor(0xAA, 0xBB, 0xCC))
        return slide

    def body_tb(slide, left=Inches(0.3), top=HEADER_H + Inches(0.1),
                width=None, height=None):
        w = width  or (SLIDE_W - Inches(0.6))
        h = height or (SLIDE_H - HEADER_H - FOOTER_H - Inches(0.15))
        tb = add_textbox(slide, left, top, w, h)
        return tb.text_frame

    def score_col(v):
        if v is None: return C_GREY
        return C_GREEN if v >= 70 else (C_YELLOW if v >= 30 else C_RED)

    def fmt_m(v):
        if v is None: return '—'
        return f'${v/1e6:.2f}M' if abs(v) >= 1_000_000 else f'${v:,.0f}'

    def pptx_table(slide, headers, col_widths_in, rows,
                   top=None, left=Inches(0.3), height=None,
                   hdr_fill=SF_NAVY, hdr_color=C_WHITE,
                   stripe=RGBColor(0xEB, 0xF3, 0xFB),
                   row_ht=Inches(0.26), font_size=8, hdr_size=8):
        from pptx.util import Inches as _In
        from pptx.oxml.ns import qn as _qn
        from lxml import etree as _et
        n_cols = len(headers)
        n_rows = len(rows) + 1
        tbl_w  = sum(col_widths_in)
        _top   = top if top is not None else HEADER_H + Inches(0.15)
        _h     = height or (SLIDE_H - _top - FOOTER_H - Inches(0.1))
        tbl = slide.shapes.add_table(n_rows, n_cols,
                                     left, _top, _In(tbl_w), _h).table
        tbl.first_row  = True
        tbl.horz_banding = True
        # Column widths
        for ci, cw in enumerate(col_widths_in):
            tbl.columns[ci].width = _In(cw)
        # Row heights  — set uniform (can't do per-row cleanly without XML)
        for ri in range(n_rows):
            tbl.rows[ri].height = row_ht

        def _cell_text(cell, text, size=font_size, bold=False,
                       color=C_BODY, align=PP_ALIGN.LEFT, italic=False):
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = align
            p.space_before = Pt(0); p.space_after = Pt(0)
            run = p.add_run()
            run.text = str(text) if text is not None else ''
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.name = FONT
            run.font.color.rgb = color

        def _cell_fill(cell, rgb):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            solidFill = _et.SubElement(tcPr, _qn('a:solidFill'))
            srgbClr   = _et.SubElement(solidFill, _qn('a:srgbClr'))
            srgbClr.set('val', f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}')

        # Header row
        for ci, hdr in enumerate(headers):
            cell = tbl.cell(0, ci)
            _cell_fill(cell, (hdr_fill[0], hdr_fill[1], hdr_fill[2]))
            align = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
            _cell_text(cell, hdr, size=hdr_size, bold=True,
                       color=hdr_color, align=align)

        # Data rows
        for ri, row_spec in enumerate(rows):
            is_odd = ri % 2 == 1
            bg_rgb = (stripe[0], stripe[1], stripe[2]) if is_odd else (0xFF, 0xFF, 0xFF)
            for ci, cell_spec in enumerate(row_spec):
                cell = tbl.cell(ri + 1, ci)
                if is_odd:
                    _cell_fill(cell, bg_rgb)
                if isinstance(cell_spec, dict):
                    text  = cell_spec.get('text', '')
                    bold  = cell_spec.get('bold', False)
                    color = cell_spec.get('color', C_BODY)
                    align = cell_spec.get('align', PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER)
                    italic = cell_spec.get('italic', False)
                    if cell_spec.get('fill'):
                        fr = cell_spec['fill']
                        _cell_fill(cell, (fr[0], fr[1], fr[2]))
                    _cell_text(cell, text, size=font_size, bold=bold,
                               color=color, align=align, italic=italic)
                else:
                    align = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
                    _cell_text(cell, cell_spec, size=font_size,
                               color=C_BODY, align=align)
        return tbl

    # ── SLIDE 1: Cover ────────────────────────────────────────────────────────
    slide = add_slide_header(f"{REGION_LABEL} Portfolio Review", REPORT_DATE, SF_NAVY)

    # Big stat block — left column
    col_w = (SLIDE_W - Inches(1.0)) / 2
    tb = add_textbox(slide, Inches(0.5), HEADER_H + Inches(0.3), col_w, Inches(5.8))
    tf = tb.text_frame; first = True

    def stat_line(label, value, color=C_BODY, size=11, bold=False):
        nonlocal first
        p = _para(tf, first); first = False; p.space_after = Pt(4)
        _run(p, f'{label}  ', size=size, color=C_SUBTEXT)
        _run(p, value, size=size, bold=bold, color=color)

    _para(tf, True); first_p = tf.paragraphs[0]
    _run(first_p, f'{REGION_LABEL}', size=26, bold=True, color=SF_NAVY)
    p2 = _para(tf); p2.space_after = Pt(12)
    _run(p2, 'Portfolio Audit Report', size=14, color=SF_BLUE)

    stat_line('Total Projects',  str(len(results)),                         SF_NAVY,   13, True)
    stat_line('Bookings',        f'${total_bk/1e6:.1f}M',                  C_BODY,    11)
    stat_line('Billings',        f'${total_bil/1e6:.1f}M',                 C_BODY,    11)
    stat_line('Backlog',         f'${total_backlog/1e6:.1f}M',             SF_BLUE,   11)
    _bid_col = C_GREEN if w_bid > 13 else (C_YELLOW if w_bid >= 5 else C_RED)
    stat_line('Weighted Bid Margin',  f'{w_bid:.1f}%',                     _bid_col,  11)
    _del_col = C_GREEN if w_close - w_bid >= 0 else (C_YELLOW if w_close - w_bid >= -5 else C_RED)
    stat_line('Delivered Margin',     f'{w_close-w_bid:+.1f}%',            _del_col,  11)
    _clo_col = C_RED if w_close < 0 else (C_RED if w_close < w_bid - 5 else C_GREEN)
    stat_line('Margin at Close',      f'{w_close:.1f}%',                   _clo_col,  11)
    p3 = _para(tf); p3.space_before = Pt(8)
    _run(p3, f'Total FAR: {fmt_m(total_far)}  ·  Overruns: {fmt_m(total_far_overrun)}',
         size=10, color=C_RED if (total_far_overrun or 0) < -100000 else C_BODY)
    p4 = _para(tf)
    _run(p4, f'Rev @ Risk: {fmt_m(total_rr_rev)}  ·  Overdue Inv: {fmt_m(total_overdue)}',
         size=10, color=C_YELLOW)

    # Status pills — right column
    pill_top = HEADER_H + Inches(0.3)
    pill_x   = Inches(0.5) + col_w + Inches(0.2)
    pill_data = [
        ('🟢 Green',      len(clean_green),  C_GREEN, RGBColor(0xE8,0xF5,0xE9)),
        ('🍉 Watermelon', len(watermelons),  C_WM,    RGBColor(0xE3,0xF2,0xFD)),
        ('🟡 Yellow',     len(yellows),      C_YELLOW,RGBColor(0xFF,0xF8,0xE1)),
        ('🔴 Red',        len(reds),         C_RED,   RGBColor(0xFF,0xEB,0xEE)),
        ('⚫ No Pulse',   len(no_pulse),     C_GREY,  RGBColor(0xF5,0xF5,0xF5)),
    ]
    for label, count, fg, bg in pill_data:
        add_rect(slide, pill_x, pill_top, col_w - Inches(0.3), Inches(0.62), fill_color=bg,
                 line_color=fg, line_width=0.75)
        tb = add_textbox(slide, pill_x + Inches(0.1), pill_top + Inches(0.08),
                         col_w - Inches(0.5), Inches(0.46))
        tf2 = tb.text_frame; p = tf2.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        _run(p, f'{label}   ', size=11, color=C_BODY)
        _run(p, str(count), size=18, bold=True, color=fg)
        pill_top += Inches(0.72)

    # ── SLIDE 2: Executive Scorecard ─────────────────────────────────────────
    slide = add_slide_header('Executive Scorecard', f'{len(results)} Projects', SF_BLUE)

    sc_headers   = ['Tier', 'Projects', 'Bookings', '🔴 Red', '🍉 WM', '🟡 Yellow', '🟢 Green', '⚫ No Pulse', 'Bid M%', 'Margin at Close%']
    sc_col_w     = [2.1, 0.8, 1.1, 0.65, 0.65, 0.85, 0.85, 1.0, 0.8, 1.25]
    sc_rows      = []
    for tn in [1, 2, 3]:
        tp     = [r for r in results if r['tier'] == tn]
        tname  = {1:'T1 Strategic', 2:'T2 Growth', 3:'T3 Volume'}[tn]
        t_bk   = sum(r['bookings'] or 0 for r in tp)
        t_red  = sum(1 for r in tp if r['health'].lower()=='red' and not r['is_watermelon'])
        t_wm   = sum(1 for r in tp if r['is_watermelon'])
        t_yel  = sum(1 for r in tp if r['health'].lower()=='yellow')
        t_grn  = sum(1 for r in tp if r['health'].lower()=='green' and not r['is_watermelon'])
        t_np   = sum(1 for r in tp if not r['has_pulse'])
        bid_pool = [(r['bid_margin'], r['bookings'] or 0) for r in tp if r['bid_margin'] is not None and not r['swe_co']]
        clo_pool = [(r['close_margin'], r['bookings'] or 0) for r in tp if r['close_margin'] is not None and not r['swe_co']]
        _bp_bk = sum(bk for _, bk in bid_pool)
        _cp_bk = sum(bk for _, bk in clo_pool)
        t_bid_v = sum(m * bk for m, bk in bid_pool) / _bp_bk if _bp_bk else (sum(m for m, _ in bid_pool) / len(bid_pool) if bid_pool else None)
        t_clo_v = sum(m * bk for m, bk in clo_pool) / _cp_bk if _cp_bk else (sum(m for m, _ in clo_pool) / len(clo_pool) if clo_pool else None)
        t_bid_s = f"{t_bid_v:.1f}%" if t_bid_v is not None else '—'
        t_clo_s = f"{t_clo_v:.1f}%" if t_clo_v is not None else '—'
        t_bid_c = (C_GREEN if t_bid_v > 13 else C_YELLOW if t_bid_v >= 5 else C_RED) if t_bid_v is not None else C_BODY
        t_clo_c = C_RED if (t_clo_v is not None and t_clo_v < 0) else C_BODY
        sc_rows.append([
            {'text': tname, 'bold': True, 'color': SF_NAVY, 'align': PP_ALIGN.LEFT},
            str(len(tp)), fmt_m(t_bk),
            {'text': str(t_red),  'color': C_RED    if t_red  else C_BODY},
            {'text': str(t_wm),   'color': C_WM     if t_wm   else C_BODY},
            {'text': str(t_yel),  'color': C_YELLOW if t_yel  else C_BODY},
            {'text': str(t_grn),  'color': C_GREEN  if t_grn  else C_BODY},
            {'text': str(t_np),   'color': C_GREY   if t_np   else C_BODY},
            {'text': t_bid_s, 'color': t_bid_c},
            {'text': t_clo_s, 'color': t_clo_c},
        ])
    # Totals row
    sc_rows.append([
        {'text': 'TOTAL', 'bold': True, 'color': SF_NAVY, 'align': PP_ALIGN.LEFT,
         'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': str(len(results)), 'bold': True, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': fmt_m(total_bk),   'bold': True, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': str(len(reds)),        'bold': True, 'color': C_RED    if reds        else C_BODY, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': str(len(watermelons)), 'bold': True, 'color': C_WM     if watermelons else C_BODY, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': str(len(yellows)),     'bold': True, 'color': C_YELLOW if yellows     else C_BODY, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': str(len(clean_green)), 'bold': True, 'color': C_GREEN  if clean_green else C_BODY, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': str(len(no_pulse)),    'bold': True, 'color': C_GREY   if no_pulse    else C_BODY, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        {'text': f"{w_bid:.1f}%",   'bold': True, 'fill': RGBColor(0xD9, 0xE8, 0xF7),
         'color': C_GREEN if w_bid > 13 else C_YELLOW if w_bid >= 5 else C_RED},
        {'text': f"{w_close:.1f}%", 'bold': True, 'fill': RGBColor(0xD9, 0xE8, 0xF7),
         'color': C_RED if w_close < 0 else C_BODY},
    ])
    pptx_table(slide, sc_headers, sc_col_w, sc_rows,
               top=HEADER_H + Inches(0.2), row_ht=Inches(0.48), font_size=10, hdr_size=10)

    # ── SLIDES 3+: Per-tier project tables ────────────────────────────────────
    TIER_LABELS = {1:'Tier 1 — Strategic (≥$7M)', 2:'Tier 2 — Growth ($750K–$7M)', 3:'Tier 3 — Volume (<$750K)'}
    ROWS_PER_SLIDE = 12

    for tn in [1, 2, 3]:
        tier_projs = sorted(
            [r for r in results if r['tier'] == tn],
            key=lambda x: (not x.get('high_watch'),
                           {'red':0,'yellow':1,'watermelon':2,'green':3}.get(x['health'].lower(), 4),
                           -bk(x))
        )
        if not tier_projs: continue

        pt_headers  = ['St', 'Project', 'Account', 'PM / PO', 'Dates', 'Bookings', 'FAR', 'Bid%', 'Margin at Close%', 'H&R', 'DQ', 'Rules', 'Summary']
        pt_col_w    = [0.28, 2.5, 1.55, 1.45, 1.0, 0.75, 0.75, 0.55, 0.6, 0.38, 0.38, 1.4, 2.2]

        for chunk_i, chunk_start in enumerate(range(0, len(tier_projs), ROWS_PER_SLIDE)):
            chunk = tier_projs[chunk_start:chunk_start + ROWS_PER_SLIDE]
            pg_label = f'Page {chunk_i+1}' if len(tier_projs) > ROWS_PER_SLIDE else ''
            slide = add_slide_header(TIER_LABELS[tn], pg_label, SF_NAVY)

            pt_rows = []
            for r in chunk:
                h       = r['health'].lower()
                hcol    = HEALTH_COL.get('watermelon' if r['is_watermelon'] else h, C_GREY)
                icon    = STATUS_ICONS.get('watermelon' if r['is_watermelon'] else h, '⚫')
                hw_pfx  = '⚑ ' if r.get('high_watch') else ''
                pulse   = '✓' if r.get('has_pulse') else '—'
                pm      = (r['pm'] or '')
                po      = (r.get('owner') or '')
                pm_po   = f"{pm}\n{po}" if po else pm
                dates   = f"Start: {r.get('start_dt','?')}\nEnd: {r.get('end_dt','?')}"
                far_val = r.get('far')
                far_col = C_RED if (far_val or 0) < 0 else C_BODY
                bid_v   = r['bid_margin']
                clo_v   = r['close_margin']
                bid_s   = f"{bid_v:.1f}%" if bid_v is not None else '—'
                clo_s   = f"{clo_v:.1f}%" if clo_v is not None else '—'
                bid_col = (C_GREEN if bid_v > 13 else C_YELLOW if bid_v >= 5 else C_RED) if bid_v is not None else C_BODY
                clo_col = C_RED if (clo_v is not None and clo_v < 0) else C_BODY
                hr_val  = r.get('health_risk_score')
                dq_val  = r.get('data_quality_score')
                codes   = '\n'.join(fmt_violation(v) for v in r['violations'][:3])
                ov      = snip(r.get('overall_summary',''), 140)
                slack_s = r.get('slack_intel','')
                summary = ov + (f'\n💬 {slack_s[:80]}' if slack_s else '')
                pt_rows.append([
                    {'text': icon, 'color': hcol, 'bold': True},
                    {'text': hw_pfx + r['name'], 'color': hcol, 'bold': bool(hw_pfx), 'align': PP_ALIGN.LEFT},
                    {'text': r['acct'], 'color': C_SUBTEXT, 'align': PP_ALIGN.LEFT},
                    {'text': pm_po, 'color': C_BODY, 'align': PP_ALIGN.LEFT},
                    {'text': dates, 'color': C_SUBTEXT, 'align': PP_ALIGN.LEFT},
                    {'text': bks(r), 'align': PP_ALIGN.RIGHT, 'color': C_BODY},
                    {'text': fmt_m(far_val), 'color': far_col, 'align': PP_ALIGN.RIGHT},
                    {'text': bid_s, 'color': bid_col},
                    {'text': clo_s, 'color': clo_col},
                    {'text': f'{hr_val:.0f}' if hr_val is not None else '—', 'bold': True, 'color': score_col(hr_val)},
                    {'text': f'{dq_val:.0f}' if dq_val is not None else '—', 'bold': True, 'color': score_col(dq_val)},
                    {'text': codes, 'color': C_RED if codes else C_BODY, 'align': PP_ALIGN.LEFT},
                    {'text': summary, 'color': C_SUBTEXT, 'align': PP_ALIGN.LEFT},
                ])
            row_h = min(Inches(0.52), Inches((SLIDE_H - HEADER_H - FOOTER_H - Inches(0.35)) / (len(pt_rows) + 1)))
            pptx_table(slide, pt_headers, pt_col_w, pt_rows,
                       top=HEADER_H + Inches(0.15), row_ht=row_h, font_size=7, hdr_size=7)

    # ── SLIDE: Top Escalations ────────────────────────────────────────────────
    top_esc = sorted(
        [r for r in results if r['is_watermelon'] or r['health'].lower() in ('red','yellow')],
        key=lambda x: (-bool(x.get('high_watch')),
                       {'red':0,'watermelon':1,'yellow':2}.get(x['health'].lower() if not x['is_watermelon'] else 'watermelon', 3),
                       -bk(x))
    )[:15]
    slide = add_slide_header('Priority Escalations', f'Top {len(top_esc)} by severity & bookings', C_RED)

    esc_headers = ['St', 'Project', 'Account', 'PM', 'Bookings', 'FAR', 'Rules', 'H&R', 'DQ', 'Summary']
    esc_col_w   = [0.28, 2.8, 1.6, 1.5, 0.85, 0.85, 1.7, 0.38, 0.38, 3.29]
    esc_rows    = []
    for r in top_esc:
        h      = r['health'].lower()
        hcol   = HEALTH_COL.get('watermelon' if r['is_watermelon'] else h, C_GREY)
        icon   = STATUS_ICONS.get('watermelon' if r['is_watermelon'] else h, '⚫')
        hw_pfx = '⚑ ' if r.get('high_watch') else ''
        codes  = '\n'.join(fmt_violation(v) for v in r['violations'][:4])
        hr_val = r.get('health_risk_score'); dq_val = r.get('data_quality_score')
        ov     = snip(r.get('overall_summary',''), 160)
        slack_s = r.get('slack_intel','')
        summary = ov + (f'\n💬 {slack_s[:120]}' if slack_s else '')
        far_val = r.get('far')
        esc_rows.append([
            {'text': icon, 'color': hcol, 'bold': True},
            {'text': hw_pfx + r['name'], 'color': hcol, 'bold': True, 'align': PP_ALIGN.LEFT},
            {'text': r['acct'], 'color': C_SUBTEXT, 'align': PP_ALIGN.LEFT},
            {'text': r['pm'] or '', 'color': C_BODY, 'align': PP_ALIGN.LEFT},
            {'text': bks(r), 'align': PP_ALIGN.RIGHT, 'color': C_BODY},
            {'text': fmt_m(far_val), 'align': PP_ALIGN.RIGHT,
             'color': C_RED if (far_val or 0) < 0 else C_BODY},
            {'text': codes, 'color': C_RED if codes else C_BODY, 'align': PP_ALIGN.LEFT},
            {'text': f'{hr_val:.0f}' if hr_val is not None else '—', 'bold': True, 'color': score_col(hr_val)},
            {'text': f'{dq_val:.0f}' if dq_val is not None else '—', 'bold': True, 'color': score_col(dq_val)},
            {'text': summary, 'color': C_SUBTEXT, 'align': PP_ALIGN.LEFT, 'italic': bool(slack_s)},
        ])
    row_h = min(Inches(0.55), Inches((SLIDE_H - HEADER_H - FOOTER_H - Inches(0.35)) / (len(esc_rows) + 1)))
    pptx_table(slide, esc_headers, esc_col_w, esc_rows,
               top=HEADER_H + Inches(0.15), hdr_fill=C_RED,
               row_ht=row_h, font_size=8, hdr_size=8)

    # ── SLIDE: Revenue at Risk ────────────────────────────────────────────────
    if rr_at_risk:
        slide = add_slide_header('Revenue at Risk — Open Resource Requests', '', C_YELLOW)

        rr_headers = ['Project', 'PM', 'Pending Revenue', 'RRs', 'Earliest RR', 'Resource Status']
        rr_col_w   = [3.8, 1.8, 1.6, 0.5, 1.1, 1.5]
        rr_rows    = []
        for r in rr_at_risk:
            elev  = ' ⚠' if r['rr_revenue'] > (r['far'] or 0) * 0.10 and (r['far'] or 0) > 0 else ''
            res   = r['resource_s'] or ''
            rflag = ' ⚠' if res in ('Red','Yellow') else ''
            res_col = C_RED if res=='Red' else (C_YELLOW if res=='Yellow' else C_BODY)
            rr_rows.append([
                {'text': r['name'], 'color': C_BODY, 'align': PP_ALIGN.LEFT},
                {'text': r['pm'] or '', 'color': C_SUBTEXT, 'align': PP_ALIGN.LEFT},
                {'text': f"${r['rr_revenue']:,.0f}{elev}", 'bold': True, 'color': C_RED, 'align': PP_ALIGN.RIGHT},
                {'text': str(r['rr_count']), 'align': PP_ALIGN.CENTER},
                {'text': str(r['rr_earliest'] or 'N/A'), 'color': C_BODY},
                {'text': f'{res}{rflag}', 'color': res_col, 'bold': res in ('Red','Yellow')},
            ])
        # Totals row
        rr_total = sum(r['rr_revenue'] for r in rr_at_risk)
        rr_cnt   = sum(r['rr_count']   for r in rr_at_risk)
        rr_rows.append([
            {'text': 'TOTAL', 'bold': True, 'color': SF_NAVY, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
            {'text': '', 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
            {'text': f'${rr_total:,.0f}', 'bold': True, 'color': C_RED,   'align': PP_ALIGN.RIGHT, 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
            {'text': str(rr_cnt), 'bold': True, 'align': PP_ALIGN.CENTER,  'fill': RGBColor(0xD9, 0xE8, 0xF7)},
            {'text': '', 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
            {'text': '', 'fill': RGBColor(0xD9, 0xE8, 0xF7)},
        ])
        row_h = min(Inches(0.42), Inches((SLIDE_H - HEADER_H - FOOTER_H - Inches(0.35)) / (len(rr_rows) + 1)))
        pptx_table(slide, rr_headers, rr_col_w, rr_rows,
                   top=HEADER_H + Inches(0.15), hdr_fill=RGBColor(0xB8, 0x76, 0x00),
                   row_ht=row_h, font_size=9, hdr_size=9)

    path = base_path + '.pptx'
    prs2.save(path)
    print(f"✅  PPTX saved: {path}")

# ── Dispatch ──────────────────────────────────────────────────────────────────
def write_html():
    import json as _json
    import html as _html
    html_version = os.environ.get('NEXT_HTML_VERSION', '')

    def status_label(r):
        if r['stage'] == 'On Hold': return 'On Hold'
        if r['is_watermelon']: return 'Watermelon'
        h = r['health'].lower()
        if h == 'red':    return 'Red'
        if h == 'yellow': return 'Yellow'
        if h == 'green':  return 'Green'
        return 'No Pulse'

    def fmt_m(v):
        if v is None: return ''
        return f'${v/1e6:.2f}M' if abs(v) >= 1_000_000 else f'${v:,.0f}'

    def fmt_pct(v):
        return f'{v:.1f}%' if v is not None else ''

    rows_data = []
    for r in results:
        dd = (r['close_margin'] - r['bid_margin']) if (r['bid_margin'] is not None and r['close_margin'] is not None) else None
        viol_codes = ', '.join(v[0] for v in r['violations'])
        bl = ' | '.join(f"{lbl}={v}" for lbl, v in baselines_list(r))
        team = []
        if r['pm']:        team.append(f"PM: {r['pm']}")
        if r.get('pm2'):   team.append(f"PM2: {r['pm2']}")
        if r.get('opp_owner'): team.append(f"AP: {r['opp_owner']}")
        if r.get('exec_sponsor'): team.append(f"ES: {r['exec_sponsor']}")
        if r.get('acct_owner'): team.append(f"AO: {r['acct_owner']}")
        if r.get('owner') and r.get('owner') != 'Unassigned': team.append(f"PO: {r['owner']}")
        po = r.get('owner','') or ''
        rows_data.append({
            'name':          r['name'],
            'acct':          r['acct'],
            'status':        status_label(r),
            'tier':          r['tier'],
            'team':          ' · '.join(team),
            'po':            po,
            'pid':           r['pid'],
            'region':        REGION_LABEL,
            'url':           f"https://org62.lightning.force.com/lightning/r/pse__Proj__c/{r['pid']}/view",
            'rev_treat':     r.get('rev_treat') or '',
            'billing_type':  r.get('billing_type') or '',
            'fmt_pipe':      fmt_m(r.get('open_pipe')) if r.get('open_pipe') else '',
            'health_risk':   r.get('health_risk_score'),
            'data_quality':  r.get('data_quality_score'),
            'bookings':      r['bookings'] or 0,
            'billings':      r['billings'] or 0,
            'far':           r['far'] or 0,
            'overdue_inv':   r['overdue_inv'] or 0,
            'rr_revenue':    r['rr_revenue'] or 0,
            'bid_margin_raw':   r['bid_margin'],
            'close_margin_raw': r['close_margin'],
            'rules':         viol_codes,
            'baselines':     bl,
            'summary':       ' '.join((r.get('overall_summary') or '').split())[:500],
            'leadership':    ' '.join((r.get('leadership_notes') or '').split())[:400],
            'swe_co':        r['swe_co'],
            'high_watch':    bool(r.get('high_watch')),
            'has_pulse':     r.get('has_pulse', False),
            'pulse_trend':   r.get('trend') or '',
            'pulse_updated': (r.get('last_updated') or '')[:10],
            'pulse_scope':   r.get('scope_s') or '',
            'pulse_sched':   r.get('sched_s') or '',
            'pulse_budget':  r.get('budget_s') or '',
            'pulse_resource':r.get('resource_s') or '',
            'pulse_customer':r.get('customer_s') or '',
            'pulse_action':  ' '.join((r.get('action_needed') or '').split())[:300],
            'pulse_steerco': r['steerco_date'].isoformat() if r.get('steerco_date') else '',
            'pulse_golive':  r['next_golive'].isoformat() if r.get('next_golive') else '',
            'slack_intel':   r.get('slack_intel') or '',
            'start_dt':      r['start_dt'].isoformat() if r.get('start_dt') else '',
            'end_dt':        r['end_dt'].isoformat() if r.get('end_dt') else '',
            'stage':         r.get('stage') or '',
            'practice':      r.get('practice') or '',
            'unsch_backlog': r.get('unsch_backlog') or 0,
            'actuals_rem':   r.get('actuals_rem') or 0,
            'eva_amt':       r.get('eva_amt'),
            'eva_pct':       r.get('eva_pct'),
            'csat_score':    r.get('csat_score'),
            'fmt_unsch':     fmt_m(r.get('unsch_backlog')) if r.get('unsch_backlog') else '',
            'fmt_actuals':   fmt_m(r.get('actuals_rem')) if r.get('actuals_rem') else '',
            'fmt_eva_amt':   (f"+${r['eva_amt']:,.0f}" if r.get('eva_amt') is not None and r['eva_amt'] >= 0 else f"-${abs(r['eva_amt']):,.0f}" if r.get('eva_amt') is not None else ''),
            'fmt_eva_pct':   (f"{r['eva_pct']:+.1f}%" if r.get('eva_pct') is not None else ''),
            'far_reason':    r.get('far_reason') or '',
            'far_subreason': r.get('far_subreason') or '',
            'fmt_bookings':  fmt_m(r['bookings']),
            'fmt_billings':  fmt_m(r['billings']),
            'fmt_bid':       fmt_pct(r['bid_margin']),
            'fmt_close':     fmt_pct(r['close_margin']),
            'fmt_delivered': (f'{dd:+.1f}%' if dd is not None else ''),
            'fmt_far':       fmt_m(r['far']),
            'fmt_overdue':   fmt_m(r['overdue_inv']) if r['overdue_inv'] else '',
            'fmt_rr':        fmt_m(r['rr_revenue']) if r['rr_revenue'] else '',
            'gdc_total':     r.get('gdc_total') or 0,
            'gdc_india':     r.get('gdc_india') or 0,
            'gdc_pct':       round(r['gdc_pct'] * 100, 1) if r.get('gdc_pct') is not None else None,
            'gdc_resources': r.get('gdc_resources') or [],
        })

    # Unique portfolio owners for dropdown
    po_list = sorted(set(r['po'] for r in rows_data if r['po'] and r['po'] != 'Unassigned'))
    po_options = '\n'.join(f'<option value="{_html.escape(p)}">{_html.escape(p)}</option>' for p in po_list)

    rows_json = _json.dumps({'generated': REPORT_DATE, 'region': REGION_LABEL, 'rows': rows_data}, ensure_ascii=False)

    # DB credentials for self-serve assignments (embedded at publish time)
    import base64 as _b64_html
    _db_cid    = os.environ.get('DB_CLIENT_ID', '')
    _db_csec   = os.environ.get('DB_CLIENT_SECRET', '')
    _tile_id_h = os.environ.get('TILE_ID', '817')
    _db_b64    = _b64_html.b64encode(f"{_db_cid}:{_db_csec}".encode()).decode() if (_db_cid and _db_csec) else ''

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ACC Portfolio Audit — {REPORT_DATE}</title>
<script src="/assets/tile-db.js"></script>
<style>
  :root {{
    --navy:    #0D2B4E;
    --blue:    #1F497D;
    --lblue:   #4472C4;
    --red:     #C0392B;
    --yellow:  #D4AC0D;
    --green:   #1E8449;
    --wm:      #E74C3C;
    --grey:    #6B7280;
    --bg:      #F8F9FA;
    --card:    #FFFFFF;
    --border:  #E5E7EB;
    --text:    #111827;
    --subtext: #6B7280;
    --stripe:  #F1F5FB;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{ height: 100%; overflow: hidden; }}
  body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; display: flex; flex-direction: column; }}

  /* ── Header ── */
  .page-header {{ background: var(--blue); color: #fff; padding: 7px 18px; display: flex; align-items: center; justify-content: space-between; }}
  .page-header h1 {{ font-size: 14px; font-weight: 700; letter-spacing: .3px; }}
  .page-header .meta {{ font-size: 10px; opacity: .75; margin-top: 1px; }}

  /* ── Scorecard ── */
  .scorecard {{ display: flex; flex-wrap: wrap; gap: 6px; padding: 8px 18px; background: var(--card); border-bottom: 1px solid var(--border); transition: padding .2s; }}
  .scorecard.collapsed {{ padding-top: 0; padding-bottom: 0; overflow: hidden; max-height: 0; }}
  .scorecard-bar {{ display: flex; align-items: center; justify-content: space-between; padding: 4px 18px; background: var(--card); border-bottom: 1px solid var(--border); cursor: pointer; user-select: none; }}
  .scorecard-bar:hover {{ background: #f0f4ff; }}
  .scorecard-bar .sc-toggle {{ font-size: 11px; color: var(--subtext); display: flex; align-items: center; gap: 5px; }}
  .scorecard-bar .sc-summary {{ font-size: 11px; color: var(--subtext); }}
  .stat {{ background: var(--bg); border: 1px solid var(--border); border-radius: 6px; padding: 6px 10px; min-width: 90px; }}
  .stat .val {{ font-size: 16px; font-weight: 700; color: var(--blue); }}
  .stat .lbl {{ font-size: 9px; text-transform: uppercase; letter-spacing: .5px; color: var(--subtext); margin-top: 1px; }}
  .stat.red    .val {{ color: var(--red); }}
  .stat.yellow .val {{ color: var(--yellow); }}
  .stat.green  .val {{ color: var(--green); }}
  .stat.wm     .val {{ color: var(--wm); }}

  /* ── Region tabs ── */
  .region-tabs {{ display: flex; gap: 0; background: var(--navy); padding: 0 18px; }}
  .rtab {{ background: none; border: none; border-bottom: 3px solid transparent; color: rgba(255,255,255,.65); font-size: 12px; font-weight: 600; padding: 7px 18px; cursor: pointer; letter-spacing: .3px; transition: color .15s, border-color .15s; }}
  .rtab:hover {{ color: #fff; }}
  .rtab.active {{ color: #fff; border-bottom-color: #4FC3F7; }}

  /* ── Controls ── */
  .controls {{ padding: 6px 18px; background: var(--card); border-bottom: 1px solid var(--border); display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }}
  .controls input[type=search] {{
    border: 1px solid var(--border); border-radius: 6px; padding: 4px 10px;
    font-size: 12px; width: 240px; outline: none;
  }}
  .controls input[type=search]:focus {{ border-color: var(--lblue); box-shadow: 0 0 0 2px #4472c420; }}
  .filter-group {{ display: flex; gap: 4px; flex-wrap: wrap; align-items: center; }}
  .pill {{
    border: 1px solid var(--border); border-radius: 20px; padding: 2px 9px;
    font-size: 11px; cursor: pointer; background: var(--bg); color: var(--subtext);
    transition: all .15s; user-select: none;
  }}
  .pill:hover {{ border-color: var(--lblue); color: var(--lblue); }}
  .pill.active {{ background: var(--blue); color: #fff; border-color: var(--blue); font-weight: 600; }}
  .pill.red-pill.active    {{ background: var(--red);    border-color: var(--red); }}
  .pill.yellow-pill.active {{ background: var(--yellow); border-color: var(--yellow); }}
  .pill.green-pill.active  {{ background: var(--green);  border-color: var(--green); }}
  .pill.wm-pill.active     {{ background: var(--wm);     border-color: var(--wm); }}
  .sep {{ width: 1px; height: 18px; background: var(--border); align-self: center; }}
  .count-badge {{ font-size: 10px; color: var(--subtext); margin-left: 3px; }}
  .po-select {{
    border: 1px solid var(--border); border-radius: 6px; padding: 3px 8px;
    font-size: 11px; background: var(--bg); color: var(--text); cursor: pointer;
    outline: none; min-width: 150px;
  }}
  .po-select:focus {{ border-color: var(--lblue); }}

  /* ── Multi-select combobox ── */
  .ms-wrap {{ position: relative; display: inline-block; }}
  .ms-btn {{
    border: 1px solid var(--border); border-radius: 6px; padding: 3px 24px 3px 8px;
    font-size: 11px; background: var(--bg); color: var(--text); cursor: pointer;
    outline: none; min-width: 120px; text-align: left; white-space: nowrap;
    user-select: none; position: relative;
  }}
  .ms-btn::after {{ content: '▾'; position: absolute; right: 7px; top: 50%; transform: translateY(-50%); font-size: 10px; color: var(--subtext); }}
  .ms-btn.ms-active {{ border-color: var(--blue); color: var(--blue); background: #eef2ff; font-weight: 600; }}
  .ms-dropdown {{
    display: none; position: absolute; top: calc(100% + 3px); left: 0; z-index: 9000;
    background: var(--card); border: 1px solid var(--border); border-radius: 8px;
    box-shadow: 0 4px 16px rgba(0,0,0,.12); min-width: 180px; max-height: 260px;
    overflow-y: auto; padding: 4px 0;
  }}
  .ms-dropdown.open {{ display: block; }}
  .ms-item {{
    display: flex; align-items: center; gap: 7px; padding: 5px 12px;
    font-size: 11px; cursor: pointer; color: var(--text);
  }}
  .ms-item:hover {{ background: #f3f6ff; }}
  .ms-item input[type=checkbox] {{ accent-color: var(--blue); cursor: pointer; margin: 0; }}
  .ms-clear {{ display:block; padding: 4px 12px 5px; font-size:10px; color:var(--blue); cursor:pointer; border-top:1px solid var(--border); }}
  .ms-clear:hover {{ text-decoration: underline; }}

  /* ── Table ── */
  .table-wrap {{ overflow: auto; padding: 0 24px 32px; flex: 1; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
  thead th {{
    background: var(--blue); color: #fff; font-size: 11px; font-weight: 600;
    text-transform: uppercase; letter-spacing: .4px; padding: 9px 10px;
    text-align: left; position: sticky; top: 0; z-index: 10; cursor: pointer;
    white-space: nowrap; user-select: none;
  }}
  thead th:hover {{ background: var(--lblue); }}
  thead th .sort-arrow {{ margin-left: 4px; opacity: .4; font-size: 10px; }}
  thead th.sorted-asc  .sort-arrow::after {{ content: '▲'; opacity: 1; }}
  thead th.sorted-desc .sort-arrow::after {{ content: '▼'; opacity: 1; }}
  thead th:not(.sorted-asc):not(.sorted-desc) .sort-arrow::after {{ content: '⇅'; }}

  tbody tr {{ border-bottom: 1px solid var(--border); transition: background .1s; }}
  tbody tr:nth-child(even) {{ background: var(--stripe); }}
  tbody tr:hover {{ background: #EBF3FB; }}
  tbody td {{ padding: 8px 10px; vertical-align: top; }}

  /* status badge */
  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; white-space: nowrap; }}
  .badge-Red        {{ background: #FDEDEC; color: var(--red); }}
  .badge-Yellow     {{ background: #FEF9E7; color: #9A7D0A; }}
  .badge-Green      {{ background: #EAFAF1; color: var(--green); }}
  .badge-Watermelon {{ background: #FDEDEC; color: var(--wm); }}
  .badge-No-Pulse   {{ background: #F3F4F6; color: var(--grey); }}
  .badge-On-Hold    {{ background: #EDE7F6; color: #6A1B9A; }}

  .hw-flag {{ display: inline-block; font-size: 10px; background: #EBF5FB; color: var(--lblue); border-radius: 4px; padding: 1px 5px; margin-right: 4px; font-weight: 600; }}
  .score-chip      {{ display: inline-block; font-size: 10px; border-radius: 4px; padding: 1px 5px; font-weight: 700; border: 1px solid; }}
  .score-green {{ background: #EAFAF1; color: #1E8449; border-color: #A9DFBF; }}
  .score-yellow {{ background: #FEF9E7; color: #9A7D0A; border-color: #F9E79F; }}
  .score-red   {{ background: #FDEDEC; color: #C0392B; border-color: #FADBD8; }}

  /* tier badge */
  .tier {{ display: inline-block; font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 4px; background: #E8EAF6; color: #3949AB; }}

  /* team list */
  .team-list {{ line-height: 1.7; }}
  .team-list span {{ display: block; font-size: 11px; color: var(--subtext); }}
  .team-list span b {{ color: var(--text); font-weight: 600; }}
  .role-abbr {{ position: relative; display: inline-block; }}
  .role-abbr .role-tip {{
    display: none; position: absolute; bottom: calc(100% + 4px); left: 0;
    background: #1e293b; color: #fff; font-size: 10px; font-weight: 400;
    white-space: nowrap; padding: 3px 7px; border-radius: 4px;
    box-shadow: 0 2px 6px rgba(0,0,0,.25); pointer-events: none; z-index: 9999;
  }}
  .role-abbr:hover .role-tip {{ display: block; }}
  .res-hover-wrap {{ position: relative; display: inline-block; cursor: pointer; }}
  .res-hover-tip {{
    display: none; position: absolute; top: calc(100% + 4px); left: 50%; transform: translateX(-50%);
    background: var(--card); color: var(--text); font-size: 10px; line-height: 1.4;
    padding: 6px 8px; border-radius: 8px; width: max-content;
    border: 1px solid var(--border); box-shadow: 0 4px 16px rgba(0,0,0,.12);
    pointer-events: auto; z-index: 9999; white-space: nowrap;
  }}
  .res-hover-tip.open {{ display: block; }}
  .res-hover-tip .res-scroll {{ max-height: 240px; overflow-y: auto; overflow-x: visible; }}
  .res-hover-tip table {{ border-collapse: collapse; font-size: 9px; }}
  .res-hover-tip th {{ font-size: 8px; color: #fff; font-weight: 600; padding: 2px 8px; text-align: left; text-transform: uppercase; letter-spacing: .3px; background: var(--blue); }}
  .res-hover-tip thead tr {{ background: var(--blue); }}
  .res-hover-tip td {{ padding: 2px 8px 2px 0; vertical-align: middle; color: var(--text); white-space: nowrap; }}
  .res-close-btn {{ font-size: 8px; cursor: pointer; color: #fff; border: none; border-radius: 3px; padding: 1px 6px; line-height: 14px; background: var(--blue); margin-left: 8px; }}
  .res-close-btn:hover {{ opacity: .8; }}
  .tier-edit {{ font-size: 11px; font-weight: 700; color: var(--blue); border: 1px solid var(--border); border-radius: 4px; padding: 1px 2px; background: var(--bg); cursor: pointer; }}
  .po-edit-lbl {{ font-size: 10px; color: var(--subtext); cursor: pointer; white-space: nowrap; }}
  .po-edit-lbl:hover {{ color: var(--blue); }}
  .po-edit-inp input {{ font-size: 10px; border: 1px solid var(--lblue); border-radius: 3px; padding: 1px 4px; width: 140px; outline: none; }}

  /* financials */
  .fin-2col  {{ display: flex; gap: 12px; }}
  .fin-col   {{ flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 1px; font-size: 11px; }}
  .fin-item  {{ display: flex; justify-content: space-between; align-items: baseline; gap: 6px; }}
  .fin-item .fk {{ color: var(--text); font-weight: 600; white-space: nowrap; flex-shrink: 0; }}
  .fin-item .fv {{ white-space: nowrap; text-align: right; min-width: 0; overflow: hidden; text-overflow: ellipsis; color: var(--subtext); }}
  .fin-item .fv.neg {{ color: var(--red); }}
  .fin-item .fv.pos {{ color: var(--green); }}
  .fin-sep   {{ border: none; border-top: 1px solid var(--border); margin: 2px 0; }}
  .fin-extras {{ margin-top: 3px; padding-top: 3px; border-top: 1px solid var(--border); display: flex; flex-direction: column; gap: 1px; font-size: 11px; }}

  /* baselines */
  .bl-list {{ font-size: 11px; line-height: 1.8; }}
  .bl-Red    {{ color: var(--red);    font-weight: 700; }}
  .bl-Yellow {{ color: var(--yellow); font-weight: 700; }}
  .bl-Green  {{ color: var(--grey); }}

  /* rules */
  .rule-code {{
    display: inline-block; font-size: 10px; background: #FFF3CD; color: #856404;
    border-radius: 3px; padding: 1px 5px; margin: 1px 2px 1px 0; border: 1px solid #FFEAA7;
    cursor: help; position: relative;
  }}
  .rule-code.neg  {{ background: #FDEDEC; color: var(--red); border-color: #FADBD8; }}
  .rule-code.warn {{ background: #FFF3CD; color: #856404; border-color: #FFEAA7; }}
  .rule-code .rule-tip {{
    display: none; position: absolute; bottom: calc(100% + 4px); left: 50%;
    transform: translateX(-50%); background: #1F2937; color: #fff;
    font-size: 11px; border-radius: 5px; padding: 6px 10px; white-space: nowrap;
    z-index: 100; pointer-events: none; box-shadow: 0 2px 8px rgba(0,0,0,.25);
    font-weight: normal; letter-spacing: 0;
  }}
  .rule-code .rule-tip::after {{
    content: ''; position: absolute; top: 100%; left: 50%; transform: translateX(-50%);
    border: 5px solid transparent; border-top-color: #1F2937;
  }}
  .rule-code:hover .rule-tip {{ display: block; }}

  /* summary expandable */
  .summary-cell {{ max-width: 320px; }}
  .summary-short {{ color: var(--text); line-height: 1.5; font-size: 11px; }}
  .summary-full {{ display: none; color: var(--subtext); font-size: 11px; margin-top: 4px; line-height: 1.5; border-top: 1px dashed var(--border); padding-top: 4px; }}
  .expand-btn {{ font-size: 10px; color: var(--lblue); cursor: pointer; margin-top: 3px; display: inline-block; }}
  .expand-btn:hover {{ text-decoration: underline; }}
  tr.expanded .summary-full {{ display: block; }}

  /* project name */
  .proj-name {{ font-weight: 600; color: var(--text); font-size: 12px; }}
  .proj-acct {{ font-size: 12px; font-weight: 600; color: var(--blue); margin-top: 2px; }}
  .proj-link {{ color: inherit; text-decoration: none; }}
  .proj-link:hover {{ color: var(--lblue); text-decoration: underline; }}
  .pulse-icon {{ display: inline-block; cursor: default; font-size: 13px; vertical-align: middle; margin-left: 4px; }}
  #pulse-tooltip {{
    display: none; position: fixed; width: 480px; max-width: 90vw; background: #fff; border: 1px solid #ddd;
    border-radius: 8px; padding: 12px 16px; box-shadow: 0 4px 20px rgba(0,0,0,.18);
    font-size: 12px; line-height: 1.65; color: var(--text); z-index: 9999; white-space: normal;
    pointer-events: none;
  }}
  .pulse-dim {{ display: flex; gap: 4px; flex-wrap: wrap; margin: 5px 0 3px; }}
  .pdim {{ font-size: 10px; border-radius: 3px; padding: 1px 5px; font-weight: 700; }}
  .pdim-green  {{ background:#EAFAF1; color:#1a6b3c; }}
  .pdim-yellow {{ background:#FEF9E7; color:#7d6608; }}
  .pdim-red    {{ background:#FDEDEC; color:#922b21; }}
  .pdim-grey   {{ background:#f0f0f0; color:#666; }}

  /* no-results */
  .no-results {{ text-align: center; padding: 48px; color: var(--subtext); font-size: 14px; }}

  /* ── Group rows ── */
  .group-row td {{
    background: #EBF3FB; border-top: 2px solid var(--lblue); border-bottom: 1px solid var(--border);
    padding: 7px 10px; font-size: 11px; font-weight: 700; color: var(--navy);
    letter-spacing: .3px; cursor: pointer; user-select: none;
  }}
  .group-tier-header td {{
    background: var(--navy) !important; color: #fff !important;
    border-top: 3px solid var(--navy) !important; border-bottom: 2px solid #0a2240 !important;
    padding: 9px 12px !important; font-size: 12px !important; letter-spacing: .4px;
  }}
  .group-tier-header td .grp-toggle {{ color: #fff !important; opacity: .8; }}
  .group-tier-header td .grp-meta {{ color: rgba(255,255,255,.75) !important; }}
  .group-tier-header td .grp-pills .grp-pill {{ opacity: .92; }}
  .group-acct-header td {{ padding-left: 28px !important; border-left: 4px solid var(--lblue); }}
  .grp-btn {{ font-size: 11px; padding: 3px 9px; border-radius: 5px; border: 1px solid var(--border); background: var(--bg); color: var(--text); cursor: pointer; }}
  .grp-btn:hover {{ background: var(--blue); color: #fff; border-color: var(--blue); }}
  .group-row td .grp-toggle {{ margin-right: 6px; font-size: 10px; opacity: .6; }}
  .group-row td .grp-meta {{ font-weight: 400; color: var(--subtext); margin-left: 10px; font-size: 11px; }}
  .group-row td .grp-pills {{ display: inline-flex; gap: 4px; margin-left: 10px; }}
  .group-row td .grp-pill {{ font-size: 10px; border-radius: 10px; padding: 1px 7px; font-weight: 700; }}
  .grp-pill-nopulse {{ background: #e8e8e8; color: #555; }}
  .grp-pill-red    {{ background:#FDEDEC; color:var(--red); }}
  .grp-pill-yellow {{ background:#FEF9E7; color:#9A7D0A; }}
  .grp-pill-wm     {{ background:#FDEDEC; color:var(--wm); }}
  .grp-pill-green  {{ background:#EAFAF1; color:var(--green); }}
  tr.grp-hidden {{ display: none; }}

  /* column widths */
  .col-status  {{ width: 90px; }}
  .col-tier    {{ width: 40px; text-align: center; }}
  .col-project {{ width: 260px; overflow-wrap: break-word; word-break: break-word; }}
  .col-team     {{ width: 180px; }}
  .col-resource {{ width: 90px; }}
  .col-fin      {{ width: 230px; }}
  .col-rules   {{ width: 130px; }}
  .col-bl      {{ width: 160px; }}
  .col-summary {{ min-width: 200px; }}
</style>
</head>
<body>
<div id="pulse-tooltip"></div>
<div id="far-tooltip" style="display:none;position:fixed;max-width:260px;background:#fff;border:1px solid #ddd;border-radius:6px;padding:8px 10px;box-shadow:0 3px 12px rgba(0,0,0,.12);font-size:11px;line-height:1.5;z-index:9999;pointer-events:none"></div>

<div class="page-header">
  <div>
    <h1>ACC Portfolio Audit</h1>
    <div class="meta">Data: <span id="data-refresh-date">loading…</span>{'&nbsp;|&nbsp;v' + html_version if html_version else ''}</div>
  </div>
  <div style="text-align:right;font-size:12px;opacity:.8" id="header-summary">
    {len(results)} projects &nbsp;|&nbsp; {fmt_m(total_bk)} bookings
  </div>
</div>

<div class="scorecard-bar" onclick="toggleScorecard()">
  <span class="sc-summary" id="sc-bar-summary">{len(results)} projects &nbsp;·&nbsp; {fmt_m(total_bk)} bookings &nbsp;·&nbsp; 🟢 {len(clean_green)+len(watermelons)} &nbsp;·&nbsp; 🟡 {len(yellows)} &nbsp;·&nbsp; 🔴 {len(reds)}</span>
  <span class="sc-toggle"><span id="sc-chevron">▲</span> Scorecard</span>
</div>
<div class="scorecard" id="scorecard">
  <div class="stat"><div class="val" id="sc-total">{len(results)}</div><div class="lbl">Projects</div></div>
  <div class="stat"><div class="val" id="sc-bookings">{fmt_m(total_bk)}</div><div class="lbl">Bookings</div></div>
  <div class="stat">
    <div class="val" id="sc-avg-bk">{fmt_m(total_bk / len(results) if results else 0)}</div>
    <div class="lbl">Avg Bookings</div>
    <div style="margin-top:4px;font-size:11px;color:var(--subtext)">Median: <span id="sc-median-bk" style="font-weight:600;color:var(--text)">{fmt_m(median_bk)}</span></div>
  </div>
  <div class="stat"><div class="val" id="sc-billings">{fmt_m(total_bil)}</div><div class="lbl">Billings</div></div>
  <div class="stat"><div class="val" id="sc-backlog">{fmt_m(total_bk - total_bil)}</div><div class="lbl">Backlog</div></div>
  <div class="stat green" style="min-width:130px">
    <div class="val" id="sc-green-total">{len(clean_green) + len(watermelons)}</div>
    <div class="lbl">Green</div>
    <div style="margin-top:5px;font-size:11px;color:var(--subtext);padding-left:10px;line-height:1.7">
      <div>✅ Genuine: <span id="sc-green">{len(clean_green)}</span></div>
      <div>🍉 Watermelon: <span id="sc-wm">{len(watermelons)}</span></div>
    </div>
  </div>
  <div class="stat yellow"><div class="val" id="sc-yellow">{len(yellows)}</div><div class="lbl">Yellow</div></div>
  <div class="stat red"><div class="val" id="sc-red">{len(reds)}</div><div class="lbl">Red</div></div>
  <div class="stat"><div class="val" id="sc-nopulse">{len(no_pulse)}</div><div class="lbl">No Pulse</div></div>
  <div class="stat"><div class="val" id="sc-far">{fmt_m(total_far)}</div><div class="lbl">Total FAR</div></div>
  <div class="stat red"><div class="val" id="sc-overdue">{fmt_m(total_overdue)}</div><div class="lbl">Overdue Inv</div></div>
  <div class="stat yellow"><div class="val" id="sc-rr">{fmt_m(total_rr_rev)}</div><div class="lbl">RR Rev Risk</div></div>
  <div class="stat"><div class="val" id="sc-bid" style="color:{'var(--green)' if w_bid > 13 else 'var(--yellow)' if w_bid >= 5 else 'var(--red)'}">{w_bid:.1f}%</div><div class="lbl">Wtd Bid Margin</div></div>
  <div class="stat {'red' if w_close - w_bid < -5 else 'yellow' if w_close - w_bid < 0 else 'green'}"><div class="val" id="sc-delivered">{w_close - w_bid:+.1f}%</div><div class="lbl">Delivered Margin</div></div>
  <div class="stat"><div class="val" id="sc-close" style="color:{'var(--red)' if w_close < 0 else ''}">{w_close:.1f}%</div><div class="lbl">Wtd Margin at Close</div></div>
  <div class="stat"><div class="val" id="sc-delta">{w_close - w_bid:+.1f}%</div><div class="lbl">Margin Delta</div></div>
  <div class="stat {'red' if avg_hr is not None and avg_hr < 60 else 'yellow' if avg_hr is not None and avg_hr < 75 else 'green' if avg_hr is not None else ''}"><div class="val" id="sc-avg-hr">{f'{avg_hr:.0f}' if avg_hr is not None else '—'}</div><div class="lbl">Avg Health Score</div></div>
  <div class="stat {'red' if avg_dq is not None and avg_dq < 60 else 'yellow' if avg_dq is not None and avg_dq < 75 else 'green' if avg_dq is not None else ''}"><div class="val" id="sc-avg-dq">{f'{avg_dq:.0f}' if avg_dq is not None else '—'}</div><div class="lbl">Avg Data Quality</div></div>
  <div class="stat {'red' if avg_csat is not None and avg_csat < 3.5 else 'yellow' if avg_csat is not None and avg_csat < 4.0 else 'green' if avg_csat is not None else ''}"><div class="val" id="sc-avg-csat">{f'{avg_csat:.1f}' if avg_csat is not None else '—'}</div><div class="lbl">Avg CSAT</div></div>
</div>

<div class="region-tabs">
  <button class="rtab active" onclick="setRegion(null,this)">ACC (All)</button>
  <button class="rtab" onclick="setRegion('AMER TMT',this)">TMT</button>
  <button class="rtab" onclick="setRegion('AMER CBS',this)">CBS</button>
</div>

<div class="controls">
  <input type="search" id="search" placeholder="🔍  Search project, account, PM, rules…" oninput="applyFilters()">
  <div class="sep"></div>
  <div class="filter-group">
    <label style="font-size:12px;color:var(--subtext);align-self:center">Status:</label>
    <div class="ms-wrap" id="ms-status-wrap">
      <button class="ms-btn" id="ms-status-btn" onclick="toggleMs('status')">All Statuses</button>
      <div class="ms-dropdown" id="ms-status-dd">
        <label class="ms-item"><input type="checkbox" value="Red" onchange="msChange('status')"> 🔴 Red</label>
        <label class="ms-item"><input type="checkbox" value="Yellow" onchange="msChange('status')"> 🟡 Yellow</label>
        <label class="ms-item"><input type="checkbox" value="Watermelon" onchange="msChange('status')"> 🍉 Watermelon</label>
        <label class="ms-item"><input type="checkbox" value="Green" onchange="msChange('status')"> 🟢 Green</label>
        <label class="ms-item"><input type="checkbox" value="No Pulse" onchange="msChange('status')"> ⚫ No Pulse</label>
        <label class="ms-item"><input type="checkbox" value="On Hold" onchange="msChange('status')"> ⏸ On Hold</label>
        <span class="ms-clear" onclick="msClear('status')">Clear</span>
      </div>
    </div>
  </div>
  <div class="sep"></div>
  <div class="filter-group">
    <span class="pill" data-filter="hw" data-val="hw" onclick="togglePill(this)">⚑ High Watch Only</span>
    <span class="pill" data-filter="swe" data-val="swe" onclick="togglePill(this)">SWE / ARI</span>
  </div>
  <div class="sep"></div>
  <div class="filter-group">
    <label style="font-size:12px;color:var(--subtext);align-self:center">Tier:</label>
    <div class="ms-wrap" id="ms-tier-wrap">
      <button class="ms-btn" id="ms-tier-btn" onclick="toggleMs('tier')">All Tiers</button>
      <div class="ms-dropdown" id="ms-tier-dd">
        <label class="ms-item"><input type="checkbox" value="1" onchange="msChange('tier')"> T1 Strategic</label>
        <label class="ms-item"><input type="checkbox" value="2" onchange="msChange('tier')"> T2 Growth</label>
        <label class="ms-item"><input type="checkbox" value="3" onchange="msChange('tier')"> T3 Volume</label>
        <span class="ms-clear" onclick="msClear('tier')">Clear</span>
      </div>
    </div>
  </div>
  <div class="sep"></div>
  <div class="filter-group">
    <label style="font-size:12px;color:var(--subtext);align-self:center">PO:</label>
    <div class="ms-wrap" id="ms-po-wrap">
      <button class="ms-btn" id="ms-po-btn" onclick="toggleMs('po')">All Owners</button>
      <div class="ms-dropdown" id="ms-po-dd">
        {chr(10).join(f'<label class="ms-item"><input type="checkbox" value="{_html.escape(p)}" onchange="msChange(\'po\')"> {_html.escape(p)}</label>' for p in po_list)}
        <span class="ms-clear" onclick="msClear('po')">Clear</span>
      </div>
    </div>
  </div>
  <div class="sep"></div>
  <div class="filter-group">
    <label style="font-size:12px;color:var(--subtext);align-self:center">Rules:</label>
    <div class="ms-wrap" id="ms-rule-wrap">
      <button class="ms-btn" id="ms-rule-btn" onclick="toggleMs('rule')">All Rules</button>
      <div class="ms-dropdown" id="ms-rule-dd">
        <label class="ms-item"><input type="checkbox" value="MARGIN_RED" onchange="msChange('rule')"> MARGIN_RED</label>
        <label class="ms-item"><input type="checkbox" value="MARGIN_YELLOW" onchange="msChange('rule')"> MARGIN_YELLOW</label>
        <label class="ms-item"><input type="checkbox" value="FAR_RED_NEG" onchange="msChange('rule')"> FAR_RED_NEG</label>
        <label class="ms-item"><input type="checkbox" value="FAR_RED_UNDERUTIL" onchange="msChange('rule')"> FAR_RED_UNDERUTIL</label>
        <label class="ms-item"><input type="checkbox" value="FAR_YELLOW" onchange="msChange('rule')"> FAR_YELLOW</label>
        <label class="ms-item"><input type="checkbox" value="RR_RISK" onchange="msChange('rule')"> RR_RISK</label>
        <label class="ms-item"><input type="checkbox" value="GDC_LOW" onchange="msChange('rule')"> GDC_LOW</label>
        <label class="ms-item"><input type="checkbox" value="OVERDUE_INV" onchange="msChange('rule')"> OVERDUE_INV</label>
        <label class="ms-item"><input type="checkbox" value="SWE_BURNING_HOT" onchange="msChange('rule')"> SWE_BURNING_HOT</label>
        <label class="ms-item"><input type="checkbox" value="NO_PULSE" onchange="msChange('rule')"> NO_PULSE</label>
        <label class="ms-item"><input type="checkbox" value="NO_STEERCO" onchange="msChange('rule')"> NO_STEERCO</label>
        <label class="ms-item"><input type="checkbox" value="MISSING_PTG" onchange="msChange('rule')"> MISSING_PTG</label>
        <label class="ms-item"><input type="checkbox" value="END_DATE_PAST" onchange="msChange('rule')"> END_DATE_PAST</label>
        <label class="ms-item"><input type="checkbox" value="END_DATE_UPCOMING" onchange="msChange('rule')"> END_DATE_UPCOMING</label>
        <label class="ms-item"><input type="checkbox" value="CSAT_OVERDUE" onchange="msChange('rule')"> CSAT_OVERDUE</label>
        <label class="ms-item"><input type="checkbox" value="CSAT_EXEMPT" onchange="msChange('rule')"> CSAT_EXEMPT</label>
        <span class="ms-clear" onclick="msClear('rule')">Clear</span>
      </div>
    </div>
  </div>
  <div class="sep"></div>
  <div class="filter-group">
    <label style="font-size:12px;color:var(--subtext);align-self:center">Group by:</label>
    <select id="group-by" class="po-select" onchange="applyFilters()">
      <option value="">None</option>
      <option value="acct">Account Name</option>
      <option value="tier_acct">Tier &amp; Account</option>
      <option value="acct_po">Account &amp; Portfolio Owner</option>
      <option value="po">Portfolio Owner</option>
      <option value="tier">Tier</option>
    </select>
    <span id="grp-collapse-btns" style="display:none;margin-left:6px">
      <button class="grp-btn" onclick="collapseAll()">⊟ Collapse All</button>
      <button class="grp-btn" onclick="expandAll()">⊞ Expand All</button>
    </span>
  </div>
</div>

<div class="table-wrap">
  <table id="main-table">
    <thead>
      <tr>
        <th class="col-status"  data-col="status">Status<span class="sort-arrow"></span></th>
        <th class="col-tier"    data-col="tier">T<span class="sort-arrow"></span></th>
        <th class="col-project" data-col="name">Project / Account<span class="sort-arrow"></span></th>
        <th class="col-team"     data-col="team">Team Leadership<span class="sort-arrow"></span></th>
        <th class="col-resource" data-col="gdc_pct">Resourcing<span class="sort-arrow"></span></th>
        <th class="col-fin"      data-col="bookings">Financials<span class="sort-arrow"></span></th>
        <th class="col-rules"   data-col="rules">Portfolio Assurance<span class="sort-arrow"></span></th>
        <th class="col-bl"      data-col="baselines">Baselines<span class="sort-arrow"></span></th>
        <th class="col-summary" data-col="summary">Summary</th>
      </tr>
    </thead>
    <tbody id="table-body"></tbody>
  </table>
  <div class="no-results" id="no-results" style="display:none">No projects match your filters.</div>
</div>

<script>
// Resource popup: open on mouseenter, close on mouseleave or close button
function resOpen(id) {{
  var el = document.getElementById(id);
  if (el) el.classList.add('open');
}}
function resClose(id) {{
  var el = document.getElementById(id);
  if (el) el.classList.remove('open');
}}

function toggleScorecard() {{
  const sc = document.getElementById('scorecard');
  const ch = document.getElementById('sc-chevron');
  const collapsed = sc.classList.toggle('collapsed');
  ch.textContent = collapsed ? '▼' : '▲';
}}

const _INLINE = {rows_json};
let RAW = _INLINE.rows || [];

async function loadData() {{
  // Try to load fresher data from _data/ (works when platform routes it correctly)
  try {{
    const [tmt, cbs] = await Promise.all([
      fetch('./_data/acc_amer_tmt_data.json', {{credentials:'include'}}).then(r=>r.ok?r.json():null).catch(()=>null),
      fetch('./_data/acc_amer_cbs_data.json', {{credentials:'include'}}).then(r=>r.ok?r.json():null).catch(()=>null),
    ]);
    const fetched = [tmt, cbs].filter(Boolean).flatMap(d => d.rows || []);
    if (fetched.length > 0) {{
      RAW = fetched;
      const latest = [tmt, cbs].filter(Boolean).map(d => d.generated).filter(Boolean).sort().pop() || '';
      document.getElementById('data-refresh-date').textContent = 'Refreshed ' + latest;
    }} else {{
      document.getElementById('data-refresh-date').textContent = 'Data: ' + (_INLINE.generated || '{REPORT_DATE}');
    }}
  }} catch(e) {{
    document.getElementById('data-refresh-date').textContent = 'Data: ' + (_INLINE.generated || '{REPORT_DATE}');
  }}
  try {{ await loadAssignments(); }} catch(e) {{ console.warn('loadAssignments (non-fatal):', e); }}
  applyFilters();
}}

// ── DB credentials (embedded at publish time) ─────────────────────────────────
const _DB_CRED  = '{_db_b64}';
const _TILE_ID  = {_tile_id_h};

// ── Self-serve assignments ────────────────────────────────────────────────────
async function loadAssignments() {{
  if (!_DB_CRED) return;
  try {{
    await TileDB.open();
    const rows = await TileDB.query('SELECT pid, tier, po FROM assignments');
    rows.forEach(a => {{
      const r = RAW.find(x => x.pid === a.pid);
      if (r) {{
        if (a.tier != null) r.tier = parseInt(a.tier);
        if (a.po)  r.po   = a.po;
      }}
    }});
  }} catch(e) {{ console.warn('loadAssignments:', e); }}
}}

async function saveAssignment(pid, field, value) {{
  if (!_DB_CRED) return;
  try {{
    await fetch('/api/tiles/' + _TILE_ID + '/db/write', {{
      method: 'POST',
      headers: {{ 'Authorization': 'Basic ' + _DB_CRED, 'Content-Type': 'application/json' }},
      body: JSON.stringify({{
        table: 'assignments', mode: 'upsert',
        key_column: 'pid', columns: ['pid', field],
        rows: [{{ pid, [field]: field === 'tier' ? parseInt(value) : value }}],
      }}),
    }});
    const r = RAW.find(x => x.pid === pid);
    if (r) r[field] = field === 'tier' ? parseInt(value) : value;
    applyFilters();
  }} catch(e) {{ console.warn('saveAssignment:', e); }}
}}

function startPoEdit(lbl, pid, currentVal) {{
  const wrap = lbl.closest('.po-edit-wrap');
  wrap.querySelector('.po-edit-lbl').style.display = 'none';
  const inpWrap = wrap.querySelector('.po-edit-inp');
  inpWrap.style.display = '';
  const inp = inpWrap.querySelector('input');
  inp.value = currentVal;
  inp.focus();
  inp.select();
}}
function commitPoEdit(inp, pid) {{
  const wrap = inp.closest('.po-edit-wrap');
  const val = inp.value.trim();
  if (val) saveAssignment(pid, 'po', val);
  wrap.querySelector('.po-val').textContent = val || '—';
  wrap.querySelector('.po-edit-inp').style.display = 'none';
  wrap.querySelector('.po-edit-lbl').style.display = '';
}}
function cancelPoEdit(inp) {{
  const wrap = inp.closest('.po-edit-wrap');
  wrap.querySelector('.po-edit-inp').style.display = 'none';
  wrap.querySelector('.po-edit-lbl').style.display = '';
}}

const STATUS_ORDER = {{Red:0, Yellow:1, Watermelon:2, Green:3, 'No Pulse':4, 'On Hold':5}};
const TIER_LABEL   = {{1:'T1',2:'T2',3:'T3'}};

const RULE_KEY = {{
  'MARGIN_RED':        'Close Margin critically below Bid Margin (RED threshold)',
  'MARGIN_YELLOW':     'Close Margin moderately below Bid Margin (YELLOW threshold)',
  'FAR_RED_NEG':       'FAR is negative — project forecasted to overrun budget',
  'FAR_RED_UNDERUTIL': 'FAR >50% remaining with <30% schedule left — severe underutilisation',
  'FAR_YELLOW':        'FAR utilisation warning — moderate schedule/spend mismatch',
  'RR_RISK':           'Open Resource Requests with pending revenue — staffing gap risk',
  'GDC_LOW':           'GDC India assigned share ≤ 65% — more than ⅓ of the team is non-GDC',
  'OVERDUE_INV':       'Overdue invoices outstanding — cash collection risk',
  'SWE_BURNING_HOT':   'SWE burn rate alert — hours/spend tracking off-plan',
  'NO_PULSE':          'No pulse submitted — project status unknown (governance violation)',
  'NO_STEERCO':     'Next Steering Committee Date required for projects ≥ $750K. ACTION: If exempt (SEH, Advisory, etc.) set date to 01/01/2100.',
  'MISSING_PTG':       'Missing project timeline/go-live date or steerco date',
  'END_DATE_PAST':     'End date has passed — project may need extension or closure',
  'END_DATE_UPCOMING': 'End date within 45 days — renewal/extension decision needed',
}};

const RULE_GROUP = {{
  'MARGIN_RED':        'Margin',
  'MARGIN_YELLOW':     'Margin',
  'FAR_RED_NEG':       'FAR',
  'FAR_RED_UNDERUTIL': 'FAR',
  'FAR_YELLOW':        'FAR',
  'RR_RISK':           'Resource',
  'GDC_LOW':           'Resource',
  'OVERDUE_INV':       'Invoice',
  'SWE_BURNING_HOT':   'SWE Burn',
  'NO_PULSE':          'Governance',
  'NO_STEERCO':        'Governance',
  'MISSING_PTG':       'Governance',
  'END_DATE_PAST':     'End Date',
  'END_DATE_UPCOMING': 'End Date',
}};

let sortCol = 'status', sortDir = 1;
let filterHW = false, filterSWE = false;
let filterStatuses = new Set(), filterTiers = new Set(), filterPOs = new Set(), filterRules = new Set();
let groupBy = '';
let filterRegion = null;
const TIER_NAME = {{1:'T1 — Strategic (≥$7M)', 2:'T2 — Growth ($750K–$7M)', 3:'T3 — Volume (<$750K)'}};

function fmtMoney(v) {{
  if (!v && v !== 0) return '—';
  const abs = Math.abs(v);
  const sign = v < 0 ? '-' : '';
  if (abs >= 1e6) return sign + '$' + (abs/1e6).toFixed(2) + 'M';
  return sign + '$' + abs.toLocaleString('en-US', {{maximumFractionDigits:0}});
}}

function fmtPct(v) {{
  if (v === null || v === undefined) return '—';
  return v.toFixed(1) + '%';
}}

function bidMarginColor(v) {{
  if (v === null || v === undefined) return '';
  if (v > 13) return 'color:var(--green)';
  if (v >= 5)  return 'color:var(--yellow)';
  return 'color:var(--red)';
}}

function statusBadge(r) {{
  const s = r.status;
  const cls = s.replace(' ','-');
  const icons = {{Red:'🔴',Yellow:'🟡',Green:'🟢',Watermelon:'🍉','No Pulse':'⚫','On Hold':'⏸'}};
  const trendArrow = {{
    'Up':'↑', 'Improving':'↑', 'Improving Slightly':'↗',
    'Stable':'→',
    'Down':'↓', 'Declining':'↓', 'Declining Slightly':'↘', 'Worsening':'↓',
  }};
  const trendColor = {{
    'Up':'var(--green)', 'Improving':'var(--green)', 'Improving Slightly':'var(--green)',
    'Stable':'var(--subtext)',
    'Down':'var(--red)', 'Declining':'var(--red)', 'Declining Slightly':'var(--yellow)', 'Worsening':'var(--red)',
  }};
  const arrow = r.pulse_trend ? (trendArrow[r.pulse_trend] || '') : '';
  const arrowCol = r.pulse_trend ? (trendColor[r.pulse_trend] || 'var(--subtext)') : '';
  const trendHtml = arrow ? ` <span style="font-size:13px;font-weight:700;color:${{arrowCol}};vertical-align:middle" title="Trend: ${{r.pulse_trend}}">${{arrow}}</span>` : '';
  let html = `<span class="badge badge-${{cls}}">${{icons[s]||''}} ${{s}}</span>${{trendHtml}}`;
  function scoreCls(v) {{ return v >= 70 ? 'score-green' : v >= 30 ? 'score-yellow' : 'score-red'; }}
  const scores = [];
  if (r.health_risk  != null) scores.push(`<span class="score-chip ${{scoreCls(r.health_risk)}}">H&amp;R&nbsp;${{r.health_risk.toFixed(0)}}</span>`);
  if (r.data_quality != null) scores.push(`<span class="score-chip ${{scoreCls(r.data_quality)}}">DQ&nbsp;${{r.data_quality.toFixed(0)}}</span>`);
  if (r.csat_score   != null) scores.push(`<span class="score-chip ${{scoreCls(r.csat_score * 20)}}">CSAT&nbsp;${{r.csat_score.toFixed(1)}}</span>`);
  if (scores.length) html += `<div style="margin-top:4px">${{scores.join(' ')}}</div>`;
  html += pulseIcon(r);
  return html;
}}

function dimCls(v) {{
  if (!v) return 'pdim-grey';
  const lv = v.toLowerCase();
  if (lv === 'green') return 'pdim-green';
  if (lv === 'yellow') return 'pdim-yellow';
  if (lv === 'red') return 'pdim-red';
  return 'pdim-grey';
}}

const _ptt = document.getElementById('pulse-tooltip');
function _showPulse(el, e) {{
  _ptt.innerHTML = el.dataset.pulse;
  _ptt.style.display = 'block';
  _movePulse(e);
}}
function _movePulse(e) {{
  const pad = 12, w = 290;
  let x = e.clientX + pad;
  if (x + w > window.innerWidth - pad) x = e.clientX - w - pad;
  _ptt.style.left = x + 'px';
  _ptt.style.top  = (e.clientY + pad) + 'px';
}}
function _hidePulse() {{ _ptt.style.display = 'none'; }}

const _ftt = document.getElementById('far-tooltip');
function _showFar(el, e) {{
  _ftt.innerHTML = '<strong style="color:#888;font-size:10px">FAR Reason</strong><div style="margin-top:3px">' + el.dataset.far + '</div>';
  _ftt.style.display = 'block'; _moveFar(e);
}}
function _moveFar(e) {{
  const pad = 12, w = 260;
  let x = e.clientX + pad;
  if (x + w > window.innerWidth - pad) x = e.clientX - w - pad;
  _ftt.style.left = x + 'px'; _ftt.style.top = (e.clientY + pad) + 'px';
}}
function _hideFar() {{ _ftt.style.display = 'none'; }}

function pulseIcon(r) {{
  if (!r.has_pulse) {{
    const tip = [
      `<strong>No Pulse on File</strong><br>This project has no active pulse record in Salesforce.`,
      (r.bookings||0) >= 150000 ? `<br><span style="color:#c0392b;font-weight:700">⚠ Governance violation — Rule 2A (Bookings ≥ $150K)</span>` : '',
      r.slack_intel ? `<div style="margin-top:6px;padding-top:6px;border-top:1px solid #eee"><span style="font-size:10px;font-weight:700;color:#5b5e6d">💬 SLACK INTEL</span><div style="margin-top:2px;color:#1a1a2e">${{r.slack_intel}}</div></div>` : '',
    ].join('');
    return `<span class="pulse-icon" data-pulse="${{tip.replace(/"/g,'&quot;')}}"
      onmouseenter="_showPulse(this,event)" onmousemove="_movePulse(event)" onmouseleave="_hidePulse()">⚫</span>`;
  }}
  const dims = [
    ['Scope',    r.pulse_scope],
    ['Schedule', r.pulse_sched],
    ['Budget',   r.pulse_budget],
    ['Resource', r.pulse_resource],
    ['Customer', r.pulse_customer],
  ];
  const dimHtml = `<div class="pulse-dim">${{dims.map(([lbl,val]) =>
    `<span class="pdim ${{dimCls(val)}}">${{lbl}}: ${{val||'—'}}</span>`).join('')}}</div>`;
  const tip = [
    `<strong>Pulse on File</strong>`,
    r.pulse_updated ? `<div style="color:#888;font-size:10px">Updated: ${{r.pulse_updated}}</div>` : '',
    dimHtml,
    r.pulse_trend   ? `<div><b>Trend:</b> ${{r.pulse_trend}}</div>` : '',
    r.pulse_steerco ? `<div><b>SteerCo:</b> ${{r.pulse_steerco}}</div>` : '',
    r.pulse_golive  ? `<div><b>Next Go-Live:</b> ${{r.pulse_golive}}</div>` : '',
    r.summary       ? `<div style="margin-top:6px;padding-top:5px;border-top:1px solid #eee"><span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px">Project Progress Summary</span><div style="margin-top:3px">${{r.summary.slice(0,250)}}${{r.summary.length>250?'…':''}}</div></div>` : '',
    r.pulse_action  ? `<div style="margin-top:6px;padding-top:5px;border-top:1px solid #eee"><span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#c0392b">Action Needed from Leadership</span><div style="margin-top:3px;color:#c0392b">${{r.pulse_action.slice(0,200)}}${{r.pulse_action.length>200?'…':''}}</div></div>` : '',
    r.slack_intel   ? `<div style="margin-top:6px;padding-top:6px;border-top:1px solid #eee"><span style="font-size:10px;font-weight:700;color:#5b5e6d">💬 SLACK INTEL</span><div style="margin-top:2px;color:#1a1a2e">${{r.slack_intel}}</div></div>` : '',
  ].join('');
  return `<span class="pulse-icon" data-pulse="${{tip.replace(/"/g,'&quot;')}}"
    onmouseenter="_showPulse(this,event)" onmousemove="_movePulse(event)" onmouseleave="_hidePulse()">📋</span>`;
}}

const ROLE_LABELS = {{
  'PM':  'Project Manager',
  'PM2': 'Program Manager',
  'AP':  'ProServ Account Partner',
  'AO':  'License Account Executive',
  'PO':  'Portfolio Owner',
  'ES':  'Executive Sponsor',
}};
function teamHtml(team) {{
  if (!team) return '';
  return '<div class="team-list">' +
    team.split(' · ').map(t => {{
      const colon = t.indexOf(':');
      const role = t.slice(0,colon);
      const name = t.slice(colon+1).trim();
      const tip  = ROLE_LABELS[role] || role;
      return `<span><b class="role-abbr">${{role}}:<span class="role-tip">${{tip}}</span></b> ${{name}}</span>`;
    }}).join('') + '</div>';
}}

function resourcingHtml(r) {{
  if (r.gdc_total === 0) {{
    return `<div style="font-size:11px;color:var(--subtext)">No assigned<br>resources</div>`;
  }}
  const isLow    = r.gdc_pct !== null && r.gdc_pct <= 65;
  const pctColor = isLow ? 'var(--red)' : 'var(--green)';
  const pctBg    = isLow ? '#FDEDEC'    : '#EAFAF1';
  const pct      = r.gdc_pct !== null ? r.gdc_pct + '%' : '?';

  // Build resource tooltip table
  const fmtH = h => (h != null && h !== 0) ? Math.round(h).toLocaleString('en-US') : '—';
  const rows = (r.gdc_resources || []).map(res => {{
    const flag = res.region === 'GDC India' ? '🇮🇳' : '🇺🇸';
    const remStyle = res.rem_hrs != null && res.rem_hrs < 0 ? ' style="color:var(--red)"' : '';
    return `<tr>
      <td>${{res.name || '—'}}</td>
      <td>${{res.role || '—'}}</td>
      <td>${{flag}} ${{res.region || '—'}}</td>
      <td style="text-align:right">${{fmtH(res.est_hrs)}}</td>
      <td style="text-align:right">${{fmtH(res.act_hrs)}}</td>
      <td style="text-align:right"${{remStyle}}>${{res.rem_hrs != null ? fmtH(res.rem_hrs) : '—'}}</td>
      <td style="text-align:right">${{fmtH(res.sch_hrs)}}</td>
      <td>${{res.start || '—'}}</td>
      <td>${{res.end || '—'}}</td>
    </tr>`;
  }}).join('');

  const tipId = 'rt-' + r.pid;
  const tip = `<div class="res-hover-tip" id="${{tipId}}">
    <div style="display:flex;align-items:center;margin-bottom:4px">
      <span style="font-size:9px;font-weight:600;color:var(--subtext);flex:1">ASSIGNED RESOURCES (${{r.gdc_total}})</span>
      <button class="res-close-btn" onclick="resClose('${{tipId}}');event.stopPropagation()">✕ Close</button>
    </div>
    <div class="res-scroll">
    <table>
      <thead><tr>
        <th>Resource</th><th>Role</th><th>Region</th>
        <th style="text-align:right">Planned</th><th style="text-align:right">Actual</th><th style="text-align:right">Remaining</th>
        <th style="text-align:right">Scheduled</th>
        <th>Start</th><th>End</th>
      </tr></thead>
      <tbody>${{rows}}</tbody>
    </table>
    </div>
    <div style="margin-top:4px;padding-top:3px;border-top:1px solid var(--border);font-size:8px;color:var(--subtext)">🇮🇳 GDC India &nbsp; 🇺🇸 Other regions</div>
  </div>`;

  return `<div class="res-hover-wrap" onmouseenter="resOpen('${{tipId}}')" onmouseleave="resClose('${{tipId}}')">
    <div style="font-size:10px;color:var(--subtext);margin-bottom:2px">GDC India</div>
    <div style="display:inline-block;padding:2px 7px;border-radius:5px;background:${{pctBg}};color:${{pctColor}};font-size:13px;font-weight:700">${{pct}}</div>
    <div style="font-size:10px;color:var(--subtext);margin-top:2px">${{r.gdc_india}}/${{r.gdc_total}} assigned</div>
    ${{tip}}
  </div>`;
}}

function finHtml(r) {{
  const neg = v => v && String(v).startsWith('-') ? ' neg' : (v && String(v).startsWith('+') ? ' pos' : '');
  const typeLabel = r.billing_type || '—';

  // FAR with hover popup for reason codes
  const farTip = [r.far_reason, r.far_subreason].filter(Boolean).join(' · ');
  const farCell = farTip
    ? `<span class="far-tip-wrap" data-far="${{farTip.replace(/"/g,'&quot;')}}"
         onmouseenter="_showFar(this,event)" onmousemove="_moveFar(event)" onmouseleave="_hideFar()"
         style="cursor:help;border-bottom:1px dotted var(--subtext)">${{r.fmt_far||'—'}} ⓘ</span>`
    : (r.fmt_far || '—');

  // Overdue invoice — always shown; red when > 0
  const overdueStyle = r.overdue_inv > 0 ? ' style="color:var(--red)"' : '';
  const overdueVal   = r.fmt_overdue || '—';

  const fi = (k, v, cls='', sty='') =>
    `<div class="fin-item"><span class="fk">${{k}}:</span><span class="fv${{cls}}"${{sty ? ` style="${{sty}}"` : ''}}>${{v}}</span></div>`;

  // Left sub-column: revenue metrics
  const leftCol = `<div class="fin-col">
    ${{fi('Type', typeLabel)}}
    ${{fi('Bookings', r.fmt_bookings||'—')}}
    ${{fi('Billings', r.fmt_billings||'—')}}
    ${{fi('FAR', farCell, r.far < 0 ? ' neg' : '')}}
    ${{fi('Overdue Inv', overdueVal, '', r.overdue_inv > 0 ? 'color:var(--red)' : '')}}
  </div>`;

  // Right sub-column: margin metrics
  const rightCol = `<div class="fin-col">
    ${{fi('Bid Margin', r.fmt_bid||'—', '', bidMarginColor(r.bid_margin_raw))}}
    ${{fi('Delivered', r.fmt_delivered||'—', neg(r.fmt_delivered))}}
    ${{fi('Margin@Close', r.fmt_close||'—', '', r.close_margin_raw !== null && r.close_margin_raw < 0 ? 'color:var(--red)' : '')}}
    ${{r.fmt_eva_amt ? fi('EvA $', r.fmt_eva_amt, neg(r.fmt_eva_amt)) : ''}}
    ${{r.fmt_eva_pct ? fi('EvA %', r.fmt_eva_pct, neg(r.fmt_eva_pct)) : ''}}
  </div>`;

  // Optional extras below both columns
  const extraItems = [
    r.fmt_unsch   ? fi('Unsch Backlog', r.fmt_unsch) : '',
    r.fmt_actuals ? fi('Actuals Rem',   r.fmt_actuals) : '',
    r.fmt_pipe    ? fi('Open Pipe',     r.fmt_pipe, '', 'color:var(--lblue)') : '',
  ].filter(Boolean).join('');

  const extrasHtml = extraItems ? `<div class="fin-extras">${{extraItems}}</div>` : '';

  return `<div><div class="fin-2col">${{leftCol}}${{rightCol}}</div>${{extrasHtml}}</div>`;
}}

function rulesHtml(rules) {{
  if (!rules) return '';
  return rules.split(', ').filter(Boolean).map(c => {{
    const isNeg = c.includes('NEG') || c.includes('RED') || c === 'END_DATE_PAST';
    const isWarn = c.includes('YELLOW') || c === 'END_DATE_UPCOMING';
    const tip = RULE_KEY[c] || c;
    const grp = RULE_GROUP[c] ? `<span style="opacity:.65;font-size:9px;margin-right:2px">${{RULE_GROUP[c]}}:</span>` : '';
    const cls = isNeg ? ' neg' : isWarn ? ' warn' : '';
    return `<span class="rule-code${{cls}}">${{grp}}${{c}}<span class="rule-tip">${{tip}}</span></span>`;
  }}).join('');
}}

function baselinesHtml(bl) {{
  if (!bl) return '';
  return '<div class="bl-list">' +
    bl.split(' | ').map(part => {{
      const eq = part.indexOf('=');
      const lbl = part.slice(0,eq);
      const val = part.slice(eq+1);
      return `<span class="bl-${{val||''}}">${{lbl}}=${{val}}</span>`;
    }}).join('<br>') + '</div>';
}}

function summaryHtml(r, idx) {{
  const short = r.summary ? r.summary.slice(0,140) + (r.summary.length > 140 ? '…' : '') : '';
  const progressBlock = r.summary
    ? `<div style="margin-top:5px"><span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--subtext)">Project Progress Summary</span><div style="margin-top:2px">${{r.summary}}</div></div>`
    : '';
  const actionBlock = r.pulse_action
    ? `<div style="margin-top:5px;padding-top:4px;border-top:1px dashed var(--border)"><span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--red)">Action Needed from Leadership</span><div style="margin-top:2px;color:var(--red)">${{r.pulse_action}}</div></div>`
    : '';
  const leadershipBlock = r.leadership
    ? `<div style="margin-top:5px;padding-top:4px;border-top:1px dashed var(--border)"><span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--subtext)">Leadership Notes</span><div style="margin-top:2px">${{r.leadership}}</div></div>`
    : '';
  const slackBlock = r.slack_intel
    ? `<div style="margin-top:5px;padding-top:5px;border-top:1px dashed var(--border)"><span style="font-size:10px;font-weight:700;color:#5b5e6d">💬 SLACK INTEL</span><div style="margin-top:2px;font-size:11px;color:var(--text)">${{r.slack_intel}}</div></div>`
    : '';
  const full = progressBlock + actionBlock + leadershipBlock + slackBlock;
  const shortDisplay = short
    || (r.pulse_action ? `<span style="font-size:11px;color:var(--red)">⚠ ${{r.pulse_action.slice(0,120)}}${{r.pulse_action.length>120?'…':''}}</span>` : '')
    || (r.slack_intel  ? `<span style="font-size:11px;color:var(--subtext)">💬 ${{r.slack_intel.slice(0,120)}}${{r.slack_intel.length>120?'…':''}}</span>` : '');
  return `<div class="summary-cell">
    <div class="summary-short">${{shortDisplay}}</div>
    ${{full.length > 140 ? `<span class="expand-btn" onclick="toggleExpand(this,${{idx}})">▼ more</span>
    <div class="summary-full">${{full}}</div>` : ''}}
  </div>`;
}}

function displayName(n) {{
  return n.replace(/[ \t-]*-[ \t-]*(CSG|CO[0-9]+).*?-Project[ \t]*$/i, '').trim();
}}

function projectRow(r, idx) {{
  const hwFlag  = r.high_watch ? '<span class="hw-flag">⚑ HW</span>' : '';
  const sweFlag = r.swe_co ? '<span class="hw-flag" style="background:#E8F5E9;color:#2E7D32">SWE</span>' : '';
  const newFlag = (r.start_dt && r.start_dt > new Date().toISOString().slice(0,10)) ? '<span class="hw-flag" style="background:#FFF8E1;color:#F57F17">🆕 NEW</span>' : '';
  return `<tr data-idx="${{idx}}" data-grp="${{(r[groupBy]||'').toString().replace(/"/g,'&quot;')}}">
    <td class="col-status">${{statusBadge(r)}}</td>
    <td class="col-tier" style="text-align:center">${{_DB_CRED && r.region !== 'AMER CBS' ? `<select class="tier-edit" onchange="saveAssignment('${{r.pid}}','tier',this.value)" title="Edit tier"><option value="1"${{r.tier===1?' selected':''}}>T1</option><option value="2"${{r.tier===2?' selected':''}}>T2</option><option value="3"${{r.tier===3?' selected':''}}>T3</option></select>` : `<span class="tier">${{TIER_LABEL[r.tier]||r.tier}}</span>`}}</td>
    <td class="col-project">
      <div class="proj-name">${{hwFlag}}${{sweFlag}}${{newFlag}}<a href="${{r.url}}" target="_blank" class="proj-link">${{displayName(r.name)}}</a></div>
      <div class="proj-acct">${{r.acct}}</div>
      ${{(r.start_dt || r.end_dt) ? (() => {{
        const today = new Date().toISOString().slice(0,10);
        const endColor = !r.end_dt ? 'var(--subtext)'
          : r.end_dt < today ? 'var(--red)'
          : (new Date(r.end_dt) - new Date(today)) / 86400000 <= 45 ? 'var(--yellow)'
          : 'var(--subtext)';
        return `<div style="font-size:10px;color:var(--subtext);margin-top:2px"><span>Start: ${{r.start_dt||'?'}}</span> &nbsp;·&nbsp; <span>End: </span><span style="color:${{endColor}};font-weight:${{endColor!=='var(--subtext)'?'600':'400'}}">${{r.end_dt||'?'}}</span></div>`;
      }})() : ''}}
      ${{(r.stage || r.practice) ? `<div style="font-size:10px;color:#4B5563">${{[r.stage,r.practice].filter(Boolean).join(' · ')}}</div>` : ''}}
      ${{r.pulse_golive ? `<div style="font-size:10px;margin-top:2px"><span style="color:var(--subtext)">Go-Live: </span><span style="font-weight:600;color:var(--blue)">${{r.pulse_golive}}</span></div>` : ''}}
    </td>
    <td class="col-team">
      ${{teamHtml(r.team)}}
      ${{_DB_CRED ? `<div class="po-edit-wrap" style="margin-top:4px">
        <span class="po-edit-lbl" onclick="startPoEdit(this,'${{r.pid}}','${{(r.po||'').replace(/'/g,"\\'")}}')" title="Click to edit Portfolio Owner">PO: <span class="po-val">${{r.po||'—'}}</span> ✏️</span>
        <span class="po-edit-inp" style="display:none"><input type="text" value="${{r.po||''}}" onblur="commitPoEdit(this,'${{r.pid}}')" onkeydown="if(event.key==='Enter')this.blur();else if(event.key==='Escape')cancelPoEdit(this)"></span>
      </div>` : ''}}
    </td>
    <td class="col-resource">${{resourcingHtml(r)}}</td>
    <td class="col-fin">${{finHtml(r)}}</td>
    <td class="col-rules">${{rulesHtml(r.rules)}}</td>
    <td class="col-bl">${{baselinesHtml(r.baselines)}}</td>
    <td>${{summaryHtml(r, idx)}}</td>
  </tr>`;
}}

function grpLabel(key) {{
  if (groupBy === 'tier') return TIER_NAME[key] || key;
  if (groupBy === 'tier_acct') {{
    const parts = key.split('|');
    return parts[1] || '(Unassigned)';
  }}
  if (groupBy === 'acct_po') {{
    const parts = key.split('|');
    return parts[1] || '(Unassigned)';
  }}
  return key || '(Unassigned)';
}}

function grpSummary(rows) {{
  const cnt  = rows.length;
  const bk   = rows.reduce((s,r)=>s+(r.bookings||0),0);
  const bil  = rows.reduce((s,r)=>s+(r.billings||0),0);
  const rem  = rows.reduce((s,r)=>s+(r.actuals_rem||0),0);
  const reds = rows.filter(r=>r.status==='Red').length;
  const yels = rows.filter(r=>r.status==='Yellow').length;
  const wms  = rows.filter(r=>r.status==='Watermelon').length;
  const grns = rows.filter(r=>r.status==='Green').length;
  const nops = rows.filter(r=>r.status==='No Pulse').length;
  let pills = '';
  if (reds) pills += `<span class="grp-pill grp-pill-red">🔴 ${{reds}}</span>`;
  if (yels) pills += `<span class="grp-pill grp-pill-yellow">🟡 ${{yels}}</span>`;
  if (wms)  pills += `<span class="grp-pill grp-pill-wm">🍉 ${{wms}}</span>`;
  if (grns) pills += `<span class="grp-pill grp-pill-green">🟢 ${{grns}}</span>`;
  if (nops) pills += `<span class="grp-pill grp-pill-nopulse">⚫ ${{nops}}</span>`;
  const finParts = [`Bkd: ${{fmtMoney(bk)}}`];
  if (bil)  finParts.push(`Bil: ${{fmtMoney(bil)}}`);
  if (rem)  finParts.push(`Rem: ${{fmtMoney(rem)}}`);
  return `<span class="grp-meta">${{cnt}} proj &nbsp;·&nbsp; ${{finParts.join(' &nbsp;·&nbsp; ')}}</span>
          <span class="grp-pills">${{pills}}</span>`;
}}

function renderRows(data) {{
  const tbody = document.getElementById('table-body');
  const noRes = document.getElementById('no-results');
  if (!data.length) {{
    tbody.innerHTML = '';
    noRes.style.display = 'block';
    return;
  }}
  noRes.style.display = 'none';

  if (!groupBy) {{
    tbody.innerHTML = data.map((r,i) => projectRow(r,i)).join('');
  }} else if (groupBy === 'acct_po') {{
    // Two-level: Account Name → Portfolio Owner
    const seen = [], groupMap = {{}};
    data.forEach(r => {{
      const k = (r.acct||'') + '|' + (r.po||'');
      if (!groupMap[k]) {{ groupMap[k] = []; seen.push(k); }}
      groupMap[k].push(r);
    }});
    seen.sort((a,b) => {{
      const [aa,ap] = a.split('|'); const [ba,bp] = b.split('|');
      const ad = aa.localeCompare(ba);
      return ad !== 0 ? ad : ap.localeCompare(bp);
    }});
    let html = '';
    let lastAcct = null;
    seen.forEach(k => {{
      const acctName = k.split('|')[0];
      if (acctName !== lastAcct) {{
        lastAcct = acctName;
        const acctRows = data.filter(r => (r.acct||'') === acctName);
        const acctGid = 'acct-sect-' + acctName.replace(/[^a-z0-9]/gi,'_');
        html += `<tr class="group-row group-tier-header" onclick="toggleGroup('${{acctGid}}',this)">
          <td colspan="8">
            <span class="grp-toggle" id="arrow-${{acctGid}}">▼</span>
            <strong>${{acctName || '(No Account)'}}</strong>
            ${{grpSummary(acctRows)}}
          </td>
        </tr>`;
      }}
      const rows = groupMap[k];
      const gid = 'grp-' + k.replace(/[^a-z0-9]/gi,'_');
      const poName = k.split('|')[1] || '(Unassigned)';
      html += `<tr class="group-row group-acct-header" onclick="toggleGroup('${{gid}}',this)">
        <td colspan="8">
          <span class="grp-toggle" id="arrow-${{gid}}">▼</span>
          <strong>${{poName}}</strong>
          ${{grpSummary(rows)}}
        </td>
      </tr>`;
      rows.forEach((r,i) => {{
        html += projectRow(r, i).replace('<tr ', `<tr data-group="${{gid}}" data-group2="acct-sect-${{(r.acct||'').replace(/[^a-z0-9]/gi,'_')}}" `);
      }});
    }});
    tbody.innerHTML = html;
  }} else if (groupBy === 'tier_acct') {{
    // Build groups keyed by "tier|acct", sorted tier-first then acct alpha
    const seen = [], groupMap = {{}};
    data.forEach(r => {{
      const k = (r.tier||9) + '|' + (r.acct||'');
      if (!groupMap[k]) {{ groupMap[k] = []; seen.push(k); }}
      groupMap[k].push(r);
    }});
    seen.sort((a,b) => {{
      const [at,aa] = a.split('|'); const [bt,ba] = b.split('|');
      const td = (parseInt(at)||99) - (parseInt(bt)||99);
      return td !== 0 ? td : aa.localeCompare(ba);
    }});
    let html = '';
    let lastTier = null;
    seen.forEach(k => {{
      const tierNum = parseInt(k.split('|')[0]) || 9;
      if (tierNum !== lastTier) {{
        lastTier = tierNum;
        const tierRows = data.filter(r => (r.tier||9) === tierNum);
        const tierGid = 'tier-sect-' + tierNum;
        html += `<tr class="group-row group-tier-header" onclick="toggleGroup('${{tierGid}}',this)">
          <td colspan="8">
            <span class="grp-toggle" id="arrow-${{tierGid}}">▼</span>
            <strong>${{TIER_NAME[tierNum]||('Tier '+tierNum)}}</strong>
            ${{grpSummary(tierRows)}}
          </td>
        </tr>`;
      }}
      const rows = groupMap[k];
      const gid = 'grp-' + k.replace(/[^a-z0-9]/gi,'_');
      html += `<tr class="group-row group-acct-header" onclick="toggleGroup('${{gid}}',this)">
        <td colspan="8">
          <span class="grp-toggle" id="arrow-${{gid}}">▼</span>
          <strong>${{grpLabel(k)}}</strong>
          ${{grpSummary(rows)}}
        </td>
      </tr>`;
      rows.forEach((r,i) => {{
        html += projectRow(r, i).replace('<tr ', `<tr data-group="${{gid}}" data-group2="tier-sect-${{r.tier||9}}" `);
      }});
    }});
    tbody.innerHTML = html;
  }} else {{
    // Build ordered groups
    const seen = [], groupMap = {{}};
    data.forEach(r => {{
      const k = (r[groupBy]||'').toString();
      if (!groupMap[k]) {{ groupMap[k] = []; seen.push(k); }}
      groupMap[k].push(r);
    }});
    // Sort group keys
    seen.sort((a,b) => {{
      if (groupBy === 'tier') return (parseInt(a)||99) - (parseInt(b)||99);
      return a.localeCompare(b);
    }});
    let html = '';
    seen.forEach(k => {{
      const rows = groupMap[k];
      const gid = 'grp-' + k.replace(/[^a-z0-9]/gi,'_');
      html += `<tr class="group-row" onclick="toggleGroup('${{gid}}',this)">
        <td colspan="8">
          <span class="grp-toggle" id="arrow-${{gid}}">▼</span>
          <strong>${{grpLabel(k)}}</strong>
          ${{grpSummary(rows)}}
        </td>
      </tr>`;
      rows.forEach((r,i) => {{
        html += projectRow(r, i).replace('<tr ', `<tr data-group="${{gid}}" `);
      }});
    }});
    tbody.innerHTML = html;
  }}
  document.getElementById('cnt-all').textContent = `(${{data.length}})`;
}}

function collapseAll() {{
  document.querySelectorAll('tr.group-row').forEach(hdr => {{
    const gid = hdr.getAttribute('onclick')?.match(/toggleGroup\\('([^']+)'/)?.[1];
    if (!gid) return;
    const rows = document.querySelectorAll(`tr[data-group="${{gid}}"]`);
    const arrow = document.getElementById('arrow-' + gid);
    rows.forEach(r => r.classList.add('grp-hidden'));
    if (arrow) arrow.textContent = '▶';
  }});
}}

function expandAll() {{
  document.querySelectorAll('tr.group-row').forEach(hdr => {{
    const gid = hdr.getAttribute('onclick')?.match(/toggleGroup\\('([^']+)'/)?.[1];
    if (!gid) return;
    const rows = document.querySelectorAll(`tr[data-group="${{gid}}"]`);
    const arrow = document.getElementById('arrow-' + gid);
    rows.forEach(r => r.classList.remove('grp-hidden'));
    if (arrow) arrow.textContent = '▼';
  }});
}}

function toggleGroup(gid, headerRow) {{
  // Primary rows tagged with this group
  const rows = document.querySelectorAll(`tr[data-group="${{gid}}"]`);
  const arrow = document.getElementById('arrow-' + gid);
  const hidden = rows.length && rows[0].classList.contains('grp-hidden');
  rows.forEach(r => r.classList.toggle('grp-hidden', !hidden));
  if (arrow) arrow.textContent = hidden ? '▼' : '▶';
  // For tier-section headers: also toggle the account sub-header rows and their data rows
  if (gid.startsWith('tier-sect-')) {{
    document.querySelectorAll(`tr[data-group2="${{gid}}"]`).forEach(r => r.classList.toggle('grp-hidden', !hidden));
    // find account group-rows whose rows are children of this tier
    document.querySelectorAll('tr.group-row:not(.group-tier-header)').forEach(subHdr => {{
      const subGid = subHdr.getAttribute('onclick')?.match(/toggleGroup\\('([^']+)'/)?.[1];
      if (!subGid) return;
      const sample = document.querySelector(`tr[data-group="${{subGid}}"][data-group2="${{gid}}"]`);
      if (sample) subHdr.classList.toggle('grp-hidden', !hidden);
    }});
  }}
}}

function updateScorecard(data) {{
  const n = data.length;
  const bk  = data.reduce((s,r) => s + (r.bookings||0), 0);
  const bil = data.reduce((s,r) => s + (r.billings||0), 0);
  const far = data.reduce((s,r) => s + (r.far||0), 0);
  const ov  = data.reduce((s,r) => s + (r.overdue_inv||0), 0);
  const rr  = data.reduce((s,r) => s + (r.rr_revenue||0), 0);
  const statuses = data.map(r => r.status);
  const cnt = s => statuses.filter(v => v===s).length;
  // weighted margin (exclude SWE/ARI — no bid/close data for those)
  const mProj = data.filter(r => !r.swe_co && r.bid_margin_raw !== null && r.close_margin_raw !== null);
  let wBid = 0, wClose = 0;
  if (mProj.length) {{
    const totBk = mProj.reduce((s,r)=>s+(r.bookings||0),0);
    if (totBk > 0) {{
      wBid   = mProj.reduce((s,r)=>s+(r.bookings||0)*(r.bid_margin_raw||0),0) / totBk;
      wClose = mProj.reduce((s,r)=>s+(r.bookings||0)*(r.close_margin_raw||0),0) / totBk;
    }}
  }}
  document.getElementById('sc-total').textContent    = n;
  document.getElementById('sc-bookings').textContent = fmtMoney(bk);
  document.getElementById('sc-avg-bk').textContent   = fmtMoney(n > 0 ? bk / n : 0);
  const _bks = data.map(r => r.bookings || 0).sort((a,b) => a - b);
  const _mid = Math.floor(_bks.length / 2);
  const _med = _bks.length === 0 ? 0 : (_bks.length % 2 ? _bks[_mid] : (_bks[_mid-1] + _bks[_mid]) / 2);
  document.getElementById('sc-median-bk').textContent = fmtMoney(_med);
  document.getElementById('sc-billings').textContent = fmtMoney(bil);
  document.getElementById('sc-backlog').textContent  = fmtMoney(bk - bil);
  document.getElementById('sc-green').textContent       = cnt('Green');
  document.getElementById('sc-wm').textContent          = cnt('Watermelon');
  document.getElementById('sc-green-total').textContent = cnt('Green') + cnt('Watermelon');
  document.getElementById('sc-yellow').textContent   = cnt('Yellow');
  document.getElementById('sc-red').textContent      = cnt('Red');
  document.getElementById('sc-nopulse').textContent  = cnt('No Pulse');
  document.getElementById('sc-far').textContent      = fmtMoney(far);
  document.getElementById('sc-overdue').textContent  = fmtMoney(ov);
  document.getElementById('sc-rr').textContent       = fmtMoney(rr);
  const bidEl = document.getElementById('sc-bid');
  bidEl.textContent = fmtPct(wBid);
  bidEl.style.color = wBid > 13 ? 'var(--green)' : wBid >= 5 ? 'var(--yellow)' : 'var(--red)';
  const wDel = wClose - wBid;
  const delEl = document.getElementById('sc-delivered');
  delEl.textContent = (wDel >= 0 ? '+' : '') + fmtPct(wDel);
  delEl.closest('.stat').className = 'stat ' + (wDel < -5 ? 'red' : wDel < 0 ? 'yellow' : 'green');
  const closeEl = document.getElementById('sc-close');
  closeEl.textContent = fmtPct(wClose);
  closeEl.style.color = wClose < 0 ? 'var(--red)' : '';
  document.getElementById('sc-delta').textContent    = (wClose - wBid >= 0 ? '+' : '') + fmtPct(wClose - wBid);
  document.getElementById('header-summary').innerHTML =
    n + ' projects &nbsp;|&nbsp; ' + fmtMoney(bk) + ' bookings';
}}

function setRegion(region, btn) {{
  filterRegion = region;
  document.querySelectorAll('.rtab').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');
  applyFilters();
}}

function applyFilters() {{
  const q = (document.getElementById('search').value || '').toLowerCase();
  let data = RAW.filter(r => {{
    if (filterRegion && r.region !== filterRegion) return false;
    if (filterStatuses.size > 0 && !filterStatuses.has(r.status)) return false;
    if (filterTiers.size > 0 && !filterTiers.has(String(r.tier))) return false;
    if (filterPOs.size > 0   && !filterPOs.has(r.po))             return false;
    if (filterRules.size > 0) {{
      const rowRules = (r.rules || '').split(', ').filter(Boolean);
      if (!rowRules.some(rc => filterRules.has(rc))) return false;
    }}
    if (filterHW  && !r.high_watch) return false;
    if (filterSWE && !r.swe_co)     return false;
    if (q && !([r.name, r.acct, r.team, r.po, r.rules, r.summary, r.baselines].join(' ').toLowerCase().includes(q))) return false;
    return true;
  }});

  data = data.slice().sort((a, b) => {{
    let av = a[sortCol], bv = b[sortCol];
    if (sortCol === 'status') {{ av = STATUS_ORDER[av]??99; bv = STATUS_ORDER[bv]??99; }}
    if (typeof av === 'string') av = av.toLowerCase();
    if (typeof bv === 'string') bv = bv.toLowerCase();
    if (av < bv) return -sortDir;
    if (av > bv) return  sortDir;
    return 0;
  }});

  groupBy = document.getElementById('group-by').value;
  document.getElementById('grp-collapse-btns').style.display = groupBy ? 'inline-flex' : 'none';
  updateScorecard(data);
  renderRows(data);
}}

function togglePill(el) {{
  const filter = el.dataset.filter;
  if (filter === 'hw') {{
    el.classList.toggle('active');
    filterHW = el.classList.contains('active');
  }} else if (filter === 'swe') {{
    el.classList.toggle('active');
    filterSWE = el.classList.contains('active');
  }}
  applyFilters();
}}

// ── Multi-select combobox helpers ──────────────────────────────────────────────
const MS_CONFIG = {{
  status: {{ set: () => filterStatuses, btnId: 'ms-status-btn', ddId: 'ms-status-dd', label: 'All Statuses', plural: (n) => n + ' Status' + (n>1?'es':'') }},
  tier:   {{ set: () => filterTiers,    btnId: 'ms-tier-btn',   ddId: 'ms-tier-dd',   label: 'All Tiers',    plural: (n) => n + ' Tier' + (n>1?'s':'') }},
  po:     {{ set: () => filterPOs,      btnId: 'ms-po-btn',     ddId: 'ms-po-dd',     label: 'All Owners',   plural: (n) => n + ' Owner' + (n>1?'s':'') }},
  rule:   {{ set: () => filterRules,    btnId: 'ms-rule-btn',   ddId: 'ms-rule-dd',   label: 'All Rules',    plural: (n) => n + ' Rule' + (n>1?'s':'') }},
}};

function toggleMs(key) {{
  const cfg = MS_CONFIG[key];
  const dd = document.getElementById(cfg.ddId);
  // close all others
  Object.keys(MS_CONFIG).forEach(k => {{
    if (k !== key) document.getElementById(MS_CONFIG[k].ddId).classList.remove('open');
  }});
  dd.classList.toggle('open');
}}

function msChange(key) {{
  const cfg = MS_CONFIG[key];
  const s = cfg.set();
  s.clear();
  document.querySelectorAll('#' + cfg.ddId + ' input[type=checkbox]:checked').forEach(cb => s.add(cb.value));
  const btn = document.getElementById(cfg.btnId);
  btn.textContent = s.size === 0 ? cfg.label : cfg.plural(s.size);
  btn.classList.toggle('ms-active', s.size > 0);
  applyFilters();
}}

function msClear(key) {{
  const cfg = MS_CONFIG[key];
  cfg.set().clear();
  document.querySelectorAll('#' + cfg.ddId + ' input[type=checkbox]').forEach(cb => cb.checked = false);
  const btn = document.getElementById(cfg.btnId);
  btn.textContent = cfg.label;
  btn.classList.remove('ms-active');
  applyFilters();
}}

// close dropdowns when clicking outside
document.addEventListener('click', e => {{
  if (!e.target.closest('.ms-wrap')) {{
    Object.values(MS_CONFIG).forEach(cfg => document.getElementById(cfg.ddId).classList.remove('open'));
  }}
}});

function toggleExpand(btn, idx) {{
  const row = btn.closest('tr');
  row.classList.toggle('expanded');
  btn.textContent = row.classList.contains('expanded') ? '▲ less' : '▼ more';
}}

document.querySelectorAll('thead th[data-col]').forEach(th => {{
  th.addEventListener('click', () => {{
    const col = th.dataset.col;
    if (sortCol === col) {{ sortDir *= -1; }}
    else {{ sortCol = col; sortDir = 1; }}
    document.querySelectorAll('thead th').forEach(t => t.classList.remove('sorted-asc','sorted-desc'));
    th.classList.add(sortDir === 1 ? 'sorted-asc' : 'sorted-desc');
    applyFilters();
  }});
}});

document.querySelector('[data-col="status"]').classList.add('sorted-asc');
loadData();
</script>
</body>
</html>"""

    import json as _json_html
    json_path = os.path.join(OUTPUT_DIR, f"acc_{REGION_SLUG.lower()}_data.json")
    with open(json_path, 'w', encoding='utf-8') as fh:
        _json_html.dump({'generated': REPORT_DATE, 'region': REGION_LABEL, 'rows': rows_data}, fh, ensure_ascii=False)
    print(f"✅  JSON data saved: {json_path}")

    path = base_path + '.html'
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write(html)
    print(f"✅  HTML saved: {path}")

    # Undated copy for Heroku / static hosting (always overwrites)
    latest_path = os.path.join(OUTPUT_DIR, f"{REGION_SLUG}_Audit_latest.html")
    with open(latest_path, 'w', encoding='utf-8') as fh:
        fh.write(html)
    print(f"✅  HTML (latest) saved: {latest_path}")

def write_json():
    import json as _json
    def fmt_m(v):
        if v is None: return ''
        return f'${v/1e6:.2f}M' if abs(v) >= 1_000_000 else f'${v:,.0f}'
    def fmt_pct(v):
        return f'{v:.1f}%' if v is not None else ''
    def status_label(r):
        if r['stage'] == 'On Hold': return 'On Hold'
        if r['is_watermelon']: return 'Watermelon'
        h = r['health'].lower()
        if h == 'red':    return 'Red'
        if h == 'yellow': return 'Yellow'
        if h == 'green':  return 'Green'
        return 'No Pulse'
    rows_data = []
    for r in results:
        dd = (r['close_margin'] - r['bid_margin']) if (r['bid_margin'] is not None and r['close_margin'] is not None) else None
        viol_codes = ', '.join(v[0] for v in r['violations'])
        bl = ' | '.join(f"{lbl}={v}" for lbl, v in baselines_list(r))
        team = []
        if r['pm']:        team.append(f"PM: {r['pm']}")
        if r.get('pm2'):   team.append(f"PM2: {r['pm2']}")
        if r.get('opp_owner'): team.append(f"AP: {r['opp_owner']}")
        if r.get('exec_sponsor'): team.append(f"ES: {r['exec_sponsor']}")
        if r.get('acct_owner'): team.append(f"AO: {r['acct_owner']}")
        if r.get('owner') and r.get('owner') != 'Unassigned': team.append(f"PO: {r['owner']}")
        po = r.get('owner','') or ''
        rows_data.append({
            'name':          r['name'],
            'acct':          r['acct'],
            'status':        status_label(r),
            'tier':          r['tier'],
            'team':          ' · '.join(team),
            'po':            po,
            'pid':           r['pid'],
            'region':        REGION_LABEL,
            'url':           f"https://org62.lightning.force.com/lightning/r/pse__Proj__c/{r['pid']}/view",
            'rev_treat':     r.get('rev_treat') or '',
            'billing_type':  r.get('billing_type') or '',
            'fmt_pipe':      fmt_m(r.get('open_pipe')) if r.get('open_pipe') else '',
            'health_risk':   r.get('health_risk_score'),
            'data_quality':  r.get('data_quality_score'),
            'bookings':      r['bookings'] or 0,
            'billings':      r['billings'] or 0,
            'far':           r['far'] or 0,
            'overdue_inv':   r['overdue_inv'] or 0,
            'rr_revenue':    r['rr_revenue'] or 0,
            'bid_margin_raw':   r['bid_margin'],
            'close_margin_raw': r['close_margin'],
            'rules':         viol_codes,
            'baselines':     bl,
            'summary':       ' '.join((r.get('overall_summary') or '').split())[:500],
            'leadership':    ' '.join((r.get('leadership_notes') or '').split())[:400],
            'swe_co':        r['swe_co'],
            'high_watch':    bool(r.get('high_watch')),
            'has_pulse':     r.get('has_pulse', False),
            'pulse_trend':   r.get('trend') or '',
            'pulse_updated': (r.get('last_updated') or '')[:10],
            'pulse_scope':   r.get('scope_s') or '',
            'pulse_sched':   r.get('sched_s') or '',
            'pulse_budget':  r.get('budget_s') or '',
            'pulse_resource':r.get('resource_s') or '',
            'pulse_customer':r.get('customer_s') or '',
            'pulse_action':  ' '.join((r.get('action_needed') or '').split())[:300],
            'pulse_steerco': r['steerco_date'].isoformat() if r.get('steerco_date') else '',
            'pulse_golive':  r['next_golive'].isoformat() if r.get('next_golive') else '',
            'slack_intel':   r.get('slack_intel') or '',
            'start_dt':      r['start_dt'].isoformat() if r.get('start_dt') else '',
            'end_dt':        r['end_dt'].isoformat() if r.get('end_dt') else '',
            'stage':         r.get('stage') or '',
            'practice':      r.get('practice') or '',
            'unsch_backlog': r.get('unsch_backlog') or 0,
            'actuals_rem':   r.get('actuals_rem') or 0,
            'eva_amt':       r.get('eva_amt'),
            'eva_pct':       r.get('eva_pct'),
            'csat_score':    r.get('csat_score'),
            'fmt_unsch':     fmt_m(r.get('unsch_backlog')) if r.get('unsch_backlog') else '',
            'fmt_actuals':   fmt_m(r.get('actuals_rem')) if r.get('actuals_rem') else '',
            'fmt_eva_amt':   (f"+${r['eva_amt']:,.0f}" if r.get('eva_amt') is not None and r['eva_amt'] >= 0 else f"-${abs(r['eva_amt']):,.0f}" if r.get('eva_amt') is not None else ''),
            'fmt_eva_pct':   (f"{r['eva_pct']:+.1f}%" if r.get('eva_pct') is not None else ''),
            'far_reason':    r.get('far_reason') or '',
            'far_subreason': r.get('far_subreason') or '',
            'fmt_bookings':  fmt_m(r['bookings']),
            'fmt_billings':  fmt_m(r['billings']),
            'fmt_bid':       fmt_pct(r['bid_margin']),
            'fmt_close':     fmt_pct(r['close_margin']),
            'fmt_delivered': (f'{dd:+.1f}%' if dd is not None else ''),
            'fmt_far':       fmt_m(r['far']),
            'fmt_overdue':   fmt_m(r['overdue_inv']) if r['overdue_inv'] else '',
            'fmt_rr':        fmt_m(r['rr_revenue']) if r['rr_revenue'] else '',
            'gdc_total':     r.get('gdc_total') or 0,
            'gdc_india':     r.get('gdc_india') or 0,
            'gdc_pct':       round(r['gdc_pct'] * 100, 1) if r.get('gdc_pct') is not None else None,
            'gdc_resources': r.get('gdc_resources') or [],
        })
    json_path = os.path.join(OUTPUT_DIR, f"acc_{REGION_SLUG.lower()}_data.json")
    with open(json_path, 'w', encoding='utf-8') as fh:
        _json.dump({'generated': REPORT_DATE, 'region': REGION_LABEL, 'rows': rows_data}, fh, ensure_ascii=False)
    print(f"✅  JSON saved: {json_path}")

# ── Dispatch ──────────────────────────────────────────────────────────────────
if 'json' in OUTPUT_FORMATS: write_json()
if 'txt'  in OUTPUT_FORMATS: write_txt()
if 'docx' in OUTPUT_FORMATS: write_docx()
if 'pptx' in OUTPUT_FORMATS: write_pptx()
if 'html' in OUTPUT_FORMATS: write_html()
# end region loop

# ── Combined HTML (multi-region only) ─────────────────────────────────────────
if 'html' in OUTPUT_FORMATS and len(SELECTED_REGION_KEYS) > 1 and _all_results:
    results    = _all_results
    REGION_LABEL = 'ACC'
    REGION_SLUG  = 'ACC'
    base_path    = f"{OUTPUT_DIR}/ACC_Audit_{FILE_STAMP}"
    write_html()

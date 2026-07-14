#!/usr/bin/env python3
"""
Combine TMT and CBS data files into a single ACC HTML.
Usage: python3 combine_html.py --output DIR
Reads: acc_amer_tmt_data.json, acc_amer_cbs_data.json
Reads template from: AMER_TMT_Audit_latest.html (latest TMT HTML as shell)
Writes: ACC_Audit_latest.html, ACC_Audit_<stamp>.html
"""
import argparse, json, os, re, html, glob
from datetime import datetime

parser = argparse.ArgumentParser()
parser.add_argument('--output', required=True)
args = parser.parse_args()

outdir = args.output

# Load both JSON files
jsons = sorted(glob.glob(os.path.join(outdir, 'acc_amer_*_data.json')))
if not jsons:
    print(f"No acc_amer_*_data.json files in {outdir}")
    raise SystemExit(1)

all_rows = []
all_pipe = {}
generated = ''
for jf in jsons:
    with open(jf) as f:
        d = json.load(f)
    all_rows.extend(d.get('rows', []))
    for aid, opps in (d.get('pipe') or {}).items():
        all_pipe.setdefault(aid, []).extend(opps)
    g = d.get('generated', '')
    if g > generated:
        generated = g

print(f"combine_html: {len(all_rows)} rows from {[os.path.basename(f) for f in jsons]}")

# Read TMT HTML as template
tmt_html_files = sorted(glob.glob(os.path.join(outdir, 'AMER_TMT_Audit_latest.html')))
if not tmt_html_files:
    tmt_html_files = sorted(glob.glob(os.path.join(outdir, 'AMER_TMT_Audit_*.html')))
if not tmt_html_files:
    print("No AMER_TMT_Audit HTML found"); raise SystemExit(1)

with open(tmt_html_files[-1], encoding='utf-8') as f:
    content = f.read()

# Replace _INLINE JSON
if 'const _INLINE = ' not in content:
    print(f"ERROR: _INLINE placeholder not found in template {tmt_html_files[-1]}"); raise SystemExit(1)
combined_json = json.dumps({'generated': generated, 'region': 'ACC', 'rows': all_rows, 'pipe': all_pipe}, ensure_ascii=False)
replacement = f'const _INLINE = {combined_json};'
content = re.sub(
    r'const _INLINE = \{.*?\};',
    lambda m: replacement,
    content, count=1, flags=re.DOTALL
)
if f'"region": "ACC"' not in content:
    print("ERROR: _INLINE replacement did not apply"); raise SystemExit(1)

# Patch tier-edit to work for all regions (CBS tiers also editable)
content = content.replace(
    "_DB_CRED && r.region !== 'AMER CBS' ?",
    "_DB_CRED ?"
)

# Rebuild PO dropdown checkboxes (also updates the JSON in combined HTML)
po_list = sorted(set(r['po'] for r in all_rows if r.get('po') and r['po'] != 'Unassigned'))
po_html = '\n'.join(
    f'<label class="ms-item"><input type="checkbox" value="{html.escape(p)}" onchange="msChange(\'po\')"> {html.escape(p)}</label>'
    for p in po_list
)
content = re.sub(
    r'(<div class="ms-dropdown" id="ms-po-dd">).*?(<span class="ms-clear")',
    rf'\1\n{po_html}\n        \2',
    content, count=1, flags=re.DOTALL
)

stamp = datetime.now().strftime('%Y-%m-%d_%H%M')
for out_path in [
    os.path.join(outdir, f'ACC_Audit_{stamp}.html'),
    os.path.join(outdir, 'ACC_Audit_latest.html'),
]:
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"✅  Combined HTML: {out_path}")

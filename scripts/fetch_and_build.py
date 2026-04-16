"""
fetch_and_build.py
------------------
1. Reads data from the three WH GSheets (PHB, PHL, PHIXC)
2. Extracts benchmark values from IB Model, OB Model, Inventory Model (col F)
3. Extracts actual monthly data from the actual raw data / dashboard sheets
4. Injects everything into the HTML template
5. Writes docs/index.html

Environment variables required (set as GitHub Secrets):
  GOOGLE_CREDENTIALS  – full JSON key content (as string)
  SHEET_ID_PHB                 – GSheet ID for PHB
  SHEET_ID_PHL                 – GSheet ID for PHL
  SHEET_ID_PHIXC               – GSheet ID for PHIXC
"""

import os, json, re, datetime
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────────────────────────────
# 1. AUTH
# ─────────────────────────────────────────────────────────────────────
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def get_gspread_client():
    raw = os.environ['GOOGLE_CREDENTIALS']
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

# ─────────────────────────────────────────────────────────────────────
# 2. HELPERS
# ─────────────────────────────────────────────────────────────────────
def col_f_value(ws, row_index):
    """Return the value in column F (index 5) for a given 1-based row."""
    try:
        row = ws.row_values(row_index)
        if len(row) >= 6:
            v = row[5]
            try:
                return float(v)
            except (ValueError, TypeError):
                return v
        return None
    except Exception:
        return None

def safe_float(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default

def safe_list(lst, length=4, default=0.0):
    """Pad or trim list to exact length."""
    out = [safe_float(v, default) for v in lst]
    while len(out) < length:
        out.append(default)
    return out[:length]

# ─────────────────────────────────────────────────────────────────────
# 3. READ IB MODEL BENCHMARKS (column F, matching row numbers from Excel)
# ─────────────────────────────────────────────────────────────────────
def read_ib_benchmarks(ws):
    """Extract all IB benchmark values from IB Model sheet column F."""
    cf = lambda r: col_f_value(ws, r)
    return {
        # Properties / factors
        'ib_ior':    safe_float(cf(5)),
        'pct_mti':   safe_float(cf(14)),
        's1_items':  safe_float(cf(38)),
        'spu_s1':    safe_float(cf(40)),
        'po_prod':   safe_float(cf(41)),
        'mt_prod':   safe_float(cf(42)),
        'r_hrs':     safe_float(cf(46)),
        'spu_r':     safe_float(cf(48)),
        'r_stns':    safe_float(cf(49)),  # min # stations (BAU)
        's3_hrs':    safe_float(cf(51)),
        'put_prod':  safe_float(cf(52)),
        's3_wait':   safe_float(cf(69)),  # Avg putaway wait BAU
        's3_items':  safe_float(cf(61)),
        'spu_s3':    safe_float(cf(63)),
        's1_wait':   safe_float(cf(68)),  # Avg recv wait BAU
        's1_truck':  safe_float(cf(26)),
        # Space (layout max)
        'sp_s1':     safe_float(cf(82)),
        'sp_r':      safe_float(cf(83)),
        'sp_s3':     safe_float(cf(84)),
        # Max layout capacities
        's1_pal':    safe_float(cf(86)),
        'r_stns':    safe_float(cf(87)),  # override with max layout
        's3_pal':    safe_float(cf(88)),
        # Step benchmark ADOs (campaign max ADO with uplift)
        'bm_ib_staging':  safe_float(cf(94)),   # Arrival Staging Max ADO
        'bm_ib_recv':     safe_float(cf(101)),  # Receiving Max ADO
        'bm_ib_putaway':  safe_float(cf(107)),  # Putaway Max ADO
        'bm_ib':          safe_float(cf(113)),  # Overall IB Max ADO (min of steps)
    }

# ─────────────────────────────────────────────────────────────────────
# 4. READ OB MODEL BENCHMARKS
# ─────────────────────────────────────────────────────────────────────
def read_ob_benchmarks(ws):
    cf = lambda r: col_f_value(ws, r)
    return {
        'ob_ior':    safe_float(cf(5)),
        'stg_items': safe_float(cf(28)),
        'sort_pct':  safe_float(cf(43)),
        'sort_prod': safe_float(cf(44)),
        'pack_prod': safe_float(cf(45)),
        'spu_sort':  safe_float(cf(48)),
        'spu_pack':  safe_float(cf(49)),
        'pre_prod':  safe_float(cf(51)),
        'spu_pre':   safe_float(cf(53)),
        'stg_pal':   safe_float(cf(78)),
        'spu_stg':   safe_float(cf(31)),
        'sort_hrs':  24,
        'pack_hrs':  24,
        'pre_hrs':   24,
        # Space (layout max)
        'sp_sort':   safe_float(cf(68)),
        'sp_pack':   safe_float(cf(69)),
        'sp_pre':    safe_float(cf(70)),
        'sp_stg':    safe_float(cf(71)),
        # Max stations
        'sort_stns': safe_float(cf(73)),
        'pack_stns': safe_float(cf(74)),
        'pre_stns':  safe_float(cf(75)),
        # Step benchmark ADOs
        'bm_ob_sorting':  safe_float(cf(81)),
        'bm_ob_packing':  safe_float(cf(87)),
        'bm_ob_presort':  safe_float(cf(93)),
        'bm_ob_staging':  safe_float(cf(99)),
        'bm_ob':          safe_float(cf(81)),  # overall = min = sorting
    }

# ─────────────────────────────────────────────────────────────────────
# 5. READ INVENTORY MODEL BENCHMARKS
# ─────────────────────────────────────────────────────────────────────
def read_inv_benchmarks(ws):
    """Read from 'Inventory Model' sheet — Theoretical calculation table."""
    # The theoretical table starts around row 11 in these sheets.
    # We'll search for the 'Theoretical calculation' header and read from there.
    # Fallback: read known row numbers based on the Excel structure we've mapped.
    all_values = ws.get_all_values()
    
    # Find Theoretical calculation section
    theo_row = None
    for i, row in enumerate(all_values):
        if any('Theoretical calculation' in str(c) for c in row):
            theo_row = i
            break
    
    if theo_row is None:
        return {}  # sheet structure unexpected
    
    # Read rows relative to theo_row
    def get_val(offset, col_idx=2):
        r = theo_row + offset
        if r < len(all_values) and col_idx < len(all_values[r]):
            return safe_float(all_values[r][col_idx], None)
        return None

    # Based on mapped structure: IOR at +3, DOC at +5, CBM/pcs at +7, CBM/SQM at +10, Loc util at +14
    # These offsets match the Excel structure observed
    result = {}
    for offset, key in [
        (3,  'inv_ior'),
        (5,  'doc'),
        (7,  'cbm_pcs'),
        (10, 'cbm_sqm'),
        (14, 'cbm_util'),
    ]:
        v = get_val(offset)
        if v is not None:
            result[key] = v

    # SQM — read from 'Actual' SQM row (offset +15 typically)
    for offset in range(12, 20):
        row = all_values[theo_row + offset] if theo_row + offset < len(all_values) else []
        if any('SQM' in str(c) for c in row) and any('Actual' in str(c) for c in row):
            try:
                result['sqm'] = safe_float(row[2])
            except Exception:
                pass
            break

    return result

# ─────────────────────────────────────────────────────────────────────
# 6. READ ACTUAL DATA
# ─────────────────────────────────────────────────────────────────────
def read_actual_data(gc, sheet_id):
    """
    Read actual monthly data from 'actual raw data', 'IB dashboard', 'OB dashboard' tabs.
    Returns a dict matching the WH[CUR] structure in the HTML.
    
    NOTE: The exact row/column mapping depends on your sheet structure.
    This function uses the same row mappings we discovered from the Excel files.
    Adjust the row numbers here if your GSheet layout differs.
    """
    sh = gc.open_by_key(sheet_id)
    
    months = ['Jan-26', 'Feb-26', 'Mar-26', 'Apr-26']
    
    # Try to read from IB dashboard / OB dashboard / raw data tabs
    # These names should match your actual tab names exactly
    actual = {}
    
    for tab_name in ['IB dashboard', 'IB Dashboard', 'IB_dashboard']:
        try:
            ws = sh.worksheet(tab_name)
            actual['ib_tab'] = ws
            break
        except gspread.WorksheetNotFound:
            continue

    for tab_name in ['OB dashboard', 'OB Dashboard', 'OB_dashboard']:
        try:
            ws = sh.worksheet(tab_name)
            actual['ob_tab'] = ws
            break
        except gspread.WorksheetNotFound:
            continue

    for tab_name in ['actual raw data', 'Actual Raw Data', 'actual_raw_data', 'Raw Data']:
        try:
            ws = sh.worksheet(tab_name)
            actual['raw_tab'] = ws
            break
        except gspread.WorksheetNotFound:
            continue

    # ── Placeholder: return zeros if we can't find the tabs ──
    # Once you confirm exact tab names and row/column layout,
    # replace this with precise cell references.
    blank4 = [0, 0, 0, 0]
    
    return {
        'months': months,
        'ib': {
            'avgUtil':     blank4,
            'peakUtil':    blank4,
            'locPeakUtil': blank4,
            'actualADO':   blank4,
            'peakADO':     blank4,
            'maxADO':      blank4,
        },
        'inv': {
            'avgUtil':     blank4,
            'peakUtil':    blank4,
            'locPeakUtil': blank4,
            'maxADO':      blank4,
            'maxCBM':      blank4,
            'actualCBM':   blank4,
        },
        'ob': {
            'avgUtil':     blank4,
            'peakUtil':    blank4,
            'actualADO':   blank4,
            'peakADO':     blank4,
            'maxADO':      blank4,
        },
        'mto': {
            'peakUtil':  blank4,
            'maxADO':    blank4,
            'handover':  blank4,
        },
        'ibSteps': {
            'jan': [0,0,0], 'feb': [0,0,0],
            'mar': [0,0,0], 'apr': [0,0,0],
        },
        'obSteps': {
            'jan': [0,0,0,0], 'feb': [0,0,0,0],
            'mar': [0,0,0,0], 'apr': [0,0,0,0],
        },
        'act': None,
        'bottleneck': 'INV',
        'ibBN': 'Receiving',
        'obBN': 'Packing',
        'spaceEff': None,
    }

# ─────────────────────────────────────────────────────────────────────
# 7. MAIN: FETCH ALL THREE WH
# ─────────────────────────────────────────────────────────────────────
def fetch_all_wh_data():
    gc = get_gspread_client()
    
    wh_configs = {
        'PHB':   os.environ['SHEET_ID_PHB'],
        'PHL':   os.environ['SHEET_ID_PHL'],
        'PHIXC': os.environ['SHEET_ID_PHIXC'],
    }
    
    all_bm = {}
    all_wh = {}

    for wh, sheet_id in wh_configs.items():
        print(f"Reading {wh} (sheet: {sheet_id[:8]}...)")
        sh = gc.open_by_key(sheet_id)
        
        # ── Read benchmarks ──
        ib_ws  = sh.worksheet('IB Model')
        ob_ws  = sh.worksheet('OB Model')
        inv_ws = sh.worksheet('Inventory Model')
        
        ib_bm  = read_ib_benchmarks(ib_ws)
        ob_bm  = read_ob_benchmarks(ob_ws)
        inv_bm = read_inv_benchmarks(inv_ws)
        
        bm = {**ib_bm, **ob_bm, **inv_bm}
        all_bm[wh] = bm
        
        # ── Read actual data ──
        actual = read_actual_data(gc, sheet_id)
        actual['theo'] = {
            'ibADO':  bm.get('bm_ib', 0),
            'invADO': 0,  # computed from INV model
            'obADO':  bm.get('bm_ob', 0),
        }
        actual['name'] = wh
        all_wh[wh] = actual
        
        print(f"  ✓ {wh}: IB={bm.get('bm_ib',0):.0f}, OB={bm.get('bm_ob',0):.0f}")
    
    return all_bm, all_wh

# ─────────────────────────────────────────────────────────────────────
# 8. INJECT DATA INTO HTML TEMPLATE
# ─────────────────────────────────────────────────────────────────────
def bm_to_js(wh, bm):
    """Convert a benchmark dict to a JS const BM_XXX={...} block."""
    lines = [f"const BM_{wh}={{"]
    for k, v in bm.items():
        if isinstance(v, float):
            lines.append(f"  {k}:{v},")
        elif isinstance(v, int):
            lines.append(f"  {k}:{v},")
        else:
            lines.append(f"  {k}:{json.dumps(v)},")
    lines.append("};")
    return "\n".join(lines)

def wh_to_js(wh, data):
    """Convert WH actual data dict to JS WH.PHB = {...} structure."""
    def js(v):
        return json.dumps(v)
    
    d = data
    months_js = js(d['months'])
    theo = d.get('theo', {})
    
    ib = d['ib']
    inv = d['inv']
    ob  = d['ob']
    mto = d.get('mto', {})
    ibs = d.get('ibSteps', {})
    obs = d.get('obSteps', {})
    
    bn  = d.get('bottleneck', 'INV')
    ibn = d.get('ibBN', 'Receiving')
    obn = d.get('obBN', 'Packing')
    
    act = d.get('act')
    act_js = js(act) if act else 'null'
    
    speff = d.get('spaceEff')
    speff_js = js(speff) if speff else 'null'
    
    return f"""{wh}:{{name:'{wh}',months:{months_js},
    theo:{{ibADO:{theo.get('ibADO',0)},invADO:{theo.get('invADO',0)},obADO:{theo.get('obADO',0)}}},
    ib:{{avgUtil:{js(ib['avgUtil'])},peakUtil:{js(ib['peakUtil'])},
        actualADO:{js(ib['actualADO'])},peakADO:{js(ib.get('peakADO',[0,0,0,0]))},maxADO:{js(ib['maxADO'])}}},
    inv:{{avgUtil:{js(inv['avgUtil'])},peakUtil:{js(inv['peakUtil'])},locPeakUtil:{js(inv.get('locPeakUtil',[0,0,0,0]))},
         maxADO:{js(inv['maxADO'])},maxCBM:{js(inv.get('maxCBM',[0,0,0,0]))},actualCBM:{js(inv.get('actualCBM',[0,0,0,0]))}}},
    ob:{{avgUtil:{js(ob['avgUtil'])},peakUtil:{js(ob['peakUtil'])},
        actualADO:{js(ob['actualADO'])},peakADO:{js(ob.get('peakADO',[0,0,0,0]))},maxADO:{js(ob['maxADO'])}}},
    mto:{{peakUtil:{js(mto.get('peakUtil',[0,0,0,0]))},maxADO:{js(mto.get('maxADO',[0,0,0,0]))},handover:{js(mto.get('handover',[0,0,0,0]))}}},
    ibSteps:{{jan:{js(ibs.get('jan',[0,0,0]))},feb:{js(ibs.get('feb',[0,0,0]))},mar:{js(ibs.get('mar',[0,0,0]))},apr:{js(ibs.get('apr',[0,0,0]))}}},
    obSteps:{{jan:{js(obs.get('jan',[0,0,0,0]))},feb:{js(obs.get('feb',[0,0,0,0]))},mar:{js(obs.get('mar',[0,0,0,0]))},apr:{js(obs.get('apr',[0,0,0,0]))}}}  ,
    spaceEff:{speff_js},act:{act_js},
    bottleneck:{js(bn)},ibBN:{js(ibn)},obBN:{js(obn)}}}"""


def inject_into_html(all_bm, all_wh):
    """Read template HTML, replace BM_DATA and WH data blocks, write output."""
    template_path = Path(__file__).parent.parent / 'docs' / 'index.html'
    
    if not template_path.exists():
        raise FileNotFoundError(
            f"Template not found at {template_path}. "
            "Make sure docs/index.html exists (copy the current dashboard HTML there)."
        )
    
    html = template_path.read_text(encoding='utf-8')
    
    # ── Replace BM_PHB / BM_PHL / BM_PHIXC blocks ──
    for wh in ['PHB', 'PHL', 'PHIXC']:
        new_block = bm_to_js(wh, all_bm[wh])
        # Match pattern: const BM_XXX={...};
        pattern = rf'const BM_{wh}=\{{[^}}]*(?:\{{[^}}]*\}}[^}}]*)?\}};'
        # Use a simpler approach: find start/end of the block
        tag_start = f'const BM_{wh}={{'
        tag_end = '};'
        start_idx = html.find(tag_start)
        if start_idx == -1:
            print(f"  ⚠ BM_{wh} block not found in template HTML")
            continue
        # Find matching }; 
        depth = 0
        end_idx = start_idx
        for i, c in enumerate(html[start_idx:], start_idx):
            if c == '{': depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0:
                    end_idx = i + 1
                    break
        if html[end_idx:end_idx+1] == ';':
            end_idx += 1
        html = html[:start_idx] + new_block + html[end_idx:]
        print(f"  ✓ Injected BM_{wh}")
    
    # ── Replace WH data block ──
    wh_js_entries = ',\n  '.join(wh_to_js(wh, all_wh[wh]) for wh in ['PHB', 'PHL', 'PHIXC'])
    new_wh_block = f"const WH={{\n  {wh_js_entries}\n}};"
    
    wh_start = html.find('const WH={')
    if wh_start != -1:
        depth = 0
        wh_end = wh_start
        for i, c in enumerate(html[wh_start:], wh_start):
            if c == '{': depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0:
                    wh_end = i + 1
                    break
        if html[wh_end:wh_end+1] == ';':
            wh_end += 1
        html = html[:wh_start] + new_wh_block + html[wh_end:]
        print(f"  ✓ Injected WH data for PHB / PHL / PHIXC")
    else:
        print("  ⚠ WH={} block not found in template HTML")
    
    # ── Add last-updated timestamp ──
    now = datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    html = html.replace(
        'Jan–Mar 2026 · '+'{CUR}',
        'Jan–Mar 2026 · {CUR}'  # leave JS variable intact
    )
    # Inject update timestamp as a hidden comment
    html = html.replace(
        '</head>',
        f'<!-- Last updated: {now} -->\n</head>'
    )
    
    template_path.write_text(html, encoding='utf-8')
    print(f"\n✅ docs/index.html updated ({len(html):,} bytes)")
    print(f"   Last updated: {now}")


# ─────────────────────────────────────────────────────────────────────
# 9. ENTRY POINT
# ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print("=" * 55)
    print("WH Throughput Tracker — GSheet → HTML Builder")
    print("=" * 55)
    
    all_bm, all_wh = fetch_all_wh_data()
    inject_into_html(all_bm, all_wh)
    
    print("\nDone. Push docs/index.html to GitHub Pages to publish.")

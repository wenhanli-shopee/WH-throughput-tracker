"""
fetch_and_build.py  (v4)
────────────────────────
- Reads months automatically from Monthly Tracker C2:F2
- Takes the 4 most recent "Actual" months (skips Projection columns)
- Reads all actual data from Monthly Tracker
- Reads benchmarks from IB Model / OB Model / Inventory Model col F
- Injects everything into docs/index.html

GitHub Secrets required:
  GOOGLE_CREDENTIALS   – Service Account JSON key (full content)
  SHEET_ID_PHB         – GSheet ID for PHB
  SHEET_ID_PHL         – GSheet ID for PHL
  SHEET_ID_PHIXC       – GSheet ID for PHIXC
"""

import os, json, time, datetime
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# ─── Auth ─────────────────────────────────────────────────────────────
def get_client():
    raw  = os.environ['GOOGLE_CREDENTIALS']
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

# ─── Helpers ──────────────────────────────────────────────────────────
def sf(v, default=0.0):
    """Safe float — strips commas, percent signs, handles None."""
    if v is None: return default
    try:
        return float(str(v).replace(',','').replace('%','').strip())
    except (TypeError, ValueError):
        return default

def sl(lst, length=4, default=0.0):
    """Safe list — pad/trim to exactly `length` elements."""
    out = [sf(v, default) for v in (lst or [])]
    while len(out) < length: out.append(default)
    return out[:length]

def open_sheet(gc, sid, retries=4, wait=30):
    """Open spreadsheet by key with exponential-backoff retry."""
    for attempt in range(retries):
        try:
            return gc.open_by_key(sid)
        except Exception as e:
            if attempt < retries - 1:
                print(f"    open_by_key failed ({e.__class__.__name__}), waiting {wait}s...")
                time.sleep(wait); wait *= 2
            else:
                raise

def read_once(ws, retries=4, wait=20):
    """Read entire sheet in ONE API call with retry on 429."""
    for attempt in range(retries):
        try:
            return ws.get_all_values()
        except gspread.exceptions.APIError as e:
            if '429' in str(e) and attempt < retries - 1:
                print(f"    Rate limited — waiting {wait}s (retry {attempt+2}/{retries})...")
                time.sleep(wait); wait *= 2
            else:
                raise
        except Exception as e:
            if attempt < retries - 1:
                print(f"    Read error ({e.__class__.__name__}), waiting {wait}s...")
                time.sleep(wait); wait *= 2
            else:
                raise
    return []

def gcf(rows, row_1based):
    """Get column-F value from pre-loaded rows (1-based row index)."""
    idx = row_1based - 1
    if 0 <= idx < len(rows) and len(rows[idx]) >= 6:
        return rows[idx][5]
    return None

def get_row(rows, row_1based, col_start_0, n=4):
    """Get n values starting at col_start_0 from a pre-loaded rows list."""
    idx = row_1based - 1
    if 0 <= idx < len(rows):
        row = rows[idx]
        return [row[c] if c < len(row) else None for c in range(col_start_0, col_start_0+n)]
    return [None]*n

# ─── Read months from Monthly Tracker R2 (C2:?2) ─────────────────────
def read_months(ws):
    """
    Row 2 of Monthly Tracker:
      Col B = None, C = '2026-01', D = '2026-02', E = '2026-03', F = '2026-04', ...
    Row 1: Col C onward = 'Actual' or 'Projection'

    We want the last 4 columns whose R1 label is 'Actual'.
    Returns (months_list, actual_col_indices_0based)
    e.g. (['Jan-26','Feb-26','Mar-26','Apr-26'], [2,3,4,5])
    """
    rows = read_once(ws)
    if len(rows) < 2:
        return [], []

    labels_row = rows[0]   # row 1: Actual / Projection labels
    months_row = rows[1]   # row 2: '2026-01', '2026-02', ...

    actual_cols = []
    for c in range(2, len(months_row)):   # start from col C (index 2)
        label = str(labels_row[c]).strip() if c < len(labels_row) else ''
        month = str(months_row[c]).strip() if c < len(months_row) else ''
        if label.lower() == 'actual' and month:
            actual_cols.append(c)

    # Take the last 4 actual columns
    actual_cols = actual_cols[-4:]

    # Format months: '2026-01' → 'Jan-26'
    month_map = {
        '01':'Jan','02':'Feb','03':'Mar','04':'Apr',
        '05':'May','06':'Jun','07':'Jul','08':'Aug',
        '09':'Sep','10':'Oct','11':'Nov','12':'Dec'
    }
    months = []
    for c in actual_cols:
        raw = str(months_row[c]).strip()   # e.g. '2026-01'
        parts = raw.split('-')
        if len(parts) == 2:
            yr = parts[0][-2:]   # '26'
            mo = month_map.get(parts[1], parts[1])
            months.append(f"{mo}-{yr}")
        else:
            months.append(raw)

    return months, actual_cols, rows   # return rows so we don't re-read

# ─── Read all actual data from Monthly Tracker ───────────────────────
def read_monthly_tracker(ws):
    """
    Read Monthly Tracker sheet ONCE and extract all actual data.
    Row references match the structure mapped from the Excel files.

    Returns a dict with all fields needed for WH[CUR] in the HTML.
    """
    months, actual_cols, rows = read_months(ws)
    n = len(actual_cols)
    if n == 0:
        print("    ⚠ No 'Actual' months found in Monthly Tracker row 2")
        return None

    print(f"    Months: {months} (cols {[c+1 for c in actual_cols]})")

    def gr(row_1based):
        """Get values at actual_cols from a given row."""
        idx = row_1based - 1
        if 0 <= idx < len(rows):
            row = rows[idx]
            return [row[c] if c < len(row) else None for c in actual_cols]
        return [None]*n

    def grsf(row_1based):
        return [sf(v) for v in gr(row_1based)]

    # ── Utilisation ───────────────────────────────────────────────────
    # R6  IB avg util
    # R7  INV (CBM Util) avg
    # R8  OB avg util
    # R11 IB peak util
    # R12 INV (CBM Util) peak
    # R13 INV (Max of Loc & CBM Util)
    # R14 OB peak util
    ib_avg_util   = grsf(6)
    inv_avg_util  = grsf(7)
    ob_avg_util   = grsf(8)
    ib_peak_util  = grsf(11)
    inv_peak_cbm  = grsf(12)
    inv_peak_loc  = grsf(13)
    ob_peak_util  = grsf(14)

    # ── Max Capacity ──────────────────────────────────────────────────
    # R18 IB max ADO (varies)
    # R19 INV max CBM
    # R20 OB max ADO
    # R21 MTO max ADO
    # R56 IB Max ADO (from Part 3 Theoretical Actual)
    # R75 INV max ADO
    # R88 OB max ADO
    # R116 MTO staging max ADO
    ib_max_ado   = grsf(56)
    inv_max_cbm  = grsf(19)
    inv_max_ado  = grsf(75)
    ob_max_ado   = grsf(88)
    mto_max_ado  = grsf(116)

    # ── Actual throughput (demand) ────────────────────────────────────
    # R39 Actual Avg ADO - OB,INV
    # R40 Actual Avg ADO - IB
    # R41 Actual Peak ADO - OB
    # R42 Actual Peak ADO - IB
    # R44 Peak MTO
    ob_actual_ado  = grsf(39)
    ib_actual_ado  = grsf(40)
    ob_peak_ado    = grsf(41)
    ib_peak_ado    = grsf(42)
    mto_peak_ado   = grsf(44)

    # ── IB Step capacities ────────────────────────────────────────────
    # R61 Max ADO: Pending Receiving Staging
    # R65 Max Avg ADO: Receiving
    # R69 Max ADO: Pending putaway Staging
    ib_step_staging  = grsf(61)
    ib_step_recv     = grsf(65)
    ib_step_putaway  = grsf(69)

    # ── OB Step capacities ────────────────────────────────────────────
    # R93 Max ADO: Sorting
    # R106 Max ADO: Packing
    # R109 Max ADO: Pre-sorting
    # R112 Max ADO: OB staging
    ob_step_sort    = grsf(93)
    ob_step_pack    = grsf(106)
    ob_step_presort = grsf(109)
    ob_step_staging = grsf(112)

    # ── MTO ───────────────────────────────────────────────────────────
    # R117 Avg handover duration
    # R15  MTO OB peak util
    mto_handover  = grsf(117)
    mto_peak_util = grsf(15)

    # ── Actual CBM ───────────────────────────────────────────────────
    # R81 CBM actual
    inv_actual_cbm = grsf(81)

    # ── INV Loc util ─────────────────────────────────────────────────
    # R86 Location Util
    inv_loc_util = grsf(86)

    # ── Space efficiency ─────────────────────────────────────────────
    # R29 AVG IB ADO/SQM
    # R30 AVG OB ADO/SQM
    # R31 AVG CBM/SQM
    # R32 AVG CBM/ADO
    # R34 MAX IB ADO/SQM
    # R35 MAX OB ADO/SQM
    avg_ib_sqm   = grsf(29)
    avg_ob_sqm   = grsf(30)
    avg_cbm_sqm  = grsf(31)
    avg_cbm_ado  = grsf(32)
    max_ib_sqm   = grsf(34)
    max_ob_sqm   = grsf(35)

    # ── Detailed act{} factors ───────────────────────────────────────
    # R62 Avg pallets pending receiving time
    # R64 item/pallet - pending recv
    # R66 PO receiving prod
    # R67 MTI receiving prod
    # R58 %MTI qty
    # R71 Avg pallets pending putaway time
    # R73 putaway hours
    # R74 item/pallet - staging (putaway)
    # R45 IOR
    # R46 DOC
    # R47 CBM/pcs
    # R85 CBM Util
    # R104 Sorting prod
    # R107 Packing prod
    # R103 Sorting items %
    # R115 item per staging pallet (OB)
    act = {
        'ib_wait_recv':  grsf(62),
        'ib_items_recv': grsf(64),
        'ib_po_prod':    grsf(66),
        'ib_mt_prod':    grsf(67),
        'ib_pct_mti':    grsf(58),
        'ib_wait_put':   grsf(71),
        'ib_put_hrs':    grsf(73),
        'ib_items_put':  grsf(74),
        'inv_ior':       grsf(45),
        'inv_doc':       grsf(46),
        'inv_cbm_pcs':   grsf(47),
        'inv_cbm_util':  grsf(85),
        'ob_sort':       grsf(104),
        'ob_pack':       grsf(107),
        'ob_sort_pct':   grsf(103),
        'ob_items_pal':  grsf(115),
    }

    # Build ibSteps / obSteps as month-keyed dicts
    # Month keys for HTML: jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec
    month_keys = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    mo_keys_used = []
    for m in months:
        mk = m[:3].lower()
        mo_keys_used.append(mk)

    ib_steps_dict = {}
    ob_steps_dict = {}
    for i, mk in enumerate(mo_keys_used):
        ib_steps_dict[mk] = [
            sf(ib_step_staging[i] if i < len(ib_step_staging) else 0),
            sf(ib_step_recv[i]    if i < len(ib_step_recv)    else 0),
            sf(ib_step_putaway[i] if i < len(ib_step_putaway) else 0),
        ]
        ob_steps_dict[mk] = [
            sf(ob_step_sort[i]    if i < len(ob_step_sort)    else 0),
            sf(ob_step_pack[i]    if i < len(ob_step_pack)    else 0),
            sf(ob_step_presort[i] if i < len(ob_step_presort) else 0),
            sf(ob_step_staging[i] if i < len(ob_step_staging) else 0),
        ]

    # Detect bottlenecks
    def detect_bn(ib_max, inv_max, ob_max):
        vals = [sf(ib_max), sf(inv_max), sf(ob_max)]
        mn = min(v for v in vals if v > 0) if any(v > 0 for v in vals) else 0
        if mn == vals[0]: return 'IB'
        if mn == vals[1]: return 'INV'
        return 'OB'

    # Use last month's values for overall bottleneck
    last = -1
    system_bn = detect_bn(
        ib_max_ado[last] if ib_max_ado else 0,
        inv_max_ado[last] if inv_max_ado else 0,
        ob_max_ado[last]  if ob_max_ado  else 0
    )

    # IB step BN: min of staging, recv, putaway
    def ib_step_bn(i):
        s = [sf(ib_step_staging[i]),sf(ib_step_recv[i]),sf(ib_step_putaway[i])]
        mn = min(v for v in s if v>0) if any(v>0 for v in s) else 0
        names = ['Arrival Staging','Receiving','Putaway Staging']
        return names[s.index(mn)] if mn > 0 else 'Receiving'

    # OB step BN: min of sort, pack, presort, staging
    def ob_step_bn(i):
        s = [sf(ob_step_sort[i]),sf(ob_step_pack[i]),sf(ob_step_presort[i]),sf(ob_step_staging[i])]
        mn = min(v for v in s if v>0) if any(v>0 for v in s) else 0
        names = ['Sorting','Packing','Pre-sorting','PO Staging']
        return names[s.index(mn)] if mn > 0 else 'Packing'

    ib_bn = ib_step_bn(last)
    ob_bn = ob_step_bn(last)

    return {
        'months':  months,
        'mo_keys': mo_keys_used,
        'ib': {
            'avgUtil':   ib_avg_util,
            'peakUtil':  ib_peak_util,
            'actualADO': ib_actual_ado,
            'peakADO':   ib_peak_ado,
            'maxADO':    ib_max_ado,
        },
        'inv': {
            'avgUtil':    inv_avg_util,
            'peakUtil':   inv_peak_cbm,
            'locPeakUtil':inv_peak_loc,
            'maxADO':     inv_max_ado,
            'maxCBM':     inv_max_cbm,
            'actualCBM':  inv_actual_cbm,
        },
        'ob': {
            'avgUtil':   ob_avg_util,
            'peakUtil':  ob_peak_util,
            'actualADO': ob_actual_ado,
            'peakADO':   ob_peak_ado,
            'maxADO':    ob_max_ado,
        },
        'mto': {
            'peakUtil': mto_peak_util,
            'maxADO':   mto_max_ado,
            'handover': mto_handover,
        },
        'ibSteps':  ib_steps_dict,
        'obSteps':  ob_steps_dict,
        'spaceEff': {
            'avgIBADOsqm': avg_ib_sqm,
            'avgOBADOsqm': avg_ob_sqm,
            'avgCBMsqm':   avg_cbm_sqm,
            'avgCBMado':   avg_cbm_ado,
            'maxIBADOsqm': max_ib_sqm,
            'maxOBADOsqm': max_ob_sqm,
        },
        'act':        act,
        'bottleneck': system_bn,
        'ibBN':       ib_bn,
        'obBN':       ob_bn,
    }

# ─── IB Model benchmarks ─────────────────────────────────────────────
def read_ib(ws):
    rows = read_once(ws)
    cf = lambda r: gcf(rows, r)
    return {
        'ib_ior':   sf(cf(5)),    'pct_mti':  sf(cf(14)),
        's1_items': sf(cf(38)),   'spu_s1':   sf(cf(40)),
        'po_prod':  sf(cf(41)),   'mt_prod':  sf(cf(42)),
        'r_hrs':    sf(cf(46)),   'spu_r':    sf(cf(48)),
        's3_hrs':   sf(cf(51)),   'put_prod': sf(cf(52)),
        's3_wait':  sf(cf(69)),   's3_items': sf(cf(61)),
        'spu_s3':   sf(cf(63)),   's1_wait':  sf(cf(68)),
        's1_truck': sf(cf(26)),
        'sp_s1':    sf(cf(82)),   'sp_r':    sf(cf(83)),   'sp_s3':  sf(cf(84)),
        's1_pal':   sf(cf(86)),   'r_stns':  sf(cf(87)),   's3_pal': sf(cf(88)),
        'bm_ib_staging': sf(cf(94)),
        'bm_ib_recv':    sf(cf(101)),
        'bm_ib_putaway': sf(cf(107)),
        'bm_ib':         sf(cf(113)),
    }

# ─── OB Model benchmarks ─────────────────────────────────────────────
def read_ob(ws):
    rows = read_once(ws)
    cf = lambda r: gcf(rows, r)
    return {
        'ob_ior':    sf(cf(5)),    'stg_items': sf(cf(28)),
        'sort_pct':  sf(cf(43)),   'sort_prod': sf(cf(44)),  'pack_prod': sf(cf(45)),
        'spu_sort':  sf(cf(48)),   'spu_pack':  sf(cf(49)),
        'pre_prod':  sf(cf(51)),   'spu_pre':   sf(cf(53)),
        'stg_pal':   sf(cf(78)),   'spu_stg':   sf(cf(31)),
        'sort_hrs':  24,  'pack_hrs': 24,  'pre_hrs': 24,
        'sp_sort':   sf(cf(68)),   'sp_pack':   sf(cf(69)),
        'sp_pre':    sf(cf(70)),   'sp_stg':    sf(cf(71)),
        'sort_stns': sf(cf(73)),   'pack_stns': sf(cf(74)),  'pre_stns': sf(cf(75)),
        'bm_ob_sorting':  sf(cf(81)),
        'bm_ob_packing':  sf(cf(87)),
        'bm_ob_presort':  sf(cf(93)),
        'bm_ob_staging':  sf(cf(99)),
        'bm_ob':          sf(cf(81)),
    }

# ─── Inventory Model benchmarks ──────────────────────────────────────
def read_inv(ws):
    rows = read_once(ws)
    theo = next((i for i, r in enumerate(rows)
                 if any('Theoretical calculation' in str(c) for c in r)), None)
    if theo is None:
        print("    ⚠ 'Theoretical calculation' not found in Inventory Model")
        return {}

    def gv(off, col=2):
        r = theo + off
        return sf(rows[r][col], None) if 0 <= r < len(rows) and col < len(rows[r]) else None

    result = {}
    for off, key in [(3,'inv_ior'),(5,'doc'),(7,'cbm_pcs'),
                     (10,'cbm_sqm'),(14,'cbm_util'),(15,'sqm')]:
        v = gv(off)
        if v is not None:
            result[key] = v

    if 'sqm' not in result:
        for off in range(12, 22):
            r = theo + off
            if r >= len(rows): break
            row_str = ' '.join(str(c) for c in rows[r])
            if 'SQM' in row_str and 'Actual' in row_str:
                try: result['sqm'] = sf(rows[r][2]); break
                except: pass
    return result

# ─── Fetch all WHs ───────────────────────────────────────────────────
def fetch_all():
    gc = get_client()
    cfg = {
        'PHB':   os.environ['SHEET_ID_PHB'],
        'PHL':   os.environ['SHEET_ID_PHL'],
        'PHIXC': os.environ['SHEET_ID_PHIXC'],
    }
    all_bm, all_wh = {}, {}

    for idx, (wh, sid) in enumerate(cfg.items()):
        if idx > 0:
            print(f"  Waiting 30s before {wh} (rate limit buffer)...")
            time.sleep(30)

        print(f"\nReading {wh} ({sid[:8]}...)")
        sh = open_sheet(gc, sid)

        # Read benchmarks (3 API calls)
        bm = {
            **read_ib(sh.worksheet('IB Model')),
            **read_ob(sh.worksheet('OB Model')),
            **read_inv(sh.worksheet('Inventory Model')),
        }
        all_bm[wh] = bm

        # Read actual data from Monthly Tracker (1 API call)
        mt_ws = sh.worksheet('Monthly Tracker')
        data  = read_monthly_tracker(mt_ws)

        if data is None:
            print(f"  ⚠ {wh}: Monthly Tracker read failed, using zeros")
            months = ['Jan-26','Feb-26','Mar-26','Apr-26']
            b4 = [0.0]*4
            data = {
                'months': months, 'mo_keys': ['jan','feb','mar','apr'],
                'ib':  {'avgUtil':b4,'peakUtil':b4,'actualADO':b4,'peakADO':b4,'maxADO':b4},
                'inv': {'avgUtil':b4,'peakUtil':b4,'locPeakUtil':b4,'maxADO':b4,'maxCBM':b4,'actualCBM':b4},
                'ob':  {'avgUtil':b4,'peakUtil':b4,'actualADO':b4,'peakADO':b4,'maxADO':b4},
                'mto': {'peakUtil':b4,'maxADO':b4,'handover':b4},
                'ibSteps':{'jan':[0,0,0],'feb':[0,0,0],'mar':[0,0,0],'apr':[0,0,0]},
                'obSteps':{'jan':[0,0,0,0],'feb':[0,0,0,0],'mar':[0,0,0,0],'apr':[0,0,0,0]},
                'spaceEff':None, 'act':None, 'bottleneck':'INV', 'ibBN':'Receiving', 'obBN':'Packing',
            }

        data['theo'] = {
            'ibADO':  bm.get('bm_ib', 0),
            'invADO': 0,
            'obADO':  bm.get('bm_ob', 0),
        }
        data['name'] = wh
        all_wh[wh] = data

        print(f"  ✓ {wh}: months={data['months']}  "
              f"bm_ib={bm.get('bm_ib',0):.0f}  "
              f"bm_ob={bm.get('bm_ob',0):.0f}  "
              f"sqm={bm.get('sqm',0):.0f}")

    return all_bm, all_wh

# ─── HTML injection ───────────────────────────────────────────────────
def bm_js(wh, bm):
    body = '\n'.join(f"  {k}:{json.dumps(v)}," for k, v in bm.items())
    return f"const BM_{wh}={{\n{body}\n}};"

def wh_js(wh, d):
    js = json.dumps
    ib  = d['ib'];   inv = d['inv'];  ob  = d['ob']
    mto = d.get('mto', {})
    ibs = d.get('ibSteps', {})
    obs = d.get('obSteps', {})
    th  = d.get('theo', {})
    mo_keys = d.get('mo_keys', ['jan','feb','mar','apr'])

    # ibSteps / obSteps — emit all available month keys
    def steps_js(steps_dict):
        parts = ','.join(f"{k}:{js(v)}" for k,v in steps_dict.items())
        return '{'+parts+'}'

    return (
        f"{wh}:{{name:{js(wh)},months:{js(d['months'])},"
        f"theo:{{ibADO:{th.get('ibADO',0)},invADO:{th.get('invADO',0)},obADO:{th.get('obADO',0)}}},"
        f"ib:{{avgUtil:{js(ib['avgUtil'])},peakUtil:{js(ib['peakUtil'])},"
        f"actualADO:{js(ib['actualADO'])},peakADO:{js(ib.get('peakADO',[0,0,0,0]))},"
        f"maxADO:{js(ib['maxADO'])}}},"
        f"inv:{{avgUtil:{js(inv['avgUtil'])},peakUtil:{js(inv['peakUtil'])},"
        f"locPeakUtil:{js(inv.get('locPeakUtil',[0,0,0,0]))},"
        f"maxADO:{js(inv['maxADO'])},maxCBM:{js(inv.get('maxCBM',[0,0,0,0]))},"
        f"actualCBM:{js(inv.get('actualCBM',[0,0,0,0]))}}},"
        f"ob:{{avgUtil:{js(ob['avgUtil'])},peakUtil:{js(ob['peakUtil'])},"
        f"actualADO:{js(ob['actualADO'])},peakADO:{js(ob.get('peakADO',[0,0,0,0]))},"
        f"maxADO:{js(ob['maxADO'])}}},"
        f"mto:{{peakUtil:{js(mto.get('peakUtil',[0,0,0,0]))},"
        f"maxADO:{js(mto.get('maxADO',[0,0,0,0]))},"
        f"handover:{js(mto.get('handover',[0,0,0,0]))}}},"
        f"ibSteps:{steps_js(ibs)},"
        f"obSteps:{steps_js(obs)},"
        f"spaceEff:{js(d.get('spaceEff'))},act:{js(d.get('act'))},"
        f"bottleneck:{js(d.get('bottleneck','INV'))},"
        f"ibBN:{js(d.get('ibBN','Receiving'))},obBN:{js(d.get('obBN','Packing'))}}}"
    )

def block_end(html, start):
    """Find the end position of a JS object block (past the closing };)."""
    depth = 0
    for i, c in enumerate(html[start:], start):
        if c == '{': depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                end = i + 1
                return end + (1 if html[end:end+1] == ';' else 0)
    return len(html)

def inject(all_bm, all_wh):
    p = Path(__file__).parent.parent / 'docs' / 'index.html'
    if not p.exists():
        raise FileNotFoundError(f"docs/index.html not found at {p}")
    html = p.read_text(encoding='utf-8')

    # Replace BM_PHB / BM_PHL / BM_PHIXC
    for wh in all_bm:
        tag = f'const BM_{wh}={{'
        s = html.find(tag)
        if s == -1:
            print(f"  ⚠ BM_{wh} not found in HTML — skipping")
            continue
        html = html[:s] + bm_js(wh, all_bm[wh]) + html[block_end(html, s):]
        print(f"  ✓ BM_{wh} injected")

    # Replace const WH={...}
    s = html.find('const WH={')
    if s != -1:
        entries = ',\n  '.join(wh_js(wh, all_wh[wh]) for wh in all_wh)
        html = html[:s] + f"const WH={{\n  {entries}\n}};" + html[block_end(html, s):]
        print(f"  ✓ WH data injected ({list(all_wh.keys())})")
    else:
        print("  ⚠ const WH not found in HTML")

    # Timestamp
    now = datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    html = html.replace('</head>', f'<!-- Last updated: {now} -->\n</head>', 1)
    p.write_text(html, encoding='utf-8')
    print(f"\n✅ docs/index.html updated ({len(html):,} bytes)  [{now}]")

# ─── Entry point ──────────────────────────────────────────────────────
if __name__ == '__main__':
    print("=" * 55)
    print("WH Throughput Tracker — GSheet → HTML Builder v4")
    print("=" * 55)
    all_bm, all_wh = fetch_all()
    inject(all_bm, all_wh)
    print("\nDone.")

"""
fetch_and_build.py  (v2 — batch reads, retry on 429, GOOGLE_CREDENTIALS secret)
"""
import os, json, time, datetime
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def get_client():
    raw  = os.environ['GOOGLE_CREDENTIALS']      # ← updated secret name
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

def sf(v, default=0.0):
    try:
        return float(str(v).replace(',','').replace('%','').strip())
    except (TypeError, ValueError):
        return default

def safe_list(lst, length=4, default=0.0):
    out = [sf(v, default) for v in lst]
    while len(out) < length: out.append(default)
    return out[:length]

def read_once(ws, retries=3, wait=20):
    """Read entire sheet in ONE API call, with exponential-backoff retry on 429."""
    for attempt in range(retries):
        try:
            return ws.get_all_values()
        except gspread.exceptions.APIError as e:
            if '429' in str(e) and attempt < retries - 1:
                print(f"    Rate limited — waiting {wait}s (retry {attempt+2}/{retries})...")
                time.sleep(wait)
                wait *= 2
            else:
                raise
    return []

def gcf(rows, row_1based):
    """Get column-F value from pre-loaded rows (1-based row index)."""
    idx = row_1based - 1
    if 0 <= idx < len(rows) and len(rows[idx]) >= 6:
        return rows[idx][5]
    return None

# ─── IB Model ────────────────────────────────────────────────────────
def read_ib(ws):
    rows = read_once(ws)
    cf = lambda r: gcf(rows, r)
    return {
        'ib_ior':  sf(cf(5)),   'pct_mti':  sf(cf(14)),
        's1_items':sf(cf(38)),  'spu_s1':   sf(cf(40)),
        'po_prod': sf(cf(41)),  'mt_prod':  sf(cf(42)),
        'r_hrs':   sf(cf(46)),  'spu_r':    sf(cf(48)),
        's3_hrs':  sf(cf(51)),  'put_prod': sf(cf(52)),
        's3_wait': sf(cf(69)),  's3_items': sf(cf(61)),
        'spu_s3':  sf(cf(63)),  's1_wait':  sf(cf(68)),
        's1_truck':sf(cf(26)),
        'sp_s1':   sf(cf(82)),  'sp_r':     sf(cf(83)),  'sp_s3':  sf(cf(84)),
        's1_pal':  sf(cf(86)),  'r_stns':   sf(cf(87)),  's3_pal': sf(cf(88)),
        'bm_ib_staging': sf(cf(94)),
        'bm_ib_recv':    sf(cf(101)),
        'bm_ib_putaway': sf(cf(107)),
        'bm_ib':         sf(cf(113)),
    }

# ─── OB Model ────────────────────────────────────────────────────────
def read_ob(ws):
    rows = read_once(ws)
    cf = lambda r: gcf(rows, r)
    return {
        'ob_ior':   sf(cf(5)),   'stg_items':sf(cf(28)),
        'sort_pct': sf(cf(43)),  'sort_prod':sf(cf(44)),  'pack_prod':sf(cf(45)),
        'spu_sort': sf(cf(48)),  'spu_pack': sf(cf(49)),
        'pre_prod': sf(cf(51)),  'spu_pre':  sf(cf(53)),
        'stg_pal':  sf(cf(78)),  'spu_stg':  sf(cf(31)),
        'sort_hrs': 24, 'pack_hrs': 24, 'pre_hrs': 24,
        'sp_sort':  sf(cf(68)),  'sp_pack':  sf(cf(69)),
        'sp_pre':   sf(cf(70)),  'sp_stg':   sf(cf(71)),
        'sort_stns':sf(cf(73)),  'pack_stns':sf(cf(74)),  'pre_stns':sf(cf(75)),
        'bm_ob_sorting':  sf(cf(81)),
        'bm_ob_packing':  sf(cf(87)),
        'bm_ob_presort':  sf(cf(93)),
        'bm_ob_staging':  sf(cf(99)),
        'bm_ob':          sf(cf(81)),
    }

# ─── Inventory Model ─────────────────────────────────────────────────
def read_inv(ws):
    rows = read_once(ws)
    # Find 'Theoretical calculation' header
    theo = next((i for i, r in enumerate(rows)
                 if any('Theoretical calculation' in str(c) for c in r)), None)
    if theo is None:
        print("    ⚠ 'Theoretical calculation' not found in Inventory Model")
        return {}

    def gv(off, col=2):
        r = theo + off
        return sf(rows[r][col], None) if 0 <= r < len(rows) and col < len(rows[r]) else None

    result = {}
    for off, key in [(3,'inv_ior'),(5,'doc'),(7,'cbm_pcs'),(10,'cbm_sqm'),(14,'cbm_util'),(15,'sqm')]:
        v = gv(off)
        if v is not None:
            result[key] = v

    if 'sqm' not in result:   # fallback search
        for off in range(12, 22):
            r = theo + off
            if r >= len(rows): break
            row_str = ' '.join(str(c) for c in rows[r])
            if 'SQM' in row_str and 'Actual' in row_str:
                try: result['sqm'] = sf(rows[r][2]); break
                except: pass
    return result

# ─── Actual data (placeholder — expand once tab layout confirmed) ─────
def read_actual(sh):
    def open_tab(*names):
        for n in names:
            try: return sh.worksheet(n)
            except gspread.WorksheetNotFound: continue
        return None

    months = ['Jan-26','Feb-26','Mar-26','Apr-26']
    b4 = [0.0,0.0,0.0,0.0]

    # Open tabs — each is read ONCE if found
    ib_ws  = open_tab('IB dashboard','IB Dashboard')
    ob_ws  = open_tab('OB dashboard','OB Dashboard')
    raw_ws = open_tab('actual raw data','Actual Raw Data','Raw Data')

    ib_rows  = read_once(ib_ws)  if ib_ws  else []
    ob_rows  = read_once(ob_ws)  if ob_ws  else []
    raw_rows = read_once(raw_ws) if raw_ws else []

    # ── TODO: map actual row/col positions here ───────────────────────
    # Example (uncomment and adjust row numbers once confirmed):
    #
    # def gr(rows, row_1based, col_start, n=4):
    #     idx = row_1based - 1
    #     return safe_list(rows[idx][col_start:col_start+n]) if idx < len(rows) else b4
    #
    # ib_avg_util  = gr(ib_rows,  5, 2)
    # ib_peak_util = gr(ib_rows,  6, 2)
    # ib_max_ado   = gr(ib_rows, 10, 2)
    # ── end TODO ──────────────────────────────────────────────────────

    return {
        'months': months,
        'ib':  {'avgUtil':b4,'peakUtil':b4,'actualADO':b4,'peakADO':b4,'maxADO':b4},
        'inv': {'avgUtil':b4,'peakUtil':b4,'locPeakUtil':b4,'maxADO':b4,'maxCBM':b4,'actualCBM':b4},
        'ob':  {'avgUtil':b4,'peakUtil':b4,'actualADO':b4,'peakADO':b4,'maxADO':b4},
        'mto': {'peakUtil':b4,'maxADO':b4,'handover':b4},
        'ibSteps':{'jan':[0,0,0],'feb':[0,0,0],'mar':[0,0,0],'apr':[0,0,0]},
        'obSteps':{'jan':[0,0,0,0],'feb':[0,0,0,0],'mar':[0,0,0,0],'apr':[0,0,0,0]},
        'act':None,'spaceEff':None,'bottleneck':'INV','ibBN':'Receiving','obBN':'Packing',
    }

# ─── Fetch all three WHs ─────────────────────────────────────────────
def fetch_all():
    gc = get_client()
    cfg = {'PHB':os.environ['SHEET_ID_PHB'],'PHL':os.environ['SHEET_ID_PHL'],'PHIXC':os.environ['SHEET_ID_PHIXC']}
    all_bm, all_wh = {}, {}

    for idx, (wh, sid) in enumerate(cfg.items()):
        if idx > 0:
            print(f"  Waiting 15s before {wh} (rate limit buffer)...")
            time.sleep(15)
        print(f"\nReading {wh} ({sid[:8]}...)")
        sh  = gc.open_by_key(sid)
        bm  = {**read_ib(sh.worksheet('IB Model')),
               **read_ob(sh.worksheet('OB Model')),
               **read_inv(sh.worksheet('Inventory Model'))}
        all_bm[wh] = bm
        act = read_actual(sh)
        act['theo'] = {'ibADO':bm.get('bm_ib',0),'invADO':0,'obADO':bm.get('bm_ob',0)}
        act['name'] = wh
        all_wh[wh] = act
        print(f"  ✓ {wh}: bm_ib={bm.get('bm_ib',0):.0f}  bm_ob={bm.get('bm_ob',0):.0f}  sqm={bm.get('sqm',0):.0f}")

    return all_bm, all_wh

# ─── HTML injection ───────────────────────────────────────────────────
def bm_js(wh, bm):
    body = '\n'.join(f"  {k}:{json.dumps(v)}," for k, v in bm.items())
    return f"const BM_{wh}={{\n{body}\n}};"

def wh_js(wh, d):
    js = json.dumps
    ib=d['ib']; inv=d['inv']; ob=d['ob']; mto=d.get('mto',{})
    ibs=d.get('ibSteps',{}); obs=d.get('obSteps',{}); th=d.get('theo',{})
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
        f"ibSteps:{{jan:{js(ibs.get('jan',[0,0,0]))},feb:{js(ibs.get('feb',[0,0,0]))},"
        f"mar:{js(ibs.get('mar',[0,0,0]))},apr:{js(ibs.get('apr',[0,0,0]))}}},"
        f"obSteps:{{jan:{js(obs.get('jan',[0,0,0,0]))},feb:{js(obs.get('feb',[0,0,0,0]))},"
        f"mar:{js(obs.get('mar',[0,0,0,0]))},apr:{js(obs.get('apr',[0,0,0,0]))}}},"
        f"spaceEff:{js(d.get('spaceEff'))},act:{js(d.get('act'))},"
        f"bottleneck:{js(d.get('bottleneck','INV'))},"
        f"ibBN:{js(d.get('ibBN','Receiving'))},obBN:{js(d.get('obBN','Packing'))}}}"
    )

def block_end(html, start):
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

    for wh in ['PHB','PHL','PHIXC']:
        tag = f'const BM_{wh}={{'
        s = html.find(tag)
        if s == -1: print(f"  ⚠ BM_{wh} not found"); continue
        html = html[:s] + bm_js(wh, all_bm[wh]) + html[block_end(html, s):]
        print(f"  ✓ BM_{wh} injected")

    s = html.find('const WH={')
    if s != -1:
        entries = ',\n  '.join(wh_js(wh, all_wh[wh]) for wh in ['PHB','PHL','PHIXC'])
        html = html[:s] + f"const WH={{\n  {entries}\n}};" + html[block_end(html, s):]
        print(f"  ✓ WH data injected")
    else:
        print("  ⚠ const WH not found")

    now = datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    html = html.replace('</head>', f'<!-- Last updated: {now} -->\n</head>', 1)
    p.write_text(html, encoding='utf-8')
    print(f"\n✅ docs/index.html updated ({len(html):,} bytes)  [{now}]")

# ─── Entry point ──────────────────────────────────────────────────────
if __name__ == '__main__':
    print("="*55)
    print("WH Throughput Tracker — GSheet → HTML Builder v2")
    print("="*55)
    all_bm, all_wh = fetch_all()
    inject(all_bm, all_wh)
    print("\nDone.")

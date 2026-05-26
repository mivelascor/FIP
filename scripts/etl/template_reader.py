"""
etl/template_reader.py  —  Fuente definitiva de datos para folletos Vantrust.

REGLAS (verificadas contra PDFs manuales):
─ CLP: resumen M/T/S/A/Ac calculado desde ODS VC (simple return).
  - Para meses antes de ODS (jul-2025), se encadena VC desde template historico.
─ USD: resumen y histórico leídos directamente del template rentabilidad sheet
  (el ODS tiene ligeras diferencias por fuente; el template es más exacto).
─ ICP: serie CLICP desde icp_clicp.json + extensión TPM para meses futuros.
─ Competencia CLP: Santander MM desde comp_clp.json.
─ Competencia USD: leída del template (col K = Banchile Corporate Dollar normalizado).
"""
import os, requests, json, calendar
from pathlib import Path
from datetime import date
from dateutil.relativedelta import relativedelta
import pandas as pd

_INPUTS_DIR = Path(__file__).parent.parent.parent / "inputs"
_TMPL_DIR   = _INPUTS_DIR / "templates"

# ── Mapa: ODS nombre → template file con datos históricos ────────────────────
FUND_TEMPLATE_MAP = {
    # Direct 1:1 mapping — each template has data for its own fund
    'FIP VANTRUST LIQUIDEZ ACTIVA':        'TEMPLATE FONDO LIQUIDEZ ACTIVA.xlsx',
    'FIP VANTRUST LIQUIDEZ ALTO APORTE':   'TEMPLATE FONDO ALTO APORTE.xlsx',
    'FIP VANTRUST LIQUIDEZ ALTO CAPITAL':  'TEMPLATE FONDO ALTO CAPITAL.xlsx',
    'FIP VANTRUST LIQUIDEZ ALTO MONTO':    'TEMPLATE FONDO LIQUIDEZ ALTO MONTO.xlsx',
    'FIP VANTRUST LIQUIDEZ CAJA':          'TEMPLATE FONDO LIQUIDEZ CAJA.xlsx',
    'FIP VANTRUST LIQUIDEZ CONTINUA':      'TEMPLATE FONDO LIQUIDEZ CONTINUA.xlsx',
    'FIP VANTRUST LIQUIDEZ CORRIENTE':     'TEMPLATE FONDO LIQUIDEZ CORRIENTE.xlsx',
    'FIP VANTRUST LIQUIDEZ CORTO PLAZO':   'TEMPLATE FONDO LIQUIDEZ CORTO PLAZO.xlsx',
    'FIP VANTRUST LIQUIDEZ DISPONIBLE I':  'TEMPLATE FONDO LIQUIDEZ Disponible I.xlsx',
    'FIP VANTRUST LIQUIDEZ DOLAR':         'TEMPLATE FONDO LIQUIDEZ DOLAR.xlsx',
    'FIP VANTRUST LIQUIDEZ DOLAR CAJA':    'TEMPLATE FONDO LIQUIDEZ DOLAR CAJA.xlsx',
    'FIP VANTRUST LIQUIDEZ EFECTIVO':      'TEMPLATE FONDO LIQUIDEZ EFECTIVO.xlsx',
    'FIP VANTRUST LIQUIDEZ FLEXIBLE':      'TEMPLATE FONDO LIQUIDEZ FLEXIBLE.xlsx',
    'FIP VANTRUST LIQUIDEZ I':             'TEMPLATE FONDO LIQUIDEZ UNO.xlsx',
    'FIP VANTRUST LIQUIDEZ LOCAL':         'TEMPLATE FONDO LIQUIDEZ LOCAL.xlsx',
    'FIP VANTRUST LIQUIDEZ MONETARIO I':   'TEMPLATE FONDO LIQUIDEZ Monetario I.xlsx',
    'FIP VANTRUST LIQUIDEZ PERMANENTE':    'TEMPLATE FONDO LIQUIDEZ Permanente.xlsx',
    'FIP VANTRUST LIQUIDEZ PLUS':          'TEMPLATE FONDO LIQUIDEZ PLUS.xlsx',
    'FIP VANTRUST LIQUIDEZ PRESENTE':      'TEMPLATE FONDO LIQUIDEZ Presente.xlsx',
    'FIP VANTRUST LIQUIDEZ RENDIMIENTO':   'TEMPLATE FONDO LIQUIDEZ RENDIMIENTO.xlsx',
    'FIP VANTRUST LIQUIDEZ RESERVA DÓLAR': 'TEMPLATE FONDO LIQUIDEZ RESERVA DOLAR.xlsx',
    'FIP VANTRUST LIQUIDEZ SENCILLO':      'TEMPLATE FONDO LIQUIDEZ SENCILLO.xlsx',
}

FONDOS_USD = {
    'FIP VANTRUST LIQUIDEZ DOLAR',
    'FIP VANTRUST LIQUIDEZ DOLAR CAJA',
    'FIP VANTRUST LIQUIDEZ RESERVA DÓLAR',
}

# Funds where template summary is used (instead of ODS calculation).
# This includes ALL funds that have FIP data in their template rentabilidad sheet.
# USD funds: always use template (different formula + data source)
# CLP funds with template: use template (SharePoint data matches PDFs better than ODS)
FONDOS_USE_TEMPLATE = FONDOS_USD | set(FUND_TEMPLATE_MAP.keys())

NOMBRE_DISPLAY = {
    'FIP VANTRUST LIQUIDEZ ACTIVA':        'FIP Liquidez Activa',
    'FIP VANTRUST LIQUIDEZ ALTO APORTE':   'FIP Liquidez Alto Aporte',
    'FIP VANTRUST LIQUIDEZ ALTO CAPITAL':  'FIP Alto Capital',
    'FIP VANTRUST LIQUIDEZ ALTO MONTO':    'FIP Liquidez Alto Monto',
    'FIP VANTRUST LIQUIDEZ CAJA':          'FIP Liquidez Caja',
    'FIP VANTRUST LIQUIDEZ CONTINUA':      'FIP Liquidez Continua',
    'FIP VANTRUST LIQUIDEZ CORRIENTE':     'FIP Liquidez Corriente',
    'FIP VANTRUST LIQUIDEZ CORTO PLAZO':   'FIP Liquidez Corto Plazo',
    'FIP VANTRUST LIQUIDEZ DISPONIBLE I':  'FIP Liquidez Disponible I',
    'FIP VANTRUST LIQUIDEZ DOLAR':         'FIP Liquidez Dólar',
    'FIP VANTRUST LIQUIDEZ DOLAR CAJA':    'FIP Liquidez Dólar Caja',
    'FIP VANTRUST LIQUIDEZ EFECTIVO':      'FIP Liquidez Efectivo',
    'FIP VANTRUST LIQUIDEZ FLEXIBLE':      'FIP Liquidez Flexible',
    'FIP VANTRUST LIQUIDEZ I':             'FIP Liquidez I',
    'FIP VANTRUST LIQUIDEZ LOCAL':         'FIP Liquidez Local',
    'FIP VANTRUST LIQUIDEZ MONETARIO I':   'FIP Liquidez Monetario I',
    'FIP VANTRUST LIQUIDEZ PERMANENTE':    'FIP Liquidez Permanente',
    'FIP VANTRUST LIQUIDEZ PLUS':          'FIP Liquidez Plus',
    'FIP VANTRUST LIQUIDEZ PRESENTE':      'FIP Liquidez Presente',
    'FIP VANTRUST LIQUIDEZ RECURRENTE':    'FIP Liquidez Recurrente',
    'FIP VANTRUST LIQUIDEZ RENDIMIENTO':   'FIP Liquidez Rendimiento',
    'FIP VANTRUST LIQUIDEZ RESERVA DÓLAR': 'FIP Liquidez Reserva Dólar',
    'FIP VANTRUST LIQUIDEZ SENCILLO':      'FIP Liquidez Sencillo',
    'FIP VANTRUST LIQUIDEZ TEMPORAL':      'FIP Liquidez Temporal',
}

_ICP_CACHE:  dict = {}
_VC_CACHE:   dict = {}
_HIST_CACHE: dict = {}
_ODS_API = "https://claudeods.vantrustcapital.cl/query"

# ── Helpers ───────────────────────────────────────────────────────────────────
def _load_json(fn):
    try:
        with open(_INPUTS_DIR / fn) as f:
            return {(int(k[:4]), int(k[5:7])): float(v) for k,v in json.load(f).items()}
    except Exception as e:
        print(f"  [WARN] {fn}: {e}"); return {}

def _eom(y, m):
    return date(y, m, calendar.monthrange(y, m)[1])

def _prev(y, m, n=1):
    d = date(y, m, 1) - relativedelta(months=n)
    return d.year, d.month

# ── ICP ───────────────────────────────────────────────────────────────────────
def _get_icp_series():
    global _ICP_CACHE
    if _ICP_CACHE: return _ICP_CACHE
    clicp = _load_json("icp_clicp.json")
    after, base = (max(clicp.keys()), clicp[max(clicp.keys())]) if clicp else ((2005,12), 10000.0)
    _extend_tpm(clicp, after, base)
    _ICP_CACHE = clicp
    return _ICP_CACHE

def _extend_tpm(series, after_ym, base_val):
    prev = base_val
    for year in range(after_ym[0], date.today().year + 1):
        try:
            r  = requests.get(f"https://mindicador.cl/api/tpm/{year}", timeout=15)
            df = pd.DataFrame([{"fecha": pd.to_datetime(i["fecha"]), "tpm": float(i["valor"])}
                                for i in r.json().get("serie", [])])
            if df.empty: continue
            df["ym"] = df["fecha"].dt.to_period("M")
            for _, row in df.groupby("ym")["tpm"].mean().reset_index().iterrows():
                ym = (row["ym"].year, row["ym"].month)
                if ym <= after_ym: continue
                prev = prev * (1 + row["tpm"] / 1200)
                series[ym] = prev
        except: pass

# ── ODS fetch ─────────────────────────────────────────────────────────────────
def _get_ods_vc(nombre):
    if nombre in _VC_CACHE: return _VC_CACHE[nombre]
    for _ in range(3):
        try:
            sql = (f"SELECT YEAR(FECHA_CIERRE) yr, MONTH(FECHA_CIERRE) mo, MAX(VALOR_CUOTA) vc "
                   f"FROM ODS.VALORES_CUOTA_GPI "
                   f"WHERE RTRIM(LTRIM(EMPRESA))=N'{nombre}' AND VALOR_CUOTA>0 "
                   f"GROUP BY YEAR(FECHA_CIERRE), MONTH(FECHA_CIERRE) ORDER BY yr, mo")
            rows = requests.post(_ODS_API, json={"Sql":sql},
                                 headers={"Content-Type":"application/json"}, timeout=30
                                 ).json().get("rows", [])
            result = {(int(x["yr"]),int(x["mo"])): float(x["vc"]) for x in rows}
            _VC_CACHE[nombre] = result; return result
        except Exception as e:
            import time; time.sleep(1)
    print(f"  [WARN] ODS failed: {nombre}"); return {}

# ── Template historico reader ─────────────────────────────────────────────────
def _get_tmpl_hist(tmpl_file):
    """Load {year: {label: [m1..m12 returns]}} from template rentabilidad sheet."""
    if tmpl_file in _HIST_CACHE: return _HIST_CACHE[tmpl_file]
    import openpyxl
    path = _TMPL_DIR / tmpl_file
    if not path.exists(): return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    if 'rentabilidad' not in wb.sheetnames:
        wb.close(); return {}
    ws  = wb['rentabilidad']
    out = {}
    cur_yr = None
    for r in range(6, ws.max_row+1):
        yr = ws.cell(r,3).value
        if isinstance(yr,(int,float)): cur_yr = int(yr)
        lbl = ws.cell(r,4).value
        if not (cur_yr and lbl and isinstance(lbl,str)): continue
        months = []
        for i in range(12):
            v = ws.cell(r, 5+i).value
            months.append(float(v) if v and v!=0 and isinstance(v,(int,float)) else None)
        if any(v is not None for v in months):
            out.setdefault(cur_yr, {})[lbl.strip()] = months
    wb.close()
    _HIST_CACHE[tmpl_file] = out
    return out

def _get_tmpl_summary(tmpl_file, is_usd):
    """
    For USD funds: read summary directly from template rentabilidad rows.
    Returns (acum_label, icp_row, comp_row, fip_row) as dicts.
    """
    import openpyxl
    path = _TMPL_DIR / tmpl_file
    if not path.exists(): return None, None, None, None
    wb = openpyxl.load_workbook(path, data_only=True)
    if 'rentabilidad' not in wb.sheetnames:
        wb.close(); return None, None, None, None
    ws = wb['rentabilidad']
    acum_label = str(ws.cell(1,24).value or '').replace('\n',' ').strip()

    def read_row(r):
        return {k: ws.cell(r,c).value
                for k,c in zip(['nombre','m','t','s','a','ac'],[19,20,21,22,23,24])}

    if is_usd:
        icp_row  = read_row(2)  # In USD templates R2=Comp (ICP-like)
        fip_row  = read_row(3)  # R3=FIP
        comp_row = read_row(2)  # Comp = same as ICP row for USD
        icp_row['nombre']  = 'ICP (Benchmark)'
        comp_row['nombre'] = 'Competencia'
    else:
        icp_row  = read_row(2)
        comp_row = read_row(3)
        fip_row  = read_row(4)
    wb.close()
    return acum_label, icp_row, comp_row, fip_row

def _fip_row_from_hist(hist_yr, hint):
    """Find the FIP monthly row in a template year dict."""
    skip = {'icp','competencia','icp (benchmark)'}
    candidates = [(k,v) for k,v in hist_yr.items() if k.lower() not in skip and 'icp' not in k.lower()]
    if not candidates: return None
    if len(candidates) == 1: return candidates[0][1]
    hint_up = hint.upper().replace('FIP','').replace('VANTRUST','').replace('LIQUIDEZ','').strip()
    for k,v in candidates:
        if hint_up in k.upper().replace('FIP','').replace('LIQUIDEZ','').strip():
            return v
    return candidates[0][1]

# ── Build CLP VC series (ODS + template chain for pre-ODS) ───────────────────
def _build_clp_vc(nombre, tmpl_file):
    vc_ods = _get_ods_vc(nombre)
    if not vc_ods: return {}
    if not tmpl_file: return vc_ods

    hist = _get_tmpl_hist(tmpl_file)
    if not hist: return vc_ods

    earliest = min(vc_ods.keys())
    combined = dict(vc_ods)
    cur_vc   = vc_ods[earliest]

    # Collect pre-ODS monthly returns in reverse order
    pre = sorted(
        ((yr, mo_idx+1, ret)
         for yr, yr_data in hist.items()
         for yr_data_row in [_fip_row_from_hist(yr_data, nombre)]
         if yr_data_row
         for mo_idx, ret in enumerate(yr_data_row)
         if ret is not None and (yr, mo_idx+1) < earliest),
        reverse=True
    )

    for yr, mo, ret in pre:
        cur_vc = cur_vc / (1.0 + ret)
        combined[(yr, mo)] = cur_vc

    # Also add one synthetic prior month so first FIP month can be calculated
    if pre:
        oldest_yr, oldest_mo, oldest_ret = pre[-1]
        py, pm = _prev(oldest_yr, oldest_mo, 1)
        if py not in combined or pm not in [k[1] for k in combined if k[0]==py]:
            combined[(py, pm)] = cur_vc / (1.0 + oldest_ret)

    return combined

# ── Return helpers ────────────────────────────────────────────────────────────
def _simple(vc, y, m, n=1):
    v1 = vc.get((y,m)); v0 = vc.get(_prev(y,m,n))
    return v1/v0-1 if v1 and v0 else None

def _annualized(vc, y, m, n=1):
    v1 = vc.get((y,m)); py,pm = _prev(y,m,n); v0 = vc.get((py,pm))
    if not v1 or not v0: return None
    days = (_eom(y,m) - _eom(py,pm)).days
    return (v1/v0-1)/days*360 if days>0 else None

def _ytd(vc, y, m):
    """Acum YYYY: (end/first_in_year - 1) / n_months * 12."""
    v_end = vc.get((y,m))
    if not v_end: return None
    v_first, n = None, 0
    for mm in range(1, m+1):
        v = vc.get((y,mm))
        if v:
            if v_first is None: v_first = v
            n += 1
    if not v_first or n==0 or v_first==v_end: return None
    return (v_end/v_first-1)/n*12

# ── Main entry point ──────────────────────────────────────────────────────────
def leer_datos_template(nombre_fondo, target_year=None, target_month=None):
    if target_year is None:
        tm = os.environ.get("TARGET_MONTH","").strip()
        if tm: target_year, target_month = map(int, tm.split("-"))
        else:
            hoy = date.today()
            prev = date(hoy.year, hoy.month, 1) - relativedelta(months=1)
            target_year, target_month = prev.year, prev.month
    y, m = target_year, target_month

    display   = NOMBRE_DISPLAY.get(nombre_fondo, nombre_fondo.replace('FIP VANTRUST ',''))
    is_usd    = nombre_fondo in FONDOS_USD
    tmpl_file = FUND_TEMPLATE_MAP.get(nombre_fondo)
    icp       = _get_icp_series()
    vc_comp   = _load_json("comp_clp.json")  # CLP Santander MM for all funds

    # ── For USD: read summary and historico directly from template ───────────
    use_tmpl = (is_usd or nombre_fondo in FONDOS_USE_TEMPLATE) and tmpl_file
    if use_tmpl:
        return _build_usd_output(nombre_fondo, display, tmpl_file, icp, vc_comp, y, m, is_usd=is_usd)

    # ── For CLP: ODS + template chain ────────────────────────────────────────
    vc_fip  = _build_clp_vc(nombre_fondo, tmpl_file)
    has_12  = bool(vc_fip.get(_prev(y, m, 12)))

    def calc(vc, es_icp, es_comp, es_fip, name):
        return {
            'nombre': name, 'es_icp': es_icp, 'es_comp': es_comp, 'es_fip': es_fip,
            'm':  _simple(vc, y, m, 1),   't': _simple(vc, y, m, 3),
            's':  _simple(vc, y, m, 6),   'a': _simple(vc, y, m, 12) if has_12 else None,
            'ac': _ytd(vc, y, m),
        }

    resumen = [
        calc(icp,     True,  False, False, 'ICP (Benchmark)'),
        calc(vc_comp, False, True,  False, 'Competencia'),
        calc(vc_fip,  False, False, True,  display),
    ]

    # Acum label from template
    acum_label = f"Acum. {y} (*)"
    if tmpl_file:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(_TMPL_DIR/tmpl_file, data_only=True)
            if 'rentabilidad' in wb.sheetnames:
                raw = wb['rentabilidad'].cell(1,24).value
                if raw: acum_label = str(raw).replace('\n',' ').strip()
            wb.close()
        except: pass

    # ── Historical table ─────────────────────────────────────────────────────
    tmpl_hist = _get_tmpl_hist(tmpl_file) if tmpl_file else {}
    historico = _build_clp_historico(y, m, icp, vc_comp, vc_fip, display, tmpl_hist, nombre_fondo)

    # ── Chart ────────────────────────────────────────────────────────────────
    labels, c_icp, c_comp, c_fip = _build_chart(y, m, icp, vc_comp, _get_ods_vc(nombre_fondo))

    return {'nombre_fip': display, 'acum_label': acum_label,
            'resumen': resumen, 'historico': historico,
            'grafico': {'labels': labels, 'icp': c_icp, 'comp': c_comp, 'fip': c_fip}}


def _build_clp_historico(y, m, icp, vc_comp, vc_fip, display, tmpl_hist, nombre_fondo):
    """Build historical table rows for CLP fund."""
    out = []
    for yr in [y-2, y-1, y]:
        last_m = m if yr==y else 12
        filas  = []
        for lbl, vc, is_icp_like in [('ICP (Benchmark)', icp, True),
                                      ('Competencia', vc_comp, False),
                                      (display, vc_fip, False)]:
            months_out = []
            for mm in range(1, 13):
                if mm > last_m:
                    months_out.append(None); continue
                # Check if ODS VC exists for this month (or it's ICP/Comp with extended series)
                if vc.get((yr, mm)) and vc.get((yr, mm-1) if mm>1 else _prev(yr,mm)):
                    months_out.append(_simple(vc, yr, mm))
                elif not is_icp_like and tmpl_hist:
                    # Pre-ODS: use template monthly returns directly
                    yr_data = tmpl_hist.get(yr, {})
                    fip_tmpl = _fip_row_from_hist(yr_data, nombre_fondo)
                    if fip_tmpl and mm <= len(fip_tmpl):
                        months_out.append(fip_tmpl[mm-1])
                    else:
                        months_out.append(_simple(vc, yr, mm))
                else:
                    months_out.append(_simple(vc, yr, mm))

            if any(v is not None for v in months_out):
                total = _ytd(vc, yr, last_m)
                filas.append({'nombre': lbl, 'meses': months_out, 'total': total})
        if filas:
            out.append({'año': yr, 'filas': filas})
    return out


def _build_usd_output(nombre_fondo, display, tmpl_file, icp, vc_comp, y, m, is_usd=True):
    """For USD funds: read everything from template (more accurate than ODS for these)."""
    import openpyxl
    path = _TMPL_DIR / tmpl_file
    wb   = openpyxl.load_workbook(path, data_only=True)
    ws   = wb['rentabilidad']

    acum_label = str(ws.cell(1,24).value or f'Acum. {y} (*)').replace('\n',' ').strip()

    def rw(r):
        return {k: ws.cell(r,c).value for k,c in zip(['nombre','m','t','s','a','ac'],[19,20,21,22,23,24])}

    if is_usd:
        # USD templates: R2=Comp/benchmark, R3=FIP, R4=empty
        r_icp, r_comp, r_fip = rw(2), rw(2), rw(3)
    else:
        # CLP templates: R2=ICP, R3=Comp, R4=FIP  
        r_icp, r_comp, r_fip = rw(2), rw(3), rw(4)

    icp_row  = {'nombre': 'ICP (Benchmark)', 'es_icp': True,  'es_comp': False, 'es_fip': False,
                'm': r_icp['m'], 't': r_icp['t'], 's': r_icp['s'], 'a': r_icp['a'], 'ac': r_icp['ac']}
    comp_row = {'nombre': 'Competencia',     'es_icp': False, 'es_comp': True,  'es_fip': False,
                'm': r_comp['m'], 't': r_comp['t'], 's': r_comp['s'], 'a': r_comp['a'], 'ac': r_comp['ac']}
    fip_row  = {'nombre': display,            'es_icp': False, 'es_comp': False, 'es_fip': True,
                'm': r_fip['m'], 't': r_fip['t'], 's': r_fip['s'], 'a': r_fip['a'], 'ac': r_fip['ac']}

    # Historical from template
    hist_raw = _get_tmpl_hist(tmpl_file)
    wb.close()

    historico = []
    for yr in [y-2, y-1, y]:
        last_m  = m if yr==y else 12
        yr_data = hist_raw.get(yr, {})
        filas   = []
        for lbl, is_icp_like in [('ICP (Benchmark)', True), ('Competencia', False), (display, False)]:
            # Find the right row in template hist
            row_months = None
            if is_icp_like:
                row_months = yr_data.get('ICP') or next(
                    (v for k,v in yr_data.items() if 'ICP' in k.upper()), None)
                if row_months is None:
                    # Use ICP series for historical ICP
                    row_months = [_simple(icp, yr, mm) for mm in range(1,13)]
            else:
                all_rows = [(k,v) for k,v in yr_data.items() if 'ICP' not in k.upper() and k.lower() not in {'competencia'}]
                comp_rows = [(k,v) for k,v in yr_data.items() if k.lower() == 'competencia']
                if lbl == 'Competencia':
                    row_months = comp_rows[0][1] if comp_rows else None
                else:
                    row_months = _fip_row_from_hist(yr_data, nombre_fondo)

            if row_months:
                months_out = [row_months[mm-1] if mm<=last_m and mm<=len(row_months) else None
                              for mm in range(1,13)]
                if any(v is not None for v in months_out):
                    # Compute total from the months
                    non_none = [v for v in months_out[:last_m] if v is not None]
                    total = None
                    if len(non_none) > 0:
                        # (1+r1)(1+r2)...(1+rn) - 1 = compound return
                        compound = 1.0
                        for r in non_none: compound *= (1+r)
                        total = compound - 1
                    filas.append({'nombre': lbl, 'meses': months_out, 'total': total})
        if filas:
            historico.append({'año': yr, 'filas': filas})

    # Chart from ODS
    vc_ods = _get_ods_vc(nombre_fondo)
    labels, c_icp, c_comp, c_fip = _build_chart(y, m, icp, vc_comp, vc_ods)

    return {'nombre_fip': display, 'acum_label': acum_label,
            'resumen': [icp_row, comp_row, fip_row],
            'historico': historico,
            'grafico': {'labels': labels, 'icp': c_icp, 'comp': c_comp, 'fip': c_fip}}


def _build_chart(y, m, icp, vc_comp, vc_ods):
    start = (y-2, 1); end = (y, m)
    months = sorted({k for k in list(vc_ods.keys())+list(icp.keys())+list(vc_comp.keys())
                     if start <= k <= end})
    if not months: return [], [], [], []
    base = months[0]
    b_f = next((v for k,v in sorted(vc_ods.items()) if k>=base), 1) or 1
    b_i = icp.get(base) or next((v for k,v in sorted(icp.items()) if k>=base), 1)
    b_c = vc_comp.get(base) or next((v for k,v in sorted(vc_comp.items()) if k>=base), 1)
    labels,ci,cc,cf = [],[],[],[]
    for k in months:
        vi = icp.get(k)
        if not vi: continue
        labels.append(f"{k[0]}-{k[1]:02d}")
        ci.append(round(vi/b_i*100, 4))
        vf = vc_ods.get(k); vc_c = vc_comp.get(k)
        cf.append(round(vf/b_f*100,4) if vf else None)
        cc.append(round(vc_c/b_c*100,4) if vc_c else None)
    return labels, ci, cc, cf

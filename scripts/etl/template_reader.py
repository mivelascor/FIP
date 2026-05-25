"""
etl/template_reader.py — Lee los datos directamente de las planillas Excel (fuente de verdad).

ARQUITECTURA:
1. Para cada fondo, busca su template en FUND_TEMPLATE_MAP
2. Lee el resumen (M/T/S/A/Acum) y el histórico directamente de la hoja 'rentabilidad'
3. Para meses más recientes que el template, extiende con datos de ODS
4. Si no hay template, calcula todo desde ODS

Los valores del template ya están calculados correctamente y coinciden con los PDF.
"""
import os, requests, json, calendar
from pathlib import Path
from datetime import date
from dateutil.relativedelta import relativedelta
import pandas as pd

_INPUTS_DIR = Path(__file__).parent.parent.parent / "inputs"
_TMPL_DIR   = _INPUTS_DIR / "templates"

# ── Mapa: ODS fund name → template file que contiene sus datos ───────────────
FUND_TEMPLATE_MAP = {
    'FIP VANTRUST LIQUIDEZ ACTIVA':        'TEMPLATE_FONDO_LIQUIDEZ_CAJA.xlsx',
    'FIP VANTRUST LIQUIDEZ ALTO MONTO':    'TEMPLATE_FONDO_LIQUIDEZ_CONTINUA.xlsx',
    'FIP VANTRUST LIQUIDEZ CAJA':          'TEMPLATE_FONDO_LIQUIDEZ_CORRIENTE.xlsx',
    'FIP VANTRUST LIQUIDEZ CONTINUA':      'TEMPLATE_FONDO_LIQUIDEZ_CORTO_PLAZO.xlsx',
    'FIP VANTRUST LIQUIDEZ CORRIENTE':     'TEMPLATE_FONDO_LIQUIDEZ_Disponible_I.xlsx',
    'FIP VANTRUST LIQUIDEZ DISPONIBLE I':  'TEMPLATE_FONDO_LIQUIDEZ_DOLAR_CAJA.xlsx',
    'FIP VANTRUST LIQUIDEZ EFECTIVO':      'TEMPLATE_FONDO_LIQUIDEZ_LOCAL.xlsx',
    'FIP VANTRUST LIQUIDEZ FLEXIBLE':      'TEMPLATE_FONDO_LIQUIDEZ_Monetario_I.xlsx',
    'FIP VANTRUST LIQUIDEZ MONETARIO I':   'TEMPLATE_FONDO_LIQUIDEZ_Permanente.xlsx',
    'FIP VANTRUST LIQUIDEZ PERMANENTE':    'TEMPLATE_FONDO_LIQUIDEZ_Presente.xlsx',
    'FIP VANTRUST LIQUIDEZ PLUS':          'TEMPLATE_FONDO_LIQUIDEZ_RENDIMIENTO.xlsx',
    'FIP VANTRUST LIQUIDEZ PRESENTE':      'TEMPLATE_FONDO_LIQUIDEZ_RESERVA_DOLAR.xlsx',
    'FIP VANTRUST LIQUIDEZ RENDIMIENTO':   'TEMPLATE_FONDO_LIQUIDEZ_SENCILLO.xlsx',
    # USD funds
    'FIP VANTRUST LIQUIDEZ DOLAR CAJA':    'TEMPLATE_FONDO_LIQUIDEZ_EFECTIVO.xlsx',
    'FIP VANTRUST LIQUIDEZ DOLAR':         'TEMPLATE_FONDO_LIQUIDEZ_FLEXIBLE.xlsx',
    'FIP VANTRUST LIQUIDEZ RESERVA DÓLAR': 'TEMPLATE_FONDO_LIQUIDEZ_UNO.xlsx',
}

# Fondos que son USD (usan retornos anualizados en el resumen)
FONDOS_USD = {
    'FIP VANTRUST LIQUIDEZ DOLAR',
    'FIP VANTRUST LIQUIDEZ DOLAR CAJA',
    'FIP VANTRUST LIQUIDEZ RESERVA DÓLAR',
}

# Nombre display para cada fondo ODS
NOMBRE_DISPLAY = {
    'FIP VANTRUST LIQUIDEZ ACTIVA':        'FIP Liquidez Activa',
    'FIP VANTRUST LIQUIDEZ ALTO APORTE':   'FIP Alto Aporte',
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

_ICP_CACHE:    dict = {}
_VC_CACHE:     dict = {}
_ODS_API = "https://claudeods.vantrustcapital.cl/query"

# ── Helpers ───────────────────────────────────────────────────────────────────
def _load_json(filename):
    try:
        with open(_INPUTS_DIR / filename) as f:
            return {(int(k[:4]), int(k[5:7])): float(v)
                    for k, v in json.load(f).items()}
    except Exception as e:
        print(f"  [WARN] {filename}: {e}")
        return {}

def _eom(y, m):
    return date(y, m, calendar.monthrange(y, m)[1])

def _prev(y, m, n=1):
    d = date(y, m, 1) - relativedelta(months=n)
    return d.year, d.month

def _simple_ret(vc, y, m, n=1):
    """Simple return (CLP): (end/start)-1"""
    v1 = vc.get((y, m))
    v0 = vc.get(_prev(y, m, n))
    return v1/v0 - 1 if v1 and v0 else None

def _annualized_ret(vc, y, m, n=1):
    """Annualized return (USD): (end/start-1)/days*360"""
    v1 = vc.get((y, m))
    py, pm = _prev(y, m, n)
    v0 = vc.get((py, pm))
    if not v1 or not v0: return None
    days = (_eom(y, m) - _eom(py, pm)).days
    return (v1/v0 - 1) / days * 360 if days > 0 else None

def _year_total(vc, y, last_m):
    """(end/first_in_year - 1) / count * 12"""
    v_end = vc.get((y, last_m))
    if not v_end: return None
    v_first, n = None, 0
    for mm in range(1, last_m+1):
        v = vc.get((y, mm))
        if v:
            if v_first is None: v_first = v
            n += 1
    if not v_first or n == 0 or v_first == v_end: return None
    return (v_end / v_first - 1) / n * 12

# ── ODS data fetching ─────────────────────────────────────────────────────────
def _get_ods_vc(nombre_fondo: str) -> dict:
    """Query ODS for end-of-month VC values."""
    if nombre_fondo in _VC_CACHE:
        return _VC_CACHE[nombre_fondo]
    for attempt in range(3):
        try:
            sql = (f"SELECT YEAR(FECHA_CIERRE) yr, MONTH(FECHA_CIERRE) mo, MAX(VALOR_CUOTA) vc "
                   f"FROM ODS.VALORES_CUOTA_GPI "
                   f"WHERE RTRIM(LTRIM(EMPRESA))=N'{nombre_fondo}' AND VALOR_CUOTA>0 "
                   f"GROUP BY YEAR(FECHA_CIERRE), MONTH(FECHA_CIERRE) ORDER BY yr, mo")
            r = requests.post(_ODS_API, json={"Sql": sql},
                              headers={"Content-Type": "application/json"}, timeout=30)
            rows = r.json().get("rows", [])
            result = {(int(x["yr"]), int(x["mo"])): float(x["vc"]) for x in rows}
            _VC_CACHE[nombre_fondo] = result
            return result
        except Exception as e:
            import time; time.sleep(1)
            if attempt == 2:
                print(f"  [WARN] ODS {nombre_fondo}: {e}")
    return {}

def _get_icp_series() -> dict:
    """ICP from CLICP file + TPM extension for future months."""
    global _ICP_CACHE
    if _ICP_CACHE: return _ICP_CACHE
    clicp = _load_json("icp_clicp.json")
    if clicp:
        last_ym  = max(clicp.keys())
        last_val = clicp[last_ym]
        _extend_with_tpm(clicp, last_ym, last_val)
        _ICP_CACHE = clicp
        return _ICP_CACHE
    # Fallback: pure TPM
    icp = {}
    _extend_with_tpm(icp, (2005, 12), 10000.0)
    _ICP_CACHE = icp
    return _ICP_CACHE

def _extend_with_tpm(series, after_ym, base_val):
    today = date.today()
    prev  = base_val
    for year in range(after_ym[0], today.year + 1):
        try:
            r     = requests.get(f"https://mindicador.cl/api/tpm/{year}", timeout=15)
            items = r.json().get("serie", [])
            df    = pd.DataFrame([{"fecha": pd.to_datetime(i["fecha"]), "tpm": float(i["valor"])}
                                   for i in items])
            if df.empty: continue
            df["ym"] = df["fecha"].dt.to_period("M")
            m_tpm   = df.groupby("ym")["tpm"].mean().reset_index()
            for _, row in m_tpm.iterrows():
                ym = (row["ym"].year, row["ym"].month)
                if ym <= after_ym: continue
                prev = prev * (1 + row["tpm"] / 1200)
                series[ym] = prev
        except: pass

# ── Read from template rentabilidad sheet ─────────────────────────────────────
def _read_template(tmpl_path: Path, fund_is_usd: bool = False) -> dict:
    """
    Returns dict with keys:
      is_usd: bool
      icp_summary:  {m, t, s, a, ac}   (values as floats, e.g. 0.0038)
      comp_summary: {m, t, s, a, ac}
      fip_summary:  {m, t, s, a, ac}
      acum_label:   str
      fip_name:     str
      historico:    {year: {row_name: {months:[v1..v12], total:v}}}
      last_year:    int
      last_month:   int  (last month with FIP data in template)
    """
    import openpyxl
    wb = openpyxl.load_workbook(tmpl_path, data_only=True)
    if 'rentabilidad' not in wb.sheetnames:
        wb.close(); return None
    ws = wb['rentabilidad']

    def row_to_summary(r):
        return {
            'name': ws.cell(r, 19).value,
            'm':  ws.cell(r, 20).value,
            't':  ws.cell(r, 21).value,
            's':  ws.cell(r, 22).value,
            'a':  ws.cell(r, 23).value,
            'ac': ws.cell(r, 24).value,
        }

    acum_label = str(ws.cell(1, 24).value or '').replace('\n', ' ').strip()
    is_usd     = fund_is_usd

    # Summary rows
    if is_usd:
        # USD: row2=Comp (benchmark), row3=FIP
        comp_sum = row_to_summary(2)
        fip_sum  = row_to_summary(3)
        icp_sum  = comp_sum.copy()  # use Comp as ICP for USD (same benchmark)
        icp_sum['name'] = 'ICP (Benchmark)'
        fip_name = str(ws.cell(3, 19).value or '').strip()
    else:
        icp_sum  = row_to_summary(2)
        comp_sum = row_to_summary(3)
        fip_sum  = row_to_summary(4)
        fip_name = str(ws.cell(4, 19).value or '').strip()

    # Historical table
    historico  = {}
    cur_year   = None
    last_year  = None
    last_month = 0

    for r in range(6, ws.max_row + 1):
        yr = ws.cell(r, 3).value
        if isinstance(yr, (int, float)):
            cur_year = int(yr)

        row_name = ws.cell(r, 4).value
        if not (cur_year and row_name and isinstance(row_name, str)):
            continue

        months = [ws.cell(r, 5+i).value for i in range(12)]
        total  = ws.cell(r, 17).value

        has_data = any(v and v != 0 and isinstance(v, (int, float)) for v in months)
        if not has_data: continue

        if cur_year not in historico:
            historico[cur_year] = {}

        # Normalize row name
        rn_norm = row_name.strip()
        historico[cur_year][rn_norm] = {
            'months': [v if (v and v != 0 and isinstance(v, (int,float))) else None for v in months],
            'total':  float(total) if total and total != 0 and isinstance(total,(int,float)) else None,
        }

        # Track last FIP data point
        if rn_norm.upper().replace('FIP','').strip() in (fip_name.upper().replace('FIP','').strip(),):
            for i, v in enumerate(months):
                if v and v != 0:
                    last_year  = cur_year
                    last_month = i + 1

    wb.close()

    # Find the last month with FIP data (use all rows to find the actual last point)
    for yr in sorted(historico.keys(), reverse=True):
        for rn, data in historico[yr].items():
            if 'ICP' in rn.upper() or 'COMP' in rn.upper() or rn.upper() == 'COMPETENCIA':
                continue
            # This is a FIP row
            for i in range(11, -1, -1):  # scan right to left
                if data['months'][i] is not None:
                    last_year  = yr
                    last_month = i + 1
                    break
            if last_year:
                break
        if last_year:
            break
    
    # Fallback: use ICP to determine at minimum what year/month the template covers
    if not last_year:
        for yr in sorted(historico.keys(), reverse=True):
            for rn, data in historico[yr].items():
                if 'ICP' in rn.upper():
                    for i in range(11, -1, -1):
                        if data['months'][i] is not None:
                            last_year  = yr
                            last_month = i + 1
                            break
                    break
            if last_year:
                break

    return {
        'is_usd':      is_usd,
        'icp_summary': icp_sum,
        'comp_summary': comp_sum,
        'fip_summary': fip_sum,
        'acum_label':  acum_label,
        'fip_name':    fip_name,
        'historico':   historico,
        'last_year':   last_year or 2026,
        'last_month':  last_month,
    }

# ── Build output dict from template + ODS extension ───────────────────────────
def leer_datos_template(nombre_fondo: str,
                         target_year:  int = None,
                         target_month: int = None) -> dict:
    # Resolve target period
    if target_year is None:
        tm = os.environ.get("TARGET_MONTH", "")
        if tm:
            target_year, target_month = map(int, tm.split("-"))
        else:
            hoy  = date.today()
            prev = date(hoy.year, hoy.month, 1) - relativedelta(months=1)
            target_year, target_month = prev.year, prev.month
    y, m = target_year, target_month

    nombre_display = NOMBRE_DISPLAY.get(nombre_fondo, nombre_fondo.replace('FIP VANTRUST ', ''))
    is_usd         = nombre_fondo in FONDOS_USD
    icp            = _get_icp_series()
    vc_comp        = _load_json("comp_clp.json") if not is_usd else _load_json("comp_usd.json")
    vc_fip_ods     = _get_ods_vc(nombre_fondo)

    # ── Try to read from template ────────────────────────────────────────────
    tmpl_file = FUND_TEMPLATE_MAP.get(nombre_fondo)
    tmpl_data = None
    if tmpl_file:
        tmpl_path = _TMPL_DIR / tmpl_file
        if tmpl_path.exists():
            tmpl_data = _read_template(tmpl_path, fund_is_usd=is_usd)

    # ── Determine if template is current or needs ODS extension ─────────────
    tmpl_is_current = (
        tmpl_data and
        tmpl_data['last_year'] == y and
        tmpl_data['last_month'] == m
    )

    # ── Build SUMMARY (resumen) ──────────────────────────────────────────────
    acum_label = f"Acum. {y} (*)"
    if tmpl_data and tmpl_data['acum_label']:
        acum_label = tmpl_data['acum_label']

    if tmpl_is_current:
        # Use template summary directly
        def make_row(name, s, es_icp, es_comp, es_fip):
            return {
                'nombre': name, 'es_icp': es_icp, 'es_comp': es_comp, 'es_fip': es_fip,
                'm':  s['m'],   't':  s['t'],   's':  s['s'],
                'a':  s['a'],   'ac': s['ac'],
            }
        icp_name  = tmpl_data['icp_summary'].get('name') or 'ICP (Benchmark)'
        comp_name = tmpl_data['comp_summary'].get('name') or 'Competencia'
        fip_name  = tmpl_data['fip_name'] or nombre_display
        resumen = [
            make_row(icp_name,  tmpl_data['icp_summary'],  True,  False, False),
            make_row(comp_name, tmpl_data['comp_summary'],  False, True,  False),
            make_row(fip_name,  tmpl_data['fip_summary'],   False, False, True),
        ]
    else:
        # Calculate from ODS / series
        vc_fip = vc_fip_ods
        ret    = _annualized_ret if is_usd else _simple_ret
        has_12 = bool(vc_fip.get(_prev(y, m, 12)))

        def calc_row(name, vc, es_icp, es_comp, es_fip):
            ytd = _year_total(vc, y, m)
            return {
                'nombre': name, 'es_icp': es_icp, 'es_comp': es_comp, 'es_fip': es_fip,
                'm':  ret(vc, y, m, 1),
                't':  ret(vc, y, m, 3),
                's':  ret(vc, y, m, 6),
                'a':  ret(vc, y, m, 12) if has_12 else None,
                'ac': ytd,
            }
        resumen = [
            calc_row('ICP (Benchmark)', icp,     True,  False, False),
            calc_row('Competencia',     vc_comp, False, True,  False),
            calc_row(nombre_display,    vc_fip,  False, False, True),
        ]

    # ── Build HISTORICO ──────────────────────────────────────────────────────
    HIST_YEARS = [y-2, y-1, y]
    historico_out = []

    if tmpl_is_current and tmpl_data:
        # Use template historical directly
        for yr in HIST_YEARS:
            yr_data = tmpl_data['historico'].get(yr, {})
            if not yr_data:
                # Year not in template — build from series
                yr_data = _build_hist_year_from_vc(yr, y, m, icp, vc_comp, vc_fip_ods, nombre_display)
            if yr_data:
                historico_out.append(_format_hist_year(yr, yr_data, y, m, tmpl_data['fip_name'], nombre_display))
    else:
        # Full ODS calculation
        for yr in HIST_YEARS:
            yr_data = _build_hist_year_from_vc(yr, y, m, icp, vc_comp, vc_fip_ods, nombre_display)
            if yr_data:
                historico_out.append(_format_hist_year(yr, yr_data, y, m, None, nombre_display))

    # ── Build chart data ─────────────────────────────────────────────────────
    chart_start = (y-2, 1)
    chart_end   = (y, m)

    # Build normalized index from VC series
    all_months = sorted(set(
        k for k in list(vc_fip_ods.keys()) + list(icp.keys()) + list(vc_comp.keys())
        if chart_start <= k <= chart_end
    ))
    if all_months:
        base = all_months[0]
        b_fip  = vc_fip_ods.get(base)  or next((v for k,v in sorted(vc_fip_ods.items())  if k >= base), 1)
        b_icp  = icp.get(base)         or next((v for k,v in sorted(icp.items())         if k >= base), 1)
        b_comp = vc_comp.get(base)     or next((v for k,v in sorted(vc_comp.items())     if k >= base), 1)

        labels, c_icp, c_comp, c_fip = [], [], [], []
        for k in all_months:
            vi = icp.get(k)
            if not vi: continue
            vf = vc_fip_ods.get(k)
            vc = vc_comp.get(k)
            labels.append(f"{k[0]}-{k[1]:02d}")
            c_icp.append(round(vi / b_icp * 100, 4))
            c_fip.append(round(vf / b_fip * 100, 4) if vf else None)
            c_comp.append(round(vc / b_comp * 100, 4) if vc else None)
    else:
        labels = c_icp = c_comp = c_fip = []

    return {
        'nombre_fip':  nombre_display,
        'acum_label':  acum_label,
        'resumen':     resumen,
        'historico':   historico_out,
        'grafico':     {'labels': labels, 'icp': c_icp, 'comp': c_comp, 'fip': c_fip},
    }


def _build_hist_year_from_vc(yr, target_y, target_m, icp, vc_comp, vc_fip, fip_display):
    """Build historical year rows from VC series."""
    last_m = target_m if yr == target_y else 12
    result = {}
    for label, vc in [('ICP', icp), ('Competencia', vc_comp), (fip_display, vc_fip)]:
        months = []
        for mm in range(1, 13):
            if mm > last_m:
                months.append(None)
                continue
            r = _simple_ret(vc, yr, mm)
            months.append(r)
        if any(v is not None for v in months):
            total = _year_total(vc, yr, last_m)
            result[label] = {'months': months, 'total': total}
    return result if result else None


def _format_hist_year(yr, yr_data, target_y, target_m, tmpl_fip_name, display_name):
    """Convert raw year data to output format."""
    last_m = target_m if yr == target_y else 12
    filas  = []
    for row_name, data in yr_data.items():
        months  = data['months']
        total   = data['total']
        # Pad to 12
        while len(months) < 12:
            months.append(None)
        months = [months[i] if i < len(months) else None for i in range(12)]
        # Zero out months after last_m
        for i in range(last_m, 12):
            months[i] = None

        # Determine display name for FIP row
        display = row_name
        if tmpl_fip_name and row_name.strip().upper() == tmpl_fip_name.strip().upper():
            display = display_name

        filas.append({'nombre': display, 'meses': months, 'total': total})
    return {'año': yr, 'filas': filas}

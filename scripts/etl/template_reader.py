"""
etl/template_reader.py — Calcula rentabilidades directamente desde SQL + ICP.
Sin dependencia de LibreOffice ni recalc.
"""
import os, requests, pandas as pd, json
from pathlib import Path
from datetime import date
from dateutil.relativedelta import relativedelta

#  Historical fallback (pre-ODS monthly returns from manual folletos) 
_HISTORICO_PATH = Path(__file__).parent.parent.parent / "inputs" / "historico_fondos.json"
_HIST_CACHE: dict = {}

def _load_historico() -> dict:
    """Load pre-ODS historical monthly returns keyed by fund name."""
    global _HIST_CACHE
    if _HIST_CACHE:
        return _HIST_CACHE
    try:
        import json as _json
        with open(_HISTORICO_PATH) as f:
            raw = _json.load(f)
        for fund, years in raw.items():
            _HIST_CACHE[fund] = {}
            for yr_str, months in years.items():
                _HIST_CACHE[fund][int(yr_str)] = {int(m): float(v) for m, v in months.items()}
    except Exception:
        pass
    return _HIST_CACHE


def _get_vc_combined(nombre_fondo: str) -> dict:
    """
    Return VC time-series merging ODS data with pre-ODS historical fallback.
    Chains VC values backwards from earliest ODS point using historical monthly returns.
    Also adds a synthetic prior-month VC so the earliest historical month's
    return can be calculated.
    """
    vc_ods = _get_vc_fondo_robust(nombre_fondo)
    hist   = _load_historico().get(nombre_fondo, {})

    if not hist:
        return vc_ods

    combined = dict(vc_ods)

    if vc_ods:
        earliest = min(vc_ods.keys())
        base_vc  = vc_ods[earliest]

        # All historical (yr, mo, ret) before ODS start, sorted descending
        hist_list = sorted(
            [(yr, mo, ret)
             for yr, months in hist.items()
             for mo, ret in months.items()
             if (yr, mo) < earliest],
            reverse=True
        )

        cur_vc = base_vc
        for yr, mo, ret in hist_list:
            if ret:
                prev_vc = cur_vc / (1.0 + ret)
                combined[(yr, mo)] = prev_vc
                cur_vc = prev_vc

        # Add synthetic baseline for the earliest historical month
        # so _ret() can compute that month's return
        if hist_list:
            last_yr, last_mo, last_ret = hist_list[-1]  # oldest month
            if last_ret and combined.get((last_yr, last_mo)):
                # prev = current / (1 + ret) already computed above
                # We need one more step back: add (yr, mo-1) as synthetic base
                from datetime import date
                from dateutil.relativedelta import relativedelta as _rdelta
                prev_date = date(last_yr, last_mo, 1) - _rdelta(months=1)
                prev_key  = (prev_date.year, prev_date.month)
                if prev_key not in combined:
                    combined[prev_key] = combined[(last_yr, last_mo)] / (1.0 + last_ret)

    return combined


API_SQL    = "https://claudeods.vantrustcapital.cl/query"
MINDICADOR = "https://mindicador.cl/api/tpm"

# Nombres exactos en la DB (con tildes si las tienen)
TEMPLATE_MAP = {
    "FIP VANTRUST LIQUIDEZ ACTIVA":        "TEMPLATE_FONDO_LIQUIDEZ_ACTIVA.xlsx",
    "FIP VANTRUST LIQUIDEZ ALTO APORTE":   "TEMPLATE_FONDO_ALTO_APORTE.xlsx",
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL":  "TEMPLATE_FONDO_ALTO_CAPITAL.xlsx",
    "FIP VANTRUST LIQUIDEZ ALTO MONTO":    "TEMPLATE_FONDO_LIQUIDEZ_ALTO_MONTO.xlsx",
    "FIP VANTRUST LIQUIDEZ CAJA":          "TEMPLATE_FONDO_LIQUIDEZ_CAJA.xlsx",
    "FIP VANTRUST LIQUIDEZ CONTINUA":      "TEMPLATE_FONDO_LIQUIDEZ_CONTINUA.xlsx",
    "FIP VANTRUST LIQUIDEZ CORRIENTE":     "TEMPLATE_FONDO_LIQUIDEZ_CORRIENTE.xlsx",
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO":   "TEMPLATE_FONDO_LIQUIDEZ_CORTO_PLAZO.xlsx",
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I":  "TEMPLATE_FONDO_LIQUIDEZ_Disponible_I.xlsx",
    "FIP VANTRUST LIQUIDEZ DOLAR":         "TEMPLATE_FONDO_LIQUIDEZ_DOLAR.xlsx",
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA":    "TEMPLATE_FONDO_LIQUIDEZ_DOLAR_CAJA.xlsx",
    "FIP VANTRUST LIQUIDEZ EFECTIVO":      "TEMPLATE_FONDO_LIQUIDEZ_EFECTIVO.xlsx",
    "FIP VANTRUST LIQUIDEZ FLEXIBLE":      "TEMPLATE_FONDO_LIQUIDEZ_FLEXIBLE.xlsx",
    "FIP VANTRUST LIQUIDEZ I":             "TEMPLATE_FONDO_LIQUIDEZ_UNO.xlsx",
    "FIP VANTRUST LIQUIDEZ LOCAL":         "TEMPLATE_FONDO_LIQUIDEZ_LOCAL.xlsx",
    "FIP VANTRUST LIQUIDEZ MONETARIO I":   "TEMPLATE_FONDO_LIQUIDEZ_Monetario_I.xlsx",
    "FIP VANTRUST LIQUIDEZ PERMANENTE":    "TEMPLATE_FONDO_LIQUIDEZ_Permanente.xlsx",
    "FIP VANTRUST LIQUIDEZ PLUS":          "TEMPLATE_FONDO_LIQUIDEZ_PLUS.xlsx",
    "FIP VANTRUST LIQUIDEZ PRESENTE":      "TEMPLATE_FONDO_LIQUIDEZ_Presente.xlsx",
    "FIP VANTRUST LIQUIDEZ RECURRENTE":    "TEMPLATE_FONDO_LIQUIDEZ.xlsx",
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO":   "TEMPLATE_FONDO_LIQUIDEZ_RENDIMIENTO.xlsx",
    "FIP VANTRUST LIQUIDEZ RESERVA DÓLAR": "TEMPLATE_FONDO_LIQUIDEZ_RESERVA_DOLAR.xlsx",
    "FIP VANTRUST LIQUIDEZ SENCILLO":      "TEMPLATE_FONDO_LIQUIDEZ_SENCILLO.xlsx",
    "FIP VANTRUST LIQUIDEZ TEMPORAL":      "TEMPLATE_FONDO_USD_MONEY_MARKET.xlsx",
}

# Nombre corto para mostrar en folleto
NOMBRE_DISPLAY = {
    "FIP VANTRUST LIQUIDEZ ACTIVA":        "FIP Liquidez Activa",
    "FIP VANTRUST LIQUIDEZ ALTO APORTE":   "FIP Alto Aporte",
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL":  "FIP Alto Capital",
    "FIP VANTRUST LIQUIDEZ ALTO MONTO":    "FIP Liquidez Alto Monto",
    "FIP VANTRUST LIQUIDEZ CAJA":          "FIP Liquidez Caja",
    "FIP VANTRUST LIQUIDEZ CONTINUA":      "FIP Liquidez Continua",
    "FIP VANTRUST LIQUIDEZ CORRIENTE":     "FIP Liquidez Corriente",
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO":   "FIP Liquidez Corto Plazo",
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I":  "FIP Liquidez Disponible I",
    "FIP VANTRUST LIQUIDEZ DOLAR":         "FIP Liquidez Dólar",
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA":    "FIP Liquidez Dólar Caja",
    "FIP VANTRUST LIQUIDEZ EFECTIVO":      "FIP Liquidez Efectivo",
    "FIP VANTRUST LIQUIDEZ FLEXIBLE":      "FIP Liquidez Flexible",
    "FIP VANTRUST LIQUIDEZ I":             "FIP Liquidez I",
    "FIP VANTRUST LIQUIDEZ LOCAL":         "FIP Liquidez Local",
    "FIP VANTRUST LIQUIDEZ MONETARIO I":   "FIP Liquidez Monetario I",
    "FIP VANTRUST LIQUIDEZ PERMANENTE":    "FIP Liquidez Permanente",
    "FIP VANTRUST LIQUIDEZ PLUS":          "FIP Liquidez Plus",
    "FIP VANTRUST LIQUIDEZ PRESENTE":      "FIP Liquidez Presente",
    "FIP VANTRUST LIQUIDEZ RECURRENTE":    "FIP Liquidez Recurrente",
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO":   "FIP Liquidez Rendimiento",
    "FIP VANTRUST LIQUIDEZ RESERVA DÓLAR": "FIP Liquidez Reserva Dólar",
    "FIP VANTRUST LIQUIDEZ SENCILLO":      "FIP Liquidez Sencillo",
    "FIP VANTRUST LIQUIDEZ TEMPORAL":      "FIP Liquidez Temporal",
}

FONDOS_USD = {
    "FIP VANTRUST LIQUIDEZ DOLAR",
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA",
    "FIP VANTRUST LIQUIDEZ RESERVA DÓLAR",
    "FIP VANTRUST LIQUIDEZ TEMPORAL",
_INPUTS_DIR = Path(__file__).parent.parent.parent / "inputs"


def _load_vc_series(filename: str) -> dict:
    """Load {(year, month): float} from JSON file."""
    try:
        with open(_INPUTS_DIR / filename) as f:
            raw = json.load(f)
        return {(int(k[:4]), int(k[5:7])): float(v) for k, v in raw.items()}
    except Exception as e:
        print(f"  [WARN] Could not load {filename}: {e}")
        return {}


VC_COMP_CLP = _load_vc_series("comp_clp.json")
VC_COMP_USD = _load_vc_series("comp_usd.json")
_ICP_CACHE: dict = {}
_VC_CACHE:  dict = {}


def _get_icp_series() -> dict:
    """Load ICP from CLICP template file, extended forward with TPM for new months."""
    global _ICP_CACHE
    if _ICP_CACHE:
        return _ICP_CACHE

    # Load CLICP from templates (matches folleto template data exactly)
    clicp = _load_vc_series("icp_clicp.json")
    if not clicp:
        # Full fallback: TPM-based approximation
        clicp = _build_icp_from_tpm(base_ym=None, base_val=None)
        _ICP_CACHE = clicp
        return _ICP_CACHE

    # Extend forward for months after the template file
    last_ym  = max(clicp.keys())
    last_val = clicp[last_ym]
    extra = _build_icp_from_tpm(base_ym=last_ym, base_val=last_val)
    clicp.update(extra)
    _ICP_CACHE = clicp
    return _ICP_CACHE


def _build_icp_from_tpm(base_ym, base_val):
    """Build ICP index continuation from TPM rates for months after base_ym."""
    result = {}
    today = date.today()
    start_year = base_ym[0] if base_ym else 2018
    prev_val = base_val if base_val else 10000.0

    for year in range(start_year, today.year + 1):
        try:
            r = requests.get(f"https://mindicador.cl/api/tpm/{year}", timeout=15)
            items = r.json().get("serie", [])
            df = pd.DataFrame([{"fecha": pd.to_datetime(i["fecha"]), "tpm": float(i["valor"])}
                                for i in items])
            if df.empty:
                continue
            df["ym"] = df["fecha"].dt.to_period("M")
            m_tpm = df.groupby("ym")["tpm"].mean().reset_index()
            for _, row in m_tpm.iterrows():
                ym = (row["ym"].year, row["ym"].month)
                if base_ym and ym <= base_ym:
                    continue
                prev_val = prev_val * (1 + row["tpm"] / 1200)
                result[ym] = prev_val
        except Exception as e:
            print(f"  [WARN] TPM fetch {year}: {e}")
    return result


def _get_vc_fondo(nombre_fondo: str) -> dict:
    global _VC_CACHE
    if nombre_fondo in _VC_CACHE:
        return _VC_CACHE[nombre_fondo]
    sql = (f"SELECT YEAR(FECHA_CIERRE) AS yr, MONTH(FECHA_CIERRE) AS mo, "
           f"MAX(VALOR_CUOTA) AS vc "
           f"FROM ODS.VALORES_CUOTA_GPI "
           f"WHERE RTRIM(LTRIM(EMPRESA))=N'{nombre_fondo}' AND VALOR_CUOTA>0 "
           f"GROUP BY YEAR(FECHA_CIERRE), MONTH(FECHA_CIERRE) "
           f"ORDER BY yr, mo")
    try:
        r = requests.post(API_SQL, json={"Sql": sql},
                          headers={"Content-Type": "application/json"}, timeout=60)
        rows = r.json().get("rows", [])
        result = {(int(row["yr"]), int(row["mo"])): float(row["vc"])
                  for row in rows if row.get("vc")}
        _VC_CACHE[nombre_fondo] = result
        return result
    except Exception as e:
        print(f"    [WARN] SQL ({nombre_fondo}): {e}")
        _VC_CACHE[nombre_fondo] = {}
        return {}


#  Rentabilidad helpers 
def _prev(y, m, n=1):
    d = date(y, m, 1) - relativedelta(months=n)
    return (d.year, d.month)

def _eom(y, m):
    """End-of-month date for (year, month)."""
    last_day = calendar.monthrange(y, m)[1]
    return datetime.date(y, m, last_day)

def _ret(vc, y, m, n=1, is_usd=False):
    v1 = vc.get((y, m))
    py, pm = _prev(y, m, n)
    v0 = vc.get((py, pm))
    if not v1 or not v0:
        return None
    simple = v1/v0 - 1
    if not is_usd:
        return simple
    import calendar as _cal
    from datetime import date as _d
    d1 = _d(y, m, _cal.monthrange(y, m)[1])
    d0 = _d(py, pm, _cal.monthrange(py, pm)[1])
    days = (d1 - d0).days
    return simple / days * 360 if days > 0 else None

def _ytd(vc, y, m):
    return _year_total(vc, y, m)

def _year_total(vc, y, last_m):
    vc_end = vc.get((y, last_m))
    if not vc_end:
        return None
    vc_first, n = None, 0
    for mm in range(1, last_m + 1):
        v = vc.get((y, mm))
        if v:
            if vc_first is None:
                vc_first = v
            n += 1
    if not vc_first or n == 0 or vc_first == vc_end:
        return None
    return (vc_end / vc_first - 1) / n * 12

def leer_datos_template(nombre_fondo: str,
                         target_year: int = None,
                         target_month: int = None) -> dict:
    # Resolve target period
    if target_year is None:
        tm = os.environ.get("TARGET_MONTH", "")
        if tm:
            target_year, target_month = map(int, tm.split("-"))
        else:
            hoy = date.today()
            prev = date(hoy.year, hoy.month, 1) - relativedelta(months=1)
            target_year, target_month = prev.year, prev.month

    y, m = target_year, target_month
    nombre_display = NOMBRE_DISPLAY.get(nombre_fondo,
                     nombre_fondo.replace("FIP VANTRUST ", "FIP "))
    is_usd     = nombre_fondo in FONDOS_USD
    vc_comp    = VC_COMP_USD if is_usd else VC_COMP_CLP

    icp    = _get_icp_series()
    vc_fip = _get_vc_combined(nombre_fondo)

    if not vc_fip:
        return {"acum_label": f"Acum. {y} (*)", "nombre_fip": nombre_display,
                "resumen": [], "historico": [],
                "grafico": {"labels":[], "icp":[], "comp":[], "fip":[]}}

    #  Resumen 
    is_usd = any(x in nombre_fondo.upper() for x in ("DOLAR", "DÓLAR", "USD", "DOLLAR"))

    def row(nombre, vc, es_icp, es_comp, es_fip):
        has_12 = bool(vc.get(_prev(y, m, 12)))
        return {"nombre": nombre,
                "m": _ret(vc,y,m,1,is_usd), "t": _ret(vc,y,m,3,is_usd),
                "s": _ret(vc,y,m,6,is_usd), "a": _ret(vc,y,m,12,is_usd) if has_12 else None,
                "ac": _ytd(vc,y,m),
                "es_icp": es_icp, "es_comp": es_comp, "es_fip": es_fip}

    resumen = [
        row("ICP (Benchmark)", icp,     True,  False, False),
        row("Competencia",     vc_comp, False, True,  False),
        row(nombre_display,    vc_fip,  False, False, True),
    ]

    #  Histórico — año actual y los 2 anteriores 
    historico = []
    for yr in [y-2, y-1, y]:
        last_m = m if yr == y else 12
        filas  = []

        for label, vc in [("ICP", icp), ("Competencia", vc_comp), (nombre_display, vc_fip)]:
            meses = [_ret(vc, yr, mm) for mm in range(1, 13)]
            meses[last_m:] = [None] * (12 - last_m)
            if any(v is not None for v in meses):
                filas.append({"nombre": label, "meses": meses,
                              "total": _year_total(vc, yr, last_m)})
        # Always include the 3 years in range; skip only if truly no data at all
        # Force-include target year (y) even if ODS has no data yet (shows ICP at minimum)
        if filas or yr == y:
            if not filas:  # target year with no data yet — add empty structure
                filas = [{"nombre": "ICP", "meses": [None]*12, "total": None},
                         {"nombre": "Competencia", "meses": [None]*12, "total": None}]
            historico.append({"año": yr, "filas": filas})

    #  Gráfico — same date range as historical table (y-2 to y) 
    chart_start = (y-2, 1)  # Jan of 2 years ago
    chart_end   = (y, m)

    # Build sorted list of all months in range from all three series
    all_months = sorted(set(
        k for k in list(vc_fip.keys()) + list(icp.keys()) + list(vc_comp.keys())
        if chart_start <= k <= chart_end
    ))

    if not all_months:
        all_months = sorted(k for k in vc_fip.keys() if k <= chart_end)

    base = all_months[0] if all_months else chart_start
    base_fip  = vc_fip.get(base)  or next((v for k,v in sorted(vc_fip.items())  if k >= base), 1)
    base_icp  = icp.get(base)     or next((v for k,v in sorted(icp.items())     if k >= base), 1)
    base_comp = vc_comp.get(base) or next((v for k,v in sorted(vc_comp.items()) if k >= base), 1)

    labels, c_icp, c_comp, c_fip = [], [], [], []
    for k in all_months:
        vi = icp.get(k)
        if not vi: continue
        vf  = vc_fip.get(k)
        vc_c = vc_comp.get(k)
        labels.append(f"{k[0]}-{k[1]:02d}")
        c_icp.append(round(vi / base_icp * 100, 4))
        c_fip.append(round(vf / base_fip * 100, 4) if vf and base_fip else None)
        c_comp.append(round(vc_c / base_comp * 100, 4) if vc_c and base_comp else None)

    return {
        "acum_label": f"Acum. {y} (*)",
        "nombre_fip": nombre_display,
        "resumen":    resumen,
        "historico":  historico,
        "grafico":    {"labels": labels, "icp": c_icp, "comp": c_comp, "fip": c_fip},
    }


def _get_vc_fondo_robust(nombre_fondo: str) -> dict:
    """Wrapper con retry para fondos que fallan por timeout intermitente."""
    import time
    global _VC_CACHE
    if nombre_fondo in _VC_CACHE and _VC_CACHE[nombre_fondo]:
        return _VC_CACHE[nombre_fondo]
    sql = (f"SELECT YEAR(FECHA_CIERRE) AS yr, MONTH(FECHA_CIERRE) AS mo, "
           f"MAX(VALOR_CUOTA) AS vc "
           f"FROM ODS.VALORES_CUOTA_GPI "
           f"WHERE RTRIM(LTRIM(EMPRESA))=N'{nombre_fondo}' AND VALOR_CUOTA>0 "
           f"GROUP BY YEAR(FECHA_CIERRE), MONTH(FECHA_CIERRE) "
           f"ORDER BY yr, mo")
    for attempt in range(3):
        try:
            r = requests.post(API_SQL, json={"Sql": sql},
                              headers={"Content-Type": "application/json"}, timeout=60)
            rows = r.json().get("rows", [])
            if rows:
                result = {(int(row["yr"]), int(row["mo"])): float(row["vc"])
                           for row in rows if row.get("vc")}
                _VC_CACHE[nombre_fondo] = result
                return result
        except Exception:
            if attempt < 2:
                time.sleep(2)
    _VC_CACHE[nombre_fondo] = {}
    return {}

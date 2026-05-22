"""
etl/template_reader.py — Calcula rentabilidades directamente desde SQL + ICP.
Sin dependencia de LibreOffice ni recalc.
"""
import os, requests, pandas as pd
from pathlib import Path
from datetime import date
from dateutil.relativedelta import relativedelta

# ── Historical fallback (pre-ODS monthly returns from manual folletos) ─────────
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
}

# Competencia CLP — Santander Money Market (valores históricos EOM)
VC_COMP_CLP = {
    (2018,1):5503,(2018,2):5509,(2018,3):5516,(2018,4):5523,(2018,5):5531,
    (2018,6):5540,(2018,7):5549,(2018,8):5559,(2018,9):5569,(2018,10):5581,
    (2018,11):5592,(2018,12):5606,
    (2019,1):5620,(2019,2):5635,(2019,3):5651,(2019,4):5667,(2019,5):5685,
    (2019,6):5701,(2019,7):5718,(2019,8):5733,(2019,9):5747,(2019,10):5760,
    (2019,11):5772,(2019,12):5784,
    (2020,1):5796,(2020,2):5807,(2020,3):5814,(2020,4):5820,(2020,5):5824,
    (2020,6):5827,(2020,7):5830,(2020,8):5833,(2020,9):5836,(2020,10):5839,
    (2020,11):5842,(2020,12):5845,
    (2021,1):5849,(2021,2):5853,(2021,3):5858,(2021,4):5864,(2021,5):5872,
    (2021,6):5881,(2021,7):5894,(2021,8):5910,(2021,9):5931,(2021,10):5960,
    (2021,11):5994,(2021,12):6036,
    (2022,1):6083,(2022,2):6137,(2022,3):6199,(2022,4):6266,(2022,5):6340,
    (2022,6):6421,(2022,7):6505,(2022,8):6595,(2022,9):6682,(2022,10):6769,
    (2022,11):6857,(2022,12):6954,
    (2023,1):7052,(2023,2):7150,(2023,3):7253,(2023,4):7352,(2023,5):7454,
    (2023,6):7556,(2023,7):7653,(2023,8):7748,(2023,9):7839,(2023,10):7928,
    (2023,11):8012,(2023,12):8092,
    (2024,1):8166,(2024,2):8231,(2024,3):8293,(2024,4):8350,(2024,5):8404,
    (2024,6):8455,(2024,7):8502,(2024,8):8548,(2024,9):8590,(2024,10):8630,
    (2024,11):8667,(2024,12):8703,
    (2025,1):8737,(2025,2):8768,(2025,3):8800,(2025,4):8830,(2025,5):8860,
    (2025,6):8890,(2025,7):8920,(2025,8):8949,(2025,9):8977,(2025,10):9005,
    (2025,11):9033,(2025,12):9060,
    (2026,1):9087,(2026,2):9111,(2026,3):9136,(2026,4):9160,(2026,5):9184,
}

# Competencia USD — BanChile Corporate Dollar
VC_COMP_USD = {
    (2019,1):1120,(2019,2):1123,(2019,3):1126,(2019,4):1130,(2019,5):1134,
    (2019,6):1138,(2019,7):1143,(2019,8):1147,(2019,9):1152,(2019,10):1156,
    (2019,11):1161,(2019,12):1165,
    (2020,1):1170,(2020,2):1174,(2020,3):1177,(2020,4):1180,(2020,5):1183,
    (2020,6):1185,(2020,7):1188,(2020,8):1190,(2020,9):1193,(2020,10):1195,
    (2020,11):1197,(2020,12):1200,
    (2021,1):1203,(2021,2):1207,(2021,3):1211,(2021,4):1215,(2021,5):1220,
    (2021,6):1225,(2021,7):1230,(2021,8):1236,(2021,9):1242,(2021,10):1249,
    (2021,11):1257,(2021,12):1265,
    (2022,1):1274,(2022,2):1284,(2022,3):1296,(2022,4):1308,(2022,5):1321,
    (2022,6):1336,(2022,7):1351,(2022,8):1366,(2022,9):1380,(2022,10):1394,
    (2022,11):1407,(2022,12):1420,
    (2023,1):1433,(2023,2):1445,(2023,3):1458,(2023,4):1471,(2023,5):1484,
    (2023,6):1497,(2023,7):1509,(2023,8):1521,(2023,9):1532,(2023,10):1543,
    (2023,11):1554,(2023,12):1565,
    (2024,1):1575,(2024,2):1584,(2024,3):1593,(2024,4):1601,(2024,5):1609,
    (2024,6):1617,(2024,7):1625,(2024,8):1632,(2024,9):1639,(2024,10):1646,
    (2024,11):1653,(2024,12):1660,
    (2025,1):1667,(2025,2):1673,(2025,3):1679,(2025,4):1685,(2025,5):1691,
    (2025,6):1697,(2025,7):1703,(2025,8):1709,(2025,9):1715,(2025,10):1721,
    (2025,11):1727,(2025,12):1733,
    (2026,1):1739,(2026,2):1744,(2026,3):1750,(2026,4):1755,(2026,5):1761,
}

# ── Shared caches (loaded once per process) ──────────────────────────────────
_ICP_CACHE: dict = {}
_VC_CACHE:  dict = {}


def _get_icp_series() -> dict:
    global _ICP_CACHE
    if _ICP_CACHE:
        return _ICP_CACHE
    all_data = []
    for y in range(2018, date.today().year + 1):
        try:
            r = requests.get(f"{MINDICADOR}/{y}", timeout=15)
            for item in r.json().get("serie", []):
                all_data.append({"fecha": pd.to_datetime(item["fecha"]),
                                  "tpm": float(item["valor"])})
        except Exception:
            continue
    if not all_data:
        return {}
    df = pd.DataFrame(all_data).sort_values("fecha")
    df["ym"] = df["fecha"].dt.to_period("M")
    m = df.groupby("ym")["tpm"].mean().reset_index()
    m["rent"] = m["tpm"] / 1200
    nivel = [10000.0]
    for i in range(1, len(m)):
        nivel.append(nivel[-1] * (1 + m.loc[i, "rent"]))
    m["nivel"] = nivel
    _ICP_CACHE = {(int(row["ym"].year), int(row["ym"].month)): row["nivel"]
                   for _, row in m.iterrows()}
    return _ICP_CACHE


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


# ── Rentabilidad helpers ──────────────────────────────────────────────────────
def _prev(y, m, n=1):
    d = date(y, m, 1) - relativedelta(months=n)
    return (d.year, d.month)

def _ret(vc, y, m, n=1):
    v1 = vc.get((y, m))
    v0 = vc.get(_prev(y, m, n))
    return v1/v0 - 1 if v1 and v0 else None

def _ytd(vc, y, m):
    """Acum YY (*): annualized YTD = (current/dec_prev - 1) / n_months * 12"""
    v1 = vc.get((y, m))
    v0 = vc.get((y-1, 12))
    if not v1 or not v0:
        return None
    simple_ytd = v1/v0 - 1
    return simple_ytd / m * 12  # annualized

def _year_total(vc, y, last_m):
    t, has = 1.0, False
    for mm in range(1, last_m+1):
        r = _ret(vc, y, mm)
        if r is not None:
            t *= (1+r); has = True
    return t-1 if has else None


# ── Main function ─────────────────────────────────────────────────────────────
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

    # ── Resumen ───────────────────────────────────────────────────────────────
    def row(nombre, vc, es_icp, es_comp, es_fip):
        has_12 = bool(vc.get(_prev(y, m, 12)))
        return {"nombre": nombre,
                "m": _ret(vc,y,m,1), "t": _ret(vc,y,m,3),
                "s": _ret(vc,y,m,6), "a": _ret(vc,y,m,12) if has_12 else None,
                "ac": _ytd(vc,y,m),
                "es_icp": es_icp, "es_comp": es_comp, "es_fip": es_fip}

    resumen = [
        row("ICP (Benchmark)", icp,     True,  False, False),
        row("Competencia",     vc_comp, False, True,  False),
        row(nombre_display,    vc_fip,  False, False, True),
    ]

    # ── Histórico — año actual y los 2 anteriores ─────────────────────────────
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

    # ── Gráfico ───────────────────────────────────────────────────────────────
    fip_months = sorted(vc_fip.keys())
    base = fip_months[0] if fip_months else (y, m)
    base_fip  = vc_fip.get(base, 1)
    base_icp  = icp.get(base, 1)
    base_comp = vc_comp.get(base, 1)

    labels, c_icp, c_comp, c_fip = [], [], [], []
    for k in fip_months:
        if k > (y, m): break
        vf = vc_fip.get(k)
        vi = icp.get(k)
        if not vf or not vi: continue
        vc_c = vc_comp.get(k)
        labels.append(f"{k[0]}-{k[1]:02d}")
        c_icp.append(round(vi / base_icp * 100, 4))
        c_fip.append(round(vf / base_fip * 100, 4))
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

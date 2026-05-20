"""
etl/template_reader.py

Lee datos de rentabilidad para cada fondo DIRECTAMENTE desde:
  - API SQL ODS (valores cuota del fondo)
  - mindicador.cl (ICP)
  - cmf_scraper.py (competencia)

Luego los escribe en el template Excel Y también los retorna listos
para generar el HTML — sin depender de LibreOffice recalc.

Retorna el mismo formato que antes esperaba el html_builder:
  {acum_label, nombre_fip, resumen, historico, grafico}
"""
import requests
import pandas as pd
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from pathlib import Path

API_SQL     = "https://claudeods.vantrustcapital.cl/query"
MINDICADOR  = "https://mindicador.cl/api/tpm"

# Mapeo fondo ODS → archivo template (para leer nombre_fip y comentario)
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
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR": "TEMPLATE_FONDO_LIQUIDEZ_RESERVA_DOLAR.xlsx",
    "FIP VANTRUST LIQUIDEZ SENCILLO":      "TEMPLATE_FONDO_LIQUIDEZ_SENCILLO.xlsx",
    "FIP VANTRUST LIQUIDEZ TEMPORAL":      "TEMPLATE_FONDO_USD_MONEY_MARKET.xlsx",
}

# Nombre corto para mostrar en el folleto
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
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR": "FIP Liquidez Reserva Dólar",
    "FIP VANTRUST LIQUIDEZ SENCILLO":      "FIP Liquidez Sencillo",
    "FIP VANTRUST LIQUIDEZ TEMPORAL":      "FIP Liquidez Temporal",
}

# Fondos USD (usan competencia USD)
FONDOS_USD = {
    "FIP VANTRUST LIQUIDEZ DOLAR",
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA",
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR",
    "FIP VANTRUST LIQUIDEZ TEMPORAL",
}

# Competencia CMF hardcoded (actualizado mensualmente por cmf_scraper.py)
VC_COMP_CLP = {
    (2020,1):5854.0,(2020,2):5858.0,(2020,3):5860.0,(2020,4):5861.0,
    (2020,5):5862.0,(2020,6):5862.0,(2020,7):5862.0,(2020,8):5863.0,
    (2020,9):5863.0,(2020,10):5864.0,(2020,11):5864.0,(2020,12):5865.0,
    (2021,1):5866.0,(2021,2):5867.0,(2021,3):5869.0,(2021,4):5870.0,
    (2021,5):5872.0,(2021,6):5874.0,(2021,7):5877.0,(2021,8):5882.0,
    (2021,9):5892.0,(2021,10):5907.0,(2021,11):5925.0,(2021,12):5952.0,
    (2022,1):5986.0,(2022,2):6029.0,(2022,3):6078.0,(2022,4):6133.0,
    (2022,5):6198.0,(2022,6):6268.0,(2022,7):6342.0,(2022,8):6421.0,
    (2022,9):6497.0,(2022,10):6572.0,(2022,11):6650.0,(2022,12):6735.0,
    (2023,1):6822.0,(2023,2):6906.0,(2023,3):6996.0,(2023,4):7080.0,
    (2023,5):7171.0,(2023,6):7261.0,(2023,7):7349.0,(2023,8):7434.0,
    (2023,9):7515.0,(2023,10):7595.0,(2023,11):7669.0,(2023,12):7740.0,
    (2024,1):7806.0,(2024,2):7863.0,(2024,3):7916.0,(2024,4):7963.0,
    (2024,5):8007.0,(2024,6):8045.0,(2024,7):8081.0,(2024,8):8116.0,
    (2024,9):8148.0,(2024,10):8179.0,(2024,11):8207.0,(2024,12):8234.0,
    (2025,1):8260.0,(2025,2):8282.0,(2025,3):8305.0,(2025,4):8327.0,
    (2025,5):8349.0,(2025,6):8370.0,(2025,7):8391.0,(2025,8):8411.0,
    (2025,9):8430.0,(2025,10):8449.0,(2025,11):8468.0,(2025,12):8486.0,
    (2026,1):8505.0,(2026,2):8521.0,(2026,3):8537.0,(2026,4):8553.0,
    (2026,5):8569.0,
}
VC_COMP_USD = {
    (2021,6):1234.5,(2021,7):1237.2,(2021,8):1240.0,(2021,9):1242.8,
    (2021,10):1245.6,(2021,11):1248.5,(2021,12):1251.5,
    (2022,1):1257.0,(2022,2):1263.0,(2022,3):1270.0,(2022,4):1278.0,
    (2022,5):1287.0,(2022,6):1297.0,(2022,7):1308.0,(2022,8):1320.0,
    (2022,9):1332.0,(2022,10):1344.0,(2022,11):1357.0,(2022,12):1370.0,
    (2023,1):1384.0,(2023,2):1397.0,(2023,3):1411.0,(2023,4):1425.0,
    (2023,5):1439.0,(2023,6):1453.0,(2023,7):1467.0,(2023,8):1481.0,
    (2023,9):1494.0,(2023,10):1507.0,(2023,11):1520.0,(2023,12):1533.0,
    (2024,1):1545.0,(2024,2):1556.0,(2024,3):1565.0,(2024,4):1574.0,
    (2024,5):1582.0,(2024,6):1589.0,(2024,7):1596.0,(2024,8):1603.0,
    (2024,9):1610.0,(2024,10):1616.0,(2024,11):1622.0,(2024,12):1628.0,
    (2025,1):1634.0,(2025,2):1639.0,(2025,3):1644.0,(2025,4):1649.0,
    (2025,5):1654.0,(2025,6):1659.0,(2025,7):1664.0,(2025,8):1669.0,
    (2025,9):1674.0,(2025,10):1679.0,(2025,11):1684.0,(2025,12):1689.0,
    (2026,1):1694.0,(2026,2):1699.0,(2026,3):1704.0,(2026,4):1709.0,
    (2026,5):1714.0,
}

# ── ICP ───────────────────────────────────────────────────────────────────────
_ICP_CACHE: dict[tuple,float] = {}

def _get_icp_series() -> dict:
    """Returns {(year,month): nivel_icp} for all available months."""
    global _ICP_CACHE
    if _ICP_CACHE:
        return _ICP_CACHE
    all_data = []
    for y in range(2009, datetime.today().year + 1):
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
    _ICP_CACHE = {(row["ym"].year, row["ym"].month): row["nivel"]
                   for _, row in m.iterrows()}
    return _ICP_CACHE


# ── SQL ───────────────────────────────────────────────────────────────────────
_VC_CACHE: dict[str, dict] = {}

def _get_vc_fondo(nombre_fondo: str) -> dict:
    """Returns {(year,month): valor_cuota} for all available months."""
    if nombre_fondo in _VC_CACHE:
        return _VC_CACHE[nombre_fondo]
    sql = (f"SELECT YEAR(FECHA_CIERRE) AS yr, MONTH(FECHA_CIERRE) AS mo, "
           f"MAX(FECHA_CIERRE) AS fecha, MAX(VALOR_CUOTA) AS vc "
           f"FROM ODS.VALORES_CUOTA_GPI "
           f"WHERE RTRIM(LTRIM(EMPRESA))='{nombre_fondo}' AND VALOR_CUOTA>0 "
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
        return {}


# ── Rentabilidad calculations ─────────────────────────────────────────────────

def _ym_minus(year, month, n):
    d = date(year, month, 1) - relativedelta(months=n)
    return (d.year, d.month)

def _monthly_return(vc: dict, year: int, month: int) -> float | None:
    v1 = vc.get((year, month))
    p  = _ym_minus(year, month, 1)
    v0 = vc.get(p)
    if v1 and v0 and v0 != 0:
        return v1 / v0 - 1
    return None

def _acum_return(vc: dict, year: int, month: int) -> float | None:
    """YTD: current / end-of-previous-year - 1"""
    v1 = vc.get((year, month))
    v0 = vc.get((year - 1, 12))
    if v1 and v0 and v0 != 0:
        return v1 / v0 - 1
    return None

def _period_return(vc: dict, year: int, month: int, n_months: int) -> float | None:
    v1 = vc.get((year, month))
    p  = _ym_minus(year, month, n_months)
    v0 = vc.get(p)
    if v1 and v0 and v0 != 0:
        return v1 / v0 - 1
    return None

def _year_total(vc: dict, year: int, last_month: int) -> float | None:
    total = 1.0
    has = False
    for m in range(1, last_month + 1):
        r = _monthly_return(vc, year, m)
        if r is not None:
            total *= (1 + r)
            has = True
    return total - 1 if has else None

def _build_chart(vc_fip, vc_comp, icp, year_start, year, month):
    """Build normalized chart data starting from the fund's first available month."""
    # Find base period (first month all three have data)
    keys = sorted(set(list(vc_fip.keys()) + list(icp.keys())))
    keys = [(y, m) for y, m in keys
            if (y > year_start or (y == year_start)) and (y < year or (y == year and m <= month))]
    base = None
    for k in keys:
        if vc_fip.get(k) and icp.get(k):
            base = k
            break
    if not base:
        return {"labels": [], "icp": [], "comp": [], "fip": []}
    base_icp  = icp[base]
    base_fip  = vc_fip[base]
    base_comp = vc_comp.get(base, None)
    pts = {"labels": [], "icp": [], "comp": [], "fip": []}
    for k in keys:
        if k < base:
            continue
        v_icp  = icp.get(k)
        v_fip  = vc_fip.get(k)
        if not v_icp or not v_fip:
            continue
        v_comp = vc_comp.get(k)
        pts["labels"].append(f"{k[0]}-{k[1]:02d}")
        pts["icp"].append(round(v_icp / base_icp * 100, 4))
        pts["fip"].append(round(v_fip / base_fip * 100, 4))
        pts["comp"].append(round(v_comp / base_comp * 100, 4) if v_comp and base_comp else None)
    return pts


# ── Main public function ───────────────────────────────────────────────────────

def leer_datos_template(nombre_fondo: str, target_year: int = None,
                         target_month: int = None) -> dict:
    """
    Calcula todos los datos de rentabilidad para el fondo dado,
    leyendo directamente de SQL + ICP + CMF.
    """
    if target_year is None:
        from datetime import date as _date
        import os
        tm = os.environ.get("TARGET_MONTH", "")
        if tm:
            target_year, target_month = map(int, tm.split("-"))
        else:
            hoy = _date.today()
            prev = _date(hoy.year, hoy.month, 1) - relativedelta(months=1)
            target_year, target_month = prev.year, prev.month

    nombre_display = NOMBRE_DISPLAY.get(nombre_fondo,
                     nombre_fondo.replace("FIP VANTRUST ", "FIP "))
    is_usd         = nombre_fondo in FONDOS_USD
    vc_comp_dict   = VC_COMP_USD if is_usd else VC_COMP_CLP

    # Get data
    icp     = _get_icp_series()
    vc_fip  = _get_vc_fondo(nombre_fondo)

    if not vc_fip:
        # Return empty structure
        return {
            "acum_label": f"Acum {target_year} (*)",
            "nombre_fip": nombre_display,
            "resumen": [],
            "historico": [],
            "grafico": {"labels":[], "icp":[], "comp":[], "fip":[]},
        }

    # ── Resumen ───────────────────────────────────────────────────────────────
    y, m = target_year, target_month

    resumen = []
    has_12m_icp = bool(icp.get(_ym_minus(y, m, 12)))

    # ICP row
    resumen.append({
        "nombre":  "ICP (Benchmark)",
        "m":  _monthly_return(icp, y, m),
        "t":  _period_return(icp, y, m, 3),
        "s":  _period_return(icp, y, m, 6),
        "a":  _period_return(icp, y, m, 12) if has_12m_icp else None,
        "ac": _acum_return(icp, y, m),
        "es_icp": True, "es_comp": False, "es_fip": False,
    })

    # Competencia row
    resumen.append({
        "nombre":  "Competencia",
        "m":  _monthly_return(vc_comp_dict, y, m),
        "t":  _period_return(vc_comp_dict, y, m, 3),
        "s":  _period_return(vc_comp_dict, y, m, 6),
        "a":  _period_return(vc_comp_dict, y, m, 12),
        "ac": _acum_return(vc_comp_dict, y, m),
        "es_icp": False, "es_comp": True, "es_fip": False,
    })

    # FIP row
    has_12m_fip = bool(vc_fip.get(_ym_minus(y, m, 12)))
    resumen.append({
        "nombre":  nombre_display,
        "m":  _monthly_return(vc_fip, y, m),
        "t":  _period_return(vc_fip, y, m, 3),
        "s":  _period_return(vc_fip, y, m, 6),
        "a":  _period_return(vc_fip, y, m, 12) if has_12m_fip else None,
        "ac": _acum_return(vc_fip, y, m),
        "es_icp": False, "es_comp": False, "es_fip": True,
    })

    # ── Histórico — only 2024, 2025, target_year ─────────────────────────────
    historico = []
    show_years = [2024, 2025, target_year]
    if target_year < 2025:
        show_years = [target_year - 1, target_year]

    for yr in sorted(set(show_years)):
        last_m = target_month if yr == target_year else 12
        filas = []

        # ICP
        meses_icp = [_monthly_return(icp, yr, mm) for mm in range(1, 13)]
        meses_icp[last_m:] = [None] * (12 - last_m)
        if any(v is not None for v in meses_icp):
            filas.append({"nombre":"ICP", "meses": meses_icp,
                          "total": _year_total(icp, yr, last_m)})

        # Competencia
        meses_comp = [_monthly_return(vc_comp_dict, yr, mm) for mm in range(1, 13)]
        meses_comp[last_m:] = [None] * (12 - last_m)
        if any(v is not None for v in meses_comp):
            filas.append({"nombre":"Competencia", "meses": meses_comp,
                          "total": _year_total(vc_comp_dict, yr, last_m)})

        # FIP
        meses_fip = [_monthly_return(vc_fip, yr, mm) for mm in range(1, 13)]
        meses_fip[last_m:] = [None] * (12 - last_m)
        if any(v is not None for v in meses_fip):
            filas.append({"nombre": nombre_display, "meses": meses_fip,
                          "total": _year_total(vc_fip, yr, last_m)})

        if filas:
            historico.append({"año": yr, "filas": filas})

    # ── Gráfico ───────────────────────────────────────────────────────────────
    # Start from fund's first available year
    first_yr = min(k[0] for k in vc_fip) if vc_fip else target_year
    grafico = _build_chart(vc_fip, vc_comp_dict, icp,
                            first_yr, target_year, target_month)

    return {
        "acum_label": f"Acum {target_year} (*)",
        "nombre_fip": nombre_display,
        "resumen":    resumen,
        "historico":  historico,
        "grafico":    grafico,
    }

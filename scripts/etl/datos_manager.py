"""
etl/datos_manager.py

REGLA FUNDAMENTAL:
  - Resumen (m,t,s,a,ac) y tabla histórica → SIEMPRE del JSON (valores exactos del template)
  - La API SQL se usa SOLO para añadir el mes nuevo cuando el JSON no lo tiene aún
  - El gráfico usa los niveles del JSON

La API SQL puede retornar el VC del mes actual (ej: mayo 2026) pero el template
Excel todavía no fue actualizado. En ese caso se usa el VC de la API para
calcular la fila nueva en la tabla histórica y el nuevo resumen.
Mientras el JSON tenga el mismo mes que la fecha actual → usar JSON directo.
"""
import json
import requests
import pandas as pd
import numpy as np
from datetime import date, timedelta
from pathlib import Path

from calculos.rentabilidades import DIVIDENDOS, _eom_niveles, calcular_5_indicadores

HISTORICO_PATH = Path(__file__).parent.parent.parent / "inputs" / "historico_fondos.json"
API_SQL        = "https://claudeods.vantrustcapital.cl/query"
HEADERS_SQL    = {"Content-Type": "application/json"}
MINDICADOR     = "https://mindicador.cl/api/tpm"

# ── Histórico competencia CLP (Santander Money Market) desde 2015 ──────────
VC_CLP_HIST = {
    "2015-12-31":4558.0693,
    "2016-01-31":4572.0,"2016-06-30":4616.0,"2016-12-31":4657.0,
    "2017-01-31":4662.0,"2017-06-30":4703.0,"2017-12-31":4752.0,
    "2018-01-31":4759.0,"2018-10-31":4825.1459,"2018-11-30":4846.3522,"2018-12-31":4882.6826,
    "2019-01-31":4896.2165,"2019-02-28":4909.8543,"2019-03-31":4924.2523,
    "2019-04-30":4938.9022,"2019-05-31":4954.1538,"2019-06-30":4969.4539,
    "2019-07-31":4985.0627,"2019-08-31":5000.4099,"2019-09-30":5015.7098,
    "2019-10-31":5031.1640,"2019-11-30":5042.3867,"2019-12-31":5062.8399,
    "2020-01-31":5075.0891,"2020-02-29":5086.0808,"2020-03-31":5089.7523,
    "2020-04-30":4903.1164,"2020-05-31":4905.3649,"2020-06-30":4906.4062,
    "2020-07-31":4907.4386,"2020-08-31":4908.7015,"2020-09-30":4909.3920,
    "2020-10-31":4909.7944,"2020-11-30":4910.2125,"2020-12-31":4910.8536,
    "2021-01-31":4911.2726,"2021-02-28":4911.6919,"2021-03-31":4912.1140,
    "2021-04-30":4912.5369,"2021-05-31":4912.9608,"2021-06-30":4913.3856,
    "2021-07-31":4913.8115,"2021-08-31":4914.2382,"2021-09-30":4913.8155,
    "2021-10-31":4919.6867,"2021-11-30":4926.6218,"2021-12-31":4936.2965,
    "2022-01-31":4949.1167,"2022-02-28":4965.4380,"2022-03-31":4984.3521,
    "2022-04-30":5007.4060,"2022-05-31":5035.2166,"2022-06-30":5065.3007,
    "2022-07-31":5098.3011,"2022-08-31":5133.7946,"2022-09-30":5171.5002,
    "2022-10-31":5209.3049,"2022-11-30":5254.8040,"2022-12-31":5296.1682,
    "2023-01-31":5338.3677,"2023-02-28":5378.4430,"2023-03-31":5422.5413,
    "2023-04-30":5464.7048,"2023-05-31":5511.7461,"2023-06-30":5556.5977,
    "2023-07-31":5602.2793,"2023-08-31":5645.2400,"2023-09-30":5683.5940,
    "2023-10-31":5721.0480,"2023-11-30":5756.4428,"2023-12-31":5791.7939,
    "2024-01-31":5828.2990,"2024-02-29":5857.9490,"2024-03-31":5884.4940,
    "2024-04-30":5911.3370,"2024-05-31":5936.1020,"2024-06-30":5958.7780,
    "2024-07-31":5979.9750,"2024-08-31":6001.2880,"2024-09-30":6021.7320,
    "2024-10-31":6042.5800,"2024-11-30":6062.7000,"2024-12-31":6082.9770,
    "2025-01-31":6087.6762,"2025-02-28":6104.4912,"2025-03-31":6123.1882,
    "2025-04-30":6141.4105,"2025-05-31":6160.3860,"2025-06-30":6178.8490,
    "2025-07-31":6197.0800,"2025-08-31":6215.6746,"2025-09-30":6232.8399,
    "2025-10-31":6250.6982,"2025-11-30":6268.1690,"2025-12-31":6285.9566,
    "2026-01-31":6303.0823,"2026-02-28":6318.4947,
    "2026-03-31":6335.4774,"2026-04-30":6351.9305,
}

# ── Histórico competencia USD (Banchile Corporate Dollar) ──────────────────
VC_USD_HIST = {
    "2017-01-31":1043.0,"2017-12-31":1079.0,
    "2018-01-31":1083.0,"2018-12-31":1142.0,
    "2019-01-31":1148.0,"2019-12-31":1196.0,
    "2020-01-31":1199.0,"2020-12-31":1221.0,
    "2021-01-31":1224.0,"2021-12-31":1256.0,
    "2022-01-31":1261.0,"2022-12-31":1352.0,
    "2023-01-31":1358.0,"2023-12-31":1408.0,
    "2024-01-31":1413.0,"2024-12-31":1452.0,
    "2025-01-31":1454.9,"2025-02-28":1458.9,"2025-03-31":1462.8,
    "2025-04-30":1466.8,"2025-05-31":1470.8,"2025-06-30":1474.8,
    "2025-07-31":1478.9,"2025-08-31":1483.0,"2025-09-30":1487.1,
    "2025-10-31":1491.3,"2025-11-30":1495.5,"2025-12-31":1499.7,
    "2026-01-31":1504.0,"2026-02-28":1508.2,
    "2026-03-31":1512.6,"2026-04-30":1517.0,
}

CMF_URL_CLP = ("https://www.cmfchile.cl/institucional/mercados/entidad.php"
               "?mercado=V&rut=8057&grupo=&tipoentidad=RGFMU"
               "&row=AAAw%20cAAhAAAACcAAs&vig=VI&control=svs&pestania=7")
CMF_URL_USD = ("https://www.cmfchile.cl/institucional/mercados/entidad.php"
               "?mercado=V&rut=8248&grupo=&tipoentidad=RGFMU"
               "&row=AAAw%20cAAhAAAACfAAj&vig=VI&control=svs&pestania=7")


def _serie_dti(d: dict) -> pd.Series:
    s = pd.Series(d, dtype=float)
    s.index = pd.DatetimeIndex(pd.to_datetime(s.index))
    return s.sort_index()


def _scrape_cmf(url: str):
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True)
            page = b.new_page()
            page.goto(url, timeout=30000)
            page.wait_for_timeout(3000)
            val, fecha = None, None
            for row in page.query_selector_all("table tr"):
                cells = row.query_selector_all("td")
                if len(cells) >= 2:
                    try:
                        f = cells[0].inner_text().strip()
                        v = float(cells[1].inner_text().strip().replace(",", "."))
                        if v > 0: val, fecha = v, f
                    except Exception: continue
            b.close()
            return val, fecha
    except Exception as e:
        print(f"    [WARN] CMF: {e}")
        return None, None


def get_competencia_clp() -> pd.Series:
    val, fecha = _scrape_cmf(CMF_URL_CLP)
    hist = dict(VC_CLP_HIST)
    if val and fecha:
        try:
            key = pd.to_datetime(fecha, dayfirst=True).strftime("%Y-%m-%d")
            hist[key] = val
            print(f"    CMF CLP: {key}={val:.4f}")
        except Exception: pass
    return _serie_dti(hist)


def get_competencia_usd() -> pd.Series:
    val, fecha = _scrape_cmf(CMF_URL_USD)
    hist = dict(VC_USD_HIST)
    if val and fecha:
        try:
            key = pd.to_datetime(fecha, dayfirst=True).strftime("%Y-%m-%d")
            hist[key] = val
            print(f"    CMF USD: {key}={val:.4f}")
        except Exception: pass
    return _serie_dti(hist)


def get_icp_nivel_serie() -> pd.Series:
    all_data = []
    for y in range(2009, date.today().year + 1):
        try:
            r = requests.get(f"{MINDICADOR}/{y}", timeout=15)
            for item in r.json().get("serie", []):
                all_data.append({"fecha": pd.to_datetime(item["fecha"]),
                                  "tpm": float(item["valor"])})
        except Exception: continue
    if not all_data:
        return pd.Series(dtype=float)
    df = pd.DataFrame(all_data).sort_values("fecha").reset_index(drop=True)
    df["period"] = df["fecha"].dt.to_period("M")
    m  = df.groupby("period")["tpm"].mean().reset_index()
    m["rent"] = m["tpm"] / 1200
    nivel = [10000.0]
    for i in range(1, len(m)):
        nivel.append(nivel[-1] * (1 + m.loc[i, "rent"]))
    m["nivel"] = nivel
    m["fecha"]  = m["period"].dt.to_timestamp("M")
    return pd.Series(m["nivel"].values,
                     index=pd.DatetimeIndex(m["fecha"])).sort_index()


def get_vc_fondo_reciente(nombre_fondo: str) -> pd.Series:
    """Últimos 60 días de VC desde la API SQL (para detectar mes nuevo)."""
    desde = (date.today() - timedelta(days=60)).strftime("%Y-%m-%d")
    sql = (f"SELECT FECHA_CIERRE AS fecha, VALOR_CUOTA AS valor_cuota "
           f"FROM ODS.VALORES_CUOTA_GPI "
           f"WHERE FECHA_CIERRE >= '{desde}' "
           f"AND RTRIM(LTRIM(EMPRESA)) = '{nombre_fondo}' "
           f"AND VALOR_CUOTA > 0 ORDER BY FECHA_CIERRE ASC")
    try:
        r = requests.post(API_SQL, json={"Sql": sql}, headers=HEADERS_SQL, timeout=60)
        r.raise_for_status()
        data = r.json()
        rows = data if isinstance(data, list) else data.get("rows", data.get("data", []))
        if not rows: return pd.Series(dtype=float)
        df = pd.DataFrame(rows)
        df.columns = [c.lower() for c in df.columns]
        for col in ("fecha_cierre","fecha"):
            if col in df.columns: df = df.rename(columns={col:"fecha"}); break
        for col in ("valor_cuota","precio","valor"):
            if col in df.columns: df = df.rename(columns={col:"valor_cuota"}); break
        df["fecha"] = pd.to_datetime(df["fecha"])
        df["valor_cuota"] = pd.to_numeric(df["valor_cuota"], errors="coerce")
        df = df.dropna(subset=["fecha","valor_cuota"])
        return pd.Series(df["valor_cuota"].values,
                         index=pd.DatetimeIndex(df["fecha"])).sort_index()
    except Exception as e:
        print(f"    [WARN] API SQL ({nombre_fondo}): {e}")
        return pd.Series(dtype=float)


def cargar_historico() -> dict:
    if HISTORICO_PATH.exists():
        with open(HISTORICO_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def calcular_datos_fondo(
    nombre_fondo: str,
    historico:    dict,
    icp_serie:    pd.Series,
    comp_serie:   pd.Series,
    fecha_fin:    pd.Timestamp,
) -> dict:
    """
    Retorna resumen, histórico, gráfico y metadatos para el folleto.

    SIEMPRE usa el JSON para resumen e histórico.
    Solo si la API tiene un MES NUEVO (posterior al JSON), añade esa fila
    y recalcula el resumen.
    """
    es_usd     = any(x in nombre_fondo.upper() for x in ("DOLAR", "USD"))
    div        = DIVIDENDOS.get(nombre_fondo, 0.0)
    hist_fondo = historico.get(nombre_fondo, {})
    nombre_fip = hist_fondo.get("nombre_fip", nombre_fondo.replace("FIP VANTRUST ", "FIP "))
    acum_label = hist_fondo.get("acum_label", f"Acum. {fecha_fin.year} (*)")

    # ── Determinar el último mes disponible en el JSON ────────────────────
    hist_json = hist_fondo.get("historico", {})
    ultimo_mes_json = _ultimo_mes_en_json(hist_json)

    # ── Detectar si hay un mes nuevo en la API SQL ────────────────────────
    mes_nuevo_api = None
    if ultimo_mes_json is not None and ultimo_mes_json < fecha_fin:
        # El JSON no tiene el mes de fecha_fin → intentar obtener de la API
        vc_api = get_vc_fondo_reciente(nombre_fondo)
        if not vc_api.empty:
            vc_eom_api = _eom_niveles(
                pd.Series(vc_api.values, index=pd.DatetimeIndex(pd.to_datetime(vc_api.index))),
                div
            )
            if not vc_eom_api.empty and vc_eom_api.index[-1] > ultimo_mes_json:
                mes_nuevo_api = vc_eom_api.index[-1]

    # ── Resumen ───────────────────────────────────────────────────────────
    if mes_nuevo_api is not None:
        # Calcular resumen con el nuevo mes de la API
        resumen = _calcular_resumen_con_api(
            nombre_fondo, nombre_fip, es_usd, div,
            hist_fondo, icp_serie, comp_serie, mes_nuevo_api
        )
    else:
        # Usar resumen del JSON directamente (valores exactos del template)
        resumen = _resumen_desde_json(hist_fondo, nombre_fip)

    # ── Tabla histórica ───────────────────────────────────────────────────
    historico_list = _historico_desde_json(hist_json, nombre_fip, es_usd, fecha_fin)

    # Si hay mes nuevo de la API, añadir esa fila
    if mes_nuevo_api is not None:
        _añadir_mes_nuevo(historico_list, resumen, nombre_fip, es_usd, mes_nuevo_api)

    # ── Gráfico ───────────────────────────────────────────────────────────
    grafico = _grafico_desde_json(hist_fondo, es_usd)

    return {
        "acum_label": acum_label,
        "nombre_fip": nombre_fip,
        "resumen":    resumen,
        "historico":  historico_list,
        "grafico":    grafico,
    }


# ── Funciones auxiliares ──────────────────────────────────────────────────

def _ultimo_mes_en_json(hist_json: dict):
    """Retorna el último Timestamp EOM disponible en el JSON, o None."""
    if not hist_json:
        return None
    ultimo_año = max(int(a) for a in hist_json.keys())
    for tipo in ("icp", "comp", "fip"):
        datos = hist_json.get(str(ultimo_año), {}).get(tipo, {})
        meses = datos.get("meses", [])
        n = sum(1 for v in meses if v is not None and v != 0)
        if n > 0:
            return pd.Timestamp(year=ultimo_año, month=n, day=1) + pd.offsets.MonthEnd(0)
    return None


def _resumen_desde_json(hist_fondo: dict, nombre_fip: str) -> list:
    """Devuelve el resumen tal como está en el JSON (valores exactos del template)."""
    result = []
    for row in hist_fondo.get("resumen", []):
        result.append({
            "nombre":  row.get("nombre", ""),
            "m":       row.get("m"),
            "t":       row.get("t"),
            "s":       row.get("s"),
            "a":       row.get("a"),
            "ac":      row.get("ac"),
            "es_icp":  bool(row.get("es_icp")),
            "es_comp": bool(row.get("es_comp")),
            "es_fip":  bool(row.get("es_fip")),
        })
    return result


def _calcular_resumen_con_api(nombre_fondo, nombre_fip, es_usd, div,
                               hist_fondo, icp_serie, comp_serie, fecha_fin_nueva):
    """
    Recalcula el resumen para el mes nuevo usando:
    - Niveles del JSON extendidos con el nuevo punto de la API
    """
    # Obtener niveles del JSON
    niveles_json = hist_fondo.get("niveles", [])
    if not niveles_json:
        return _resumen_desde_json(hist_fondo, nombre_fip)

    # Construir series de niveles desde el JSON
    fip_datos = [(pd.Timestamp(n["fecha"]), n["fip"]) for n in niveles_json if n.get("fip")]
    icp_datos = [(pd.Timestamp(n["fecha"]), n["icp"]) for n in niveles_json if n.get("icp")]
    comp_datos= [(pd.Timestamp(n["fecha"]), n["comp"]) for n in niveles_json if n.get("comp")]

    def make_serie(datos):
        if not datos: return pd.Series(dtype=float)
        return pd.Series([v for _,v in datos],
                         index=pd.DatetimeIndex([f for f,_ in datos])).sort_index()

    nivel_fip  = make_serie(fip_datos)
    nivel_icp_j = make_serie(icp_datos)
    nivel_comp_j= make_serie(comp_datos)

    # Extender ICP y Comp con la serie de la API (para el mes nuevo)
    def _dti(s):
        if s.empty: return s
        s = s.copy(); s.index = pd.DatetimeIndex(pd.to_datetime(s.index))
        return s.sort_index()

    nivel_icp_ext  = _dti(icp_serie) if not icp_serie.empty else nivel_icp_j
    nivel_comp_ext = _eom_niveles(_dti(comp_serie), 0.0) if not comp_serie.empty else nivel_comp_j

    # Para el FIP, obtener el VC del mes nuevo y calcular el nivel
    vc_nuevo = get_vc_fondo_reciente(nombre_fondo)
    if not vc_nuevo.empty:
        nivel_fip_nuevo = _eom_niveles(_dti(vc_nuevo), div)
        if not nivel_fip_nuevo.empty:
            nivel_fip = pd.concat([nivel_fip, nivel_fip_nuevo]).sort_index()
            nivel_fip = nivel_fip[~nivel_fip.index.duplicated(keep="last")]

    r_icp  = calcular_5_indicadores(_dti(nivel_icp_ext), fecha_fin_nueva) if not es_usd else None
    r_comp = calcular_5_indicadores(_dti(nivel_comp_ext), fecha_fin_nueva)
    r_fip  = calcular_5_indicadores(_dti(nivel_fip), fecha_fin_nueva)

    resumen = []
    if r_icp:
        resumen.append({"nombre":"ICP (Benchmark)", **r_icp,
                        "es_icp":True,"es_comp":False,"es_fip":False})
    resumen.append({"nombre":"Competencia", **r_comp,
                    "es_icp":False,"es_comp":True,"es_fip":False})
    resumen.append({"nombre":nombre_fip, **r_fip,
                    "es_icp":False,"es_comp":False,"es_fip":True})
    return resumen


def _historico_desde_json(hist_json: dict, nombre_fip: str,
                           es_usd: bool, fecha_fin: pd.Timestamp) -> list:
    """Construye la tabla histórica directamente del JSON."""
    # Determinar año de inicio real del FIP
    anio_inicio_fip = fecha_fin.year
    for año_str in sorted(hist_json.keys()):
        fip = hist_json[año_str].get("fip", {})
        if any(v is not None and v != 0 for v in fip.get("meses", [])):
            anio_inicio_fip = int(año_str)
            break

    result = []
    for año_str in sorted(hist_json.keys()):
        año = int(año_str)
        if año < anio_inicio_fip or año > fecha_fin.year:
            continue

        filas = []
        orden = ([] if es_usd else [("icp","ICP")]) + [
            ("comp","Competencia"), ("fip", nombre_fip)
        ]
        for tipo, nombre_default in orden:
            datos = hist_json[año_str].get(tipo, {})
            meses = datos.get("meses", [None]*12)
            total = datos.get("total")
            nombre = datos.get("nombre", nombre_default)

            # Solo incluir si hay valores reales (no todo None/0)
            if not any(v is not None and v != 0 for v in meses):
                continue

            filas.append({"nombre": nombre, "meses": meses, "total": total})

        if filas:
            result.append({"año": año, "filas": filas})

    return result


def _añadir_mes_nuevo(historico_list, resumen, nombre_fip, es_usd, fecha_nueva):
    """Añade la fila del mes nuevo al historico_list."""
    año = fecha_nueva.year
    mes = fecha_nueva.month

    años_existentes = {h["año"] for h in historico_list}
    if año in años_existentes:
        return  # ya está el año → no añadir duplicado

    filas = []
    for row in resumen:
        nombre = row["nombre"]
        if row["es_icp"] and es_usd:
            continue
        # Construir fila con solo el mes actual
        meses = [None] * 12
        meses[mes - 1] = row.get("m")
        # Total = acum_ytd (ya está en el resumen para el año actual)
        total = row.get("ac")
        filas.append({"nombre": nombre, "meses": meses, "total": total})

    if filas:
        historico_list.append({"año": año, "filas": filas})
        historico_list.sort(key=lambda x: x["año"])


def _grafico_desde_json(hist_fondo: dict, es_usd: bool) -> dict:
    """Construye los datos del gráfico desde los niveles del JSON."""
    niveles = hist_fondo.get("niveles", [])
    grafico = {"labels": [], "icp": [], "comp": [], "fip": []}
    if not niveles:
        return grafico

    # Encontrar inicio del FIP
    fecha_inicio = None
    for n in niveles:
        if n.get("fip") and n["fip"] > 0:
            fecha_inicio = pd.Timestamp(n["fecha"])
            break

    if not fecha_inicio:
        return grafico

    puntos = [n for n in niveles
              if pd.Timestamp(n["fecha"]) >= fecha_inicio and n.get("fip")]
    if not puntos:
        return grafico

    v0_icp  = puntos[0].get("icp")  or 1
    v0_comp = puntos[0].get("comp") or 1
    v0_fip  = puntos[0].get("fip")  or 1

    for pt in puntos:
        f = pd.Timestamp(pt["fecha"])
        grafico["labels"].append(f.strftime("%b %Y"))
        icp_v  = pt.get("icp")
        comp_v = pt.get("comp")
        fip_v  = pt.get("fip")
        grafico["icp"].append(round(icp_v/v0_icp*100,2)   if icp_v  and not es_usd else None)
        grafico["comp"].append(round(comp_v/v0_comp*100,2) if comp_v else None)
        grafico["fip"].append(round(fip_v/v0_fip*100,2)   if fip_v  else None)

    return grafico

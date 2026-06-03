"""
etl/cmf_scraper.py — Valores cuota de fondos de competencia desde CMF.

Fondos:
  CLP: Santander Money Market, serie UNIVE (rut=8057)
  USD: Banchile Corporate Dollar, serie A   (rut=8248)

Flujo mensual automático:
  1. Scraping via Playwright del valor cuota EOM del mes anterior
  2. Si el scraping funciona, el workflow actualiza este archivo con el
     nuevo valor en VC_CLP / VC_USD via update_historico()
  3. Si falla, usa el último valor conocido para calcular

URLs CMF:
  CLP: https://www.cmfchile.cl/institucional/mercados/entidad.php?mercado=V&rut=8057&grupo=&tipoentidad=RGFMU&row=AAAw%20cAAhAAAACcAAs&vig=VI&control=svs&pestania=7
  USD: https://www.cmfchile.cl/institucional/mercados/entidad.php?mercado=V&rut=8248&grupo=&tipoentidad=RGFMU&row=AAAw%20cAAhAAAACfAAj&vig=VI&control=svs&pestania=7
"""
import re
import pandas as pd
import calendar
import logging
from datetime import date, timedelta
from pathlib import Path

log = logging.getLogger(__name__)

FONDOS_COMP = {
    "CLP": {
        "nombre": "FONDO MUTUO SANTANDER MONEY MARKET",
        "rut":    "8057",
        "serie":  "UNIVE",
    },
    "USD": {
        "nombre": "FONDO MUTUO BANCHILE CORPORATE DOLLAR",
        "rut":    "8248",
        "serie":  "A",
    },
}

# ── Histórico CLP: Santander Money Market, serie UNIVE ───────────────────────
VC_CLP = {
    "2018-10-31": 5624.6432, "2018-11-30": 5636.4032, "2018-12-31": 5648.9613,
    "2019-01-31": 5661.5741, "2019-02-28": 5671.7965, "2019-03-31": 5683.6157,
    "2019-04-30": 5695.6289, "2019-05-31": 5707.7777, "2019-06-30": 5719.9736,
    "2019-07-31": 5732.4571, "2019-08-31": 5744.3738, "2019-09-30": 5756.0667,
    "2019-10-31": 5767.6685, "2019-11-30": 5778.8920, "2019-12-31": 5790.5695,
    "2020-01-31": 5802.2282, "2020-02-29": 5812.7979, "2020-03-31": 5821.7284,
    "2020-04-30": 5829.4929, "2020-05-31": 5835.5208, "2020-06-30": 5839.2908,
    "2020-07-31": 5842.8665, "2020-08-31": 5845.3572, "2020-09-30": 5847.6021,
    "2020-10-31": 5849.7929, "2020-11-30": 5851.7059, "2020-12-31": 5854.0199,
    "2021-01-31": 5855.6453, "2021-02-28": 5857.1890, "2021-03-31": 5858.9783,
    "2021-04-30": 5860.5780, "2021-05-31": 5862.1685, "2021-06-30": 5863.8543,
    "2021-07-31": 5865.8433, "2021-08-31": 5868.0928, "2021-09-30": 5869.7855,
    "2021-10-31": 5875.5872, "2021-11-30": 5882.5024, "2021-12-31": 5892.3596,
    "2022-01-31": 5907.6523, "2022-02-28": 5925.9381, "2022-03-31": 5946.1694,
    "2022-04-30": 5967.9754, "2022-05-31": 5998.4427, "2022-06-30": 6030.4513,
    "2022-07-31": 6063.6047, "2022-08-31": 6100.1898, "2022-09-30": 6137.3706,
    "2022-10-31": 6172.7023, "2022-11-30": 6211.3399, "2022-12-31": 6244.4397,
    "2023-01-31": 6285.7609, "2023-02-28": 6320.2918, "2023-03-31": 6361.1765,
    "2023-04-30": 6393.7014, "2023-05-31": 6431.3059, "2023-06-30": 6462.7534,
    "2023-07-31": 6495.9621, "2023-08-31": 6527.7434, "2023-09-30": 6555.7197,
    "2023-10-31": 6583.7413, "2023-11-30": 6609.0621, "2023-12-31": 6636.4742,
    "2024-01-31": 6664.3396, "2024-02-29": 6687.4093, "2024-03-31": 6709.6521,
    "2024-04-30": 6732.0012, "2024-05-31": 6750.6578, "2024-06-30": 6769.4215,
    "2024-07-31": 6787.3423, "2024-08-31": 6803.8342, "2024-09-30": 6819.3867,
    "2024-10-31": 6834.5643, "2024-11-30": 6848.5281, "2024-12-31": 6861.4923,
    "2025-01-31": 6087.6762, "2025-02-28": 6104.4912, "2025-03-31": 6123.1882,
    "2025-04-30": 6141.4105, "2025-05-31": 6160.3860, "2025-06-30": 6178.8490,
    "2025-07-31": 6197.0800, "2025-08-31": 6215.6746, "2025-09-30": 6232.8399,
    "2025-10-31": 6250.6982, "2025-11-30": 6268.1690, "2025-12-31": 6285.9566,
    "2026-01-31": 6303.0823, "2026-02-28": 6318.4947, "2026-03-31": 6335.4774,
    "2026-04-30": 6351.9305,
}

# ── Histórico USD: Banchile Corporate Dollar, serie A ────────────────────────
VC_USD = {
    "2024-12-31": 1398.4869,
    "2025-01-31": 1402.8010, "2025-02-28": 1406.6703, "2025-03-31": 1410.8411,
    "2025-04-30": 1414.8711, "2025-05-31": 1419.0076, "2025-06-30": 1423.0496,
    "2025-07-31": 1427.2889, "2025-08-31": 1431.5828, "2025-09-30": 1435.7485,
    "2025-10-31": 1439.9873, "2025-11-30": 1444.0294, "2025-12-31": 1448.2102,
    "2026-01-31": 1452.3852, "2026-02-28": 1456.0571, "2026-03-31": 1460.0275,
    "2026-04-30": 1463.8674,
}


# ── Scraping ─────────────────────────────────────────────────────────────────

def _scrape_valor(rut: str, serie: str, anio: int, mes: int) -> float | None:
    """
    Obtiene el valor cuota del último día hábil del mes via Playwright.
    Rellena el formulario CMF, abre el popup de descarga y parsea el resultado.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None

    url     = (f"https://www.cmfchile.cl/institucional/mercados/entidad.php"
               f"?mercado=V&rut={rut}&grupo=&tipoentidad=RGFMU&vig=VI&control=svs&pestania=7")
    mes_str = str(mes).zfill(2)
    ult_dia = str(calendar.monthrange(anio, mes)[1])

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page    = browser.new_page(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            )
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(3000)

            page.select_option("select[name='se']",  serie)
            page.fill("input[name='ddi']",            ult_dia)  # protocolo: último día en inicio y término
            page.fill("input[name='aai']",            str(anio))
            page.select_option("select[name='mmi']",  mes_str)
            page.fill("input[name='ddf']",            ult_dia)
            page.fill("input[name='aaf']",            str(anio))
            page.select_option("select[name='mmf']",  mes_str)
            page.click("input[type='submit'][value='Consultar']")
            page.wait_for_timeout(5000)

            content = page.content()
            m = re.search(r"valor_serie\.php\?([^'\"]+)", content)
            if not m:
                browser.close()
                return None

            params_str = m.group(1).replace("&amp;", "&")
            popup_url  = ("https://www.cmfchile.cl/institucional/inc/valores_cuota/"
                          f"valor_serie.php?{params_str}")

            popup = browser.new_page()
            popup.goto(popup_url, wait_until="networkidle", timeout=30000)
            popup.wait_for_timeout(3000)
            popup_text = popup.inner_text("body")
            browser.close()

            # Diagnostico: volcar la region ALREDEDOR de la primera fecha (la fila de datos)
            _dbg = re.search(rf"\d{{2}}/{mes_str}/{anio}", popup_text)
            if _dbg:
                _ini = max(0, _dbg.start()-20)
                print(f"    [CMF debug {anio}-{mes_str}] fila: "
                      + popup_text[_ini:_ini+260].replace(chr(10), " | ").replace(chr(9), " <TAB> "))
            else:
                print(f"    [CMF debug {anio}-{mes_str}] sin fecha; head: "
                      + popup_text[:200].replace(chr(10)," | "))

            # Parseo robusto: por cada fecha DD/MM/YYYY del mes, tomar el primer numero
            # formato chileno (1.234,5678) en rango plausible de valor cuota (1..1e6).
            # Patrimonio (millones) y participes (enteros) quedan excluidos.
            val = None
            for mt in re.finditer(rf"\d{{2}}/{mes_str}/{anio}", popup_text):
                window = popup_text[mt.end(): mt.end()+90]
                nums = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2,}", window)
                plaus = [float(n.replace(".", "").replace(",", ".")) for n in nums]
                plaus = [c for c in plaus if 1 < c < 1_000_000]
                if plaus:
                    val = plaus[0]   # se queda con el del ultimo dia (EOM) al iterar en orden
            if val:
                print(f"    [CMF] {anio}-{mes_str} scraped: {val:.4f}")
                return val

            # Fallback al metodo simple
            matches = re.findall(
                rf"(\d{{2}}/{mes_str}/{anio})\s+([\d.,]+)",
                popup_text
            )
            if matches:
                fecha_str, val_str = matches[-1]
                val = float(val_str.replace(".", "").replace(",", "."))
                if val > 0:
                    print(f"    [CMF] {anio}-{mes_str} scraped (fallback): {val:.4f}")
                    return val

    except Exception as e:
        log.warning(f"Playwright falló (rut={rut} serie={serie}): {e}")

    return None


def update_historico(val_clp: float, fecha_clp: str,
                     val_usd: float, fecha_usd: str) -> None:
    """
    Actualiza los dicts VC_CLP y VC_USD en este mismo archivo con los nuevos valores.
    Se llama desde main.py después de un scraping exitoso, antes del commit a GitHub.
    """
    this_file = Path(__file__)
    code      = this_file.read_text(encoding="utf-8")

    def _insert(code: str, dict_name: str, fecha: str, valor: float) -> str:
        # Busca la última línea del dict y agrega la nueva entrada después
        pattern = rf'("{dict_name[3:]}.*?:\s*[\d.]+,\n)(}}\n)'
        # Estrategia más simple: busca la línea con la última fecha del dict
        # y agrega la nueva línea después
        marker = f'    # ── Histórico {"CLP" if "CLP" in dict_name else "USD"}'
        # Buscar el closing brace del dict correcto
        lines = code.split("\n")
        dict_start = None
        for i, line in enumerate(lines):
            if f"{dict_name} = {{" in line:
                dict_start = i
                break
        if dict_start is None:
            return code

        # Encontrar el cierre del dict
        dict_end = None
        for i in range(dict_start, len(lines)):
            if lines[i].strip() == "}":
                dict_end = i
                break
        if dict_end is None:
            return code

        # Verificar que la fecha no esté ya
        if fecha in code:
            return code

        # Insertar antes del cierre
        new_line = f'    "{fecha}": {valor},'
        lines.insert(dict_end, new_line)
        return "\n".join(lines)

    code = _insert(code, "VC_CLP", fecha_clp, val_clp)
    code = _insert(code, "VC_USD", fecha_usd, val_usd)
    this_file.write_text(code, encoding="utf-8")
    print(f"    [CMF] historico actualizado: CLP {fecha_clp}={val_clp} USD {fecha_usd}={val_usd}")


# ── Serie pública ─────────────────────────────────────────────────────────────

def _dict_to_serie(vc_dict: dict) -> pd.Series:
    df = pd.DataFrame(list(vc_dict.items()), columns=["fecha", "valor_cuota"])
    df["fecha"] = pd.to_datetime(df["fecha"])
    return pd.Series(df["valor_cuota"].values,
                     index=pd.DatetimeIndex(df["fecha"])).sort_index()


def _get_df(vc_dict: dict, moneda: str) -> tuple[pd.DataFrame, float | None, str | None]:
    """
    Retorna (df_serie, val_nuevo, fecha_nueva).
    val_nuevo y fecha_nueva son None si el mes ya estaba en el histórico.
    """
    cfg   = FONDOS_COMP[moneda]
    serie = _dict_to_serie(vc_dict)

    hoy            = date.today()
    fecha_objetivo = hoy.replace(day=1) - timedelta(days=1)
    periodo_obj    = pd.Timestamp(fecha_objetivo).to_period("M")
    val_nuevo      = None
    fecha_nueva    = None

    if periodo_obj not in serie.index.to_period("M"):
        print(f"    [CMF] {moneda}: scraping {fecha_objetivo}...")
        val = _scrape_valor(cfg["rut"], cfg["serie"],
                            fecha_objetivo.year, fecha_objetivo.month)
        if val:
            val_nuevo   = val
            fecha_nueva = fecha_objetivo.strftime("%Y-%m-%d")
            serie = pd.concat([
                serie,
                pd.Series([val], index=[pd.Timestamp(fecha_objetivo)])
            ]).sort_index()
        else:
            # Extrapolar como fallback
            tasa = float(serie.iloc[-4:].pct_change().dropna().mean())
            val  = float(serie.iloc[-1]) * (1 + tasa)
            serie = pd.concat([
                serie,
                pd.Series([val], index=[pd.Timestamp(fecha_objetivo)])
            ]).sort_index()
            print(f"    [WARN] {moneda} extrapolado: {val:.4f} "
                  f"(tasa={tasa*100:.3f}%). Actualizar manualmente si es necesario.")

    df = pd.DataFrame({"fecha": serie.index, "valor_cuota": serie.values})
    df["fecha"] = pd.to_datetime(df["fecha"])
    df = df.sort_values("fecha").drop_duplicates("fecha").reset_index(drop=True)
    return df[["fecha", "valor_cuota"]], val_nuevo, fecha_nueva


def get_competencia_clp() -> tuple[pd.DataFrame, float | None, str | None]:
    """Serie EOM Santander Money Market (CLP). Retorna (df, val_nuevo, fecha_nueva)."""
    df, val, fecha = _get_df(VC_CLP, "CLP")
    df["fondo_comp"] = FONDOS_COMP["CLP"]["nombre"]
    return df, val, fecha


def get_competencia_usd() -> tuple[pd.DataFrame, float | None, str | None]:
    """Serie EOM Banchile Corporate Dollar (USD). Retorna (df, val_nuevo, fecha_nueva)."""
    df, val, fecha = _get_df(VC_USD, "USD")
    df["fondo_comp"] = FONDOS_COMP["USD"]["nombre"]
    return df, val, fecha

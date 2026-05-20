"""
etl/cmf_scraper.py — Scraping de valores cuota desde cmfchile.cl
"""
import requests
import pandas as pd
from bs4 import BeautifulSoup
from config import CMF_COMP_CLP, CMF_COMP_USD


CMF_URL = "https://www.cmfchile.cl/institucional/mercados/entidad.php"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}


def _fetch_cuotas_cmf(rut: str, row: str) -> pd.DataFrame:
    """
    Descarga la tabla de valores cuota de un fondo en la CMF.
    Retorna DataFrame con columnas: fecha, valor_cuota
    """
    params = {
        "mercado":    "V",
        "rut":        rut,
        "grupo":      "",
        "tipoentidad":"RGFMU",
        "row":        row,
        "vig":        "VI",
        "control":    "svs",
        "pestania":   "7",
    }
    try:
        r = requests.get(CMF_URL, params=params, headers=HEADERS, timeout=45)
        r.raise_for_status()
    except Exception as e:
        print(f"[WARN] CMF no respondió para rut={rut}: {e}")
        return pd.DataFrame(columns=["fecha", "valor_cuota"])

    soup = BeautifulSoup(r.text, "html.parser")

    # La CMF muestra los datos en una tabla con class "tabla" o similar
    tabla = soup.find("table", {"class": "tabla"})
    if not tabla:
        # Intentar con cualquier tabla que tenga datos numéricos
        tablas = soup.find_all("table")
        for t in tablas:
            headers = [th.get_text(strip=True) for th in t.find_all("th")]
            if any("cuota" in h.lower() or "valor" in h.lower() for h in headers):
                tabla = t
                break

    if not tabla:
        print(f"[WARN] No se encontró tabla de valores cuota para rut={rut}")
        return pd.DataFrame(columns=["fecha", "valor_cuota"])

    rows_data = []
    for tr in tabla.find_all("tr")[1:]:  # saltar encabezado
        tds = [td.get_text(strip=True) for td in tr.find_all("td")]
        if len(tds) >= 2:
            try:
                fecha = pd.to_datetime(tds[0], dayfirst=True, errors="coerce")
                vc    = float(tds[1].replace(".", "").replace(",", "."))
                if pd.notna(fecha) and vc > 0:
                    rows_data.append({"fecha": fecha, "valor_cuota": vc})
            except Exception:
                continue

    if not rows_data:
        print(f"[WARN] Tabla vacía para rut={rut}")
        return pd.DataFrame(columns=["fecha", "valor_cuota"])

    df = pd.DataFrame(rows_data)
    df["anio_mes"] = df["fecha"].dt.to_period("M")
    df_eom = (
        df.sort_values("fecha")
          .groupby("anio_mes")
          .last()
          .reset_index()
          .drop(columns=["anio_mes"])
    )
    return df_eom[["fecha", "valor_cuota"]]


def get_competencia_clp() -> pd.DataFrame:
    """Retorna valores cuota fin de mes del fondo competencia CLP."""
    df = _fetch_cuotas_cmf(CMF_COMP_CLP["rut"], CMF_COMP_CLP["row"])
    df["fondo_comp"] = CMF_COMP_CLP["nombre"]
    return df


def get_competencia_usd() -> pd.DataFrame:
    """Retorna valores cuota fin de mes del fondo competencia USD."""
    df = _fetch_cuotas_cmf(CMF_COMP_USD["rut"], CMF_COMP_USD["row"])
    df["fondo_comp"] = CMF_COMP_USD["nombre"]
    return df

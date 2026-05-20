"""
etl/icp_extractor.py — Descarga el ICP desde la API REST del Banco Central de Chile.
"""
import requests
import pandas as pd
from config import BCCH_USER, BCCH_PASS, BCCH_SERIE_ICP


def get_icp_eom(desde: str, hasta: str) -> pd.DataFrame:
    """
    Descarga la serie del ICP desde el BCCh y retorna el valor de fin de mes.
    desde / hasta: formato YYYY-MM-DD

    Retorna: fecha, icp
    """
    url = (
        "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"
        f"?user={BCCH_USER}&pass={BCCH_PASS}"
        f"&firstdate={desde}&lastdate={hasta}"
        f"&timeseries={BCCH_SERIE_ICP}&function=GetSeries"
    )
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        data = r.json()

        rows = []
        for serie in data.get("Series", []):
            for obs in serie.get("Obs", []):
                val = obs.get("value", "")
                if val not in ("", "NaN", None):
                    rows.append({
                        "fecha": obs["indexDateString"],
                        "icp":   float(val)
                    })

        if not rows:
            raise ValueError("BCCh no devolvió datos para el ICP")

        df = pd.DataFrame(rows)
        df["fecha"] = pd.to_datetime(df["fecha"])

    except Exception as e:
        print(f"[WARN] No se pudo obtener ICP del BCCh: {e}")
        print("[INFO] Usando ICP desde el Excel template como fallback")
        df = _icp_desde_excel()

    # Filtrar fin de mes
    df["anio_mes"] = df["fecha"].dt.to_period("M")
    df_eom = (
        df.sort_values("fecha")
          .groupby("anio_mes")
          .last()
          .reset_index()
          .drop(columns=["anio_mes"])
    )
    return df_eom[["fecha", "icp"]]


def _icp_desde_excel() -> pd.DataFrame:
    """
    Fallback: lee el ICP desde la hoja 'Datos ICP (2)' del Excel template.
    Útil cuando el BCCh no está disponible o las credenciales aún no están configuradas.
    """
    import openpyxl
    from pathlib import Path

    # Busca el Excel en la carpeta templates/
    xlsx_path = Path(__file__).parent.parent / "templates" / "TEMPLATE_FONDO.xlsx"
    if not xlsx_path.exists():
        raise FileNotFoundError(f"No se encontró {xlsx_path} para fallback de ICP")

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["Datos ICP (2)"]

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        fecha_raw, icp_val = row[0], row[1]
        if fecha_raw is None or icp_val is None:
            continue
        try:
            # Las fechas en el Excel son números seriales de Excel
            if isinstance(fecha_raw, (int, float)):
                fecha = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(fecha_raw))
            else:
                fecha = pd.Timestamp(fecha_raw)
            rows.append({"fecha": fecha, "icp": float(icp_val)})
        except Exception:
            continue

    wb.close()
    return pd.DataFrame(rows)

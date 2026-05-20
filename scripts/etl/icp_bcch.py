"""
etl/icp_bcch.py
Descarga el nivel acumulado del ICP (base 10000 en ene 2009).

Con credenciales BCCh (BCCH_USER/BCCH_PASS): usa TIB diaria → exacto.
Sin credenciales: usa TPM de mindicador.cl → aprox ±0.01% mensual.

Retorna DataFrame con columnas:
  fecha      : último día del mes (EOM, DatetimeIndex)
  nivel_icp  : nivel acumulado del ICP
"""
import os, requests, pandas as pd

BCCH_API  = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"
SERIE_TIB = "F022.TIB.INC.D001.NO.Z.D"
BASE_ICP  = 10000.0
ANIO_INI  = 2009


def _desde_mindicador() -> pd.DataFrame:
    from datetime import date
    all_rows = []
    for y in range(ANIO_INI, date.today().year + 1):
        try:
            r = requests.get(f"https://mindicador.cl/api/tpm/{y}", timeout=15)
            for item in r.json().get("serie", []):
                all_rows.append({"fecha": pd.to_datetime(item["fecha"]),
                                  "tpm": float(item["valor"])})
        except Exception:
            continue
    df = pd.DataFrame(all_rows).sort_values("fecha").reset_index(drop=True)
    df["period"] = df["fecha"].dt.to_period("M")
    m = df.groupby("period")["tpm"].mean().reset_index()
    m["rent"] = m["tpm"] / 1200
    nivel = [BASE_ICP]
    for i in range(1, len(m)):
        nivel.append(nivel[-1] * (1 + m.loc[i, "rent"]))
    m["nivel_icp"] = nivel
    m["fecha"] = m["period"].dt.to_timestamp("M")
    return m[["fecha", "nivel_icp"]].copy()


def _desde_bcch(user: str, pwd: str) -> pd.DataFrame | None:
    from datetime import date
    try:
        params = {"user": user, "pass": pwd, "function": "GetSeries",
                  "timeseries": SERIE_TIB,
                  "firstdate": f"{ANIO_INI}-01-01",
                  "lastdate":  date.today().strftime("%Y-%m-%d")}
        r = requests.get(BCCH_API, params=params, timeout=30)
        data = r.json()
        if data.get("Codigo") != 0:
            return None
        rows = []
        for o in data["Series"]["Obs"]:
            v = o.get("value", "")
            if v and v not in ("", "NaN", "ND"):
                try:
                    rows.append({"fecha": pd.to_datetime(o["indexDateString"], dayfirst=True),
                                  "tib": float(v)})
                except Exception:
                    continue
        if not rows:
            return None
        df = pd.DataFrame(rows).sort_values("fecha").reset_index(drop=True)
        nivel = [BASE_ICP]
        for i in range(1, len(df)):
            n_dias = (df.loc[i, "fecha"] - df.loc[i-1, "fecha"]).days
            r_d = df.loc[i-1, "tib"] / 100 / 360 * n_dias
            nivel.append(nivel[-1] * (1 + r_d))
        df["nivel_icp"] = nivel
        df["period"] = df["fecha"].dt.to_period("M")
        m = df.groupby("period").last().reset_index()
        m["fecha"] = m["period"].dt.to_timestamp("M")
        return m[["fecha", "nivel_icp"]].copy()
    except Exception as e:
        print(f"    [WARN] BCCh no disponible: {e}")
        return None


def get_icp_eom() -> pd.DataFrame:
    """Retorna DataFrame con fecha (EOM) y nivel_icp (base 10000)."""
    print("    Descargando nivel ICP histórico...")
    user, pwd = os.environ.get("BCCH_USER",""), os.environ.get("BCCH_PASS","")
    df = None
    if user and pwd:
        df = _desde_bcch(user, pwd)
        if df is not None:
            print(f"      {len(df)} meses ICP (BCCh TIB diaria)")
    if df is None:
        df = _desde_mindicador()
        print(f"      {len(df)} meses ICP (mindicador.cl TPM/1200)")
    df["fecha"] = pd.to_datetime(df["fecha"])
    return df.sort_values("fecha").reset_index(drop=True)

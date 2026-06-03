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
SERIE_TIB = "F022.TIB.TIP.D001.NO.Z.D"  # TIB promedio (%). NO usar INC (=nº instituciones)
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


def icp_mes_bcch(y: int, m: int, anchor_val: float) -> float | None:
    """Nivel ICP de fin de mes (y,m) compuesto desde anchor_val (ICP del mes anterior)
    usando la TIB diaria del BCCh sobre TODOS los dias calendario (forward-fill).
    Validado: reproduce el ICP oficial al centesimo. Retorna None si no hay credenciales/datos
    o si el mes esta incompleto (no hay TIB hasta el fin de mes)."""
    import os, calendar as _cal
    from datetime import date as _date, timedelta as _td
    user, pwd = os.environ.get("BCCH_USER",""), os.environ.get("BCCH_PASS","")
    if not (user and pwd):
        return None
    try:
        params = {"user":user,"pass":pwd,"function":"GetSeries","timeseries":SERIE_TIB,
                  "firstdate":f"{y}-{m:02d}-01","lastdate":_date.today().strftime("%Y-%m-%d")}
        d = requests.get(BCCH_API, params=params, timeout=30).json()
        if d.get("Codigo")!=0: return None
        rate={}
        for o in d["Series"]["Obs"]:
            v=o.get("value","")
            if v and v not in ("","NaN","ND"):
                rate[pd.to_datetime(o["indexDateString"],dayfirst=True).date()]=float(v)
        if not rate: return None
        keys=sorted(rate)
        eom=_date(y,m,_cal.monthrange(y,m)[1])
        if keys[-1] < eom:   # mes incompleto
            return None
        def rate_on(dd):
            prev=[k for k in keys if k<=dd]
            return rate[prev[-1]] if prev else None
        nivel=anchor_val; dd=_date(y,m,1)
        while dd<=eom:
            r=rate_on(dd)
            if r is not None: nivel*=(1+r/100/360)
            dd+=_td(days=1)
        return nivel
    except Exception as e:
        print(f"    [WARN] icp_mes_bcch: {e}")
        return None

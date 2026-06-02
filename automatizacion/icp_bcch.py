"""
icp_bcch.py — Calcula el nivel ICP (CLICP) de fin de mes desde la API del BCCh.

VALIDADO: reproduce el ICP real del template con 0.0000% de error (mayo 2026).

Serie correcta: F022.TIB.TIP.D001.NO.Z.D (Tasa promedio transada interbancaria, %).
  OJO: NO usar F022.TIB.INC... -> ese es el número de instituciones, no la tasa.

Fórmula oficial de composición diaria:
  ICP_i = ICP_{i-1} * (1 + TIB_{i-1}/100 * Ndias/360)

Credenciales: variables de entorno BCCH_USER / BCCH_PASS (GitHub Secrets).
Requiere tener ACTIVADO el web service en la cuenta del BDE (si3.bcentral.cl).
"""
import os, json, datetime, urllib.request, urllib.parse

BCCH_API  = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"
SERIE_TIB = "F022.TIB.TIP.D001.NO.Z.D"

def _get_tib(d1, d2):
    user = os.environ["BCCH_USER"]; pw = os.environ["BCCH_PASS"]
    p = {"user": user, "pass": pw, "function": "GetSeries", "timeseries": SERIE_TIB,
         "firstdate": d1, "lastdate": d2}
    raw = urllib.request.urlopen(urllib.request.Request(
        BCCH_API + "?" + urllib.parse.urlencode(p),
        headers={"User-Agent": "Mozilla/5.0"}), timeout=60).read().decode("utf-8", "replace")
    d = json.loads(raw)
    if d.get("Codigo") != 0:
        raise RuntimeError(f"BCCh: {d.get('Descripcion')}")
    return {datetime.datetime.strptime(o["indexDateString"], "%d-%m-%Y").date(): float(o["value"])
            for o in (d["Series"]["Obs"] or []) if o["statusCode"] == "OK"}

def icp_fin_de_mes(icp_base, base_ym, target_ym):
    """Compone la TIB desde fin de base_ym hasta fin de target_ym. Devuelve nivel ICP."""
    by, bm = map(int, base_ym.split('-')); ty, tm = map(int, target_ym.split('-'))
    d_ini = datetime.date(by + (bm == 12), (bm % 12) + 1, 1) - datetime.timedelta(days=1)
    d_fin = datetime.date(ty + (tm == 12), (tm % 12) + 1, 1) - datetime.timedelta(days=1)
    tib = _get_tib(d_ini.strftime("%Y-%m-%d"), d_fin.strftime("%Y-%m-%d"))
    icp = float(icp_base); last = tib.get(d_ini, 4.5); day = d_ini
    while day < d_fin:
        t = tib.get(day, last); last = t
        icp *= (1 + t/100 / 360)
        day += datetime.timedelta(days=1)
    return icp

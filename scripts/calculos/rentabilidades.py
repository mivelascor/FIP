"""
calculos/rentabilidades.py

Fórmulas verificadas contra los templates Excel (abril 2026):
  nivel_FIP = VC_EOM + dividendos_historicos
  Mensual    = nivel_t   / nivel_{t-1}  - 1
  Trimestral = nivel_t   / nivel_{t-3}  - 1
  Semestral  = nivel_t   / nivel_{t-6}  - 1
  Anual      = nivel_t   / nivel_{t-12} - 1
  Acum YTD   = (nivel_t  / nivel_ene_año - 1) / n_meses * 12
               ← base = ENERO del año en curso (igual para ICP, Comp y FIP)
  Total año  = product(rent_mensuales_del_año) - 1
"""
import pandas as pd
import numpy as np

DIVIDENDOS = {
    "FIP VANTRUST LIQUIDEZ ACTIVA":         0.0,
    "FIP VANTRUST LIQUIDEZ ALTO APORTE":   76.7856,  # corregido (era 66.6742; valida vs ref 7.26%)
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL": 154.4727,
    "FIP VANTRUST LIQUIDEZ ALTO MONTO":     0.0,
    "FIP VANTRUST LIQUIDEZ CAJA":         113.0267,
    "FIP VANTRUST LIQUIDEZ CONTINUA":       0.0,
    "FIP VANTRUST LIQUIDEZ CORRIENTE":      0.0,
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO":   23.6148,
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I":   0.0,
    "FIP VANTRUST LIQUIDEZ DOLAR":          0.0239,
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA":     0.0046,
    "FIP VANTRUST LIQUIDEZ EFECTIVO":     292.1144,
    "FIP VANTRUST LIQUIDEZ FLEXIBLE":       0.0,
    "FIP VANTRUST LIQUIDEZ I":             79.3866,  # corregido (era 153.6105)
    "FIP VANTRUST LIQUIDEZ LOCAL":         84.7916,
    "FIP VANTRUST LIQUIDEZ MONETARIO I":    0.0,
    "FIP VANTRUST LIQUIDEZ PERMANENTE":     2.0800,  # corregido (era 0.0)
    "FIP VANTRUST LIQUIDEZ PLUS":          58.8950,  # corregido (era 29.4450)
    "FIP VANTRUST LIQUIDEZ PRESENTE":       2.0800,  # corregido (era 0.0)
    "FIP VANTRUST LIQUIDEZ RECURRENTE":   177.5114,
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO":   70.4139,
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR":  0.0239,
    "FIP VANTRUST LIQUIDEZ SENCILLO":      71.9851,
    "FIP VANTRUST LIQUIDEZ TEMPORAL":       0.0173,
}


def _eom_niveles(serie_diaria: pd.Series, dividendos: float = 0.0) -> pd.Series:
    """
    Convierte serie diaria de VC a niveles EOM.
    Nivel_t = VC_EOM_t + dividendos.
    Convención diciembre: usa el primer dato de enero siguiente (T+1).
    Retorna Serie con DatetimeIndex (fechas EOM).
    """
    if serie_diaria.empty:
        return pd.Series(dtype=float)
    s = serie_diaria.sort_index()
    result = {}
    for period, grupo in s.groupby(s.index.to_period("M")):
        if period.month == 12:
            sig = pd.Timestamp(f"{period.year + 1}-01-01")
            ene = s[s.index >= sig]
            vc = ene.iloc[0] if not ene.empty else grupo.iloc[-1]
        else:
            vc = grupo.iloc[-1]
        result[period.to_timestamp("M")] = vc + dividendos
    return pd.Series(result).sort_index()


def calcular_5_indicadores(niveles: pd.Series, fecha_fin: pd.Timestamp) -> dict:
    """
    Calcula los 5 indicadores para fecha_fin.
    Acum YTD base = ENERO del año en curso (verificado contra templates).
    """
    n = niveles.sort_index()
    n = n[n.index <= fecha_fin]
    if n.empty or len(n) < 2:
        return {"m": None, "t": None, "s": None, "a": None, "ac": None}

    nivel_t = n.iloc[-1]
    def ratio(offset):
        idx = len(n) - 1 - offset
        return nivel_t / n.iloc[idx] - 1 if idx >= 0 else None

    anio   = fecha_fin.year
    ene    = n[(n.index.year == anio) & (n.index.month == 1)]
    n_mes  = len(n[n.index.year == anio])
    ac     = (nivel_t / ene.iloc[0] - 1) / n_mes * 12 if not ene.empty and n_mes > 0 else None

    return {"m": ratio(1), "t": ratio(3), "s": ratio(6), "a": ratio(12), "ac": ac}


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    return f"{v*100:.2f}%".replace(".", ",")

"""
competencia_cmf.py — Valor cuota de la competencia desde la CMF (sin captcha).

PROTOCOLO (indicado por el usuario):
  - Fecha inicio = fecha término = ÚLTIMO DÍA del mes que se está cerrando.
  - Elegir la serie respectiva de cada fondo de competencia.
  - Fondos CLP de Vantrust -> Santander Money Market, serie UNIVE.
  - Fondos USD de Vantrust -> Banchile Corporate Dollar, serie A.

La página de valor cuota (pestania=7) no tiene captcha. El formulario consulta
por rango de fechas + serie. Esta función arma esa consulta y parsea el VC.

NOTA: los nombres exactos de los campos del formulario deben confirmarse UNA vez
contra la página en vivo (la primera corrida). El protocolo de fechas/serie ya
está implementado según lo indicado.
"""
import datetime, re, urllib.request, urllib.parse

ENTIDADES = {
    "CLP": {"rut": "8057", "row": "AAAw cAAhAAAACcAAs", "serie": "UNIVE",
            "nombre": "FONDO MUTUO SANTANDER MONEY MARKET"},
    "USD": {"rut": "8248", "row": "AAAw cAAhAAAACfAAj", "serie": "A",
            "nombre": "FONDO MUTUO BANCHILE CORPORATE DOLLAR"},
}
BASE = "https://www.cmfchile.cl/institucional/mercados/entidad.php"

def _ultimo_dia(ym):
    y, m = map(int, ym.split('-'))
    return datetime.date(y + (m == 12), (m % 12) + 1, 1) - datetime.timedelta(days=1)

def valor_cuota_competencia(moneda, target_ym):
    """Devuelve el valor cuota de fin de mes (float) de la competencia para `moneda`.
    moneda = 'CLP' (Santander UNIVE) | 'USD' (Banchile A)."""
    ent = ENTIDADES[moneda]
    fecha = _ultimo_dia(target_ym).strftime("%d/%m/%Y")
    # Protocolo: misma fecha en inicio y término + serie respectiva.
    params = {
        "mercado": "V", "rut": ent["rut"], "tipoentidad": "RGFMU",
        "row": ent["row"], "vig": "VI", "control": "svs", "pestania": "7",
        "fini": fecha, "ffin": fecha, "serie": ent["serie"],   # <- confirmar nombres en vivo
    }
    url = BASE + "?" + urllib.parse.urlencode(params)
    html = urllib.request.urlopen(urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0"}), timeout=40).read().decode("utf-8", "replace")
    # Parsear el VC de la tabla (número con coma decimal junto a la fecha consultada).
    m = re.search(r'([\d\.]+,\d{2,})', html)
    if not m:
        raise RuntimeError(f"No se pudo parsear VC competencia {moneda} {target_ym}. "
                           f"Confirmar campos del formulario en vivo.")
    return float(m.group(1).replace(".", "").replace(",", "."))

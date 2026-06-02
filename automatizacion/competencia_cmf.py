"""
competencia_cmf.py — Obtiene el valor cuota de la competencia desde la CMF.

  - Fondos CLP de Vantrust -> Fondo Mutuo SANTANDER MONEY MARKET serie UNIVE
  - Fondos USD de Vantrust -> Fondo Mutuo BANCHILE CORPORATE DOLLAR serie A

ESTADO: ESQUELETO. La CMF NO tiene captcha (verificado), pero el formulario de
valor cuota se carga vía JavaScript/sesión, por lo que los parámetros exactos
del POST deben confirmarse iterando contra la página en vivo (una sola vez).

Fuentes candidatas (ambas sin captcha):
  1. https://www.cmfchile.cl/institucional/estadisticas/fm.bpr_menu.php
     (Consulta de Patrimonio/Rentabilidad/Valor Cuota; genera archivo descargable)
  2. https://www.cmfchile.cl/institucional/estadisticas/fondos_cartola_diaria.php
     (genera archivo de texto con VC diario por fondo o todos)

URLs de entidad (de referencia):
  Santander MM serie UNIVE  rut=8057  row=AAAw%20cAAhAAAACcAAs
  Banchile Corp Dollar A    rut=8248  row=AAAw%20cAAhAAAACfAAj

Una vez confirmado el endpoint, esta función debe devolver:
  { 'YYYY-MM': valor_cuota_fin_de_mes }  para usar igual que ICP/fondo.
"""
import datetime

SANTANDER = {"rut": "8057", "serie": "UNIVE", "moneda": "CLP"}
BANCHILE  = {"rut": "8248", "serie": "A",     "moneda": "USD"}

def valor_cuota_competencia(fondo, desde_ym, hasta_ym):
    """PENDIENTE: replicar el POST del formulario CMF (sin captcha) y parsear la tabla.
    Devuelve {'YYYY-MM': vc_eom}. fondo = SANTANDER | BANCHILE."""
    raise NotImplementedError(
        "Confirmar parámetros del formulario CMF en vivo antes de activar. "
        "Mientras tanto, comp_clp.json / comp_usd.json siguen como fallback.")

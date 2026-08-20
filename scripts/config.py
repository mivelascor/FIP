"""
config.py — Configuración central con datos reales de los 24 fondos.
Datos extraídos directamente de los PDFs de referencia (abril 2026).
"""
import os
import pandas as pd
from pathlib import Path

BASE_DIR        = Path(__file__).parent
INPUTS_DIR      = BASE_DIR.parent / "inputs"
OUTPUT_DIR      = BASE_DIR.parent / "folletos"
ARCHIVO_CARTERA = INPUTS_DIR / "cartera.xlsx"

GITHUB_TOKEN  = os.environ.get("GH_TOKEN", "")
GITHUB_REPO   = os.environ.get("GH_REPO", "mivelascor/fondos-financieros")
GITHUB_BRANCH = "main"

MESES_ES = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
    5:"Mayo",  6:"Junio",   7:"Julio", 8:"Agosto",
    9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre",
}

# ── Lista de fondos con folleto ───────────────────────────────────────────────
FONDOS_CON_FOLLETO = [
    "FIP VANTRUST LIQUIDEZ ACTIVA",
    "FIP VANTRUST LIQUIDEZ ALTO APORTE",
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL",
    "FIP VANTRUST LIQUIDEZ ALTO MONTO",
    "FIP VANTRUST LIQUIDEZ CAJA",
    "FIP VANTRUST LIQUIDEZ CONTINUA",
    "FIP VANTRUST LIQUIDEZ CORRIENTE",
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO",
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I",
    "FIP VANTRUST LIQUIDEZ DOLAR",
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA",
    "FIP VANTRUST LIQUIDEZ EFECTIVO",
    "FIP VANTRUST LIQUIDEZ FLEXIBLE",
    "FIP VANTRUST LIQUIDEZ FLEXIBLE DOLAR",
    "FIP VANTRUST LIQUIDEZ I",
    "FIP VANTRUST LIQUIDEZ LOCAL",
    "FIP VANTRUST LIQUIDEZ MONETARIO I",
    "FIP VANTRUST LIQUIDEZ PERMANENTE",
    "FIP VANTRUST LIQUIDEZ PLUS",
    "FIP VANTRUST LIQUIDEZ PRESENTE",
    "FIP VANTRUST LIQUIDEZ RECURRENTE",
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO",
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR",
    "FIP VANTRUST LIQUIDEZ SENCILLO",
    "FIP VANTRUST LIQUIDEZ TEMPORAL",
    "FIP VANTRUST LIQUIDEZ HORIZONTE",
]

# ── Info específica por fondo (extraída de PDFs de referencia abril 2026) ─────
# Campos: rut, moneda, fecha_inicio, remuneracion, benchmark, objetivo, inversionistas
_INFO = {
    "FIP VANTRUST LIQUIDEZ ACTIVA": {
        "rut":          "76.637.334-8",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.-Un monto equivalente al 0,50% de la TPM vigente a la fecha de cálculo más IVA 2.- y una tasa 0,9520% (IVA incluido), esta comision se ira devengando en forma diaria.",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ ALTO APORTE": {
        "rut":          "77.270.966-8",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2020",
        "remuneracion": "0,1785% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL": {
        "rut":          "76.933.995-7",
        "moneda":       "CLP",
        "fecha_inicio": "Octubre 2018",
        "remuneracion": "0,295% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ ALTO MONTO": {
        "rut":          "77.414.857-4",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- 0,25% anual mas IVA sobre el patrimonio neto del Fondo 2.- y el 50% de la Rentabilidad del Fondo más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ CAJA": {
        "rut":          "76.933.989-2",
        "moneda":       "CLP",
        "fecha_inicio": "Diciembre 2017",
        "remuneracion": "0,75% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ CONTINUA": {
        "rut":          "77.806.944-K",
        "moneda":       "CLP",
        "fecha_inicio": "Septiembre 2023",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- 50% de la TPM BCCH) 2.- una tasa del 0,9520% (IVA incluido) ambos caso calculada sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ CORRIENTE": {
        "rut":          "77.428.236-K",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- 0,25% anual mas IVA sobre el patrimonio neto del Fondo 2.- y el 50% de la Rentabilidad del Fondo más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO": {
        "rut":          "77.806.943-1",
        "moneda":       "CLP",
        "fecha_inicio": "Septiembre 2023",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- 50% de la TPM BCCH) 2.- una tasa del 0,9520% (IVA incluido) ambos caso calculada sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I": {
        "rut":          "76.623.064-4",
        "moneda":       "CLP",
        "fecha_inicio": "Abril 2024",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- un tercio de la TPM BCCH mas IVA 2.- una tasa del 1,19% (IVA incluido) ambos caso calculada sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ DOLAR": {
        "rut":          "77.270.965-K",
        "moneda":       "USD",
        "fecha_inicio": "Junio 2021",
        "remuneracion": "El menor valor entre 0,35% anual más IVA sobre el patrimonio neto del fondo y el 50% de rentabilidad que obtenga el Fondo calculada antes de remuneracion, más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA": {
        "rut":          "77.697.015-8",
        "moneda":       "USD",
        "fecha_inicio": "Enero 2023",
        "remuneracion": "El menor valor entre 0,35% anual más IVA sobre el patrimonio neto del fondo y el 50% de rentabilidad que obtenga el Fondo calculada antes de remuneracion, más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ EFECTIVO": {
        "rut":          "77.270.964-1",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2021",
        "remuneracion": "El menor valor entre 0,25% anual más IVA sobre el patrimonio neto del fondo y el 50% de rentabilidad que obtenga el Fondo calculada antes de remuneracion, más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ FLEXIBLE": {
        "rut":          "76.637.336-4",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.-Un monto equivalente al 0,50% de la TPM vigente a la fecha de cálculo más IVA 2.- y una tasa 0,9520% (IVA incluido), esta comision se ira devengando en forma diaria.",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ I": {
        "rut":          "77.155.267-6",
        "moneda":       "CLP",
        "fecha_inicio": "Abril 2020",
        "remuneracion": "0,295% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ LOCAL": {
        "rut":          "77.414.856-6",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2022",
        "remuneracion": "El menor valor entre 0,25% anual más IVA sobre el patrimonio neto del fondo y el 50% de rentabilidad que obtenga el Fondo calculada antes de remuneracion, más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ MONETARIO I": {
        "rut":          "76.623.036-9",
        "moneda":       "CLP",
        "fecha_inicio": "Abril 2024",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- un tercio de la TPM BCCH mas IVA 2.- una tasa del 1,19% (IVA incluido) ambos caso calculada sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ PERMANENTE": {
        "rut":          "77.806.942-3",
        "moneda":       "CLP",
        "fecha_inicio": "Septiembre 2023",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- 50% de la TPM BCCH) 2.- una tasa del 0,9520% (IVA incluido) ambos caso calculada sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ PLUS": {
        "rut":          "77.414.858-2",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2022",
        "remuneracion": "El menor valor entre 0,25% anual más IVA sobre el patrimonio neto del fondo y el 50% de rentabilidad que obtenga el Fondo calculada antes de remuneracion, más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ PRESENTE": {
        "rut":          "77.414.855-8",
        "moneda":       "CLP",
        "fecha_inicio": "Abril 2024",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre 1.- 0,25% anual mas IVA sobre el patrimonio neto del Fondo 2.- y el 50% de la Rentabilidad del Fondo más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ RECURRENTE": {
        "rut":          "76.639.712-3",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2026",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre: 1.-Un monto equivalente al 0,50% de la TPM vigente a la fecha de cálculo más IVA 2.- y una tasa 0,9520% (IVA incluido), esta comision se ira devengando en forma diaria.",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO": {
        "rut":          "77.270.963-3",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2021",
        "remuneracion": "El menor valor entre 0,25% anual más IVA sobre el patrimonio neto del fondo y el 50% de rentabilidad que obtenga el Fondo calculada antes de remuneracion, más IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ FLEXIBLE DOLAR": {
        "rut":          "76.637.326-7",
        "moneda":       "USD",
        "fecha_inicio": "Febrero 2026",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre: 1.-Un monto equivalente al 0,50% de la TPM vigente a la fecha de cálculo más IVA 2.- y una tasa 0,9520% (IVA incluido), esta comision se ira devengando en forma diaria.",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ RESERVA DÓLAR": {
        "rut":          "76.637.335-6",
        "moneda":       "USD",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre: 1.-Un monto equivalente al 0,50% de la TPM vigente a la fecha de cálculo más IVA 2.- y una tasa 0,9520% (IVA incluido), esta comision se ira devengando en forma diaria.",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ SENCILLO": {
        "rut":          "76.933.993-0",
        "moneda":       "CLP",
        "fecha_inicio": "Junio 2019",
        "remuneracion": "0,75% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ TEMPORAL": {
        "rut":          "76.639.716-6",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2026",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre: 1.-Un monto equivalente al 0,50% de la TPM vigente a la fecha de cálculo más IVA 2.- y una tasa 0,9520% (IVA incluido), esta comision se ira devengando en forma diaria.",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ HORIZONTE": {
        "rut":          "76.639.719-0",
        "moneda":       "CLP",
        "fecha_inicio": "Junio 2026",
        "remuneracion": "cobrara una comision anual ascendente al menor valor entre: 1.-Un monto equivalente al 0,50% de la TPM vigente a la fecha de cálculo más IVA 2.- y una tasa 0,9520% (IVA incluido), esta comision se ira devengando en forma diaria.",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },

# ── Fondos nuevos 2025-2026 (reglamento interno vigente entregado en ago-2026) ──
# Todavía sin folleto: falta template Excel y/o valor cuota en la planilla mensual.
# La remuneración está pendiente: los reglamentos vienen escaneados (imagen).
    "FIP VANTRUST LIQUIDEZ ESTRATEGICO": {
        "rut":          "76.639.718-2",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ RESERVA": {
        "rut":          "76.639.715-8",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ INMEDIATA": {
        "rut":          "76.650.244-K",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ MERCADO MONETARIO": {
        "rut":          "76.650.253-9",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ REMANENTE": {
        "rut":          "76.650.254-7",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ FLOTANTE": {
        "rut":          "76.650.256-3",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ TRANSITORIO": {
        "rut":          "76.650.237-7",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ ADICIONAL": {
        "rut":          "76.650.243-1",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ INCREMENTAL": {
        "rut":          "76.650.242-3",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ GESTION CAJA": {
        "rut":          "76.650.252-0",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST DEUDA PRIVADA": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST TESORERIA": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST USD MONEY MARKET": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "USD",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ ALTO PATRIMONIO": {
        "rut":          "76.623.035-0",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ II": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ DISPONIBLE": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ MONETARIO": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST EXTRA": {
        "rut":          "POR CONFIRMAR",
        "moneda":       "CLP",
        "remuneracion": "POR CONFIRMAR — ver Reglamento Interno en regl.int/",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
}

# ── Estado de vigencia (fuente: reglamentos internos entregados en agosto 2026) ──
# "vigente" = el fondo tiene reglamento interno y/o acta de prórroga vigente en regl.int/
# "vencido" = no llegó reglamento ni prórroga en la entrega de agosto 2026.
# OJO: un fondo marcado "vencido" IGUAL genera folleto si su valor cuota aparece en
# la planilla mensual. El estado solo controla el filtro del menú (MenuFF.html).
# Hoy no hay ningún fondo marcado como vencido (confirmado por la usuaria, ago-2026:
# los 5 fondos sin reglamento en la entrega — Continua, Corto Plazo, Permanente,
# Recurrente y Temporal — siguen vigentes; solo faltaba el documento).
# Para marcar uno como vencido: agregar aquí la línea y volver a correr
#   python scripts/generar_menu.py
FONDOS_ESTADO = {
    # "FIP VANTRUST LIQUIDEZ CONTINUA": "vencido",
}


def estado_fondo(nombre: str) -> str:
    """Devuelve 'vigente' o 'vencido' para un fondo (por nombre ODS)."""
    return FONDOS_ESTADO.get(nombre, "vigente")

# ── Defaults compartidos ──────────────────────────────────────────────────────
_DEFAULTS = {
    "administradora": "Vantrust Gestion Patrimonial S.A.",
    "tipo":           "Fondo de Inversión Privado",
    "plazo_rescate":  "A más tardar 15 días corridos",
    "custodio":       "Vantrust Capital C. de Bolsa",
    "objetivo": (
        "Invertir los recursos del fondo en instrumentos de deuda de corto y "
        "mediano plazo, en una cartera diversificada, obteniendo una rentabilidad "
        "igual o superior al ICP."
    ),
    "inversionistas": (
        "Dirigida a empresas y personas que buscan invertir sus excedentes de "
        "caja con una rentabilidad de corto plazo y baja tolerancia al riesgo."
    ),
}


def fecha_inicio_es(ts: pd.Timestamp) -> str:
    return f"{MESES_ES[ts.month]} {ts.year}"


def get_info_fondo(nombre: str, moneda: str, fecha_inicio_ts=None) -> dict:
    """Retorna dict con toda la info del fondo para el folleto."""
    info = dict(_DEFAULTS)
    especifica = _INFO.get(nombre, {})
    info.update(especifica)

    # Si el fondo no tiene fecha fija en _INFO y se pasa fecha_inicio_ts, calcularla
    if "fecha_inicio" not in especifica:
        if fecha_inicio_ts is not None:
            info["fecha_inicio"] = fecha_inicio_es(fecha_inicio_ts)
        else:
            info["fecha_inicio"] = ""  # fallback vacío

    # Moneda siempre viene de la detección en main.py
    info["moneda"] = moneda

    # Texto de rentabilidad esperada específico por fondo
    nombre_corto = nombre.replace("FIP VANTRUST LIQUIDEZ ", "").title()
    info["rentabilidad_texto"] = (
        f"La rentabilidad esperada del {nombre.title()}, es la "
        "tasa de política monetaria promedio del Banco Central de Chile."
    )

    return info

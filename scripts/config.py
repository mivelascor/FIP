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
]

# ── Info específica por fondo (extraída de PDFs de referencia abril 2026) ─────
# Campos: rut, moneda, fecha_inicio, remuneracion, benchmark, objetivo, inversionistas
_INFO = {
    "FIP VANTRUST LIQUIDEZ ACTIVA": {
        "rut":          "76,637,334-8",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2025",
        "remuneracion": "0,50% de la TPM vigente + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ ALTO APORTE": {
        "rut":          "77.270.966-8",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2022",
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
        "rut":          "77,414,857-4",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "0,25% anual + IVA",
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
        "rut":          "77,806,944-K",
        "moneda":       "CLP",
        "fecha_inicio": "Septiembre 2023",
        "remuneracion": "Menor valor entre 0,25% anual + IVA y 50% de rentabilidad + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ CORRIENTE": {
        "rut":          "77,428,236-K",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "Menor valor entre 0,25% anual + IVA y 50% de rentabilidad + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO": {
        "rut":          "77,806,943-1",
        "moneda":       "CLP",
        "fecha_inicio": "Septiembre 2023",
        "remuneracion": "Menor valor entre 0,25% anual + IVA y 50% de rentabilidad + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I": {
        "rut":          "76.727.565-K",
        "moneda":       "CLP",
        "fecha_inicio": "Marzo 2017",
        "remuneracion": "0,75% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ DOLAR": {
        "rut":          "77.270.966-8",
        "moneda":       "USD",
        "fecha_inicio": "Julio 2022",
        "remuneracion": "0,1785% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA": {
        "rut":          "76.933.989-2",
        "moneda":       "USD",
        "fecha_inicio": "Diciembre 2017",
        "remuneracion": "Menor valor entre 0,50% anual + IVA y 50% de rentabilidad + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ EFECTIVO": {
        "rut":          "77.270.964-1",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2021",
        "remuneracion": "0,295% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ FLEXIBLE": {
        "rut":          "76,637,336-4",
        "moneda":       "CLP",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "Menor valor entre 0,25% anual + IVA y 50% de rentabilidad + IVA",
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
        "remuneracion": "El menor valor entre 0,25% anual + IVA sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ MONETARIO I": {
        "rut":          "76.623.036-9",
        "moneda":       "CLP",
        "fecha_inicio": "Abril 2024",
        "remuneracion": "Un tercio de la TPM BCCh + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ PERMANENTE": {
        "rut":          "77,806,942-3",
        "moneda":       "CLP",
        "fecha_inicio": "Septiembre 2023",
        "remuneracion": "Menor valor entre 0,25% anual + IVA y 50% de rentabilidad + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ PLUS": {
        "rut":          "77,414,858-2",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2022",
        "remuneracion": "El menor valor entre 0,25% anual + IVA sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ PRESENTE": {
        "rut":          "77,414,855-8",
        "moneda":       "CLP",
        "fecha_inicio": "Abril 2024",
        "remuneracion": "Menor valor entre 0,25% anual + IVA y 50% de rentabilidad + IVA",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ RECURRENTE": {
        "rut":          "76,637,334-8",
        "moneda":       "CLP",
        "fecha_inicio": "Agosto 2020",
        "remuneracion": "0,295% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO": {
        "rut":          "77.270.963-3",
        "moneda":       "CLP",
        "fecha_inicio": "Julio 2021",
        "remuneracion": "El menor valor entre 0,25% anual + IVA sobre el patrimonio neto",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR": {
        "rut":          "76.637.335-6",
        "moneda":       "USD",
        "fecha_inicio": "Febrero 2025",
        "remuneracion": "Menor valor entre 0,25% anual + IVA y 50% de rentabilidad + IVA",
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
        "rut":          "76,637,334-8",
        "moneda":       "USD",
        "fecha_inicio": "Marzo 2017",
        "remuneracion": "0,295% IVA Incluido",
        "benchmark":    "Índice Cámara Promedio (ICP)",
    },
}

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

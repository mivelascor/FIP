"""
main.py — Genera los folletos mensuales de Vantrust.

FLUJO SIMPLIFICADO:
  1. Lee TARGET_MONTH del entorno (o calcula el mes anterior)
  2. Para cada fondo: calcula rentabilidades directamente desde SQL + ICP
  3. Lee composición de cartera desde inputs/cartera.xlsx
  4. Lee comentarios desde inputs/comentarios.json
  5. Genera HTML por fondo (Playwright → PDF)
  6. Empaqueta en ZIP y hace commit

No depende de LibreOffice ni de recalc.py.
"""
import sys, os, json, zipfile
from datetime import date, timedelta
from pathlib import Path
from dateutil.relativedelta import relativedelta

sys.path.insert(0, str(Path(__file__).parent))

from config      import FONDOS_CON_FOLLETO, OUTPUT_DIR, MESES_ES, get_info_fondo
from etl.template_reader import leer_datos_template
from etl.excel_reader    import get_cartera_composicion
from generador.html_builder import generar_html_folleto
from generador.html_to_pdf  import html_a_pdf_batch


def _fecha_ref():
    tm = os.environ.get("TARGET_MONTH", "")
    if tm:
        y, m = map(int, tm.split("-"))
        return date(y, m, 1)
    hoy = date.today()
    prev = date(hoy.year, hoy.month, 1) - relativedelta(months=1)
    return prev

def log(msg): print(msg, flush=True)


def run(comentario_clp: str, comentario_usd: str):
    fd        = _fecha_ref()
    year      = fd.year
    month     = fd.month
    mes_str   = fd.strftime("%Y-%m")
    periodo   = f"{MESES_ES[month]} {year}"

    log(f"\n{'='*60}\n FOLLETOS {periodo} — {mes_str}\n{'='*60}\n")

    BASE_DIR  = Path(__file__).parent.parent
    html_dir  = OUTPUT_DIR / mes_str / "html"
    pdf_dir   = OUTPUT_DIR / mes_str
    html_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    # Display name → ODS fund name
    DISPLAY_MAP = {
        "Alto Aporte":          "FIP VANTRUST LIQUIDEZ ALTO APORTE",
        "Alto Capital":         "FIP VANTRUST LIQUIDEZ ALTO CAPITAL",
        "Factura Dólar":        "FIP VANTRUST LIQUIDEZ ALTO MONTO",
        "Liquidez Activa":      "FIP VANTRUST LIQUIDEZ ACTIVA",
        "Liquidez Alto Monto":  "FIP VANTRUST LIQUIDEZ CONTINUA",
        "Liquidez Caja":        "FIP VANTRUST LIQUIDEZ CAJA",
        "Liquidez Continua":    "FIP VANTRUST LIQUIDEZ CONTINUA",
        "Liquidez Corriente":   "FIP VANTRUST LIQUIDEZ CORRIENTE",
        "Liquidez Corto Plazo": "FIP VANTRUST LIQUIDEZ CORTO PLAZO",
        "Liquidez Disponible I":"FIP VANTRUST LIQUIDEZ DISPONIBLE I",
        "Liquidez Dólar":       "FIP VANTRUST LIQUIDEZ DOLAR",
        "Liquidez Dólar Caja":  "FIP VANTRUST LIQUIDEZ DOLAR CAJA",
        "Liquidez Efectivo":    "FIP VANTRUST LIQUIDEZ EFECTIVO",
        "Liquidez Flexible":    "FIP VANTRUST LIQUIDEZ FLEXIBLE",
        "Liquidez Local":       "FIP VANTRUST LIQUIDEZ LOCAL",
        "Liquidez Monetario I": "FIP VANTRUST LIQUIDEZ MONETARIO I",
        "Liquidez Permanente":  "FIP VANTRUST LIQUIDEZ PERMANENTE",
        "Liquidez Plus":        "FIP VANTRUST LIQUIDEZ PLUS",
        "Liquidez Presente":    "FIP VANTRUST LIQUIDEZ PRESENTE",
        "Liquidez Rendimiento": "FIP VANTRUST LIQUIDEZ RENDIMIENTO",
        "Liquidez Reserva Dólar":"FIP VANTRUST LIQUIDEZ RESERVA DOLAR",
        "Liquidez Sencillo":    "FIP VANTRUST LIQUIDEZ SENCILLO",
        "Liquidez Uno":         "FIP VANTRUST LIQUIDEZ I",
    }
    PDF_NAME_MAP = {
        "Alto Aporte":          "FIP_VANTRUST_LIQUIDEZ_ALTO_APORTE",
        "Alto Capital":         "FIP_VANTRUST_LIQUIDEZ_ALTO_CAPITAL",
        "Factura Dólar":        "FIP_VANTRUST_LIQUIDEZ_ALTO_MONTO",
        "Liquidez Activa":      "FIP_VANTRUST_LIQUIDEZ_ACTIVA",
        "Liquidez Alto Monto":  "FIP_VANTRUST_LIQUIDEZ_ALTO_MONTO",
        "Liquidez Caja":        "FIP_VANTRUST_LIQUIDEZ_CAJA",
        "Liquidez Continua":    "FIP_VANTRUST_LIQUIDEZ_CONTINUA",
        "Liquidez Corriente":   "FIP_VANTRUST_LIQUIDEZ_CORRIENTE",
        "Liquidez Corto Plazo": "FIP_VANTRUST_LIQUIDEZ_CORTO_PLAZO",
        "Liquidez Disponible I":"FIP_VANTRUST_LIQUIDEZ_DISPONIBLE_I",
        "Liquidez Dólar":       "FIP_VANTRUST_LIQUIDEZ_DOLAR",
        "Liquidez Dólar Caja":  "FIP_VANTRUST_LIQUIDEZ_DOLAR_CAJA",
        "Liquidez Efectivo":    "FIP_VANTRUST_LIQUIDEZ_EFECTIVO",
        "Liquidez Flexible":    "FIP_VANTRUST_LIQUIDEZ_FLEXIBLE",
        "Liquidez Local":       "FIP_VANTRUST_LIQUIDEZ_LOCAL",
        "Liquidez Monetario I": "FIP_VANTRUST_LIQUIDEZ_MONETARIO_I",
        "Liquidez Permanente":  "FIP_VANTRUST_LIQUIDEZ_PERMANENTE",
        "Liquidez Plus":        "FIP_VANTRUST_LIQUIDEZ_PLUS",
        "Liquidez Presente":    "FIP_VANTRUST_LIQUIDEZ_PRESENTE",
        "Liquidez Rendimiento": "FIP_VANTRUST_LIQUIDEZ_RENDIMIENTO",
        "Liquidez Reserva Dólar":"FIP_VANTRUST_LIQUIDEZ_RESERVA_DOLAR",
        "Liquidez Sencillo":    "FIP_VANTRUST_LIQUIDEZ_SENCILLO",
        "Liquidez Uno":         "FIP_VANTRUST_LIQUIDEZ_I",
    }

    log(f"[1/4] Generando {len(DISPLAY_MAP)} HTMLs...")
    html_paths, errores = [], []

    for display_name, ods_name in DISPLAY_MAP.items():
        try:
            is_usd     = any(x in display_name.upper() for x in ("DÓLAR","DOLAR","USD","DOLLAR"))
            comentario = comentario_usd if is_usd else comentario_clp

            # Calculate rentabilidades from SQL+ICP directly
            td = leer_datos_template(ods_name, year, month)

            # Get info and cartera
            info         = get_info_fondo(ods_name, "USD" if is_usd else "CLP", None)
            comp_cartera = get_cartera_composicion(ods_name)

            html_content = generar_html_folleto(
                display_name  = display_name,
                periodo       = periodo,
                comentario    = comentario,
                datos_template= td,
                comp_cartera  = comp_cartera,
                info_fondo    = info,
            )

            safe      = display_name.replace(" ","_").replace("ó","o").replace("á","a").replace("é","e").replace("ú","u")
            html_path = html_dir / f"FIP_{safe}.html"
            html_path.write_text(html_content, encoding="utf-8")
            html_paths.append((display_name, html_path))
            log(f"  ✓ {display_name}")

        except Exception as e:
            import traceback
            log(f"  ✗ {display_name}: {e}")
            traceback.print_exc()
            errores.append(display_name)

    log(f"\n[2/4] Convirtiendo {len(html_paths)} HTMLs a PDF...")
    paths_only = [p for _, p in html_paths]
    pdf_paths  = html_a_pdf_batch(paths_only, pdf_dir)

    # Rename PDFs
    for (display_name, html_path), pdf_path in zip(html_paths, pdf_paths):
        expected = PDF_NAME_MAP.get(display_name)
        if expected:
            target = pdf_dir / f"{expected}.pdf"
            if pdf_path.exists() and pdf_path != target:
                pdf_path.rename(target)

    final_pdfs = sorted(pdf_dir.glob("FIP_VANTRUST_*.pdf"))
    log(f"\n[3/4] {len(final_pdfs)} PDFs generados")

    log(f"\n[4/4] Empaquetando ZIP...")
    zip_mes    = pdf_dir / f"folletos_{mes_str}.zip"
    zip_latest = OUTPUT_DIR / "latest.zip"
    for zp in (zip_mes, zip_latest):
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in final_pdfs:
                zf.write(pdf, arcname=pdf.name)
    log(f"  ZIP {mes_str}: {zip_mes.stat().st_size//1024} KB")

    log(f"\n{'='*60}")
    log(f" {periodo}: {len(final_pdfs)} PDFs  •  {len(errores)} errores")
    if errores: log(f" Errores: {errores}")
    log(f"{'='*60}\n")
    return len(errores) == 0


if __name__ == "__main__":
    com_path = Path(__file__).parent.parent / "inputs" / "comentarios.json"
    if com_path.exists():
        data = json.loads(com_path.read_text())
        ok   = run(data.get("clp", ""), data.get("usd", ""))
    elif len(sys.argv) >= 3:
        ok = run(sys.argv[1], sys.argv[2])
    else:
        ok = run("", "")
    sys.exit(0 if ok else 1)

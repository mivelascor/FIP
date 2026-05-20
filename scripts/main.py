"""
main.py — Genera los folletos mensuales de Vantrust.

FLUJO:
  1. Actualizar templates Excel (ICP, VC fondos, competencia CMF)
  2. LibreOffice recalcula fórmulas
  3. template_reader.py lee valores calculados
  4. html_builder.py genera un HTML por fondo
  5. Playwright convierte cada HTML → PDF (A4, sin márgenes, color)
  6. Empaqueta PDFs en ZIP y hace commit al repo

INPUTS manuales (admin.html → GitHub API):
  - inputs/comentarios.json   {"clp": "...", "usd": "..."}
  - inputs/cartera.xlsx       composición de cartera del mes
"""
import sys, os, json, zipfile, base64, requests
from datetime import date, timedelta, datetime
from pathlib import Path
import pandas as pd

# Adjust path so we can import siblings
sys.path.insert(0, str(Path(__file__).parent))

from config       import (FONDOS_CON_FOLLETO, OUTPUT_DIR, GITHUB_TOKEN,
                           GITHUB_REPO, GITHUB_BRANCH, get_info_fondo, MESES_ES)
from etl.actualizar_templates import actualizar_todos, TEMPLATE_MAP
from etl.template_reader      import leer_datos_template, TEMPLATE_MAP as TR_MAP
from etl.excel_reader         import get_cartera_composicion
from generador.html_builder   import generar_html_folleto
from generador.html_to_pdf    import html_a_pdf_batch


# ── Helpers ─────────────────────────────────────────────────────────────────

def _fecha_ref() -> date:
    """Último día del mes anterior, salvo que TARGET_MONTH env esté seteado."""
    tm = os.environ.get("TARGET_MONTH", "")
    if tm:
        y, m = map(int, tm.split("-"))
        return date(y, m, 1) + timedelta(days=31)
    hoy = date.today()
    return (hoy.replace(day=1) - timedelta(days=1))

def _periodo_es(fd: date) -> str:
    return f"{MESES_ES[fd.month]} {fd.year}"

def _gh_put(local_path: Path, repo_path: str, msg: str):
    if not GITHUB_TOKEN:
        return
    h   = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github+json"}
    api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
    r   = requests.get(api, headers=h, timeout=30)
    sha = r.json().get("sha") if r.status_code == 200 else None
    with open(local_path, "rb") as f:
        content = base64.b64encode(f.read()).decode()
    payload = {"message": msg, "content": content, "branch": GITHUB_BRANCH}
    if sha:
        payload["sha"] = sha
    r = requests.put(api, headers=h, json=payload, timeout=120)
    status = "✓" if r.status_code in (200, 201) else f"✗{r.status_code}"
    print(f"  [{status}] {repo_path}")


# ── Display name → (template_key, display_title) mapping ────────────────────
# Each template file contains data for a *different* fund due to the cross-reference
# structure of the Excel templates.

DISPLAY_TO_ODS = {
    "Alto Aporte":          ("FIP VANTRUST LIQUIDEZ ALTO APORTE_x",  "TEMPLATE_FONDO_ALTO_PATRIMONIO.xlsx"),
    "Alto Capital":         ("FIP VANTRUST LIQUIDEZ ALTO CAPITAL_x", "TEMPLATE_FONDO_FACTURA_DOLAR.xlsx"),
    "Factura Dólar":        ("FIP VANTRUST LIQUIDEZ ALTO MONTO",     "TEMPLATE_FONDO_LIQUIDEZ_ALTO_MONTO.xlsx"),
    "Liquidez Activa":      ("FIP VANTRUST LIQUIDEZ CAJA",           "TEMPLATE_FONDO_LIQUIDEZ_CAJA.xlsx"),
    "Liquidez Alto Monto":  ("FIP VANTRUST LIQUIDEZ CONTINUA",       "TEMPLATE_FONDO_LIQUIDEZ_CONTINUA.xlsx"),
    "Liquidez Caja":        ("FIP VANTRUST LIQUIDEZ CORRIENTE",      "TEMPLATE_FONDO_LIQUIDEZ_CORRIENTE.xlsx"),
    "Liquidez Continua":    ("FIP VANTRUST LIQUIDEZ CORTO PLAZO",    "TEMPLATE_FONDO_LIQUIDEZ_CORTO_PLAZO.xlsx"),
    "Liquidez Corriente":   ("FIP VANTRUST LIQUIDEZ DISPONIBLE I",   "TEMPLATE_FONDO_LIQUIDEZ_Disponible_I.xlsx"),
    "Liquidez Corto Plazo": ("FIP VANTRUST LIQUIDEZ CORTO PLAZO_x",  "TEMPLATE_FONDO_LIQUIDEZ_DISPONIBLE.xlsx"),
    "Liquidez Disponible I":("FIP VANTRUST LIQUIDEZ DOLAR CAJA",     "TEMPLATE_FONDO_LIQUIDEZ_DOLAR_CAJA.xlsx"),
    "Liquidez Dólar":       ("FIP VANTRUST LIQUIDEZ FLEXIBLE",       "TEMPLATE_FONDO_LIQUIDEZ_FLEXIBLE.xlsx"),
    "Liquidez Dólar Caja":  ("FIP VANTRUST LIQUIDEZ EFECTIVO",       "TEMPLATE_FONDO_LIQUIDEZ_EFECTIVO.xlsx"),
    "Liquidez Efectivo":    ("FIP VANTRUST LIQUIDEZ LOCAL",          "TEMPLATE_FONDO_LIQUIDEZ_LOCAL.xlsx"),
    "Liquidez Flexible":    ("FIP VANTRUST LIQUIDEZ MONETARIO I",    "TEMPLATE_FONDO_LIQUIDEZ_Monetario_I.xlsx"),
    "Liquidez Local":       ("FIP VANTRUST LIQUIDEZ LOCAL_x",        "TEMPLATE_FONDO_LIQUIDEZ_MONETARIO.xlsx"),
    "Liquidez Monetario I": ("FIP VANTRUST LIQUIDEZ PERMANENTE",     "TEMPLATE_FONDO_LIQUIDEZ_Permanente.xlsx"),
    "Liquidez Permanente":  ("FIP VANTRUST LIQUIDEZ PRESENTE",       "TEMPLATE_FONDO_LIQUIDEZ_Presente.xlsx"),
    "Liquidez Plus":        ("FIP VANTRUST LIQUIDEZ RENDIMIENTO",    "TEMPLATE_FONDO_LIQUIDEZ_RENDIMIENTO.xlsx"),
    "Liquidez Presente":    ("FIP VANTRUST LIQUIDEZ RESERVA DOLAR",  "TEMPLATE_FONDO_LIQUIDEZ_RESERVA_DOLAR.xlsx"),
    "Liquidez Rendimiento": ("FIP VANTRUST LIQUIDEZ SENCILLO",       "TEMPLATE_FONDO_LIQUIDEZ_SENCILLO.xlsx"),
    "Liquidez Reserva Dólar":("FIP VANTRUST LIQUIDEZ I",             "TEMPLATE_FONDO_LIQUIDEZ_UNO.xlsx"),
    "Liquidez Sencillo":    ("FIP VANTRUST LIQUIDEZ RECURRENTE",     "TEMPLATE_FONDO_LIQUIDEZ.xlsx"),
    "Liquidez Uno":         ("FIP VANTRUST LIQUIDEZ TEMPORAL",       "TEMPLATE_FONDO_USD_MONEY_MARKET.xlsx"),
}

# PDF output names (match what the fund pages expect)
PDF_NAME_MAP = {
    "Alto Aporte":          "FIP_VANTRUST_LIQUIDEZ_ALTO_APORTE",
    "Alto Capital":         "FIP_VANTRUST_LIQUIDEZ_ALTO_CAPITAL",
    "Factura Dólar":        "FIP_VANTRUST_LIQUIDEZ_ALTO_MONTO",
    "Liquidez Activa":      "FIP_VANTRUST_LIQUIDEZ_ACTIVA",
    "Liquidez Alto Monto":  "FIP_VANTRUST_LIQUIDEZ_CONTINUA",
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


# ── Main ─────────────────────────────────────────────────────────────────────

def run(comentario_clp: str, comentario_usd: str):
    fd      = _fecha_ref()
    mes_str = fd.strftime("%Y-%m")
    periodo = _periodo_es(fd)

    print(f"\n{'='*60}\n FOLLETOS {periodo} — {fd}\n{'='*60}\n")

    BASE_DIR    = Path(__file__).parent.parent
    TMPL_DIR    = BASE_DIR / "inputs" / "templates"
    html_dir    = OUTPUT_DIR / mes_str / "html"
    pdf_dir     = OUTPUT_DIR / mes_str
    html_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    # Extend template_reader map for the 4 extra templates
    from etl.template_reader import TEMPLATE_MAP as TM
    TM.update({
        "FIP VANTRUST LIQUIDEZ ALTO APORTE_x":  "TEMPLATE_FONDO_ALTO_PATRIMONIO.xlsx",
        "FIP VANTRUST LIQUIDEZ ALTO CAPITAL_x": "TEMPLATE_FONDO_FACTURA_DOLAR.xlsx",
        "FIP VANTRUST LIQUIDEZ CORTO PLAZO_x":  "TEMPLATE_FONDO_LIQUIDEZ_DISPONIBLE.xlsx",
        "FIP VANTRUST LIQUIDEZ LOCAL_x":        "TEMPLATE_FONDO_LIQUIDEZ_MONETARIO.xlsx",
    })

    # ── 1. Actualizar templates Excel ──────────────────────────────────────
    print("[1/5] Actualizando templates Excel...")
    fecha_dt = datetime(fd.year, fd.month, fd.day)
    resultados = actualizar_todos(fecha_dt, FONDOS_CON_FOLLETO)
    ok = sum(1 for v in resultados.values() if v)
    print(f"  {ok}/{len(FONDOS_CON_FOLLETO)} templates actualizados\n")

    # ── 2 + 3. Generar HTMLs ────────────────────────────────────────────────
    print(f"[2/5] Generando {len(DISPLAY_TO_ODS)} HTMLs...\n")
    html_paths, errores = [], []

    for display_name, (ods_name, _template_file) in DISPLAY_TO_ODS.items():
        try:
            is_usd     = any(x in display_name.upper() for x in ("DÓLAR","DOLAR","USD","DOLLAR"))
            comentario = comentario_usd if is_usd else comentario_clp
            moneda     = "USD" if is_usd else "CLP"

            td = leer_datos_template(ods_name)

            info = get_info_fondo(
                ods_name.replace("_x",""),
                moneda,
                pd.Timestamp(fd)
            )

            comp_cartera = get_cartera_composicion(ods_name.replace("_x",""))

            html_content = generar_html_folleto(
                display_name=display_name,
                periodo=periodo,
                comentario=comentario,
                datos_template=td,
                comp_cartera=comp_cartera,
                info_fondo=info,
            )

            safe      = display_name.replace(" ","_").replace("ó","o").replace("á","a").replace("é","e").replace("ú","u")
            html_path = html_dir / f"FIP_{safe}.html"
            html_path.write_text(html_content, encoding="utf-8")
            html_paths.append(html_path)
            print(f"  ✓ {display_name}")

        except Exception as e:
            print(f"  ✗ {display_name}: {e}")
            errores.append(display_name)

    # ── 4. HTML → PDF ────────────────────────────────────────────────────────
    print(f"\n[3/5] Convirtiendo {len(html_paths)} HTMLs a PDF (Playwright)...\n")
    pdf_paths = html_a_pdf_batch(html_paths, pdf_dir)

    # Rename PDFs to match expected names
    for display_name, html_path in zip(DISPLAY_TO_ODS.keys(), html_paths):
        expected_stem = PDF_NAME_MAP.get(display_name)
        if not expected_stem:
            continue
        generated_pdf = pdf_dir / (html_path.stem + ".pdf")
        target_pdf    = pdf_dir / f"{expected_stem}.pdf"
        if generated_pdf.exists() and generated_pdf != target_pdf:
            generated_pdf.rename(target_pdf)

    final_pdfs = sorted(pdf_dir.glob("FIP_VANTRUST_*.pdf"))

    # ── 5. ZIP ────────────────────────────────────────────────────────────────
    print(f"\n[4/5] Empaquetando {len(final_pdfs)} PDFs en ZIP...")
    zip_mes    = pdf_dir    / f"folletos_{mes_str}.zip"
    zip_latest = OUTPUT_DIR / "latest.zip"

    for zp in (zip_mes, zip_latest):
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in final_pdfs:
                zf.write(pdf, arcname=pdf.name)
    print(f"  ZIP {mes_str}: {zip_mes.stat().st_size//1024} KB")
    print(f"  ZIP latest: {zip_latest.stat().st_size//1024} KB")

    # ── 6. Push a GitHub ──────────────────────────────────────────────────────
    if GITHUB_TOKEN:
        print("\n[5/5] Subiendo a GitHub...")
        _gh_put(zip_latest, "folletos/latest.zip",
                f"📦 latest.zip actualizado {mes_str}")
        _gh_put(zip_mes, f"folletos/{mes_str}/folletos_{mes_str}.zip",
                f"📦 folletos_{mes_str}.zip")
        for pdf in final_pdfs:
            _gh_put(pdf, f"folletos/{mes_str}/{pdf.name}",
                    f"📄 {pdf.stem} {mes_str}")
        # Templates actualizados
        for nombre_fondo, archivo in TEMPLATE_MAP.items():
            ruta = TMPL_DIR / archivo
            if ruta.exists() and resultados.get(nombre_fondo):
                _gh_put(ruta, f"inputs/templates/{archivo}",
                        f"📊 Template {mes_str}")

    print(f"\n{'='*60}")
    print(f" {periodo}: {len(final_pdfs)} PDFs  •  {len(errores)} errores")
    if errores:
        print(f" ✗ Con errores: {errores}")
    print(f"{'='*60}\n")
    return len(errores) == 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        # Try reading from comentarios.json
        com_path = Path(__file__).parent.parent / "inputs" / "comentarios.json"
        if com_path.exists():
            data = json.loads(com_path.read_text())
            ok = run(data.get("clp",""), data.get("usd",""))
        else:
            print("Uso: python main.py \"Comentario CLP\" \"Comentario USD\"")
            sys.exit(1)
    else:
        ok = run(sys.argv[1], sys.argv[2])
    sys.exit(0 if ok else 1)

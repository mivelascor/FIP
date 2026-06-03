"""
main.py — Genera folletos mensuales Vantrust.

LÓGICA DE FECHA:
  Si TARGET_MONTH está seteado (desde admin.html o workflow_dispatch), úsalo.
  Si no, detecta automáticamente el ÚLTIMO MES COMPLETO en ODS:
    - Consulta el último día disponible por mes en ODS
    - Si el mes más reciente tiene datos solo hasta día < 28 → es incompleto → usa el anterior
    - Así nunca genera folletos con datos parciales
"""
import sys, os, json, zipfile, requests
from datetime import date
from pathlib import Path
from dateutil.relativedelta import relativedelta

sys.path.insert(0, str(Path(__file__).parent))

from config      import MESES_ES, get_info_fondo
from etl.template_reader import leer_datos_template, _get_icp_series
from etl.excel_reader    import get_cartera_composicion
from generador.html_builder import generar_html_folleto
from generador.html_to_pdf  import html_a_pdf_batch

API_SQL = "https://claudeods.vantrustcapital.cl/query"

DISPLAY_MAP = {
    "Alto Aporte":           "FIP VANTRUST LIQUIDEZ ALTO APORTE",
    "Alto Capital":          "FIP VANTRUST LIQUIDEZ ALTO CAPITAL",
    "Factura Dólar":         "FIP VANTRUST LIQUIDEZ ALTO MONTO",
    "Liquidez Activa":       "FIP VANTRUST LIQUIDEZ ACTIVA",
    "Liquidez Caja":         "FIP VANTRUST LIQUIDEZ CAJA",
    "Liquidez Continua":     "FIP VANTRUST LIQUIDEZ CONTINUA",
    "Liquidez Corriente":    "FIP VANTRUST LIQUIDEZ CORRIENTE",
    "Liquidez Corto Plazo":  "FIP VANTRUST LIQUIDEZ CORTO PLAZO",
    "Liquidez Disponible I": "FIP VANTRUST LIQUIDEZ DISPONIBLE I",
    "Liquidez Dólar":        "FIP VANTRUST LIQUIDEZ DOLAR",
    "Liquidez Dólar Caja":   "FIP VANTRUST LIQUIDEZ DOLAR CAJA",
    "Liquidez Efectivo":     "FIP VANTRUST LIQUIDEZ EFECTIVO",
    "Liquidez Flexible":     "FIP VANTRUST LIQUIDEZ FLEXIBLE",
    "Liquidez Uno":          "FIP VANTRUST LIQUIDEZ I",
    "Liquidez Local":        "FIP VANTRUST LIQUIDEZ LOCAL",
    "Liquidez Monetario I":  "FIP VANTRUST LIQUIDEZ MONETARIO I",
    "Liquidez Permanente":   "FIP VANTRUST LIQUIDEZ PERMANENTE",
    "Liquidez Plus":         "FIP VANTRUST LIQUIDEZ PLUS",
    "Liquidez Presente":     "FIP VANTRUST LIQUIDEZ PRESENTE",
    "Liquidez Recurrente":   "FIP VANTRUST LIQUIDEZ RECURRENTE",
    "Liquidez Rendimiento":  "FIP VANTRUST LIQUIDEZ RENDIMIENTO",
    "Liquidez Reserva Dólar":"FIP VANTRUST LIQUIDEZ RESERVA DÓLAR",
    "Liquidez Sencillo":     "FIP VANTRUST LIQUIDEZ SENCILLO",
    "Liquidez Temporal":     "FIP VANTRUST LIQUIDEZ TEMPORAL",
}

PDF_NAME_MAP = {
    "Alto Aporte":           "FIP_VANTRUST_LIQUIDEZ_ALTO_APORTE",
    "Alto Capital":          "FIP_VANTRUST_LIQUIDEZ_ALTO_CAPITAL",
    "Factura Dólar":         "FIP_VANTRUST_LIQUIDEZ_ALTO_MONTO",
    "Liquidez Activa":       "FIP_VANTRUST_LIQUIDEZ_ACTIVA",
    "Liquidez Caja":         "FIP_VANTRUST_LIQUIDEZ_CAJA",
    "Liquidez Continua":     "FIP_VANTRUST_LIQUIDEZ_CONTINUA",
    "Liquidez Corriente":    "FIP_VANTRUST_LIQUIDEZ_CORRIENTE",
    "Liquidez Corto Plazo":  "FIP_VANTRUST_LIQUIDEZ_CORTO_PLAZO",
    "Liquidez Disponible I": "FIP_VANTRUST_LIQUIDEZ_DISPONIBLE_I",
    "Liquidez Dólar":        "FIP_VANTRUST_LIQUIDEZ_DOLAR",
    "Liquidez Dólar Caja":   "FIP_VANTRUST_LIQUIDEZ_DOLAR_CAJA",
    "Liquidez Efectivo":     "FIP_VANTRUST_LIQUIDEZ_EFECTIVO",
    "Liquidez Flexible":     "FIP_VANTRUST_LIQUIDEZ_FLEXIBLE",
    "Liquidez Uno":          "FIP_VANTRUST_LIQUIDEZ_I",
    "Liquidez Local":        "FIP_VANTRUST_LIQUIDEZ_LOCAL",
    "Liquidez Monetario I":  "FIP_VANTRUST_LIQUIDEZ_MONETARIO_I",
    "Liquidez Permanente":   "FIP_VANTRUST_LIQUIDEZ_PERMANENTE",
    "Liquidez Plus":         "FIP_VANTRUST_LIQUIDEZ_PLUS",
    "Liquidez Presente":     "FIP_VANTRUST_LIQUIDEZ_PRESENTE",
    "Liquidez Recurrente":   "FIP_VANTRUST_LIQUIDEZ_RECURRENTE",
    "Liquidez Rendimiento":  "FIP_VANTRUST_LIQUIDEZ_RENDIMIENTO",
    "Liquidez Reserva Dólar":"FIP_VANTRUST_LIQUIDEZ_RESERVA_DOLAR",
    "Liquidez Sencillo":     "FIP_VANTRUST_LIQUIDEZ_SENCILLO",
    "Liquidez Temporal":     "FIP_VANTRUST_LIQUIDEZ_TEMPORAL",
}


def _detect_last_complete_month() -> tuple[int, int]:
    """
    Consulta ODS y devuelve (year, month) del último mes CON datos completos.
    Un mes es completo si su último día disponible es >= 28.
    """
    sql = """SELECT TOP 4
        YEAR(FECHA_CIERRE) yr, MONTH(FECHA_CIERRE) mo, MAX(DAY(FECHA_CIERRE)) last_day
    FROM ODS.VALORES_CUOTA_GPI
    WHERE VALOR_CUOTA > 0
      AND RTRIM(LTRIM(EMPRESA)) = 'FIP VANTRUST LIQUIDEZ ALTO APORTE'
    GROUP BY YEAR(FECHA_CIERRE), MONTH(FECHA_CIERRE)
    ORDER BY yr DESC, mo DESC"""
    try:
        r = requests.post(API_SQL, json={"Sql": sql},
                          headers={"Content-Type": "application/json"}, timeout=30)
        rows = r.json().get("rows", [])
        for row in rows:
            yr, mo, last_day = int(row["yr"]), int(row["mo"]), int(row["last_day"])
            if last_day >= 28:   # complete month
                print(f"  → Último mes completo detectado: {yr}-{mo:02d} (último día: {last_day})")
                return yr, mo
    except Exception as e:
        print(f"  [WARN] No se pudo detectar mes completo: {e}")

    # Fallback: mes anterior al actual
    hoy  = date.today()
    prev = date(hoy.year, hoy.month, 1) - relativedelta(months=1)
    print(f"  → Fallback: {prev.year}-{prev.month:02d}")
    return prev.year, prev.month


def _fecha_ref() -> tuple[int, int]:
    tm = os.environ.get("TARGET_MONTH", "").strip()
    if tm:
        y, m = map(int, tm.split("-"))
        print(f"  → TARGET_MONTH forzado: {y}-{m:02d}")
        return y, m
    return _detect_last_complete_month()


def log(msg): print(msg, flush=True)


def _actualizar_referencias(year: int, month: int):
    """Antes de actualizar templates: asegura que ICP (BCCh), Santander (CMF) y
    Banchile (CMF) del mes objetivo esten en los JSON. Asi template_updater puede
    construir la fila nueva sin ingreso manual. Degrada con gracia (no rompe si falla)."""
    inp = Path(__file__).parent.parent / "inputs"
    key = f"{year}-{month:02d}"

    def _load(name):
        p = inp / name
        try: return p, json.loads(p.read_text(encoding="utf-8"))
        except Exception: return p, {}

    # ── ICP via BCCh ──────────────────────────────────────────────────────────
    try:
        p, icp = _load("icp_clicp.json")
        if key not in icp:
            ks = sorted(icp.keys())
            if ks:
                from etl.icp_bcch import icp_mes_bcch
                val = icp_mes_bcch(year, month, float(icp[ks[-1]]))
                if val:
                    icp[key] = round(val, 2)
                    p.write_text(json.dumps(icp, indent=2, ensure_ascii=False), encoding="utf-8")
                    log(f"  ✓ ICP {key} = {icp[key]} (BCCh)")
                else:
                    log(f"  [WARN] ICP {key}: BCCh no disponible (revisa secrets BCCH_USER/BCCH_PASS)")
        else:
            log(f"  · ICP {key} ya presente ({icp[key]})")
    except Exception as e:
        log(f"  [WARN] ICP no actualizado: {e}")

    # ── Competencia CLP (Santander) y USD (Banchile) via CMF ───────────────────
    try:
        from etl.cmf_scraper import get_competencia_clp, get_competencia_usd
        for name, getter, label in [("comp_clp.json", get_competencia_clp, "Santander"),
                                     ("comp_usd.json", get_competencia_usd, "Banchile")]:
            p, comp = _load(name)
            if key in comp:
                log(f"  · {label} {key} ya presente ({comp[key]})")
                continue
            df, val_nuevo, fecha_nueva = getter()
            v = None
            if val_nuevo and fecha_nueva and fecha_nueva.startswith(key):
                v = float(val_nuevo)
            elif not df.empty:
                row = df[df["fecha"].astype(str).str.startswith(key)]
                if not row.empty: v = float(row.iloc[-1]["valor_cuota"])
            # guardrail: rango plausible (Santander ~3000-20000; Banchile ~500-5000)
            rng = (500, 5000) if label == "Banchile" else (3000, 20000)
            if v and not (rng[0] <= v <= rng[1]):
                log(f"  [WARN] {label} {key}={v} fuera de rango {rng}; se ignora (revisar parseo CMF)")
                v = None
            if v:
                comp[key] = round(v, 4)
                p.write_text(json.dumps(comp, indent=2, ensure_ascii=False), encoding="utf-8")
                src_lbl = "CMF" if (val_nuevo and fecha_nueva and fecha_nueva.startswith(key)) else "CMF/extrapolado"
                log(f"  ✓ {label} {key} = {comp[key]} ({src_lbl})")
            else:
                log(f"  [WARN] {label} {key}: CMF no devolvio valor")
    except Exception as e:
        log(f"  [WARN] Competencia no actualizada: {e}")


def run(comentario_clp: str, comentario_usd: str):
    year, month = _fecha_ref()
    # Write detected month for workflow commit message
    try:
        with open('/tmp/target_month.txt', 'w') as _f:
            _f.write(f"{year}-{month:02d}")
    except: pass
    mes_str   = f"{year}-{month:02d}"
    periodo   = f"{MESES_ES[month]} {year}"

    log(f"\n{'='*60}\n FOLLETOS {periodo} — {mes_str}\n{'='*60}\n")

    BASE_DIR = Path(__file__).parent.parent

    # ── Step 0: Update templates with current month data ──────────────────────
    # ── Referencias automaticas (ICP BCCh + competencia CMF) ANTES de templates ──
    log("[0/5] Actualizando referencias (ICP/competencia) automaticamente...")
    _actualizar_referencias(year, month)

    log("[0/4] Actualizando templates con datos del mes...")
    try:
        from etl.template_updater import run_update
        update_results = run_update(year, month)
        updated_count = sum(1 for v in update_results.values() if v == 'updated')
        log(f"  ✓ Templates actualizados: {updated_count}/{len(update_results)}\n")
    except Exception as e:
        import traceback; traceback.print_exc()
        log(f"  [WARN] Template update failed: {e} — usando templates existentes\n")
    html_dir = BASE_DIR / "folletos" / mes_str / "html"
    pdf_dir  = BASE_DIR / "folletos" / mes_str
    html_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    # Pre-load ICP once for all funds
    log("[1/4] Pre-cargando ICP...")
    _get_icp_series()
    log("  ✓ ICP cargado\n")

    log(f"[2/4] Generando {len(DISPLAY_MAP)} HTMLs...")
    html_paths, errores = [], []

    for display_name, ods_name in DISPLAY_MAP.items():
        try:
            is_usd     = any(x in display_name.upper() for x in ("DÓLAR","DOLAR","USD","DOLLAR"))
            comentario = comentario_usd if is_usd else comentario_clp

            td   = leer_datos_template(ods_name, year, month)
            info = get_info_fondo(ods_name, "USD" if is_usd else "CLP", None)
            comp = get_cartera_composicion(ods_name)

            html = generar_html_folleto(display_name, periodo, comentario, td, comp, info)

            safe      = display_name.replace(" ","_").replace("ó","o").replace("á","a").replace("é","e").replace("ú","u")
            html_path = html_dir / f"FIP_{safe}.html"
            html_path.write_text(html, encoding="utf-8")
            html_paths.append((display_name, html_path))

            fip_row = next((r for r in td['resumen'] if r['es_fip']), None)
            m_val   = round(fip_row['m']*100, 3) if fip_row and fip_row['m'] else "EMPTY"
            log(f"  ✓ {display_name:28} m={m_val}%")

        except Exception as e:
            import traceback; traceback.print_exc()
            log(f"  ✗ {display_name}: {e}")
            errores.append(display_name)

    log(f"\n[3/4] Convirtiendo {len(html_paths)} HTMLs a PDF...")
    paths_only = [p for _, p in html_paths]
    pdf_paths  = html_a_pdf_batch(paths_only, pdf_dir)

    for (display_name, _), pdf_path in zip(html_paths, pdf_paths):
        expected = PDF_NAME_MAP.get(display_name)
        if expected:
            target = pdf_dir / f"{expected}.pdf"
            if pdf_path and pdf_path.exists() and pdf_path != target:
                pdf_path.rename(target)

    final_pdfs = sorted(pdf_dir.glob("FIP_VANTRUST_*.pdf"))
    log(f"  {len(final_pdfs)} PDFs generados")

    log(f"\n[4/4] Empaquetando ZIP...")
    # zip_mes goes to folletos/ ROOT so MenuFF can find it via API listing
    zip_mes    = BASE_DIR / "folletos" / f"folletos_{mes_str}.zip"
    zip_latest = BASE_DIR / "folletos" / "latest.zip"
    for zp in (zip_mes, zip_latest):
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in final_pdfs:
                zf.write(pdf, arcname=pdf.name)
    log(f"  ✓ {zip_mes.name}: {zip_mes.stat().st_size//1024} KB")

    log(f"\n{'='*60}")
    log(f" {periodo}: {len(final_pdfs)} PDFs  •  {len(errores)} errores")
    if errores: log(f" Errores: {errores}")
    log(f"{'='*60}\n")
    return len(errores) == 0


if __name__ == "__main__":
    com_path = Path(__file__).parent.parent / "inputs" / "comentarios.json"
    if com_path.exists():
        data = json.loads(com_path.read_text(encoding="utf-8"))
        ok   = run(data.get("clp", ""), data.get("usd", ""))
    elif len(sys.argv) >= 3:
        ok = run(sys.argv[1], sys.argv[2])
    else:
        ok = run("", "")
    sys.exit(0 if ok else 1)

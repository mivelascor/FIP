"""
scripts/main.py — Orquestador principal del pipeline de folletos Vantrust.

Proceso mensual automático:
  1. Determina el mes objetivo (último mes cerrado)
  2. Extrae valores cuota de nuestros fondos (SQL)
  3. Extrae ICP del BCCh (mindicador.cl)
  4. Extrae valores cuota de la competencia (CMF scraping)
  5. Lee composición de cartera (cartera.xlsx — subido via admin.html)
  6. Lee comentarios del PM (comentarios.json — subido via admin.html)
  7. Calcula todas las métricas de rentabilidad
  8. Genera un HTML por fondo
  9. Empaqueta los HTMLs en folletos/YYYY-MM/
 10. Commit de vuelta al repo

Inputs desde admin.html (via GitHub API):
  - inputs/cartera.xlsx
  - inputs/comentarios.json  {"clp": "texto...", "usd": "texto..."}

Todos los demás datos son automáticos.
"""
import os, sys, json, shutil, zipfile
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from config import FUND_CONFIG, MESES_ES
from etl.sql_extractor    import get_valores_cuota
from etl.icp_bcch         import get_icp_serie
from etl.cmf_scraper      import get_competencia, calcular_rentabilidades_comp
from etl.cartera_reader   import leer_cartera
from calculos.rentabilidades import (calcular_resumen, calcular_historia, calcular_chart)
from generador.html_builder  import generar_html


def mes_objetivo() -> tuple:
    """Retorna (year, month) del último mes cerrado."""
    hoy = date.today()
    # Usamos el mes anterior al actual (ya que estamos generando el folleto del mes pasado)
    objetivo = hoy - relativedelta(months=1)
    # Pero si ya estamos a más de 5 días del mes actual, usamos el mes actual
    # (por si el workflow se dispara manualmente a fin de mes)
    target_env = os.environ.get("TARGET_MONTH", "")  # formato "2026-04"
    if target_env:
        year, month = map(int, target_env.split("-"))
        return year, month
    return objetivo.year, objetivo.month


def log(msg):
    print(f"[main] {msg}", flush=True)


def main():
    year, month = mes_objetivo()
    mes_str = f"{year}-{month:02d}"
    log(f"Generando folletos para {MESES_ES.get(month, month)} {year}")

    base = Path(__file__).parent.parent  # raíz del repo
    inputs_dir = base / "inputs"
    out_dir    = base / "folletos" / mes_str
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── 1. Comentarios del PM ───────────────────────────────────────────────
    comentarios_path = inputs_dir / "comentarios.json"
    comentarios = {"clp": "", "usd": ""}
    if comentarios_path.exists():
        with open(comentarios_path) as f:
            comentarios = json.load(f)
        log("Comentarios del PM cargados.")
    else:
        log("⚠ No se encontró inputs/comentarios.json — usando comentarios vacíos.")

    # ── 2. Datos de valor cuota (SQL) ───────────────────────────────────────
    log("Extrayendo valores cuota desde SQL...")
    vc_todos = get_valores_cuota()
    log(f"  Fondos con VC: {list(vc_todos.keys())[:3]}...")

    # ── 3. ICP desde BCCh ───────────────────────────────────────────────────
    log("Descargando ICP desde mindicador.cl...")
    icp_niveles = get_icp_serie()
    log(f"  ICP: {len(icp_niveles)} meses disponibles.")

    # ── 4. Composición de cartera ───────────────────────────────────────────
    cartera_path = inputs_dir / "cartera.xlsx"
    cartera = {}
    if cartera_path.exists():
        cartera = leer_cartera(str(cartera_path))
        log(f"Cartera: {len(cartera)} fondos.")
    else:
        log("⚠ No se encontró inputs/cartera.xlsx — composición vacía.")

    # ── 5. Generar folleto por fondo ────────────────────────────────────────
    generados = []
    errores   = []

    for sql_name, cfg in FUND_CONFIG.items():
        try:
            display_name = cfg["display_name"]
            log(f"  Procesando: {display_name}...")

            is_usd = cfg.get("is_usd", False)

            # Valor cuota del fondo
            vc_fip_raw = vc_todos.get(sql_name, {})
            if not vc_fip_raw:
                # Intentar variantes del nombre
                for k in vc_todos:
                    if cfg["display_name"].upper() in k.upper() or sql_name.upper() in k.upper():
                        vc_fip_raw = vc_todos[k]
                        break

            if not vc_fip_raw:
                log(f"  ⚠ Sin datos VC para {sql_name}")

            # Competencia
            vc_comp = {}
            # Reconstruir dict completo desde histórico CMF
            from etl.cmf_scraper import VC_CLP, VC_USD
            vc_comp_raw = VC_USD if is_usd else VC_CLP
            # Asegurar que tenemos el mes objetivo
            get_competencia(year, month, is_usd)
            vc_comp = dict(vc_comp_raw)

            # Composición de cartera para este fondo
            comp_cartera = cartera.get(sql_name, {})
            if not comp_cartera:
                for k in cartera:
                    if display_name.upper() in k.upper() or sql_name.split()[-1] in k:
                        comp_cartera = cartera[k]
                        break

            # Comentario del PM
            comentario = comentarios.get("usd" if is_usd else "clp", "")

            # Cálculos
            summary = calcular_resumen(vc_fip_raw, icp_niveles, vc_comp, year, month)

            year_inicio = cfg["fecha_inicio"][0]
            historia = calcular_historia(
                vc_fip_raw, icp_niveles, vc_comp,
                year_inicio, year, month,
                cfg["nombre_hist"]
            )

            chart_data = calcular_chart(
                vc_fip_raw, icp_niveles, vc_comp,
                year_inicio, year, month
            )

            # Generar HTML
            html = generar_html(
                cfg=cfg,
                year=year, month=month,
                summary=summary,
                historia=historia,
                chart_data=chart_data,
                comp_cartera=comp_cartera,
                comentario=comentario,
                has_icp=True,
            )

            safe = display_name.replace(" ", "_")
            out_path = out_dir / f"FIP_{safe}_{mes_str}.html"
            out_path.write_text(html, encoding="utf-8")
            generados.append(str(out_path))
            log(f"  ✅ {display_name}")

        except Exception as e:
            import traceback
            log(f"  ❌ {cfg['display_name']}: {e}")
            traceback.print_exc()
            errores.append(cfg["display_name"])

    # ── 6. ZIP ──────────────────────────────────────────────────────────────
    zip_path = base / "folletos" / f"folletos_{mes_str}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in generados:
            zf.write(f, Path(f).name)

    log(f"\n✅ {len(generados)} folletos generados en folletos/{mes_str}/")
    log(f"   ZIP: folletos/folletos_{mes_str}.zip")
    if errores:
        log(f"⚠  Errores en: {errores}")
        sys.exit(1)


if __name__ == "__main__":
    main()

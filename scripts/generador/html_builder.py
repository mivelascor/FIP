"""
generador/html_builder.py
Genera el HTML de un folleto: 1 página con todo el contenido + 1 página de composición.
"""
import re
from pathlib import Path
from config import MESES_ES, _INFO as FUND_INFO, _DEFAULTS

DUR_MAP = {
    'Hasta 30 días':      'Menos de 1 mes',
    'Entre 31 y 90 días': '1-3 meses',
    'Entre 91 y 120 días':'3-4 meses',
    'Entre 1 y 2 años':   'Más de 4 meses',
}
DUR_ORDER   = ['Menos de 1 mes','1-3 meses','3-4 meses','Más de 4 meses']
INSTR_ORDER = ['FMX','Inmob','Nota estructurada','Revisar','Garantia','Pagare']
INSTR_DISPLAY = {
    'FMX':              'Financiamiento',
    'Inmob':            'Finan. Inmobiliario',
    'Nota estructurada':'Nota Estructurada',
    'Revisar':          'Efectos de Comercio',
    'Garantia':         'Garantías',
    'Pagare':           'Pagaré',
}

def _f(v, d=2):
    if v is None: return "—"
    try: return f"{float(v)*100:.{d}f}".replace('.', ',') + "%"
    except: return "—"

def _title(name):
    words = name.split()
    mid = max(1, len(words)//2)
    return "<br>".join([" ".join(words[:mid]), " ".join(words[mid:])]) if len(words)>=2 else name

def _parse_year(lbl):
    s = str(lbl)
    if re.match(r'\d{4}-\d{2}', s):
        try: return int(s[:4])
        except: pass
    parts = s.split()
    if len(parts)==2:
        try: return int(parts[1])
        except: pass
    return 0

def _svg(chart_pts, has_icp):
    if not chart_pts: return '<svg viewBox="0 0 500 110" width="100%" height="110"></svg>'
    all_vals = [float(v) for p in chart_pts for k in ('icp','fund','comp') if p.get(k) is not None for v in [p[k]]]
    if not all_vals: return '<svg viewBox="0 0 500 110" width="100%" height="110"></svg>'
    ymi, yma = min(all_vals), max(all_vals)
    yr = max(yma - ymi, 1)
    xl, xr, yt, yb = 32, 492, 4, 94
    ch, cw, n = yb-yt, xr-xl, len(chart_pts)
    def xy(i,v): return round(xl+(i/max(n-1,1))*cw,2), round(yb-((float(v)-ymi)/yr)*ch,2)
    lines=[]
    for i in range(5):
        f=i/4; yv=ymi+f*yr; yp=yb-f*ch
        lines.append(f'  <line x1="{xl}" x2="{xr}" y1="{yp:.1f}" y2="{yp:.1f}" stroke="#e0e0e0" stroke-width="0.5"/>')
        lines.append(f'  <text x="{xl-2}" y="{yp+2:.1f}" text-anchor="end" font-size="5.5" fill="#999" font-family="Barlow,sans-serif">{round(yv)}</text>')
    if has_icp:
        pts=[(i,p['icp']) for i,p in enumerate(chart_pts) if p.get('icp') is not None]
        if pts:
            s=" ".join(f"{xy(i,v)[0]},{xy(i,v)[1]}" for i,v in pts)
            lines.append(f'  <polyline points="{s}" fill="none" stroke="#aaa" stroke-width="1.2" stroke-dasharray="3,2"/>')
    for col,key,sw in [('#444','comp','1.7'),('#0a0a0a','fund','1.8')]:
        pts=[(i,p[key]) for i,p in enumerate(chart_pts) if p.get(key) is not None]
        if pts:
            s=" ".join(f"{xy(i,v)[0]},{xy(i,v)[1]}" for i,v in pts)
            lines.append(f'  <polyline points="{s}" fill="none" stroke="{col}" stroke-width="{sw}"/>')
    seen=set()
    for i,p in enumerate(chart_pts):
        yr2=_parse_year(p['date'])
        if yr2%2==0 and yr2 not in seen and yr2>2000:
            x,_=xy(i,ymi)
            lines.append(f'  <text x="{x}" y="106.0" text-anchor="middle" font-size="5.5" fill="#999" font-family="Barlow,sans-serif">{yr2}</text>')
            seen.add(yr2)
    lines.append(f'  <line x1="{xl}" x2="{xr}" y1="{yb}" y2="{yb}" stroke="#ccc" stroke-width="0.5"/>')
    return f'<svg viewBox="0 0 500 110" width="100%" height="110" xmlns="http://www.w3.org/2000/svg">\n'+'\n'.join(lines)+'\n</svg>'

def _cbar(data, order, labels=None):
    html=""
    for k in order:
        v=data.get(k,0.0)
        lbl=labels.get(k,k) if labels else k
        ps=f"{v*100:.2f}".replace('.',',')+"%"
        bar=f'\n        <div class="bar-track"><div class="bar-fill" style="width:{v*100:.2f}%"></div></div>' if v>0.001 else ""
        html+=f'      <div class="comp-item">\n        <div class="comp-row"><span class="comp-label">{lbl}</span><span class="comp-val">{ps}</span></div>{bar}\n      </div>\n'
    return html

CSS = """@import url('https://fonts.googleapis.com/css2?family=EB+Garamond:ital,wght@0,400;0,500;0,600;1,400&family=Barlow:wght@300;400;500;600;700&display=swap');
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Barlow',Arial,sans-serif;background:#fff;color:#0a0a0a;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
.page{display:table;width:210mm;min-height:297mm;height:297mm;table-layout:fixed;page-break-after:always;background:linear-gradient(to right,#0a0a0a 68mm,#ffffff 68mm);}
.col-left{display:table-cell;width:68mm;vertical-align:top;padding:0;}
.col-right{display:table-cell;vertical-align:top;background:#fff;padding:0;}
.logo-block{padding:32pt 14pt 0pt 14pt;}
.logo-nombre{font-family:'Barlow',Arial,sans-serif;font-size:20pt;font-weight:700;line-height:1.1;color:#ffffff;letter-spacing:-0.3pt;}
.logo-fondo-row{display:table;width:100%;margin-top:8pt;padding-bottom:8pt;border-bottom:0.75pt solid #ffffff;}
.logo-fondo-spacer{display:table-cell;}
.logo-sub{display:table-cell;font-size:5.5pt;letter-spacing:3pt;color:#ffffff;font-weight:400;text-transform:uppercase;text-align:right;white-space:nowrap;vertical-align:bottom;}
.left-section{padding:7pt 14pt 6pt;border-bottom:1pt solid rgba(255,255,255,0.08);}
.left-section-title{font-size:7pt;font-weight:700;color:#ffffff;margin-bottom:7pt;}
.info-table{width:100%;border-collapse:collapse;}
.info-table tr td{padding:2pt 0;border-bottom:1pt solid rgba(255,255,255,0.05);vertical-align:top;font-size:7pt;line-height:1.35;}
.info-table tr:last-child td{border-bottom:none;}
.info-table td.label{color:rgba(255,255,255,0.55);white-space:nowrap;padding-right:4pt;width:44%;}
.info-table td.value{color:#ffffff;font-weight:400;}
.left-text{font-size:7pt;color:rgba(255,255,255,0.50);line-height:1.5;}
.comp-item{padding:2.5pt 0;border-bottom:1pt solid rgba(255,255,255,0.05);}
.comp-item:last-child{border-bottom:none;}
.comp-row{display:table;width:100%;}
.comp-label{display:table-cell;color:rgba(255,255,255,0.45);font-size:7pt;}
.comp-val{display:table-cell;color:#fff;font-size:7pt;font-weight:700;text-align:right;white-space:nowrap;}
.bar-track{background:rgba(255,255,255,0.1);border-radius:1pt;height:2pt;margin-top:1.5pt;}
.bar-fill{height:2pt;border-radius:1pt;background:rgba(255,255,255,0.55);}
.right-inner{padding:18pt 18pt 12pt 16pt;}
.section-title{font-family:'EB Garamond',serif;font-size:13pt;font-weight:500;border-bottom:1.5pt solid #0a0a0a;padding-bottom:3pt;margin-bottom:6pt;letter-spacing:-0.2pt;color:#0a0a0a;}
.comentario-box{border-left:2.5pt solid #0a0a0a;padding:5pt 8pt;font-size:7.5pt;color:#1c1c1c;line-height:1.55;font-style:italic;background:#f7f7f5;margin-bottom:10pt;}
.chart-section{margin-bottom:10pt;}
.chart-container{border:0.75pt solid #d8d8d8;background:#f7f7f5;padding:5pt 6pt 3pt;}
.chart-label-row{display:table;width:100%;margin-bottom:2pt;}
.chart-label{display:table-cell;font-size:6pt;letter-spacing:0.8pt;text-transform:uppercase;color:#aaa;}
.chart-label-right{display:table-cell;text-align:right;font-size:5.5pt;color:#ccc;}
.chart-legend-row{margin-top:3pt;}
.legend-item{display:inline-block;margin-right:10pt;font-size:6.5pt;color:#666;vertical-align:middle;}
.legend-line-solid{display:inline-block;width:12pt;height:1.5pt;vertical-align:middle;margin-right:2pt;border-radius:1pt;}
.legend-line-dashed{display:inline-block;width:12pt;height:0;border-top:1.5pt dashed #aaa;vertical-align:middle;margin-right:2pt;}
/* Summary table — compact */
.tbl-resumen{width:100%;border-collapse:collapse;margin-bottom:8pt;font-size:7pt;}
.tbl-resumen thead th{background:#0a0a0a;color:#fff;padding:3.5pt 3pt;text-align:center;font-weight:500;font-size:6.5pt;white-space:nowrap;}
.tbl-resumen thead th:first-child{text-align:left;padding-left:5pt;}
.tbl-resumen tbody td{padding:3.5pt 3pt;text-align:center;border-bottom:0.75pt solid #d8d8d8;font-size:7.5pt;color:#1c1c1c;}
.tbl-resumen tbody td:first-child{text-align:left;padding-left:5pt;font-weight:600;}
.tbl-resumen tbody tr:nth-child(odd) td{background:#f7f7f5;}
.tbl-resumen tbody tr.fip td{background:#0a0a0a;color:#fff;font-weight:700;border-bottom:none;}
/* History table — compact */
.tbl-hist{width:100%;border-collapse:collapse;font-size:5.5pt;}
.tbl-hist thead th{background:#0a0a0a;color:#fff;padding:2.5pt 1.5pt;text-align:center;font-weight:500;font-size:5.5pt;white-space:nowrap;}
.tbl-hist thead th:first-child{text-align:left;padding-left:4pt;width:18pt;}
.tbl-hist thead th:nth-child(2){text-align:left;width:46pt;}
.tbl-hist tbody tr.tr-icp td{background:#fff;}
.tbl-hist tbody tr.tr-comp td{background:#f7f7f5;}
.tbl-hist tbody tr.tr-fip td{background:#ebebeb;font-weight:700;}
.tbl-hist tbody tr.tr-sp td{height:2pt;padding:0;border:none;background:transparent!important;}
.tbl-hist td{padding:1.8pt 1.5pt;text-align:center;color:#333;border-bottom:0.5pt solid rgba(0,0,0,0.04);}
.tbl-hist td:first-child{text-align:left;padding-left:4pt;font-weight:700;color:#0a0a0a;}
.tbl-hist td:nth-child(2){text-align:left;font-size:5pt;color:#666;}
.tbl-hist td:last-child{font-weight:700;}
/* Glossary */
.glos-item{margin-bottom:7pt;}
.glos-term{font-weight:700;font-size:8pt;margin-bottom:2pt;color:#0a0a0a;}
.glos-def{font-size:7.5pt;color:#666;line-height:1.5;}
.disclaimer-box{border:0.75pt solid #d8d8d8;padding:7pt 9pt;font-size:7.5pt;color:#666;line-height:1.5;background:#f7f7f5;}
@media print{body{background:#fff;}.page{page-break-after:always;}}
@page{size:A4;margin:0;}"""

GLOSARIO = """
      <div class="glos-item"><div class="glos-term">Riesgo de Mercado</div><div class="glos-def">Variación adversa en el precio o tasa de mercado en los instrumentos en que invierte el Fondo.</div></div>
      <div class="glos-item"><div class="glos-term">Riesgo de Crédito</div><div class="glos-def">Posible pérdida por incumplimiento de las obligaciones de los emisores de los instrumentos en que invierte el Fondo.</div></div>
      <div class="glos-item"><div class="glos-term">Riesgo de Liquidez</div><div class="glos-def">Riesgo asociado a la capacidad del Fondo para cumplir con sus obligaciones de rescate o vencimiento.</div></div>
      <div class="glos-item"><div class="glos-term">Riesgo de Tasa de interés</div><div class="glos-def">Exposición a pérdidas por cambios adversos en las tasas de interés de mercado.</div></div>
      <div class="glos-item"><div class="glos-term">Riesgo de Moneda</div><div class="glos-def">Exposición a pérdidas por cambios adversos en el valor de las monedas extranjeras del balance del Fondo.</div></div>
      <div class="glos-item"><div class="glos-term">Gastos del fondo</div><div class="glos-def">Gastos directos e indirectos necesarios para el correcto funcionamiento del fondo, detallados en el Reglamento Interno.</div></div>"""


def generar_html_folleto(display_name, periodo, comentario, datos_template,
                          comp_cartera, info_fondo) -> str:
    td   = datos_template
    info = info_fondo

    title = _title(display_name)

    # ── Summary rows ─────────────────────────────────────────────────────────
    summary_html = ""
    for row in td['resumen']:
        css = ' class="fip"' if row['es_fip'] else ''
        summary_html += f'          <tr{css}><td>{row["nombre"]}</td><td>{_f(row["m"])}</td><td>{_f(row["t"])}</td><td>{_f(row["s"])}</td><td>{_f(row["a"])}</td><td>{_f(row["ac"])}</td></tr>\n'

    # ── Historical table ─────────────────────────────────────────────────────
    by_year = {}
    for yr_data in td['historico']:
        by_year.setdefault(yr_data['año'], []).extend(yr_data['filas'])
    hist_html = ""
    for yr in sorted(by_year):
        rows = by_year[yr]; n = len(rows); first = True
        for fila in rows:
            fn = fila['nombre'].upper()
            css = 'tr-icp' if 'ICP' in fn else ('tr-comp' if 'COMP' in fn else 'tr-fip')
            mths = "".join(f"<td>{_f(v)}</td>" if v is not None else "<td>—</td>" for v in fila['meses'])
            tot  = f"<td>{_f(fila.get('total'))}</td>" if fila.get('total') is not None else "<td>—</td>"
            yr_cell = f'<td rowspan="{n}">{yr}</td>' if first else ""; first = False
            hist_html += f'          <tr class="{css}">{yr_cell}<td>{fila["nombre"]}</td>{mths}{tot}</tr>\n'
        hist_html += '          <tr class="tr-sp"><td colspan="15"></td></tr>\n'

    # ── Chart ────────────────────────────────────────────────────────────────
    chart_pts = [{'date': td['grafico']['labels'][i],
                  'icp':  td['grafico']['icp'][i]  if i < len(td['grafico']['icp'])  else None,
                  'fund': td['grafico']['fip'][i]  if i < len(td['grafico']['fip'])  else None,
                  'comp': td['grafico']['comp'][i] if i < len(td['grafico']['comp']) else None}
                 for i in range(len(td['grafico']['labels']))]
    has_icp = any(r['es_icp'] for r in td['resumen'])
    svg = _svg(chart_pts, has_icp)
    leg_icp  = '<span class="legend-item"><span class="legend-line-dashed"></span>ICP Norm.</span>' if has_icp else ''
    leg_comp = '<span class="legend-item"><span class="legend-line-solid" style="background:#444;"></span>Competencia Relevante</span>'
    leg_fund = f'<span class="legend-item"><span class="legend-line-solid" style="background:#0a0a0a;"></span>{td["nombre_fip"]}</span>'
    acum_lbl = td['acum_label'].replace('\n', ' ')

    # ── Cartera composition ───────────────────────────────────────────────────
    dur, curr, instr = {}, {}, {}
    if comp_cartera:
        def pp(s):
            try: return float(str(s).replace('%','').replace(',','.').strip())/100
            except: return 0.0

        # Moneda
        for l, p in comp_cartera.get('moneda', []):
            key = 'Pesos' if l == 'CLP' else ('UF' if l == 'CLF' else str(l))
            curr[key] = pp(p)

        # Duración
        for l, p in comp_cartera.get('duracion', []):
            key = DUR_MAP.get(str(l), str(l))
            dur[key] = pp(p)

        # Instrumento
        for l, p in comp_cartera.get('instrumento', []):
            instr[str(l)] = pp(p)

    mon_html   = _cbar(curr,  [k for k in ['Pesos','UF','USD'] if k in curr or curr.get(k,0)>0])
    dur_html   = _cbar(dur,   DUR_ORDER)
    instr_html = _cbar(instr, [k for k in INSTR_ORDER if k in instr],
                       labels=INSTR_DISPLAY)

    # ── Static info ───────────────────────────────────────────────────────────
    ods_clean = display_name  # fallback
    fi        = FUND_INFO.get(display_name, {})
    rut       = str(info.get('rut') or fi.get('rut', '')).strip()
    moneda    = str(info.get('moneda') or fi.get('moneda','CLP')).strip()
    tipo      = str(fi.get('tipo', _DEFAULTS.get('tipo','Fondo de Inversión Privado'))).strip()
    fecha_s   = str(info.get('fecha_inicio','') or fi.get('fecha_inicio','')).strip()
    remun     = str(fi.get('remuneracion','—')).strip()
    custodio  = str(_DEFAULTS.get('custodio','Vantrust Capital C. de Bolsa')).strip()
    forma     = ("La moneda en que el inversionista entra al fondo es aportando dólares estadounidenses, y al rescate de las cuotas, el fondo le entrega dólares estadounidenses." if moneda=="USD"
                 else "La moneda en que el inversionista entra al fondo es aportando pesos chilenos, y al rescate de las cuotas, el fondo le entrega pesos chilenos.")

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>FIP {display_name} – {periodo}</title>
<style>{CSS}</style></head>
<body>
<!-- ═══════════ PÁGINA 1: todo el contenido ═══════════ -->
<div class="page">
  <div class="col-left">
    <div class="logo-block"><div class="logo-nombre">{title}</div><div class="logo-fondo-row"><span class="logo-fondo-spacer"></span><span class="logo-sub">F O N D O</span></div></div>
    <div class="left-section"><div class="left-section-title">Información General</div>
      <table class="info-table">
        <tr><td class="label">Administradora</td><td class="value">Vantrust Gestión Patrimonial S.A.</td></tr>
        <tr><td class="label">RUT Fondo</td><td class="value">{rut}</td></tr>
        <tr><td class="label">Moneda</td><td class="value">{moneda}</td></tr>
        <tr><td class="label">Tipo de Fondo</td><td class="value">{tipo}</td></tr>
        <tr><td class="label">Fecha Inicio</td><td class="value">{fecha_s}</td></tr>
        <tr><td class="label">Benchmark</td><td class="value">Índice Cámara Promedio (ICP)</td></tr>
        <tr><td class="label">Fondo Rescatable</td><td class="value">Sí</td></tr>
        <tr><td class="label">Plazo Rescate</td><td class="value">A más tardar 15 días corridos</td></tr>
        <tr><td class="label">Riesgos</td><td class="value">Mercado – Crédito – Liquidez – Tasa interés – Derivados</td></tr>
        <tr><td class="label">Remuneración</td><td class="value">{remun}</td></tr>
        <tr><td class="label">Custodio</td><td class="value">{custodio}</td></tr>
      </table></div>
    <div class="left-section"><div class="left-section-title">Objetivo</div><p class="left-text">Invertir los recursos del fondo en instrumentos de deuda de corto y mediano plazo, en una cartera diversificada, obteniendo una rentabilidad igual o superior al ICP.</p></div>
    <div class="left-section"><div class="left-section-title">Rentabilidad</div><p class="left-text">La rentabilidad esperada del Fondo Vantrust {display_name} es la tasa de política monetaria promedio del Banco Central de Chile.</p></div>
    <div class="left-section"><div class="left-section-title">Inversionistas</div><p class="left-text">Dirigida a empresas y personas que buscan invertir sus excedentes de caja con una rentabilidad de corto plazo y baja tolerancia al riesgo.</p></div>
  </div>
  <div class="col-right"><div class="right-inner">
    <div class="section-title">Comentario Portafolio Manager</div>
    <div class="comentario-box">{comentario or '[Comentario del Portafolio Manager]'}</div>
    <div class="chart-section">
      <div class="section-title">Evolución Rentabilidad</div>
      <div class="chart-container">
        <div class="chart-label-row"><span class="chart-label">{td['nombre_fip']}</span><span class="chart-label-right">Índice base 100</span></div>
        {svg}
        <div class="chart-legend-row">{leg_icp}{leg_comp}{leg_fund}</div>
      </div>
    </div>
    <table class="tbl-resumen">
      <thead><tr><th>Rentabilidad</th><th>Mensual</th><th>Trimestral</th><th>Semestral</th><th>Anual</th><th>{acum_lbl}</th></tr></thead>
      <tbody>
{summary_html}      </tbody>
    </table>
    <table class="tbl-hist">
      <thead><tr><th>Año</th><th>Fondo</th><th>Ene</th><th>Feb</th><th>Mar</th><th>Abr</th><th>May</th><th>Jun</th><th>Jul</th><th>Ago</th><th>Sep</th><th>Oct</th><th>Nov</th><th>Dic</th><th>Total</th></tr></thead>
      <tbody>
{hist_html}      </tbody>
    </table>
  </div></div>
</div>
<!-- ═══════════ PÁGINA 2: composición + glosario ═══════════ -->
<div class="page">
  <div class="col-left">
    <div class="logo-block"><div class="logo-nombre">{title}</div><div class="logo-fondo-row"><span class="logo-fondo-spacer"></span><span class="logo-sub">F O N D O</span></div></div>
    <div class="left-section"><div class="left-section-title">Composición por Moneda</div>{mon_html}</div>
    <div class="left-section"><div class="left-section-title">Composición por Instrumento</div>{instr_html}</div>
    <div class="left-section"><div class="left-section-title">Composición por Duración</div>{dur_html}</div>
  </div>
  <div class="col-right"><div class="right-inner">
    <div class="section-title">Glosario</div>
{GLOSARIO}
    <div class="glos-item"><div class="glos-term">Forma de Ingreso y Pago del Fondo</div><div class="glos-def">{forma}</div></div>
    <div class="section-title" style="margin-top:10pt;">Disclaimer</div>
    <div class="disclaimer-box">Conforme a la Ley Única de Fondos, las administradoras de fondos de inversión privados están sujetas a las obligaciones de información establecidas por la Comisión para el Mercado Financiero. Tales fondos no están sometidos a fiscalización de la Comisión y no hacemos oferta pública de sus cuotas.</div>
  </div></div>
</div>
</body></html>"""

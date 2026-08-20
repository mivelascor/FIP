# -*- coding: utf-8 -*-
"""
generar_menu.py — Regenera MenuFF.html y las páginas Oxx*.html de cada fondo
a partir de scripts/fondos_registry.py (única fuente de verdad del menú).

Uso:
    python scripts/generar_menu.py

Para agregar un fondo, marcar uno como vencido o cambiar su reglamento,
editar scripts/fondos_registry.py y volver a correr este script.
Columnas del registro:
    id, nombre_menu, titulo_l1, titulo_l2, rut, moneda, estado,
    reglamento (archivo en regl.int/), prorroga (en regl.int/prorrogas/),
    folleto_pdf (nombre base del PDF, o None si todavía no se genera)
"""
import sys, os, html, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fondos_registry import R
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def esc(s): return html.escape(s or '', quote=True)
def url(p): return p.replace(' ','%20').replace('°','%C2%B0')

# ─────────────────────────── PÁGINA DE FONDO ───────────────────────────
PAGE = '''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{titulo_doc}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');
:root{{--bg:#0a0a0a;--surface:#111;--surface2:#161616;--border:#1e1e1e;--border2:#2a2a2a;--text:#e8e8e8;--muted:#555;--muted2:#777;--green:#22c55e;--blue:#60a5fa;--amber:#d19a3a;}}
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:60px;}}
.topbar{{border-bottom:1px solid var(--border);padding:15px 32px;display:flex;align-items:center;gap:14px;position:sticky;top:0;background:var(--bg);z-index:50;}}
.tb-back{{font-size:12px;color:var(--muted);text-decoration:none;transition:color .15s;letter-spacing:.3px;}}
.tb-back:hover{{color:var(--text);}}
.tb-sep{{color:var(--border2);font-size:16px;}}
.tb-name{{font-size:13px;font-weight:500;color:var(--muted2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
main{{max-width:540px;margin:0 auto;padding:44px 24px 0;}}
.badges{{display:flex;gap:8px;align-items:center;margin-bottom:14px;flex-wrap:wrap;}}
.badge{{font-size:9px;font-weight:700;letter-spacing:3px;text-transform:uppercase;background:var(--surface);border:1px solid var(--border2);color:var(--muted2);padding:4px 10px;display:inline-block;}}
.badge.usd{{color:var(--blue);border-color:#1e3a5f;background:#0d1f33;}}
.badge.vig{{color:var(--green);border-color:#14532d;background:#0c1f14;}}
.badge.ven{{color:var(--amber);border-color:#4a3308;background:#231a06;}}
.fund-title{{font-family:'DM Serif Display',serif;font-size:34px;font-weight:400;letter-spacing:-0.8px;line-height:1.1;margin-bottom:6px;}}
.fund-rut{{font-size:12px;color:var(--muted);margin-bottom:36px;}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px;}}
.action-card{{background:var(--surface);border:1px solid var(--border);padding:22px 18px;text-decoration:none;color:var(--text);transition:all .18s;display:flex;flex-direction:column;gap:10px;cursor:pointer;border-radius:1px;}}
.action-card:hover{{background:var(--surface2);border-color:var(--border2);transform:translateY(-1px);}}
.action-card.off{{opacity:.45;cursor:default;}}
.action-card.off:hover{{background:var(--surface);border-color:var(--border);transform:none;}}
.ac-icon{{font-size:22px;}}
.ac-title{{font-size:14px;font-weight:600;line-height:1.3;}}
.ac-meta{{font-size:11px;color:var(--muted);}}
.ac-meta.ok{{color:var(--green);}}
.hist-btn{{background:var(--surface);border:1px solid var(--border);padding:18px;text-align:center;cursor:pointer;transition:all .18s;width:100%;border-radius:1px;display:flex;align-items:center;justify-content:center;gap:10px;}}
.hist-btn:hover{{background:var(--surface2);border-color:var(--border2);}}
.hist-btn-icon{{font-size:20px;}}
.hist-btn-lbl{{font-size:13px;font-weight:500;}}
.hist-panel{{display:none;margin-top:10px;background:var(--surface);border:1px solid var(--border);padding:20px;border-radius:1px;}}
.hp-label{{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--muted);margin-bottom:12px;}}
.hp-row{{display:flex;gap:8px;align-items:center;}}
select{{background:var(--bg);border:1px solid var(--border2);color:var(--text);font-family:'DM Sans',sans-serif;font-size:12px;padding:8px 12px;outline:none;transition:border-color .15s;border-radius:1px;flex:1;cursor:pointer;}}
select:focus{{border-color:var(--muted2);}}
.btn-dl{{background:var(--text);color:var(--bg);border:none;font-family:'DM Sans',sans-serif;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;padding:9px 16px;cursor:pointer;transition:opacity .15s;border-radius:1px;white-space:nowrap;}}
.btn-dl:hover{{opacity:.86;}}
.dl-msg{{font-size:11px;color:var(--muted);margin-top:8px;min-height:16px;}}
.note{{font-size:11px;color:var(--muted);line-height:1.6;margin-top:16px;}}
@keyframes fadeUp{{from{{opacity:0;transform:translateY(6px)}}to{{opacity:1;transform:translateY(0)}}}}
.hist-panel.open{{display:block;animation:fadeUp .2s ease;}}
@media(max-width:480px){{.grid2{{grid-template-columns:1fr;}}}}
.site-footer{{border-top:1px solid var(--border,#1e1e1e);margin-top:60px;padding:20px 40px;font-size:11px;color:var(--muted,#555);line-height:1.6;}}
</style>
</head>
<body>
<div class="topbar">
  <a class="tb-back" href="MenuFF.html">← Fondos</a>
  <span class="tb-sep">|</span>
  <span class="tb-name">{titulo_doc}</span>
</div>
<main>
  <div class="badges">
    <span class="badge{usd_cls}">FIP VANTRUST · {moneda}</span>
    <span class="badge {est_cls}">{est_lbl}</span>
  </div>
  <h1 class="fund-title">{h1}</h1>
  <p class="fund-rut">{rut}</p>

  <div class="grid2">
    {card_regl}
    {card_folleto}
  </div>
  {card_prorroga}

  {bloque_hist}
  {nota}
</main>
<script>
const REPO="mivelascor/FIP",BRANCH="main",PDF={pdf_js};
let abierto=false;

async function cargarMeses(){{
  if(!PDF) return;
  try{{
    const r=await fetch(`https://api.github.com/repos/${{REPO}}/contents/folletos`,
      {{headers:{{"Accept":"application/vnd.github+json"}}}});
    if(!r.ok)return;
    const items=await r.json();
    const meses=items.filter(i=>i.type==="dir"&&/^\\d{{4}}-\\d{{2}}$/.test(i.name)).map(i=>i.name).sort().reverse();
    const sel=document.getElementById("selMes");
    const meta=document.getElementById("folleto-meta");
    if(!meses.length){{if(sel)sel.innerHTML="<option>Sin folletos</option>";if(meta)meta.textContent="No disponible";return;}}
    if(sel)sel.innerHTML=meses.map(m=>`<option value="${{m}}">${{m}}</option>`).join("");
    if(meta)meta.textContent=meses[0];
  }}catch(e){{const meta=document.getElementById("folleto-meta");if(meta)meta.textContent="No disponible";}}
}}

function dlPDF(mes){{
  const url=`https://raw.githubusercontent.com/${{REPO}}/${{BRANCH}}/folletos/${{mes}}/${{PDF}}.pdf`;
  const a=document.createElement("a");a.href=url;a.download=`${{PDF}}_${{mes}}.pdf`;a.click();
}}
function dlUltimo(){{
  const sel=document.getElementById("selMes");
  const mes=sel&&sel.value;
  if(!mes){{alert("Sin folletos disponibles.");return;}}
  dlPDF(mes);
}}
function dlHistorial(){{
  const mes=document.getElementById("selMes").value;
  if(!mes){{alert("Selecciona un mes.");return;}}
  dlPDF(mes);
  const msg=document.getElementById("dlMsg");
  msg.textContent=`Descargando PDF ${{mes}}...`;
  setTimeout(()=>msg.textContent="",2500);
}}
function toggleHist(){{
  abierto=!abierto;
  document.getElementById("histPanel").classList.toggle("open",abierto);
}}
cargarMeses();
</script>
<footer class="site-footer">Conforme a la Ley Única de Fondos, las administradoras de fondos de inversión privados están sujetas a las obligaciones de información establecidas por la Comisión para el Mercado Financiero. Tales fondos no están sometidos a fiscalización de la Comisión y no hacemos oferta pública de sus cuotas.</footer>
</body>
</html>
'''

HIST = '''<div class="hist-btn" onclick="toggleHist()">
    <span class="hist-btn-icon">🗂️</span>
    <span class="hist-btn-lbl">Historial de folletos</span>
  </div>

  <div class="hist-panel" id="histPanel">
    <div class="hp-label">Selecciona un mes</div>
    <div class="hp-row">
      <select id="selMes"><option value="">Cargando...</option></select>
      <button class="btn-dl" onclick="dlHistorial()">Descargar PDF</button>
    </div>
    <div class="dl-msg" id="dlMsg"></div>
  </div>'''

paginas = 0
for (pid,nombre,l1,l2,rut,moneda,estado,regl,prorroga,pdf) in R:
    titulo_doc = f"FIP VANTRUST {l1} {l2}".replace('  ',' ').strip().upper()
    h1 = (f"{l1}<br>{l2}" if l1 else l2)
    if regl:
        card_regl = (f'<a class="action-card" href="regl.int/{url(esc(regl))}" download>\n'
                     f'      <div class="ac-icon">📄</div>\n'
                     f'      <div class="ac-title">Reglamento<br>Interno</div>\n'
                     f'      <div class="ac-meta">Documento vigente</div>\n    </a>')
    else:
        card_regl = ('<div class="action-card off">\n'
                     '      <div class="ac-icon">📄</div>\n'
                     '      <div class="ac-title">Reglamento<br>Interno</div>\n'
                     '      <div class="ac-meta">No disponible en el repositorio</div>\n    </div>')
    if pdf:
        card_folleto = ('<div class="action-card" onclick="dlUltimo()">\n'
                        '      <div class="ac-icon">📑</div>\n'
                        '      <div class="ac-title">Folleto<br>Comercial — PDF</div>\n'
                        '      <div class="ac-meta ok" id="folleto-meta">Cargando...</div>\n    </div>')
    else:
        card_folleto = ('<div class="action-card off">\n'
                        '      <div class="ac-icon">📑</div>\n'
                        '      <div class="ac-title">Folleto<br>Comercial — PDF</div>\n'
                        '      <div class="ac-meta">Aún no se genera</div>\n    </div>')
    if prorroga:
        card_prorroga = (f'<div class="grid2"><a class="action-card" href="regl.int/prorrogas/{url(esc(prorroga))}" download>\n'
                         f'      <div class="ac-icon">🗓️</div>\n'
                         f'      <div class="ac-title">Prórroga de<br>vigencia</div>\n'
                         f'      <div class="ac-meta">Acta de directorio 2026</div>\n    </a></div>')
    else:
        card_prorroga = ''
    nota = ('' if pdf else
            '<p class="note">Este fondo todavía no tiene folleto comercial. Se generará automáticamente '
            'en cuanto aparezca su valor cuota en la planilla mensual.</p>')
    page = PAGE.format(
        titulo_doc=titulo_doc, h1=h1, rut=esc(rut or '—'), moneda=moneda,
        usd_cls=' usd' if moneda=='USD' else '',
        est_cls='vig' if estado=='vigente' else 'ven',
        est_lbl='Vigente' if estado=='vigente' else 'Vencido',
        card_regl=card_regl, card_folleto=card_folleto, card_prorroga=card_prorroga,
        bloque_hist=HIST if pdf else '', nota=nota,
        pdf_js=json.dumps(pdf) if pdf else 'null')
    open(os.path.join(REPO,pid+'.html'),'w',encoding='utf-8').write(page)
    paginas += 1
print('páginas generadas:', paginas)


# ─────────────────────────── MENÚ ───────────────────────────
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def esc(s): return html.escape(s or '', quote=True)

orden = sorted(R, key=lambda x:(x[5]=='USD', x[1]))
cards=[]
for (pid,nombre,l1,l2,rut,moneda,estado,regl,prorroga,pdf) in orden:
    chips=[]
    if estado=='vencido': chips.append('<span class="chip ven">Vencido</span>')
    if not pdf:           chips.append('<span class="chip sin">Sin folleto</span>')
    cards.append(
      f'  <a class="card" href="{pid}.html" data-estado="{estado}" data-folleto="{"si" if pdf else "no"}">'
      f'<div class="card-name">{esc(nombre)}</div>'
      f'<div class="card-foot"><span class="card-tag{" usd" if moneda=="USD" else ""}">{moneda}</span>'
      f'{"".join(chips)}</div></a>')
GRID='\n'.join(cards)
n_tot=len(R); n_vig=sum(1 for x in R if x[6]=='vigente'); n_ven=n_tot-n_vig

MENU = '''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fichas de Fondos — Vantrust</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');
:root{--bg:#0a0a0a;--surface:#111;--s2:#161616;--border:#1e1e1e;--border2:#2a2a2a;--text:#e8e8e8;--muted:#555;--muted2:#777;--blue:#3b82f6;--green:#22c55e;--amber:#d19a3a;}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:80px;}

.topbar{border-bottom:1px solid var(--border);padding:15px 40px;display:flex;align-items:center;gap:14px;position:sticky;top:0;background:var(--bg);z-index:50;}
.tb-back{font-size:12px;color:var(--muted);text-decoration:none;transition:color .15s;letter-spacing:.3px;}
.tb-back:hover{color:var(--text);}
.tb-sep{color:var(--border2);font-size:16px;}
.tb-title{font-family:'DM Serif Display',serif;font-size:17px;font-weight:400;}

.controls{padding:28px 40px 0;display:flex;align-items:center;gap:16px;flex-wrap:wrap;}
.search-box{width:300px;background:var(--surface);border:1px solid var(--border);color:var(--text);font-family:'DM Sans',sans-serif;font-size:13px;padding:10px 14px;outline:none;transition:border-color .15s;border-radius:2px;}
.search-box:focus{border-color:var(--border2);}
.search-box::placeholder{color:var(--muted);}

/* ── Toggle Todos / Vigentes / Vencidos ── */
.seg{display:inline-flex;border:1px solid var(--border2);border-radius:2px;overflow:hidden;background:var(--surface);}
.seg button{background:transparent;border:none;border-right:1px solid var(--border2);color:var(--muted2);
  font-family:'DM Sans',sans-serif;font-size:12px;font-weight:500;letter-spacing:.3px;
  padding:10px 18px;cursor:pointer;transition:all .15s;display:flex;align-items:center;gap:7px;}
.seg button:last-child{border-right:none;}
.seg button:hover{color:var(--text);background:var(--s2);}
.seg button.on{background:var(--text);color:var(--bg);}
.seg button .n{font-size:10px;font-weight:700;opacity:.6;}
.seg button.on .n{opacity:.5;}
.search-count{font-size:12px;color:var(--muted);}

.eyebrow{padding:22px 40px 14px;font-size:10px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:var(--muted);}

.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--border);border:1px solid var(--border);margin:0 40px 44px;}
.card{background:var(--bg);padding:22px;text-decoration:none;color:var(--text);transition:background .15s;display:flex;flex-direction:column;justify-content:space-between;gap:12px;min-height:104px;}
.card:hover{background:var(--surface);}
.card-name{font-size:13px;font-weight:500;line-height:1.4;}
.card-foot{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}
.card-tag{font-size:10px;color:var(--muted);letter-spacing:.3px;font-weight:500;}
.card-tag.usd{color:var(--blue);}
.chip{font-size:9px;font-weight:600;letter-spacing:.5px;padding:2px 6px;border-radius:2px;border:1px solid;}
.chip.ven{color:var(--amber);border-color:#4a3308;background:#231a06;}
.chip.sin{color:var(--muted2);border-color:var(--border2);background:var(--surface);}
.empty{padding:40px;text-align:center;color:var(--muted);font-size:13px;grid-column:1/-1;background:var(--bg);}

/* ── Downloads ── */
.dl-section{margin:0 40px;background:var(--surface);border:1px solid var(--border);padding:32px 32px 28px;}
.dl-title{font-family:'DM Serif Display',serif;font-size:24px;font-weight:400;margin-bottom:6px;}
.dl-sub{font-size:12px;color:var(--muted);margin-bottom:28px;line-height:1.55;}
.dl-primary-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:24px;}
.btn-big{display:flex;align-items:center;gap:14px;background:var(--s2);border:1px solid var(--border2);padding:18px 20px;cursor:pointer;transition:all .18s;border-radius:2px;width:100%;text-align:left;color:var(--text);font-family:'DM Sans',sans-serif;}
.btn-big:hover{border-color:var(--text);background:var(--border);}
.btn-big:disabled{opacity:.35;cursor:not-allowed;}
.btn-big.primary{background:var(--text);border-color:var(--text);color:var(--bg);}
.btn-big.primary:hover{opacity:.88;}
.btn-icon{font-size:26px;flex-shrink:0;}
.btn-lbl{display:flex;flex-direction:column;gap:3px;}
.btn-lbl strong{font-size:13px;font-weight:600;}
.btn-lbl span{font-size:11px;opacity:.65;}

.dl-divider{border:none;border-top:1px solid var(--border);margin:8px 0 24px;}
.hist-label{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--muted);margin-bottom:14px;}
.hist-row{display:flex;gap:10px;align-items:center;}
select{background:var(--s2);border:1px solid var(--border2);color:var(--text);font-family:'DM Sans',sans-serif;font-size:13px;padding:9px 13px;outline:none;transition:border-color .15s;border-radius:2px;cursor:pointer;flex:1;max-width:200px;}
select:focus{border-color:var(--muted2);}
.btn-sm{background:transparent;border:1px solid var(--border2);color:var(--muted2);font-family:'DM Sans',sans-serif;font-size:12px;padding:9px 18px;cursor:pointer;transition:all .15s;border-radius:2px;white-space:nowrap;}
.btn-sm:hover{border-color:var(--text);color:var(--text);}
.dl-hint{font-size:11px;color:var(--muted);margin-top:10px;min-height:16px;}

.site-footer{border-top:1px solid var(--border);margin-top:60px;padding:20px 40px;font-size:11px;color:var(--muted);line-height:1.6;}

@media(max-width:900px){.grid{grid-template-columns:repeat(2,1fr);}.dl-primary-row{grid-template-columns:1fr;}
  .topbar,.controls,.eyebrow,.dl-section{padding-left:20px;padding-right:20px;}
  .grid,.dl-section{margin-left:0;margin-right:0;}.search-box{width:100%;}}
</style>
</head>
<body>

<div class="topbar">
  <a class="tb-back" href="menu_principal.html">← Menú</a>
  <span class="tb-sep">|</span>
  <span class="tb-title">Fichas de Fondos</span>
</div>

<div class="controls">
  <input class="search-box" id="q" type="search" placeholder="Buscar fondo..." oninput="filtrar()">
  <div class="seg" id="seg">
    <button class="on" data-f="todos">Todos <span class="n">__TOT__</span></button>
    <button data-f="vigente">Vigentes <span class="n">__VIG__</span></button>
    <button data-f="vencido">Vencidos <span class="n">__VEN__</span></button>
  </div>
  <span class="search-count" id="ct"></span>
</div>

<div class="eyebrow" id="eyebrow">Todos los fondos</div>

<div class="grid" id="grid">
__GRID__
  <div class="empty" id="empty" style="display:none">No hay fondos que coincidan.</div>
</div>

<div class="dl-section">
  <div class="dl-title">Descargas</div>
  <div class="dl-sub">Folletos comerciales en PDF — datos del último día del mes anterior.</div>

  <div class="dl-primary-row">
    <button class="btn-big primary" id="btnLatest" onclick="dlLatest()">
      <span class="btn-icon">📦</span>
      <span class="btn-lbl">
        <strong>Folletos del mes — PDF</strong>
        <span id="lbl-mes">Cargando...</span>
      </span>
    </button>
    <button class="btn-big" onclick="dlReglamentos()">
      <span class="btn-icon">📄</span>
      <span class="btn-lbl">
        <strong>Todos los reglamentos</strong>
        <span>__NREG__ PDFs de reglamentos internos</span>
      </span>
    </button>
    <button class="btn-big" onclick="dlProrrogas()">
      <span class="btn-icon">🗓️</span>
      <span class="btn-lbl">
        <strong>Prórrogas de vigencia</strong>
        <span>Actas de directorio 2026</span>
      </span>
    </button>
  </div>

  <hr class="dl-divider">

  <div class="hist-label">Historial de folletos anteriores</div>
  <div class="hist-row">
    <select id="selMes"><option value="">Cargando...</option></select>
    <button class="btn-sm" onclick="dlHistorial()">⬇ Descargar ZIP</button>
  </div>
  <div class="dl-hint" id="hint"></div>
</div>

<script>
const REPO="mivelascor/FIP", BRANCH="main";
let filtro="todos";

/* ── Toggle Todos / Vigentes / Vencidos ── */
document.getElementById("seg").addEventListener("click", e => {
  const b = e.target.closest("button"); if(!b) return;
  [...document.querySelectorAll("#seg button")].forEach(x => x.classList.toggle("on", x===b));
  filtro = b.dataset.f;
  document.getElementById("eyebrow").textContent =
    filtro==="vigente" ? "Fondos vigentes" : filtro==="vencido" ? "Fondos vencidos" : "Todos los fondos";
  filtrar();
});

/* quita tildes para que "dolar" encuentre "Dólar" */
const norm = s => (s||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();

function filtrar(){
  const q = norm(document.getElementById("q").value);
  const cards = [...document.querySelectorAll(".card")];
  let vis = 0;
  cards.forEach(c => {
    const okEstado = filtro==="todos" || c.dataset.estado===filtro;
    const okTexto  = !q || norm(c.innerText).includes(q);
    const show = okEstado && okTexto;
    c.style.display = show ? "flex" : "none";
    if(show) vis++;
  });
  const empty = document.getElementById("empty");
  empty.style.display = vis ? "none" : "block";
  empty.textContent = (!vis && !q && filtro==="vencido")
    ? "No hay fondos marcados como vencidos."
    : "No hay fondos que coincidan.";
  document.getElementById("ct").textContent =
    `${vis} fondo${vis!==1?"s":""}` + (q ? " · filtrado" : "");
}

async function cargarMeses(){
  try{
    const r = await fetch(`https://api.github.com/repos/${REPO}/contents/folletos`,
      {headers:{"Accept":"application/vnd.github+json"}});
    if(!r.ok) throw new Error(r.status);
    const items = await r.json();
    const zips = items
      .filter(i => i.type==="file" && /^folletos_\\d{4}-\\d{2}\\.zip$/.test(i.name))
      .map(i => i.name.replace("folletos_","").replace(".zip",""))
      .sort().reverse();
    if(!zips.length){
      document.getElementById("lbl-mes").textContent = "Sin folletos disponibles";
      document.getElementById("btnLatest").disabled = true;
      document.getElementById("selMes").innerHTML = "<option value=''>Sin historial</option>";
      return;
    }
    document.getElementById("lbl-mes").textContent = zips[0];
    const historial = zips.slice(1);
    document.getElementById("selMes").innerHTML = historial.length
      ? historial.map(m => `<option value="${m}">${m}</option>`).join("")
      : "<option value=''>Sin historial anterior</option>";
  }catch(e){ document.getElementById("lbl-mes").textContent = "Error al cargar"; }
}

function dl(url, name){ const a=document.createElement("a"); a.href=url; a.download=name; a.click(); }

function dlLatest(){
  const mes = document.getElementById("lbl-mes").textContent;
  hint(`Descargando folletos ${mes}...`);
  dl(`https://raw.githubusercontent.com/${REPO}/${BRANCH}/folletos/latest.zip`, `folletos_${mes}.zip`);
  setTimeout(()=>hint(""), 3000);
}
function dlHistorial(){
  const mes = document.getElementById("selMes").value;
  if(!mes){ alert("Selecciona un mes."); return; }
  hint(`Descargando folletos ${mes}...`);
  dl(`https://raw.githubusercontent.com/${REPO}/${BRANCH}/folletos/folletos_${mes}.zip`, `folletos_${mes}.zip`);
  setTimeout(()=>hint(""), 3000);
}
async function dlCarpeta(ruta, etiqueta){
  try{
    hint(`Descargando ${etiqueta}...`);
    const r = await fetch(`https://api.github.com/repos/${REPO}/contents/${ruta}`);
    const data = await r.json();
    const pdfs = (Array.isArray(data)?data:[]).filter(f => f.type==="file" && f.name.toLowerCase().endsWith(".pdf"));
    if(!pdfs.length){ hint("Sin documentos disponibles."); return; }
    pdfs.forEach((f,i) => setTimeout(()=>dl(f.download_url, f.name), i*400));
    setTimeout(()=>hint(""), pdfs.length*400 + 500);
  }catch(e){ hint("Error al descargar."); }
}
function dlReglamentos(){ dlCarpeta("regl.int", "reglamentos"); }
function dlProrrogas(){  dlCarpeta("regl.int/prorrogas", "prórrogas"); }

function hint(msg){ document.getElementById("hint").textContent = msg; }

document.addEventListener("DOMContentLoaded", ()=>{ filtrar(); cargarMeses(); });
</script>
<footer class="site-footer">Conforme a la Ley Única de Fondos, las administradoras de fondos de inversión privados están sujetas a las obligaciones de información establecidas por la Comisión para el Mercado Financiero. Tales fondos no están sometidos a fiscalización de la Comisión y no hacemos oferta pública de sus cuotas.</footer>
</body>
</html>
'''
MENU = (MENU.replace('__GRID__',GRID).replace('__TOT__',str(n_tot))
            .replace('__VIG__',str(n_vig)).replace('__VEN__',str(n_ven))
            .replace('__NREG__',str(sum(1 for x in R if x[7]))))
open(os.path.join(REPO,'MenuFF.html'),'w',encoding='utf-8').write(MENU)
print('MenuFF.html ok —', n_tot,'fondos /',n_vig,'vigentes /',n_ven,'vencidos')

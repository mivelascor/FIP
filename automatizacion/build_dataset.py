#!/usr/bin/env python3
"""
build_dataset.py - Motor de datos para los folletos FIP Vantrust (final).

Principio: los templates son la fuente historica INTOCABLE. Solo se calcula y
agrega el MES NUEVO. La historia se lee del template.

Insumos automaticos:
  ICP         -> API BCCh (icp_bcch.py). VALIDADO 0.0000% error.
  Fondo       -> VC del query (valor_cuota.xlsx) + dividendo.
  Competencia -> CMF sin captcha (competencia_cmf.py): Santander UNIVE / Banchile A.

DOS CONVENCIONES DE RETORNO (confirmadas con VC_para_fichas):
  CLP: retorno ACUMULADO del periodo. nivel = VC + d.
       M/T/S/A = nivel_fin/nivel_{1,3,6,12} - 1 ; Acum = (fin/enero-1)/n*12.
       d se recupera de la hoja `rentabilidad` (validado: Alto Aporte 0.62%).
  USD: retorno ANUALIZADO. nivel H = VC_query + dividendo (columna H del template).
       cada periodo = (H_fin/H_ini - 1) / dias_calendario * 360.
       (validado: Dolar Caja mayo = 5.07%). Competencia USD tambien anualizada.
"""
import openpyxl, datetime, json, os, sys
from collections import defaultdict

FUNDS_CLP = {
 "TEMPLATE FONDO LIQUIDEZ ACTIVA.xlsx":      "FIP VANTRUST LIQUIDEZ ACTIVA",
 "TEMPLATE FONDO ALTO APORTE.xlsx":          "FIP VANTRUST LIQUIDEZ ALTO APORTE",
 "TEMPLATE FONDO ALTO CAPITAL.xlsx":         "FIP VANTRUST LIQUIDEZ ALTO CAPITAL",
 "TEMPLATE FONDO LIQUIDEZ ALTO MONTO.xlsx":  "FIP VANTRUST LIQUIDEZ ALTO MONTO",
 "TEMPLATE FONDO LIQUIDEZ CAJA.xlsx":        "FIP VANTRUST LIQUIDEZ CAJA",
 "TEMPLATE FONDO LIQUIDEZ CONTINUA.xlsx":    "FIP VANTRUST LIQUIDEZ CONTINUA",
 "TEMPLATE FONDO LIQUIDEZ CORRIENTE.xlsx":   "FIP VANTRUST LIQUIDEZ CORRIENTE",
 "TEMPLATE FONDO LIQUIDEZ CORTO PLAZO.xlsx": "FIP VANTRUST LIQUIDEZ CORTO PLAZO",
 "TEMPLATE FONDO LIQUIDEZ Disponible I.xlsx":"FIP VANTRUST LIQUIDEZ DISPONIBLE I",
 "TEMPLATE FONDO LIQUIDEZ EFECTIVO.xlsx":    "FIP VANTRUST LIQUIDEZ EFECTIVO",
 "TEMPLATE FONDO LIQUIDEZ FLEXIBLE.xlsx":    "FIP VANTRUST LIQUIDEZ FLEXIBLE",
 "TEMPLATE FONDO LIQUIDEZ UNO.xlsx":         "FIP VANTRUST LIQUIDEZ I",
 "TEMPLATE FONDO LIQUIDEZ LOCAL.xlsx":       "FIP VANTRUST LIQUIDEZ LOCAL",
 "TEMPLATE FONDO LIQUIDEZ Monetario I.xlsx": "FIP VANTRUST LIQUIDEZ MONETARIO I",
 "TEMPLATE FONDO LIQUIDEZ Permanente.xlsx":  "FIP VANTRUST LIQUIDEZ PERMANENTE",
 "TEMPLATE FONDO LIQUIDEZ PLUS.xlsx":        "FIP VANTRUST LIQUIDEZ PLUS",
 "TEMPLATE FONDO LIQUIDEZ Presente.xlsx":    "FIP VANTRUST LIQUIDEZ PRESENTE",
 "TEMPLATE FONDO LIQUIDEZ RENDIMIENTO.xlsx": "FIP VANTRUST LIQUIDEZ RENDIMIENTO",
 "TEMPLATE FONDO LIQUIDEZ SENCILLO.xlsx":    "FIP VANTRUST LIQUIDEZ SENCILLO",
}
FUNDS_USD = {
 "TEMPLATE FONDO LIQUIDEZ RESERVA DOLAR.xlsx":"FIP VANTRUST LIQUIDEZ RESERVA DOLAR",
 "TEMPLATE FONDO LIQUIDEZ DOLAR.xlsx":        "FIP VANTRUST LIQUIDEZ DOLAR",
 "TEMPLATE FONDO LIQUIDEZ DOLAR CAJA.xlsx":   "FIP VANTRUST LIQUIDEZ DOLAR CAJA",
}


def shift(ym, k=1):
    y, m = map(int, ym.split('-')); m -= k
    while m <= 0: m += 12; y -= 1
    return f"{y}-{m:02d}"


def load_valor_cuota(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Consulta1 (2)"] if "Consulta1 (2)" in wb.sheetnames else wb[wb.sheetnames[0]]
    eom, ld = defaultdict(dict), defaultdict(dict)
    for row in ws.iter_rows(min_row=2, values_only=True):
        f, n, p = row[0], row[1], row[4]
        if n is None or not isinstance(f, datetime.datetime): continue
        nm = str(n).strip().strip("'"); ym = f"{f.year}-{f.month:02d}"
        if ym not in ld[nm] or f.day > ld[nm][ym]:
            ld[nm][ym] = f.day; eom[nm][ym] = float(p)
    return eom


def read_template(tmpl):
    """Niveles ICP(B), comp(K), VC(G), H(H) + fechas(F), y retornos FIP (rentabilidad)."""
    wb = openpyxl.load_workbook(tmpl, data_only=True)
    ws = wb["Datos ICP (2)"]
    icp_lv, comp_lv, G_lv, H_lv, dates = {}, {}, {}, {}, {}
    for row in ws.iter_rows():
        a = row[0].value if len(row) > 0 else None
        f = row[5].value if len(row) > 5 else None
        if isinstance(a, datetime.datetime):
            ym = f"{a.year}-{a.month:02d}"
            if isinstance(row[1].value, (int, float)): icp_lv[ym] = float(row[1].value)
            if len(row) > 10 and isinstance(row[10].value, (int, float)): comp_lv[ym] = float(row[10].value)
        if isinstance(f, datetime.datetime):
            ym = f"{f.year}-{f.month:02d}"; dates[ym] = f.date()
            if isinstance(row[6].value, (int, float)): G_lv[ym] = float(row[6].value)
            if isinstance(row[7].value, (int, float)): H_lv[ym] = float(row[7].value)
    fip = {}
    if "rentabilidad" in wb.sheetnames:
        wr = wb["rentabilidad"]; cy = None
        for r in range(1, wr.max_row + 1):
            c = wr.cell(r, 3).value; d = wr.cell(r, 4).value
            if isinstance(c, (int, float)) and 2000 < c < 2100: cy = int(c)
            if d and isinstance(d, str):
                du = d.strip().upper()
                if du != "ICP" and "COMPETENCIA" not in du and ("FIP" in du or "LIQUIDEZ" in du or "ALTO" in du) and cy:
                    for mi in range(12):
                        v = wr.cell(r, 5 + mi).value
                        if isinstance(v, (int, float)) and v != 0:
                            fip[f"{cy}-{mi+1:02d}"] = float(v)
    return icp_lv, comp_lv, G_lv, H_lv, dates, fip


def recover_dividend(vc, fip_ret):
    """CLP: d del mes oficial mas reciente, validado contra el previo. 0.0 si no calza."""
    oms = [m for m in sorted(fip_ret) if m in vc and shift(m) in vc]
    if not oms: return 0.0
    m = oms[-1]; r = fip_ret[m]
    if r == 0: return 0.0
    d = (vc[m] - (1 + r) * vc[shift(m)]) / r
    if len(oms) >= 2:
        p = oms[-2]; calc = (vc[p] + d) / (vc[shift(p)] + d) - 1
        if abs(calc - fip_ret[p]) > 0.0005: return 0.0
    return d


def metrics_clp(levels, end):
    def ret(e, a):
        return None if (e not in levels or a not in levels or not levels[a]) else round((levels[e]/levels[a]-1)*100, 2)
    y = end.split('-')[0]; jan = f"{y}-01"
    months = [k for k in levels if k.startswith(y) and k <= end]
    acum = round((levels[end]/levels[jan]-1)/len(months)*12*100, 2) if (jan in levels and levels.get(jan) and months) else None
    return {"mensual": ret(end, shift(end, 1)), "trimestral": ret(end, shift(end, 3)),
            "semestral": ret(end, shift(end, 6)), "anual": ret(end, shift(end, 12)), "acum": acum}


def metrics_usd(levels, dates, end):
    """Anualizado: (H_fin/H_ini - 1) / dias_calendario * 360."""
    def ann(k):
        a = shift(end, k)
        if end in levels and a in levels and a in dates and end in dates and levels[a]:
            dias = (dates[end] - dates[a]).days
            return round((levels[end]/levels[a]-1)/dias*360*100, 2) if dias else None
        return None
    y = end.split('-')[0]; jan = f"{y}-01"
    acum = None
    if jan in levels and end in dates and jan in dates and levels.get(jan):
        dias = (dates[end] - dates[jan]).days
        acum = round((levels[end]/levels[jan]-1)/dias*360*100, 2) if dias else None
    return {"mensual": ann(1), "trimestral": ann(3), "semestral": ann(6), "anual": ann(12), "acum": acum}


def build(valor_cuota_path, templates_dir, end):
    eom = load_valor_cuota(valor_cuota_path)
    dataset = {}
    for funds, moneda in [(FUNDS_CLP, "CLP"), (FUNDS_USD, "USD")]:
        for tmpl, nemo in funds.items():
            path = os.path.join(templates_dir, tmpl)
            if not os.path.exists(path): continue
            vc = eom.get(nemo, {})
            if end not in vc: continue
            icp_lv, comp_lv, G_lv, H_lv, dates, fip_ret = read_template(path)
            if not dates.get(end):
                y, m = map(int, end.split('-')); dates[end] = datetime.date(y+(m == 12), (m % 12)+1, 1) - datetime.timedelta(days=1)
            # ICP mes nuevo desde BCCh
            base = max(icp_lv) if icp_lv else None
            try:
                from icp_bcch import icp_fin_de_mes
                if base and end not in icp_lv: icp_lv[end] = icp_fin_de_mes(icp_lv[base], base, end)
            except Exception:
                if base and end not in icp_lv: icp_lv[end] = icp_lv[base]
            # competencia mes nuevo desde CMF
            try:
                from competencia_cmf import valor_cuota_competencia
                comp_lv[end] = valor_cuota_competencia(moneda, end)
            except Exception:
                if comp_lv and end not in comp_lv: comp_lv[end] = comp_lv[max(comp_lv)]

            if moneda == "CLP":
                d = recover_dividend(vc, fip_ret)
                Hf = {m: v + d for m, v in vc.items()}
                fondo = metrics_clp(Hf, end)
                comp = metrics_clp(comp_lv, end) if comp_lv else None
                icp = metrics_clp(icp_lv, end)
            else:  # USD: anualizado, nivel H = VC_query + dividendo
                common = sorted(set(H_lv) & set(G_lv))
                d = (H_lv[common[-1]] - G_lv[common[-1]]) if common else 0.0
                Hf = dict(H_lv)
                if end not in Hf: Hf[end] = vc[end] + d
                fondo = metrics_usd(Hf, dates, end)
                comp = metrics_usd(comp_lv, dates, end) if comp_lv else None
                icp = metrics_usd(icp_lv, dates, end)
            dataset[nemo] = {"template": tmpl, "moneda": moneda, "mes": end,
                             "dividendo": round(d, 6), "vc": round(vc[end], 4),
                             "fondo": fondo, "icp": icp, "competencia": comp}
    return dataset


if __name__ == "__main__":
    vc_path = sys.argv[1] if len(sys.argv) > 1 else "inputs/valor_cuota.xlsx"
    tdir    = sys.argv[2] if len(sys.argv) > 2 else "inputs/templates"
    end     = sys.argv[3] if len(sys.argv) > 3 else datetime.date.today().strftime("%Y-%m")
    ds = build(vc_path, tdir, end)
    print(f"[{end}] {len(ds)} fondos")
    for nemo, x in ds.items():
        f = x["fondo"]
        print(f"  [{x['moneda']}] {nemo.replace('FIP VANTRUST LIQUIDEZ ', ''):<14} "
              f"M={f['mensual']} T={f['trimestral']} S={f['semestral']} A={f['anual']} Acum={f['acum']}")

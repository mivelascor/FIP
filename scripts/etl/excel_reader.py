"""
etl/excel_reader.py — Lee cartera.xlsx subida via admin.html.

Estructura del cartera.xlsx:
  Hoja 'instrumento': fila 1 = encabezados, fila 2+ = montos, bloque % desde fila 15
  Hoja 'duracion':    fila 2 = encabezados, fila 3-6 = montos, bloque % desde fila 13
  Hoja 'moneda':      fila 2 = encabezados, fila 3-5 = montos, bloque % desde fila 12
"""
import openpyxl
from config import ARCHIVO_CARTERA


def _find_col(hdr_row, nombre_fondo: str) -> int:
    """Busca la columna del fondo en la fila de encabezados (match exacto o parcial)."""
    nombre_up = nombre_fondo.upper()
    # Match exacto primero
    for j, h in enumerate(hdr_row):
        if h and str(h).upper().strip() == nombre_up:
            return j
    # Match parcial (los primeros 20 chars)
    nombre_short = nombre_up[:20]
    for j, h in enumerate(hdr_row):
        if h and nombre_short in str(h).upper():
            return j
    return -1


def _find_total_col(hdr_row) -> int:
    """Busca la columna TOTAL."""
    for j, h in enumerate(hdr_row):
        if h and str(h).upper().strip() == "TOTAL":
            return j
    return -1


def _bloque_pct(rows, col_fondo: int, col_label: int = 1) -> dict:
    """
    Lee un bloque de porcentajes (valores entre 0 y 1).
    Retorna dict {label: pct_float}.
    """
    result = {}
    for row in rows:
        label = row[col_label]
        val   = row[col_fondo] if col_fondo < len(row) else None
        if (label and str(label).strip().lower() not in ("", "total", "nan")
                and isinstance(val, (int, float))
                and 0 <= val <= 1):
            result[str(label).strip()] = float(val)
    return result


def get_cartera_composicion(nombre_fondo: str) -> dict:
    """
    Lee composición (moneda, duración, instrumento) desde cartera.xlsx.
    Retorna dict con listas de tuplas (nombre, "XX,XX%").
    """
    result = {"moneda": [], "duracion": [], "instrumento": []}

    try:
        wb = openpyxl.load_workbook(str(ARCHIVO_CARTERA), read_only=True, data_only=True)
    except Exception as e:
        print(f"    [WARN] No se pudo leer cartera.xlsx: {e}")
        return result

    def fmt(v: float) -> str:
        return f"{v*100:.2f}%".replace(".", ",")

    # ── MONEDA ───────────────────────────────────────────────────────────────
    try:
        ws_m  = wb['moneda']
        rows_m = list(ws_m.iter_rows(values_only=True))

        # Buscar el segundo bloque de porcentajes (buscar segunda ocurrencia de COD_MONEDA)
        hdr_rows = [i for i, r in enumerate(rows_m)
                    if r[1] and str(r[1]).upper().strip() == "COD_MONEDA"]

        col_f = -1
        pct_block = []

        if len(hdr_rows) >= 2:
            # Segundo bloque = porcentajes
            hdr_idx = hdr_rows[1]
            hdr     = rows_m[hdr_idx]
            col_f   = _find_col(hdr, nombre_fondo)
            if col_f >= 0:
                pct_block = rows_m[hdr_idx+1:]
        elif len(hdr_rows) == 1:
            # Solo un bloque, buscar si tiene valores 0-1
            hdr_idx = hdr_rows[0]
            hdr     = rows_m[hdr_idx]
            col_f   = _find_col(hdr, nombre_fondo)
            if col_f >= 0:
                col_t = _find_total_col(hdr)
                for row in rows_m[hdr_idx+1:]:
                    monto = row[col_f] if col_f < len(row) else None
                    total = row[col_t] if col_t >= 0 and col_t < len(row) else None
                    if isinstance(monto, (int, float)) and isinstance(total, (int, float)) and total > 0:
                        pct_block.append(list(row[:2]) + [monto/total] + list(row[3:]))
                    else:
                        pct_block.append(row)
                col_f = 2  # ya está en col 2

        if col_f >= 0:
            moneda_dict = _bloque_pct(pct_block, col_f)
            for mon, pct_val in sorted(moneda_dict.items(), key=lambda x: -x[1]):
                if pct_val > 0.001:
                    result["moneda"].append((mon, fmt(pct_val)))
    except Exception as e:
        print(f"    [WARN] Error leyendo moneda: {e}")

    # ── DURACION ─────────────────────────────────────────────────────────────
    ORDEN_TRAMOS = ["Hasta 30 días", "Entre 31 y 90 días", "Entre 91 y 120 días", "Entre 1 y 2 años"]
    try:
        ws_d  = wb['duracion']
        rows_d = list(ws_d.iter_rows(values_only=True))

        # Buscar el segundo bloque de TRAMO (porcentajes)
        tramo_rows = [i for i, r in enumerate(rows_d)
                      if r[1] and str(r[1]).upper().strip() == "TRAMO"]

        col_f = -1
        pct_block = []

        if len(tramo_rows) >= 2:
            hdr_idx = tramo_rows[1]
            hdr     = rows_d[hdr_idx]
            col_f   = _find_col(hdr, nombre_fondo)
            if col_f >= 0:
                pct_block = rows_d[hdr_idx+1:]
        elif len(tramo_rows) == 1:
            # Solo un bloque — calcular % desde montos
            hdr_idx = tramo_rows[0]
            hdr     = rows_d[hdr_idx]
            col_f   = _find_col(hdr, nombre_fondo)
            col_t   = _find_total_col(hdr)
            if col_f >= 0 and col_t >= 0:
                for row in rows_d[hdr_idx+1:]:
                    monto = row[col_f] if col_f < len(row) else None
                    total = row[col_t] if col_t < len(row) else None
                    if isinstance(monto, (int, float)) and isinstance(total, (int, float)) and total > 0:
                        new_row = list(row)
                        new_row[col_f] = monto / total
                        pct_block.append(tuple(new_row))
                    else:
                        pct_block.append(row)

        if col_f >= 0 and pct_block:
            dur_dict = _bloque_pct(pct_block, col_f)
            for tramo in ORDEN_TRAMOS:
                pct_val = dur_dict.get(tramo, 0.0)
                result["duracion"].append((tramo, fmt(pct_val)))
    except Exception as e:
        print(f"    [WARN] Error leyendo duracion: {e}")

    # ── INSTRUMENTO ──────────────────────────────────────────────────────────
    try:
        ws_i  = wb['instrumento']
        rows_i = list(ws_i.iter_rows(values_only=True))

        # Buscar la segunda ocurrencia de CLASIFICACION_OBS (bloque de %)
        cls_rows = [i for i, r in enumerate(rows_i)
                    if r[1] and str(r[1]).upper().strip() == "CLASIFICACION_OBS"]

        col_f = -1
        pct_block = []

        if len(cls_rows) >= 2:
            hdr_idx = cls_rows[1]
            hdr     = rows_i[hdr_idx]
            col_f   = _find_col(hdr, nombre_fondo)
            if col_f >= 0:
                pct_block = rows_i[hdr_idx+1:]
        elif len(cls_rows) == 1:
            # Solo un bloque con montos, calcular %
            hdr_idx = cls_rows[0]
            hdr     = rows_i[hdr_idx]
            col_f   = _find_col(hdr, nombre_fondo)
            col_t   = _find_total_col(hdr)
            if col_f >= 0 and col_t >= 0:
                for row in rows_i[hdr_idx+1:]:
                    monto = row[col_f] if col_f < len(row) else None
                    total = row[col_t] if col_t < len(row) else None
                    if isinstance(monto, (int, float)) and isinstance(total, (int, float)) and total > 0:
                        new_row = list(row)
                        new_row[col_f] = monto / total
                        pct_block.append(tuple(new_row))
                    else:
                        pct_block.append(row)

        if col_f >= 0 and pct_block:
            inst_dict = _bloque_pct(pct_block, col_f)
            for inst, pct_val in sorted(inst_dict.items(), key=lambda x: -x[1]):
                if pct_val > 0.001:
                    result["instrumento"].append((inst, fmt(pct_val)))
    except Exception as e:
        print(f"    [WARN] Error leyendo instrumento: {e}")

    wb.close()
    return result

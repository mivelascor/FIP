"""
generador/graficos.py — Genera imágenes PNG que reemplazan los OLE del PPTX.
Cada función retorna bytes PNG listos para insertar en el PPTX.
"""
import io
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")   # sin pantalla (servidor)
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.patches import FancyBboxPatch

# ── Paleta corporativa ────────────────────────────────────────────────────────
C_FONDO = "#003366"
C_COMP  = "#7F7F7F"
C_ICP   = "#BFA060"
C_BG    = "#FFFFFF"
C_GRID  = "#E8E8E8"
FONT    = "DejaVu Sans"


def _to_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=C_BG, edgecolor="none")
    buf.seek(0)
    plt.close(fig)
    return buf.read()


def _fmt_pct(v, decimals=2) -> str:
    if pd.isna(v):
        return "—"
    return f"{v*100:+.{decimals}f}%"


# ── 1. Gráfico de evolución (líneas base 1000) ────────────────────────────────
def grafico_evolucion(df: pd.DataFrame) -> bytes:
    """
    df columnas: fecha, b1000_fondo, b1000_comp, b1000_icp
    """
    fig, ax = plt.subplots(figsize=(7.5, 4), facecolor=C_BG)
    ax.set_facecolor(C_BG)

    # Rellenar área bajo el fondo
    ax.fill_between(df["fecha"], df["b1000_fondo"],
                    alpha=0.08, color=C_FONDO)

    ax.plot(df["fecha"], df["b1000_fondo"], label="Fondo",
            color=C_FONDO, linewidth=2.2, zorder=3)
    ax.plot(df["fecha"], df["b1000_comp"], label="Competencia",
            color=C_COMP, linewidth=1.5, linestyle="--", zorder=2)
    ax.plot(df["fecha"], df["b1000_icp"], label="ICP (Benchmark)",
            color=C_ICP, linewidth=1.5, linestyle=":", zorder=2)

    ax.legend(fontsize=8, framealpha=0.9, loc="upper left")
    ax.set_ylabel("Base 1.000", fontsize=8, color="#555")
    ax.yaxis.set_major_formatter(mtick.FormatStrFormatter("%.0f"))
    ax.tick_params(labelsize=7, colors="#555")
    ax.grid(axis="y", color=C_GRID, linewidth=0.7)
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color("#CCC")

    fig.tight_layout(pad=0.5)
    return _to_bytes(fig)


# ── 2. Tabla de rentabilidades ────────────────────────────────────────────────
def tabla_rentabilidades_img(metricas_df: pd.DataFrame) -> bytes:
    """
    metricas_df: nombre, mensual, trimestral, semestral, anual, ytd
    """
    fig, ax = plt.subplots(figsize=(7.5, 1.8), facecolor=C_BG)
    ax.set_facecolor(C_BG)
    ax.axis("off")

    col_labels = ["", "Mensual", "Trimestral", "Semestral", "Anual\n(12M)", "YTD"]
    data = []
    colors_rows = []

    for _, r in metricas_df.iterrows():
        row = [
            r["nombre"],
            _fmt_pct(r["mensual"]),
            _fmt_pct(r["trimestral"]),
            _fmt_pct(r["semestral"]),
            _fmt_pct(r["anual"]),
            _fmt_pct(r["ytd"]),
        ]
        data.append(row)
        colors_rows.append(["#FFFFFF"] * 6)

    # Colorear fila del fondo
    colors_rows[0] = ["#EDF2FA"] * 6

    t = ax.table(
        cellText=data,
        colLabels=col_labels,
        cellLoc="center",
        loc="center",
        cellColours=colors_rows,
    )
    t.auto_set_font_size(False)
    t.set_fontsize(8.5)
    t.scale(1, 1.7)

    # Encabezados
    for ci in range(len(col_labels)):
        cell = t[0, ci]
        cell.set_facecolor(C_FONDO)
        cell.set_text_props(color="white", fontweight="bold", fontsize=8)

    # Primera columna en negrita
    for ri in range(1, len(data) + 1):
        t[ri, 0].set_text_props(fontweight="bold", ha="left")

    fig.tight_layout(pad=0.2)
    return _to_bytes(fig)


# ── 3. Gráfico composición por duración (donut) ───────────────────────────────
def grafico_composicion_duracion(df_cartera: pd.DataFrame) -> bytes:
    """
    df_cartera: duracion, pct
    """
    if df_cartera.empty:
        fig, ax = plt.subplots(figsize=(4, 2.5))
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        ax.axis("off")
        return _to_bytes(fig)

    dur = df_cartera.groupby("duracion")["pct"].sum()
    dur = dur[dur > 0]

    COLORES_DUR = ["#003366", "#1A5599", "#4A90D9", "#A8C8F0", "#D0E5F7"]

    fig, ax = plt.subplots(figsize=(4.5, 2.8), facecolor=C_BG)
    ax.set_facecolor(C_BG)

    wedges, texts, autotexts = ax.pie(
        dur.values,
        labels=None,
        autopct=lambda p: f"{p:.1f}%" if p > 3 else "",
        colors=COLORES_DUR[:len(dur)],
        startangle=90,
        pctdistance=0.75,
        wedgeprops={"width": 0.55, "edgecolor": "white", "linewidth": 1.5},
    )
    for at in autotexts:
        at.set_fontsize(7.5)
        at.set_color("white")
        at.set_fontweight("bold")

    ax.legend(
        wedges, dur.index.tolist(),
        loc="center left", bbox_to_anchor=(0.85, 0.5),
        fontsize=7, frameon=False,
    )
    ax.set_title("Por Duración", fontsize=8.5, color="#333", pad=4)
    fig.tight_layout(pad=0.3)
    return _to_bytes(fig)


# ── 4. Tabla comparación mensual ──────────────────────────────────────────────
def tabla_comparacion_mensual(df_rentm: pd.DataFrame) -> bytes:
    """
    df_rentm: mes, rent_fondo, rent_comp, rent_icp
    """
    fig, ax = plt.subplots(figsize=(7.5, 3.2), facecolor=C_BG)
    ax.set_facecolor(C_BG)
    ax.axis("off")

    col_labels = ["Mes", "Fondo", "Competencia", "ICP"]
    data = []
    cell_colors = []

    for _, r in df_rentm.iterrows():
        rf = r["rent_fondo"]
        rc = r["rent_comp"]
        ri = r["rent_icp"]
        data.append([r["mes"], _fmt_pct(rf), _fmt_pct(rc), _fmt_pct(ri)])

        # Color verde si fondo supera benchmark, rojo si no
        if pd.notna(rf) and pd.notna(ri):
            c_fondo = "#E8F5E9" if rf >= ri else "#FFEBEE"
        else:
            c_fondo = "#FFFFFF"
        cell_colors.append(["#FFFFFF", c_fondo, "#FFFFFF", "#FFFFFF"])

    t = ax.table(
        cellText=data,
        colLabels=col_labels,
        cellLoc="center",
        loc="center",
        cellColours=cell_colors,
    )
    t.auto_set_font_size(False)
    t.set_fontsize(8)
    t.scale(1, 1.45)

    for ci in range(len(col_labels)):
        cell = t[0, ci]
        cell.set_facecolor(C_FONDO)
        cell.set_text_props(color="white", fontweight="bold", fontsize=8)

    fig.tight_layout(pad=0.2)
    return _to_bytes(fig)

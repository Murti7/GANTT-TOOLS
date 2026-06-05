"""
Tokens de estilo para los tres módulos de generación de reportes.

Importa colores y fuentes de palette.py y construye los objetos openpyxl
y las constantes matplotlib que cada módulo necesita.
Ningún módulo de generación define colores ni fuentes propios.

Secciones:
  PRES_*   → presupuesto_exporter  (openpyxl, documentos licitación)
  GANTT_*  → excel_exporter        (openpyxl, hojas análisis / Gantt)
  CHART_*  → analysis_charts       (matplotlib)
"""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from gantt.reporting.palette import (
    FONT_PRES, FONT_EXCEL, FONT_CHART,
    NEGRO, BLANCO,
    GRIS_OSCURO, GRIS_MEDIO, GRIS_SUAVE, GRIS_LINEA, GRIS_TEXTO,
    AZUL_OSCURO, AZUL_MEDIO, AZUL_NEUTRO,
    VERDE_EXITO, VERDE_EXITO_FONDO, ROJO_ALERTA, ROJO_ALERTA_FONDO,
    AMARILLO_ADV, AMARILLO_EDIT_PRES, AMARILLO_EDIT_GANTT,
    CHART_AZUL, CHART_AZUL_CLARO, CHART_AZUL_MEDIO,
    CHART_VERDE, CHART_VERDE_CLARO, CHART_NARANJA, CHART_NARANJA_CL,
    CHART_TIERRA, CHART_ROSAOSCURO, CHART_MORADO, CHART_CYAN,
    CHART_GRIS, CHART_GRIS_CLARO, CHART_BORDE, CHART_GRID,
    CHART_TEXTO, CHART_MUTED,
)

# ── Presupuesto — Fills ───────────────────────────────────────────────────────
PRES_FC  = PatternFill('solid', fgColor=GRIS_OSCURO)    # cabecera principal
PRES_FCP = PatternFill('solid', fgColor=GRIS_MEDIO)     # cabecera tabla / capítulo
PRES_FSC = PatternFill('solid', fgColor=GRIS_SUAVE)     # subcapítulo
PRES_FD  = PatternFill('solid', fgColor=GRIS_SUAVE)     # detalle / alterno
PRES_FB  = PatternFill('solid', fgColor=BLANCO)         # blanco
PRES_FA  = PatternFill('solid', fgColor=AMARILLO_EDIT_PRES)  # editable

# ── Presupuesto — Fonts ───────────────────────────────────────────────────────
PRES_FT  = Font(name=FONT_PRES, size=12, bold=True, color=BLANCO)   # título sección
PRES_FCF = Font(name=FONT_PRES, size=10, bold=True, color=NEGRO)    # capítulo / negrita
PRES_FN  = Font(name=FONT_PRES, size=10,            color=NEGRO)    # normal
PRES_FDE = Font(name=FONT_PRES, size=9,             color=GRIS_TEXTO)  # detalle
PRES_FTO = Font(name=FONT_PRES, size=10, bold=True, color=NEGRO)    # total

# ── Presupuesto — Alineaciones ────────────────────────────────────────────────
PRES_AD  = Alignment(horizontal='left',   vertical='top',    wrap_text=True)
PRES_AN  = Alignment(horizontal='right',  vertical='top')
PRES_ANC = Alignment(horizontal='right',  vertical='center')
PRES_AC  = Alignment(horizontal='center', vertical='center')
PRES_AT  = Alignment(horizontal='left',   vertical='top')
PRES_ALC = Alignment(horizontal='left',   vertical='center')

# ── Presupuesto — Bordes ──────────────────────────────────────────────────────
PRES_BORDE_FINO  = Border(bottom=Side(style='thin', color=GRIS_LINEA))
PRES_BORDE_TABLA = Border(
    left=Side(style='thin', color=GRIS_LINEA),
    right=Side(style='thin', color=GRIS_LINEA),
    top=Side(style='thin', color=GRIS_LINEA),
    bottom=Side(style='thin', color=GRIS_LINEA),
)
PRES_BORDE_TOTAL = Border(top=Side(style='medium', color=NEGRO))

# ── Presupuesto — Formatos numéricos ─────────────────────────────────────────
PRES_FE = '#,##0.00 €'   # euros con símbolo
PRES_FQ = '#,##0.###'    # cantidades sin ceros sobrantes
PRES_FU = '#,##0.000'    # cantidades de recursos (3 decimales fijos)
PRES_FP = '0.###%'       # porcentajes sin ceros sobrantes

# ── Gantt / Análisis — Fills ──────────────────────────────────────────────────
GANTT_FILL_HEADER  = PatternFill('solid', fgColor=AZUL_OSCURO)
GANTT_FILL_SUBHEAD = PatternFill('solid', fgColor=AZUL_MEDIO)
GANTT_FILL_CAP     = PatternFill('solid', fgColor=GRIS_MEDIO)
GANTT_FILL_ALT     = PatternFill('solid', fgColor=GRIS_SUAVE)
GANTT_FILL_WHITE   = PatternFill('solid', fgColor=BLANCO)
GANTT_FILL_GREEN   = PatternFill('solid', fgColor=VERDE_EXITO_FONDO)
GANTT_FILL_RED     = PatternFill('solid', fgColor=ROJO_ALERTA_FONDO)
GANTT_FILL_YELLOW  = PatternFill('solid', fgColor=AMARILLO_EDIT_GANTT)

# ── Gantt / Análisis — Fonts ──────────────────────────────────────────────────
GANTT_FONT_HEADER = Font(bold=True, color=BLANCO,  name=FONT_EXCEL, size=10)
GANTT_FONT_BOLD   = Font(bold=True,                name=FONT_EXCEL, size=10)
GANTT_FONT_ITALIC = Font(italic=True,              name=FONT_EXCEL, size=10)
GANTT_FONT_NORMAL = Font(                          name=FONT_EXCEL, size=10)

# ── Gantt / Análisis — Alineaciones ──────────────────────────────────────────
GANTT_ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
GANTT_ALIGN_LEFT   = Alignment(horizontal='left',   vertical='center')

# ── Gantt / Análisis — Formatos numéricos ─────────────────────────────────────
GANTT_FMT_EUROS  = '#,##0.00'
GANTT_FMT_HORAS  = '#,##0.0'
GANTT_FMT_ENTERO = '#,##0'
GANTT_FMT_PCT    = '0.0'

# ── Gráficos matplotlib ───────────────────────────────────────────────────────
CHART_PALETA = [
    CHART_AZUL, CHART_VERDE, CHART_NARANJA,
    CHART_MORADO, CHART_CYAN, CHART_GRIS,
]
# Paleta extendida para gráficos apilados con muchas series (pares oscuro/claro)
CHART_PALETA_APILADA = [
    CHART_AZUL, CHART_AZUL_CLARO,
    CHART_VERDE, CHART_VERDE_CLARO,
    CHART_NARANJA, CHART_NARANJA_CL,
    CHART_TIERRA, CHART_ROSAOSCURO,
    CHART_GRIS,
]
CHART_DPI = 300

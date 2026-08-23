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

import re
from dataclasses import dataclass
from datetime import date

from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from gantt.bc3.models import BrandingConfig
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
PRES_FI  = Font(name=FONT_PRES, size=10, italic=True, color=NEGRO)  # nota editable
PRES_FDEI = Font(name=FONT_PRES, size=9, italic=True, color=NEGRO)  # nota detalle editable
PRES_FDEB = Font(name=FONT_PRES, size=9, bold=True, color=GRIS_TEXTO)  # detalle destacado
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


HEX_COLOR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def validate_hex_color(value: str) -> str:
    """Valida y normaliza un color hexadecimal #RRGGBB."""
    if not isinstance(value, str) or not HEX_COLOR_RE.match(value):
        raise ValueError(f"Color hexadecimal invalido: {value!r}")
    return value.upper()


def as_matplotlib_hex(value: str) -> str:
    """Devuelve un color con # para APIs matplotlib."""
    return value if value.startswith("#") else f"#{value}"


@dataclass(frozen=True)
class ChartStyle:
    """Tokens visuales para graficos matplotlib."""
    primary: str
    secondary: str
    accent: str
    series: tuple[str, ...]
    stacked: tuple[str, ...]
    grid: str
    text: str
    muted: str
    border: str
    background: str
    font_family: str
    dpi: int


@dataclass(frozen=True)
class DiagramStyle:
    """Tokens visuales para el diagrama temporal de planificacion."""
    critical_fill: str
    critical_border: str
    task_fill: str
    task_border: str
    milestone_fill: str
    milestone_border: str
    finish_fill: str
    post_fill: str
    post_border: str
    dependency: str
    dependency_critical: str
    document: str
    grid: str
    text: str
    band_default: str
    panel_separator: str


# ── Branding corporativo — paleta dinámica por empresa ────────────────────────
# A diferencia de las secciones anteriores (paleta fija del proyecto), esta
# sección construye estilos openpyxl a partir del BrandingConfig de la empresa
# cargada para el proyecto, con fallback a colores neutros si no hay empresa.

def color_texto_sobre(hex_color: str) -> str:
    """
    Retorna '#FFFFFF' o '#000000' según el contraste del fondo.
    Usa la fórmula de luminancia relativa W3C.
    """
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    luminancia = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return '#000000' if luminancia > 0.5 else '#FFFFFF'


@dataclass
class DocumentPalette:
    """
    Paleta de estilos openpyxl construida a partir de BrandingConfig.
    Se pasa como parámetro a todas las funciones de exportación.
    """
    fill_header:   PatternFill
    fill_subhead:  PatternFill
    fill_alt:      PatternFill
    fill_white:    PatternFill
    fill_yellow:   PatternFill
    font_header:   Font
    font_subhead:  Font
    font_bold:     Font
    font_normal:   Font
    font_small:    Font
    color_primario:    str
    color_secundario:  str
    color_texto:       str
    color_acento:      str
    color_blanco:      str
    color_fondo_alt:   str
    font_family:       str
    chart:             ChartStyle
    diagram:           DiagramStyle


def build_palette(company: BrandingConfig | None = None) -> DocumentPalette:
    """
    Construye la DocumentPalette a partir del BrandingConfig.
    Si company es None usa la paleta corporativa por defecto.
    """
    c = company or BrandingConfig()

    fuente = c.fuente_principal
    primary = validate_hex_color(c.color_primario)
    secondary = validate_hex_color(c.color_secundario)
    accent = validate_hex_color(c.color_acento)
    text = validate_hex_color(c.color_texto)
    alt = validate_hex_color(c.color_fondo_alt)
    white = validate_hex_color(c.color_blanco)

    def fill(hex_color: str) -> PatternFill:
        return PatternFill('solid', fgColor=hex_color.lstrip('#'))

    def font_sobre(hex_color: str, bold: bool = False, size: int = 10) -> Font:
        color_txt = color_texto_sobre(hex_color).lstrip('#')
        return Font(name=fuente, bold=bold, size=size, color=color_txt)

    return DocumentPalette(
        fill_header   = fill(primary),
        fill_subhead  = fill(secondary),
        fill_alt      = fill(alt),
        fill_white    = fill(white),
        fill_yellow   = fill('#FFF3CD'),
        font_header   = font_sobre(primary, bold=True, size=10),
        font_subhead  = font_sobre(secondary, bold=True, size=10),
        font_bold     = Font(name=fuente, bold=True, size=10,
                            color=text.lstrip('#')),
        font_normal   = Font(name=fuente, size=10,
                            color=text.lstrip('#')),
        font_small    = Font(name=fuente, size=8,
                            color=text.lstrip('#')),
        color_primario   = primary,
        color_secundario = secondary,
        color_texto      = text,
        color_acento     = accent,
        color_blanco     = white,
        color_fondo_alt  = alt,
        font_family      = fuente,
        chart=ChartStyle(
            primary=primary,
            secondary=secondary,
            accent=accent,
            series=(primary, secondary, accent, CHART_NARANJA, CHART_MORADO, CHART_CYAN),
            stacked=(
                primary, secondary, accent, CHART_AZUL_CLARO, CHART_VERDE_CLARO,
                CHART_NARANJA, CHART_NARANJA_CL, CHART_TIERRA, CHART_GRIS,
            ),
            grid=CHART_GRID,
            text=text,
            muted=CHART_MUTED,
            border=CHART_BORDE,
            background=white,
            font_family=FONT_CHART,
            dpi=CHART_DPI,
        ),
        diagram=DiagramStyle(
            critical_fill="#FFD6D6",
            critical_border=as_matplotlib_hex(ROJO_ALERTA),
            task_fill=alt,
            task_border=secondary,
            milestone_fill=as_matplotlib_hex(AMARILLO_ADV),
            milestone_border=text,
            finish_fill=as_matplotlib_hex(VERDE_EXITO_FONDO),
            post_fill=as_matplotlib_hex(GRIS_SUAVE),
            post_border=as_matplotlib_hex(GRIS_TEXTO),
            dependency=text,
            dependency_critical=as_matplotlib_hex(ROJO_ALERTA),
            document=as_matplotlib_hex(GRIS_TEXTO),
            grid=as_matplotlib_hex(GRIS_LINEA),
            text=text,
            band_default=as_matplotlib_hex(GRIS_SUAVE),
            panel_separator=as_matplotlib_hex(GRIS_LINEA),
        ),
    )


# Literales por idioma para la cabecera y el pie de documento
LITERALES: dict[str, dict[str, str]] = {
    'es': {
        'generado':    'Documento generado el',
        'revision':    'Rev.',
        'pagina':      'Pág.',
        'pendiente':   'PENDIENTE',
        'transferencia': 'Transferencia bancaria',
        'web':         'Web',
        'email':       'Email',
        'tel':         'Tel',
        'pres_03_01':  'Mediciones',
        'pres_03_02':  'Mediciones ciegas',
    },
    'ca': {
        'generado':    'Document generat el',
        'revision':    'Rev.',
        'pagina':      'Pàg.',
        'pendiente':   'PENDENT',
        'transferencia': 'Transferència bancària',
        'web':         'Web',
        'email':       'Email',
        'tel':         'Tel',
        'pres_03_01':  'Amidaments',
        'pres_03_02':  'Amidaments cecs',
    },
    'en': {
        'generado':    'Document generated on',
        'revision':    'Rev.',
        'pagina':      'Page',
        'pendiente':   'PENDING',
        'transferencia': 'Bank transfer',
        'web':         'Web',
        'email':       'Email',
        'tel':         'Tel',
        'pres_03_01':  'Measurements',
        'pres_03_02':  'Blind measurements',
    },
}


def inserir_cabecera(
    ws,
    palette: DocumentPalette,
    company: BrandingConfig | None,
    titol_document: str,
    num_columnes: int,
    ref_projecte: str = "",
    data_generacio: str = "",
    revisio: str = "00",
) -> int:
    """
    Inserta la cabecera corporativa estándar en el worksheet.
    Ocupa las filas 1-4. Retorna 6 (primera fila de contenido,
    tras fila 5 vacía separadora).

    Estructura:
        Fila 1: Logo (col A-B) | Nombre empresa (col C+)
        Fila 2: Logo           | NIF | Tel | Email | Web
        Fila 3: Logo           | Dirección fiscal
        Fila 4: Título documento | Ref. proyecto | Fecha | Rev.
        Fila 5: vacía separadora
    """
    lit = LITERALES.get(company.idioma if company else 'es', LITERALES['es'])
    data = data_generacio or date.today().strftime('%d/%m/%Y')
    col_max = get_column_letter(num_columnes)

    # ── Fila 1: nombre empresa ─────────────────────────────────────────
    ws.row_dimensions[1].height = 28
    ws.merge_cells(f'C1:{col_max}1')
    c = ws['C1']
    c.value = company.nombre if company else ''
    c.font = Font(name=palette.font_normal.name, bold=True, size=13,
                  color=palette.color_primario.lstrip('#'))
    c.alignment = Alignment(horizontal='left', vertical='center')

    # ── Fila 2: NIF | Tel | Email ──────────────────────────────────────
    ws.row_dimensions[2].height = 16
    ws.merge_cells(f'C2:{col_max}2')
    if company:
        nif = company.nif_cif or lit['pendiente']
        tel = f"{lit['tel']}: {company.telefono}" if company.telefono else ''
        email = f"{lit['email']}: {company.email}" if company.email else ''
        parts = [f'NIF: {nif}', tel, email]
        ws['C2'].value = '  |  '.join(p for p in parts if p)
    ws['C2'].font = palette.font_small
    ws['C2'].alignment = Alignment(horizontal='left', vertical='center')

    # ── Fila 3: dirección | web ────────────────────────────────────────
    ws.row_dimensions[3].height = 16
    ws.merge_cells(f'C3:{col_max}3')
    if company:
        dir_str = company.direccion_fiscal or ''
        web_str = f"{lit['web']}: {company.web}" if company.web else ''
        parts = [p for p in [dir_str, web_str] if p]
        ws['C3'].value = '  |  '.join(parts)
    ws['C3'].font = palette.font_small
    ws['C3'].alignment = Alignment(horizontal='left', vertical='center')

    # ── Fila 4: título documento ───────────────────────────────────────
    ws.row_dimensions[4].height = 22
    ws.merge_cells(f'A4:{col_max}4')
    c = ws['A4']
    parts_titol = [titol_document]
    if ref_projecte:
        parts_titol.append(ref_projecte)
    rev_str = f"{lit['revision']} {revisio}"
    data_str = f"{lit['generado']} {data}"
    parts_titol += [rev_str, data_str]
    c.value = '   |   '.join(parts_titol)
    c.font = Font(name=palette.font_header.name, bold=True, size=10,
                  color='FFFFFF')
    c.fill = palette.fill_header
    c.alignment = Alignment(horizontal='left', vertical='center')

    # ── Fila 5: separadora vacía ───────────────────────────────────────
    # Se toca la celda A5 (sin valor) para que ws.max_row registre esta fila;
    # de lo contrario openpyxl no la contabiliza y el siguiente ws.append()
    # de la hoja sobreescribiría la fila separadora en lugar de empezar en 6.
    ws.cell(row=5, column=1)
    ws.row_dimensions[5].height = 6

    # ── Logo (filas 1-3, columnas A-B) ────────────────────────────────
    if company and company.logo_path:
        try:
            img = XLImage(str(company.logo_path))
            img.width  = 160
            img.height = 45
            img.anchor = 'A1'
            ws.add_image(img)
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 12
        except Exception:
            pass   # logo no disponible (formato no soportado por openpyxl) — continuar sin él

    return 6   # primera fila de contenido


def inserir_peu(
    ws,
    palette: DocumentPalette,
    company: BrandingConfig | None,
    num_columnes: int,
) -> None:
    """Inserta pie de documento: fila vacía + fila de datos empresa."""
    col_max = get_column_letter(num_columnes)
    fila_peu = ws.max_row + 2
    ws.row_dimensions[fila_peu].height = 14
    ws.merge_cells(f'A{fila_peu}:{col_max}{fila_peu}')
    c = ws.cell(row=fila_peu, column=1)
    if company:
        c.value = (f'{company.nombre}  |  NIF: {company.nif_cif}  |  '
                   f'{company.web}  |  {company.email}')
    c.font = palette.font_small
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.fill = palette.fill_alt

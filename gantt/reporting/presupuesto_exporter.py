"""
Genera los documentos de presupuesto para el pliego de contratación.

Siete ficheros Excel independientes: PRES.01, PRES.02.01, PRES.02.02,
PRES.02.03, PRES.02.04, PRES.05 y JUST_PRECIOS_Recursos.xlsx.
"""

from pathlib import Path
import math
import re
import textwrap
from weakref import WeakKeyDictionary

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side  # usados inline en helpers de layout
from openpyxl.utils import get_column_letter

from gantt.bc3.models import Capitulo, Partida, Presupuesto, ProjectConfig
from gantt.bc3.parser import expandir_descompuesto


from gantt.reporting.styles import (
    PRES_FC, PRES_FCP, PRES_FSC, PRES_FD, PRES_FB, PRES_FA,
    PRES_FT, PRES_FCF, PRES_FN, PRES_FDE, PRES_FTO,
    PRES_AD, PRES_AN, PRES_ANC, PRES_AC, PRES_AT, PRES_ALC,
    PRES_FE, PRES_FQ, PRES_FU, PRES_FP,
    PRES_BORDE_FINO, PRES_BORDE_TABLA, PRES_BORDE_TOTAL,
)
from gantt.reporting.palette import FONT_PRES, GRIS_LINEA, GRIS_TEXTO


# ── Tablas para conversión numérica ──────────────────────────────────────────
_UNIDADES = [
    '', 'UN', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE',
    'OCHO', 'NUEVE', 'DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE',
    'QUINCE', 'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE',
    'VEINTE', 'VEINTIÚN', 'VEINTIDÓS', 'VEINTITRÉS', 'VEINTICUATRO',
    'VEINTICINCO', 'VEINTISÉIS', 'VEINTISIETE', 'VEINTIOCHO', 'VEINTINUEVE',
]
_DECENAS  = ['', '', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA',
             'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA']
_CENTENAS = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS',
             'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS']


def grupo_a_letras(n: int) -> str:
    """Convierte 1-999 a palabras en castellano (uso interno)."""
    if n == 100:
        return 'CIEN'
    partes = []
    if n >= 100:
        partes.append(_CENTENAS[n // 100])
        n %= 100
    if n == 0:
        return ' '.join(partes)
    if n < 30:
        partes.append(_UNIDADES[n])
    else:
        uni = n % 10
        partes.append(_DECENAS[n // 10] + (' Y ' + _UNIDADES[uni] if uni else ''))
    return ' '.join(p for p in partes if p)


def numero_a_letras(importe: float) -> str:
    """
    Convierte un importe numérico a texto en castellano en MAYÚSCULAS.
    Cubre importes hasta 999.999,99 €.
    Ejemplo: 37.66 → 'TREINTA Y SIETE EUROS CON SESENTA Y SEIS CÉNTIMOS'
    """
    total_cents = round(importe * 100)
    euros = total_cents // 100
    cents = total_cents % 100

    if total_cents == 0:
        return 'CERO EUROS'

    partes: list[str] = []

    if euros > 0:
        miles = euros // 1000
        resto = euros % 1000
        euros_txt: list[str] = []
        if miles == 1:
            euros_txt.append('MIL')
        elif miles > 1:
            euros_txt.append(grupo_a_letras(miles) + ' MIL')
        if resto > 0:
            euros_txt.append(grupo_a_letras(resto))
        partes.extend(euros_txt)
        partes.append('EURO' if euros == 1 else 'EUROS')

    if cents > 0:
        if euros > 0:
            partes.append('CON')
        partes.append(grupo_a_letras(cents))
        partes.append('CÉNTIMO' if cents == 1 else 'CÉNTIMOS')

    return ' '.join(partes)


# ── Helpers de estilo ─────────────────────────────────────────────────────────

def aplicar_estilo_cabecera(ws, fila: int, n_cols: int) -> None:
    """Aplica estilo de cabecera (azul oscuro, texto blanco) a una fila fusionada."""
    if n_cols > 1:
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=n_cols)
    cell = ws.cell(fila, 1)
    cell.fill      = PRES_FC
    cell.font      = PRES_FT
    cell.alignment = PRES_AC
    for c in range(1, n_cols + 1):
        ws.cell(fila, c).border = PRES_BORDE_TABLA
    ws.row_dimensions[fila].height = 26.0
    registrar_altura_manual(ws, fila)


def aplicar_estilo_capitulo(ws, fila: int, n_cols: int) -> None:
    """Aplica estilo de capítulo (azul claro, negrita) a las celdas de una fila."""
    for c in range(1, n_cols + 1):
        ws.cell(fila, c).fill = PRES_FCP
        ws.cell(fila, c).font = PRES_FCF
        ws.cell(fila, c).border = PRES_BORDE_TABLA




def aplicar_estilo_subtotal(ws, fila: int, n_cols: int) -> None:
    """Aplica estilo de subtotal: blanco, negrita y borde superior fino/medio."""
    for c in range(1, n_cols + 1):
        ws.cell(fila, c).fill = PRES_FB
        ws.cell(fila, c).font = PRES_FCF
        ws.cell(fila, c).border = PRES_BORDE_TOTAL
        ws.cell(fila, c).alignment = PRES_ANC
    ws.row_dimensions[fila].height = 22.0
    registrar_altura_manual(ws, fila)

def aplicar_estilo_total(ws, fila: int, n_cols: int) -> None:
    """Aplica estilo de total sobrio: blanco, negrita y borde superior."""
    for c in range(1, n_cols + 1):
        ws.cell(fila, c).fill = PRES_FB
        ws.cell(fila, c).font = PRES_FTO
        ws.cell(fila, c).border = PRES_BORDE_TOTAL
    ws.row_dimensions[fila].height = 24.0
    registrar_altura_manual(ws, fila)


def ajustar_columnas(ws, config: dict[str, int]) -> None:
    """Ajusta anchos de columna según config {letra_columna: ancho}."""
    for col_letter, width in config.items():
        ws.column_dimensions[col_letter].width = width


def aplicar_encabezado_documental(
    ws, config: ProjectConfig, anexo: str, titulo: str, n_cols: int
) -> None:
    """Escribe el encabezado documental. Las líneas vacías del config se omiten."""
    exp = f'Nº Exp.: {config.numero_expediente}' if config.numero_expediente else ''
    candidatas = [
        config.entidad,
        config.proyecto,
        config.edificio,
        exp,
        f'ANEXO ECONÓMICO {anexo}',
        titulo,
    ]
    filas = [f for f in candidatas if f]
    for i, texto in enumerate(filas):
        es_titulo = (i == len(filas) - 1)
        ws.append([texto] + [''] * (n_cols - 1))
        fila = ws.max_row
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=n_cols)
        cell = ws.cell(fila, 1)
        cell.alignment = PRES_AC
        cell.fill = PRES_FC if es_titulo else PRES_FB
        cell.font = (
            Font(name=FONT_PRES, size=12, bold=True, color='FFFFFF')
            if es_titulo else
            Font(name=FONT_PRES, size=10, bold=True, color='000000')
        )
        ws.row_dimensions[fila].height = 18
        registrar_altura_manual(ws, fila)


def configurar_impresion(
    ws,
    orientacion: str,
    fila_cabecera: int | None = None,
    config: ProjectConfig | None = None,
) -> None:
    """Configura pie de página, márgenes e impresión ajustada a ancho."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.font.name in (None, 'Calibri', FONT_PRES):
                cell.font = PRES_FN
    ws.oddFooter.left.text   = config.footer_org if config else ''
    ws.oddFooter.center.text = config.footer_exp if config else ''
    ws.oddFooter.right.text = 'Página &[Page] de &[Pages]'
    ws.page_setup.orientation = orientacion
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.7
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3
    ws.sheet_view.showGridLines = False
    ws.print_area = ws.dimensions
    if fila_cabecera is not None:
        ws.print_title_rows = f'{fila_cabecera}:{fila_cabecera}'


# ── Helpers de datos ─────────────────────────────────────────────────────────

def get_recurso_info(codigo: str, presupuesto: Presupuesto) -> tuple[str, str, float]:
    """
    Retorna (descripcion, unidad, precio) para un código de recurso o partida alzada.
    Busca en MO, MT, MQ y PA; devuelve el código como descripción si no se encuentra.
    """
    if codigo in presupuesto.recursos_mo:
        r = presupuesto.recursos_mo[codigo]
        return r.descripcion, 'h', r.precio_hora
    if codigo in presupuesto.recursos_mt:
        r = presupuesto.recursos_mt[codigo]
        return r.descripcion, r.unidad, r.precio_unidad
    if codigo in presupuesto.recursos_mq:
        r = presupuesto.recursos_mq[codigo]
        return r.descripcion, r.unidad, r.precio_unidad
    if codigo in presupuesto.recursos_pa:
        r = presupuesto.recursos_pa[codigo]
        return r.descripcion, r.unidad, r.precio_unidad
    return codigo, '-', 0.0


def recursos_expandidos(codigo: str, presupuesto: Presupuesto) -> dict[str, float]:
    """Retorna {codigo_recurso: cantidad_efectiva} expandiendo PAs recursivamente."""
    return expandir_descompuesto(
        codigo,
        presupuesto.descompuestos_raw,
        set(presupuesto.recursos_pa),
    )


def recursos_para_display(codigo: str, presupuesto: Presupuesto) -> dict[str, float]:
    """Retorna recursos del descompuesto sin expandir auxiliares (PAs).
    Uso exclusivo para documentos: los auxiliares aparecen como línea única 'AUX'."""
    return expandir_descompuesto(codigo, presupuesto.descompuestos_raw, set())


def tipo_recurso(codigo: str, presupuesto: Presupuesto) -> str:
    """
    Clasifica un código en MO, MT, MQ, PA, % o '?' (no reconocido).
    Usa los diccionarios del presupuesto como fuente de verdad.
    Los costes indirectos se detectan por el código antes del fallback '?',
    para que recursos genuinamente desconocidos no contaminen cálculos de coste.
    """
    if codigo.startswith('%'):
        return '%'
    if codigo in presupuesto.recursos_mo:
        return 'MO'
    if codigo in presupuesto.recursos_mt:
        return 'MT'
    if codigo in presupuesto.recursos_mq:
        return 'MQ'
    if codigo in presupuesto.recursos_pa:
        return 'PA'
    return '?'


def iter_todas_partidas(presupuesto: Presupuesto):
    """
    Genera tuplas (capitulo_raiz, subcapitulo_o_none, partida) en orden.
    Recorre recursivamente todos los capítulos y subcapítulos.
    """
    def recurrir(raiz, nodo, sub_actual):
        for partida in nodo.partidas:
            yield raiz, sub_actual, partida
        for sub in nodo.subcapitulos:
            yield from recurrir(raiz, sub, sub)

    for cap in presupuesto.capitulos:
        yield from recurrir(cap, cap, None)


def es_partida_liquidable(codigo: str, presupuesto: Presupuesto) -> bool:
    """
    Retorna True si la partida tiene mano de obra Y material simultáneamente,
    expandiendo PAs recursivamente para detectar recursos en descompuestos anidados.
    Criterio para incluir en la base de cálculo de la liquidación máxima.
    """
    recursos = recursos_expandidos(codigo, presupuesto)
    tiene_mo = any(tipo_recurso(r, presupuesto) == 'MO' for r in recursos)
    tiene_mt = any(tipo_recurso(r, presupuesto) == 'MT' for r in recursos)
    return tiene_mo and tiene_mt




def normalizar_texto_largo(texto: str | None) -> str:
    """
    Normaliza descripciones largas para que sean legibles en Excel.

    El parser BC3 puede entregar textos con párrafos compactados. Esta función no
    cambia el contenido técnico, pero recupera estructura visual básica: apartados,
    listas e hitos contractuales habituales de las descripciones de partidas.
    """
    if not texto:
        return ''

    limpio = str(texto).replace('\r\n', '\n').replace('\r', '\n')
    limpio = limpio.replace('|', '')
    limpio = re.sub(r'[ \t]+', ' ', limpio)
    limpio = re.sub(r' *\n *', '\n', limpio)

    patrones_apartado = [
        'Incluye:',
        'Documentación mínima a entregar:',
        'Criterio de medición en proyecto:',
        'Criterio de medida en proyecto:',
        'Criterio de medición en obra:',
        'Criterio de medida en obra:',
    ]
    for patron in patrones_apartado:
        limpio = limpio.replace(f' {patron}', f'\n\n{patron}')

    # Convierte listados compactados " - Elemento" en líneas independientes.
    limpio = re.sub(r'\s+-\s+', '\n- ', limpio)
    limpio = re.sub(r'\n{3,}', '\n\n', limpio)
    return limpio.strip()


def forzar_altura_fila(ws, fila: int, altura: float) -> None:
    """Aplica la mayor altura calculada a una fila sin reducir alturas previas."""
    actual = ws.row_dimensions[fila].height or 0
    ws.row_dimensions[fila].height = max(actual, altura)


_alturas_manuales: WeakKeyDictionary = WeakKeyDictionary()


def registrar_altura_manual(ws, fila: int) -> None:
    """Marca una fila para que el ajuste global de alturas no la recalcule."""
    if ws not in _alturas_manuales:
        _alturas_manuales[ws] = set()
    _alturas_manuales[ws].add(fila)


def tiene_altura_manual(ws, fila: int) -> bool:
    """Indica si una fila tiene altura calculada y fijada explícitamente."""
    return fila in _alturas_manuales.get(ws, set())

# ── Helpers internos de escritura Excel ──────────────────────────────────────

def row_height(
    desc: str,
    width_chars: int = 55,
    line_height: float = 13.0,
    min_height: float = 18.0,
    max_height: float = 180.0,
) -> float:
    """Estima altura de fila según saltos reales y longitud de texto."""
    texto = str(desc or '').strip()
    if not texto:
        return min_height

    lineas = 0
    for bloque in texto.splitlines() or ['']:
        bloque = bloque.strip()
        if not bloque:
            lineas += 1
        else:
            lineas += max(1, (len(bloque) // max(width_chars, 1)) + 1)

    return max(min_height, min(max_height, lineas * line_height))


def cel_desc(
    ws,
    fila: int,
    col: int,
    texto: str,
    font=None,
    width_chars: int = 55,
    line_height: float = 13.0,
    min_height: float = 18.0,
    max_height: float = 180.0,
) -> None:
    """Escribe una celda de descripción con wrap_text y ajuste de altura."""
    cell = ws.cell(fila, col, texto)
    cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    cell.font = font or PRES_FN
    cell.border = PRES_BORDE_TABLA
    altura = row_height(
        texto,
        width_chars=width_chars,
        line_height=line_height,
        min_height=min_height,
        max_height=max_height,
    )
    forzar_altura_fila(ws, fila, altura)


def cel_eur(ws, fila: int, col: int, valor: float, font=None) -> None:
    """Escribe una celda numérica en euros con formato y alineación derecha."""
    cell              = ws.cell(fila, col, valor)
    cell.number_format = PRES_FE
    cell.alignment    = PRES_AN
    cell.font         = font or PRES_FN
    cell.border       = PRES_BORDE_TABLA


def cel_cantidad(ws, fila: int, col: int, valor: float, font=None, recurso: bool = False) -> None:
    """Escribe una cantidad con cero inicial y sin coma decimal sobrante."""
    cell = ws.cell(fila, col, valor)
    cell.number_format = PRES_FU if recurso else PRES_FQ
    cell.alignment = PRES_ANC
    cell.font = font or PRES_FN
    cell.border = PRES_BORDE_TABLA



# ── Ajuste global de alturas de fila ─────────────────────────────────────────
# Calibrado empíricamente para UIBSans 10 pt. Si se cambia la fuente o tamaño,
# el parámetro más sensible es PRES_ROW_FONT_WIDTH_PT.
PRES_ROW_FONT_WIDTH_PT = 5.8
PRES_ROW_LINE_HEIGHT_PT = 16.0
PRES_ROW_PADDING_PT = 12.0
PRES_ROW_MIN_HEIGHT_PT = 15.0
PRES_ROW_MAX_HEIGHT_PT = 409.0


def col_width_to_points(col_width_chars: float) -> float:
    """Convierte ancho de columna openpyxl aproximado a puntos."""
    return col_width_chars * 7.0


def _points_torow_height(points: float) -> float:
    """openpyxl espera la altura de fila directamente en puntos."""
    return points


def cell_has_text(value) -> bool:
    """Indica si una celda contiene texto o dato visible."""
    return value is not None and str(value).strip() != ''


def calcular_altura_fila_por_ancho_efectivo(row_cells, col_widths_points) -> float:
    """
    Calcula la altura necesaria de una fila.

    Para cada celda con contenido, se considera como ancho efectivo el ancho de
    su columna más las columnas vacías adyacentes a la derecha, hasta encontrar
    otra celda con contenido. Esto reproduce mejor cómo se percibe el texto largo
    en Excel cuando la fila contiene una descripción y varias columnas vacías.
    """
    max_lines = 1

    for col_index, cell_text in enumerate(row_cells):
        if not cell_has_text(cell_text):
            continue

        effective_width = col_widths_points[col_index]
        for next_col in range(col_index + 1, len(row_cells)):
            if cell_has_text(row_cells[next_col]):
                break
            effective_width += col_widths_points[next_col]

        chars_per_line = max(int(effective_width / PRES_ROW_FONT_WIDTH_PT), 1)

        lines_for_cell = 0
        for paragraph in str(cell_text).split('\n'):
            if paragraph.strip() == '':
                lines_for_cell += 1
            else:
                lines_for_cell += max(math.ceil(len(paragraph) / chars_per_line), 1)

        max_lines = max(max_lines, lines_for_cell)

    height_points = round(max_lines * PRES_ROW_LINE_HEIGHT_PT + PRES_ROW_PADDING_PT)
    return max(PRES_ROW_MIN_HEIGHT_PT, min(height_points, PRES_ROW_MAX_HEIGHT_PT))


def ajustar_alturas_filas_por_contenido(ws, max_col: int, min_row: int = 1) -> None:
    """
    Ajusta alturas de todas las filas no vacías según ancho efectivo disponible.

    No modifica filas completamente vacías, para respetar separadores con altura
    reducida definidos explícitamente por el exporter.
    """
    col_widths_points = []
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        width_chars = ws.column_dimensions[col_letter].width or 8
        col_widths_points.append(col_width_to_points(width_chars))

    for row_idx in range(min_row, ws.max_row + 1):
        if tiene_altura_manual(ws, row_idx):
            continue

        row_cells = [ws.cell(row_idx, col_idx).value for col_idx in range(1, max_col + 1)]
        if not any(cell_has_text(value) for value in row_cells):
            continue

        height_points = calcular_altura_fila_por_ancho_efectivo(row_cells, col_widths_points)
        calculated_height = _points_torow_height(height_points)

        # Nunca se reduce una altura ya asignada. Esto evita que el ajuste global
        # colapse filas de texto, subtotales o filas que han sido dimensionadas
        # durante la escritura del documento.
        current_height = ws.row_dimensions[row_idx].height or 0
        ws.row_dimensions[row_idx].height = max(current_height, calculated_height)



def ancho_rango_en_puntos(ws, start_col: int, end_col: int) -> float:
    """Devuelve el ancho aproximado en puntos de un rango de columnas."""
    total = 0.0
    for col_idx in range(start_col, end_col + 1):
        col_letter = get_column_letter(col_idx)
        width_chars = ws.column_dimensions[col_letter].width or 8
        total += col_width_to_points(width_chars)
    return total


def lineas_estimadas_texto(texto: str, chars_per_line: int) -> int:
    """Estima líneas visibles respetando saltos de línea explícitos."""
    total = 0
    for paragraph in str(texto or '').split('\n'):
        if paragraph.strip() == '':
            total += 1
        else:
            total += max(math.ceil(len(paragraph) / max(chars_per_line, 1)), 1)
    return max(total, 1)


def altura_por_lineas(lineas: int) -> float:
    """Calcula altura de fila en puntos para un número de líneas."""
    altura = round(lineas * PRES_ROW_LINE_HEIGHT_PT + PRES_ROW_PADDING_PT)
    return max(PRES_ROW_MIN_HEIGHT_PT, min(altura, PRES_ROW_MAX_HEIGHT_PT))


def dividir_texto_en_bloques_visibles(
    texto: str,
    width_points: float,
    max_height_points: float = 300.0,
) -> list[str]:
    """
    Divide textos largos en bloques seguros para Excel.

    La división se realiza ANTES de escribir en la hoja. Cada bloque contiene
    como máximo un número de líneas que cabe holgadamente en una fila de Excel,
    por debajo del límite físico de 409 puntos. Las líneas se parten por ancho
    efectivo y se conserva la estructura de párrafos y listas mediante saltos de
    línea explícitos.
    """
    texto = str(texto or '').strip()
    if not texto:
        return ['']

    chars_per_line = max(int(width_points / PRES_ROW_FONT_WIDTH_PT), 1)
    max_lines = max(int((max_height_points - PRES_ROW_PADDING_PT) / PRES_ROW_LINE_HEIGHT_PT), 1)

    lineas_fisicas: list[str] = []
    for paragraph in texto.split('\n'):
        if paragraph.strip() == '':
            lineas_fisicas.append('')
            continue

        partes = textwrap.wrap(
            paragraph,
            width=chars_per_line,
            break_long_words=False,
            replace_whitespace=False,
            drop_whitespace=True,
        )
        lineas_fisicas.extend(partes or [''])

    bloques: list[str] = []
    actual: list[str] = []
    for linea in lineas_fisicas:
        if len(actual) >= max_lines:
            bloques.append('\n'.join(actual).strip())
            actual = []
        actual.append(linea)

    if actual:
        bloques.append('\n'.join(actual).strip())

    return bloques or ['']


def border_continuacion_texto() -> Border:
    """Borde para filas de continuación: sin borde superior."""
    return Border(
        left=Side(style='thin', color=GRIS_LINEA),
        right=Side(style='thin', color=GRIS_LINEA),
        bottom=Side(style='thin', color=GRIS_LINEA),
    )


def aplicar_estilo_rango_texto(
    ws,
    fila: int,
    start_col: int,
    end_col: int,
    font=None,
    fill=None,
    continuation: bool = False,
) -> None:
    """Aplica estilo a una fila de descripción, incluyendo continuaciones."""
    borde = border_continuacion_texto() if continuation else PRES_BORDE_TABLA
    for col in range(start_col, end_col + 1):
        cell = ws.cell(fila, col)
        cell.font = font or PRES_FN
        cell.fill = fill or PRES_FB
        cell.border = borde
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)


def append_descripcion_fusionada_partida(
    ws,
    n_cols: int,
    start_col: int,
    end_col: int,
    texto: str,
    font=None,
    fill=None,
    max_height_points: float = 300.0,
) -> list[int]:
    """
    Añade una descripción larga en una o varias filas fusionadas.

    Si el texto no cabe en una fila por el límite de Excel, se crean filas de
    continuación debajo, sin borde superior, para que visualmente pertenezcan a
    la misma descripción.
    """
    width_points = ancho_rango_en_puntos(ws, start_col, end_col)
    chunks = dividir_texto_en_bloques_visibles(texto, width_points, max_height_points)
    filas: list[int] = []

    for idx, chunk in enumerate(chunks):
        ws.append([''] * n_cols)
        fila = ws.max_row
        filas.append(fila)
        ws.cell(fila, start_col, chunk)
        if end_col > start_col:
            ws.merge_cells(start_row=fila, start_column=start_col, end_row=fila, end_column=end_col)
        aplicar_estilo_rango_texto(
            ws,
            fila,
            start_col,
            end_col,
            font=font or PRES_FN,
            fill=fill or PRES_FB,
            continuation=idx > 0,
        )
        lineas = lineas_estimadas_texto(chunk, max(int(width_points / PRES_ROW_FONT_WIDTH_PT), 1))
        ws.row_dimensions[fila].height = altura_por_lineas(lineas)
        registrar_altura_manual(ws, fila)

    return filas


def append_fila_cuadro_precios_1(
    ws,
    codigo: str,
    texto: str,
    precio: float,
) -> None:
    """Añade una partida de PRES.02.01, dividiendo la descripción si es necesario."""
    width_points = ancho_rango_en_puntos(ws, 2, 2)
    chunks = dividir_texto_en_bloques_visibles(texto, width_points, max_height_points=300.0)
    chars_per_line = max(int(width_points / PRES_ROW_FONT_WIDTH_PT), 1)

    for idx, chunk in enumerate(chunks):
        if idx == 0:
            ws.append([codigo, chunk, precio])
        else:
            ws.append(['', chunk, ''])
        fila = ws.max_row
        for col in range(1, 4):
            cell = ws.cell(fila, col)
            cell.fill = PRES_FB
            cell.font = PRES_FN
            cell.border = border_continuacion_texto() if idx > 0 else PRES_BORDE_TABLA
        ws.cell(fila, 1).alignment = PRES_AT
        ws.cell(fila, 2).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        if idx == 0:
            cel_eur(ws, fila, 3, precio)
        else:
            ws.cell(fila, 3).alignment = PRES_AN
        lineas = lineas_estimadas_texto(chunk, chars_per_line)
        ws.row_dimensions[fila].height = altura_por_lineas(lineas)
        registrar_altura_manual(ws, fila)

def cabecera_tabla(ws, headers: list[str]) -> None:
    """Escribe una fila de cabecera de tabla con estilo de cabecera."""
    ws.append(headers)
    r = ws.max_row
    for c, _ in enumerate(headers, 1):
        ws.cell(r, c).fill      = PRES_FCP
        ws.cell(r, c).font      = PRES_FCF
        ws.cell(r, c).alignment = PRES_AC
        ws.cell(r, c).border    = PRES_BORDE_TABLA
    ws.row_dimensions[r].height = 22.0
    registrar_altura_manual(ws, r)


def sep_vacia(ws, altura: float = 4.0, n_cols: int = 1) -> None:
    """Inserta una fila separadora real sin modificar la fila anterior.

    openpyxl no siempre incrementa ws.max_row cuando se usa ws.append([]).
    Por eso se crean celdas vacías explícitas en una nueva fila. La fila se
    marca como manual para que el ajuste global de alturas no la recalcule.
    """
    fila = ws.max_row + 1
    for col in range(1, n_cols + 1):
        cell = ws.cell(fila, col, '')
        cell.fill = PRES_FB
        cell.border = Border()
    ws.row_dimensions[fila].height = altura
    registrar_altura_manual(ws, fila)


def capitulo_limpio(codigo: str) -> str:
    """Devuelve el código de capítulo sin el sufijo '#'."""
    return codigo.rstrip('#')


def get_cap_gr(presupuesto: Presupuesto) -> 'Capitulo | None':
    """Localiza el capítulo de gestión de residuos por código."""
    return next((c for c in presupuesto.capitulos if c.codigo == presupuesto.config.codigo_capitulo_gr), None)


def fila_separador_cascade(ws, n_cols: int) -> None:
    """Inserta una fila separadora con fondo azul muy claro entre bloques de cascada."""
    ws.append([''] * n_cols)
    r = ws.max_row
    for c in range(1, n_cols + 1):
        ws.cell(r, c).fill = PRES_FSC
        ws.cell(r, c).border = PRES_BORDE_FINO
    ws.row_dimensions[r].height = 6.0
    registrar_altura_manual(ws, r)


def escribir_cascade(ws, filas: list[tuple], n_cols: int = 3) -> None:
    """
    Escribe una tabla en cascada de 3 columnas: (etiqueta, pct_str, importe).
    None como etiqueta inserta un separador visual.
    True como etiqueta aplica estilo de total.
    """
    for item in filas:
        if item is None:
            fila_separador_cascade(ws, n_cols)
            continue
        label, pct, valor, es_total = item
        ws.append([label, pct, ''] if n_cols >= 3 else [label, pct])
        r = ws.max_row
        ws.cell(r, 1).font = PRES_FTO if es_total else PRES_FN
        ws.cell(r, 1).fill = PRES_FB
        ws.cell(r, 2).font = PRES_FTO if es_total else PRES_FN
        ws.cell(r, 2).fill = PRES_FB
        ws.cell(r, 2).alignment = PRES_AC
        cel_eur(ws, r, 3, valor, font=PRES_FTO if es_total else PRES_FN)
        ws.cell(r, 3).fill = PRES_FB
        for c in range(1, n_cols + 1):
            ws.cell(r, c).border = PRES_BORDE_TOTAL if es_total else PRES_BORDE_FINO
        ws.row_dimensions[r].height = 24.0 if es_total else 21.0
        registrar_altura_manual(ws, r)


def descripcion_completa(partida: Partida) -> str:
    """Retorna la descripción larga (~T) normalizada para lectura documental."""
    return normalizar_texto_largo(partida.descripcion_larga or partida.descripcion)


def fmt_numero_texto(valor: float, decimales: int = 3) -> str:
    """Formatea un número para textos visibles con separadores españoles."""
    texto = f'{valor:,.{decimales}f}'.rstrip('0').rstrip('.')
    return texto.translate(str.maketrans(',.', '.,'))


# ── PRES.01 — Cuadro de Oferta ───────────────────────────────────────────────

def generar_pres01(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Cuadro de Oferta con columnas amarillas para el licitador."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Cuadro de Oferta'
    N = 8
    ajustar_columnas(ws, {'A': 16, 'B': 55, 'C': 6, 'D': 12,
                          'E': 16, 'F': 16, 'G': 18, 'H': 18})

    aplicar_encabezado_documental(ws, presupuesto.config, 'PRES.01', 'CUADRO DE OFERTA', N)

    ws.append(['Rellene las columnas en amarillo con su precio unitario ofertado']
              + [''] * (N - 1))
    r2 = ws.max_row
    ws.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=N)
    ws.cell(r2, 1).fill      = PRES_FA
    ws.cell(r2, 1).font      = Font(name=FONT_PRES, size=10, italic=True, color='000000')
    ws.cell(r2, 1).alignment = PRES_AC
    ws.row_dimensions[r2].height = 16

    cabecera_tabla(ws, [
        'Código', 'Descripción', 'Ud.', 'Cantidad',
        'P. Unit. ref. (€)', 'Importe ref. (€)',
        'P. Unit. ofertado (€)', 'Importe ofertado (€)',
    ])
    fila_cabecera = ws.max_row
    ws.freeze_panes = 'A8'

    current_cap = None
    current_sub = None
    alt = True

    for cap, sub, partida in iter_todas_partidas(presupuesto):
        if cap is not current_cap:
            current_cap = cap
            current_sub = None
            ws.append([capitulo_limpio(cap.codigo), cap.descripcion]
                      + [''] * (N - 2))
            r = ws.max_row
            aplicar_estilo_capitulo(ws, r, N)
            cel_desc(ws, r, 2, cap.descripcion, font=PRES_FCF, width_chars=55)

        if sub is not current_sub:
            current_sub = sub
            if sub is not None:
                ws.append([capitulo_limpio(sub.codigo), '  ' + sub.descripcion]
                          + [''] * (N - 2))
                r = ws.max_row
                for c in range(1, N + 1):
                    ws.cell(r, c).fill = PRES_FSC
                    ws.cell(r, c).font = PRES_FCF
                cel_desc(ws, r, 2, '  ' + sub.descripcion, font=PRES_FCF, width_chars=55)
                ws.cell(r, 2).fill = PRES_FSC

        importe = partida.precio_unitario * partida.cantidad
        fill = PRES_FD if alt else PRES_FB
        alt  = not alt

        ws.append([partida.codigo, partida.descripcion, partida.unidad,
                   partida.cantidad, partida.precio_unitario, importe, '', ''])
        r = ws.max_row
        for c in range(1, N + 1):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = PRES_FN
            ws.cell(r, c).border = PRES_BORDE_TABLA
        cel_desc(ws, r, 2, partida.descripcion, width_chars=55)
        ws.cell(r, 2).fill = fill
        cel_cantidad(ws, r, 4, partida.cantidad)
        ws.cell(r, 4).fill = fill
        ws.cell(r, 5).number_format = PRES_FE
        ws.cell(r, 5).alignment = PRES_AN
        ws.cell(r, 6).number_format = PRES_FE
        ws.cell(r, 6).alignment = PRES_AN
        ws.cell(r, 7).fill = PRES_FA   # precio ofertado
        ws.cell(r, 8).fill = PRES_FA   # importe ofertado

    # Fila TOTAL PEM
    ws.append(['', 'TOTAL PRESUPUESTO DE EJECUCIÓN MATERIAL'] + [''] * 3
              + [presupuesto.importe_total, '', ''])
    r = ws.max_row
    aplicar_estilo_total(ws, r, N)
    cel_eur(ws, r, 6, presupuesto.importe_total, font=PRES_FTO)
    ws.cell(r, 6).fill = PRES_FB

    ajustar_columnas(ws, {'A': 16, 'B': 55, 'C': 6, 'D': 12,
                          'E': 16, 'F': 16, 'G': 18, 'H': 18})
    ajustar_alturas_filas_por_contenido(ws, max_col=N)
    configurar_impresion(ws, 'landscape', fila_cabecera, presupuesto.config)
    wb.save(filepath)


# ── PRES.02.01 — Cuadro de Precios n.º 1 ─────────────────────────────────────

def generar_pres0201(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Cuadro de Precios n.º 1 con el precio unitario en letra."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Cuadro de Precios 1'
    N = 3
    ajustar_columnas(ws, {'A': 16, 'B': 72, 'C': 18})

    aplicar_encabezado_documental(ws, presupuesto.config, 'PRES.02.01', 'CUADRO DE PRECIOS N.º 1', N)

    cabecera_tabla(ws, ['Código', 'Descripción del precio', 'Precio unit. (€)'])
    fila_cabecera = ws.max_row
    ws.freeze_panes = 'A7'
    ajustar_columnas(ws, {'A': 16, 'B': 72, 'C': 18})

    # Recopilar partidas con sus capítulos
    current_cap = None
    for cap, sub, partida in iter_todas_partidas(presupuesto):
        if cap is not current_cap:
            current_cap = cap
            ws.append([capitulo_limpio(cap.codigo), cap.descripcion, ''])
            r = ws.max_row
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N)
            aplicar_estilo_capitulo(ws, r, N)
            cel_desc(ws, r, 1,
                      f'CAPÍTULO {capitulo_limpio(cap.codigo)} — {cap.descripcion}',
                      font=PRES_FCF)

        desc_precio  = descripcion_completa(partida)
        precio_letra = numero_a_letras(partida.precio_unitario)
        contenido_b  = f'{desc_precio}\n\nPRECIO: {precio_letra}'
        append_fila_cuadro_precios_1(
            ws,
            partida.codigo,
            contenido_b,
            partida.precio_unitario,
        )
        sep_vacia(ws, 4.0, N)

    ajustar_columnas(ws, {'A': 16, 'B': 72, 'C': 18})
    ajustar_alturas_filas_por_contenido(ws, max_col=N)
    configurar_impresion(ws, 'portrait', fila_cabecera, presupuesto.config)
    wb.save(filepath)


# ── PRES.02.02 — Cuadro de Precios n.º 2 ─────────────────────────────────────

def generar_pres0202(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Cuadro de Precios n.º 2 con descomposición de cada partida."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Cuadro de Precios 2'
    N = 7
    ajustar_columnas(ws, {'A': 8, 'B': 16, 'C': 55, 'D': 8, 'E': 12, 'F': 16, 'G': 16})

    aplicar_encabezado_documental(ws, presupuesto.config, 'PRES.02.02', 'CUADRO DE PRECIOS N.º 2', N)

    cabecera_tabla(ws, ['Tipo', 'Código', 'Descripción',
                         'Ud.', 'Cant./ud', 'P. unit./Base (€)', 'Coste (€)'])
    fila_cabecera = ws.max_row
    ws.freeze_panes = 'A7'
    ajustar_columnas(ws, {'A': 8, 'B': 16, 'C': 55, 'D': 8, 'E': 12, 'F': 16, 'G': 16})

    for cap, sub, partida in iter_todas_partidas(presupuesto):
        # Fila resumen de partida
        ws.append([partida.codigo, partida.descripcion, partida.unidad,
                   '', '', '', ''])
        r = ws.max_row
        aplicar_estilo_capitulo(ws, r, N)
        cel_desc(ws, r, 2, partida.descripcion, font=PRES_FCF, width_chars=38, max_height=80.0)
        ws.cell(r, 2).fill = PRES_FCP
        for c in (1, 3, 4, 5, 6, 7):
            ws.cell(r, c).alignment = PRES_AC
        ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 20, 22)
        registrar_altura_manual(ws, r)

        # Fila descripción técnica completa. Si supera el límite de altura
        # de Excel, se parte en filas de continuación sin borde superior.
        desc_larga = descripcion_completa(partida)
        append_descripcion_fusionada_partida(
            ws,
            n_cols=N,
            start_col=2,
            end_col=N,
            texto=desc_larga,
            font=PRES_FN,
            fill=PRES_FB,
            max_height_points=300.0,
        )

        recursos = recursos_para_display(partida.codigo, presupuesto)

        if not recursos:
            # Partida de precio alzado
            ws.append(['', 'Precio alzado', '', '', '', '', partida.precio_unitario])
            r = ws.max_row
            for c in range(1, N + 1):
                ws.cell(r, c).fill = PRES_FD
                ws.cell(r, c).font = PRES_FDE
                ws.cell(r, c).border = PRES_BORDE_TABLA
            cel_eur(ws, r, 7, partida.precio_unitario, font=PRES_FDE)
            ws.cell(r, 7).fill = PRES_FD
        else:
            # Agrupar por tipo sin expandir auxiliares (PA)
            grupos = {'MO': [], 'MQ': [], 'MT': [], 'PA': [], '%': [], '?': []}
            for cod_rec, cant in recursos.items():
                grupos[tipo_recurso(cod_rec, presupuesto)].append((cod_rec, cant))

            suma_directos = 0.0
            for tipo, label in [('MO', 'Mano de obra'), ('MQ', 'Maquinaria'),
                                ('MT', 'Materiales'), ('PA', 'Unidades auxiliares'),
                                ('?', 'Sin clasificar')]:
                if not grupos[tipo]:
                    continue
                # Subtítulo del grupo
                ws.append(['', label, '', '', '', '', ''])
                r = ws.max_row
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
                for c in range(1, N + 1):
                    ws.cell(r, c).fill = PRES_FD
                    ws.cell(r, c).font = Font(name=FONT_PRES, size=9,
                                              bold=True, color=GRIS_TEXTO)
                    ws.cell(r, c).border = PRES_BORDE_TABLA

                for cod_rec, cant in grupos[tipo]:
                    desc, unidad, precio = get_recurso_info(cod_rec, presupuesto)
                    coste = round(cant * precio, 4)
                    suma_directos += coste
                    fill = PRES_FB
                    col_tipo = 'AUX' if tipo == 'PA' else tipo
                    ws.append([col_tipo, cod_rec, desc, unidad, cant, precio, coste])
                    r = ws.max_row
                    for c in range(1, N + 1):
                        ws.cell(r, c).fill = fill
                        ws.cell(r, c).font = PRES_FDE
                        ws.cell(r, c).border = PRES_BORDE_TABLA
                        ws.cell(r, c).alignment = PRES_AC
                    cel_desc(ws, r, 3, desc, font=PRES_FDE, width_chars=55, max_height=90.0)
                    ws.cell(r, 3).fill = fill
                    ws.cell(r, 1).alignment = PRES_ALC
                    ws.cell(r, 2).alignment = PRES_ALC
                    cel_cantidad(ws, r, 5, cant, font=PRES_FDE, recurso=True)
                    ws.cell(r, 5).fill = fill
                    ws.cell(r, 6).number_format = PRES_FE
                    ws.cell(r, 6).alignment = PRES_ANC
                    ws.cell(r, 7).number_format = PRES_FE
                    ws.cell(r, 7).alignment = PRES_ANC

            # Subtotal directos
            ws.append(['', '', '', '', '', 'Suma directos:', round(suma_directos, 2)])
            r = ws.max_row
            aplicar_estilo_subtotal(ws, r, N)
            ws.row_dimensions[r].height = 22.0
            ws.cell(r, 6).alignment = PRES_ANC
            cel_eur(ws, r, 7, round(suma_directos, 2), font=PRES_FCF)
            ws.cell(r, 7).fill = PRES_FB

            total_auxiliares = 0.0
            if grupos['%']:
                ws.append(['', 'Auxiliares / costes indirectos', '', '', '', '', ''])
                r = ws.max_row
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
                for c in range(1, N + 1):
                    ws.cell(r, c).fill = PRES_FB
                    ws.cell(r, c).font = Font(name=FONT_PRES, size=9,
                                              bold=True, color=GRIS_TEXTO)
                    ws.cell(r, c).border = PRES_BORDE_FINO

                base_aux = suma_directos  # base acumulada: cada % se aplica sobre directos + anteriores %
                for cod_rec, cant in grupos['%']:
                    desc, unidad, _precio = get_recurso_info(cod_rec, presupuesto)
                    porcentaje = cant / 100 if abs(cant) > 1 else cant
                    importe_auxiliar = round(base_aux * porcentaje, 4)
                    total_auxiliares += importe_auxiliar
                    unidad_aux = '%' if not unidad or unidad == '-' else unidad
                    ws.append(['%', cod_rec, desc, unidad_aux,
                               porcentaje, base_aux, importe_auxiliar])
                    base_aux += importe_auxiliar
                    r = ws.max_row
                    for c in range(1, N + 1):
                        ws.cell(r, c).fill = PRES_FB
                        ws.cell(r, c).font = PRES_FDE
                        ws.cell(r, c).border = PRES_BORDE_TABLA
                        ws.cell(r, c).alignment = PRES_AC
                    cel_desc(ws, r, 3, desc, font=PRES_FDE, width_chars=55, max_height=90.0)
                    ws.cell(r, 1).alignment = PRES_ALC
                    ws.cell(r, 2).alignment = PRES_ALC
                    ws.cell(r, 5).number_format = PRES_FP
                    ws.cell(r, 5).alignment = PRES_ANC
                    ws.cell(r, 6).number_format = PRES_FE
                    ws.cell(r, 6).alignment = PRES_ANC
                    ws.cell(r, 7).number_format = PRES_FE
                    ws.cell(r, 7).alignment = PRES_ANC

            if grupos['%']:
                ws.append(['', '', '', '', '', 'Total auxiliares:', round(total_auxiliares, 2)])
                r = ws.max_row
                aplicar_estilo_subtotal(ws, r, N)
                ws.row_dimensions[r].height = 22.0
                cel_eur(ws, r, 7, round(total_auxiliares, 2), font=PRES_FCF)
                ws.cell(r, 7).fill = PRES_FB

        # Fila total partida
        ws.append(['', '', '', '', '', 'Precio unitario total:', partida.precio_unitario])
        r = ws.max_row
        for c in range(1, N + 1):
            ws.cell(r, c).fill = PRES_FB
            ws.cell(r, c).font = PRES_FTO
            ws.cell(r, c).border = PRES_BORDE_TOTAL
            ws.cell(r, c).alignment = PRES_ANC
        ws.cell(r, 6).alignment = Alignment(horizontal='right', vertical='center')
        cel_eur(ws, r, 7, partida.precio_unitario, font=PRES_FTO)
        ws.cell(r, 7).fill = PRES_FB
        ws.row_dimensions[r].height = 24.0
        registrar_altura_manual(ws, r)

        sep_vacia(ws, 6.0, N)

    ajustar_columnas(ws, {'A': 8, 'B': 16, 'C': 55, 'D': 8, 'E': 12, 'F': 16, 'G': 16})
    ajustar_alturas_filas_por_contenido(ws, max_col=N)
    configurar_impresion(ws, 'portrait', fila_cabecera, presupuesto.config)
    wb.save(filepath)


# ── PRES.02.03 / PRES.03 — Presupuesto Descompuesto y Mediciones ─────────────

def generar_descompuesto_doc(
    presupuesto: Presupuesto,
    filepath: Path,
    mostrar_precios: bool,
) -> None:
    """
    Generador compartido para PRES.02.03 (con precios) y PRES.03 (mediciones ciegas).
    Con mostrar_precios=True: 6 columnas, precios, importes, horas MO visibles, PEM final.
    Con mostrar_precios=False: 4 columnas, sin precios, sin importes, horas MO ocultas.
    """
    N          = 6 if mostrar_precios else 4
    col_widths = {'A': 8, 'B': 55, 'C': 6, 'D': 12, 'E': 16, 'F': 16} if mostrar_precios \
                 else {'A': 8, 'B': 55, 'C': 6, 'D': 12}
    headers    = (['Código', 'Descripción', 'Ud.', 'Cantidad', 'P. unit. (€)', 'Importe (€)']
                  if mostrar_precios else ['Código', 'Descripción', 'Ud.', 'Cantidad'])
    anexo      = 'PRES.02.03' if mostrar_precios else 'PRES.03'
    titulo     = 'PRESUPUESTO DESCOMPUESTO Y MEDICIONES' if mostrar_precios else 'MEDICIONES'

    wb = Workbook()
    ws = wb.active
    ws.title = 'Presupuesto Descompuesto' if mostrar_precios else 'Mediciones Ciegas'
    ajustar_columnas(ws, col_widths)
    aplicar_encabezado_documental(ws, presupuesto.config, anexo, titulo, N)
    cabecera_tabla(ws, headers)
    fila_cabecera = ws.max_row
    ws.freeze_panes = 'A7'
    ajustar_columnas(ws, col_widths)

    def escribir_partidas_de(nodo: Capitulo) -> None:
        for sub in nodo.subcapitulos:
            ws.append(['', f'{capitulo_limpio(sub.codigo)} — {sub.descripcion}']
                      + [''] * (N - 2))
            r_sub = ws.max_row
            ws.merge_cells(start_row=r_sub, start_column=1, end_row=r_sub, end_column=N)
            for c in range(1, N + 1):
                ws.cell(r_sub, c).fill = PRES_FSC
                ws.cell(r_sub, c).font = PRES_FCF
            cel_desc(ws, r_sub, 1,
                      f'{capitulo_limpio(sub.codigo)} — {sub.descripcion}', font=PRES_FCF)
            ws.cell(r_sub, 1).fill = PRES_FSC
            escribir_partidas_de(sub)

        for partida in nodo.partidas:
            importe = partida.precio_unitario * partida.cantidad
            fila_p  = [partida.codigo, partida.descripcion, partida.unidad, partida.cantidad]
            if mostrar_precios:
                fila_p += [partida.precio_unitario, importe]
            ws.append(fila_p)
            r_p = ws.max_row
            aplicar_estilo_capitulo(ws, r_p, N)
            cel_desc(ws, r_p, 2, partida.descripcion, font=PRES_FCF, width_chars=42, max_height=90.0)
            ws.cell(r_p, 2).fill = PRES_FCP
            cel_cantidad(ws, r_p, 4, partida.cantidad, font=PRES_FCF)
            ws.cell(r_p, 4).fill = PRES_FCP
            if mostrar_precios:
                cel_eur(ws, r_p, 5, partida.precio_unitario, font=PRES_FCF)
                ws.cell(r_p, 5).fill = PRES_FCP
                cel_eur(ws, r_p, 6, importe, font=PRES_FCF)
                ws.cell(r_p, 6).fill = PRES_FCP

            append_descripcion_fusionada_partida(
                ws, n_cols=N, start_col=2, end_col=N,
                texto=descripcion_completa(partida),
                font=PRES_FN, fill=PRES_FB, max_height_points=300.0,
            )

            recursos = recursos_para_display(partida.codigo, presupuesto)
            if recursos:
                grupos = {'MO': [], 'MQ': [], 'MT': [], 'PA': [], '%': [], '?': []}
                for cod_rec, cant in recursos.items():
                    grupos[tipo_recurso(cod_rec, presupuesto)].append((cod_rec, cant))

                suma_directos_local = 0.0
                if mostrar_precios:
                    for tipo_d in ['MO', 'MQ', 'MT', 'PA']:
                        for cod_tmp, cant_tmp in grupos[tipo_d]:
                            _, _, precio_tmp = get_recurso_info(cod_tmp, presupuesto)
                            suma_directos_local += round(cant_tmp * precio_tmp, 4)
                base_aux_local = suma_directos_local  # base acumulada para % en cascada

                for tipo, _ in [('MO', None), ('MQ', None), ('MT', None),
                                 ('PA', None), ('%', None), ('?', None)]:
                    if not grupos[tipo]:
                        continue
                    for cod_rec, cant in grupos[tipo]:
                        desc_r, unidad_r, precio_r = get_recurso_info(cod_rec, presupuesto)
                        if tipo == '%':
                            cant = cant / 100 if abs(cant) > 1 else cant
                            unidad_r = '%'
                            if mostrar_precios:
                                precio_r = base_aux_local

                        col_tipo = 'AUX' if tipo == 'PA' else tipo
                        if mostrar_precios:
                            coste = round(base_aux_local * cant, 4) if tipo == '%' \
                                    else round(cant * precio_r, 4)
                            ws.append([col_tipo, f'{cod_rec} — {desc_r}',
                                       unidad_r, cant, precio_r, coste])
                            if tipo == '%':
                                base_aux_local += coste
                        else:
                            cant_visible = '' if tipo == 'MO' else cant
                            ws.append([col_tipo, f'{cod_rec} — {desc_r}', unidad_r, cant_visible])

                        r_r = ws.max_row
                        for c in range(1, N + 1):
                            ws.cell(r_r, c).fill = PRES_FB
                            ws.cell(r_r, c).font = PRES_FDE
                            ws.cell(r_r, c).border = PRES_BORDE_TABLA
                        cel_desc(ws, r_r, 2, f'{cod_rec} — {desc_r}',
                                  font=PRES_FDE, width_chars=55, max_height=90.0)
                        ws.cell(r_r, 2).fill = PRES_FB

                        if mostrar_precios:
                            cel_cantidad(ws, r_r, 4, cant, font=PRES_FDE, recurso=True)
                            ws.cell(r_r, 4).fill = PRES_FB
                            ws.cell(r_r, 5).number_format = PRES_FE
                            ws.cell(r_r, 5).alignment    = PRES_AN
                            ws.cell(r_r, 6).number_format = PRES_FE
                            ws.cell(r_r, 6).alignment    = PRES_AN
                        else:
                            cant_visible = '' if tipo == 'MO' else cant
                            if cant_visible != '':
                                cel_cantidad(ws, r_r, 4, cant_visible, font=PRES_FDE,
                                              recurso=(tipo not in ('%', '?')))
                                ws.cell(r_r, 4).fill = PRES_FB

            sep_vacia(ws, 5.0, N)

    for cap in presupuesto.capitulos:
        ws.append([f'CAPÍTULO {capitulo_limpio(cap.codigo)} — {cap.descripcion}']
                  + [''] * (N - 1))
        r = ws.max_row
        aplicar_estilo_cabecera(ws, r, N)
        cel_desc(ws, r, 1,
                  f'CAPÍTULO {capitulo_limpio(cap.codigo)} — {cap.descripcion}', font=PRES_FT)
        ws.cell(r, 1).fill = PRES_FC
        ws.row_dimensions[r].height = 18
        registrar_altura_manual(ws, r)
        escribir_partidas_de(cap)

        # Total capítulo
        if mostrar_precios:
            ws.append(['', f'TOTAL CAPÍTULO {capitulo_limpio(cap.codigo)}: {cap.descripcion}',
                       '', '', '', cap.importe_total])
        else:
            ws.append([f'TOTAL CAPÍTULO {capitulo_limpio(cap.codigo)}: {cap.descripcion}']
                      + [''] * (N - 1))
        r_tc = ws.max_row
        for c in range(1, N + 1):
            ws.cell(r_tc, c).fill = PRES_FCP
            ws.cell(r_tc, c).font = PRES_FCF
            ws.cell(r_tc, c).border = PRES_BORDE_TABLA
        if mostrar_precios:
            ws.merge_cells(start_row=r_tc, start_column=1, end_row=r_tc, end_column=5)
            cel_eur(ws, r_tc, 6, cap.importe_total, font=PRES_FCF)
            ws.cell(r_tc, 6).fill = PRES_FCP
        else:
            ws.merge_cells(start_row=r_tc, start_column=1, end_row=r_tc, end_column=N)
        registrar_altura_manual(ws, r_tc)
        sep_vacia(ws, 6.0, N)

    if mostrar_precios:
        ws.append(['PRESUPUESTO DE EJECUCIÓN MATERIAL (PEM)'] + [''] * (N - 2)
                  + [presupuesto.importe_total])
        r = ws.max_row
        aplicar_estilo_total(ws, r, N)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N - 1)
        cel_eur(ws, r, N, presupuesto.importe_total, font=PRES_FTO)
        ws.cell(r, N).fill = PRES_FB

    ajustar_columnas(ws, col_widths)
    ajustar_alturas_filas_por_contenido(ws, max_col=N)
    configurar_impresion(ws, 'portrait', fila_cabecera, presupuesto.config)
    wb.save(filepath)


def generar_pres0203(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Presupuesto Descompuesto y Mediciones por capítulos."""
    generar_descompuesto_doc(presupuesto, filepath, mostrar_precios=True)


def generar_pres03_mediciones(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera las Mediciones ciegas: descripción y unidades sin precios ni horas MO."""
    generar_descompuesto_doc(presupuesto, filepath, mostrar_precios=False)


# ── PRES.02.04 — Resumen por Capítulos ───────────────────────────────────────

def generar_pres0204(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Resumen por Capítulos con cascada financiera PEM → PGL."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Resumen Capítulos'
    N3 = 3
    ajustar_columnas(ws, {'A': 52, 'B': 14, 'C': 18})

    cap_gr     = get_cap_gr(presupuesto)
    importe_gr = cap_gr.importe_total if cap_gr else 0.0
    pem_sin_gr = presupuesto.importe_total - importe_gr
    gg  = pem_sin_gr * presupuesto.config.porcentaje_gg
    bi  = pem_sin_gr * presupuesto.config.porcentaje_bi
    pec = pem_sin_gr + gg + bi
    iva_obra = pec * presupuesto.config.iva_obra
    iva_gr   = importe_gr * presupuesto.config.iva_gr
    pgl      = pec + importe_gr + iva_obra + iva_gr

    # ── Sección 1: tabla por capítulos ────────────────────────────────────────
    aplicar_encabezado_documental(ws, presupuesto.config, 'PRES.02.04', 'RESUMEN POR CAPÍTULOS', N3)

    cabecera_tabla(ws, ['Capítulo', 'Descripción', 'Importe (€)'])
    fila_cabecera = ws.max_row
    ws.freeze_panes = 'A7'

    for i, cap in enumerate(presupuesto.capitulos):
        ws.append([capitulo_limpio(cap.codigo), cap.descripcion, cap.importe_total])
        r = ws.max_row
        fill = PRES_FD if i % 2 == 0 else PRES_FB
        for c in range(1, N3 + 1):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = PRES_FN
            ws.cell(r, c).border = PRES_BORDE_TABLA
        cel_desc(ws, r, 2, cap.descripcion)
        ws.cell(r, 2).fill = fill
        cel_eur(ws, r, 3, cap.importe_total)
        ws.cell(r, 3).fill = fill
        ws.row_dimensions[r].height = 21.0

    ws.append(['', 'TOTAL PRESUPUESTO DE EJECUCIÓN MATERIAL (PEM)',
               presupuesto.importe_total])
    r = ws.max_row
    aplicar_estilo_total(ws, r, N3)
    cel_eur(ws, r, 3, presupuesto.importe_total, font=PRES_FTO)
    ws.cell(r, 3).fill = PRES_FB

    sep_vacia(ws, 6.0, N3)
    sep_vacia(ws, 6.0, N3)

    # ── Sección 2: cascada general ────────────────────────────────────────────
    ws.append(['RESUMEN GENERAL DEL PRESUPUESTO'] + [''] * (N3 - 1))
    r = ws.max_row
    aplicar_estilo_cabecera(ws, r, N3)
    ws.row_dimensions[r].height = 22
    registrar_altura_manual(ws, r)

    cabecera_tabla(ws, ['Concepto', '% aplicado', 'Importe (€)'])

    filas = [
        ('PEM (sin residuos)',        '—',   pem_sin_gr, False),
        (f'Gastos Generales',         f'{presupuesto.config.porcentaje_gg*100:.0f}%', gg,  False),
        (f'Beneficio Industrial',     f'{presupuesto.config.porcentaje_bi*100:.0f}%', bi,  False),
        None,
        ('PEC',                       '—',   pec,        True),
        ('Gestión de Residuos (GR)',  '—',   importe_gr, False),
        None,
        (f'IVA obra',                 f'{presupuesto.config.iva_obra*100:.0f}%', iva_obra, False),
        (f'IVA residuos',             f'{presupuesto.config.iva_gr*100:.0f}%',   iva_gr,   False),
        None,
        ('PGL (Presupuesto Global de Licitación)', '—', pgl, True),
    ]
    escribir_cascade(ws, filas, N3)

    ajustar_columnas(ws, {'A': 52, 'B': 14, 'C': 18})
    ajustar_alturas_filas_por_contenido(ws, max_col=N3)
    configurar_impresion(ws, 'portrait', fila_cabecera, presupuesto.config)
    wb.save(filepath)


# ── PRES.05 — VEC y Liquidación ──────────────────────────────────────────────

def generar_pres05(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el VEC y Liquidación — Bloque 1 (VEC) primero, Bloque 2 después."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'VEC y Liquidación'
    N3 = 3
    ajustar_columnas(ws, {'A': 55, 'B': 14, 'C': 18})

    cap_gr     = get_cap_gr(presupuesto)
    importe_gr = cap_gr.importe_total if cap_gr else 0.0
    pem_sin_gr = presupuesto.importe_total - importe_gr
    gg  = pem_sin_gr * presupuesto.config.porcentaje_gg
    bi  = pem_sin_gr * presupuesto.config.porcentaje_bi
    pec = pem_sin_gr + gg + bi

    # ── Calcular base de liquidación (Bloque 2) ───────────────────────────────
    pem_liq = sum(
        p.precio_unitario * p.cantidad
        for cap, sub, p in iter_todas_partidas(presupuesto)
        if es_partida_liquidable(p.codigo, presupuesto)
    )
    gg_liq   = pem_liq * presupuesto.config.porcentaje_gg
    bi_liq   = pem_liq * presupuesto.config.porcentaje_bi
    pec_liq  = pem_liq + gg_liq + bi_liq
    liq_max  = (pec_liq + importe_gr) * presupuesto.config.porcentaje_liquidacion

    vec      = pec + importe_gr + liq_max
    iva_obra = pec * presupuesto.config.iva_obra
    iva_gr   = importe_gr * presupuesto.config.iva_gr
    pgl      = pec + importe_gr + iva_obra + iva_gr

    # ── BLOQUE 1: VEC ─────────────────────────────────────────────────────────
    aplicar_encabezado_documental(
        ws,
        presupuesto.config,
        'PRES.05',
        'VALOR ESTIMADO DEL CONTRATO (VEC) Y LIQUIDACIÓN',
        N3,
    )

    cabecera_tabla(ws, ['Concepto', '% aplicado', 'Importe (€)'])
    fila_cabecera = ws.max_row
    ws.freeze_panes = 'A7'

    filas_b1 = [
        ('PEM (sin residuos)',                       '—',  pem_sin_gr, False),
        (f'Gastos Generales',                        f'{presupuesto.config.porcentaje_gg*100:.0f}%',   gg,      False),
        (f'Beneficio Industrial',                    f'{presupuesto.config.porcentaje_bi*100:.0f}%',   bi,      False),
        None,
        ('PEC',                                      '—',  pec,        True),
        ('Gestión de Residuos (GR)',                 '—',  importe_gr, False),
        (f'Liquidación máxima ({presupuesto.config.porcentaje_liquidacion*100:.0f}%)', '—', liq_max, False),
        None,
        ('VEC (sin IVA)',                            '—',  vec,        True),
        (f'IVA obra ({presupuesto.config.iva_obra*100:.0f}%)',  f'{presupuesto.config.iva_obra*100:.0f}%', iva_obra, False),
        (f'IVA residuos ({presupuesto.config.iva_gr*100:.0f}%)', f'{presupuesto.config.iva_gr*100:.0f}%', iva_gr, False),
        None,
        ('PGL (Presupuesto Global de Licitación)',   '—',  pgl,        True),
    ]
    escribir_cascade(ws, filas_b1, N3)

    sep_vacia(ws, 6.0, N3)
    sep_vacia(ws, 6.0, N3)
    sep_vacia(ws, 6.0, N3)

    # ── BLOQUE 2: Base de liquidación ─────────────────────────────────────────
    ws.append(['BLOQUE 2 — BASE DE CÁLCULO DE LA LIQUIDACIÓN MÁXIMA'] + [''] * (N3 - 1))
    r = ws.max_row
    aplicar_estilo_cabecera(ws, r, N3)
    ws.row_dimensions[r].height = 22

    nota = ('Base de liquidación: partidas con mano de obra y material simultáneos. '
            'Se excluyen partidas de auditoría, documentación, puesta en servicio, '
            'seguridad y salud, y gestión de residuos.')
    ws.append([nota] + [''] * (N3 - 1))
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N3)
    ws.cell(r, 1).fill      = PRES_FA
    ws.cell(r, 1).font      = Font(name=FONT_PRES, size=9, italic=True, color='000000')
    ws.cell(r, 1).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.cell(r, 1).border    = PRES_BORDE_TABLA
    ws.row_dimensions[r].height = 40
    registrar_altura_manual(ws, r)
    sep_vacia(ws, 6.0, N3)

    cabecera_tabla(ws, ['Código', 'Descripción', 'Importe (€)'])

    alt = True
    for cap, sub, partida in iter_todas_partidas(presupuesto):
        if not es_partida_liquidable(partida.codigo, presupuesto):
            continue
        importe_p = partida.precio_unitario * partida.cantidad
        fill = PRES_FD if alt else PRES_FB
        alt  = not alt
        ws.append([partida.codigo, partida.descripcion, importe_p])
        r = ws.max_row
        for c in range(1, N3 + 1):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = PRES_FN
            ws.cell(r, c).border = PRES_BORDE_TABLA
        cel_desc(ws, r, 2, partida.descripcion, width_chars=55)
        ws.cell(r, 2).fill = fill
        cel_eur(ws, r, 3, importe_p)
        ws.cell(r, 3).fill = fill

    ws.append(['', 'TOTAL BASE IMPONIBLE LIQUIDACIÓN', pem_liq])
    r = ws.max_row
    for c in range(1, N3 + 1):
        ws.cell(r, c).fill = PRES_FCP
        ws.cell(r, c).font = PRES_FCF
        ws.cell(r, c).border = PRES_BORDE_TABLA
    cel_eur(ws, r, 3, pem_liq, font=PRES_FCF)
    ws.cell(r, 3).fill = PRES_FCP
    registrar_altura_manual(ws, r)
    sep_vacia(ws, 6.0, N3)

    filas_b2 = [
        ('PEM liquidación',          '—',   pem_liq,    False),
        (f'Gastos Generales',        f'{presupuesto.config.porcentaje_gg*100:.0f}%',  gg_liq,  False),
        (f'Beneficio Industrial',    f'{presupuesto.config.porcentaje_bi*100:.0f}%',  bi_liq,  False),
        None,
        ('PEC liquidación',          '—',   pec_liq,    True),
        ('Gestión de Residuos (GR)', '—',   importe_gr, False),
        None,
        (f'LIQUIDACIÓN MÁXIMA ({presupuesto.config.porcentaje_liquidacion*100:.0f}%)', '—', liq_max, True),
    ]
    escribir_cascade(ws, filas_b2, N3)

    ajustar_columnas(ws, {'A': 55, 'B': 14, 'C': 18})
    ajustar_alturas_filas_por_contenido(ws, max_col=N3)
    configurar_impresion(ws, 'portrait', fila_cabecera, presupuesto.config)
    wb.save(filepath)


# ── JUST_PRECIOS — Justificación de Precios ───────────────────────────────────

def generar_just_precios(presupuesto: Presupuesto, filepath: Path) -> None:
    """
    Genera la Justificación de Precios con cantidades totales de cada recurso.
    """
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Justificación de Precios'
    N = 6
    ajustar_columnas(ws, {'A': 18, 'B': 55, 'C': 8, 'D': 14, 'E': 18, 'F': 18})

    aplicar_encabezado_documental(
        ws,
        presupuesto.config,
        'JUST_PRECIOS',
        'JUSTIFICACIÓN DE PRECIOS — RECURSOS EMPLEADOS EN EL PROYECTO',
        N,
    )
    ws.freeze_panes = 'A7'

    # Acumular cantidades totales por recurso
    mo_tot: dict[str, float] = {}
    mt_tot: dict[str, float] = {}
    mq_tot: dict[str, float] = {}

    for cap, sub, partida in iter_todas_partidas(presupuesto):
        for cod_rec, cant_ud in presupuesto.descompuestos_raw.get(partida.codigo, []):
            total = cant_ud * partida.cantidad
            tr = tipo_recurso(cod_rec, presupuesto)
            if tr == 'MO':
                mo_tot[cod_rec] = mo_tot.get(cod_rec, 0) + total
            elif tr == 'MT':
                mt_tot[cod_rec] = mt_tot.get(cod_rec, 0) + total
            elif tr == 'MQ':
                mq_tot[cod_rec] = mq_tot.get(cod_rec, 0) + total

    def escribir_seccion(titulo: str, recursos_map, totales: dict[str, float],
                         label_total: str) -> None:
        ws.append([titulo] + [''] * (N - 1))
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N)
        for c in range(1, N + 1):
            ws.cell(r, c).fill = PRES_FCP
            ws.cell(r, c).font = PRES_FCF
        ws.cell(r, 1).alignment = PRES_AT

        cabecera_tabla(ws, ['Código', 'Descripción', 'Ud.',
                             'P. unit. (€)', 'Cant. total proyecto', 'Coste total (€)'])

        coste_total_seccion = 0.0
        recursos_ordenados  = sorted(
            ((c, r) for c, r in recursos_map.items() if c in totales),
            key=lambda x: getattr(x[1], 'descripcion', x[0]),
        )
        for i, (cod, recurso) in enumerate(recursos_ordenados):
            cant     = totales.get(cod, 0)
            precio   = getattr(recurso, 'precio_hora', None) or getattr(recurso, 'precio_unidad', 0)
            unidad   = 'h' if hasattr(recurso, 'precio_hora') else recurso.unidad
            coste    = cant * precio
            coste_total_seccion += coste
            fill = PRES_FD if i % 2 == 0 else PRES_FB
            ws.append([cod, recurso.descripcion, unidad, precio, cant, coste])
            r = ws.max_row
            for c in range(1, N + 1):
                ws.cell(r, c).fill = fill
                ws.cell(r, c).font = PRES_FN
                ws.cell(r, c).border = PRES_BORDE_TABLA
            cel_desc(ws, r, 2, recurso.descripcion)
            ws.cell(r, 2).fill = fill
            ws.cell(r, 4).number_format = PRES_FE
            ws.cell(r, 4).alignment    = PRES_AN
            cel_cantidad(ws, r, 5, cant)
            ws.cell(r, 5).fill = fill
            cel_eur(ws, r, 6, coste)
            ws.cell(r, 6).fill = fill

        ws.append(['', f'TOTAL {label_total}', '', '', '', coste_total_seccion])
        r = ws.max_row
        aplicar_estilo_total(ws, r, N)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cel_eur(ws, r, 6, coste_total_seccion, font=PRES_FTO)
        ws.cell(r, 6).fill = PRES_FB
        sep_vacia(ws, 6.0, N)

    escribir_seccion('MANO DE OBRA', presupuesto.recursos_mo, mo_tot, 'MANO DE OBRA')
    escribir_seccion('MAQUINARIA',   presupuesto.recursos_mq, mq_tot, 'MAQUINARIA')
    escribir_seccion('MATERIALES',   presupuesto.recursos_mt, mt_tot, 'MATERIALES')

    ajustar_columnas(ws, {'A': 18, 'B': 55, 'C': 8, 'D': 14, 'E': 18, 'F': 18})
    ajustar_alturas_filas_por_contenido(ws, max_col=N)
    configurar_impresion(ws, 'portrait', 7, presupuesto.config)
    wb.save(filepath)


# ── Punto de entrada ─────────────────────────────────────────────────────────

def generar_todos(presupuesto: Presupuesto, output_dir: Path) -> list[Path]:
    """
    Genera los ocho documentos de presupuesto en output_dir.
    Retorna lista de rutas de los ficheros generados.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    documentos = [
        ('PRES.01_Cuadro_Oferta.xlsx',               generar_pres01),
        ('PRES.02.01_Cuadro_Precios_1.xlsx',          generar_pres0201),
        ('PRES.02.02_Cuadro_Precios_2.xlsx',          generar_pres0202),
        ('PRES.02.03_Presupuesto_Descompuesto.xlsx',  generar_pres0203),
        ('PRES.02.04_Resumen_Capitulos.xlsx',         generar_pres0204),
        ('PRES.03_Mediciones_Ciegas.xlsx',            generar_pres03_mediciones),
        ('PRES.05_VEC_Liquidacion.xlsx',              generar_pres05),
        ('JUST_PRECIOS_Recursos.xlsx',                generar_just_precios),
    ]
    rutas: list[Path] = []
    for nombre, funcion in documentos:
        ruta = output_dir / nombre
        funcion(presupuesto, ruta)
        rutas.append(ruta)
    return rutas

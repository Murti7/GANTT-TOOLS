"""
Genera los documentos de presupuesto para el pliego de contratación.

Siete ficheros Excel independientes: PRES.01, PRES.02.01, PRES.02.02,
PRES.02.03, PRES.02.04, PRES.05 y JUST_PRECIOS_Recursos.xlsx.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from gantt.bc3.models import Capitulo, Partida, Presupuesto

# ── Constantes financieras ────────────────────────────────────────────────────
PORCENTAJE_GG          = 0.13
PORCENTAJE_BI          = 0.06
PORCENTAJE_IVA_OBRA    = 0.21
PORCENTAJE_IVA_GR      = 0.10
PORCENTAJE_LIQUIDACION = 0.10
CODIGO_CAPITULO_GR     = '13#'

# ── Paleta de colores (específica de documentos de presupuesto) ───────────────
PRES_AZUL_OSCURO    = '1B3A5C'
PRES_AZUL_CLARO     = 'D9E1EC'
PRES_AZUL_MUY_CLARO = 'EEF2F7'
PRES_GRIS_CLARO     = 'F5F7FA'
PRES_BLANCO         = 'FFFFFF'
PRES_AMARILLO       = 'FFF3CD'
PRES_GRIS_TEXTO     = '444444'

# Fills internos
_FC  = PatternFill('solid', fgColor=PRES_AZUL_OSCURO)    # cabecera / total
_FCP = PatternFill('solid', fgColor=PRES_AZUL_CLARO)     # capítulo
_FSC = PatternFill('solid', fgColor=PRES_AZUL_MUY_CLARO) # subcapítulo
_FD  = PatternFill('solid', fgColor=PRES_GRIS_CLARO)     # detalle / alterno
_FB  = PatternFill('solid', fgColor=PRES_BLANCO)          # blanco
_FA  = PatternFill('solid', fgColor=PRES_AMARILLO)        # a rellenar

# Fonts internos
_FT  = Font(name='Calibri', size=11, bold=True,  color='FFFFFF')  # título sección
_FCF = Font(name='Calibri', size=10, bold=True,  color='000000')  # capítulo
_FN  = Font(name='Calibri', size=10,             color='000000')  # normal
_FDE = Font(name='Calibri', size=9,              color=PRES_GRIS_TEXTO)  # detalle
_FTO = Font(name='Calibri', size=10, bold=True,  color='FFFFFF')  # total

# Alineaciones internas
_AD  = Alignment(horizontal='left',   vertical='top',    wrap_text=True)
_AN  = Alignment(horizontal='right',  vertical='top')
_AC  = Alignment(horizontal='center', vertical='center')
_AT  = Alignment(horizontal='left',   vertical='top')

# Formatos numéricos internos
_FE  = '#,##0.00'   # euros
_FQ  = '#,##0.###'  # cantidades (hasta 3 decimales)


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


def _grupo_a_letras(n: int) -> str:
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
            euros_txt.append(_grupo_a_letras(miles) + ' MIL')
        if resto > 0:
            euros_txt.append(_grupo_a_letras(resto))
        partes.extend(euros_txt)
        partes.append('EURO' if euros == 1 else 'EUROS')

    if cents > 0:
        if euros > 0:
            partes.append('CON')
        partes.append(_grupo_a_letras(cents))
        partes.append('CÉNTIMO' if cents == 1 else 'CÉNTIMOS')

    return ' '.join(partes)


# ── Helpers de estilo ─────────────────────────────────────────────────────────

def aplicar_estilo_cabecera(ws, fila: int, n_cols: int) -> None:
    """Aplica estilo de cabecera (azul oscuro, texto blanco) a una fila fusionada."""
    if n_cols > 1:
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=n_cols)
    cell = ws.cell(fila, 1)
    cell.fill      = _FC
    cell.font      = _FT
    cell.alignment = _AC


def aplicar_estilo_capitulo(ws, fila: int, n_cols: int) -> None:
    """Aplica estilo de capítulo (azul claro, negrita) a las celdas de una fila."""
    for c in range(1, n_cols + 1):
        ws.cell(fila, c).fill = _FCP
        ws.cell(fila, c).font = _FCF


def aplicar_estilo_total(ws, fila: int, n_cols: int) -> None:
    """Aplica estilo de total (azul oscuro, texto blanco, negrita) a una fila."""
    for c in range(1, n_cols + 1):
        ws.cell(fila, c).fill = _FC
        ws.cell(fila, c).font = _FTO


def ajustar_columnas(ws, config: dict[str, int]) -> None:
    """Ajusta anchos de columna según config {letra_columna: ancho}."""
    for col_letter, width in config.items():
        ws.column_dimensions[col_letter].width = width


# ── Helpers de datos ─────────────────────────────────────────────────────────

def get_recurso_info(codigo: str, presupuesto: Presupuesto) -> tuple[str, str, float]:
    """
    Retorna (descripcion, unidad, precio) para un código de recurso.
    Busca en MO, MT y MQ; devuelve el código como descripción si no se encuentra.
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
    return codigo, '-', 0.0


def tipo_recurso(codigo: str) -> str:
    """Clasifica un código de recurso en MO, MT, MQ o %."""
    if codigo.startswith('MO-'):
        return 'MO'
    if codigo.startswith('MT-') or codigo.startswith('%MT'):
        return 'MT'
    if codigo.startswith('MQ-'):
        return 'MQ'
    return '%'


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
    Retorna True si la partida tiene mano de obra Y material simultáneamente.
    Criterio para incluir en la base de cálculo de la liquidación máxima.
    """
    items = presupuesto.descompuestos_raw.get(codigo, [])
    tiene_mo = any(r.startswith('MO-') for r, _ in items)
    tiene_mt = any(r.startswith('MT-') or r.startswith('%MT') for r, _ in items)
    return tiene_mo and tiene_mt


# ── Helpers internos de escritura Excel ──────────────────────────────────────

def _row_height(desc: str) -> float:
    """Estima la altura de fila en puntos según la longitud de la descripción."""
    return max(15.0, min(60.0, len(str(desc)) // 55 * 15 + 15))


def _cel_desc(ws, fila: int, col: int, texto: str, font=None) -> None:
    """Escribe una celda de descripción con wrap_text y ajuste de altura."""
    cell              = ws.cell(fila, col, texto)
    cell.alignment    = _AD
    cell.font         = font or _FN
    ws.row_dimensions[fila].height = _row_height(texto)


def _cel_eur(ws, fila: int, col: int, valor: float, font=None) -> None:
    """Escribe una celda numérica en euros con formato y alineación derecha."""
    cell              = ws.cell(fila, col, valor)
    cell.number_format = _FE
    cell.alignment    = _AN
    cell.font         = font or _FN


def _cabecera_tabla(ws, headers: list[str]) -> None:
    """Escribe una fila de cabecera de tabla con estilo de cabecera."""
    ws.append(headers)
    r = ws.max_row
    for c, _ in enumerate(headers, 1):
        ws.cell(r, c).fill      = _FC
        ws.cell(r, c).font      = _FT
        ws.cell(r, c).alignment = _AC


def _sep_vacia(ws, altura: float = 4.0) -> None:
    """Inserta una fila vacía separadora con altura reducida."""
    ws.append([])
    ws.row_dimensions[ws.max_row].height = altura


def _capitulo_limpio(codigo: str) -> str:
    """Devuelve el código de capítulo sin el sufijo '#'."""
    return codigo.rstrip('#')


def _get_cap_gr(presupuesto: Presupuesto) -> 'Capitulo | None':
    """Localiza el capítulo de gestión de residuos por código."""
    return next((c for c in presupuesto.capitulos if c.codigo == CODIGO_CAPITULO_GR), None)


def _fila_separador_cascade(ws, n_cols: int) -> None:
    """Inserta una fila separadora con fondo azul muy claro entre bloques de cascada."""
    ws.append([''] * n_cols)
    r = ws.max_row
    for c in range(1, n_cols + 1):
        ws.cell(r, c).fill = _FSC
    ws.row_dimensions[r].height = 6.0


def _escribir_cascade(ws, filas: list[tuple], n_cols: int = 3) -> None:
    """
    Escribe una tabla en cascada de 3 columnas: (etiqueta, pct_str, importe).
    None como etiqueta inserta un separador visual.
    True como etiqueta aplica estilo de total.
    """
    for item in filas:
        if item is None:
            _fila_separador_cascade(ws, n_cols)
            continue
        label, pct, valor, es_total = item
        ws.append([label, pct, ''] if n_cols >= 3 else [label, pct])
        r = ws.max_row
        ws.cell(r, 1).font = _FTO if es_total else _FN
        ws.cell(r, 1).fill = _FC if es_total else _FB
        ws.cell(r, 2).font = _FTO if es_total else _FN
        ws.cell(r, 2).fill = _FC if es_total else _FB
        ws.cell(r, 2).alignment = _AC
        _cel_eur(ws, r, 3, valor, font=_FTO if es_total else _FN)
        ws.cell(r, 3).fill = _FC if es_total else _FB


def descripcion_completa(partida: Partida) -> str:
    """Retorna la descripción larga (~T) si existe; si no, la descripción corta."""
    return partida.descripcion_larga or partida.descripcion


# ── PRES.01 — Cuadro de Oferta ───────────────────────────────────────────────

def generar_pres01(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Cuadro de Oferta con columnas amarillas para el licitador."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Cuadro de Oferta'
    N = 8

    # Encabezado del documento
    ws.append(['CUADRO DE OFERTA'] + [''] * (N - 1))
    aplicar_estilo_cabecera(ws, 1, N)
    ws.row_dimensions[1].height = 22

    ws.append(['Rellene las columnas en amarillo con su precio unitario ofertado']
              + [''] * (N - 1))
    r2 = ws.max_row
    ws.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=N)
    ws.cell(r2, 1).fill      = _FA
    ws.cell(r2, 1).font      = Font(name='Calibri', size=10, italic=True, color='000000')
    ws.cell(r2, 1).alignment = _AC
    ws.row_dimensions[r2].height = 16

    _cabecera_tabla(ws, [
        'Código', 'Descripción', 'Ud.', 'Cantidad',
        'P. Unit. ref. (€)', 'Importe ref. (€)',
        'P. Unit. ofertado (€)', 'Importe ofertado (€)',
    ])
    ws.freeze_panes = 'A4'

    current_cap = None
    current_sub = None
    alt = True

    for cap, sub, partida in iter_todas_partidas(presupuesto):
        if cap is not current_cap:
            current_cap = cap
            current_sub = None
            ws.append([_capitulo_limpio(cap.codigo), cap.descripcion]
                      + [''] * (N - 2))
            r = ws.max_row
            aplicar_estilo_capitulo(ws, r, N)
            _cel_desc(ws, r, 2, cap.descripcion, font=_FCF)

        if sub is not current_sub:
            current_sub = sub
            if sub is not None:
                ws.append([_capitulo_limpio(sub.codigo), '  ' + sub.descripcion]
                          + [''] * (N - 2))
                r = ws.max_row
                for c in range(1, N + 1):
                    ws.cell(r, c).fill = _FSC
                    ws.cell(r, c).font = _FCF
                _cel_desc(ws, r, 2, '  ' + sub.descripcion, font=_FCF)
                ws.cell(r, 2).fill = _FSC

        importe = partida.precio_unitario * partida.cantidad
        fill = _FD if alt else _FB
        alt  = not alt

        ws.append([partida.codigo, partida.descripcion, partida.unidad,
                   partida.cantidad, partida.precio_unitario, importe, '', ''])
        r = ws.max_row
        for c in range(1, N + 1):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = _FN
        _cel_desc(ws, r, 2, partida.descripcion)
        ws.cell(r, 2).fill = fill
        ws.cell(r, 4).number_format = _FQ
        ws.cell(r, 4).alignment = _AN
        ws.cell(r, 5).number_format = _FE
        ws.cell(r, 5).alignment = _AN
        ws.cell(r, 6).number_format = _FE
        ws.cell(r, 6).alignment = _AN
        ws.cell(r, 7).fill = _FA   # precio ofertado
        ws.cell(r, 8).fill = _FA   # importe ofertado

    # Fila TOTAL PEM
    ws.append(['', 'TOTAL PRESUPUESTO DE EJECUCIÓN MATERIAL'] + [''] * 3
              + [presupuesto.importe_total, '', ''])
    r = ws.max_row
    aplicar_estilo_total(ws, r, N)
    _cel_eur(ws, r, 6, presupuesto.importe_total, font=_FTO)
    ws.cell(r, 6).fill = _FC

    ajustar_columnas(ws, {'A': 16, 'B': 55, 'C': 6, 'D': 12,
                          'E': 16, 'F': 16, 'G': 18, 'H': 18})
    wb.save(filepath)


# ── PRES.02.01 — Cuadro de Precios n.º 1 ─────────────────────────────────────

def generar_pres0201(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Cuadro de Precios n.º 1 con el precio unitario en letra."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Cuadro de Precios 1'
    N = 3

    ws.append(['CUADRO DE PRECIOS N.º 1'] + [''] * (N - 1))
    aplicar_estilo_cabecera(ws, 1, N)
    ws.row_dimensions[1].height = 22

    _cabecera_tabla(ws, ['Código', 'Descripción del precio', 'Precio unit. (€)'])
    ws.freeze_panes = 'A3'

    # Recopilar partidas con sus capítulos
    current_cap = None
    for cap, sub, partida in iter_todas_partidas(presupuesto):
        if cap is not current_cap:
            current_cap = cap
            ws.append([_capitulo_limpio(cap.codigo), cap.descripcion, ''])
            r = ws.max_row
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N)
            aplicar_estilo_capitulo(ws, r, N)
            _cel_desc(ws, r, 1,
                      f'CAPÍTULO {_capitulo_limpio(cap.codigo)} — {cap.descripcion}',
                      font=_FCF)

        desc_mayus   = descripcion_completa(partida).upper()
        precio_letra = numero_a_letras(partida.precio_unitario)
        contenido_b  = f'{desc_mayus}\n\nPRECIO: {precio_letra}'
        ws.append([partida.codigo, contenido_b, partida.precio_unitario])
        r = ws.max_row
        for c in range(1, N + 1):
            ws.cell(r, c).fill = _FB
            ws.cell(r, c).font = _FN
        ws.cell(r, 1).alignment = _AT
        _cel_desc(ws, r, 2, contenido_b)
        _cel_eur(ws, r, 3, partida.precio_unitario)
        _sep_vacia(ws, 4.0)

    ajustar_columnas(ws, {'A': 16, 'B': 72, 'C': 18})
    wb.save(filepath)


# ── PRES.02.02 — Cuadro de Precios n.º 2 ─────────────────────────────────────

def generar_pres0202(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Cuadro de Precios n.º 2 con descomposición de cada partida."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Cuadro de Precios 2'
    N = 7

    ws.append(['CUADRO DE PRECIOS N.º 2'] + [''] * (N - 1))
    aplicar_estilo_cabecera(ws, 1, N)
    ws.row_dimensions[1].height = 22

    _cabecera_tabla(ws, ['Tipo', 'Código', 'Descripción',
                         'Ud.', 'Cant./ud', 'P. unit. (€)', 'Coste (€)'])
    ws.freeze_panes = 'A3'

    for cap, sub, partida in iter_todas_partidas(presupuesto):
        # Fila resumen de partida
        ws.append([partida.codigo, partida.descripcion[:60], partida.unidad,
                   '', '', '', ''])
        r = ws.max_row
        aplicar_estilo_capitulo(ws, r, N)
        _cel_desc(ws, r, 2, partida.descripcion[:60], font=_FCF)
        ws.cell(r, 2).fill = _FCP

        # Fila descripción técnica completa
        desc_larga = descripcion_completa(partida)
        ws.append(['', desc_larga, '', '', '', '', ''])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
        ws.cell(r, 2).fill = _FB
        _cel_desc(ws, r, 2, desc_larga)

        items = presupuesto.descompuestos_raw.get(partida.codigo, [])

        if not items:
            # Partida de precio alzado
            ws.append(['', 'Precio alzado', '', '', '', '', partida.precio_unitario])
            r = ws.max_row
            for c in range(1, N + 1):
                ws.cell(r, c).fill = _FD
                ws.cell(r, c).font = _FDE
            _cel_eur(ws, r, 7, partida.precio_unitario, font=_FDE)
            ws.cell(r, 7).fill = _FD
        else:
            # Agrupar por tipo
            grupos = {'MO': [], 'MQ': [], 'MT': [], '%': []}
            for cod_rec, cant in items:
                grupos[tipo_recurso(cod_rec)].append((cod_rec, cant))

            suma_directos = 0.0
            for tipo, label in [('MO', 'Mano de obra'), ('MQ', 'Maquinaria'), ('MT', 'Materiales'), ('%', 'Costes indirectos')]:
                if not grupos[tipo]:
                    continue
                # Subtítulo del grupo
                ws.append(['', label, '', '', '', '', ''])
                r = ws.max_row
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
                for c in range(1, N + 1):
                    ws.cell(r, c).fill = _FD
                    ws.cell(r, c).font = Font(name='Calibri', size=9,
                                              bold=True, color=PRES_GRIS_TEXTO)

                for i, (cod_rec, cant) in enumerate(grupos[tipo]):
                    desc, unidad, precio = get_recurso_info(cod_rec, presupuesto)
                    coste = round(cant * precio, 4)
                    suma_directos += coste
                    fill = _FD if i % 2 == 0 else _FB
                    ws.append([tipo, cod_rec, desc, unidad, cant, precio, coste])
                    r = ws.max_row
                    for c in range(1, N + 1):
                        ws.cell(r, c).fill = fill
                        ws.cell(r, c).font = _FDE
                    _cel_desc(ws, r, 3, desc, font=_FDE)
                    ws.cell(r, 3).fill = fill
                    ws.cell(r, 5).number_format = _FQ
                    ws.cell(r, 5).alignment = _AN
                    ws.cell(r, 6).number_format = _FE
                    ws.cell(r, 6).alignment = _AN
                    ws.cell(r, 7).number_format = _FE
                    ws.cell(r, 7).alignment = _AN

            # Subtotal directos
            ws.append(['', '', '', '', '', 'Suma directos:', round(suma_directos, 2)])
            r = ws.max_row
            for c in range(1, N + 1):
                ws.cell(r, c).fill = _FD
                ws.cell(r, c).font = _FCF
            _cel_eur(ws, r, 7, round(suma_directos, 2), font=_FCF)
            ws.cell(r, 7).fill = _FD

        # Fila total partida
        ws.append(['', 'PRECIO UNITARIO TOTAL', '', '', '', '', partida.precio_unitario])
        r = ws.max_row
        aplicar_estilo_total(ws, r, N)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        _cel_eur(ws, r, 7, partida.precio_unitario, font=_FTO)
        ws.cell(r, 7).fill = _FC

        _sep_vacia(ws, 6.0)

    ajustar_columnas(ws, {'A': 8, 'B': 16, 'C': 55, 'D': 8, 'E': 12, 'F': 16, 'G': 16})
    wb.save(filepath)


# ── PRES.02.03 — Presupuesto Descompuesto y Mediciones ───────────────────────

def generar_pres0203(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Presupuesto Descompuesto y Mediciones por capítulos."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Presupuesto Descompuesto'
    N = 6

    ws.append(['PRESUPUESTO DESCOMPUESTO Y MEDICIONES'] + [''] * (N - 1))
    aplicar_estilo_cabecera(ws, 1, N)
    ws.row_dimensions[1].height = 22

    _cabecera_tabla(ws, ['Código', 'Descripción', 'Ud.',
                         'Cantidad', 'P. unit. (€)', 'Importe (€)'])
    ws.freeze_panes = 'A3'

    for cap in presupuesto.capitulos:
        ws.append([f'CAPÍTULO {_capitulo_limpio(cap.codigo)} — {cap.descripcion}']
                  + [''] * (N - 1))
        r = ws.max_row
        aplicar_estilo_cabecera(ws, r, N)
        _cel_desc(ws, r, 1,
                  f'CAPÍTULO {_capitulo_limpio(cap.codigo)} — {cap.descripcion}',
                  font=_FT)
        ws.cell(r, 1).fill = _FC
        ws.row_dimensions[r].height = 18

        def escribir_partidas_de(nodo: Capitulo) -> None:
            for sub in nodo.subcapitulos:
                ws.append(['', f'{_capitulo_limpio(sub.codigo)} — {sub.descripcion}']
                          + [''] * (N - 2))
                r_sub = ws.max_row
                ws.merge_cells(start_row=r_sub, start_column=1, end_row=r_sub, end_column=N)
                for c in range(1, N + 1):
                    ws.cell(r_sub, c).fill = _FSC
                    ws.cell(r_sub, c).font = _FCF
                _cel_desc(ws, r_sub, 1,
                          f'{_capitulo_limpio(sub.codigo)} — {sub.descripcion}', font=_FCF)
                ws.cell(r_sub, 1).fill = _FSC
                escribir_partidas_de(sub)

            for partida in nodo.partidas:
                importe = partida.precio_unitario * partida.cantidad

                # Fila resumen de la partida
                ws.append([partida.codigo, partida.descripcion[:60],
                           partida.unidad, partida.cantidad,
                           partida.precio_unitario, importe])
                r_p = ws.max_row
                aplicar_estilo_capitulo(ws, r_p, N)
                _cel_desc(ws, r_p, 2, partida.descripcion[:60], font=_FCF)
                ws.cell(r_p, 2).fill = _FCP
                ws.cell(r_p, 4).number_format = _FQ
                ws.cell(r_p, 4).alignment = _AN
                _cel_eur(ws, r_p, 5, partida.precio_unitario, font=_FCF)
                ws.cell(r_p, 5).fill = _FCP
                _cel_eur(ws, r_p, 6, importe, font=_FCF)
                ws.cell(r_p, 6).fill = _FCP

                # Fila descripción técnica completa
                desc_larga = descripcion_completa(partida)
                ws.append(['', desc_larga, '', '', '', ''])
                r_desc = ws.max_row
                ws.merge_cells(start_row=r_desc, start_column=2, end_row=r_desc, end_column=N)
                ws.cell(r_desc, 2).fill = _FB
                _cel_desc(ws, r_desc, 2, desc_larga)

                # Descompost (7 cols ajustados a 6 del documento)
                items = presupuesto.descompuestos_raw.get(partida.codigo, [])
                if items:
                    grupos = {'MO': [], 'MQ': [], 'MT': [], '%': []}
                    for cod_rec, cant in items:
                        grupos[tipo_recurso(cod_rec)].append((cod_rec, cant))

                    for tipo, label in [('MO', 'Mano de obra'), ('MQ', 'Maquinaria'),
                                        ('MT', 'Materiales'), ('%', 'Costes indirectos')]:
                        if not grupos[tipo]:
                            continue
                        for i, (cod_rec, cant) in enumerate(grupos[tipo]):
                            desc_r, unidad_r, precio_r = get_recurso_info(cod_rec, presupuesto)
                            coste = round(cant * precio_r, 4)
                            fill  = _FD if i % 2 == 0 else _FB
                            ws.append([tipo, f'{cod_rec} — {desc_r}',
                                       unidad_r, cant, precio_r, coste])
                            r_r = ws.max_row
                            for c in range(1, N + 1):
                                ws.cell(r_r, c).fill = fill
                                ws.cell(r_r, c).font = _FDE
                            _cel_desc(ws, r_r, 2, f'{cod_rec} — {desc_r}', font=_FDE)
                            ws.cell(r_r, 2).fill = fill
                            ws.cell(r_r, 4).number_format = _FQ
                            ws.cell(r_r, 4).alignment    = _AN
                            ws.cell(r_r, 5).number_format = _FE
                            ws.cell(r_r, 5).alignment    = _AN
                            ws.cell(r_r, 6).number_format = _FE
                            ws.cell(r_r, 6).alignment    = _AN

                # Fila medición
                ws.append(['', 'Medición:', partida.unidad, partida.cantidad, '', ''])
                r_med = ws.max_row
                for c in range(1, N + 1):
                    ws.cell(r_med, c).fill = _FD
                    ws.cell(r_med, c).font = _FCF
                ws.cell(r_med, 4).number_format = _FQ
                ws.cell(r_med, 4).alignment = _AN

                # Fila cálculo total de la partida
                desc_calc = (f'Total partida: {partida.cantidad} {partida.unidad} '
                             f'× {partida.precio_unitario:.2f} €/ud')
                ws.append(['', desc_calc, '', '', '', importe])
                r_tot = ws.max_row
                aplicar_estilo_total(ws, r_tot, N)
                ws.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=5)
                _cel_eur(ws, r_tot, 6, importe, font=_FTO)
                ws.cell(r_tot, 6).fill = _FC

                _sep_vacia(ws, 5.0)

        escribir_partidas_de(cap)

        # Total capítulo
        ws.append(['', f'TOTAL CAPÍTULO {_capitulo_limpio(cap.codigo)}: {cap.descripcion}',
                   '', '', '', cap.importe_total])
        r_tc = ws.max_row
        for c in range(1, N + 1):
            ws.cell(r_tc, c).fill = _FCP
            ws.cell(r_tc, c).font = _FCF
        ws.merge_cells(start_row=r_tc, start_column=1, end_row=r_tc, end_column=5)
        _cel_eur(ws, r_tc, 6, cap.importe_total, font=_FCF)
        ws.cell(r_tc, 6).fill = _FCP
        ws.append([])

    # PEM final
    ws.append(['PRESUPUESTO DE EJECUCIÓN MATERIAL (PEM)'] + [''] * (N - 2)
              + [presupuesto.importe_total])
    r = ws.max_row
    aplicar_estilo_total(ws, r, N)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N - 1)
    _cel_eur(ws, r, N, presupuesto.importe_total, font=_FTO)
    ws.cell(r, N).fill = _FC

    ajustar_columnas(ws, {'A': 8, 'B': 55, 'C': 6, 'D': 12, 'E': 16, 'F': 16})
    wb.save(filepath)


# ── PRES.02.04 — Resumen por Capítulos ───────────────────────────────────────

def generar_pres0204(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el Resumen por Capítulos con cascada financiera PEM → PGL."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Resumen Capítulos'
    N3 = 3

    cap_gr     = _get_cap_gr(presupuesto)
    importe_gr = cap_gr.importe_total if cap_gr else 0.0
    pem_sin_gr = presupuesto.importe_total - importe_gr
    gg  = pem_sin_gr * PORCENTAJE_GG
    bi  = pem_sin_gr * PORCENTAJE_BI
    pec = pem_sin_gr + gg + bi
    iva_obra = pec * PORCENTAJE_IVA_OBRA
    iva_gr   = importe_gr * PORCENTAJE_IVA_GR
    pgl      = pec + importe_gr + iva_obra + iva_gr

    # ── Sección 1: tabla por capítulos ────────────────────────────────────────
    ws.append(['RESUMEN POR CAPÍTULOS'] + [''] * (N3 - 1))
    aplicar_estilo_cabecera(ws, 1, N3)
    ws.row_dimensions[1].height = 22

    _cabecera_tabla(ws, ['Capítulo', 'Descripción', 'Importe (€)'])
    ws.freeze_panes = 'A3'

    for i, cap in enumerate(presupuesto.capitulos):
        ws.append([_capitulo_limpio(cap.codigo), cap.descripcion, cap.importe_total])
        r = ws.max_row
        fill = _FD if i % 2 == 0 else _FB
        for c in range(1, N3 + 1):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = _FN
        _cel_desc(ws, r, 2, cap.descripcion)
        ws.cell(r, 2).fill = fill
        _cel_eur(ws, r, 3, cap.importe_total)
        ws.cell(r, 3).fill = fill

    ws.append(['', 'TOTAL PRESUPUESTO DE EJECUCIÓN MATERIAL (PEM)',
               presupuesto.importe_total])
    r = ws.max_row
    aplicar_estilo_total(ws, r, N3)
    _cel_eur(ws, r, 3, presupuesto.importe_total, font=_FTO)
    ws.cell(r, 3).fill = _FC

    ws.append([])
    ws.append([])

    # ── Sección 2: cascada general ────────────────────────────────────────────
    ws.append(['RESUMEN GENERAL DEL PRESUPUESTO'] + [''] * (N3 - 1))
    r = ws.max_row
    aplicar_estilo_cabecera(ws, r, N3)
    ws.row_dimensions[r].height = 22

    _cabecera_tabla(ws, ['Concepto', '% aplicado', 'Importe (€)'])

    filas = [
        ('PEM (sin residuos)',        '—',   pem_sin_gr, False),
        (f'Gastos Generales',         f'{PORCENTAJE_GG*100:.0f}%', gg,  False),
        (f'Beneficio Industrial',     f'{PORCENTAJE_BI*100:.0f}%', bi,  False),
        None,
        ('PEC',                       '—',   pec,        True),
        ('Gestión de Residuos (GR)',  '—',   importe_gr, False),
        None,
        (f'IVA obra',                 f'{PORCENTAJE_IVA_OBRA*100:.0f}%', iva_obra, False),
        (f'IVA residuos',             f'{PORCENTAJE_IVA_GR*100:.0f}%',   iva_gr,   False),
        None,
        ('PGL (Presupuesto Global de Licitación)', '—', pgl, True),
    ]
    _escribir_cascade(ws, filas, N3)

    ajustar_columnas(ws, {'A': 52, 'B': 14, 'C': 18})
    wb.save(filepath)


# ── PRES.05 — VEC y Liquidación ──────────────────────────────────────────────

def generar_pres05(presupuesto: Presupuesto, filepath: Path) -> None:
    """Genera el VEC y Liquidación — Bloque 1 (VEC) primero, Bloque 2 después."""
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'VEC y Liquidación'
    N3 = 3

    cap_gr     = _get_cap_gr(presupuesto)
    importe_gr = cap_gr.importe_total if cap_gr else 0.0
    pem_sin_gr = presupuesto.importe_total - importe_gr
    gg  = pem_sin_gr * PORCENTAJE_GG
    bi  = pem_sin_gr * PORCENTAJE_BI
    pec = pem_sin_gr + gg + bi

    # ── Calcular base de liquidación (Bloque 2) ───────────────────────────────
    pem_liq = sum(
        p.precio_unitario * p.cantidad
        for cap, sub, p in iter_todas_partidas(presupuesto)
        if es_partida_liquidable(p.codigo, presupuesto)
    )
    gg_liq   = pem_liq * PORCENTAJE_GG
    bi_liq   = pem_liq * PORCENTAJE_BI
    pec_liq  = pem_liq + gg_liq + bi_liq
    liq_max  = (pec_liq + importe_gr) * PORCENTAJE_LIQUIDACION

    vec      = pec + importe_gr + liq_max
    iva_obra = pec * PORCENTAJE_IVA_OBRA
    iva_gr   = importe_gr * PORCENTAJE_IVA_GR
    pgl      = vec + iva_obra + iva_gr

    # ── BLOQUE 1: VEC ─────────────────────────────────────────────────────────
    ws.append(['BLOQUE 1 — VALOR ESTIMADO DEL CONTRATO (VEC)'] + [''] * (N3 - 1))
    aplicar_estilo_cabecera(ws, 1, N3)
    ws.row_dimensions[1].height = 22

    _cabecera_tabla(ws, ['Concepto', '% aplicado', 'Importe (€)'])

    filas_b1 = [
        ('PEM (sin residuos)',                       '—',  pem_sin_gr, False),
        (f'Gastos Generales',                        f'{PORCENTAJE_GG*100:.0f}%',   gg,      False),
        (f'Beneficio Industrial',                    f'{PORCENTAJE_BI*100:.0f}%',   bi,      False),
        None,
        ('PEC',                                      '—',  pec,        True),
        ('Gestión de Residuos (GR)',                 '—',  importe_gr, False),
        (f'Liquidación máxima ({PORCENTAJE_LIQUIDACION*100:.0f}%)', '—', liq_max, False),
        None,
        ('VEC (sin IVA)',                            '—',  vec,        True),
        (f'IVA obra ({PORCENTAJE_IVA_OBRA*100:.0f}%)',  f'{PORCENTAJE_IVA_OBRA*100:.0f}%', iva_obra, False),
        (f'IVA residuos ({PORCENTAJE_IVA_GR*100:.0f}%)', f'{PORCENTAJE_IVA_GR*100:.0f}%', iva_gr, False),
        None,
        ('PGL (Presupuesto Global de Licitación)',   '—',  pgl,        True),
    ]
    _escribir_cascade(ws, filas_b1, N3)

    ws.append([])
    ws.append([])
    ws.append([])

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
    ws.cell(r, 1).fill      = _FA
    ws.cell(r, 1).font      = Font(name='Calibri', size=9, italic=True, color='000000')
    ws.cell(r, 1).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.row_dimensions[r].height = 40
    ws.append([])

    _cabecera_tabla(ws, ['Código', 'Descripción', 'Importe (€)'])

    alt = True
    for cap, sub, partida in iter_todas_partidas(presupuesto):
        if not es_partida_liquidable(partida.codigo, presupuesto):
            continue
        importe_p = partida.precio_unitario * partida.cantidad
        fill = _FD if alt else _FB
        alt  = not alt
        ws.append([partida.codigo, partida.descripcion, importe_p])
        r = ws.max_row
        for c in range(1, N3 + 1):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = _FN
        _cel_desc(ws, r, 2, partida.descripcion)
        ws.cell(r, 2).fill = fill
        _cel_eur(ws, r, 3, importe_p)
        ws.cell(r, 3).fill = fill

    ws.append(['', 'TOTAL BASE IMPONIBLE LIQUIDACIÓN', pem_liq])
    r = ws.max_row
    for c in range(1, N3 + 1):
        ws.cell(r, c).fill = _FCP
        ws.cell(r, c).font = _FCF
    _cel_eur(ws, r, 3, pem_liq, font=_FCF)
    ws.cell(r, 3).fill = _FCP
    ws.append([])

    filas_b2 = [
        ('PEM liquidación',          '—',   pem_liq,    False),
        (f'Gastos Generales',        f'{PORCENTAJE_GG*100:.0f}%',  gg_liq,  False),
        (f'Beneficio Industrial',    f'{PORCENTAJE_BI*100:.0f}%',  bi_liq,  False),
        None,
        ('PEC liquidación',          '—',   pec_liq,    True),
        ('Gestión de Residuos (GR)', '—',   importe_gr, False),
        None,
        (f'LIQUIDACIÓN MÁXIMA ({PORCENTAJE_LIQUIDACION*100:.0f}%)', '—', liq_max, True),
    ]
    _escribir_cascade(ws, filas_b2, N3)

    ajustar_columnas(ws, {'A': 55, 'B': 14, 'C': 18})
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

    ws.append(['JUSTIFICACIÓN DE PRECIOS — RECURSOS EMPLEADOS EN EL PROYECTO']
              + [''] * (N - 1))
    aplicar_estilo_cabecera(ws, 1, N)
    ws.row_dimensions[1].height = 22

    # Acumular cantidades totales por recurso
    mo_tot: dict[str, float] = {}
    mt_tot: dict[str, float] = {}
    mq_tot: dict[str, float] = {}

    for cap, sub, partida in iter_todas_partidas(presupuesto):
        for cod_rec, cant_ud in presupuesto.descompuestos_raw.get(partida.codigo, []):
            total = cant_ud * partida.cantidad
            tr = tipo_recurso(cod_rec)
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
            ws.cell(r, c).fill = _FCP
            ws.cell(r, c).font = _FCF
        ws.cell(r, 1).alignment = _AT

        _cabecera_tabla(ws, ['Código', 'Descripción', 'Ud.',
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
            fill = _FD if i % 2 == 0 else _FB
            ws.append([cod, recurso.descripcion, unidad, precio, cant, coste])
            r = ws.max_row
            for c in range(1, N + 1):
                ws.cell(r, c).fill = fill
                ws.cell(r, c).font = _FN
            _cel_desc(ws, r, 2, recurso.descripcion)
            ws.cell(r, 2).fill = fill
            ws.cell(r, 4).number_format = _FE
            ws.cell(r, 4).alignment    = _AN
            ws.cell(r, 5).number_format = _FQ
            ws.cell(r, 5).alignment    = _AN
            _cel_eur(ws, r, 6, coste)
            ws.cell(r, 6).fill = fill

        ws.append(['', f'TOTAL {label_total}', '', '', '', coste_total_seccion])
        r = ws.max_row
        aplicar_estilo_total(ws, r, N)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        _cel_eur(ws, r, 6, coste_total_seccion, font=_FTO)
        ws.cell(r, 6).fill = _FC
        ws.append([])

    escribir_seccion('MANO DE OBRA', presupuesto.recursos_mo, mo_tot, 'MANO DE OBRA')
    escribir_seccion('MAQUINARIA',   presupuesto.recursos_mq, mq_tot, 'MAQUINARIA')
    escribir_seccion('MATERIALES',   presupuesto.recursos_mt, mt_tot, 'MATERIALES')

    ajustar_columnas(ws, {'A': 18, 'B': 55, 'C': 8, 'D': 14, 'E': 18, 'F': 18})
    wb.save(filepath)


# ── Punto de entrada ─────────────────────────────────────────────────────────

def generar_todos(presupuesto: Presupuesto, output_dir: Path) -> list[Path]:
    """
    Genera los siete documentos de presupuesto en output_dir.
    Retorna lista de rutas de los ficheros generados.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    documentos = [
        ('PRES.01_Cuadro_Oferta.xlsx',               generar_pres01),
        ('PRES.02.01_Cuadro_Precios_1.xlsx',          generar_pres0201),
        ('PRES.02.02_Cuadro_Precios_2.xlsx',          generar_pres0202),
        ('PRES.02.03_Presupuesto_Descompuesto.xlsx',  generar_pres0203),
        ('PRES.02.04_Resumen_Capitulos.xlsx',         generar_pres0204),
        ('PRES.05_VEC_Liquidacion.xlsx',              generar_pres05),
        ('JUST_PRECIOS_Recursos.xlsx',                generar_just_precios),
    ]
    rutas: list[Path] = []
    for nombre, funcion in documentos:
        ruta = output_dir / nombre
        funcion(presupuesto, ruta)
        rutas.append(ruta)
    return rutas

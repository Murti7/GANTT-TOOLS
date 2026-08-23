"""
Módulo responsable de generar el archivo Excel de análisis del presupuesto.

Responsabilidad: recibir el Presupuesto parseado desde BC3 y —opcionalmente—
las planificaciones calculadas, y escribir un archivo .xlsx con hasta cuatro
tipos de hojas: resumen por capítulo, detalle de partidas, simulador
multi-escenario y tablas de importación MS Project por escenario.
"""

import io
import math
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from gantt.bc3.models import Capitulo, Presupuesto, RecursoMO
from gantt.planning.analyser import (
    calcular_carga_semanal_recursos,
    calcular_holguras,
    calcular_recomendacion,
    calcular_sensibilidad,
    calcular_sobredimensionados,
    dias_habiles_entre,
    duracion_total_proyecto,
)
from gantt.planning.models import PlanificacionProyecto
from gantt.reporting.styles import (
    GANTT_FILL_ALT, GANTT_FILL_CAP, GANTT_FILL_GREEN, GANTT_FILL_HEADER, GANTT_FILL_RED, GANTT_FILL_WHITE, GANTT_FILL_YELLOW,
    GANTT_FONT_BOLD, GANTT_FONT_HEADER, GANTT_FONT_ITALIC,
    DocumentPalette, build_palette,
)
from gantt.reporting.documents import DocumentMetadata, DocumentPurpose, DocumentType, PlanningMetadata
from gantt.reporting.presentation import PresentationContext
from gantt.reporting.rendering import (
    apply_document_footer,
    configure_excel_printing,
    metadata_from_project_config,
    presentation_from_palette_company,
    render_budget_header,
    render_analysis_header,
    render_planning_header,
)


def _apply_palette_to_legacy_styles(palette: DocumentPalette) -> None:
    """
    Puente temporal para el exporter historico, que todavia usa constantes
    globales. Mantiene API interna estable mientras se completa la migracion a
    estilos inyectados por funcion.
    """
    global GANTT_FILL_ALT, GANTT_FILL_CAP, GANTT_FILL_HEADER, GANTT_FILL_WHITE, GANTT_FILL_YELLOW
    global GANTT_FONT_BOLD, GANTT_FONT_HEADER, GANTT_FONT_ITALIC

    GANTT_FILL_HEADER = palette.fill_header
    GANTT_FILL_CAP = palette.fill_subhead
    GANTT_FILL_ALT = palette.fill_alt
    GANTT_FILL_WHITE = palette.fill_white
    GANTT_FILL_YELLOW = palette.fill_yellow
    GANTT_FONT_HEADER = palette.font_header
    GANTT_FONT_BOLD = palette.font_bold
    GANTT_FONT_ITALIC = Font(name=palette.font_family, italic=True, size=10)


def horas_capitulo(capitulo: Capitulo, codigos: list[str]) -> dict[str, float]:
    """
    Calcula recursivamente las horas totales de MO por recurso en un capítulo.
    Acumula las partidas directas y las de todos los subcapítulos hijos.
    """
    horas = {c: 0.0 for c in codigos}
    for partida in capitulo.partidas:
        for linea in partida.lineas_mo:
            if linea.codigo_recurso in horas:
                horas[linea.codigo_recurso] += linea.cantidad * partida.cantidad
    for subcap in capitulo.subcapitulos:
        for c, h in horas_capitulo(subcap, codigos).items():
            horas[c] += h
    return horas


def iter_partidas(capitulo: Capitulo, codigo_cap: str, codigo_subcap: str = ''):
    """
    Genera tuplas (codigo_cap, codigo_subcap, partida) en orden de aparición.
    Primero las partidas directas del capítulo, luego las de cada subcapítulo.
    """
    for partida in capitulo.partidas:
        yield codigo_cap, codigo_subcap, partida
    for subcap in capitulo.subcapitulos:
        yield from iter_partidas(subcap, codigo_cap, subcap.codigo)


def adjust_col_widths(ws) -> None:
    """Ajusta el ancho de cada columna al máximo contenido de sus celdas."""
    for col in ws.columns:
        max_len = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 60)


def first_filterable_row(ws, start_row: int) -> int:
    """Primera fila usable como header de autofiltro, evitando merges."""
    merged_rows = {row for merged in ws.merged_cells.ranges for row, _col in merged.cells}
    for row in range(start_row, ws.max_row + 1):
        non_empty = sum(
            1
            for col in range(1, ws.max_column + 1)
            if ws.cell(row, col).value not in (None, "")
        )
        if row not in merged_rows and non_empty >= 2:
            return row
    return start_row



def write_resumen(ws, presupuesto: Presupuesto, recursos: list[RecursoMO]) -> None:
    """
    Escribe la hoja 'Resumen por capítulo': tabla principal de capítulos/subcapítulos
    seguida de una sección de resumen de mano de obra del proyecto.
    """
    codigos = [r.codigo for r in recursos]
    headers = (
        ['Código', 'Descripción', 'Importe total (€)', 'H MO total']
        + [r.descripcion for r in recursos]
        + ['% MO / Importe']
    )
    n_main_cols = len(headers)

    ws.append(headers)
    for cell in ws[1]:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    tot_importe = 0.0
    tot_hmo    = 0.0
    tot_h      = [0.0] * len(codigos)

    for cap in presupuesto.capitulos:
        horas   = horas_capitulo(cap, codigos)
        hmo     = sum(horas.values())
        coste   = sum(horas[c] * presupuesto.recursos_mo[c].precio_hora for c in codigos)
        pct     = round(coste / cap.importe_total * 100, 1) if cap.importe_total else 0.0

        ws.append([cap.codigo, cap.descripcion, cap.importe_total, hmo]
                  + [horas[c] for c in codigos] + [pct])
        row = ws[ws.max_row]
        for cell in row:
            cell.fill = GANTT_FILL_CAP
            cell.font = GANTT_FONT_BOLD
        row[2].number_format = '#,##0.00'
        for cell in row[3:]:
            cell.number_format = '#,##0.0'

        tot_importe += cap.importe_total
        tot_hmo     += hmo
        for i, c in enumerate(codigos):
            tot_h[i] += horas[c]

        for subcap in cap.subcapitulos:
            sub_h     = horas_capitulo(subcap, codigos)
            sub_hmo   = sum(sub_h.values())
            sub_coste = sum(sub_h[c] * presupuesto.recursos_mo[c].precio_hora for c in codigos)
            sub_pct   = round(sub_coste / subcap.importe_total * 100, 1) if subcap.importe_total else 0.0

            ws.append([subcap.codigo, '  ' + subcap.descripcion, subcap.importe_total, sub_hmo]
                      + [sub_h[c] for c in codigos] + [sub_pct])
            sub_row = ws[ws.max_row]
            sub_row[2].number_format = '#,##0.00'
            for cell in sub_row[3:]:
                cell.number_format = '#,##0.0'

    # Fila TOTAL tabla principal
    tot_coste = sum(tot_h[i] * presupuesto.recursos_mo[c].precio_hora
                    for i, c in enumerate(codigos))
    tot_pct   = round(tot_coste / tot_importe * 100, 1) if tot_importe else 0.0
    ws.append(['', 'TOTAL', tot_importe, tot_hmo] + tot_h + [tot_pct])
    tot_row = ws[ws.max_row]
    for cell in tot_row:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
    tot_row[2].number_format = '#,##0.00'
    for cell in tot_row[3:]:
        cell.number_format = '#,##0.0'

    # --- SECCIÓN: RESUMEN DE MANO DE OBRA DEL PROYECTO ---
    ws.append([])
    ws.append([])
    ws.append(['RESUMEN DE MANO DE OBRA DEL PROYECTO'])
    mo_title = ws.max_row
    # Fusionado en el mismo ancho que la tabla principal
    ws.merge_cells(start_row=mo_title, start_column=1, end_row=mo_title, end_column=n_main_cols)
    ws.cell(mo_title, 1).fill = GANTT_FILL_HEADER
    ws.cell(mo_title, 1).font = GANTT_FONT_HEADER
    ws.cell(mo_title, 1).alignment = Alignment(horizontal='center')

    ws.append(['Código', 'Descripción', '€/hora', 'Horas totales', 'Coste total (€)', '% coste MO'])
    mo_hdr = ws[ws.max_row]
    for cell in mo_hdr:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    for i, recurso in enumerate(recursos):
        horas_r = tot_h[i]
        coste_r = horas_r * recurso.precio_hora
        pct_r   = round(coste_r / tot_coste * 100, 1) if tot_coste else 0.0
        fill    = GANTT_FILL_ALT if i % 2 == 0 else GANTT_FILL_WHITE
        ws.append([recurso.codigo, recurso.descripcion, recurso.precio_hora, horas_r, coste_r, pct_r])
        mo_row = ws[ws.max_row]
        for cell in mo_row:
            cell.fill = fill
        mo_row[2].number_format = '#,##0.00'
        mo_row[3].number_format = '#,##0.00'
        mo_row[4].number_format = '#,##0.00'
        mo_row[5].number_format = '0.0'

    ws.append(['', 'TOTAL', '', sum(tot_h), tot_coste, 100.0])
    mo_tot = ws[ws.max_row]
    for cell in mo_tot:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
    mo_tot[3].number_format = '#,##0.00'
    mo_tot[4].number_format = '#,##0.00'
    mo_tot[5].number_format = '0.0'

    adjust_col_widths(ws)


def write_partidas(ws, presupuesto: Presupuesto, recursos: list[RecursoMO]) -> None:
    """
    Escribe la hoja 'Partidas detalladas' con una fila por partida y desglose MO.
    Filas alternas en blanco y gris. Cabecera fija con freeze_panes.
    """
    codigos = [r.codigo for r in recursos]
    headers = (
        ['Capítulo', 'Subcapítulo', 'Código partida', 'Descripción', 'Unidad',
         'Cantidad', 'Precio unitario (€)', 'Importe (€)', 'H MO total']
        + [r.descripcion for r in recursos]
    )
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'A2'

    row_idx = 2
    for cap in presupuesto.capitulos:
        for cod_cap, cod_subcap, partida in iter_partidas(cap, cap.codigo):
            horas_por_cod = {c: 0.0 for c in codigos}
            for linea in partida.lineas_mo:
                if linea.codigo_recurso in horas_por_cod:
                    horas_por_cod[linea.codigo_recurso] = linea.cantidad * partida.cantidad
            hmo_total = sum(horas_por_cod.values())
            importe   = partida.precio_unitario * partida.cantidad

            ws.append([cod_cap, cod_subcap, partida.codigo, partida.descripcion,
                       partida.unidad, partida.cantidad, partida.precio_unitario,
                       importe, hmo_total]
                      + [horas_por_cod[c] for c in codigos])

            row  = ws[row_idx]
            fill = GANTT_FILL_ALT if row_idx % 2 == 0 else GANTT_FILL_WHITE
            for cell in row:
                cell.fill = fill
            for cell in row[5:]:
                cell.number_format = '#,##0.00'
            row_idx += 1

    adjust_col_widths(ws)


def write_simulador(
    ws, presupuesto: Presupuesto, planificaciones: list[PlanificacionProyecto]
) -> None:
    """
    Escribe la hoja 'Simulador de cuadrillas' con dos secciones:
    (A) parámetros de cuadrilla por escenario y (B) tabla comparativa de duraciones.
    """
    recursos = list(presupuesto.recursos_mo.values())
    n_escenarios = len(planificaciones)
    # Ancho sección B: ID + Nombre + 3 cols/escenario + Perfil limitante
    n_cols_b = 2 + 3 * n_escenarios + 1

    # ── SECCIÓN A: PARÁMETROS POR ESCENARIO ──────────────────────────────────
    ws.append(['PARÁMETROS DE CUADRILLA POR ESCENARIO'])
    title_a = ws.max_row
    ws.merge_cells(start_row=title_a, start_column=1, end_row=title_a, end_column=n_cols_b)
    ws.cell(title_a, 1).fill = GANTT_FILL_HEADER
    ws.cell(title_a, 1).font = GANTT_FONT_HEADER
    ws.cell(title_a, 1).alignment = Alignment(horizontal='center')

    for plan_idx, plan in enumerate(planificaciones):
        if plan_idx > 0:
            ws.append([])

        # Título del escenario
        ws.append([plan.escenario.nombre])
        scen_row = ws.max_row
        ws.merge_cells(start_row=scen_row, start_column=1, end_row=scen_row, end_column=5)
        ws.cell(scen_row, 1).fill = GANTT_FILL_CAP
        ws.cell(scen_row, 1).font = GANTT_FONT_BOLD

        # Cabecera
        ws.append(['Código recurso', 'Descripción', '€/hora', 'Operarios asignados', 'Horas/día'])
        for cell in ws[ws.max_row]:
            cell.fill = GANTT_FILL_HEADER
            cell.font = GANTT_FONT_HEADER
            cell.alignment = Alignment(horizontal='center')

        # Filas de recursos
        for i, recurso in enumerate(recursos):
            ops  = plan.escenario.operarios_por_recurso.get(recurso.codigo, 1)
            fill = GANTT_FILL_ALT if i % 2 == 0 else GANTT_FILL_WHITE
            ws.append([recurso.codigo, recurso.descripcion, recurso.precio_hora,
                       ops, plan.parametros.horas_dia])
            row = ws[ws.max_row]
            for cell in row:
                cell.fill = fill
            row[2].number_format = '#,##0.00'
            row[3].fill = GANTT_FILL_YELLOW  # Operarios: editable en el yaml

    # ── SECCIÓN B: DURACIONES POR ESCENARIO ──────────────────────────────────
    ws.append([])
    ws.append([])
    ws.append(['DURACIONES CALCULADAS POR ESCENARIO'])
    title_b = ws.max_row
    ws.merge_cells(start_row=title_b, start_column=1, end_row=title_b, end_column=n_cols_b)
    ws.cell(title_b, 1).fill = GANTT_FILL_HEADER
    ws.cell(title_b, 1).font = GANTT_FONT_HEADER
    ws.cell(title_b, 1).alignment = Alignment(horizontal='center')

    # Fila de cabecera 1: nombres de escenarios (merged 3 cols cada uno)
    hdr1_values = ['ID', 'Nombre']
    for plan in planificaciones:
        hdr1_values += [plan.escenario.nombre, '', '']
    hdr1_values += ['Perfil limitante']
    ws.append(hdr1_values)
    hdr1_row = ws.max_row
    for i in range(n_escenarios):
        col = 3 + i * 3
        ws.merge_cells(start_row=hdr1_row, start_column=col, end_row=hdr1_row, end_column=col + 2)
        ws.cell(hdr1_row, col).alignment = Alignment(horizontal='center')
    for cell in ws[hdr1_row]:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER

    # Fila de cabecera 2: Días / Inicio / Fin por escenario
    ws.append(['', ''] + ['Días', 'Inicio', 'Fin'] * n_escenarios + [''])
    for cell in ws[ws.max_row]:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    # Lookup por ID para cada planificación
    by_id_per_plan = [{t.id: t for t in plan.tareas} for plan in planificaciones]
    base_tareas    = planificaciones[0].tareas

    for tarea_base in base_tareas:
        row_data = [tarea_base.id, tarea_base.nombre]
        for plan_dict in by_id_per_plan:
            t     = plan_dict.get(tarea_base.id, tarea_base)
            dias  = t.duracion_dias if t.duracion_dias is not None else 0.0
            inicio = t.fecha_inicio.strftime('%d/%m/%Y') if t.fecha_inicio else ''
            fin    = t.fecha_fin.strftime('%d/%m/%Y')    if t.fecha_fin   else ''
            row_data += [dias, inicio, fin]
        row_data += [tarea_base.perfil_limitante or '']

        ws.append(row_data)
        row = ws[ws.max_row]

        if tarea_base.tipo == 'hito':
            for cell in row:
                cell.fill = GANTT_FILL_YELLOW
        else:
            for cell in row:
                cell.fill = GANTT_FILL_CAP
            row[0].font = GANTT_FONT_BOLD
            row[1].font = GANTT_FONT_BOLD

        # Formato numérico en columnas Días
        for i in range(n_escenarios):
            row[2 + i * 3].number_format = '#,##0.0'

    # Fila DURACIÓN TOTAL hasta el hito de fin de plazo contractual
    tot_data  = ['', 'DURACIÓN TOTAL']
    tot_fills = []
    for plan in planificaciones:
        hito_fin = next((t for t in plan.tareas if t.es_fin_plazo), None)
        if hito_fin and hito_fin.fecha_fin:
            dias_tot = dias_habiles_entre(plan.parametros.fecha_inicio, hito_fin.fecha_fin)
            cumple = dias_tot <= plan.parametros.plazo_contractual_dias
            label  = f'{"CUMPLE" if cumple else "NO CUMPLE"} ({dias_tot}d)'
            tot_fills.append(GANTT_FILL_GREEN if cumple else GANTT_FILL_RED)
        else:
            label = '—'
            tot_fills.append(GANTT_FILL_WHITE)
        tot_data += [label, '', '']
    tot_data += ['']

    ws.append(tot_data)
    tot_row = ws[ws.max_row]
    tot_row[0].font = GANTT_FONT_BOLD
    tot_row[1].font = GANTT_FONT_BOLD
    for i, fill in enumerate(tot_fills):
        for j in range(3):
            tot_row[2 + i * 3 + j].fill = fill
            tot_row[2 + i * 3 + j].font = GANTT_FONT_BOLD

    adjust_col_widths(ws)


def write_ms_project_sheet(
    ws, planificacion: PlanificacionProyecto, presupuesto: Presupuesto
) -> None:
    """
    Escribe una hoja de importación MS Project para un escenario dado.
    Columnas: ID, Name, Duration, Start, Finish, Predecessors,
    Resource Names, Outline Level, Milestone, Notes.
    """
    headers = ['ID', 'Name', 'Duration', 'Start', 'Finish', 'Predecessors',
               'Resource Names', 'Outline Level', 'Milestone', 'Notes']
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'A2'

    # Tabla inversa: descripción de recurso → código, para obtener operarios
    desc_to_codigo = {r.descripcion: c for c, r in presupuesto.recursos_mo.items()}

    # Mapa ID_tarea → número de fila para construir predecesoras
    id_map = {t.id: idx + 1 for idx, t in enumerate(planificacion.tareas)}

    task_id = 1
    for i, tarea in enumerate(planificacion.tareas):
        es_hito         = tarea.tipo == 'hito'
        name            = ('  ' + tarea.nombre) if es_hito else tarea.nombre
        dias_redondeados = round(tarea.duracion_dias or 0, 1)
        duration        = '0d' if (es_hito or dias_redondeados == 0) else f'{dias_redondeados:.1f}d'
        start           = tarea.fecha_inicio.strftime('%d/%m/%Y') if tarea.fecha_inicio else ''
        finish          = tarea.fecha_fin.strftime('%d/%m/%Y')    if tarea.fecha_fin   else ''
        predecessors    = ','.join(
            str(id_map[dep]) for dep in tarea.dependencias if dep in id_map
        )

        # Recurso limitante con operarios asignados en este escenario
        resource_names = ''
        if tarea.perfil_limitante and tarea.perfil_limitante != 'estacional' and tarea.horas_mo_total:
            codigo_lim = desc_to_codigo.get(tarea.perfil_limitante)
            if codigo_lim:
                n_ops = planificacion.escenario.operarios_por_recurso.get(codigo_lim, 1)
                resource_names = f'{tarea.perfil_limitante}[{n_ops}]'

        outline_level = 1  # todos al mismo nivel jerárquico
        milestone     = 'YES' if es_hito else 'NO'

        if tarea.capitulos_bc3 and tarea.horas_mo_total:
            notes = f'BC3: {", ".join(tarea.capitulos_bc3)} | H MO: {tarea.horas_mo_total:.1f}h'
        elif tarea.capitulos_bc3:
            notes = f'BC3: {", ".join(tarea.capitulos_bc3)}'
        else:
            notes = ''

        ws.append([task_id, name, duration, start, finish, predecessors,
                   resource_names, outline_level, milestone, notes])

        row = ws[ws.max_row]
        if es_hito:
            for cell in row:
                cell.fill = GANTT_FILL_YELLOW
                cell.font = GANTT_FONT_ITALIC
        else:
            fill = GANTT_FILL_ALT if i % 2 == 0 else GANTT_FILL_WHITE
            for cell in row:
                cell.fill = fill

        task_id += 1

    adjust_col_widths(ws)
    ws.column_dimensions['B'].width = max(ws.column_dimensions['B'].width, 60)


def write_carga_recursos(
    ws, presupuesto: Presupuesto, planificacion: PlanificacionProyecto
) -> None:
    """
    Escribe la hoja 'Carga recursos': histograma semanal de horas de mano de
    obra por perfil (Sección A, pivote semana x recurso) y detalle de las
    tareas que aportan horas a cada semana/recurso (Sección B).
    """
    filas = calcular_carga_semanal_recursos(planificacion, presupuesto)

    if not filas:
        ws.append(['Sin horas de mano de obra planificadas para mostrar.'])
        return

    recursos_codigos = sorted({f['recurso_codigo'] for f in filas})
    recursos_nombres = {f['recurso_codigo']: f['recurso_nombre'] for f in filas}
    semanas = sorted({(f['semana_inicio'], f['semana_iso']) for f in filas})
    valores = {(f['semana_inicio'], f['recurso_codigo']): f['horas_semana'] for f in filas}

    # ── SECCIÓN A: HISTOGRAMA SEMANAL (pivote semana x recurso) ──────────────
    n_cols = 2 + len(recursos_codigos) + 1
    ws.append(['HISTOGRAMA SEMANAL DE CARGA DE RECURSOS'])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    ws.cell(r, 1).fill = GANTT_FILL_HEADER
    ws.cell(r, 1).font = GANTT_FONT_HEADER
    ws.cell(r, 1).alignment = Alignment(horizontal='center')

    headers = (
        ['Semana (lunes)', 'Semana ISO']
        + [recursos_nombres[codigo] for codigo in recursos_codigos]
        + ['Total semana']
    )
    ws.append(headers)
    for cell in ws[ws.max_row]:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    for i, (semana_inicio, semana_iso_str) in enumerate(semanas):
        horas_por_recurso = [valores.get((semana_inicio, c), 0.0) for c in recursos_codigos]
        ws.append(
            [semana_inicio.strftime('%d/%m/%Y'), semana_iso_str]
            + horas_por_recurso
            + [sum(horas_por_recurso)]
        )
        row = ws[ws.max_row]
        fill = GANTT_FILL_ALT if i % 2 == 0 else GANTT_FILL_WHITE
        for cell in row:
            cell.fill = fill
        for cell in row[2:]:
            cell.number_format = '#,##0.0'

    # ── SECCIÓN B: DETALLE POR TAREA, SEMANA Y RECURSO ────────────────────────
    ws.append([])
    ws.append([])
    ws.append(['DETALLE DE TAREAS POR SEMANA Y RECURSO'])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(r, 1).fill = GANTT_FILL_HEADER
    ws.cell(r, 1).font = GANTT_FONT_HEADER
    ws.cell(r, 1).alignment = Alignment(horizontal='center')

    ws.append(['Semana ISO', 'Recurso', 'Descripción', 'Horas semana', 'Tareas incluidas'])
    for cell in ws[ws.max_row]:
        cell.fill = GANTT_FILL_HEADER
        cell.font = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    for i, fila in enumerate(filas):
        ws.append([
            fila['semana_iso'],
            fila['recurso_codigo'],
            fila['recurso_nombre'],
            fila['horas_semana'],
            ', '.join(fila['tareas_incluidas']),
        ])
        row = ws[ws.max_row]
        fill = GANTT_FILL_ALT if i % 2 == 0 else GANTT_FILL_WHITE
        for cell in row:
            cell.fill = fill
        row[3].number_format = '#,##0.0'

    adjust_col_widths(ws)


def write_dashboard(
    ws,
    presupuesto: Presupuesto,
    planificaciones: list[PlanificacionProyecto],
    palette: DocumentPalette | None = None,
) -> None:
    """
    Escribe la hoja Dashboard: resumen ejecutivo, diagrama de red embebido,
    análisis de sensibilidad y recomendación automática de operarios.
    """
    from gantt.reporting.network_diagram import generar_diagrama_red

    params = planificaciones[0].parametros

    # ── SECCIÓN A: RESUMEN EJECUTIVO ──────────────────────────────────────────
    ws.append(['RESUMEN EJECUTIVO DEL PROYECTO'])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(r, 1).fill      = GANTT_FILL_HEADER
    ws.cell(r, 1).font      = GANTT_FONT_HEADER
    ws.cell(r, 1).alignment = Alignment(horizontal='center')

    for label, value in [
        ('Proyecto:',               params.nombre),
        ('Fecha inicio obra:',      params.fecha_inicio.strftime('%d/%m/%Y')),
        ('Horas dia / Dias semana:', f'{params.horas_dia}h / {params.dias_semana}d'),
        ('Plazo contractual:',       f'{params.plazo_contractual_dias} dias habiles'),
        ('', ''),
    ]:
        ws.append([label, value])
        ws[ws.max_row][0].font = GANTT_FONT_BOLD

    for plan in planificaciones:
        hito_fin = next((t for t in plan.tareas if t.es_fin_plazo), None)
        if hito_fin and hito_fin.fecha_fin:
            dias   = dias_habiles_entre(params.fecha_inicio, hito_fin.fecha_fin)
            cumple = dias <= params.plazo_contractual_dias
            ws.append([
                f'Escenario {plan.escenario.nombre}:',
                f'{dias}d habiles - {"CUMPLE" if cumple else "NO CUMPLE"}',
            ])
            row = ws[ws.max_row]
            row[0].font = GANTT_FONT_BOLD
            row[1].fill = GANTT_FILL_GREEN if cumple else GANTT_FILL_RED

    # ── SECCIÓN B: DIAGRAMA DE RED ────────────────────────────────────────────
    ws.append([])
    ws.append(['DIAGRAMA DE RED - DEPENDENCIAS Y CAMINO CRITICO'])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(r, 1).fill      = GANTT_FILL_HEADER
    ws.cell(r, 1).font      = GANTT_FONT_HEADER
    ws.cell(r, 1).alignment = Alignment(horizontal='center')

    image_row  = ws.max_row + 1
    png_bytes  = generar_diagrama_red(planificaciones[-1], palette=palette)
    img        = XLImage(io.BytesIO(png_bytes))
    img.width  = 1400
    img.height = 600
    ws.add_image(img, f'A{image_row}')
    for _ in range(40):   # reservar filas para la imagen
        ws.append([])

    # ── SECCIONES C, D, E: TABLAS ANALÍTICAS ─────────────────────────────────
    ws.append([])
    ws.append([])
    plan_analisis = planificaciones[-1]
    horas_dia     = plan_analisis.parametros.horas_dia
    holguras      = calcular_holguras(plan_analisis.tareas)

    # Identificar nodos post-plazo via BFS en el grafo de sucesores
    hito_fin = next((t for t in plan_analisis.tareas if t.es_fin_plazo), None)
    sucesores_map: dict[str, list[str]] = {t.id: [] for t in plan_analisis.tareas}
    for t in plan_analisis.tareas:
        for dep in t.dependencias:
            if dep in sucesores_map:
                sucesores_map[dep].append(t.id)

    nodos_post_plazo: set[str] = set()
    if hito_fin:
        cola = list(sucesores_map.get(hito_fin.id, []))
        visitados: set[str] = set()
        while cola:
            nid = cola.pop()
            if nid in visitados:
                continue
            visitados.add(nid)
            nodos_post_plazo.add(nid)
            cola.extend(sucesores_map.get(nid, []))

    tareas_camino_critico = sorted(
        [
            t for t in plan_analisis.tareas
            if holguras.get(t.id, 1) <= 0
            and not t.es_fin_plazo
            and t.id not in nodos_post_plazo
            and t.tipo == 'tarea'
            and t.horas_por_recurso
        ],
        key=lambda t: t.fecha_inicio or date.min,
    )

    ws.append([f'Escenario analizado: {plan_analisis.escenario.nombre}'])
    ws[ws.max_row][0].font = Font(italic=True)

    # ── TABLA 1: CAMINO CRÍTICO ───────────────────────────────────────────────
    ws.append(['CAMINO CRITICO - TAREAS CON HOLGURA CERO'])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.cell(r, 1).fill      = GANTT_FILL_HEADER
    ws.cell(r, 1).font      = GANTT_FONT_HEADER
    ws.cell(r, 1).alignment = Alignment(horizontal='center')

    ws.append(['ID', 'Nombre', 'Dias', 'Recurso limitante', 'H limitante', 'Op', 'Holgura'])
    for cell in ws[ws.max_row]:
        cell.fill      = GANTT_FILL_HEADER
        cell.font      = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    for i, tarea in enumerate(tareas_camino_critico):
        if tarea.horas_por_recurso:
            dias_por_rec = {
                c: h / (plan_analisis.escenario.operarios_por_recurso.get(c, 1) * horas_dia)
                for c, h in tarea.horas_por_recurso.items() if h > 0
            }
            cod_lim    = max(dias_por_rec, key=lambda c: dias_por_rec[c])
            rec_lim    = presupuesto.recursos_mo.get(cod_lim)
            nombre_lim = rec_lim.descripcion if rec_lim else cod_lim
            h_lim      = round(tarea.horas_por_recurso[cod_lim], 1)
            ops_lim    = plan_analisis.escenario.operarios_por_recurso.get(cod_lim, 1)
        else:
            nombre_lim = 'Duracion fija'
            h_lim      = '—'
            ops_lim    = '—'

        holgura = round(holguras.get(tarea.id, 0), 1)
        ws.append([
            tarea.id, tarea.nombre,
            round(tarea.duracion_dias or 0, 1),
            nombre_lim, h_lim, ops_lim, holgura,
        ])
        row = ws[ws.max_row]
        for cell in row:
            cell.fill = GANTT_FILL_ALT if i % 2 == 0 else GANTT_FILL_WHITE
        row[6].fill = GANTT_FILL_RED

    # ── TABLA 2: ANÁLISIS DE SENSIBILIDAD ─────────────────────────────────────
    ws.append([])
    ws.append([])
    ws.append(['ANALISIS DE SENSIBILIDAD - IMPACTO DE +1 OPERARIO POR TAREA'])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    ws.cell(r, 1).fill      = GANTT_FILL_HEADER
    ws.cell(r, 1).font      = GANTT_FONT_HEADER
    ws.cell(r, 1).alignment = Alignment(horizontal='center')

    ws.append([f'Escenario analizado: {plan_analisis.escenario.nombre}'])
    ws[ws.max_row][0].font = Font(italic=True)

    ws.append([
        'Tarea', 'Nombre tarea', 'Recurso MO', 'Descripcion',
        'H tarea', 'Op', 'Dias tarea', '+1op dias',
        'Red. proyecto (d)', 'Impacto',
    ])
    for cell in ws[ws.max_row]:
        cell.fill      = GANTT_FILL_HEADER
        cell.font      = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    sensibilidad = calcular_sensibilidad(plan_analisis, presupuesto)

    # Colores alternos por bloque de tarea
    tarea_ids_orden: list[str] = []
    for item in sensibilidad:
        if item['tarea_id'] not in tarea_ids_orden:
            tarea_ids_orden.append(item['tarea_id'])
    color_por_tarea = {
        tid: (GANTT_FILL_ALT if idx % 2 == 0 else GANTT_FILL_WHITE)
        for idx, tid in enumerate(tarea_ids_orden)
    }

    for item in sensibilidad:
        fill = color_por_tarea[item['tarea_id']]
        red  = round(item['reduccion_proyecto'], 1)
        if item['es_limitante']:
            limitante_str = 'directo'
        elif item['reduccion_proyecto'] > 0:
            limitante_str = 'indirecto'
        else:
            limitante_str = ''

        ws.append([
            item['tarea_id'],
            item['tarea_nombre'][:40],
            item['codigo'],
            item['descripcion'],
            round(item['horas'], 1),
            item['operarios_actual'],
            round(item['dias_tarea_actual'], 1),
            round(item['dias_tarea_nuevo'], 1),
            red,
            limitante_str,
        ])
        row = ws[ws.max_row]
        for cell in row:
            cell.fill = fill
        if red > 5:
            row[8].fill = GANTT_FILL_GREEN
        elif red > 0:
            row[8].fill = GANTT_FILL_YELLOW
        if item['es_limitante']:
            for cell in row:
                cell.font = GANTT_FONT_BOLD

    # ── TABLA 3: CONFIGURACIÓN ÓPTIMA DE RECURSOS ─────────────────────────────
    ws.append([])
    ws.append([])
    ws.append(['CONFIGURACION OPTIMA DE RECURSOS'])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.cell(r, 1).fill      = GANTT_FILL_HEADER
    ws.cell(r, 1).font      = GANTT_FONT_HEADER
    ws.cell(r, 1).alignment = Alignment(horizontal='center')

    ws.append([f'Objetivo: cumplir {params.plazo_contractual_dias} dias habiles'])
    ws[ws.max_row][0].font = Font(italic=True)

    ajustes, dur_final = calcular_recomendacion(
        plan_analisis, presupuesto, params.plazo_contractual_dias
    )

    # Construir progresión de duración por perfil
    operarios_optimos  = dict(plan_analisis.escenario.operarios_por_recurso)
    progresion_por_perfil: dict[str, list[int]] = {}
    dur_progresion = duracion_total_proyecto(plan_analisis)

    for ajuste in ajustes:
        codigo = ajuste['codigo']
        operarios_optimos[codigo] = ajuste['operarios_despues']
        dur_nueva = dur_progresion - ajuste['impacto_dias']
        if codigo not in progresion_por_perfil:
            progresion_por_perfil[codigo] = [int(dur_progresion)]
        progresion_por_perfil[codigo].append(int(dur_nueva))
        dur_progresion = dur_nueva

    sobredim     = calcular_sobredimensionados(plan_analisis, presupuesto, operarios_optimos)
    sobredim_map = {s['codigo']: s for s in sobredim}
    ajustes_map  = {a['codigo']: a for a in ajustes}

    ws.append(['Perfil MO', 'Descripcion', 'Actual', 'Optimo', 'Ajuste', 'Progresion duracion', 'Tipo'])
    for cell in ws[ws.max_row]:
        cell.fill      = GANTT_FILL_HEADER
        cell.font      = GANTT_FONT_HEADER
        cell.alignment = Alignment(horizontal='center')

    for i, codigo in enumerate(plan_analisis.escenario.operarios_por_recurso):
        recurso    = presupuesto.recursos_mo.get(codigo)
        ops_actual = plan_analisis.escenario.operarios_por_recurso[codigo]

        if codigo in ajustes_map:
            ajuste         = ajustes_map[codigo]
            ops_optimo     = ajuste['operarios_despues']
            prog           = progresion_por_perfil.get(codigo, [])
            progresion_str = ' -> '.join(f'{d}d' for d in prog)
            tipo           = 'Incremento necesario'
            fill_tipo      = GANTT_FILL_RED
        elif codigo in sobredim_map:
            ops_optimo     = sobredim_map[codigo]['operarios_optimo']
            progresion_str = 'Sin impacto en plazo'
            tipo           = 'Reduccion posible'
            fill_tipo      = GANTT_FILL_GREEN
        else:
            ops_optimo     = ops_actual
            progresion_str = '-'
            tipo           = 'Optimo'
            fill_tipo      = GANTT_FILL_WHITE

        delta      = ops_optimo - ops_actual
        ajuste_str = f'+{delta}' if delta > 0 else str(delta) if delta < 0 else '='

        ws.append([
            codigo, recurso.descripcion if recurso else codigo,
            ops_actual, ops_optimo, ajuste_str, progresion_str, tipo,
        ])
        row = ws[ws.max_row]
        for cell in row:
            cell.fill = GANTT_FILL_ALT if i % 2 == 0 else GANTT_FILL_WHITE
        row[6].fill = fill_tipo

    # Fila resultado final
    dur_r        = math.ceil(dur_final)
    cumple_final = dur_r <= params.plazo_contractual_dias
    ws.append([
        f'Duracion estimada: {dur_r}d habiles - '
        f'{"CUMPLE" if cumple_final else "NO CUMPLE"} '
        f'(plazo: {params.plazo_contractual_dias}d)',
        '', '', '', '', '', '',
    ])
    fila_res = ws[ws.max_row]
    ws.merge_cells(
        start_row=fila_res[0].row, start_column=1,
        end_row=fila_res[0].row,   end_column=7,
    )
    for cell in fila_res:
        cell.fill = GANTT_FILL_GREEN if cumple_final else GANTT_FILL_RED
        cell.font = GANTT_FONT_BOLD

    ws.append(['Validar disponibilidad de recursos antes de aplicar ajustes', '', '', '', '', '', ''])
    fila_av = ws[ws.max_row]
    ws.merge_cells(
        start_row=fila_av[0].row, start_column=1,
        end_row=fila_av[0].row,   end_column=7,
    )
    for cell in fila_av:
        cell.fill = GANTT_FILL_YELLOW
        cell.font = GANTT_FONT_ITALIC

    # Anchos de columna ajustados a 10 columnas (Tabla 2 es la más ancha)
    for col, w in [('A', 12), ('B', 35), ('C', 18), ('D', 30), ('E', 10), ('F', 6), ('G', 12), ('H', 12), ('I', 16), ('J', 10)]:
        ws.column_dimensions[col].width = w


def exportar_analisis(
    presupuesto: Presupuesto,
    output_path: Path,
    planificaciones: list[PlanificacionProyecto] | None = None,
    palette: DocumentPalette | None = None,
    metadata: DocumentMetadata | None = None,
    presentation: PresentationContext | None = None,
) -> None:
    """
    Genera el Excel de análisis del presupuesto BC3 y lo escribe en output_path.

    Parámetros
    ----------
    presupuesto : Presupuesto
        Objeto presupuesto parseado desde el BC3.
    output_path : Path
        Ruta del archivo .xlsx a generar o sobreescribir.
    planificaciones : list[PlanificacionProyecto] | None
        Lista de planificaciones calculadas (una por escenario). Si se proporciona,
        se generan también la hoja 'Simulador de cuadrillas', 'Carga recursos'
        y una hoja 'MSProject_{escenario}' por cada escenario.
    """
    pal = palette or (presentation.palette if presentation else None) or build_palette(presupuesto.company)
    presentation_context = presentation or presentation_from_palette_company(pal, presupuesto.company)
    document_type = DocumentType.PLANNING_ANALYSIS if planificaciones else DocumentType.BUDGET_ANALYSIS
    metadata = metadata or metadata_from_project_config(
        presupuesto.config,
        document_type,
        title="Planificacion y analisis de escenarios" if planificaciones else "Analisis economico del presupuesto",
    )
    if planificaciones and not metadata.planning:
        metadata = metadata.model_copy(
            update={
                "planning": PlanningMetadata(
                    scenario=planificaciones[-1].escenario.nombre,
                    contractual_deadline=planificaciones[-1].parametros.plazo_contractual_dias,
                )
            }
        )
    _apply_palette_to_legacy_styles(pal)

    recursos = list(presupuesto.recursos_mo.values())
    wb = Workbook()

    ws1 = wb.active
    ws1.title = 'Resumen por capítulo'
    write_resumen(ws1, presupuesto, recursos)

    ws2 = wb.create_sheet('Partidas detalladas')
    write_partidas(ws2, presupuesto, recursos)

    if planificaciones:
        ws3 = wb.create_sheet('Simulador de cuadrillas')
        write_simulador(ws3, presupuesto, planificaciones)

        ws_carga = wb.create_sheet('Carga recursos')
        write_carga_recursos(ws_carga, presupuesto, planificaciones[-1])

        for plan in planificaciones:
            sheet_name = f'MSProject_{plan.escenario.nombre}'
            write_ms_project_sheet(wb.create_sheet(sheet_name), plan, presupuesto)

        # Dashboard como primera hoja (índice 0, se inserta delante de todo)
        write_dashboard(wb.create_sheet('Dashboard', 0), presupuesto, planificaciones, palette=pal)

    header_rows = 5
    first_content_row = header_rows + 1
    first_scrollable_row = first_content_row + 1
    for ws in wb.worksheets:
        ws.insert_rows(1, amount=header_rows)
        n_cols = max(ws.max_column, 6)
        if planificaciones:
            render_planning_header(ws, metadata, presentation_context, n_cols)
        else:
            render_analysis_header(ws, metadata, presentation_context, n_cols)
        data_last_row = ws.max_row
        filter_row = first_filterable_row(ws, first_content_row)
        if ws.max_row >= filter_row + 1:
            ws.freeze_panes = f"A{filter_row + 1}"
            ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(n_cols)}{data_last_row}"
        apply_document_footer(ws, metadata, presentation_context, n_cols)
        configure_excel_printing(
            ws,
            orientation="landscape" if n_cols > 6 or planificaciones else "portrait",
            repeat_row=filter_row if ws.max_row > filter_row else None,
            purpose=DocumentPurpose.ANALYSIS,
        )

    wb.save(output_path)



if __name__ == '__main__':
    import sys
    from pathlib import Path
    from gantt.bc3.parser import parse_bc3
    from gantt.planning.models import cargar_planificacion_yaml
    from gantt.planning.calculator import calcular_planificacion

    project_name = sys.argv[1] if len(sys.argv) > 1 else 'Complexe-Balear'
    bc3_filename  = sys.argv[2] if len(sys.argv) > 2 else None
    input_dir  = Path('projects') / project_name / 'input'
    output_dir = Path('projects') / project_name / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)

    if bc3_filename:
        bc3_path = input_dir / bc3_filename
        if not bc3_path.exists():
            raise FileNotFoundError(f'No se encontró: {bc3_path}')
    else:
        bc3_files = list(input_dir.glob('*.bc3'))
        if not bc3_files:
            raise FileNotFoundError(f'No se encontró ningún .bc3 en {input_dir}')
        if len(bc3_files) > 1:
            raise ValueError(
                f'Multiples .bc3 en {input_dir}. '
                f'Especifica el archivo como segundo argumento.\n'
                f'Disponibles: {[f.name for f in bc3_files]}'
            )
        bc3_path = bc3_files[0]

    presupuesto = parse_bc3(bc3_path)

    planificaciones = None
    yaml_path = input_dir / 'planificacion.yaml'
    if yaml_path.exists():
        parametros, escenarios, tareas, bandas = cargar_planificacion_yaml(yaml_path)
        planificaciones = [
            calcular_planificacion(presupuesto, parametros, tareas, escenario, bandas)
            for escenario in escenarios
        ]
        print(f'Planificaciones calculadas: {[p.escenario.nombre for p in planificaciones]}')
    else:
        print('planificacion.yaml no encontrado — solo hojas 1 y 2')

    output_path = output_dir / f'{project_name}_analisis.xlsx'
    exportar_analisis(presupuesto, output_path, planificaciones)
    print(f'Excel generado: {output_path}')

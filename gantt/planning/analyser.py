"""
Módulo de análisis temporal del proyecto.

Responsabilidad: calcular holguras, camino crítico, sensibilidad
de recursos y recomendaciones de optimización a partir de una
PlanificacionProyecto ya calculada.
"""

import math
from collections import deque
from datetime import date, timedelta

from gantt.bc3.models import Presupuesto
from gantt.planning.calculator import calcular_planificacion
from gantt.planning.models import EscenarioRecursos, PlanificacionProyecto, TareaGantt


def dias_habiles_entre(d1: date, d2: date) -> int:
    """
    Cuenta los días hábiles entre dos fechas.
    Retorna valor negativo si d2 < d1.
    """
    if d1 is None or d2 is None:
        return 0
    total = 0
    d = min(d1, d2)
    fin = max(d1, d2)
    while d < fin:
        if d.weekday() < 5:
            total += 1
        d += timedelta(days=1)
    return total if d2 >= d1 else -total


def calcular_holguras(tareas: list[TareaGantt]) -> dict[str, float]:
    """
    Calcula la holgura de cada tarea mediante backward pass correcto.
    El backward pass se ancla al hito con es_fin_plazo=True.
    Tarea crítica = holgura <= 0.
    """
    if not tareas:
        return {}

    sucesores: dict[str, list[str]] = {t.id: [] for t in tareas}
    for t in tareas:
        for dep in t.dependencias:
            if dep in sucesores:
                sucesores[dep].append(t.id)

    # Orden topológico con algoritmo de Kahn
    in_degree = {t.id: len(t.dependencias) for t in tareas}
    queue = deque([t.id for t in tareas if in_degree[t.id] == 0])
    orden_topologico = []
    while queue:
        nid = queue.popleft()
        orden_topologico.append(nid)
        for suc in sucesores[nid]:
            in_degree[suc] -= 1
            if in_degree[suc] == 0:
                queue.append(suc)

    nodos_finales = [tid for tid in orden_topologico if not sucesores[tid]]

    # Hito que marca el fin del plazo contractual — ancla el backward pass
    hito_fin = next((t for t in tareas if t.es_fin_plazo), None)

    fecha_proyecto = next(
        (t.fecha_inicio for t in tareas if t.fecha_inicio is not None), None
    )

    inicio_rel: dict[str, int] = {}
    fin_rel: dict[str, int] = {}
    for t in tareas:
        if t.fecha_inicio and t.fecha_fin:
            inicio_rel[t.id] = dias_habiles_entre(fecha_proyecto, t.fecha_inicio)
            fin_rel[t.id] = dias_habiles_entre(fecha_proyecto, t.fecha_fin)

    # max_fin anclado al hito de fin de plazo; fallback al máximo de terminales
    if hito_fin and hito_fin.id in fin_rel:
        max_fin = fin_rel[hito_fin.id]
    else:
        max_fin = max((fin_rel.get(nid, 0) for nid in nodos_finales), default=0)

    # Backward pass en orden topológico inverso
    fin_tard_rel: dict[str, int] = {}
    ini_tard_rel: dict[str, int] = {}

    # Inicializar solo el hito de fin de plazo; los terminales fuera de
    # la cadena contractual se anclan a su propia fecha fin
    if hito_fin and hito_fin.id in fin_rel:
        nid = hito_fin.id
        fin_tard_rel[nid] = max_fin
        ini_tard_rel[nid] = max_fin - (fin_rel.get(nid, 0) - inicio_rel.get(nid, 0))
    else:
        for nid in nodos_finales:
            fin_tard_rel[nid] = max_fin
            ini_tard_rel[nid] = max_fin - (fin_rel.get(nid, 0) - inicio_rel.get(nid, 0))

    for tid in reversed(orden_topologico):
        if tid in fin_tard_rel:
            continue
        sucs_calculados = [s for s in sucesores[tid] if s in ini_tard_rel]
        if sucs_calculados:
            fin_tard_rel[tid] = min(ini_tard_rel[s] for s in sucs_calculados)
        else:
            # Terminal fuera de la cadena contractual: sin holgura respecto a sí mismo
            fin_tard_rel[tid] = fin_rel.get(tid, max_fin)
        ini_tard_rel[tid] = fin_tard_rel[tid] - (fin_rel.get(tid, 0) - inicio_rel.get(tid, 0))

    # Holgura = inicio_tardio - inicio_temprano
    return {
        t.id: ini_tard_rel.get(t.id, 0) - inicio_rel.get(t.id, 0)
        for t in tareas
    }


def duracion_total_proyecto(planificacion: PlanificacionProyecto) -> float:
    """
    Días hábiles desde fecha_inicio hasta el hito es_fin_plazo=True.
    Retorna 999.0 si no existe el hito o no tiene fecha calculada.
    """
    hito_fin = next((t for t in planificacion.tareas if t.es_fin_plazo), None)
    if not hito_fin or not hito_fin.fecha_fin:
        return 999.0

    return sum(
        1 for i in range(
            (hito_fin.fecha_fin - planificacion.parametros.fecha_inicio).days
        )
        if (
            planificacion.parametros.fecha_inicio + timedelta(days=i)
        ).weekday() < planificacion.parametros.dias_semana
    )


def calcular_sensibilidad(
    planificacion: PlanificacionProyecto,
    presupuesto: Presupuesto,
) -> list[dict]:
    """
    Retorna una entrada por recurso por tarea crítica,
    con horas, operarios, días actuales y reducción real del proyecto.
    Ordenado por tarea (fecha_inicio) y dentro de la tarea por horas descendente.
    """
    holguras        = calcular_holguras(planificacion.tareas)
    horas_dia       = planificacion.parametros.horas_dia
    duracion_actual = duracion_total_proyecto(planificacion)

    hito_fin        = next((t for t in planificacion.tareas if t.es_fin_plazo), None)
    fecha_fin_plazo = hito_fin.fecha_fin if hito_fin else None

    tareas_criticas = [
        t for t in planificacion.tareas
        if holguras.get(t.id, 1) <= 0
        and t.tipo == 'tarea'
        and t.horas_por_recurso
        and not t.es_fin_plazo
        and (
            fecha_fin_plazo is None
            or t.fecha_inicio is None
            or t.fecha_inicio <= fecha_fin_plazo
        )
    ]
    tareas_criticas.sort(key=lambda t: t.fecha_inicio or date.min)

    resultado: list[dict] = []
    for tarea in tareas_criticas:
        dias_por_recurso = {
            codigo: horas / (
                planificacion.escenario.operarios_por_recurso.get(codigo, 1) * horas_dia
            )
            for codigo, horas in tarea.horas_por_recurso.items()
            if horas > 0
        }

        if not dias_por_recurso:
            continue

        max_dias = max(dias_por_recurso.values())

        # Ordenar recursos por horas descendente dentro de la tarea
        recursos_ordenados = sorted(
            tarea.horas_por_recurso.items(),
            key=lambda x: x[1],
            reverse=True,
        )

        for codigo, horas in recursos_ordenados:
            if horas <= 0:
                continue

            operarios_actual = planificacion.escenario.operarios_por_recurso.get(codigo, 1)
            dias_actual      = horas / (operarios_actual * horas_dia)
            dias_nuevo       = horas / ((operarios_actual + 1) * horas_dia)
            es_limitante     = abs(dias_actual - max_dias) < 0.01

            nuevos_operarios = dict(planificacion.escenario.operarios_por_recurso)
            nuevos_operarios[codigo] = operarios_actual + 1

            nuevo_escenario = EscenarioRecursos(
                nombre='sensibilidad_temp',
                operarios_por_recurso=nuevos_operarios,
            )
            plan_nuevo = calcular_planificacion(
                presupuesto,
                planificacion.parametros,
                planificacion.tareas_originales,
                nuevo_escenario,
            )
            reduccion = duracion_actual - duracion_total_proyecto(plan_nuevo)

            recurso = presupuesto.recursos_mo.get(codigo)
            resultado.append({
                'tarea_id':           tarea.id,
                'tarea_nombre':       tarea.nombre,
                'tarea_fecha_inicio': tarea.fecha_inicio,
                'codigo':             codigo,
                'descripcion':        recurso.descripcion if recurso else codigo,
                'horas':              horas,
                'operarios_actual':   operarios_actual,
                'dias_tarea_actual':  dias_actual,
                'dias_tarea_nuevo':   dias_nuevo,
                'reduccion_proyecto': max(reduccion, 0.0),
                'es_limitante':       es_limitante,
            })

    return resultado


def calcular_recomendacion(
    planificacion: PlanificacionProyecto,
    presupuesto: Presupuesto,
    plazo_dias: int,
) -> tuple[list[dict], float]:
    """
    Calcula la combinación óptima de operarios para cumplir el plazo.
    Algoritmo greedy con recálculo completo de forward pass en cada iteración.
    Retorna (lista_ajustes, duracion_final_estimada).
    """
    operarios_trabajo = dict(planificacion.escenario.operarios_por_recurso)

    escenario_actual = EscenarioRecursos(
        nombre='recomendacion_temp',
        operarios_por_recurso=dict(operarios_trabajo),
    )
    plan_actual = calcular_planificacion(
        presupuesto,
        planificacion.parametros,
        planificacion.tareas_originales,
        escenario_actual,
    )
    dur_actual = duracion_total_proyecto(plan_actual)

    if dur_actual <= plazo_dias:
        return [], dur_actual

    ajustes: list[dict] = []

    for _ in range(30):
        if dur_actual <= plazo_dias:
            break

        sensibilidad = calcular_sensibilidad(plan_actual, presupuesto)
        if not sensibilidad:
            break
        ajuste_optimo = max(sensibilidad, key=lambda x: x['reduccion_proyecto'])
        if ajuste_optimo['reduccion_proyecto'] <= 0:
            break

        codigo    = ajuste_optimo['codigo']
        ops_antes = operarios_trabajo.get(codigo, 1)
        operarios_trabajo[codigo] = ops_antes + 1

        escenario_actual = EscenarioRecursos(
            nombre='recomendacion_temp',
            operarios_por_recurso=dict(operarios_trabajo),
        )
        plan_actual = calcular_planificacion(
            presupuesto,
            planificacion.parametros,
            planificacion.tareas_originales,
            escenario_actual,
        )
        dur_nueva = duracion_total_proyecto(plan_actual)

        ajustes.append({
            'codigo':            codigo,
            'descripcion':       ajuste_optimo['descripcion'],
            'operarios_antes':   ops_antes,
            'operarios_despues': operarios_trabajo[codigo],
            'impacto_dias':      dur_actual - dur_nueva,
        })
        dur_actual = dur_nueva

    return ajustes, dur_actual


def calcular_sobredimensionados(
    planificacion: PlanificacionProyecto,
    presupuesto: Presupuesto,
    operarios_trabajo: dict[str, int],
) -> list[dict]:
    """
    Detecta perfiles MO innecesarios: reduce en 1 operario y verifica
    que la duración total no aumenta respecto a planificacion.
    Retorna lista de perfiles reducibles con operarios_actual y operarios_optimo.
    """
    resultado = []
    dur_referencia = duracion_total_proyecto(planificacion)

    for codigo, ops_actual in operarios_trabajo.items():
        if ops_actual <= 1:
            continue

        operarios_reducidos = dict(operarios_trabajo)
        operarios_reducidos[codigo] = ops_actual - 1

        escenario_reducido = EscenarioRecursos(
            nombre='sobredimensionado_check',
            operarios_por_recurso=operarios_reducidos,
        )
        plan_reducido = calcular_planificacion(
            presupuesto,
            planificacion.parametros,
            planificacion.tareas_originales,
            escenario_reducido,
        )
        dur_reducida = duracion_total_proyecto(plan_reducido)

        if dur_reducida <= dur_referencia:
            recurso = presupuesto.recursos_mo.get(codigo)
            resultado.append({
                'codigo':           codigo,
                'descripcion':      recurso.descripcion if recurso else codigo,
                'operarios_actual': ops_actual,
                'operarios_optimo': ops_actual - 1,
            })

    return resultado


def dias_trabajados_tarea(tarea: TareaGantt, dias_semana: int) -> list[date]:
    """
    Devuelve los días hábiles que ocupa una tarea: los ceil(duracion_dias)
    días hábiles a partir de fecha_inicio (incluida).

    Este recuento reproduce exactamente el que usa calculator.sumar_dias_habiles
    para llegar a fecha_fin, así que los días de tareas encadenadas nunca se
    solapan ni dejan huecos: fecha_fin de una tarea es siempre el primer día
    hábil fuera del rango devuelto aquí.
    """
    if not tarea.fecha_inicio or not tarea.duracion_dias:
        return []

    dias_objetivo = math.ceil(tarea.duracion_dias)
    dias: list[date] = []
    fecha = tarea.fecha_inicio
    while len(dias) < dias_objetivo:
        if fecha.weekday() < dias_semana:
            dias.append(fecha)
        fecha += timedelta(days=1)
    return dias


def semana_iso(fecha: date) -> tuple[date, str]:
    """Retorna (lunes de la semana ISO, 'AAAA-Www') para una fecha dada."""
    lunes = fecha - timedelta(days=fecha.isoweekday() - 1)
    iso_year, iso_week, _ = fecha.isocalendar()
    return lunes, f'{iso_year}-W{iso_week:02d}'


def calcular_carga_semanal_recursos(
    planificacion: PlanificacionProyecto,
    presupuesto: Presupuesto,
) -> list[dict]:
    """
    Reparte linealmente las horas MO de cada tarea entre sus días hábiles
    y agrega el resultado por semana ISO y recurso.

    Cada fila del resultado contiene: semana_inicio (lunes de la semana ISO),
    semana_iso ('AAAA-Www'), recurso_codigo, recurso_nombre, horas_semana y
    tareas_incluidas (IDs de las tareas que aportan horas a ese recurso en
    esa semana). Las tareas sin horas de mano de obra o sin fechas calculadas
    se ignoran. No muta planificacion ni presupuesto.
    """
    dias_semana = planificacion.parametros.dias_semana
    acumulado: dict[tuple[str, date], dict] = {}

    for tarea in planificacion.tareas:
        if not tarea.horas_por_recurso:
            continue

        dias = dias_trabajados_tarea(tarea, dias_semana)
        if not dias:
            continue

        for codigo, horas_totales in tarea.horas_por_recurso.items():
            if horas_totales <= 0:
                continue
            horas_por_dia = horas_totales / len(dias)

            for dia in dias:
                lunes, etiqueta_iso = semana_iso(dia)
                clave = (codigo, lunes)
                if clave not in acumulado:
                    recurso = presupuesto.recursos_mo.get(codigo)
                    acumulado[clave] = {
                        'semana_inicio':    lunes,
                        'semana_iso':       etiqueta_iso,
                        'recurso_codigo':   codigo,
                        'recurso_nombre':   recurso.descripcion if recurso else codigo,
                        'horas_semana':     0.0,
                        'tareas_incluidas': [],
                    }
                fila = acumulado[clave]
                fila['horas_semana'] += horas_por_dia
                if tarea.id not in fila['tareas_incluidas']:
                    fila['tareas_incluidas'].append(tarea.id)

    filas = list(acumulado.values())
    for fila in filas:
        fila['horas_semana'] = round(fila['horas_semana'], 4)
    filas.sort(key=lambda f: (f['semana_inicio'], f['recurso_codigo']))
    return filas

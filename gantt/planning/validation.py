"""
Validaciones previas de planificación.

Comprueba relaciones entre presupuesto, tareas y escenarios antes de calcular
duraciones y fechas, para que los errores sean localizables.
"""

from collections import Counter, deque

from gantt.bc3.models import Presupuesto
from gantt.planning.calculator import partidas_para_codigo
from gantt.planning.models import EscenarioRecursos, ParametrosProyecto, TareaGantt


def recursos_mo_para_tarea(tarea: TareaGantt, presupuesto: Presupuesto) -> set[str]:
    """Devuelve recursos MO usados por los códigos BC3 asociados a una tarea."""
    recursos: set[str] = set()
    for bc3_codigo in tarea.capitulos_bc3:
        for partida in partidas_para_codigo(bc3_codigo, presupuesto):
            for linea in partida.lineas_mo:
                recursos.add(linea.codigo_recurso)
    return recursos


def detectar_ciclo_tareas(tareas: list[TareaGantt]) -> bool:
    """Detecta ciclos considerando solo dependencias entre IDs existentes."""
    by_id = {t.id: t for t in tareas}
    in_degree = {t.id: 0 for t in tareas}
    sucesores: dict[str, list[str]] = {t.id: [] for t in tareas}

    for tarea in tareas:
        for dep_id in tarea.dependencias:
            if dep_id in by_id:
                in_degree[tarea.id] += 1
                sucesores[dep_id].append(tarea.id)

    cola = deque(t_id for t_id, grado in in_degree.items() if grado == 0)
    visitados = 0
    while cola:
        t_id = cola.popleft()
        visitados += 1
        for sucesor_id in sucesores[t_id]:
            in_degree[sucesor_id] -= 1
            if in_degree[sucesor_id] == 0:
                cola.append(sucesor_id)

    return visitados != len(tareas)


def validar_planificacion_previa(
    presupuesto: Presupuesto,
    parametros: ParametrosProyecto,
    escenarios: list[EscenarioRecursos],
    tareas: list[TareaGantt],
) -> list[str]:
    """
    Valida una planificación antes de ejecutar el cálculo.
    Lanza ValueError con todos los problemas encontrados.
    """
    errores: list[str] = []
    warnings: list[str] = []

    if parametros.horas_dia <= 0:
        errores.append('proyecto.horas_dia debe ser mayor que 0')
    if parametros.dias_semana <= 0 or parametros.dias_semana > 7:
        errores.append('proyecto.dias_semana debe estar entre 1 y 7')
    if parametros.plazo_contractual_dias <= 0:
        errores.append('proyecto.plazo_contractual_dias debe ser mayor que 0')

    ids = [t.id for t in tareas]
    repetidos = sorted(tid for tid, count in Counter(ids).items() if count > 1)
    if repetidos:
        errores.append(f'IDs de tarea duplicados: {repetidos}')

    ids_validos = set(ids)
    for tarea in tareas:
        if tarea.tipo not in {'tarea', 'hito'}:
            errores.append(f'Tarea {tarea.id}: tipo inválido "{tarea.tipo}"')
        for dep_id in tarea.dependencias:
            if dep_id not in ids_validos:
                errores.append(f'Tarea {tarea.id}: dependencia inexistente "{dep_id}"')

    if detectar_ciclo_tareas(tareas):
        errores.append('Las dependencias de tareas contienen un ciclo')

    hitos_fin = [t for t in tareas if t.es_fin_plazo]
    if len(hitos_fin) != 1:
        errores.append(
            f'Debe existir exactamente un hito con es_fin_plazo=true; encontrados {len(hitos_fin)}'
        )
    for hito in hitos_fin:
        if hito.tipo != 'hito':
            errores.append(f'Tarea {hito.id}: es_fin_plazo=true solo debe usarse en hitos')

    recursos_necesarios: set[str] = set()
    for tarea in tareas:
        if tarea.tipo == 'hito':
            continue

        if tarea.duracion_dias_fija is None and not tarea.capitulos_bc3:
            errores.append(
                f'Tarea {tarea.id}: necesita duracion_dias_fija o capitulos_bc3'
            )

        for bc3_codigo in tarea.capitulos_bc3:
            partidas = partidas_para_codigo(bc3_codigo, presupuesto)
            if not partidas:
                errores.append(f'Tarea {tarea.id}: código BC3 inexistente "{bc3_codigo}"')

        recursos_tarea = recursos_mo_para_tarea(tarea, presupuesto)
        if tarea.duracion_dias_fija is None and tarea.capitulos_bc3 and not recursos_tarea:
            errores.append(
                f'Tarea {tarea.id}: sus códigos BC3 no aportan horas de mano de obra'
            )
        recursos_necesarios.update(recursos_tarea)

    if not escenarios:
        errores.append('Debe existir al menos un escenario de recursos')

    recursos_presupuesto = set(presupuesto.recursos_mo)
    for escenario in escenarios:
        recursos_escenario = set(escenario.operarios_por_recurso)
        inexistentes = sorted(recursos_escenario - recursos_presupuesto)
        if inexistentes:
            errores.append(
                f'Escenario {escenario.nombre}: recursos MO inexistentes {inexistentes}'
            )

        incompletos = sorted(recursos_necesarios - recursos_escenario)
        if incompletos:
            warnings.append(
                f'Escenario {escenario.nombre}: faltan recursos MO necesarios {incompletos}'
            )

        no_positivos = sorted(
            codigo for codigo, operarios in escenario.operarios_por_recurso.items()
            if operarios <= 0
        )
        if no_positivos:
            errores.append(
                f'Escenario {escenario.nombre}: operarios no positivos {no_positivos}'
            )

    if errores:
        detalle = '\n'.join(f'- {error}' for error in errores)
        raise ValueError(f'Planificación inválida:\n{detalle}')

    return warnings

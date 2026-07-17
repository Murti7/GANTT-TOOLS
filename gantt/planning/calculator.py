"""
Módulo responsable de calcular la planificación temporal del proyecto.

Responsabilidad: ejecutar el pipeline completo —horas MO desde BC3,
duraciones por escenario de recursos y forward pass de fechas— para producir
una PlanificacionProyecto con todas las tareas fechadas.
"""

import math
from collections import deque
from datetime import date, timedelta

from gantt.bc3.models import Capitulo, Partida, Presupuesto
from gantt.planning.models import (
    BandaVisual,
    EscenarioRecursos,
    ParametrosProyecto,
    PlanificacionProyecto,
    TareaGantt,
)


def sumar_dias_habiles(fecha_inicio: date, dias: float, dias_semana: int) -> date:
    """
    Suma un número de días hábiles a una fecha de inicio.
    Con dias <= 0 retorna fecha_inicio sin modificar.
    Soporta duraciones decimales redondeando al entero superior.
    """
    if dias <= 0:
        return fecha_inicio
    enteros = math.ceil(dias)
    fecha = fecha_inicio
    for _ in range(enteros):
        fecha += timedelta(days=1)
        while fecha.weekday() >= dias_semana:
            fecha += timedelta(days=1)
    return fecha


def ordenar_topologico(tareas: list[TareaGantt]) -> list[TareaGantt]:
    """
    Ordena las tareas respetando dependencias mediante el algoritmo de Kahn (BFS).
    Lanza ValueError si hay dependencias circulares.
    Las dependencias que referencian IDs no presentes en la lista se ignoran.
    """
    by_id = {t.id: t for t in tareas}
    in_degree = {t.id: 0 for t in tareas}
    sucesores: dict[str, list[str]] = {t.id: [] for t in tareas}

    for tarea in tareas:
        for dep_id in tarea.dependencias:
            if dep_id in by_id:
                in_degree[tarea.id] += 1
                sucesores[dep_id].append(tarea.id)

    cola = deque(t_id for t_id, grado in in_degree.items() if grado == 0)
    resultado = []

    while cola:
        t_id = cola.popleft()
        resultado.append(by_id[t_id])
        for sucesor_id in sucesores[t_id]:
            in_degree[sucesor_id] -= 1
            if in_degree[sucesor_id] == 0:
                cola.append(sucesor_id)

    if len(resultado) != len(tareas):
        raise ValueError('Ciclo detectado en las dependencias de las tareas')

    return resultado


def buscar_capitulo(codigo: str, capitulos: list[Capitulo]) -> Capitulo | None:
    """
    Busca recursivamente un capítulo o subcapítulo por código exacto en el árbol.
    Retorna None si no se encuentra.
    """
    for cap in capitulos:
        if cap.codigo == codigo:
            return cap
        encontrado = buscar_capitulo(codigo, cap.subcapitulos)
        if encontrado is not None:
            return encontrado
    return None


def partidas_para_codigo(bc3_codigo: str, presupuesto: Presupuesto) -> list[Partida]:
    """
    Retorna las partidas del presupuesto que corresponden al código BC3 indicado.

    - Código con '#': localiza el capítulo en el árbol y recoge recursivamente
      todas sus partidas directas y de subcapítulos hijos.
    - Código sin '#': busca en el árbol la partida con ese código exacto.
    """
    if bc3_codigo.endswith('#'):
        cap = buscar_capitulo(bc3_codigo, presupuesto.capitulos)
        if cap is None:
            return []
        partidas: list[Partida] = []
        pila = [cap]
        while pila:
            nodo = pila.pop()
            partidas.extend(nodo.partidas)
            pila.extend(nodo.subcapitulos)
        return partidas

    # Código exacto: recorrer el árbol completo hasta encontrarlo
    pila_caps = list(presupuesto.capitulos)
    while pila_caps:
        nodo = pila_caps.pop()
        for partida in nodo.partidas:
            if partida.codigo == bc3_codigo:
                return [partida]
        pila_caps.extend(nodo.subcapitulos)
    return []


def calcular_planificacion(
    presupuesto: Presupuesto,
    parametros: ParametrosProyecto,
    tareas: list[TareaGantt],
    escenario: EscenarioRecursos,
    bandas: list[BandaVisual] | None = None,
) -> PlanificacionProyecto:
    """
    Calcula la planificación temporal del proyecto para un escenario dado.
    Ejecuta horas MO, duraciones y forward pass en orden topológico.
    No muta los objetos de entrada.

    bandas son las bandas visuales del diagrama de red definidas en
    planificacion.yaml; se adjuntan a la PlanificacionProyecto resultante
    sin intervenir en el cálculo. Si no se indican, queda una lista vacía.
    """
    # --- PASOS 1 + 2: horas MO y duración por tarea ---
    tareas_con_duracion: list[TareaGantt] = []

    for tarea in tareas:
        # Paso 1: acumular horas por recurso MO desde los capítulos BC3
        horas_por_recurso: dict[str, float] = {}
        for bc3_codigo in tarea.capitulos_bc3:
            for partida in partidas_para_codigo(bc3_codigo, presupuesto):
                for linea in partida.lineas_mo:
                    horas_por_recurso[linea.codigo_recurso] = (
                        horas_por_recurso.get(linea.codigo_recurso, 0.0)
                        + linea.cantidad * partida.cantidad
                    )

        # Paso 2: calcular duración y perfil limitante
        if tarea.duracion_dias_fija is not None:
            duracion    = float(tarea.duracion_dias_fija)
            perfil      = 'estacional' if tarea.estacional else None
            horas_total = 0.0

        elif horas_por_recurso:
            dias_por_recurso = {
                codigo: horas / (escenario.operarios_por_recurso.get(codigo, 1) * parametros.horas_dia)
                for codigo, horas in horas_por_recurso.items()
                if horas > 0
            }
            if dias_por_recurso:
                limitante   = max(dias_por_recurso, key=lambda c: dias_por_recurso[c])
                recurso_lim = presupuesto.recursos_mo.get(limitante)
                perfil      = recurso_lim.descripcion if recurso_lim else limitante
                duracion    = dias_por_recurso[limitante]
                horas_total = sum(horas_por_recurso.values())
            else:
                duracion, perfil, horas_total = 0.0, None, 0.0

        else:
            duracion, perfil, horas_total = 0.0, None, 0.0

        tarea_calculada = tarea.model_copy(update={
            'duracion_dias':    duracion,
            'horas_mo_total':   horas_total,
            'perfil_limitante': perfil,
            'horas_por_recurso': horas_por_recurso,
        })
        tareas_con_duracion.append(tarea_calculada)

    # --- PASO 3: forward pass —calcular fechas en orden topológico ---
    tareas_ordenadas = ordenar_topologico(tareas_con_duracion)
    by_id: dict[str, TareaGantt] = {}
    resultado: list[TareaGantt] = []

    ids_validos = {t.id for t in tareas_ordenadas}

    for tarea in tareas_ordenadas:
        fechas_prev = []
        for dep_id in tarea.dependencias:
            if dep_id in by_id and by_id[dep_id].fecha_fin is not None:
                fechas_prev.append(by_id[dep_id].fecha_fin)
            elif dep_id not in ids_validos:
                raise ValueError(
                    f'Tarea {tarea.id} depende de {dep_id} '
                    f'que no existe en la planificacion'
                )
        fecha_ini = max(fechas_prev) if fechas_prev else parametros.fecha_inicio
        fecha_f   = sumar_dias_habiles(fecha_ini, tarea.duracion_dias or 0.0, parametros.dias_semana)

        tarea_fechada = tarea.model_copy(update={
            'fecha_inicio': fecha_ini,
            'fecha_fin':    fecha_f,
        })
        by_id[tarea.id] = tarea_fechada
        resultado.append(tarea_fechada)

    return PlanificacionProyecto(
        parametros=parametros,
        escenario=escenario,
        tareas=resultado,
        tareas_originales=tareas,
        bandas=bandas or [],
    )


if __name__ == '__main__':
    import sys
    from pathlib import Path
    from gantt.bc3.parser import parse_bc3
    from gantt.planning.models import cargar_planificacion_yaml

    project_name = sys.argv[1] if len(sys.argv) > 1 else 'complex-balear'
    input_dir = Path('projects') / project_name / 'input'

    bc3_files = list(input_dir.glob('*.bc3'))
    if not bc3_files:
        raise FileNotFoundError(f'No se encontró .bc3 en {input_dir}')

    yaml_path = input_dir / 'planificacion.yaml'
    if not yaml_path.exists():
        raise FileNotFoundError(f'No se encontró planificacion.yaml en {input_dir}')

    presupuesto = parse_bc3(bc3_files[0])
    parametros, escenarios, tareas, bandas = cargar_planificacion_yaml(yaml_path)

    print(f'Proyecto: {parametros.nombre}')
    print(f'Plazo contractual: {parametros.plazo_contractual_dias} días hábiles')
    print()

    for escenario in escenarios:
        planificacion = calcular_planificacion(
            presupuesto, parametros, tareas, escenario, bandas
        )
        print(f'=== ESCENARIO: {escenario.nombre} ===')
        print(f'{"ID":<12} {"Nombre":<45} {"Días":>6} {"Inicio":<12} '
              f'{"Fin":<12} {"Perfil limitante":<30} {"H MO":>8}')
        print('-' * 130)
        for tarea in planificacion.tareas:
            perfil = tarea.perfil_limitante or ('estacional' if tarea.estacional else '—')
            horas  = f'{tarea.horas_mo_total:.1f}' if tarea.horas_mo_total else '—'
            dias   = f'{tarea.duracion_dias:.1f}' if tarea.duracion_dias else '0'
            print(f'{tarea.id:<12} {tarea.nombre[:44]:<45} {dias:>6} '
                  f'{str(tarea.fecha_inicio):<12} {str(tarea.fecha_fin):<12} '
                  f'{perfil:<30} {horas:>8}')

        hito_fin = next((t for t in planificacion.tareas if t.es_fin_plazo), None)
        if hito_fin and hito_fin.fecha_fin:
            dias_totales = sum(
                1 for i in range((hito_fin.fecha_fin - parametros.fecha_inicio).days)
                if (parametros.fecha_inicio + timedelta(days=i)).weekday() < parametros.dias_semana
            )
            cumple = 'CUMPLE' if dias_totales <= parametros.plazo_contractual_dias else 'NO CUMPLE'
            print()
            print(f'Duracion total: {dias_totales} dias habiles '
                  f'(plazo: {parametros.plazo_contractual_dias}) -> {cumple}')
        print()

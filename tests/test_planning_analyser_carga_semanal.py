"""
Tests de la carga semanal de recursos.

Responsabilidad: verificar que calcular_carga_semanal_recursos reparte
linealmente las horas MO de cada tarea entre sus días hábiles y agrega
correctamente por semana ISO y recurso, sin mutar la planificación de entrada.
"""

from datetime import date

from gantt.bc3.models import Presupuesto, RecursoMO
from gantt.planning.analyser import calcular_carga_semanal_recursos
from gantt.planning.models import (
    EscenarioRecursos,
    ParametrosProyecto,
    PlanificacionProyecto,
    TareaGantt,
)

PARAMETROS = ParametrosProyecto(
    nombre='Proyecto de prueba',
    fecha_inicio=date(2025, 1, 6),
    horas_dia=8,
    dias_semana=5,
    plazo_contractual_dias=30,
)
ESCENARIO = EscenarioRecursos(nombre='base', operarios_por_recurso={'MO-x': 1})


def _presupuesto() -> Presupuesto:
    return Presupuesto(
        codigo='OBRA##',
        descripcion='Proyecto de prueba',
        importe_total=0.0,
        capitulos=[],
        recursos_mo={'MO-x': RecursoMO(codigo='MO-x', descripcion='Oficial X', precio_hora=20.0)},
    )


def _tarea(
    id_: str,
    fecha_inicio: date,
    duracion_dias: float,
    horas_por_recurso: dict[str, float],
) -> TareaGantt:
    return TareaGantt(
        id=id_,
        nombre=f'Tarea {id_}',
        tipo='tarea',
        capitulos_bc3=[],
        dependencias=[],
        fecha_inicio=fecha_inicio,
        duracion_dias=duracion_dias,
        horas_por_recurso=horas_por_recurso,
    )


def _planificacion(tareas: list[TareaGantt]) -> PlanificacionProyecto:
    return PlanificacionProyecto(
        parametros=PARAMETROS,
        escenario=ESCENARIO,
        tareas=tareas,
        tareas_originales=tareas,
    )


def test_tarea_dentro_de_una_sola_semana():
    # Lunes 6 a miércoles 8 de enero 2025: 3 días hábiles, misma semana ISO.
    tarea = _tarea('T1', date(2025, 1, 6), 3.0, {'MO-x': 24.0})
    planificacion = _planificacion([tarea])

    filas = calcular_carga_semanal_recursos(planificacion, _presupuesto())

    assert len(filas) == 1
    fila = filas[0]
    assert fila['semana_inicio'] == date(2025, 1, 6)
    assert fila['semana_iso'] == '2025-W02'
    assert fila['recurso_codigo'] == 'MO-x'
    assert fila['recurso_nombre'] == 'Oficial X'
    assert fila['horas_semana'] == 24.0
    assert fila['tareas_incluidas'] == ['T1']


def test_tarea_que_cruza_dos_semanas():
    # Viernes 10 (semana W02) + lunes 13 (semana W03): 2 días hábiles, cruzando fin de semana.
    tarea = _tarea('T1', date(2025, 1, 10), 2.0, {'MO-x': 10.0})
    planificacion = _planificacion([tarea])

    filas = calcular_carga_semanal_recursos(planificacion, _presupuesto())

    assert len(filas) == 2
    filas_por_semana = {f['semana_iso']: f for f in filas}
    assert set(filas_por_semana) == {'2025-W02', '2025-W03'}
    assert filas_por_semana['2025-W02']['horas_semana'] == 5.0
    assert filas_por_semana['2025-W03']['horas_semana'] == 5.0
    assert filas_por_semana['2025-W02']['tareas_incluidas'] == ['T1']
    assert filas_por_semana['2025-W03']['tareas_incluidas'] == ['T1']


def test_varias_tareas_comparten_recurso_en_la_misma_semana():
    t1 = _tarea('T1', date(2025, 1, 6), 1.0, {'MO-x': 4.0})
    t2 = _tarea('T2', date(2025, 1, 7), 1.0, {'MO-x': 6.0})
    planificacion = _planificacion([t1, t2])

    filas = calcular_carga_semanal_recursos(planificacion, _presupuesto())

    assert len(filas) == 1
    fila = filas[0]
    assert fila['semana_iso'] == '2025-W02'
    assert fila['horas_semana'] == 10.0
    assert fila['tareas_incluidas'] == ['T1', 'T2']


def test_tareas_sin_horas_mo_se_ignoran():
    hito = _tarea('H1', date(2025, 1, 6), 0.0, {})
    tarea_normal = _tarea('T1', date(2025, 1, 6), 1.0, {'MO-x': 8.0})
    planificacion = _planificacion([hito, tarea_normal])

    filas = calcular_carga_semanal_recursos(planificacion, _presupuesto())

    assert len(filas) == 1
    assert filas[0]['tareas_incluidas'] == ['T1']


def test_no_muta_planificacion_de_entrada():
    tarea = _tarea('T1', date(2025, 1, 6), 3.0, {'MO-x': 24.0})
    planificacion = _planificacion([tarea])
    copia_tareas_antes = list(planificacion.tareas)

    calcular_carga_semanal_recursos(planificacion, _presupuesto())

    assert planificacion.tareas == copia_tareas_antes
    assert planificacion.tareas[0].horas_por_recurso == {'MO-x': 24.0}

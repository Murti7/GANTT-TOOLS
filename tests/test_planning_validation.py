"""
Tests de validación previa de planificación.
"""

from datetime import date

import pytest

from gantt.bc3.models import (
    Capitulo,
    LineaDescompuesto,
    Partida,
    Presupuesto,
    RecursoMO,
)
from gantt.planning.models import EscenarioRecursos, ParametrosProyecto, TareaGantt
from gantt.planning.validation import validar_planificacion_previa


PARAMETROS = ParametrosProyecto(
    nombre='Proyecto',
    fecha_inicio=date(2026, 1, 1),
    horas_dia=8,
    dias_semana=5,
    plazo_contractual_dias=20,
)


def _presupuesto() -> Presupuesto:
    partida = Partida(
        codigo='P01',
        descripcion='Partida',
        unidad='ud',
        precio_unitario=100.0,
        cantidad=2.0,
        lineas_mo=[LineaDescompuesto(codigo_recurso='MO-1', cantidad=4.0)],
    )
    return Presupuesto(
        codigo='OBRA##',
        descripcion='Proyecto',
        importe_total=200.0,
        capitulos=[
            Capitulo(
                codigo='A#',
                descripcion='Capitulo',
                importe_total=200.0,
                partidas=[partida],
                subcapitulos=[],
            )
        ],
        recursos_mo={'MO-1': RecursoMO(codigo='MO-1', descripcion='Oficial', precio_hora=20.0)},
    )


def _tareas() -> list[TareaGantt]:
    return [
        TareaGantt(
            id='T1',
            nombre='Tarea',
            tipo='tarea',
            capitulos_bc3=['A#'],
            dependencias=[],
        ),
        TareaGantt(
            id='FIN',
            nombre='Fin',
            tipo='hito',
            capitulos_bc3=[],
            duracion_dias_fija=0,
            dependencias=['T1'],
            es_fin_plazo=True,
        ),
    ]


def test_validar_planificacion_previa_acepta_planificacion_consistente():
    warnings = validar_planificacion_previa(
        _presupuesto(),
        PARAMETROS,
        [EscenarioRecursos(nombre='base', operarios_por_recurso={'MO-1': 1})],
        _tareas(),
    )

    assert warnings == []


def test_validar_planificacion_previa_acumula_errores_relevantes():
    tareas = _tareas()
    tareas.append(
        TareaGantt(
            id='T2',
            nombre='Sin datos',
            tipo='tarea',
            capitulos_bc3=['NO_EXISTE'],
            dependencias=['NO_DEP'],
        )
    )

    with pytest.raises(ValueError) as exc_info:
        validar_planificacion_previa(
            _presupuesto(),
            PARAMETROS,
            [EscenarioRecursos(nombre='base', operarios_por_recurso={})],
            tareas,
        )

    mensaje = str(exc_info.value)
    assert 'dependencia inexistente "NO_DEP"' in mensaje
    assert 'código BC3 inexistente "NO_EXISTE"' in mensaje


def test_validar_planificacion_previa_avisa_si_escenario_esta_incompleto():
    warnings = validar_planificacion_previa(
        _presupuesto(),
        PARAMETROS,
        [EscenarioRecursos(nombre='base', operarios_por_recurso={})],
        _tareas(),
    )

    assert warnings == ["Escenario base: faltan recursos MO necesarios ['MO-1']"]


def test_validar_planificacion_previa_detecta_ciclos_y_recursos_inexistentes():
    tareas = _tareas()
    tareas[0].dependencias.append('FIN')

    with pytest.raises(ValueError) as exc_info:
        validar_planificacion_previa(
            _presupuesto(),
            PARAMETROS,
            [EscenarioRecursos(nombre='base', operarios_por_recurso={'MO-1': 1, 'MO-X': 1})],
            tareas,
        )

    mensaje = str(exc_info.value)
    assert 'contienen un ciclo' in mensaje
    assert "recursos MO inexistentes ['MO-X']" in mensaje

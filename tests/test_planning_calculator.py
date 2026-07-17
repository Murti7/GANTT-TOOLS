"""
Tests del cálculo de planificación.

Responsabilidad: verificar el ordenamiento topológico de tareas,
en particular la detección de dependencias circulares.
"""

import pytest

from gantt.planning.calculator import ordenar_topologico
from gantt.planning.models import TareaGantt


def _tarea(id_: str, dependencias: list[str]) -> TareaGantt:
    return TareaGantt(
        id=id_,
        nombre=f'Tarea {id_}',
        tipo='tarea',
        capitulos_bc3=[],
        dependencias=dependencias,
    )


def test_ordenar_topologico_orden_valido():
    t1 = _tarea('T1', [])
    t2 = _tarea('T2', ['T1'])
    t3 = _tarea('T3', ['T2'])

    orden = ordenar_topologico([t3, t1, t2])

    assert [t.id for t in orden] == ['T1', 'T2', 'T3']


def test_ordenar_topologico_detecta_dependencia_circular():
    t1 = _tarea('T1', ['T2'])
    t2 = _tarea('T2', ['T1'])

    with pytest.raises(ValueError, match='[Cc]iclo'):
        ordenar_topologico([t1, t2])

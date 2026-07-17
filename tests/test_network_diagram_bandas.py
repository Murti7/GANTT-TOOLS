"""
Tests de las bandas visuales del diagrama de red.

Responsabilidad: verificar que las bandas definidas en planificacion.yaml
se usan tal cual, que una tarea con grupo_visual desconocido se reasigna
a la banda automática "Otros" en vez de descartarse, y que un yaml antiguo
sin sección 'bandas' sigue funcionando con las bandas por defecto.
"""

from datetime import date

from gantt.planning.models import BandaVisual, TareaGantt
from gantt.reporting.network_diagram import (
    BANDA_OTROS_ID,
    BANDAS_DEFAULT,
    agrupar_tareas,
    calcular_y,
    resolver_bandas,
)


def _tarea(id_: str, grupo_visual: str | None, orden_visual: int | None = None) -> TareaGantt:
    return TareaGantt(
        id=id_,
        nombre=f'Tarea {id_}',
        tipo='tarea',
        capitulos_bc3=[],
        dependencias=[],
        grupo_visual=grupo_visual,
        orden_visual=orden_visual,
        fecha_inicio=date(2025, 1, 1),
        fecha_fin=date(2025, 1, 5),
    )


def test_planificacion_usa_bandas_definidas_en_yaml():
    bandas = [
        BandaVisual(id='auditorias', titulo='Auditorías', descripcion='', color='#D9EAF7'),
        BandaVisual(id='ejecucion', titulo='Ejecución', descripcion='', color='#E2F0D9'),
    ]
    tareas = [
        _tarea('T1', 'auditorias'),
        _tarea('T2', 'ejecucion'),
    ]

    resueltas = resolver_bandas(bandas)
    assert resueltas == bandas

    y_tareas, bandas_y, bandas_render = calcular_y(tareas, resueltas)

    assert [b.id for b in bandas_render] == ['auditorias', 'ejecucion']
    assert 'T1' in y_tareas and 'T2' in y_tareas
    assert set(bandas_y) == {'auditorias', 'ejecucion'}


def test_tarea_con_grupo_visual_desconocido_va_a_otros(capsys):
    bandas = [BandaVisual(id='auditorias', titulo='Auditorías', descripcion='')]
    tareas = [
        _tarea('T1', 'auditorias'),
        _tarea('T2', 'fase_inexistente'),
    ]

    grupos = agrupar_tareas(tareas, {'auditorias'})

    # La tarea con grupo desconocido no se descarta: aparece en "Otros".
    assert grupos['auditorias'] == [tareas[0]]
    assert grupos[BANDA_OTROS_ID] == [tareas[1]]

    salida = capsys.readouterr().out
    assert 'fase_inexistente' in salida
    assert 'AVISO' in salida

    # calcular_y también debe incluir la tarea y crear la banda "Otros".
    y_tareas, bandas_y, bandas_render = calcular_y(tareas, bandas)
    assert 'T1' in y_tareas
    assert 'T2' in y_tareas
    assert any(b.id == BANDA_OTROS_ID for b in bandas_render)
    assert BANDA_OTROS_ID in bandas_y


def test_compatibilidad_yaml_sin_seccion_bandas():
    # Un yaml antiguo sin 'bandas' produce una lista vacía en el modelo;
    # resolver_bandas debe recurrir a BANDAS_DEFAULT sin lanzar error.
    resueltas = resolver_bandas([])

    assert resueltas == BANDAS_DEFAULT
    assert any(b.id == 'Cierre contractual' for b in resueltas)

    tarea = _tarea('T1', None)  # sin grupo_visual, como en proyectos existentes
    y_tareas, bandas_y, bandas_render = calcular_y([tarea], resueltas)

    assert 'T1' in y_tareas
    assert len(bandas_render) == len(BANDAS_DEFAULT)

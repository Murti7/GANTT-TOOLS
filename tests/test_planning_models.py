"""
Tests de carga de planificacion.yaml.

Responsabilidad: verificar que el loader personalizado preserva como texto
los IDs de tarea con cero inicial (ej: '07'), que PyYAML convertiría a
entero por defecto.
"""

from pathlib import Path

from gantt.planning.models import cargar_planificacion_yaml

PLANIFICACION_YAML = """
proyecto:
  nombre: Proyecto de prueba
  fecha_inicio: '2025-01-01'
  horas_dia: 8
  dias_semana: 5
  plazo_contractual_dias: 30

escenarios:
  base:
    MO-x: 1

tareas:
  07:
    nombre: Tarea con cero inicial
    tipo: tarea
    capitulos_bc3: []
    dependencias: []
  FIN:
    nombre: Cierre
    tipo: hito
    capitulos_bc3: []
    duracion_dias_fija: 0
    dependencias: ['07']
    es_fin_plazo: true
"""


def test_cargar_planificacion_yaml_preserva_cero_inicial(tmp_path: Path):
    ruta = tmp_path / 'planificacion.yaml'
    ruta.write_text(PLANIFICACION_YAML, encoding='utf-8')

    parametros, escenarios, tareas, bandas = cargar_planificacion_yaml(ruta)

    ids = [t.id for t in tareas]
    assert '07' in ids
    assert 7 not in ids
    assert '7' not in ids

    tarea_fin = next(t for t in tareas if t.id == 'FIN')
    assert tarea_fin.dependencias == ['07']

    assert parametros.nombre == 'Proyecto de prueba'
    assert escenarios[0].operarios_por_recurso == {'MO-x': 1}
    assert bandas == []

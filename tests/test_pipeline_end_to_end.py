"""
Test end-to-end del pipeline principal en un workspace temporal.
"""

import json
from pathlib import Path

from main import main


BC3_MINIMO = r"""~C|OBRA##||Proyecto mini|100.00||0|
~D|OBRA##
|A#\\1
\|
~C|A#||Capitulo A|100.00||0|
~D|A#
|A01\\2
\|
~C|A01|ud|Partida de ejemplo|50.00||0|
~D|A01
|MO-OFI1\\3
\|
~C|MO-OFI1|h|Oficial 1a|20.00||1|
"""


PLANIFICACION_YAML = """
proyecto:
  nombre: Proyecto mini
  fecha_inicio: '2026-01-05'
  horas_dia: 8
  dias_semana: 5
  plazo_contractual_dias: 20

escenarios:
  base:
    MO-OFI1: 1

tareas:
  T1:
    nombre: Ejecutar partida
    tipo: tarea
    capitulos_bc3: ['A#']
    dependencias: []
  FIN:
    nombre: Fin
    tipo: hito
    capitulos_bc3: []
    duracion_dias_fija: 0
    dependencias: [T1]
    es_fin_plazo: true
"""


def test_pipeline_end_to_end_en_tmp_path(monkeypatch, tmp_path: Path):
    project_input = tmp_path / 'projects' / 'Mini' / 'input'
    project_input.mkdir(parents=True)
    (project_input / 'mini.bc3').write_text(BC3_MINIMO, encoding='latin-1')
    (project_input / 'config.yaml').write_text(
        'proyecto: Proyecto mini configurado\nporcentaje_gg: 0.13\n',
        encoding='utf-8',
    )
    (project_input / 'planificacion.yaml').write_text(PLANIFICACION_YAML, encoding='utf-8')

    monkeypatch.chdir(tmp_path)

    main('Mini')

    output_dir = tmp_path / 'projects' / 'Mini' / 'output'
    manifest_path = output_dir / 'run_manifest.json'
    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))

    assert (output_dir / 'PRES.01_Cuadro_Oferta.xlsx').exists()
    assert (output_dir / 'PRES.05_VEC_Liquidacion.xlsx').exists()
    assert (output_dir / 'Mini_gantt.xlsx').exists()
    assert (output_dir / 'Mini_gantt.png').exists()
    assert (output_dir / 'reporting' / '01_economico' / '01_importe_por_capitulo.png').exists()
    assert manifest['project'] == 'Mini'
    assert manifest['bc3_original_name'] == 'mini.bc3'
    assert manifest['planning_present'] is True
    assert manifest['scenarios'] == ['base']
    assert any(path.endswith('Mini_gantt.png') for path in manifest['outputs'])

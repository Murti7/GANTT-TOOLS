from pathlib import Path

from gantt.application.budgeting import run_budgeting
from gantt.application.context import resolver_contexto_ejecucion
from gantt.application.planning import run_planning


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


def _workspace(tmp_path: Path) -> Path:
    input_dir = tmp_path / "projects" / "Mini" / "input"
    input_dir.mkdir(parents=True)
    (input_dir / "mini.bc3").write_text(BC3_MINIMO, encoding="latin-1")
    (input_dir / "config.yaml").write_text("proyecto: Proyecto mini\n", encoding="utf-8")
    (input_dir / "planificacion.yaml").write_text(PLANIFICACION_YAML, encoding="utf-8")
    return tmp_path


def test_run_budgeting_escribe_output_por_capability(tmp_path: Path):
    workspace = _workspace(tmp_path)
    context = resolver_contexto_ejecucion("Mini", workspace_root=workspace)

    result = run_budgeting(context)

    root = workspace / "projects" / "Mini" / "output" / "budgeting" / "mini"
    assert result.output_root == root
    assert (root / "documents" / "PRES.01_Cuadro_Oferta.xlsx").exists()
    assert (root / "analysis" / "budget_analysis.xlsx").exists()
    assert (root / "analysis" / "charts" / "01_economico" / "01_importe_por_capitulo.png").exists()
    assert (root / "run_manifest.json").exists()


def test_run_planning_escribe_output_por_capability(tmp_path: Path):
    workspace = _workspace(tmp_path)
    context = resolver_contexto_ejecucion("Mini", workspace_root=workspace)

    result = run_planning(context)

    root = workspace / "projects" / "Mini" / "output" / "planning" / "mini"
    assert result.output_root == root
    assert result.scenarios == ["base"]
    assert (root / "planning_analysis.xlsx").exists()
    assert (root / "gantt.png").exists()
    assert (root / "run_manifest.json").exists()

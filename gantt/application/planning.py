"""Caso de uso Planning: BC3 + planning.yaml -> analisis/diagrama."""

from pathlib import Path

from gantt.application.context import ProjectExecutionContext
from gantt.application.documents import apply_document_identity_fallbacks, planning_document_metadata
from gantt.application.manifest import write_run_manifest
from gantt.application.outputs import planning_output_dirs, source_id_from_path
from gantt.application.product import Capability
from gantt.bc3.parser import parse_bc3
from gantt.planning.calculator import calcular_planificacion
from gantt.planning.models import cargar_planificacion_yaml
from gantt.planning.validation import validar_planificacion_previa
from gantt.reporting.excel_exporter import exportar_analisis
from gantt.reporting.documents import DocumentType
from gantt.reporting.network_diagram import generar_diagrama_red
from gantt.reporting.presentation import build_presentation_context


class PlanningRunResult:
    def __init__(self, output_root: Path, outputs: list[Path], scenarios: list[str], warnings: list[str]) -> None:
        self.output_root = output_root
        self.outputs = outputs
        self.scenarios = scenarios
        self.warnings = warnings


def run_planning(context: ProjectExecutionContext) -> PlanningRunResult:
    """Ejecuta solo la capability Planning en la estructura ENG-3."""
    if context.planificacion_path is None:
        raise FileNotFoundError("No se encontro planificacion.yaml para Planning.")

    dirs = planning_output_dirs(context.project_root, source_id_from_path(context.bc3_path))
    for path in dirs.values():
        path.mkdir(parents=True, exist_ok=True)

    presupuesto = parse_bc3(context.bc3_path)
    if context.config_path:
        presupuesto = presupuesto.model_copy(
            update={"config": presupuesto.config.model_copy(update=context.project_config_update)}
        )
    presupuesto = apply_document_identity_fallbacks(presupuesto, context)
    presupuesto.company = context.company_config
    presentation = build_presentation_context(presupuesto.company)

    parametros, escenarios, tareas, bandas = cargar_planificacion_yaml(context.planificacion_path)
    warnings = validar_planificacion_previa(presupuesto, parametros, escenarios, tareas)

    planificaciones = [
        calcular_planificacion(presupuesto, parametros, tareas, escenario, bandas)
        for escenario in escenarios
    ]

    outputs: list[Path] = []
    analysis_path = dirs["root"] / "planning_analysis.xlsx"
    analysis_metadata = planning_document_metadata(
        context,
        DocumentType.PLANNING_ANALYSIS,
        scenario=planificaciones[-1].escenario.nombre,
        contractual_deadline=planificaciones[-1].parametros.plazo_contractual_dias,
    ).model_copy(
        update={
            "project_code": presupuesto.config.numero_expediente,
            "project_name": presupuesto.config.proyecto,
            "site_name": presupuesto.config.edificio,
            "client_name": presupuesto.config.entidad,
            "reference": presupuesto.config.numero_expediente,
            "revision": presupuesto.config.revision,
        }
    )
    exportar_analisis(
        presupuesto,
        analysis_path,
        planificaciones,
        metadata=analysis_metadata,
        presentation=presentation,
    )
    outputs.append(analysis_path)

    diagram_path = dirs["root"] / "gantt.png"
    diagram_path.write_bytes(generar_diagrama_red(planificaciones[0], palette=presentation.palette))
    outputs.append(diagram_path)

    manifest_context = context.model_copy(update={"output_dir": dirs["root"]})
    write_run_manifest(
        manifest_context,
        outputs,
        [plan.escenario.nombre for plan in planificaciones],
        warnings,
        presentation=presentation,
        capability=Capability.PLANNING,
    )
    return PlanningRunResult(
        dirs["root"],
        outputs,
        [plan.escenario.nombre for plan in planificaciones],
        list(context.warnings) + warnings,
    )

"""Caso de uso Budgeting: BC3 -> documentos/analisis/charts."""

from pathlib import Path

from gantt.application.context import ProjectExecutionContext
from gantt.application.documents import apply_document_identity_fallbacks, budget_document_metadata
from gantt.application.manifest import write_run_manifest
from gantt.application.outputs import budget_output_dirs, source_id_from_path
from gantt.application.product import Capability
from gantt.bc3.parser import parse_bc3
from gantt.reporting.analysis_charts import AnalysisChartsReport
from gantt.reporting.documents import DocumentType
from gantt.reporting.excel_exporter import exportar_analisis
from gantt.reporting.presentation import build_presentation_context
from gantt.reporting.presupuesto_exporter import generar_todos


class BudgetingRunResult:
    """Resultado simple de una ejecucion budgeting."""

    def __init__(self, output_root: Path, outputs: list[Path], warnings: list[str]) -> None:
        self.output_root = output_root
        self.outputs = outputs
        self.warnings = warnings


def run_budgeting(context: ProjectExecutionContext) -> BudgetingRunResult:
    """Ejecuta solo la capability Budgeting en la estructura ENG-3."""
    dirs = budget_output_dirs(context.project_root, source_id_from_path(context.bc3_path))
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
    outputs: list[Path] = []

    document_paths = generar_todos(presupuesto, dirs["documents"], palette=presentation.palette)
    outputs.extend(document_paths)

    analysis_path = dirs["analysis"] / "budget_analysis.xlsx"
    analysis_metadata = budget_document_metadata(context, DocumentType.BUDGET_ANALYSIS).model_copy(
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
        None,
        metadata=analysis_metadata,
        presentation=presentation,
    )
    outputs.append(analysis_path)

    AnalysisChartsReport.from_presupuesto(
        presupuesto,
        dirs["charts"],
        palette=presentation.palette,
    ).generate()
    outputs.extend(sorted(dirs["charts"].rglob("*.png")))

    manifest_context = context.model_copy(update={"output_dir": dirs["root"]})
    write_run_manifest(
        manifest_context,
        outputs,
        [],
        [],
        presentation=presentation,
        capability=Capability.BUDGETING,
    )
    return BudgetingRunResult(dirs["root"], outputs, list(context.warnings))

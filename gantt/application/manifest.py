"""
Manifest estructurado de ejecución.

El manifest permite reconstruir qué entrada produjo cada carpeta de outputs sin
introducir un sistema de logging complejo.
"""

import hashlib
import json
import platform
from datetime import datetime, timezone
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import TYPE_CHECKING

from pydantic import BaseModel, ConfigDict, Field

from gantt.application.context import ProjectExecutionContext
from gantt.application.product import Capability
from gantt.reporting.documents import DOCUMENT_TYPES

if TYPE_CHECKING:
    from gantt.reporting.presentation import PresentationContext


class RunManifest(BaseModel):
    """Contrato del archivo run_manifest.json."""
    model_config = ConfigDict(arbitrary_types_allowed=True)

    project: str
    capability: str | None = None
    run_id: str
    timestamp: str
    python_version: str
    gantt_tools_version: str | None = None
    bc3_file: str
    bc3_original_name: str
    bc3_sha256: str
    config_file: str | None = None
    company: str | None = None
    presentation_theme: str | None = None
    presentation_language: str | None = None
    planning_file: str | None = None
    planning_present: bool
    scenarios: list[str] = Field(default_factory=list)
    outputs: list[str] = Field(default_factory=list)
    document_outputs: list["ManifestDocumentOutput"] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ManifestDocumentOutput(BaseModel):
    """Metadata documental minima por fichero generado."""

    document_type: str
    document_code: str
    capability: str | None = None
    source: str
    path: str


def sha256_file(path: Path) -> str:
    """Calcula SHA-256 de un archivo sin cargarlo completo en memoria."""
    digest = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def relative_to_workspace(path: Path | None, workspace_root: Path) -> str | None:
    """Devuelve una ruta relativa al workspace cuando es posible."""
    if path is None:
        return None
    try:
        return str(path.resolve().relative_to(workspace_root.resolve()))
    except ValueError:
        return str(path)


def obtener_version_gantt_tools() -> str | None:
    """Devuelve la versión instalada si el paquete está instalado."""
    try:
        return version('gantt-tools')
    except PackageNotFoundError:
        return None


def build_run_manifest(
    context: ProjectExecutionContext,
    outputs: list[Path],
    scenarios: list[str],
    warnings: list[str] | None = None,
    timestamp: datetime | None = None,
    presentation: "PresentationContext | None" = None,
    capability: Capability | str | None = None,
) -> RunManifest:
    """Construye el modelo de manifest a partir del contexto y los outputs."""
    ts = timestamp or datetime.now(timezone.utc)
    merged_warnings = list(context.warnings)
    if warnings:
        merged_warnings.extend(warnings)

    relative_outputs = [
        relative_to_workspace(path, context.workspace_root) or str(path)
        for path in outputs
    ]
    return RunManifest(
        project=context.project_name,
        capability=capability.value if isinstance(capability, Capability) else capability,
        run_id=context.run_id,
        timestamp=ts.isoformat(),
        python_version=platform.python_version(),
        gantt_tools_version=obtener_version_gantt_tools(),
        bc3_file=relative_to_workspace(context.bc3_path, context.workspace_root) or '',
        bc3_original_name=context.bc3_path.name,
        bc3_sha256=sha256_file(context.bc3_path),
        config_file=relative_to_workspace(context.config_path, context.workspace_root),
        company=context.company_slug,
        presentation_theme=presentation.theme_id if presentation else None,
        presentation_language=presentation.language if presentation else None,
        planning_file=relative_to_workspace(context.planificacion_path, context.workspace_root),
        planning_present=context.planificacion_path is not None,
        scenarios=scenarios,
        outputs=relative_outputs,
        document_outputs=_document_outputs(outputs, relative_outputs, context, capability),
        warnings=merged_warnings,
    )


def _document_outputs(
    outputs: list[Path],
    relative_outputs: list[str],
    context: ProjectExecutionContext,
    capability: Capability | str | None,
) -> list[ManifestDocumentOutput]:
    by_filename = {info.default_filename: info for info in DOCUMENT_TYPES.values()}
    capability_value = capability.value if isinstance(capability, Capability) else capability
    result: list[ManifestDocumentOutput] = []
    for output, relative_output in zip(outputs, relative_outputs):
        info = by_filename.get(output.name)
        if not info:
            continue
        result.append(
            ManifestDocumentOutput(
                document_type=info.identifier.value,
                document_code=info.document_code,
                capability=capability_value,
                source=context.bc3_path.name,
                path=relative_output,
            )
        )
    return result


def write_run_manifest(
    context: ProjectExecutionContext,
    outputs: list[Path],
    scenarios: list[str],
    warnings: list[str] | None = None,
    presentation: "PresentationContext | None" = None,
    capability: Capability | str | None = None,
) -> Path:
    """Escribe run_manifest.json en la carpeta de salida."""
    manifest = build_run_manifest(
        context,
        outputs,
        scenarios,
        warnings,
        presentation=presentation,
        capability=capability,
    )
    manifest_path = context.output_dir / 'run_manifest.json'
    manifest_path.write_text(
        json.dumps(manifest.model_dump(), ensure_ascii=False, indent=2) + '\n',
        encoding='utf-8',
    )
    return manifest_path

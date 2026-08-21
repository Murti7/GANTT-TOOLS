"""
Resolución del contexto de ejecución de un proyecto.

Centraliza la relación entre workspace, proyecto, BC3 seleccionado, config.yaml,
empresa y carpeta de salida. Esta es infraestructura de aplicación: no parsea el
BC3 ni calcula planificación.
"""

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import yaml
from pydantic import BaseModel, ConfigDict, Field, ValidationError

from gantt.bc3.models import BrandingConfig, ProjectConfig, cargar_company
from gantt.application.project import ProjectModel, resolve_project_model


class ProjectInputConfig(ProjectConfig):
    """
    Contrato validado de config.yaml.

    empresa identifica la empresa emisora y pertenece al contexto de ejecución;
    el resto de campos se transforman en ProjectConfig para los documentos.
    """
    empresa: str | None = None


class ProjectExecutionContext(BaseModel):
    """
    Contexto resuelto de una ejecución de gantt-tools sobre un proyecto.
    """
    model_config = ConfigDict(arbitrary_types_allowed=True)

    workspace_root: Path
    project_name: str
    project_root: Path
    input_dir: Path
    output_dir: Path
    bc3_path: Path
    available_bc3_files: list[Path]
    config_path: Path | None = None
    project_config: ProjectConfig = Field(default_factory=ProjectConfig)
    project_model: ProjectModel | None = None
    project_config_update: dict[str, Any] = Field(default_factory=dict)
    company_slug: str | None = None
    company_config: BrandingConfig | None = None
    planificacion_path: Path | None = None
    run_id: str
    warnings: list[str] = Field(default_factory=list)


def generar_run_id() -> str:
    """Genera un identificador UTC legible y estable para la ejecución."""
    return datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')


def listar_bc3(input_dir: Path) -> list[Path]:
    """Devuelve los BC3 disponibles en input ordenados por nombre."""
    return sorted(input_dir.glob('*.bc3'), key=lambda ruta: ruta.name.lower())


def seleccionar_bc3(input_dir: Path, bc3_filename: str | None = None) -> tuple[Path, list[Path]]:
    """
    Selecciona el BC3 a procesar.

    Si hay varios BC3 exige nombre explícito para evitar ejecuciones ambiguas.
    """
    bc3_files = listar_bc3(input_dir)

    if bc3_filename:
        bc3_path = input_dir / bc3_filename
        if not bc3_path.exists():
            raise FileNotFoundError(
                f'No se encontró el archivo BC3 indicado: {bc3_path}'
            )
        return bc3_path, bc3_files

    if not bc3_files:
        raise FileNotFoundError(f'No se encontró ningún .bc3 en {input_dir}')
    if len(bc3_files) > 1:
        raise ValueError(
            f'Múltiples .bc3 en {input_dir}. '
            f'Especifica el archivo: '
            f'python main.py <nombre-proyecto> <archivo.bc3>\n'
            f'Disponibles: {[f.name for f in bc3_files]}'
        )
    return bc3_files[0], bc3_files


def resolver_output_dir(project_root: Path, bc3_path: Path, bc3_files: list[Path]) -> Path:
    """
    Resuelve la carpeta de salida.

    Con un solo BC3 conserva output/. Con varias versiones, separa por stem BC3.
    """
    project_output_dir = project_root / 'output'
    if len(bc3_files) > 1:
        return project_output_dir / bc3_path.stem
    return project_output_dir


def cargar_config_proyecto(config_path: Path | None) -> tuple[ProjectConfig, dict[str, Any], str | None]:
    """
    Carga config.yaml y separa configuración documental de empresa emisora.
    """
    if config_path is None or not config_path.exists():
        return ProjectConfig(), {}, None

    with open(config_path, encoding='utf-8') as f:
        raw = yaml.safe_load(f) or {}

    try:
        input_config = ProjectInputConfig.model_validate(raw)
    except ValidationError as exc:
        raise ValueError(f'config.yaml no es válido en {config_path}:\n{exc}') from exc

    project_fields = set(ProjectConfig.model_fields)
    update = {key: value for key, value in raw.items() if key in project_fields}
    project_config = ProjectConfig.model_validate(update)
    return project_config, update, input_config.empresa


def resolver_contexto_ejecucion(
    project_name: str,
    bc3_filename: str | None = None,
    workspace_root: Path = Path('.'),
    run_id: str | None = None,
) -> ProjectExecutionContext:
    """
    Construye el contexto completo de ejecución de un proyecto.
    """
    workspace_root = workspace_root.resolve()
    project_root = workspace_root / 'projects' / project_name
    input_dir = project_root / 'input'
    if not input_dir.exists():
        raise FileNotFoundError(f'No se encontró el directorio input del proyecto: {input_dir}')

    bc3_path, bc3_files = seleccionar_bc3(input_dir, bc3_filename)
    output_dir = resolver_output_dir(project_root, bc3_path, bc3_files)

    config_path = input_dir / 'config.yaml'
    config_path = config_path if config_path.exists() else None
    project_config, project_config_update, company_slug = cargar_config_proyecto(config_path)
    project_model, project_model_warnings = resolve_project_model(
        project_root,
        project_config,
        company_slug,
    )
    if project_model.issuer and not company_slug:
        company_slug = project_model.issuer

    company_config = None
    if company_slug:
        company_config = cargar_company(company_slug, workspace_root / 'companies')

    planificacion_path = input_dir / 'planificacion.yaml'
    planificacion_path = planificacion_path if planificacion_path.exists() else None

    return ProjectExecutionContext(
        workspace_root=workspace_root,
        project_name=project_name,
        project_root=project_root,
        input_dir=input_dir,
        output_dir=output_dir,
        bc3_path=bc3_path,
        available_bc3_files=bc3_files,
        config_path=config_path,
        project_config=project_config,
        project_model=project_model,
        project_config_update=project_config_update,
        company_slug=company_slug,
        company_config=company_config,
        planificacion_path=planificacion_path,
        run_id=run_id or generar_run_id(),
        warnings=project_model_warnings,
    )

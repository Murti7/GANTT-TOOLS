"""Modelo canonico de proyecto y compatibilidad con config.yaml legacy."""

from pathlib import Path

import yaml
from pydantic import BaseModel, ConfigDict, Field, ValidationError

from gantt.bc3.models import ProjectConfig
from gantt.billing.models import Client, ClientType


class ProjectIdentity(BaseModel):
    model_config = ConfigDict(extra="forbid")

    code: str = ""
    name: str = ""
    site: str = ""
    revision: str = "00"


class ProjectDefaults(BaseModel):
    model_config = ConfigDict(extra="forbid")

    language: str = "es"


class ProjectDocumentOverrides(BaseModel):
    model_config = ConfigDict(extra="forbid")

    budget: dict[str, str] = Field(default_factory=dict)
    planning: dict[str, str] = Field(default_factory=dict)
    billing: dict[str, str] = Field(default_factory=dict)


class ProjectModel(BaseModel):
    """Configuracion canonica futura de un proyecto."""

    model_config = ConfigDict(extra="forbid")

    project: ProjectIdentity = Field(default_factory=ProjectIdentity)
    client: Client | None = None
    issuer: str | None = None
    defaults: ProjectDefaults = Field(default_factory=ProjectDefaults)
    documents: ProjectDocumentOverrides = Field(default_factory=ProjectDocumentOverrides)
    legacy_config: ProjectConfig | None = None
    source_path: Path | None = None
    is_legacy: bool = False


def client_from_project_yaml(raw: dict) -> Client | None:
    data = raw.get("client")
    if not data:
        return None
    return Client(
        legal_name=data["legal_name"],
        client_type=ClientType(data.get("type", data.get("client_type", "company"))),
        tax_id=data.get("tax_id", ""),
        vat_id=data.get("vat_id", ""),
        fiscal_address=data.get("address", data.get("fiscal_address", "")),
        country=data.get("country", "ES"),
        billing_email=data.get("billing_email", ""),
        contact_person=data.get("contact_person", ""),
        language=data.get("language", "es"),
        contract_reference=data.get("contract_reference", ""),
        purchase_order=data.get("purchase_order", ""),
        expediente=data.get("expediente", ""),
    )


def load_project_yaml(path: Path) -> ProjectModel:
    """Carga project.yaml canonico."""
    raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    try:
        return ProjectModel(
            project=ProjectIdentity.model_validate(raw.get("project", {})),
            client=client_from_project_yaml(raw),
            issuer=raw.get("issuer"),
            defaults=ProjectDefaults.model_validate(raw.get("defaults", {})),
            documents=ProjectDocumentOverrides.model_validate(raw.get("documents", {})),
            source_path=path,
            is_legacy=False,
        )
    except (ValidationError, KeyError, ValueError) as exc:
        raise ValueError(f"project.yaml no es valido en {path}:\n{exc}") from exc


def adapt_legacy_config(config: ProjectConfig, issuer: str | None, path: Path | None = None) -> ProjectModel:
    """Adapta config.yaml legacy al modelo canonico sin duplicar dominio."""
    client = None
    if config.client:
        client = Client(
            legal_name=config.client.legal_name,
            client_type=ClientType(config.client.type),
            tax_id=config.client.tax_id,
            vat_id=config.client.vat_id,
            fiscal_address=config.client.address,
            country=config.client.country,
            billing_email=config.client.billing_email,
            contact_person=config.client.contact_person,
            language=config.client.language,
            contract_reference=config.client.contract_reference,
            purchase_order=config.client.purchase_order,
            expediente=config.client.expediente,
        )
    return ProjectModel(
        project=ProjectIdentity(
            code=config.numero_expediente,
            name=config.proyecto,
            site=config.edificio,
            revision=config.revision,
        ),
        client=client,
        issuer=issuer,
        defaults=ProjectDefaults(language=client.language if client else "es"),
        legacy_config=config,
        source_path=path,
        is_legacy=True,
    )


def resolve_project_model(project_root: Path, legacy_config: ProjectConfig, legacy_issuer: str | None) -> tuple[ProjectModel, list[str]]:
    """Resuelve project.yaml preferente o config.yaml legacy compatible."""
    project_yaml = project_root / "project.yaml"
    if project_yaml.exists():
        return load_project_yaml(project_yaml), []
    config_path = project_root / "input" / "config.yaml"
    warning = (
        "Proyecto usando input/config.yaml legacy; project.yaml sera la fuente canonica futura."
        if config_path.exists()
        else "Proyecto sin project.yaml; usando defaults y configuracion legacy vacia."
    )
    return adapt_legacy_config(legacy_config, legacy_issuer, config_path if config_path.exists() else None), [warning]

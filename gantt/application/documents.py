"""Construccion centralizada de metadata documental desde contexto."""

from datetime import date

from gantt.application.context import ProjectExecutionContext
from gantt.billing.models import Invoice
from gantt.reporting.documents import (
    BillingMetadata,
    BudgetMetadata,
    DocumentMetadata,
    DocumentStatus,
    DocumentType,
    PlanningMetadata,
    document_type_info,
)


def looks_like_imported_source_title(value: str) -> bool:
    normalized = value.strip().lower()
    return normalized.startswith("importado desde") or normalized.startswith("imported from")


def apply_document_identity_fallbacks(presupuesto, context: ProjectExecutionContext):
    """Evita que descripciones tecnicas del BC3 gobiernen titulos documentales."""
    config = presupuesto.config
    project_model = context.project_model
    project = project_model.project if project_model else None

    project_name = (project.name if project else "") or config.proyecto
    if not project_name or looks_like_imported_source_title(project_name):
        project_name = context.project_name

    site_name = (project.site if project else "") or config.edificio or context.project_name
    project_code = (project.code if project else "") or config.numero_expediente or context.bc3_path.stem
    revision = (project.revision if project else "") or config.revision or "00"

    return presupuesto.model_copy(
        update={
            "config": config.model_copy(
                update={
                    "proyecto": project_name,
                    "edificio": site_name,
                    "numero_expediente": project_code,
                    "revision": revision,
                }
            )
        }
    )


def base_metadata(context: ProjectExecutionContext, document_type: DocumentType) -> DocumentMetadata:
    project_model = context.project_model
    project = project_model.project if project_model else None
    client = project_model.client if project_model else None
    info = document_type_info(document_type)
    return DocumentMetadata(
        document_type=document_type,
        title=info.default_title,
        project_code=project.code if project else context.project_config.numero_expediente,
        project_name=project.name if project else context.project_config.proyecto,
        site_name=project.site if project else context.project_config.edificio,
        client_name=client.legal_name if client else context.project_config.entidad,
        reference=context.project_config.numero_expediente,
        revision=project.revision if project else context.project_config.revision,
        issue_date=date.today(),
        source_name=context.bc3_path.name,
    )


def budget_document_metadata(
    context: ProjectExecutionContext,
    document_type: DocumentType,
    *,
    economic_reference: str = "",
) -> DocumentMetadata:
    metadata = base_metadata(context, document_type)
    metadata.budget = BudgetMetadata(
        budget_revision=metadata.revision,
        economic_reference=economic_reference,
    )
    return metadata


def planning_document_metadata(
    context: ProjectExecutionContext,
    document_type: DocumentType,
    *,
    scenario: str = "",
    contractual_deadline: int | None = None,
) -> DocumentMetadata:
    metadata = base_metadata(context, document_type)
    metadata.planning = PlanningMetadata(
        scenario=scenario,
        contractual_deadline=contractual_deadline,
    )
    return metadata


def invoice_document_metadata(
    context: ProjectExecutionContext,
    invoice: Invoice,
) -> DocumentMetadata:
    metadata = base_metadata(context, DocumentType.INVOICE)
    metadata.status = (
        DocumentStatus.ISSUED if invoice.status.value == "issued" else DocumentStatus.DRAFT
    )
    metadata.client_name = invoice.client.legal_name
    metadata.project_name = invoice.project_name or metadata.project_name
    metadata.reference = invoice.project_reference or metadata.reference
    metadata.billing = BillingMetadata(
        invoice_number=invoice.invoice_number,
        invoice_status=invoice.status.value,
        due_date=invoice.payment_terms.due_date,
        billing_preset=invoice.billing_preset,
        billing_source_type=invoice.billing_source_type,
        billing_source_reference=invoice.billing_source_reference,
        economic_basis=invoice.economic_basis,
    )
    return metadata

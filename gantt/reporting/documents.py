"""Sistema documental: tipos, metadata, titulos y nombres de archivo."""

import re
from datetime import date
from enum import StrEnum

from pydantic import BaseModel, ConfigDict, Field


class DocumentFamily(StrEnum):
    BUDGET = "budget"
    PLANNING = "planning"
    BILLING = "billing"


class DocumentPurpose(StrEnum):
    DELIVERY = "delivery"
    ANALYSIS = "analysis"


class DocumentStatus(StrEnum):
    DRAFT = "draft"
    FINAL = "final"
    ISSUED = "issued"


class DocumentType(StrEnum):
    BUDGET_OFFER = "budget_offer"
    UNIT_PRICE_TABLE_1 = "unit_price_table_1"
    UNIT_PRICE_TABLE_2 = "unit_price_table_2"
    DECOMPOSED_BUDGET = "decomposed_budget"
    CHAPTER_SUMMARY = "chapter_summary"
    MEASUREMENTS = "measurements"
    BLIND_MEASUREMENTS = "blind_measurements"
    VEC_SETTLEMENT = "vec_settlement"
    RESOURCE_PRICE_JUSTIFICATION = "resource_price_justification"
    BUDGET_ANALYSIS = "budget_analysis"
    PLANNING_ANALYSIS = "planning_analysis"
    PLANNING_DIAGRAM = "planning_diagram"
    RESOURCE_ANALYSIS = "resource_analysis"
    INVOICE = "invoice"
    CERTIFICATION = "certification"
    FINAL_SETTLEMENT = "final_settlement"


class DocumentTypeInfo(BaseModel):
    model_config = ConfigDict(extra="forbid")

    identifier: DocumentType
    document_code: str
    family: DocumentFamily
    default_title: str
    default_subtitle: str = ""
    default_filename: str


DOCUMENT_TYPES: dict[DocumentType, DocumentTypeInfo] = {
    DocumentType.BUDGET_OFFER: DocumentTypeInfo(identifier=DocumentType.BUDGET_OFFER, document_code="PRES.01", family=DocumentFamily.BUDGET, default_title="Cuadro de oferta", default_filename="PRES.01_Cuadro_Oferta.xlsx"),
    DocumentType.UNIT_PRICE_TABLE_1: DocumentTypeInfo(identifier=DocumentType.UNIT_PRICE_TABLE_1, document_code="PRES.02.01", family=DocumentFamily.BUDGET, default_title="Cuadro de precios n. 1", default_filename="PRES.02.01_Cuadro_Precios_1.xlsx"),
    DocumentType.UNIT_PRICE_TABLE_2: DocumentTypeInfo(identifier=DocumentType.UNIT_PRICE_TABLE_2, document_code="PRES.02.02", family=DocumentFamily.BUDGET, default_title="Cuadro de precios n. 2", default_filename="PRES.02.02_Cuadro_Precios_2.xlsx"),
    DocumentType.DECOMPOSED_BUDGET: DocumentTypeInfo(identifier=DocumentType.DECOMPOSED_BUDGET, document_code="PRES.02.03", family=DocumentFamily.BUDGET, default_title="Presupuesto descompuesto", default_filename="PRES.02.03_Presupuesto_Descompuesto.xlsx"),
    DocumentType.CHAPTER_SUMMARY: DocumentTypeInfo(identifier=DocumentType.CHAPTER_SUMMARY, document_code="PRES.02.04", family=DocumentFamily.BUDGET, default_title="Resumen de capitulos", default_filename="PRES.02.04_Resumen_Capitulos.xlsx"),
    DocumentType.MEASUREMENTS: DocumentTypeInfo(identifier=DocumentType.MEASUREMENTS, document_code="PRES.03.01", family=DocumentFamily.BUDGET, default_title="Mediciones", default_filename="PRES.03.01_Mediciones.xlsx"),
    DocumentType.BLIND_MEASUREMENTS: DocumentTypeInfo(identifier=DocumentType.BLIND_MEASUREMENTS, document_code="PRES.03.02", family=DocumentFamily.BUDGET, default_title="Mediciones ciegas", default_filename="PRES.03.02_Mediciones_Ciegas.xlsx"),
    DocumentType.VEC_SETTLEMENT: DocumentTypeInfo(identifier=DocumentType.VEC_SETTLEMENT, document_code="PRES.05", family=DocumentFamily.BUDGET, default_title="VEC y liquidacion", default_filename="PRES.05_VEC_Liquidacion.xlsx"),
    DocumentType.RESOURCE_PRICE_JUSTIFICATION: DocumentTypeInfo(identifier=DocumentType.RESOURCE_PRICE_JUSTIFICATION, document_code="JUST.01", family=DocumentFamily.BUDGET, default_title="Justificacion de precios de recursos", default_filename="JUST_PRECIOS_Recursos.xlsx"),
    DocumentType.BUDGET_ANALYSIS: DocumentTypeInfo(identifier=DocumentType.BUDGET_ANALYSIS, document_code="BUD.ANALYSIS", family=DocumentFamily.BUDGET, default_title="Analisis economico del presupuesto", default_filename="budget_analysis.xlsx"),
    DocumentType.PLANNING_ANALYSIS: DocumentTypeInfo(identifier=DocumentType.PLANNING_ANALYSIS, document_code="PLN.01", family=DocumentFamily.PLANNING, default_title="Planificacion y analisis de escenarios", default_filename="planning_analysis.xlsx"),
    DocumentType.PLANNING_DIAGRAM: DocumentTypeInfo(identifier=DocumentType.PLANNING_DIAGRAM, document_code="PLN.02", family=DocumentFamily.PLANNING, default_title="Diagrama temporal de planificacion", default_filename="gantt.png"),
    DocumentType.RESOURCE_ANALYSIS: DocumentTypeInfo(identifier=DocumentType.RESOURCE_ANALYSIS, document_code="PLN.03", family=DocumentFamily.PLANNING, default_title="Analisis de recursos", default_filename="resource_analysis.xlsx"),
    DocumentType.INVOICE: DocumentTypeInfo(identifier=DocumentType.INVOICE, document_code="F01", family=DocumentFamily.BILLING, default_title="Factura", default_filename="invoice.xlsx"),
    DocumentType.CERTIFICATION: DocumentTypeInfo(identifier=DocumentType.CERTIFICATION, document_code="F05", family=DocumentFamily.BILLING, default_title="Certificacion parcial", default_filename="certification.xlsx"),
    DocumentType.FINAL_SETTLEMENT: DocumentTypeInfo(identifier=DocumentType.FINAL_SETTLEMENT, document_code="F06", family=DocumentFamily.BILLING, default_title="Liquidacion final", default_filename="final_settlement.xlsx"),
}


class BudgetMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    budget_revision: str = ""
    economic_reference: str = ""


class PlanningMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    scenario: str = ""
    contractual_deadline: int | None = None


class BillingMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    invoice_number: str | None = None
    invoice_status: str = "draft"
    due_date: date | None = None
    billing_preset: str = "F01"
    billing_source_type: str = ""
    billing_source_reference: str = ""
    economic_basis: str = ""


class DocumentMetadata(BaseModel):
    """Identidad documental comun y metadata especifica por composicion."""

    model_config = ConfigDict(extra="forbid")

    document_type: DocumentType
    title: str = ""
    subtitle: str = ""
    project_code: str = ""
    project_name: str = ""
    site_name: str = ""
    client_name: str = ""
    reference: str = ""
    revision: str = "00"
    issue_date: date = Field(default_factory=date.today)
    status: DocumentStatus | None = None
    source_name: str = ""
    source_hash: str = ""
    budget: BudgetMetadata | None = None
    planning: PlanningMetadata | None = None
    billing: BillingMetadata | None = None


def document_type_info(document_type: DocumentType) -> DocumentTypeInfo:
    return DOCUMENT_TYPES[document_type]


ANALYSIS_DOCUMENT_TYPES = {
    DocumentType.BUDGET_ANALYSIS,
    DocumentType.PLANNING_ANALYSIS,
    DocumentType.PLANNING_DIAGRAM,
    DocumentType.RESOURCE_ANALYSIS,
}


def document_purpose(document_type: DocumentType) -> DocumentPurpose:
    return (
        DocumentPurpose.ANALYSIS
        if document_type in ANALYSIS_DOCUMENT_TYPES
        else DocumentPurpose.DELIVERY
    )


def resolve_document_title(
    metadata: DocumentMetadata,
    *,
    explicit_title: str = "",
    configured_title: str = "",
    source_description: str = "",
    filename_fallback: str = "",
) -> str:
    """Politica unica de titulo editorial."""
    info = document_type_info(metadata.document_type)
    for candidate in (
        explicit_title,
        configured_title,
        metadata.title,
        info.default_title,
        source_description,
        filename_fallback,
    ):
        if candidate and candidate.strip():
            return candidate.strip()
    return info.default_title


def safe_filename_part(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("_") or "document"


def filename_for_document(metadata: DocumentMetadata) -> str:
    """Nombre estable, corto y trazable."""
    info = document_type_info(metadata.document_type)
    if metadata.document_type == DocumentType.INVOICE and metadata.billing:
        if metadata.billing.invoice_status == "issued" and metadata.billing.invoice_number:
            preset = safe_filename_part(metadata.billing.billing_preset or "F01")
            return f"{safe_filename_part(metadata.billing.invoice_number)}_{preset}.xlsx"
        preset = safe_filename_part(metadata.billing.billing_preset or "F01")
        source = metadata.billing.invoice_number or metadata.reference or metadata.source_name or "invoice"
        return f"{preset}_DRAFT_{safe_filename_part(source)}.xlsx"
    return info.default_filename

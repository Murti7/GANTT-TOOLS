"""Politicas de layout documental por familia."""

from enum import StrEnum

from pydantic import BaseModel, ConfigDict

from gantt.reporting.documents import DocumentFamily, DocumentMetadata, document_type_info
from gantt.reporting.presentation import PresentationContext


class LayoutFamily(StrEnum):
    BUDGET = "budget"
    PLANNING = "planning"
    INVOICE = "invoice"


class HeaderBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    issuer_name: str = ""
    project_line: str = ""
    site_line: str = ""
    document_line: str = ""
    reference_line: str = ""
    revision_line: str = ""
    date_line: str = ""
    client_line: str = ""


def layout_family_for_document(metadata: DocumentMetadata) -> LayoutFamily:
    family = document_type_info(metadata.document_type).family
    if family == DocumentFamily.BILLING:
        return LayoutFamily.INVOICE
    if family == DocumentFamily.PLANNING:
        return LayoutFamily.PLANNING
    return LayoutFamily.BUDGET


def header_block(metadata: DocumentMetadata, presentation: PresentationContext) -> HeaderBlock:
    """Compone la identidad visible del documento sin consultar filesystem."""
    info = document_type_info(metadata.document_type)
    issuer_name = presentation.company.nombre if presentation.company else ""
    document_line = f"{info.document_code} - {metadata.title or info.default_title}"
    return HeaderBlock(
        issuer_name=issuer_name,
        project_line=metadata.project_name,
        site_line=metadata.site_name,
        document_line=document_line,
        reference_line=metadata.reference or metadata.project_code,
        revision_line=metadata.revision,
        date_line=metadata.issue_date.isoformat(),
        client_line=metadata.client_name,
    )

"""Interfaz de aplicacion para exportar artefactos de facturacion."""

import json
from dataclasses import dataclass
from pathlib import Path
from enum import StrEnum

from pydantic import BaseModel, ConfigDict, Field

from gantt.billing.calculator import calculate_invoice
from gantt.billing.models import Invoice
from gantt.billing.persistence import build_invoice_snapshot, write_invoice_json
from gantt.application.outputs import billing_output_dir
from gantt.reporting.invoice_exporter import export_invoice_excel
from gantt.reporting.documents import DocumentMetadata
from gantt.reporting.presentation import PresentationContext


class InvoiceRunManifest(BaseModel):
    """Manifest no sensible de una ejecucion de facturacion."""

    model_config = ConfigDict(extra="forbid")

    invoice_id: str
    invoice_number: str | None = None
    invoice_status: str
    issuer: str
    client: str
    project: str = ""
    source: str = ""
    billing_preset: str = "F01"
    billing_source_type: str = ""
    billing_source_reference: str = ""
    economic_basis: str = ""
    ready_to_issue: bool = False
    completeness_status: str = "draft_valid"
    missing_required_fields: list[str] = Field(default_factory=list)
    data_hash: str
    outputs: list[str] = Field(default_factory=list)


@dataclass(frozen=True)
class InvoiceExportResult:
    excel_path: Path
    json_path: Path
    manifest_path: Path


class DocumentCompletenessStatus(StrEnum):
    DRAFT_VALID = "draft_valid"
    ISSUE_READY = "issue_ready"
    ISSUE_BLOCKED = "issue_blocked"


def invoice_document_completeness(invoice: Invoice) -> tuple[DocumentCompletenessStatus, list[str]]:
    """Valida campos administrativos visibles sin modificar logica fiscal."""
    missing: list[str] = []
    if not invoice.issuer.tax_id or invoice.issuer.tax_id.upper() == "PENDIENTE":
        missing.append("issuer.tax_id")
    if not invoice.issuer.fiscal_address or invoice.issuer.fiscal_address.upper() == "PENDIENTE":
        missing.append("issuer.address")
    if not (invoice.client.tax_id or invoice.client.vat_id):
        missing.append("client.tax_id")
    if not invoice.client.fiscal_address:
        missing.append("client.address")
    if invoice.status.value == "issued":
        return (
            DocumentCompletenessStatus.ISSUE_BLOCKED if missing else DocumentCompletenessStatus.ISSUE_READY,
            missing,
        )
    return (
        DocumentCompletenessStatus.DRAFT_VALID if missing else DocumentCompletenessStatus.ISSUE_READY,
        missing,
    )


def build_invoice_run_manifest(
    invoice: Invoice,
    outputs: list[Path],
    *,
    source: str = "",
) -> InvoiceRunManifest:
    snapshot = build_invoice_snapshot(invoice)
    completeness_status, missing = invoice_document_completeness(invoice)
    return InvoiceRunManifest(
        invoice_id=invoice.invoice_id,
        invoice_number=invoice.invoice_number,
        invoice_status=invoice.status.value,
        issuer=invoice.issuer.legal_name,
        client=invoice.client.legal_name,
        project=invoice.project_name,
        source=source,
        billing_preset=invoice.billing_preset,
        billing_source_type=invoice.billing_source_type,
        billing_source_reference=invoice.billing_source_reference,
        economic_basis=invoice.economic_basis,
        ready_to_issue=completeness_status == DocumentCompletenessStatus.ISSUE_READY,
        completeness_status=completeness_status.value,
        missing_required_fields=missing,
        data_hash=snapshot.data_hash,
        outputs=[str(path) for path in outputs],
    )


def write_invoice_run_manifest(
    invoice: Invoice,
    outputs: list[Path],
    manifest_path: Path,
    *,
    source: str = "",
) -> Path:
    manifest = build_invoice_run_manifest(invoice, outputs, source=source)
    manifest_path.write_text(
        json.dumps(manifest.model_dump(mode="json"), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return manifest_path


def export_invoice_artifacts(
    invoice: Invoice,
    presentation: PresentationContext,
    output_dir: Path,
    *,
    basename: str = "invoice",
    source: str = "",
    metadata: DocumentMetadata | None = None,
) -> InvoiceExportResult:
    """Exporta Excel, JSON y manifest de una factura."""
    output_dir.mkdir(parents=True, exist_ok=True)
    calculation = calculate_invoice(invoice)

    excel_path = export_invoice_excel(
        invoice,
        calculation,
        presentation,
        output_dir / f"{basename}.xlsx",
        metadata=metadata,
    )
    json_path = write_invoice_json(invoice, output_dir / f"{basename}.json")
    manifest_path = write_invoice_run_manifest(
        invoice,
        [excel_path, json_path],
        output_dir / f"{basename}_manifest.json",
        source=source,
    )
    return InvoiceExportResult(
        excel_path=excel_path,
        json_path=json_path,
        manifest_path=manifest_path,
    )


def export_project_invoice_artifacts(
    invoice: Invoice,
    presentation: PresentationContext,
    project_root: Path,
    *,
    source: str = "",
    metadata: DocumentMetadata | None = None,
) -> InvoiceExportResult:
    """Exporta factura en output/billing/drafts|issued segun estado."""
    preset = invoice.billing_preset or "F01"
    basename = (
        f"{invoice.invoice_number}_{preset}"
        if invoice.invoice_number and invoice.status.value == "issued"
        else f"{preset}_DRAFT_{invoice.invoice_id}"
    )
    return export_invoice_artifacts(
        invoice,
        presentation,
        billing_output_dir(project_root, invoice),
        basename=basename,
        source=source,
        metadata=metadata,
    )

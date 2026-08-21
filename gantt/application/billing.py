"""Interfaz de aplicacion para exportar artefactos de facturacion."""

import json
from dataclasses import dataclass
from pathlib import Path

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
    data_hash: str
    outputs: list[str] = Field(default_factory=list)


@dataclass(frozen=True)
class InvoiceExportResult:
    excel_path: Path
    json_path: Path
    manifest_path: Path


def build_invoice_run_manifest(
    invoice: Invoice,
    outputs: list[Path],
    *,
    source: str = "",
) -> InvoiceRunManifest:
    snapshot = build_invoice_snapshot(invoice)
    return InvoiceRunManifest(
        invoice_id=invoice.invoice_id,
        invoice_number=invoice.invoice_number,
        invoice_status=invoice.status.value,
        issuer=invoice.issuer.legal_name,
        client=invoice.client.legal_name,
        project=invoice.project_name,
        source=source,
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
    return export_invoice_artifacts(
        invoice,
        presentation,
        billing_output_dir(project_root, invoice),
        basename="invoice",
        source=source,
        metadata=metadata,
    )

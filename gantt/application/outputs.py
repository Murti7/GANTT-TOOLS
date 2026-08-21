"""Politica de filesystem de outputs por capability."""

from pathlib import Path

from gantt.application.product import Capability
from gantt.billing.models import Invoice, InvoiceStatus


def source_id_from_path(path: Path) -> str:
    return path.stem


def capability_output_root(project_root: Path, capability: Capability, source_id: str | None = None) -> Path:
    root = project_root / "output" / capability.value
    return root / source_id if source_id else root


def budget_output_dirs(project_root: Path, source_id: str) -> dict[str, Path]:
    root = capability_output_root(project_root, Capability.BUDGETING, source_id)
    return {
        "root": root,
        "documents": root / "documents",
        "analysis": root / "analysis",
        "charts": root / "analysis" / "charts",
    }


def planning_output_dirs(project_root: Path, source_id: str) -> dict[str, Path]:
    root = capability_output_root(project_root, Capability.PLANNING, source_id)
    return {
        "root": root,
        "charts": root / "charts",
    }


def billing_output_dir(project_root: Path, invoice: Invoice) -> Path:
    if invoice.status == InvoiceStatus.ISSUED and invoice.invoice_number:
        return project_root / "output" / Capability.BILLING.value / "issued" / invoice.invoice_number
    return project_root / "output" / Capability.BILLING.value / "drafts" / invoice.invoice_id

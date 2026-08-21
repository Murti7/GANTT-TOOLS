"""Persistencia JSON trazable de facturas."""

import hashlib
import json
import platform
from datetime import datetime, timezone
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from pydantic import BaseModel, ConfigDict

from gantt.billing.calculator import calculate_invoice
from gantt.billing.models import Invoice, InvoiceCalculation


class InvoiceSnapshot(BaseModel):
    """Instantanea reproducible de una factura y su calculo."""

    model_config = ConfigDict(extra="forbid")

    schema_version: str = "1"
    created_at: str
    python_version: str
    gantt_tools_version: str | None = None
    invoice: Invoice
    calculation: InvoiceCalculation
    data_hash: str


def gantt_tools_version() -> str | None:
    try:
        return version("gantt-tools")
    except PackageNotFoundError:
        return None


def invoice_hash(invoice: Invoice, calculation: InvoiceCalculation) -> str:
    payload = {
        "invoice": invoice.model_dump(mode="json"),
        "calculation": calculation.model_dump(mode="json"),
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def build_invoice_snapshot(invoice: Invoice, timestamp: datetime | None = None) -> InvoiceSnapshot:
    calculation = calculate_invoice(invoice)
    ts = timestamp or datetime.now(timezone.utc)
    return InvoiceSnapshot(
        created_at=ts.isoformat(),
        python_version=platform.python_version(),
        gantt_tools_version=gantt_tools_version(),
        invoice=invoice,
        calculation=calculation,
        data_hash=invoice_hash(invoice, calculation),
    )


def write_invoice_json(invoice: Invoice, output_path: Path) -> Path:
    snapshot = build_invoice_snapshot(invoice)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(snapshot.model_dump(mode="json"), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return output_path

"""Servicios de aplicacion para construir facturas desde fuentes del sistema."""

from datetime import timedelta
from decimal import Decimal

from gantt.bc3.economics import FinancialSummary, calcular_resumen_financiero
from gantt.bc3.models import Presupuesto
from gantt.billing.models import (
    BillingBaseKind,
    Client,
    ClientType,
    Currency,
    Invoice,
    InvoiceLine,
    InvoiceStatus,
    Issuer,
    PaymentTerms,
    SourceReference,
    SourceType,
)
from gantt.billing.policies import default_tax_policy


def payment_terms_from_issuer(issuer: Issuer, issue_date=None) -> PaymentTerms:
    due_date = issue_date + timedelta(days=issuer.default_payment_days) if issue_date else None
    return PaymentTerms(
        method=issuer.default_payment_method,
        payment_days=issuer.default_payment_days,
        due_date=due_date,
        iban=issuer.iban,
        bic=issuer.bic,
        account_holder=issuer.account_holder,
    )


def create_invoice(
    issuer: Issuer,
    client: Client,
    lines: list[InvoiceLine],
    *,
    project_name: str = "",
    project_reference: str = "",
    status: InvoiceStatus = InvoiceStatus.DRAFT,
    invoice_number: str | None = None,
    issue_date=None,
    notes: str = "",
) -> Invoice:
    """Crea una factura aplicando politica fiscal y condiciones por defecto."""
    return Invoice(
        issuer=issuer,
        client=client,
        project_name=project_name,
        project_reference=project_reference,
        status=status,
        invoice_number=invoice_number,
        issue_date=issue_date,
        lines=lines,
        tax_policy=default_tax_policy(issuer, client),
        payment_terms=payment_terms_from_issuer(issuer, issue_date),
        notes=notes,
        currency=issuer.currency,
    )


def client_from_project_config(project_config) -> Client:
    """Mapea `ProjectConfig.client` al modelo de dominio Client."""
    if not getattr(project_config, "client", None):
        raise ValueError("El proyecto no define client en config.yaml")
    cfg = project_config.client
    return Client(
        legal_name=cfg.legal_name,
        client_type=ClientType(cfg.type),
        tax_id=cfg.tax_id,
        vat_id=cfg.vat_id,
        fiscal_address=cfg.address,
        country=cfg.country,
        billing_email=cfg.billing_email,
        contact_person=cfg.contact_person,
        language=cfg.language,
        contract_reference=cfg.contract_reference,
        purchase_order=cfg.purchase_order,
        expediente=cfg.expediente,
    )


def line_from_billing_base(base_kind: BillingBaseKind, amount, description: str, source: SourceReference) -> InvoiceLine:
    """Transforma una base economica en una linea facturable."""
    return InvoiceLine(
        code=base_kind.value,
        description=description,
        unit="ud",
        quantity=Decimal("1"),
        unit_price=amount,
        source=source,
    )


def financial_summary_value(summary: FinancialSummary, base_kind: BillingBaseKind) -> Decimal:
    """Extrae una base F02-F04 desde FinancialSummary sin duplicar formulas."""
    mapping = {
        BillingBaseKind.PEM: summary.pem_sin_gr,
        BillingBaseKind.PEM_BI: summary.pem_sin_gr + summary.bi,
        BillingBaseKind.PEC: summary.pec,
    }
    if base_kind not in mapping:
        raise ValueError(f"Base economica no soportada desde BC3: {base_kind}")
    return Decimal(str(mapping[base_kind]))


def invoice_line_from_presupuesto(
    presupuesto: Presupuesto,
    base_kind: BillingBaseKind,
    *,
    description: str | None = None,
    bc3_file: str = "",
) -> InvoiceLine:
    """Crea una linea F02-F04 desde Presupuesto via gantt.bc3.economics."""
    summary = calcular_resumen_financiero(presupuesto)
    amount = financial_summary_value(summary, base_kind)
    label = description or {
        BillingBaseKind.PEM: "Presupuesto de ejecucion material (PEM)",
        BillingBaseKind.PEM_BI: "PEM + beneficio industrial",
        BillingBaseKind.PEC: "Presupuesto de ejecucion por contrata (PEC)",
    }[base_kind]
    return line_from_billing_base(
        base_kind,
        amount,
        label,
        SourceReference(
            source_type=SourceType.BC3_BUDGET,
            reference=base_kind.value,
            project=presupuesto.config.proyecto or presupuesto.descripcion,
            bc3_file=bc3_file,
        ),
    )


def milestone_invoice_line(description: str, contract_amount, percentage, milestone_id: str) -> InvoiceLine:
    """Crea una linea por hito contractual."""
    amount = Decimal(str(contract_amount)) * Decimal(str(percentage))
    return InvoiceLine(
        code=milestone_id,
        description=description,
        unit="%",
        quantity=Decimal("1"),
        unit_price=amount,
        source=SourceReference(
            source_type=SourceType.CONTRACT_MILESTONE,
            reference=f"{Decimal(str(percentage)) * Decimal('100')}%",
            milestone_id=milestone_id,
        ),
    )

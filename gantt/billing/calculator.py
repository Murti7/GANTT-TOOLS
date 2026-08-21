"""Motor determinista de calculo de facturas."""

from decimal import Decimal, ROUND_HALF_UP

from gantt.billing.models import (
    Invoice,
    InvoiceCalculation,
    LineCalculation,
    TaxComponent,
    TaxTreatment,
)

CENT = Decimal("0.01")


def money(value: Decimal) -> Decimal:
    """Redondeo monetario unico: dos decimales, HALF_UP."""
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def calculate_invoice(invoice: Invoice) -> InvoiceCalculation:
    """Calcula subtotal, impuestos, retencion, total factura e importe a cobrar."""
    lines: list[LineCalculation] = []

    for line in invoice.lines:
        gross = money(line.quantity * line.unit_price)
        discount = money(line.discount_amount)
        taxable = money(gross - discount)
        lines.append(
            LineCalculation(
                code=line.code,
                description=line.description,
                gross_amount=gross,
                discount_amount=discount,
                taxable_amount=taxable,
                source=line.source,
            )
        )

    subtotal = money(sum((line.gross_amount for line in lines), Decimal("0")))
    discounts = money(sum((line.discount_amount for line in lines), Decimal("0")))
    taxable_base = money(sum((line.taxable_amount for line in lines), Decimal("0")))

    vat_policy = invoice.tax_policy.vat
    vat_amount = Decimal("0") if vat_policy.treatment in {TaxTreatment.EXEMPT, TaxTreatment.NOT_SUBJECT} else taxable_base * vat_policy.rate
    vat = TaxComponent(
        name="IVA",
        base=taxable_base,
        rate=vat_policy.rate,
        amount=money(vat_amount),
        treatment=vat_policy.treatment,
    )

    withholding_policy = invoice.tax_policy.withholding
    withholding_amount = taxable_base * withholding_policy.rate if withholding_policy.enabled else Decimal("0")
    withholding = TaxComponent(
        name="Retencion",
        base=taxable_base,
        rate=withholding_policy.rate if withholding_policy.enabled else Decimal("0"),
        amount=money(withholding_amount),
    )

    invoice_total = money(taxable_base + vat.amount)
    amount_due = money(invoice_total - withholding.amount)

    return InvoiceCalculation(
        line_calculations=lines,
        subtotal=subtotal,
        discounts=discounts,
        taxable_base=taxable_base,
        vat=vat,
        withholding=withholding,
        invoice_total=invoice_total,
        amount_due=amount_due,
        currency=invoice.currency,
    )

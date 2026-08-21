"""Exporter Excel para facturas.

El exporter representa un Invoice ya calculado. No contiene logica fiscal.
"""

from pathlib import Path
from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from gantt.billing.models import Invoice, InvoiceCalculation, TaxTreatment
from gantt.reporting.documents import BillingMetadata, DocumentMetadata, DocumentStatus, DocumentType
from gantt.reporting.presentation import PresentationContext
from gantt.reporting.rendering import (
    apply_document_footer,
    configure_excel_printing,
    render_invoice_header,
)


THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(bottom=THIN)
MONEY_FMT = '#,##0.00 "EUR"'
QTY_FMT = "#,##0.###"
PCT_FMT = "0.00%"


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color.lstrip("#"))


def _write_label_value(ws, row: int, label: str, value, value_col: int = 2) -> None:
    ws.cell(row, 1).value = label
    ws.cell(row, 1).font = Font(bold=True)
    ws.cell(row, value_col).value = value


def _section(ws, row: int, title: str, end_col: int, context: PresentationContext) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    cell = ws.cell(row, 1)
    cell.value = title
    cell.fill = context.palette.fill_header
    cell.font = context.palette.font_header
    cell.alignment = Alignment(horizontal="center")
    return row + 1


def _format_money_cell(cell) -> None:
    cell.number_format = MONEY_FMT
    cell.alignment = Alignment(horizontal="right")


def export_invoice_excel(
    invoice: Invoice,
    calculation: InvoiceCalculation,
    context: PresentationContext,
    output_path: Path,
    metadata: DocumentMetadata | None = None,
) -> Path:
    """Genera el Excel profesional de factura desde dominio + calculo."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Factura"

    for col, width in {
        "A": 18,
        "B": 42,
        "C": 12,
        "D": 14,
        "E": 16,
        "F": 16,
    }.items():
        ws.column_dimensions[col].width = width

    metadata = metadata or DocumentMetadata(
        document_type=DocumentType.INVOICE,
        title="Factura",
        project_name=invoice.project_name,
        reference=invoice.project_reference,
        status=DocumentStatus.ISSUED if invoice.status.value == "issued" else DocumentStatus.DRAFT,
        billing=BillingMetadata(
            invoice_number=invoice.invoice_number,
            invoice_status=invoice.status.value,
            due_date=invoice.payment_terms.due_date,
        ),
    )
    row = render_invoice_header(ws, metadata, context, 6)
    _write_label_value(ws, row, "Numero", invoice.invoice_number or "BORRADOR")
    _write_label_value(ws, row + 1, "Fecha", invoice.issue_date.isoformat() if invoice.issue_date else "")
    _write_label_value(ws, row + 2, "Proyecto", invoice.project_name)
    _write_label_value(ws, row + 3, "Referencia", invoice.project_reference)

    row += 6
    row = _section(ws, row, "EMISOR", 3, context)
    issuer_rows = [
        ("Nombre", invoice.issuer.legal_name),
        ("NIF/CIF", invoice.issuer.tax_id),
        ("Direccion fiscal", invoice.issuer.fiscal_address),
        ("Email", invoice.issuer.email),
        ("Telefono", invoice.issuer.phone),
    ]
    for label, value in issuer_rows:
        _write_label_value(ws, row, label, value)
        row += 1

    row += 1
    row = _section(ws, row, "CLIENTE", 3, context)
    client_rows = [
        ("Nombre", invoice.client.legal_name),
        ("NIF/CIF/VAT", invoice.client.tax_id or invoice.client.vat_id),
        ("Direccion fiscal", invoice.client.fiscal_address),
        ("Pais", invoice.client.country),
        ("Email facturacion", invoice.client.billing_email),
        ("Pedido/expediente", invoice.client.purchase_order or invoice.client.expediente),
    ]
    for label, value in client_rows:
        _write_label_value(ws, row, label, value)
        row += 1

    row += 1
    row = _section(ws, row, "CONCEPTOS", 6, context)
    headers = ["Codigo", "Descripcion", "Unidad", "Cantidad", "Precio", "Importe"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row, col)
        cell.value = header
        cell.fill = context.palette.fill_subhead
        cell.font = context.palette.font_subhead
        cell.alignment = Alignment(horizontal="center")
    row += 1

    line_by_code_desc = {
        (line.code, line.description): line
        for line in invoice.lines
    }
    for line_calc in calculation.line_calculations:
        line = line_by_code_desc[(line_calc.code, line_calc.description)]
        values = [
            line.code,
            line.description,
            line.unit,
            line.quantity,
            line.unit_price,
            line_calc.taxable_amount,
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row, col)
            cell.value = value
            cell.border = BORDER
            if col in (5, 6):
                _format_money_cell(cell)
            elif col == 4:
                cell.number_format = QTY_FMT
        row += 1

    row += 1
    row = _section(ws, row, "RESUMEN", 6, context)
    summary_rows = [
        ("Subtotal", calculation.subtotal),
        ("Descuentos", calculation.discounts),
        ("Base imponible", calculation.taxable_base),
        (f"IVA ({calculation.vat.rate:.2%})", calculation.vat.amount),
    ]
    if calculation.vat.treatment in {TaxTreatment.EXEMPT, TaxTreatment.NOT_SUBJECT}:
        summary_rows[-1] = (f"IVA ({calculation.vat.treatment.value})", calculation.vat.amount)
    if calculation.withholding.amount:
        summary_rows.append((f"Retencion ({calculation.withholding.rate:.2%})", -calculation.withholding.amount))
    summary_rows.extend(
        [
            ("TOTAL FACTURA", calculation.invoice_total),
            ("LIQUIDO A PERCIBIR", calculation.amount_due),
        ]
    )

    for label, value in summary_rows:
        ws.cell(row, 4).value = label
        ws.cell(row, 4).font = Font(bold=label in {"TOTAL FACTURA", "LIQUIDO A PERCIBIR"})
        ws.cell(row, 6).value = value
        _format_money_cell(ws.cell(row, 6))
        if label in {"TOTAL FACTURA", "LIQUIDO A PERCIBIR"}:
            ws.cell(row, 4).fill = _fill(context.palette.color_fondo_alt)
            ws.cell(row, 6).fill = _fill(context.palette.color_fondo_alt)
            ws.cell(row, 6).font = Font(bold=True)
        row += 1

    row += 1
    row = _section(ws, row, "FORMA DE PAGO", 6, context)
    payment_rows = [
        ("Forma de pago", invoice.payment_terms.method.value),
        ("Vencimiento", invoice.payment_terms.due_date.isoformat() if invoice.payment_terms.due_date else ""),
        ("IBAN", invoice.payment_terms.iban),
        ("BIC", invoice.payment_terms.bic),
        ("Titular", invoice.payment_terms.account_holder),
        ("Referencia", invoice.payment_terms.payment_reference),
    ]
    for label, value in payment_rows:
        _write_label_value(ws, row, label, value)
        row += 1

    if invoice.notes:
        row += 1
        row = _section(ws, row, "NOTAS", 6, context)
        ws.merge_cells(start_row=row, start_column=1, end_row=row + 2, end_column=6)
        ws.cell(row, 1).value = invoice.notes
        ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")

    for row_cells in ws.iter_rows():
        for cell in row_cells:
            if cell.value is not None:
                font = copy(cell.font)
                font.name = context.palette.font_family
                cell.font = font

    apply_document_footer(ws, metadata, context, 6, invoice=True)
    configure_excel_printing(ws, orientation="portrait", repeat_row=None, fit_to_height=1)
    ws.freeze_panes = "A25"
    wb.save(output_path)
    return output_path

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from gantt.bc3.models import BrandingConfig, ProjectConfig, cargar_company
from gantt.bc3.parser import parse_bc3
from gantt.billing.calculator import calculate_invoice
from gantt.billing.models import (
    BillingBaseKind,
    Client,
    ClientType,
    Currency,
    Invoice,
    InvoiceLine,
    InvoiceStatus,
    Issuer,
    TaxPolicy,
    TaxRate,
    TaxTreatment,
)
from gantt.billing.persistence import build_invoice_snapshot
from gantt.billing.services import (
    client_from_project_config,
    create_invoice,
    invoice_line_from_presupuesto,
    milestone_invoice_line,
)


BC3_MINIMO = r"""~C|OBRA##||Proyecto mini|100.00||0|
~D|OBRA##
|A#\\1
\|
~C|A#||Capitulo A|100.00||0|
~D|A#
|A01\\2
\|
~C|A01|ud|Partida de ejemplo|50.00||0|
~D|A01
|MO-OFI1\\3
\|
~C|MO-OFI1|h|Oficial 1a|20.00||1|
"""


def issuer_company() -> Issuer:
    return Issuer(
        legal_name="BAFRAS Engineering S.L.",
        entity_type="company",
        tax_id="B12345678",
        fiscal_address="Calle Demo 1",
        default_vat_rate=Decimal("0.21"),
        default_withholding_rate=Decimal("0"),
        withholding_enabled=False,
        iban="ES0000000000000000000000",
        account_holder="BAFRAS Engineering S.L.",
    )


def issuer_professional() -> Issuer:
    return Issuer(
        legal_name="David Murti",
        entity_type="professional",
        tax_id="12345678Z",
        fiscal_address="Calle Demo 2",
        default_vat_rate=Decimal("0.21"),
        default_withholding_rate=Decimal("0.15"),
        withholding_enabled=True,
        iban="ES0000000000000000000000",
        account_holder="David Murti",
    )


def client_company() -> Client:
    return Client(
        legal_name="Cliente S.A.",
        client_type=ClientType.COMPANY,
        tax_id="A12345678",
        fiscal_address="Avenida Cliente 1",
        billing_email="facturas@example.com",
    )


def test_factura_simple_1000_mas_iva():
    invoice = create_invoice(
        issuer_company(),
        client_company(),
        [InvoiceLine(description="Honorarios profesionales", quantity=1, unit_price=1000)],
    )

    calc = calculate_invoice(invoice)

    assert calc.taxable_base == Decimal("1000.00")
    assert calc.vat.amount == Decimal("210.00")
    assert calc.withholding.amount == Decimal("0.00")
    assert calc.invoice_total == Decimal("1210.00")
    assert calc.amount_due == Decimal("1210.00")


def test_autonomo_con_retencion_distingue_total_y_liquido():
    invoice = create_invoice(
        issuer_professional(),
        client_company(),
        [InvoiceLine(description="Servicios de ingenieria", quantity=1, unit_price=1000)],
    )

    calc = calculate_invoice(invoice)

    assert calc.taxable_base == Decimal("1000.00")
    assert calc.vat.amount == Decimal("210.00")
    assert calc.withholding.amount == Decimal("150.00")
    assert calc.invoice_total == Decimal("1210.00")
    assert calc.amount_due == Decimal("1060.00")


def test_sociedad_sin_retencion_misma_base():
    invoice = create_invoice(
        issuer_company(),
        client_company(),
        [InvoiceLine(description="Servicios de ingenieria", quantity=1, unit_price=1000)],
    )

    assert calculate_invoice(invoice).amount_due == Decimal("1210.00")


def test_varias_lineas_y_redondeo_decimal_half_up():
    invoice = create_invoice(
        issuer_company(),
        client_company(),
        [
            InvoiceLine(description="Linea A", quantity=3, unit_price="19.995"),
            InvoiceLine(description="Linea B", quantity=2, unit_price="10.005"),
        ],
    )

    calc = calculate_invoice(invoice)

    assert [line.taxable_amount for line in calc.line_calculations] == [
        Decimal("59.99"),
        Decimal("20.01"),
    ]
    assert calc.taxable_base == Decimal("80.00")
    assert calc.vat.amount == Decimal("16.80")
    assert calc.invoice_total == Decimal("96.80")


def test_iva_reducido_y_exento_son_configurables():
    invoice = create_invoice(
        issuer_company(),
        client_company(),
        [InvoiceLine(description="Servicio reducido", unit_price=1000)],
    )
    invoice.tax_policy = TaxPolicy(vat=TaxRate(treatment=TaxTreatment.REDUCED, rate=Decimal("0.10")))

    assert calculate_invoice(invoice).vat.amount == Decimal("100.00")

    invoice.tax_policy = TaxPolicy(vat=TaxRate(treatment=TaxTreatment.EXEMPT, rate=Decimal("0")))
    calc = calculate_invoice(invoice)
    assert calc.vat.amount == Decimal("0.00")
    assert calc.invoice_total == Decimal("1000.00")


def test_draft_no_requiere_numero_definitivo():
    invoice = create_invoice(
        issuer_company(),
        client_company(),
        [InvoiceLine(description="Borrador", unit_price=100)],
    )

    assert invoice.status == InvoiceStatus.DRAFT
    assert invoice.invoice_number is None


def test_issued_requiere_numero_fecha_y_datos_fiscales():
    with pytest.raises(ValueError, match="requiere numero"):
        create_invoice(
            issuer_company(),
            client_company(),
            [InvoiceLine(description="Emitida", unit_price=100)],
            status=InvoiceStatus.ISSUED,
            issue_date=date(2026, 1, 15),
        )

    issued = create_invoice(
        issuer_company(),
        client_company(),
        [InvoiceLine(description="Emitida", unit_price=100)],
        status=InvoiceStatus.ISSUED,
        invoice_number="BAF-2026-001",
        issue_date=date(2026, 1, 15),
    )

    assert issued.invoice_number == "BAF-2026-001"


def test_cliente_particular_y_empresa_son_representables():
    particular = Client(legal_name="Persona Particular", client_type=ClientType.INDIVIDUAL)
    empresa = client_company()

    assert particular.tax_id == ""
    assert empresa.tax_id == "A12345678"


def test_facturacion_por_hito_contractual():
    line = milestone_invoice_line(
        "30% adjudicacion",
        contract_amount=Decimal("10000"),
        percentage=Decimal("0.30"),
        milestone_id="H-ADJ",
    )

    invoice = create_invoice(issuer_company(), client_company(), [line])
    calc = calculate_invoice(invoice)

    assert calc.taxable_base == Decimal("3000.00")
    assert invoice.lines[0].source.milestone_id == "H-ADJ"


def test_facturacion_bc3_pem_y_pec_desde_economics(tmp_path: Path):
    bc3 = tmp_path / "mini.bc3"
    bc3.write_text(BC3_MINIMO, encoding="latin-1")
    presupuesto = parse_bc3(bc3)

    pem_line = invoice_line_from_presupuesto(presupuesto, BillingBaseKind.PEM, bc3_file=bc3.name)
    pec_line = invoice_line_from_presupuesto(presupuesto, BillingBaseKind.PEC, bc3_file=bc3.name)

    assert pem_line.unit_price == Decimal("100.0")
    assert pec_line.unit_price == Decimal("119.000")
    assert pec_line.source.source_type == "bc3_budget"


def test_client_from_project_config_soporta_bloque_client():
    config = ProjectConfig(
        client={
            "legal_name": "Ajuntament Demo",
            "tax_id": "P0700000A",
            "type": "public_administration",
            "expediente": "EXP-1",
        }
    )

    client = client_from_project_config(config)

    assert client.client_type == ClientType.PUBLIC_ADMINISTRATION
    assert client.expediente == "EXP-1"


def test_snapshot_json_incluye_hash_y_calculo():
    invoice = create_invoice(
        issuer_company(),
        client_company(),
        [InvoiceLine(description="Snapshot", unit_price=100)],
    )

    snapshot = build_invoice_snapshot(invoice)

    assert snapshot.calculation.invoice_total == Decimal("121.00")
    assert len(snapshot.data_hash) == 64


def test_multiempresa_bafras_y_murti_configuran_retencion_diferente():
    bafras = Issuer.from_company(
        BrandingConfig(nombre="BAFRAS", tipo_entidad="sl", iva_tipo=0.21, irpf_tipo=0.0, moneda="EUR")
    )
    murti = Issuer.from_company(
        BrandingConfig(nombre="David Murti", tipo_entidad="autonomo", iva_tipo=0.21, irpf_tipo=0.15, moneda="EUR")
    )

    bafras_calc = calculate_invoice(create_invoice(bafras, client_company(), [InvoiceLine(description="A", unit_price=1000)]))
    murti_calc = calculate_invoice(create_invoice(murti, client_company(), [InvoiceLine(description="A", unit_price=1000)]))

    assert bafras_calc.withholding.amount == Decimal("0.00")
    assert murti_calc.withholding.amount == Decimal("150.00")

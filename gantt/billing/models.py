"""Modelos de dominio para facturacion.

Los modelos describen la factura y sus conceptos. Los importes derivados se
calculan en `calculator.py`; no deben introducirse manualmente en exporters.
"""

from datetime import date
from decimal import Decimal
from enum import StrEnum
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from gantt.bc3.models import BrandingConfig


class Currency(StrEnum):
    EUR = "EUR"


class ClientType(StrEnum):
    COMPANY = "company"
    PROFESSIONAL = "professional"
    INDIVIDUAL = "individual"
    PUBLIC_ADMINISTRATION = "public_administration"
    FOREIGN = "foreign"


class EntityType(StrEnum):
    COMPANY = "company"
    PROFESSIONAL = "professional"
    PUBLIC_ADMINISTRATION = "public_administration"
    INDIVIDUAL = "individual"
    OTHER = "other"


class InvoiceStatus(StrEnum):
    DRAFT = "draft"
    ISSUED = "issued"
    PARTIALLY_PAID = "partially_paid"
    PAID = "paid"
    OVERDUE = "overdue"
    CANCELLED = "cancelled"
    CREDITED = "credited"


class TaxTreatment(StrEnum):
    STANDARD = "standard"
    REDUCED = "reduced"
    EXEMPT = "exempt"
    NOT_SUBJECT = "not_subject"


class SourceType(StrEnum):
    MANUAL = "manual"
    BC3_BUDGET = "bc3_budget"
    CONTRACT_MILESTONE = "contract_milestone"
    CERTIFICATION = "certification"
    ADVANCE = "advance"
    FINAL_SETTLEMENT = "final_settlement"
    CREDIT_NOTE = "credit_note"


class BillingBaseKind(StrEnum):
    PROFESSIONAL_FEES = "professional_fees"
    PEM = "pem"
    PEM_BI = "pem_bi"
    PEC = "pec"
    CERTIFICATION = "certification"
    FINAL_SETTLEMENT = "final_settlement"
    MANUAL = "manual"


class PaymentMethod(StrEnum):
    BANK_TRANSFER = "bank_transfer"
    DIRECT_DEBIT = "direct_debit"
    CHECK = "check"
    OTHER = "other"


def decimal_from_any(value) -> Decimal:
    """Convierte entradas numericas a Decimal sin pasar por float binario."""
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def require_rate(value: Decimal) -> Decimal:
    if value < Decimal("0") or value > Decimal("1"):
        raise ValueError("El tipo debe estar entre 0 y 1")
    return value


class Issuer(BaseModel):
    """Emisor legal de la factura."""

    model_config = ConfigDict(extra="forbid")

    legal_name: str
    entity_type: EntityType = EntityType.COMPANY
    tax_id: str = ""
    fiscal_address: str = ""
    email: str = ""
    phone: str = ""
    website: str = ""
    iban: str = ""
    bic: str = ""
    bank_name: str = ""
    account_holder: str = ""
    default_payment_days: int = 30
    default_payment_method: PaymentMethod = PaymentMethod.BANK_TRANSFER
    invoice_series: str = "A"
    currency: Currency = Currency.EUR
    default_vat_rate: Decimal = Decimal("0.21")
    default_withholding_rate: Decimal = Decimal("0")
    withholding_enabled: bool = False

    @field_validator("legal_name")
    @classmethod
    def legal_name_required(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("El emisor requiere nombre legal")
        return value.strip()

    @field_validator("default_vat_rate", "default_withholding_rate", mode="before")
    @classmethod
    def parse_rate(cls, value) -> Decimal:
        return require_rate(decimal_from_any(value))

    @classmethod
    def from_company(cls, company: BrandingConfig) -> "Issuer":
        entity_type = (
            EntityType.PROFESSIONAL
            if company.tipo_entidad == "autonomo"
            else EntityType.COMPANY
        )
        return cls(
            legal_name=company.nombre,
            entity_type=entity_type,
            tax_id=company.nif_cif,
            fiscal_address=company.direccion_fiscal,
            email=company.email,
            phone=company.telefono,
            website=company.web,
            iban=company.iban,
            bic=company.bic_swift,
            bank_name=company.entitat_bancaria,
            account_holder=company.titular_compte or company.nombre,
            default_payment_days=company.termini_pagament_dies,
            default_payment_method=PaymentMethod.BANK_TRANSFER,
            invoice_series=company.serie_factura,
            currency=Currency(company.moneda),
            default_vat_rate=decimal_from_any(company.iva_tipo),
            default_withholding_rate=decimal_from_any(company.irpf_tipo),
            withholding_enabled=company.irpf_tipo > 0,
        )


class Client(BaseModel):
    """Cliente contractual o destinatario administrativo."""

    model_config = ConfigDict(extra="forbid")

    legal_name: str
    client_type: ClientType = ClientType.COMPANY
    tax_id: str = ""
    vat_id: str = ""
    fiscal_address: str = ""
    country: str = "ES"
    billing_email: str = ""
    contact_person: str = ""
    language: str = "es"
    contract_reference: str = ""
    purchase_order: str = ""
    expediente: str = ""
    administrative_recipient: "Client | None" = None

    @field_validator("legal_name")
    @classmethod
    def client_name_required(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("El cliente requiere nombre legal")
        return value.strip()


class SourceReference(BaseModel):
    """Referencia trazable al origen economico de una linea."""

    model_config = ConfigDict(extra="forbid")

    source_type: SourceType = SourceType.MANUAL
    reference: str = ""
    project: str = ""
    bc3_file: str = ""
    capitulo_bc3: str = ""
    partida_bc3: str = ""
    milestone_id: str = ""
    certification_id: str = ""


class BillingBase(BaseModel):
    """Base economica que origina una o varias lineas de factura."""

    model_config = ConfigDict(extra="forbid")

    kind: BillingBaseKind
    description: str
    amount: Decimal
    source: SourceReference = Field(default_factory=SourceReference)
    currency: Currency = Currency.EUR

    @field_validator("amount", mode="before")
    @classmethod
    def parse_amount(cls, value) -> Decimal:
        return decimal_from_any(value)


class TaxRate(BaseModel):
    """Tratamiento y tipo de IVA."""

    model_config = ConfigDict(extra="forbid")

    treatment: TaxTreatment = TaxTreatment.STANDARD
    rate: Decimal = Decimal("0.21")
    reason: str = ""

    @field_validator("rate", mode="before")
    @classmethod
    def parse_rate(cls, value) -> Decimal:
        return require_rate(decimal_from_any(value))

    @model_validator(mode="after")
    def validate_treatment(self) -> "TaxRate":
        if self.treatment in {TaxTreatment.EXEMPT, TaxTreatment.NOT_SUBJECT} and self.rate != 0:
            raise ValueError("IVA exento/no sujeto debe tener tipo 0")
        return self


class WithholdingPolicy(BaseModel):
    """Politica configurable de retencion."""

    model_config = ConfigDict(extra="forbid")

    enabled: bool = False
    rate: Decimal = Decimal("0")
    reason: str = ""

    @field_validator("rate", mode="before")
    @classmethod
    def parse_rate(cls, value) -> Decimal:
        return require_rate(decimal_from_any(value))


class TaxPolicy(BaseModel):
    """Politica fiscal aplicable a una factura concreta."""

    model_config = ConfigDict(extra="forbid")

    vat: TaxRate = Field(default_factory=TaxRate)
    withholding: WithholdingPolicy = Field(default_factory=WithholdingPolicy)


class PaymentTerms(BaseModel):
    """Condiciones de pago de la factura."""

    model_config = ConfigDict(extra="forbid")

    method: PaymentMethod = PaymentMethod.BANK_TRANSFER
    payment_days: int = 30
    due_date: date | None = None
    iban: str = ""
    bic: str = ""
    account_holder: str = ""
    notes: str = ""
    payment_reference: str = ""

    @field_validator("payment_days")
    @classmethod
    def non_negative_days(cls, value: int) -> int:
        if value < 0:
            raise ValueError("El plazo de pago no puede ser negativo")
        return value


class InvoiceLine(BaseModel):
    """Linea facturable tipada y trazable."""

    model_config = ConfigDict(extra="forbid")

    code: str = ""
    description: str
    unit: str = "ud"
    quantity: Decimal = Decimal("1")
    unit_price: Decimal
    discount_amount: Decimal = Decimal("0")
    source: SourceReference = Field(default_factory=SourceReference)
    currency: Currency = Currency.EUR

    @field_validator("description")
    @classmethod
    def description_required(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("La linea requiere descripcion")
        return value.strip()

    @field_validator("quantity", "unit_price", "discount_amount", mode="before")
    @classmethod
    def parse_decimal(cls, value) -> Decimal:
        return decimal_from_any(value)

    @model_validator(mode="after")
    def validate_line(self) -> "InvoiceLine":
        if self.quantity == 0:
            raise ValueError("La cantidad no puede ser cero")
        if self.discount_amount < 0:
            raise ValueError("El descuento no puede ser negativo")
        return self


class LineCalculation(BaseModel):
    """Resultado calculado por linea."""

    model_config = ConfigDict(extra="forbid")

    code: str = ""
    description: str
    gross_amount: Decimal
    discount_amount: Decimal
    taxable_amount: Decimal
    source: SourceReference


class TaxComponent(BaseModel):
    """Componente fiscal calculado."""

    model_config = ConfigDict(extra="forbid")

    name: str
    base: Decimal
    rate: Decimal
    amount: Decimal
    treatment: TaxTreatment | None = None


class InvoiceCalculation(BaseModel):
    """Resumen economico oficial de una factura."""

    model_config = ConfigDict(extra="forbid")

    line_calculations: list[LineCalculation]
    subtotal: Decimal
    discounts: Decimal
    taxable_base: Decimal
    vat: TaxComponent
    withholding: TaxComponent
    invoice_total: Decimal
    amount_due: Decimal
    currency: Currency


class Invoice(BaseModel):
    """Factura como objeto de dominio independiente de Excel/PDF."""

    model_config = ConfigDict(extra="forbid")

    invoice_id: str = Field(default_factory=lambda: str(uuid4()))
    status: InvoiceStatus = InvoiceStatus.DRAFT
    invoice_number: str | None = None
    issuer: Issuer
    client: Client
    project_name: str = ""
    project_reference: str = ""
    issue_date: date | None = None
    creation_date: date = Field(default_factory=date.today)
    operation_date: date | None = None
    payment_date: date | None = None
    lines: list[InvoiceLine]
    tax_policy: TaxPolicy = Field(default_factory=TaxPolicy)
    payment_terms: PaymentTerms = Field(default_factory=PaymentTerms)
    notes: str = ""
    currency: Currency = Currency.EUR
    original_invoice_id: str | None = None
    correction_reason: str = ""

    @model_validator(mode="after")
    def validate_invoice(self) -> "Invoice":
        if not self.lines:
            raise ValueError("La factura requiere al menos una linea")
        if any(line.currency != self.currency for line in self.lines):
            raise ValueError("Todas las lineas deben usar la moneda de la factura")
        if self.issuer.currency != self.currency:
            raise ValueError("La moneda del emisor no coincide con la factura")
        if self.payment_terms.due_date and self.issue_date and self.payment_terms.due_date < self.issue_date:
            raise ValueError("El vencimiento no puede ser anterior a la emision")
        if self.status != InvoiceStatus.DRAFT:
            self.validate_for_issue()
        return self

    def validate_for_issue(self) -> None:
        """Validacion estricta para una factura emitida."""
        if not self.invoice_number:
            raise ValueError("Una factura emitida requiere numero")
        if self.issue_date is None:
            raise ValueError("Una factura emitida requiere fecha de emision")
        if not self.issuer.tax_id or self.issuer.tax_id == "PENDIENTE":
            raise ValueError("Una factura emitida requiere identificacion fiscal del emisor")
        if not self.issuer.fiscal_address or self.issuer.fiscal_address == "PENDIENTE":
            raise ValueError("Una factura emitida requiere domicilio fiscal del emisor")
        if not self.client.tax_id and self.client.client_type != ClientType.INDIVIDUAL:
            raise ValueError("Una factura emitida a cliente no particular requiere identificacion fiscal")


class Certification(BaseModel):
    """Valoracion tecnica/economica previa a una factura."""

    model_config = ConfigDict(extra="forbid")

    certification_id: str
    contract_amount: Decimal
    accumulated_executed_amount: Decimal
    previously_certified_amount: Decimal = Decimal("0")
    currency: Currency = Currency.EUR

    @field_validator("contract_amount", "accumulated_executed_amount", "previously_certified_amount", mode="before")
    @classmethod
    def parse_decimal(cls, value) -> Decimal:
        return decimal_from_any(value)

    @property
    def current_certification_amount(self) -> Decimal:
        return self.accumulated_executed_amount - self.previously_certified_amount

    @property
    def executed_percentage(self) -> Decimal:
        if self.contract_amount == 0:
            return Decimal("0")
        return self.accumulated_executed_amount / self.contract_amount


class FinalSettlement(BaseModel):
    """Liquidacion conceptual de contrato."""

    model_config = ConfigDict(extra="forbid")

    settlement_id: str
    original_contract_amount: Decimal
    modifications_amount: Decimal = Decimal("0")
    already_invoiced_amount: Decimal = Decimal("0")
    adjustments_amount: Decimal = Decimal("0")
    currency: Currency = Currency.EUR

    @field_validator(
        "original_contract_amount",
        "modifications_amount",
        "already_invoiced_amount",
        "adjustments_amount",
        mode="before",
    )
    @classmethod
    def parse_decimal(cls, value) -> Decimal:
        return decimal_from_any(value)

    @property
    def final_balance(self) -> Decimal:
        return self.original_contract_amount + self.modifications_amount + self.adjustments_amount - self.already_invoiced_amount

"""Presets F01-F06 como origen facturable, no como motores de factura."""

from enum import StrEnum

from pydantic import BaseModel, ConfigDict

from gantt.billing.models import BillingBaseKind, SourceType


class BillingPresetCode(StrEnum):
    F01 = "F01"
    F02 = "F02"
    F03 = "F03"
    F04 = "F04"
    F05 = "F05"
    F06 = "F06"


class BillingPreset(BaseModel):
    model_config = ConfigDict(extra="forbid")

    code: BillingPresetCode
    name: str
    document_title: str
    document_title_i18n: dict[str, str] = {}
    base_kind: BillingBaseKind
    source_type: SourceType


BILLING_PRESETS: dict[BillingPresetCode, BillingPreset] = {
    BillingPresetCode.F01: BillingPreset(code=BillingPresetCode.F01, name="Professional Fees", document_title="Factura de honorarios profesionales", document_title_i18n={"es": "Factura de honorarios profesionales"}, base_kind=BillingBaseKind.PROFESSIONAL_FEES, source_type=SourceType.MANUAL),
    BillingPresetCode.F02: BillingPreset(code=BillingPresetCode.F02, name="PEM", document_title="Factura sobre presupuesto de ejecucion material", document_title_i18n={"es": "Factura sobre presupuesto de ejecucion material"}, base_kind=BillingBaseKind.PEM, source_type=SourceType.BC3_BUDGET),
    BillingPresetCode.F03: BillingPreset(code=BillingPresetCode.F03, name="PEM + BI", document_title="Factura sobre presupuesto con beneficio industrial", document_title_i18n={"es": "Factura sobre presupuesto con beneficio industrial"}, base_kind=BillingBaseKind.PEM_BI, source_type=SourceType.BC3_BUDGET),
    BillingPresetCode.F04: BillingPreset(code=BillingPresetCode.F04, name="PEM + GG + BI", document_title="Factura sobre presupuesto completo", document_title_i18n={"es": "Factura sobre presupuesto completo"}, base_kind=BillingBaseKind.PEC, source_type=SourceType.BC3_BUDGET),
    BillingPresetCode.F05: BillingPreset(code=BillingPresetCode.F05, name="Partial Certification", document_title="Factura de certificacion parcial", document_title_i18n={"es": "Factura de certificacion parcial"}, base_kind=BillingBaseKind.CERTIFICATION, source_type=SourceType.CERTIFICATION),
    BillingPresetCode.F06: BillingPreset(code=BillingPresetCode.F06, name="Final Settlement", document_title="Factura de liquidacion final", document_title_i18n={"es": "Factura de liquidacion final"}, base_kind=BillingBaseKind.FINAL_SETTLEMENT, source_type=SourceType.FINAL_SETTLEMENT),
}


def billing_preset(code: BillingPresetCode | str) -> BillingPreset:
    return BILLING_PRESETS[BillingPresetCode(code)]

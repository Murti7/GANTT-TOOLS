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
    base_kind: BillingBaseKind
    source_type: SourceType


BILLING_PRESETS: dict[BillingPresetCode, BillingPreset] = {
    BillingPresetCode.F01: BillingPreset(code=BillingPresetCode.F01, name="Professional Fees", base_kind=BillingBaseKind.PROFESSIONAL_FEES, source_type=SourceType.MANUAL),
    BillingPresetCode.F02: BillingPreset(code=BillingPresetCode.F02, name="PEM", base_kind=BillingBaseKind.PEM, source_type=SourceType.BC3_BUDGET),
    BillingPresetCode.F03: BillingPreset(code=BillingPresetCode.F03, name="PEM + BI", base_kind=BillingBaseKind.PEM_BI, source_type=SourceType.BC3_BUDGET),
    BillingPresetCode.F04: BillingPreset(code=BillingPresetCode.F04, name="PEM + GG + BI", base_kind=BillingBaseKind.PEC, source_type=SourceType.BC3_BUDGET),
    BillingPresetCode.F05: BillingPreset(code=BillingPresetCode.F05, name="Partial Certification", base_kind=BillingBaseKind.CERTIFICATION, source_type=SourceType.CERTIFICATION),
    BillingPresetCode.F06: BillingPreset(code=BillingPresetCode.F06, name="Final Settlement", base_kind=BillingBaseKind.FINAL_SETTLEMENT, source_type=SourceType.FINAL_SETTLEMENT),
}


def billing_preset(code: BillingPresetCode | str) -> BillingPreset:
    return BILLING_PRESETS[BillingPresetCode(code)]

from gantt.billing.models import BillingBaseKind, SourceType
from gantt.billing.presets import BillingPresetCode, billing_preset


def test_f01_f06_son_presets_no_motores_duplicados():
    assert billing_preset("F01").base_kind == BillingBaseKind.PROFESSIONAL_FEES
    assert billing_preset(BillingPresetCode.F02).base_kind == BillingBaseKind.PEM
    assert billing_preset("F03").base_kind == BillingBaseKind.PEM_BI
    assert billing_preset("F04").base_kind == BillingBaseKind.PEC
    assert billing_preset("F05").source_type == SourceType.CERTIFICATION
    assert billing_preset("F06").source_type == SourceType.FINAL_SETTLEMENT

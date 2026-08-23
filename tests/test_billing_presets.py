from gantt.billing.models import BillingBaseKind, SourceType
from gantt.billing.presets import BillingPresetCode, billing_preset


def test_f01_f06_son_presets_no_motores_duplicados():
    assert billing_preset("F01").base_kind == BillingBaseKind.PROFESSIONAL_FEES
    assert billing_preset("F01").document_title == "Factura de honorarios profesionales"
    assert billing_preset(BillingPresetCode.F02).base_kind == BillingBaseKind.PEM
    assert billing_preset("F02").document_title == "Factura sobre presupuesto de ejecucion material"
    assert billing_preset("F03").base_kind == BillingBaseKind.PEM_BI
    assert billing_preset("F03").document_title == "Factura sobre presupuesto con beneficio industrial"
    assert billing_preset("F04").base_kind == BillingBaseKind.PEC
    assert billing_preset("F04").document_title == "Factura sobre presupuesto completo"
    assert billing_preset("F05").source_type == SourceType.CERTIFICATION
    assert billing_preset("F05").document_title == "Factura de certificacion parcial"
    assert billing_preset("F06").source_type == SourceType.FINAL_SETTLEMENT
    assert billing_preset("F06").document_title == "Factura de liquidacion final"

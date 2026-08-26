from datetime import date

from gantt.reporting.documents import (
    BillingMetadata,
    DocumentMetadata,
    DocumentType,
    DocumentFamily,
    document_type_info,
    filename_for_document,
    resolve_document_title,
)


def test_document_type_catalogo_incluye_identidad_documental():
    info = document_type_info(DocumentType.UNIT_PRICE_TABLE_1)

    assert info.document_code == "PRES.02.01"
    assert info.family == DocumentFamily.BUDGET
    assert info.default_title == "Cuadro de precios n. 1"


def test_title_resolution_prioriza_titulo_explicito_y_no_source_description():
    metadata = DocumentMetadata(
        document_type=DocumentType.UNIT_PRICE_TABLE_1,
        project_name="Auditoria energetica integral",
        source_name="BAF-VID-2026-001_VC.bc3",
    )

    title = resolve_document_title(
        metadata,
        explicit_title="PRES.02.01 - Cuadro de precios n. 1",
        source_description="Importado desde ficheros CSV",
    )

    assert title == "PRES.02.01 - Cuadro de precios n. 1"


def test_title_resolution_usa_default_antes_que_descripcion_bc3():
    metadata = DocumentMetadata(document_type=DocumentType.BUDGET_OFFER)

    title = resolve_document_title(
        metadata,
        source_description="Importado desde ficheros CSV",
        filename_fallback="archivo.bc3",
    )

    assert title == "Cuadro de oferta"


def test_filename_policy_budget_y_invoice_draft_issued():
    budget = DocumentMetadata(document_type=DocumentType.CHAPTER_SUMMARY)
    measurements = DocumentMetadata(document_type=DocumentType.MEASUREMENTS)
    blind_measurements = DocumentMetadata(document_type=DocumentType.BLIND_MEASUREMENTS)
    draft = DocumentMetadata(
        document_type=DocumentType.INVOICE,
        reference="invoice-123",
        billing=BillingMetadata(invoice_status="draft"),
    )
    issued = DocumentMetadata(
        document_type=DocumentType.INVOICE,
        billing=BillingMetadata(invoice_status="issued", invoice_number="BAF-2026-001"),
    )

    assert filename_for_document(budget) == "PRES.02.04_Resumen_Capitulos.xlsx"
    assert filename_for_document(blind_measurements) == "PRES.03.01_Mediciones_Ciegas.xlsx"
    assert filename_for_document(measurements) == "PRES.03.02_Mediciones.xlsx"
    assert filename_for_document(draft) == "F01_DRAFT_invoice-123.xlsx"
    assert filename_for_document(issued) == "BAF-2026-001_F01.xlsx"


def test_metadata_por_familia_no_obliga_campos_ajenos():
    metadata = DocumentMetadata(
        document_type=DocumentType.PLANNING_ANALYSIS,
        title="Planificacion",
        project_code="BAF-VID-2026-001",
        issue_date=date(2026, 8, 21),
    )

    assert metadata.billing is None
    assert metadata.budget is None

from datetime import date
from pathlib import Path

from gantt.application.context import resolver_contexto_ejecucion
from gantt.application.documents import (
    apply_document_identity_fallbacks,
    budget_document_metadata,
    invoice_document_metadata,
    planning_document_metadata,
)
from gantt.bc3.models import Presupuesto
from gantt.billing.models import Client, ClientType, InvoiceLine, Issuer
from gantt.billing.services import create_invoice
from gantt.reporting.documents import DocumentType
from gantt.reporting.layouts import LayoutFamily, header_block, layout_family_for_document
from gantt.reporting.presentation import build_presentation_context


def _context(tmp_path: Path):
    project_root = tmp_path / "projects" / "Demo"
    input_dir = project_root / "input"
    input_dir.mkdir(parents=True)
    (input_dir / "demo.bc3").write_text("", encoding="latin-1")
    (input_dir / "config.yaml").write_text(
        """
empresa: demo
entidad: Cliente Legacy
proyecto: Auditoria energetica integral
edificio: Viding Fitness Calvia
numero_expediente: BAF-VID-2026-001
revision: "00"
""",
        encoding="utf-8",
    )
    company_dir = tmp_path / "companies" / "demo"
    company_dir.mkdir(parents=True)
    (company_dir / "company.yaml").write_text(
        """
identitat:
  nombre: BAFRAS Engineering S.L.
  tipo_entidad: sl
fiscal: {}
bancari: {}
branding: {}
facturacio: {}
""",
        encoding="utf-8",
    )
    return resolver_contexto_ejecucion("Demo", workspace_root=tmp_path)


def test_budget_header_block_procede_de_metadata_y_presentation(tmp_path: Path):
    context = _context(tmp_path)
    metadata = budget_document_metadata(context, DocumentType.UNIT_PRICE_TABLE_1)
    presentation = build_presentation_context(context.company_config)

    header = header_block(metadata, presentation)

    assert layout_family_for_document(metadata) == LayoutFamily.BUDGET
    assert header.issuer_name == "BAFRAS Engineering S.L."
    assert header.project_line == "Auditoria energetica integral"
    assert header.site_line == "Viding Fitness Calvia"
    assert header.document_line == "PRES.02.01 - Cuadro de precios n. 1"
    assert header.reference_line == "BAF-VID-2026-001"
    assert header.revision_line == "00"
    assert "Importado desde ficheros CSV" not in header.document_line


def test_planning_metadata_especifica_escenario_y_plazo(tmp_path: Path):
    context = _context(tmp_path)
    metadata = planning_document_metadata(
        context,
        DocumentType.PLANNING_ANALYSIS,
        scenario="base",
        contractual_deadline=90,
    )

    assert layout_family_for_document(metadata) == LayoutFamily.PLANNING
    assert metadata.planning is not None
    assert metadata.planning.scenario == "base"
    assert metadata.planning.contractual_deadline == 90


def test_invoice_metadata_usa_cliente_y_estado_de_invoice(tmp_path: Path):
    context = _context(tmp_path)
    issuer = Issuer(legal_name="BAFRAS Engineering S.L.", tax_id="B123", fiscal_address="Calle")
    client = Client(legal_name="Viding Fitness", client_type=ClientType.COMPANY, tax_id="B456")
    invoice = create_invoice(
        issuer,
        client,
        [InvoiceLine(description="Servicio", unit_price=100)],
        project_name="Auditoria energetica integral",
        project_reference="BAF-VID-2026-001",
    )
    metadata = invoice_document_metadata(context, invoice)

    assert layout_family_for_document(metadata) == LayoutFamily.INVOICE
    assert metadata.client_name == "Viding Fitness"
    assert metadata.billing is not None
    assert metadata.billing.invoice_status == "draft"


def test_document_identity_fallback_evita_descripcion_importada_como_titulo(tmp_path: Path):
    context = _context(tmp_path)
    presupuesto = Presupuesto(
        codigo="OBRA",
        descripcion="Importado desde ficheros CSV",
        importe_total=0,
        capitulos=[],
        recursos_mo={},
    )

    saneado = apply_document_identity_fallbacks(presupuesto, context)

    assert saneado.config.proyecto == "Auditoria energetica integral"
    assert "Importado desde" not in saneado.config.proyecto

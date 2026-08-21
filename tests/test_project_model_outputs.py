from decimal import Decimal
from datetime import date
from pathlib import Path

import pytest

from gantt.application.outputs import budget_output_dirs, billing_output_dir, planning_output_dirs
from gantt.application.project import ProjectModel, adapt_legacy_config, load_project_yaml, resolve_project_model
from gantt.bc3.models import ProjectConfig
from gantt.billing.models import Client, ClientType, InvoiceLine, Issuer
from gantt.billing.services import create_invoice


def test_project_yaml_valido_carga_cliente_issuer_defaults_y_documentos(tmp_path: Path):
    path = tmp_path / "project.yaml"
    path.write_text(
        """
project:
  code: BAF-VID-2026-001
  name: Auditoria energetica integral
  site: Viding Fitness Calvia
  revision: "00"
client:
  legal_name: Viding Fitness
  tax_id: B00000000
  type: company
issuer: bafras-engineering
defaults:
  language: es
documents:
  budget:
    title: Presupuesto de auditoria energetica
""",
        encoding="utf-8",
    )

    model = load_project_yaml(path)

    assert model.project.code == "BAF-VID-2026-001"
    assert model.client is not None
    assert model.client.legal_name == "Viding Fitness"
    assert model.issuer == "bafras-engineering"
    assert model.documents.budget["title"] == "Presupuesto de auditoria energetica"
    assert not model.is_legacy


def test_project_yaml_invalido_falla(tmp_path: Path):
    path = tmp_path / "project.yaml"
    path.write_text("client:\n  tax_id: X\n", encoding="utf-8")

    with pytest.raises(ValueError, match="project.yaml no es valido"):
        load_project_yaml(path)


def test_legacy_config_se_adapta_con_warning(tmp_path: Path):
    project_root = tmp_path / "projects" / "Demo"
    (project_root / "input").mkdir(parents=True)
    (project_root / "input" / "config.yaml").write_text("empresa: demo\n", encoding="utf-8")
    config = ProjectConfig(proyecto="Proyecto legacy", edificio="Edificio", numero_expediente="EXP-1")

    model, warnings = resolve_project_model(project_root, config, "demo")

    assert model.is_legacy
    assert model.project.name == "Proyecto legacy"
    assert model.issuer == "demo"
    assert warnings


def test_output_policy_por_capability():
    root = Path("projects") / "Demo"

    budget = budget_output_dirs(root, "source-a")
    planning = planning_output_dirs(root, "source-a")

    assert budget["documents"] == root / "output" / "budgeting" / "source-a" / "documents"
    assert budget["analysis"] == root / "output" / "budgeting" / "source-a" / "analysis"
    assert planning["root"] == root / "output" / "planning" / "source-a"


def test_billing_output_draft_vs_issued():
    issuer = Issuer(legal_name="Demo S.L.", tax_id="B123", fiscal_address="Calle")
    client = Client(legal_name="Cliente", client_type=ClientType.COMPANY, tax_id="A123")
    invoice = create_invoice(issuer, client, [InvoiceLine(description="Servicio", unit_price=100)])
    root = Path("projects") / "Demo"

    assert billing_output_dir(root, invoice) == root / "output" / "billing" / "drafts" / invoice.invoice_id

    issued = create_invoice(
        issuer,
        client,
        [InvoiceLine(description="Servicio", unit_price=Decimal("100"))],
        status="issued",
        invoice_number="BAF-2026-001",
        issue_date=date(2026, 1, 15),
    )
    assert billing_output_dir(root, issued) == root / "output" / "billing" / "issued" / "BAF-2026-001"

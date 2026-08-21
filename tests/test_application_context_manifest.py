"""
Tests del contexto de ejecución y manifest.
"""

from datetime import datetime, timezone
from pathlib import Path

import pytest

from gantt.application.context import cargar_config_proyecto, resolver_contexto_ejecucion
from gantt.application.manifest import build_run_manifest
from gantt.bc3.models import cargar_company


def _write_company(root: Path) -> None:
    company_dir = root / 'companies' / 'demo'
    company_dir.mkdir(parents=True)
    (company_dir / 'company.yaml').write_text(
        """
identitat:
  nombre: Demo Engineering
  tipo_entidad: sl
fiscal: {}
bancari: {}
branding:
  color_primario: "#111111"
facturacio:
  idioma: es
""",
        encoding='utf-8',
    )


def test_resolver_contexto_ejecucion_separa_output_por_bc3_si_hay_varios(tmp_path: Path):
    _write_company(tmp_path)
    input_dir = tmp_path / 'projects' / 'Proyecto' / 'input'
    input_dir.mkdir(parents=True)
    (input_dir / 'a.bc3').write_text('', encoding='latin-1')
    (input_dir / 'b.bc3').write_text('', encoding='latin-1')
    (input_dir / 'config.yaml').write_text(
        'empresa: demo\nproyecto: Proyecto demo\nporcentaje_gg: 0.17\n',
        encoding='utf-8',
    )

    context = resolver_contexto_ejecucion(
        'Proyecto',
        'b.bc3',
        workspace_root=tmp_path,
        run_id='RUN-1',
    )

    assert context.bc3_path == input_dir / 'b.bc3'
    assert context.output_dir == tmp_path / 'projects' / 'Proyecto' / 'output' / 'b'
    assert context.company_slug == 'demo'
    assert context.company_config is not None
    assert context.project_config_update == {
        'proyecto': 'Proyecto demo',
        'porcentaje_gg': 0.17,
    }
    assert not hasattr(context.project_config, 'empresa')


def test_cargar_config_proyecto_rechaza_claves_desconocidas(tmp_path: Path):
    config_path = tmp_path / 'config.yaml'
    config_path.write_text('empresa: demo\ncampo_inventado: true\n', encoding='utf-8')

    with pytest.raises(ValueError, match='campo_inventado'):
        cargar_config_proyecto(config_path)


def test_cargar_company_valida_company_yaml(tmp_path: Path):
    _write_company(tmp_path)

    company = cargar_company('demo', tmp_path / 'companies')

    assert company.nombre == 'Demo Engineering'
    assert company.color_primario == '#111111'
    assert company.idioma == 'es'


def test_build_run_manifest_incluye_hash_y_rutas_relativas(tmp_path: Path):
    input_dir = tmp_path / 'projects' / 'Proyecto' / 'input'
    output_dir = tmp_path / 'projects' / 'Proyecto' / 'output'
    input_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True)
    bc3_path = input_dir / 'a.bc3'
    bc3_path.write_text('contenido', encoding='latin-1')

    context = resolver_contexto_ejecucion(
        'Proyecto',
        workspace_root=tmp_path,
        run_id='RUN-1',
    )
    output = output_dir / 'resultado.xlsx'
    output.write_text('', encoding='utf-8')

    manifest = build_run_manifest(
        context,
        [output],
        ['base'],
        timestamp=datetime(2026, 1, 1, tzinfo=timezone.utc),
    )

    assert manifest.project == 'Proyecto'
    assert manifest.bc3_original_name == 'a.bc3'
    assert manifest.bc3_sha256 == (
        '6f9566ef46386b8cf372671cb9eddff5488256eff5ea5fb99e1041c3e27082bf'
    )
    assert manifest.outputs == ['projects\\Proyecto\\output\\resultado.xlsx'] or manifest.outputs == [
        'projects/Proyecto/output/resultado.xlsx'
    ]
    assert manifest.scenarios == ['base']

"""
Tests del enrutado de salidas del pipeline principal.

Responsabilidad: verificar que los proyectos con varias versiones BC3 no
mezclan sus documentos generados en la misma carpeta output.
"""

from pathlib import Path

import pytest

from main import resolver_bc3_path, resolver_output_dir


def test_resolver_output_dir_usa_output_base_si_hay_un_solo_bc3(tmp_path: Path):
    input_dir = tmp_path / 'input'
    input_dir.mkdir()
    bc3_path = input_dir / 'presupuesto.bc3'
    bc3_path.write_text('', encoding='latin-1')

    output_dir = resolver_output_dir('Proyecto', bc3_path, [bc3_path])

    assert output_dir == Path('projects') / 'Proyecto' / 'output'


def test_resolver_output_dir_crea_subcarpeta_si_hay_varios_bc3(tmp_path: Path):
    input_dir = tmp_path / 'input'
    input_dir.mkdir()
    bc3_path = input_dir / 'BAF-VID-2026-001_V2.2.bc3'
    otro_bc3_path = input_dir / 'BAF-VID-2026-001 PT.bc3'
    bc3_path.write_text('', encoding='latin-1')
    otro_bc3_path.write_text('', encoding='latin-1')

    output_dir = resolver_output_dir('Viding Fitness Calvià', bc3_path, [bc3_path, otro_bc3_path])

    assert output_dir == (
        Path('projects')
        / 'Viding Fitness Calvià'
        / 'output'
        / 'BAF-VID-2026-001_V2.2'
    )


def test_resolver_bc3_path_exige_nombre_si_hay_varios_bc3(tmp_path: Path):
    input_dir = tmp_path / 'input'
    input_dir.mkdir()
    (input_dir / 'a.bc3').write_text('', encoding='latin-1')
    (input_dir / 'b.bc3').write_text('', encoding='latin-1')

    with pytest.raises(ValueError, match='Múltiples .bc3'):
        resolver_bc3_path(input_dir)


def test_resolver_bc3_path_devuelve_bc3_indicado_y_lista_disponible(tmp_path: Path):
    input_dir = tmp_path / 'input'
    input_dir.mkdir()
    elegido = input_dir / 'b.bc3'
    (input_dir / 'a.bc3').write_text('', encoding='latin-1')
    elegido.write_text('', encoding='latin-1')

    bc3_path, bc3_files = resolver_bc3_path(input_dir, 'b.bc3')

    assert bc3_path == elegido
    assert [ruta.name for ruta in bc3_files] == ['a.bc3', 'b.bc3']

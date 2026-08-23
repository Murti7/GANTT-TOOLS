from pathlib import Path
import re
import zipfile

from openpyxl import load_workbook

from gantt.bc3.models import Capitulo, Partida, Presupuesto, ProjectConfig, RecursoMO
from gantt.reporting.presupuesto_exporter import (
    generar_pres0301_mediciones,
    generar_pres0302_mediciones_ciegas,
)
from gantt.reporting.xlsx_integrity import validate_xlsx_integrity


def _presupuesto_pres03() -> Presupuesto:
    partida = Partida(
        codigo="01.01",
        descripcion="Levantamiento inicial",
        descripcion_larga=(
            "Servicio de recopilacion y caracterizacion tecnica.\n\n"
            "Incluye:\n"
            "- Revision documental\n"
            "- Mediciones en m2 y validacion de datos"
        ),
        unidad="m2",
        precio_unitario=1234.56,
        cantidad=0.125,
        lineas_mo=[],
    )
    subpartida = Partida(
        codigo="01.02",
        descripcion="Inspeccion complementaria",
        descripcion_larga="Descripcion tecnica complementaria.",
        unidad="ud",
        precio_unitario=4321.0,
        cantidad=1.5,
        lineas_mo=[],
    )
    subcapitulo = Capitulo(
        codigo="01.1#",
        descripcion="Subcapitulo tecnico",
        importe_total=9876.54,
        partidas=[subpartida],
        subcapitulos=[],
    )
    capitulo = Capitulo(
        codigo="01#",
        descripcion="Informacion tecnica",
        importe_total=9876.54,
        partidas=[partida],
        subcapitulos=[subcapitulo],
    )
    return Presupuesto(
        codigo="OBRA##",
        descripcion="Presupuesto demo",
        importe_total=9876.54,
        capitulos=[capitulo],
        recursos_mo={"MO-01": RecursoMO(codigo="MO-01", descripcion="Oficial especialista", precio_hora=88.88)},
        descompuestos_raw={"01.01": [("MO-01", 2.0)]},
        config=ProjectConfig(
            proyecto="Auditoria energetica integral",
            edificio="Viding Fitness Calvia",
            entidad="Cliente demo",
            numero_expediente="BAF-VID-2026-001",
            revision="00",
        ),
    )


def _values(path: Path) -> tuple[list, str]:
    wb = load_workbook(path, data_only=False)
    try:
        ws = wb.active
        values = [cell.value for row in ws.iter_rows() for cell in row if cell.value is not None]
        return values, "\n".join(str(value) for value in values)
    finally:
        wb.close()


def test_pres0302_mediciones_ciegas_contenido_layout_integridad_y_confidencialidad(tmp_path: Path):
    path = tmp_path / "PRES.03.02_Mediciones_Ciegas.xlsx"

    generar_pres0302_mediciones_ciegas(_presupuesto_pres03(), path)

    report = validate_xlsx_integrity(path)
    assert report.passed, report.issues

    wb = load_workbook(path, data_only=False)
    try:
        assert wb.sheetnames == ["MEDICIONES"]
        ws = wb["MEDICIONES"]
        values = [cell.value for row in ws.iter_rows() for cell in row if cell.value is not None]
        text = "\n".join(str(value) for value in values)

        assert "MEDICIONES CIEGAS" in values
        assert "PRES.03.02" in values
        assert "CAPÍTULO 01 — Informacion tecnica" in values
        assert "01.1 — Subcapitulo tecnico" in values
        assert "01.01" in values
        assert "Levantamiento inicial" in values
        assert "Servicio de recopilacion" in text
        assert "m2" in values
        assert 0.125 in values
        assert 1.5 in values

        forbidden_terms = [
            "P. unit.",
            "Importe",
            "Precio unitario",
            "TOTAL CAPÍTULO",
            "PEM",
            "PEC",
            "GG",
            "BI",
            "IVA",
            "MO",
            "MT",
            "MQ",
            "PA",
        ]
        for term in forbidden_terms:
            assert not re.search(rf"(?<!\w){re.escape(term)}(?!\w)", text, flags=re.IGNORECASE)
        assert "1234.56" not in text
        assert "9876.54" not in text

        assert ws.freeze_panes is None
        assert str(ws.page_setup.paperSize) == str(ws.PAPERSIZE_A4)
        assert ws.page_setup.orientation == "portrait"
        assert ws.print_title_rows
        assert not any(dim.hidden for dim in ws.column_dimensions.values())
        assert not any(dim.hidden for dim in ws.row_dimensions.values())
        assert not any(isinstance(value, str) and value.startswith("=") for value in values)
    finally:
        wb.close()

    with zipfile.ZipFile(path) as archive:
        serialized = b"\n".join(archive.read(name) for name in archive.namelist() if name.endswith(".xml"))
    assert b"1234.56" not in serialized
    assert b"9876.54" not in serialized
    assert b"88.88" not in serialized


def test_pres0301_mediciones_recupera_comportamiento_historico(tmp_path: Path):
    path = tmp_path / "PRES.03.01_Mediciones.xlsx"

    generar_pres0301_mediciones(_presupuesto_pres03(), path)

    report = validate_xlsx_integrity(path)
    assert report.passed, report.issues

    wb = load_workbook(path, data_only=False)
    try:
        ws = wb["Mediciones"]
        values = [cell.value for row in ws.iter_rows() for cell in row if cell.value is not None]
        text = "\n".join(str(value) for value in values)

        assert "MEDICIONES" in values
        assert "PRES.03.01" in values
        assert "CAPÍTULO 01 — Informacion tecnica" in values
        assert "01.01" in values
        assert "Levantamiento inicial" in values
        assert "Servicio de recopilacion" in text
        assert "MO" in values
        assert "MO-01 — Oficial especialista" in values
        assert "TOTAL CAPÍTULO 01: Informacion tecnica" in values
        assert "P. unit." not in text
        assert "Importe" not in text
        assert ws.freeze_panes is None
        assert str(ws.page_setup.paperSize) == str(ws.PAPERSIZE_A4)
        assert ws.page_setup.orientation == "portrait"
        assert ws.print_title_rows
    finally:
        wb.close()


def test_pres0301_y_pres0302_son_documentos_distintos_desde_mismo_presupuesto(tmp_path: Path):
    presupuesto = _presupuesto_pres03()
    mediciones = tmp_path / "PRES.03.01_Mediciones.xlsx"
    ciegas = tmp_path / "PRES.03.02_Mediciones_Ciegas.xlsx"

    generar_pres0301_mediciones(presupuesto, mediciones)
    generar_pres0302_mediciones_ciegas(presupuesto, ciegas)

    valores_mediciones, texto_mediciones = _values(mediciones)
    valores_ciegas, texto_ciegas = _values(ciegas)

    assert mediciones.name != ciegas.name
    assert "PRES.03.01" in valores_mediciones
    assert "PRES.03.02" in valores_ciegas
    assert "01.01" in valores_mediciones
    assert "01.01" in valores_ciegas
    assert "m2" in valores_mediciones
    assert "m2" in valores_ciegas
    assert 0.125 in valores_mediciones
    assert 0.125 in valores_ciegas
    assert texto_mediciones != texto_ciegas
    assert "MO-01 — Oficial especialista" in valores_mediciones
    assert "MO-01 — Oficial especialista" not in valores_ciegas
    assert "TOTAL CAPÍTULO 01: Informacion tecnica" in valores_mediciones
    assert "TOTAL CAPÍTULO 01: Informacion tecnica" not in valores_ciegas
    assert "1234.56" not in texto_ciegas

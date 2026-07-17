"""
Tests del parser BC3.

Responsabilidad: verificar que parse_bc3 construye correctamente la
estructura Presupuesto/Capitulo/Partida a partir de ficheros BC3 mínimos,
y que valida la presencia del código raíz antes de continuar.
"""

from pathlib import Path

import pytest

from gantt.bc3.parser import parse_bc3

# Formato de continuación de un ~D: primera línea con '|', siguientes con '\',
# pares "codigo\\cantidad" (doble barra invertida) y terminador "\|".
# Ver gantt/bc3/parser.py::read_bc3_records para el detalle de parseo.

BC3_MINIMO = r"""~C|OBRA##||Proyecto de ejemplo|100.00||0|
~D|OBRA##
|A#\\1
\|
~C|A#||Capitulo A|100.00||0|
~D|A#
|A01\\2
\|
~C|A01|ud|Partida de ejemplo|50.00||0|
~D|A01
|MO-OFI1\\3
\|
~C|MO-OFI1|h|Oficial 1a|20.00||1|
"""

BC3_SIN_MO = r"""~C|OBRA##||Proyecto sin MO|50.00||0|
~D|OBRA##
|B#\\1
\|
~C|B#||Capitulo B|50.00||0|
~D|B#
|B01\\1
\|
~C|B01|kg|Partida sin mano de obra|10.00||0|
~D|B01
|MT-CEMENTO\\5
\|
~C|MT-CEMENTO|kg|Cemento|2.00||3|
"""

BC3_SIN_RAIZ = r"""~C|A#||Capitulo A|100.00||0|
~D|A#
|A01\\2
\|
~C|A01|ud|Partida de ejemplo|50.00||0|
"""


def escribir_bc3(tmp_path: Path, contenido: str) -> Path:
    ruta = tmp_path / 'test.bc3'
    ruta.write_text(contenido, encoding='latin-1')
    return ruta


def test_parse_bc3_minimo_un_capitulo_una_partida_un_recurso_mo(tmp_path):
    ruta = escribir_bc3(tmp_path, BC3_MINIMO)

    presupuesto = parse_bc3(ruta)

    assert presupuesto.codigo == 'OBRA##'
    assert presupuesto.importe_total == 100.0
    assert len(presupuesto.capitulos) == 1

    capitulo = presupuesto.capitulos[0]
    assert capitulo.codigo == 'A#'
    assert len(capitulo.partidas) == 1
    assert capitulo.subcapitulos == []

    partida = capitulo.partidas[0]
    assert partida.codigo == 'A01'
    assert partida.cantidad == 2.0
    assert partida.precio_unitario == 50.0
    assert len(partida.lineas_mo) == 1
    assert partida.lineas_mo[0].codigo_recurso == 'MO-OFI1'
    assert partida.lineas_mo[0].cantidad == 3.0

    assert len(presupuesto.recursos_mo) == 1
    recurso = presupuesto.recursos_mo['MO-OFI1']
    assert recurso.descripcion == 'Oficial 1a'
    assert recurso.precio_hora == 20.0


def test_parse_bc3_sin_recursos_mo(tmp_path):
    ruta = escribir_bc3(tmp_path, BC3_SIN_MO)

    presupuesto = parse_bc3(ruta)

    assert presupuesto.recursos_mo == {}
    assert len(presupuesto.recursos_mt) == 1
    assert 'MT-CEMENTO' in presupuesto.recursos_mt

    partida = presupuesto.capitulos[0].partidas[0]
    assert partida.codigo == 'B01'
    assert partida.lineas_mo == []


def test_parse_bc3_sin_codigo_raiz_lanza_error_con_contexto(tmp_path):
    ruta = escribir_bc3(tmp_path, BC3_SIN_RAIZ)

    with pytest.raises(ValueError) as exc_info:
        parse_bc3(ruta)

    # El error debe incluir la ruta del fichero para poder identificarlo.
    assert str(ruta) in str(exc_info.value)

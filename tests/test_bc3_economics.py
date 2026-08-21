"""
Tests de cálculos económicos críticos.
"""

from gantt.bc3.economics import calcular_resumen_financiero
from gantt.bc3.models import (
    Capitulo,
    LineaDescompuesto,
    Partida,
    Presupuesto,
    ProjectConfig,
    RecursoElemental,
    RecursoMO,
)


def test_calcular_resumen_financiero_con_residuos_y_liquidacion():
    partida_liquidable = Partida(
        codigo='P01',
        descripcion='Partida liquidable',
        unidad='ud',
        precio_unitario=100.0,
        cantidad=2.0,
        lineas_mo=[LineaDescompuesto(codigo_recurso='MO-1', cantidad=1.0)],
    )
    partida_no_liquidable = Partida(
        codigo='P02',
        descripcion='Solo mano de obra',
        unidad='ud',
        precio_unitario=50.0,
        cantidad=1.0,
        lineas_mo=[LineaDescompuesto(codigo_recurso='MO-1', cantidad=1.0)],
    )
    presupuesto = Presupuesto(
        codigo='OBRA##',
        descripcion='Proyecto',
        importe_total=300.0,
        capitulos=[
            Capitulo(
                codigo='01#',
                descripcion='Obra',
                importe_total=250.0,
                partidas=[partida_liquidable, partida_no_liquidable],
                subcapitulos=[],
            ),
            Capitulo(
                codigo='GR#',
                descripcion='Residuos',
                importe_total=50.0,
                partidas=[],
                subcapitulos=[],
            ),
        ],
        recursos_mo={'MO-1': RecursoMO(codigo='MO-1', descripcion='Oficial', precio_hora=20.0)},
        recursos_mt={'MT-1': RecursoElemental(codigo='MT-1', descripcion='Material', unidad='ud', precio_unidad=10.0)},
        descompuestos_raw={
            'P01': [('MO-1', 1.0), ('MT-1', 2.0)],
            'P02': [('MO-1', 1.0)],
        },
        config=ProjectConfig(
            porcentaje_gg=0.10,
            porcentaje_bi=0.05,
            iva_obra=0.21,
            iva_gr=0.10,
            porcentaje_liquidacion=0.10,
            codigo_capitulo_gr='GR#',
        ),
    )

    resumen = calcular_resumen_financiero(presupuesto)

    assert resumen.importe_gr == 50.0
    assert resumen.pem_sin_gr == 250.0
    assert resumen.gg == 25.0
    assert resumen.bi == 12.5
    assert resumen.pec == 287.5
    assert resumen.iva_obra == 60.375
    assert resumen.iva_gr == 5.0
    assert resumen.pgl == 402.875
    assert resumen.pem_liq == 200.0
    assert resumen.gg_liq == 20.0
    assert resumen.bi_liq == 10.0
    assert resumen.pec_liq == 230.0
    assert resumen.liq_max == 28.0
    assert resumen.vec == 365.5

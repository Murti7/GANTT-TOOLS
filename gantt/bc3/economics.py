"""
Cálculos económicos del presupuesto.

Este módulo contiene reglas de dominio financiero que los exporters solo deben
representar en documentos.
"""

from pydantic import BaseModel

from gantt.bc3.models import Capitulo, Presupuesto
from gantt.bc3.parser import expandir_descompuesto


class FinancialSummary(BaseModel):
    """Resultado de la cascada económica principal del presupuesto."""

    importe_gr: float
    pem_sin_gr: float
    gg: float
    bi: float
    pec: float
    iva_obra: float
    iva_gr: float
    pgl: float
    pem_liq: float
    gg_liq: float
    bi_liq: float
    pec_liq: float
    liq_max: float
    vec: float


def iter_partidas_presupuesto(presupuesto: Presupuesto):
    """Recorre todas las partidas del presupuesto en orden jerárquico."""
    def recurrir(raiz, nodo, sub_actual):
        for partida in nodo.partidas:
            yield raiz, sub_actual, partida
        for sub in nodo.subcapitulos:
            yield from recurrir(raiz, sub, sub)

    for cap in presupuesto.capitulos:
        yield from recurrir(cap, cap, None)


def capitulo_gestion_residuos(presupuesto: Presupuesto) -> Capitulo | None:
    """Localiza el capítulo de gestión de residuos configurado."""
    return next(
        (c for c in presupuesto.capitulos if c.codigo == presupuesto.config.codigo_capitulo_gr),
        None,
    )


def tipo_recurso_presupuesto(codigo: str, presupuesto: Presupuesto) -> str:
    """Clasifica un código de recurso usando el presupuesto como fuente de verdad."""
    if codigo.startswith('%'):
        return '%'
    if codigo in presupuesto.recursos_mo:
        return 'MO'
    if codigo in presupuesto.recursos_mt:
        return 'MT'
    if codigo in presupuesto.recursos_mq:
        return 'MQ'
    if codigo in presupuesto.recursos_pa:
        return 'PA'
    return '?'


def recursos_expandidos_presupuesto(codigo: str, presupuesto: Presupuesto) -> dict[str, float]:
    """Expande descompuestos auxiliares para cálculos de dominio."""
    return expandir_descompuesto(
        codigo,
        presupuesto.descompuestos_raw,
        set(presupuesto.recursos_pa),
    )


def es_partida_liquidable_presupuesto(codigo: str, presupuesto: Presupuesto) -> bool:
    """
    Una partida es liquidable si combina mano de obra y material.
    """
    recursos = recursos_expandidos_presupuesto(codigo, presupuesto)
    tiene_mo = any(tipo_recurso_presupuesto(r, presupuesto) == 'MO' for r in recursos)
    tiene_mt = any(tipo_recurso_presupuesto(r, presupuesto) == 'MT' for r in recursos)
    return tiene_mo and tiene_mt


def calcular_resumen_financiero(presupuesto: Presupuesto) -> FinancialSummary:
    """Calcula la cascada PEM, PEC, PGL, VEC y liquidación máxima."""
    cap_gr = capitulo_gestion_residuos(presupuesto)
    importe_gr = cap_gr.importe_total if cap_gr else 0.0
    pem_sin_gr = presupuesto.importe_total - importe_gr

    gg = pem_sin_gr * presupuesto.config.porcentaje_gg
    bi = pem_sin_gr * presupuesto.config.porcentaje_bi
    pec = pem_sin_gr + gg + bi
    iva_obra = pec * presupuesto.config.iva_obra
    iva_gr = importe_gr * presupuesto.config.iva_gr
    pgl = pec + importe_gr + iva_obra + iva_gr

    pem_liq = sum(
        partida.precio_unitario * partida.cantidad
        for _cap, _sub, partida in iter_partidas_presupuesto(presupuesto)
        if es_partida_liquidable_presupuesto(partida.codigo, presupuesto)
    )
    gg_liq = pem_liq * presupuesto.config.porcentaje_gg
    bi_liq = pem_liq * presupuesto.config.porcentaje_bi
    pec_liq = pem_liq + gg_liq + bi_liq
    liq_max = (pec_liq + importe_gr) * presupuesto.config.porcentaje_liquidacion
    vec = pec + importe_gr + liq_max

    return FinancialSummary(
        importe_gr=importe_gr,
        pem_sin_gr=pem_sin_gr,
        gg=gg,
        bi=bi,
        pec=pec,
        iva_obra=iva_obra,
        iva_gr=iva_gr,
        pgl=pgl,
        pem_liq=pem_liq,
        gg_liq=gg_liq,
        bi_liq=bi_liq,
        pec_liq=pec_liq,
        liq_max=liq_max,
        vec=vec,
    )

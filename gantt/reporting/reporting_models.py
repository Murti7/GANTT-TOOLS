"""Modelos intermedios para graficos de reporting.

Esta capa evita que los PNG dependan del Excel ya exportado. El Excel sigue
existiendo como salida y como fuente legacy, pero los graficos pueden renderizar
directamente desde el modelo de dominio Presupuesto.
"""

from dataclasses import dataclass, field

from gantt.bc3.models import Capitulo, Partida, Presupuesto


@dataclass(frozen=True)
class ChapterCostSummary:
    codigo: str
    descripcion: str
    importe_total: float
    horas_mo_total: float
    horas_por_perfil: dict[str, float]
    is_main: bool


@dataclass(frozen=True)
class PartidaCostSummary:
    codigo: str
    descripcion: str
    importe: float


@dataclass(frozen=True)
class BudgetAnalysisData:
    chapters: list[ChapterCostSummary]
    partidas: list[PartidaCostSummary]
    profiles: list[str] = field(default_factory=list)

    def resumen_headers(self) -> list[str]:
        return [
            "Codigo",
            "Descripcion",
            "Importe total",
            "H MO total",
            *self.profiles,
        ]

    def resumen_rows(self) -> list[dict]:
        rows: list[dict] = []
        for chapter in self.chapters:
            row = {
                "Codigo": chapter.codigo,
                "Descripcion": chapter.descripcion,
                "Importe total": chapter.importe_total,
                "H MO total": chapter.horas_mo_total,
                "_is_main": chapter.is_main,
            }
            row.update({profile: chapter.horas_por_perfil.get(profile, 0.0) for profile in self.profiles})
            rows.append(row)
        return rows

    def partidas_headers(self) -> list[str]:
        return ["Codigo partida", "Descripcion", "Importe"]

    def partidas_rows(self) -> list[dict]:
        return [
            {
                "Codigo partida": partida.codigo,
                "Descripcion": partida.descripcion,
                "Importe": partida.importe,
            }
            for partida in self.partidas
        ]


def _horas_capitulo(capitulo: Capitulo, codigos: list[str]) -> dict[str, float]:
    horas = {codigo: 0.0 for codigo in codigos}
    for partida in capitulo.partidas:
        for linea in partida.lineas_mo:
            if linea.codigo_recurso in horas:
                horas[linea.codigo_recurso] += linea.cantidad * partida.cantidad
    for subcapitulo in capitulo.subcapitulos:
        for codigo, valor in _horas_capitulo(subcapitulo, codigos).items():
            horas[codigo] += valor
    return horas


def _iter_partidas(capitulo: Capitulo) -> list[Partida]:
    partidas = list(capitulo.partidas)
    for subcapitulo in capitulo.subcapitulos:
        partidas.extend(_iter_partidas(subcapitulo))
    return partidas


def build_budget_analysis_data(presupuesto: Presupuesto) -> BudgetAnalysisData:
    """Extrae los datos de reporting economico y MO desde Presupuesto."""
    recursos = list(presupuesto.recursos_mo.values())
    codigos = [recurso.codigo for recurso in recursos]
    profiles = [recurso.descripcion for recurso in recursos]
    profile_by_code = {recurso.codigo: recurso.descripcion for recurso in recursos}

    chapters: list[ChapterCostSummary] = []
    partidas: list[PartidaCostSummary] = []

    for capitulo in presupuesto.capitulos:
        horas_codigo = _horas_capitulo(capitulo, codigos)
        chapters.append(
            ChapterCostSummary(
                codigo=capitulo.codigo,
                descripcion=capitulo.descripcion,
                importe_total=capitulo.importe_total,
                horas_mo_total=sum(horas_codigo.values()),
                horas_por_perfil={
                    profile_by_code[codigo]: horas
                    for codigo, horas in horas_codigo.items()
                },
                is_main=True,
            )
        )

        for subcapitulo in capitulo.subcapitulos:
            horas_sub = _horas_capitulo(subcapitulo, codigos)
            chapters.append(
                ChapterCostSummary(
                    codigo=subcapitulo.codigo,
                    descripcion=subcapitulo.descripcion,
                    importe_total=subcapitulo.importe_total,
                    horas_mo_total=sum(horas_sub.values()),
                    horas_por_perfil={
                        profile_by_code[codigo]: horas
                        for codigo, horas in horas_sub.items()
                    },
                    is_main=False,
                )
            )

        for partida in _iter_partidas(capitulo):
            partidas.append(
                PartidaCostSummary(
                    codigo=partida.codigo,
                    descripcion=partida.descripcion,
                    importe=partida.precio_unitario * partida.cantidad,
                )
            )

    return BudgetAnalysisData(chapters=chapters, partidas=partidas, profiles=profiles)

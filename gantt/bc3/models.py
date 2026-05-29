"""
Modelos Pydantic que representan la estructura de un archivo BC3 (FIEBDC-3).

Responsabilidad: definir los tipos de datos del presupuesto tal como aparecen
en el formato estándar BC3 de la construcción española.
"""

from pydantic import BaseModel, Field


class RecursoMO(BaseModel):
    """
    Recurso de mano de obra definido en el BC3.
    Representa un tipo de operario con su coste por hora.
    """

    codigo: str
    descripcion: str
    precio_hora: float


class RecursoElemental(BaseModel):
    """Recurso elemental del presupuesto: material, o maquinaria con unidad y precio."""

    codigo: str
    descripcion: str
    unidad: str
    precio_unidad: float


class LineaDescompuesto(BaseModel):
    """
    Línea de un descompuesto BC3.
    Representa un recurso o subconcepto con su cantidad asignada.
    """

    codigo_recurso: str
    cantidad: float


class Partida(BaseModel):
    """
    Unidad de obra del presupuesto BC3.
    Contiene su composición de mano de obra con horas por unidad
    y la cantidad total de unidades a ejecutar en obra.
    """

    codigo: str
    descripcion: str
    descripcion_larga: str = ''         # descripción técnica completa del registro ~T del BC3
    unidad: str
    precio_unitario: float
    cantidad: float                     # unidades a ejecutar (del ~D del capítulo padre)
    lineas_mo: list[LineaDescompuesto]  # recursos MO con h/unidad


class Capitulo(BaseModel):
    """
    Capítulo o subcapítulo del presupuesto BC3.
    Agrupa partidas de obra bajo un código y descripción comunes.
    Puede contener subcapítulos (estructura jerárquica).
    """

    codigo: str
    descripcion: str
    descripcion_larga: str = ''         # descripción técnica completa del registro ~T del BC3
    importe_total: float
    partidas: list[Partida]
    subcapitulos: list['Capitulo']


class Presupuesto(BaseModel):
    """
    Presupuesto completo parseado desde un archivo BC3.
    Contiene todos los capítulos principales y el catálogo
    de recursos de mano de obra disponibles.
    """

    codigo: str
    descripcion: str
    importe_total: float
    capitulos: list[Capitulo]
    recursos_mo: dict[str, RecursoMO]       # clave: codigo MO
    recursos_mt: dict[str, RecursoElemental] = Field(default_factory=dict)  # materiales
    recursos_mq: dict[str, RecursoElemental] = Field(default_factory=dict)  # maquinaria
    descompuestos_raw: dict[str, list[tuple[str, float]]] = Field(default_factory=dict)


Capitulo.model_rebuild()

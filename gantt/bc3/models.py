"""
Modelos Pydantic que representan la estructura de un archivo BC3 (FIEBDC-3).

Responsabilidad: definir los tipos de datos del presupuesto tal como aparecen
en el formato estándar BC3 de la construcción española.
"""

from pathlib import Path

import yaml
from pydantic import BaseModel, ConfigDict, Field


class ProjectConfig(BaseModel):
    """
    Configuración completa del proyecto: metadatos documentales y parámetros
    financieros del contrato. Se carga desde config.yaml en input/ del proyecto.
    Los campos con valor por defecto permiten funcionar sin yaml.
    Los campos de texto vacíos se omiten en el encabezado documental.
    """

    model_config = ConfigDict(extra='forbid')

    # ── Metadatos documentales ────────────────────────────────────────────────
    entidad:            str = ''   # entidad contratante
    proyecto:           str = ''   # nombre o descripción del proyecto
    edificio:           str = ''   # ubicación o nombre del inmueble
    numero_expediente:  str = ''   # número de expediente de licitación
    footer_org:         str = ''   # texto izquierdo del pie de página
    footer_exp:         str = ''   # texto central del pie de página
    revision:           str = '00' # revisión del documento generado

    # ── Parámetros financieros del contrato ───────────────────────────────────
    porcentaje_gg:          float = 0.13   # gastos generales
    porcentaje_bi:          float = 0.06   # beneficio industrial
    iva_obra:               float = 0.21   # IVA aplicable a la obra
    iva_gr:                 float = 0.10   # IVA aplicable a gestión de residuos
    porcentaje_liquidacion: float = 0.10   # porcentaje para cálculo de liquidación máxima

    # ── Estructura del presupuesto ────────────────────────────────────────────
    codigo_capitulo_gr:  str = ''   # código BC3 del capítulo de gestión de residuos
    client: 'ProjectClientConfig | None' = None


class ProjectClientConfig(BaseModel):
    """Cliente contractual configurado en input/config.yaml."""

    model_config = ConfigDict(extra='forbid')

    legal_name: str
    tax_id: str = ''
    vat_id: str = ''
    address: str = ''
    country: str = 'ES'
    type: str = 'company'
    billing_email: str = ''
    contact_person: str = ''
    language: str = 'es'
    contract_reference: str = ''
    purchase_order: str = ''
    expediente: str = ''


ProjectConfig.model_rebuild()


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


class BrandingConfig(BaseModel):
    """
    Configuración de branding corporativo cargada de company.yaml.
    Todos los campos tienen valor por defecto para garantizar
    compatibilidad con proyectos sin empresa configurada.
    """
    model_config = ConfigDict(extra='forbid')

    # Identidad
    nombre: str = "Sin empresa"
    tipo_entidad: str = "sl"
    nif_cif: str = ""
    direccion_fiscal: str = ""
    telefono: str = ""
    email: str = ""
    web: str = ""

    # Fiscal
    irpf_tipo: float = 0.0
    iva_tipo: float = 0.21
    irpf_inicio_actividad: bool = False
    nota_irpf: str = ""

    # Bancario
    iban: str = ""
    bic_swift: str = ""
    entitat_bancaria: str = ""
    titular_compte: str = ""
    termini_pagament_dies: int = 30
    forma_pagament: str = "transferencia"

    # Branding visual
    color_primario: str = "#1B3A5C"
    color_secundario: str = "#2E6DA4"
    color_acento: str = "#4F7B6E"
    color_texto: str = "#1A1C1E"
    color_fondo_alt: str = "#F5F7FA"
    color_blanco: str = "#FFFFFF"
    fuente_principal: str = "Calibri"
    fuente_fallback: str = "Calibri"
    logo_path: Path | None = None

    # Facturación
    serie_factura: str = "A"
    moneda: str = "EUR"
    idioma: str = "es"


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
    recursos_pa: dict[str, RecursoElemental] = Field(default_factory=dict)  # partidas alzadas / unidades auxiliares con descompuesto propio
    descompuestos_raw: dict[str, list[tuple[str, float]]] = Field(default_factory=dict)
    config: ProjectConfig = Field(default_factory=ProjectConfig)
    company: BrandingConfig | None = None


Capitulo.model_rebuild()


class CompanyIdentityConfig(BaseModel):
    """Bloque de identidad corporativa de company.yaml."""
    model_config = ConfigDict(extra='forbid')

    nombre: str
    tipo_entidad: str
    nif_cif: str = ''
    direccion_fiscal: str = ''
    telefono: str = ''
    email: str = ''
    web: str = ''


class CompanyFiscalConfig(BaseModel):
    """Bloque fiscal de company.yaml."""
    model_config = ConfigDict(extra='forbid')

    tipo_entidad: str = ''
    irpf_tipo: float = 0.0
    iva_tipo: float = 0.21
    irpf_inicio_actividad: bool = False
    nota_irpf: str = ''


class CompanyBankConfig(BaseModel):
    """Bloque bancario de company.yaml."""
    model_config = ConfigDict(extra='forbid')

    iban: str = ''
    bic_swift: str = ''
    entitat_bancaria: str = ''
    titular_compte: str = ''
    termini_pagament_dies: int = 30
    forma_pagament: str = 'transferencia'


class CompanyBrandingYamlConfig(BaseModel):
    """Bloque visual de company.yaml antes de resolver la ruta del logo."""
    model_config = ConfigDict(extra='forbid')

    color_primario: str = '#1B3A5C'
    color_secundario: str = '#2E6DA4'
    color_acento: str = '#4F7B6E'
    color_texto: str = '#1A1C1E'
    color_fondo_alt: str = '#F5F7FA'
    color_blanco: str = '#FFFFFF'
    fuente_principal: str = 'Calibri'
    fuente_fallback: str = 'Calibri'
    logo: str = 'logo.png'


class CompanyBillingConfig(BaseModel):
    """Bloque de facturación de company.yaml."""
    model_config = ConfigDict(extra='forbid')

    serie_factura: str = 'A'
    moneda: str = 'EUR'
    idioma: str = 'es'


class CompanyYamlConfig(BaseModel):
    """Contrato validado del archivo company.yaml."""
    model_config = ConfigDict(extra='forbid')

    identitat: CompanyIdentityConfig
    fiscal: CompanyFiscalConfig = Field(default_factory=CompanyFiscalConfig)
    bancari: CompanyBankConfig = Field(default_factory=CompanyBankConfig)
    branding: CompanyBrandingYamlConfig = Field(default_factory=CompanyBrandingYamlConfig)
    facturacio: CompanyBillingConfig = Field(default_factory=CompanyBillingConfig)


def cargar_company(empresa_slug: str, companies_dir: Path) -> BrandingConfig:
    """
    Carga el company.yaml de la empresa indicada.
    Lanza FileNotFoundError si no existe la carpeta o el yaml.
    """
    company_dir = companies_dir / empresa_slug
    yaml_path = company_dir / 'company.yaml'

    if not yaml_path.exists():
        raise FileNotFoundError(
            f'No se encontró company.yaml en {company_dir}\n'
            f'Empresas disponibles: '
            f'{[d.name for d in companies_dir.iterdir() if d.is_dir()]}'
        )

    with open(yaml_path, encoding='utf-8') as f:
        data = CompanyYamlConfig.model_validate(yaml.safe_load(f) or {})

    logo_filename = data.branding.logo
    logo_path = company_dir / logo_filename
    logo_path = logo_path if logo_path.exists() else None

    return BrandingConfig(
        nombre=data.identitat.nombre,
        tipo_entidad=data.identitat.tipo_entidad,
        nif_cif=data.identitat.nif_cif,
        direccion_fiscal=data.identitat.direccion_fiscal,
        telefono=data.identitat.telefono,
        email=data.identitat.email,
        web=data.identitat.web,
        irpf_tipo=data.fiscal.irpf_tipo,
        iva_tipo=data.fiscal.iva_tipo,
        irpf_inicio_actividad=data.fiscal.irpf_inicio_actividad,
        nota_irpf=data.fiscal.nota_irpf,
        iban=data.bancari.iban,
        bic_swift=data.bancari.bic_swift,
        entitat_bancaria=data.bancari.entitat_bancaria,
        titular_compte=data.bancari.titular_compte,
        termini_pagament_dies=data.bancari.termini_pagament_dies,
        forma_pagament=data.bancari.forma_pagament,
        color_primario=data.branding.color_primario,
        color_secundario=data.branding.color_secundario,
        color_acento=data.branding.color_acento,
        color_texto=data.branding.color_texto,
        color_fondo_alt=data.branding.color_fondo_alt,
        color_blanco=data.branding.color_blanco,
        fuente_principal=data.branding.fuente_principal,
        fuente_fallback=data.branding.fuente_fallback,
        logo_path=logo_path,
        serie_factura=data.facturacio.serie_factura,
        moneda=data.facturacio.moneda,
        idioma=data.facturacio.idioma,
    )

"""Contexto de presentacion corporativa para los reportes."""

from dataclasses import dataclass

from gantt.bc3.models import BrandingConfig
from gantt.reporting.styles import DocumentPalette, build_palette


@dataclass(frozen=True)
class PresentationContext:
    """Contrato explicito de branding, idioma y tema visual de una ejecucion."""

    company: BrandingConfig | None
    palette: DocumentPalette
    language: str
    theme_id: str


def build_presentation_context(company: BrandingConfig | None = None) -> PresentationContext:
    """Construye el contexto visual completo con fallback neutro."""
    palette = build_palette(company)
    language = company.idioma if company else "es"
    company_name = company.nombre if company else "default"
    theme_id = (
        company_name.lower()
        .replace(" ", "-")
        .replace(".", "")
        .replace("/", "-")
    )
    return PresentationContext(
        company=company,
        palette=palette,
        language=language,
        theme_id=theme_id,
    )

"""Primitivas de rendering documental para Excel.

Este modulo contiene geometria, logo, cabecera, footer e impresion. No calcula
valores de dominio ni resuelve configuracion desde filesystem.
"""

from dataclasses import dataclass
from datetime import date

from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from gantt.bc3.models import BrandingConfig, ProjectConfig
from gantt.reporting.documents import (
    BillingMetadata,
    DocumentMetadata,
    DocumentType,
    document_type_info,
)
from gantt.reporting.presentation import PresentationContext
from gantt.reporting.styles import DocumentPalette, LITERALES, build_palette


THIN_GREY = Side(style="thin", color="D9D9D9")
BOX_BORDER = Border(
    left=THIN_GREY,
    right=THIN_GREY,
    top=THIN_GREY,
    bottom=THIN_GREY,
)


@dataclass(frozen=True)
class LogoRenderPolicy:
    """Politica unica de dimensionado de logo."""

    container_width: int = 170
    container_height: int = 56
    max_width: int = 150
    max_height: int = 44
    alignment: str = "center"
    clear_space: int = 8
    preserve_aspect_ratio: bool = True


@dataclass(frozen=True)
class RenderedLogoSize:
    width: int
    height: int
    original_width: int
    original_height: int


def _font(palette: DocumentPalette, *, size: int, bold: bool = False, color: str | None = None) -> Font:
    return Font(
        name=palette.font_family,
        size=size,
        bold=bold,
        color=(color or palette.color_texto).lstrip("#"),
    )


def calculate_logo_size(
    original_width: int,
    original_height: int,
    policy: LogoRenderPolicy | None = None,
) -> RenderedLogoSize:
    """Calcula el mayor tamano posible preservando aspecto."""
    policy = policy or LogoRenderPolicy()
    available_width = max(1, min(policy.container_width, policy.max_width) - policy.clear_space)
    available_height = max(1, min(policy.container_height, policy.max_height) - policy.clear_space)
    if original_width <= 0 or original_height <= 0:
        return RenderedLogoSize(0, 0, original_width, original_height)
    scale = min(available_width / original_width, available_height / original_height)
    return RenderedLogoSize(
        width=max(1, round(original_width * scale)),
        height=max(1, round(original_height * scale)),
        original_width=original_width,
        original_height=original_height,
    )


def add_logo(ws, company: BrandingConfig | None, policy: LogoRenderPolicy | None = None) -> RenderedLogoSize | None:
    """Inserta logo si existe y devuelve dimensiones renderizadas."""
    if not company or not company.logo_path:
        return None
    try:
        img = XLImage(str(company.logo_path))
        size = calculate_logo_size(img.width, img.height, policy)
        img.width = size.width
        img.height = size.height
        img.anchor = "A1"
        ws.add_image(img)
        return size
    except Exception:
        return None


def metadata_from_project_config(
    config: ProjectConfig,
    document_type: DocumentType,
    *,
    title: str = "",
) -> DocumentMetadata:
    """Adaptador legacy: crea DocumentMetadata sin consultar fuentes externas."""
    info = document_type_info(document_type)
    return DocumentMetadata(
        document_type=document_type,
        title=title or info.default_title,
        project_code=config.numero_expediente,
        project_name=config.proyecto,
        site_name=config.edificio,
        client_name=config.entidad,
        reference=config.numero_expediente,
        revision=config.revision,
        issue_date=date.today(),
    )


def render_budget_header(
    ws,
    metadata: DocumentMetadata,
    presentation: PresentationContext,
    num_columns: int,
) -> int:
    """Cabecera profesional para familia Budget/Analysis."""
    palette = presentation.palette
    company = presentation.company
    col_max = get_column_letter(num_columns)
    has_logo = add_logo(ws, company) is not None

    for row, height in [(1, 26), (2, 18), (3, 18), (4, 8), (5, 26), (6, 20), (7, 26), (8, 20), (9, 8)]:
        ws.row_dimensions[row].height = height

    if has_logo:
        ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 14)
        if num_columns >= 2:
            ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 12)
        start_col = 3 if num_columns >= 4 else 2
    else:
        start_col = 1

    ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=num_columns)
    ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=num_columns)
    ws.cell(1, start_col).value = company.nombre if company else ""
    ws.cell(1, start_col).font = _font(palette, size=15, bold=True, color=palette.color_primario)
    ws.cell(2, start_col).value = _company_claim(company)
    ws.cell(2, start_col).font = _font(palette, size=9, color=palette.color_texto)

    info = document_type_info(metadata.document_type)
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=num_columns)
    ws.cell(5, 1).value = (metadata.project_name or "").upper()
    ws.cell(5, 1).font = _font(palette, size=13, bold=True, color=palette.color_primario)
    ws.cell(5, 1).alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=num_columns)
    site_client = " - ".join(p for p in [metadata.site_name, metadata.client_name] if p)
    ws.cell(6, 1).value = site_client
    ws.cell(6, 1).font = _font(palette, size=10)

    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=num_columns)
    ws.cell(7, 1).value = f"{info.document_code} - {(metadata.title or info.default_title).upper()}"
    ws.cell(7, 1).fill = palette.fill_header
    ws.cell(7, 1).font = palette.font_header
    ws.cell(7, 1).alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(start_row=8, start_column=1, end_row=8, end_column=num_columns)
    date_text = metadata.issue_date.strftime("%d/%m/%Y")
    parts = [f"Ref. {metadata.reference or metadata.project_code}", f"Rev. {metadata.revision}", date_text]
    if metadata.planning and metadata.planning.scenario:
        parts.insert(0, f"Escenario: {metadata.planning.scenario}")
    ws.cell(8, 1).value = "   |   ".join(p for p in parts if p)
    ws.cell(8, 1).font = _font(palette, size=9)

    for row in range(1, 9):
        for col in range(1, num_columns + 1):
            ws.cell(row, col).border = BOX_BORDER
            if row in (1, 7):
                ws.cell(row, col).fill = palette.fill_header
            else:
                ws.cell(row, col).fill = palette.fill_white
            ws.cell(row, col).alignment = Alignment(vertical="center", wrap_text=True)
    return 10


def render_planning_header(
    ws,
    metadata: DocumentMetadata,
    presentation: PresentationContext,
    num_columns: int,
) -> int:
    """Cabecera especifica para planificacion."""
    metadata = metadata.model_copy(update={"title": metadata.title or "Planificacion del proyecto"})
    first_content = render_budget_header(ws, metadata, presentation, num_columns)
    ws.cell(7, 1).value = "PLANIFICACION DEL PROYECTO"
    if metadata.planning and metadata.planning.scenario:
        ws.cell(8, 1).value = (
            f"Escenario: {metadata.planning.scenario}   |   "
            f"Ref. {metadata.reference or metadata.project_code}   |   "
            f"Rev. {metadata.revision}   |   {metadata.issue_date.strftime('%d/%m/%Y')}"
        )
    return first_content


def render_invoice_header(
    ws,
    metadata: DocumentMetadata,
    presentation: PresentationContext,
    num_columns: int,
) -> int:
    """Cabecera especifica para facturas."""
    palette = presentation.palette
    company = presentation.company
    add_logo(ws, company)
    for row, height in [(1, 26), (2, 18), (3, 18), (4, 18), (5, 8)]:
        ws.row_dimensions[row].height = height
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=3)
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=3)
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=3)
    ws.cell(1, 2).value = company.nombre if company else ""
    ws.cell(1, 2).font = _font(palette, size=15, bold=True, color=palette.color_primario)
    ws.cell(2, 2).value = f"CIF/NIF: {company.nif_cif}" if company else ""
    ws.cell(3, 2).value = company.direccion_fiscal if company else ""
    ws.cell(2, 2).font = _font(palette, size=9)
    ws.cell(3, 2).font = _font(palette, size=9)

    billing = metadata.billing or BillingMetadata()
    ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=num_columns)
    ws.cell(1, 4).value = "FACTURA"
    ws.cell(1, 4).font = _font(palette, size=18, bold=True)
    ws.cell(1, 4).alignment = Alignment(horizontal="right")
    lines = [
        f"Numero: {billing.invoice_number or 'BORRADOR'}",
        f"Estado: {billing.invoice_status.upper()}",
        f"Fecha: {metadata.issue_date.strftime('%d/%m/%Y')}",
        f"Vence: {billing.due_date.strftime('%d/%m/%Y') if billing.due_date else ''}",
    ]
    for idx, text in enumerate(lines, start=2):
        ws.merge_cells(start_row=idx, start_column=4, end_row=idx, end_column=num_columns)
        ws.cell(idx, 4).value = text
        ws.cell(idx, 4).font = _font(palette, size=9, bold=idx == 2)
        ws.cell(idx, 4).alignment = Alignment(horizontal="right")
    for row in range(1, 5):
        for col in range(1, num_columns + 1):
            ws.cell(row, col).border = BOX_BORDER
            ws.cell(row, col).alignment = Alignment(vertical="center", wrap_text=True)
    return 6


def apply_document_footer(
    ws,
    metadata: DocumentMetadata,
    presentation: PresentationContext,
    num_columns: int,
    *,
    invoice: bool = False,
) -> None:
    """Footer visible y footer de impresion."""
    palette = presentation.palette
    company = presentation.company
    col_max = get_column_letter(num_columns)
    row = ws.max_row + 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_columns)
    if company:
        if invoice:
            text = " | ".join(p for p in [company.nombre, company.nif_cif, company.direccion_fiscal, company.web] if p)
        else:
            text = " | ".join(
                p for p in [
                    f"{company.nombre} - {company.web}" if company.web else company.nombre,
                    f"{metadata.reference or metadata.project_code} - Rev. {metadata.revision}",
                ] if p
            )
    else:
        text = f"{metadata.reference or metadata.project_code} - Rev. {metadata.revision}"
    c = ws.cell(row, 1)
    c.value = text
    c.font = palette.font_small
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = palette.fill_alt
    ws.row_dimensions[row].height = 16
    ws.oddFooter.left.text = (company.nombre if company else "")
    ws.oddFooter.center.text = f"{metadata.reference or metadata.project_code} - Rev. {metadata.revision}"
    ws.oddFooter.right.text = "Pagina &[Page] de &[Pages]"
    ws.print_area = f"$A$1:${col_max}${ws.max_row}"


def configure_excel_printing(
    ws,
    *,
    orientation: str = "portrait",
    repeat_row: int | None = None,
    fit_to_height: int = 0,
) -> None:
    ws.page_setup.orientation = orientation
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_to_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    ws.page_margins.header = 0.25
    ws.page_margins.footer = 0.25
    ws.sheet_view.showGridLines = False
    ws.sheet_view.view = "pageLayout"
    if repeat_row is not None:
        ws.print_title_rows = f"{repeat_row}:{repeat_row}"


def _company_claim(company: BrandingConfig | None) -> str:
    if not company:
        return ""
    parts = [company.web, company.email, company.telefono]
    contact = " | ".join(p for p in parts if p and p != "PENDIENTE")
    return contact or "Ingenieria - Energia - Sistemas"


def presentation_from_palette_company(
    palette: DocumentPalette | None,
    company: BrandingConfig | None,
) -> PresentationContext:
    pal = palette or build_palette(company)
    return PresentationContext(
        company=company,
        palette=pal,
        language=company.idioma if company else "es",
        theme_id=(company.nombre.lower().replace(" ", "-") if company else "default"),
    )

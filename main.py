"""
Punto de entrada principal del sistema gantt-tools.

Orquesta el pipeline completo: parseo BC3, generación de documentos de
presupuesto y, si existe planificacion.yaml, cálculo de planificación
y exportación Excel + diagrama de red.

Uso:
    python main.py <nombre-proyecto> [archivo.bc3]

Ejemplos:
    python main.py 4t-2t-Sotano
    python main.py Complexe-Balear
    python main.py Complexe-Balear pressupost_revisat.bc3
"""

import sys
from pathlib import Path

from gantt.application.context import (
    resolver_contexto_ejecucion,
    seleccionar_bc3,
    resolver_output_dir as resolver_output_dir_contexto,
)
from gantt.application.manifest import write_run_manifest
from gantt.application.documents import apply_document_identity_fallbacks
from gantt.bc3.parser import parse_bc3
from gantt.planning.analyser import duracion_total_proyecto
from gantt.planning.validation import validar_planificacion_previa
from gantt.reporting.presentation import build_presentation_context


def verificar_entorno() -> None:
    """Comprueba dependencias del entorno antes de procesar. Avisa si algo falta."""
    try:
        from matplotlib.font_manager import findfont, FontProperties
        from gantt.reporting.palette import FONT_PRES
    except ImportError:
        return

    if FONT_PRES and 'DejaVu' in findfont(FontProperties(family=FONT_PRES)):
        print(f'AVISO: fuente "{FONT_PRES}" no encontrada — los documentos de presupuesto'
              f' usarán la fuente por defecto del sistema. Instala {FONT_PRES} para'
              f' obtener el formato corporativo correcto.')
from gantt.planning.calculator import calcular_planificacion
from gantt.planning.models import cargar_planificacion_yaml
from gantt.reporting.analysis_charts import AnalysisChartsReport
from gantt.reporting.excel_exporter import exportar_analisis
from gantt.reporting.network_diagram import generar_diagrama_red
from gantt.reporting.presupuesto_exporter import generar_todos


def resolver_bc3_path(input_dir: Path, bc3_filename: str | None = None) -> tuple[Path, list[Path]]:
    """
    Localiza el BC3 a procesar y devuelve tambien todos los BC3 disponibles.
    """
    return seleccionar_bc3(input_dir, bc3_filename)


def resolver_output_dir(project_name: str, bc3_path: Path, bc3_files: list[Path]) -> Path:
    """
    Si un proyecto tiene varias versiones BC3, separa sus salidas por fichero.
    """
    return resolver_output_dir_contexto(Path('projects') / project_name, bc3_path, bc3_files)


def main(project_name: str, bc3_filename: str | None = None) -> None:
    """
    Pipeline completo para el proyecto indicado.
    El bloque de presupuesto y reporting se ejecuta siempre.
    El bloque de planificación solo si existe planificacion.yaml.
    """
    context = resolver_contexto_ejecucion(project_name, bc3_filename)
    input_dir = context.input_dir
    bc3_path = context.bc3_path
    output_dir = context.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    yaml_path = context.planificacion_path
    total = 4 if yaml_path is not None else 3
    output_path = output_dir / f'{project_name}_gantt.xlsx'
    outputs: list[Path] = []
    warnings: list[str] = []

    # ── BLOQUE 1: Presupuesto ──────────────────────────────────────────
    print(f'[1/{total}] Parseando BC3: {bc3_path.name}')
    presupuesto = parse_bc3(bc3_path)
    print(f'      {len(presupuesto.capitulos)} capítulos, '
          f'{len(presupuesto.recursos_mo)} recursos MO, '
          f'{len(presupuesto.recursos_mt)} recursos MT')

    if context.config_path:
        presupuesto = presupuesto.model_copy(
            update={'config': presupuesto.config.model_copy(update=context.project_config_update)}
        )
        print(f'      config.yaml cargado')
    presupuesto = apply_document_identity_fallbacks(presupuesto, context)

    if context.company_config:
        presupuesto.company = context.company_config
        print(f'      Empresa: {presupuesto.company.nombre} ({presupuesto.company.idioma})')
    else:
        print('      Sin empresa configurada — documentos sin branding')
        warnings.append('No hay empresa configurada para esta ejecución.')

    presentation = build_presentation_context(presupuesto.company)
    palette = presentation.palette

    print(f'[2/{total}] Generando documentos de presupuesto...')
    rutas = generar_todos(presupuesto, output_dir, palette=palette)
    outputs.extend(rutas)
    for ruta in rutas:
        print(f'      {ruta.name}')

    print(f'[3/{total}] Generando gráficos de reporting...')
    exportar_analisis(presupuesto, output_path, None, palette=palette)
    report = AnalysisChartsReport.from_presupuesto(
        presupuesto=presupuesto,
        output_dir=output_dir / 'reporting',
        palette=palette,
    )
    report.generate()
    outputs.append(output_path)
    outputs.extend(sorted((output_dir / 'reporting').rglob('*.png')))
    print(f'      {output_path.name}')
    print(f'      reporting/ ({len(list((output_dir / "reporting").rglob("*.png")))} gráficos)')

    # ── BLOQUE 2: Planificación (opcional) ────────────────────────────
    if yaml_path is None:
        warnings.append('No se encontró planificacion.yaml; se omitió planificación temporal.')
        write_run_manifest(context, outputs, [], warnings, presentation=presentation)
        print()
        print('No se encontró planificacion.yaml — omitiendo Gantt y diagrama de red.')
        print(f'Documentos generados en: {output_dir}')
        return

    print('[4/4] Calculando planificación y exportando...')
    parametros, escenarios, tareas, bandas = cargar_planificacion_yaml(yaml_path)
    warnings.extend(
        validar_planificacion_previa(presupuesto, parametros, escenarios, tareas)
    )

    planificaciones = []
    for escenario in escenarios:
        planificacion = calcular_planificacion(
            presupuesto, parametros, tareas, escenario, bandas
        )
        planificaciones.append(planificacion)
        dias_totales = duracion_total_proyecto(planificacion)
        cumple = (
            'CUMPLE'
            if dias_totales <= parametros.plazo_contractual_dias
            else 'NO CUMPLE'
        )
        print(
            f'      Escenario "{escenario.nombre}" calculado: '
            f'{dias_totales} dias habiles - {cumple}'
        )

    exportar_analisis(presupuesto, output_path, planificaciones, palette=palette)
    print(f'      {output_path.name} (actualizado con planificación)')

    ruta_png = output_dir / f'{project_name}_gantt.png'
    png_bytes = generar_diagrama_red(planificaciones[0], palette=palette)
    ruta_png.write_bytes(png_bytes)
    outputs.append(ruta_png)
    print(f'      {ruta_png.name}')

    write_run_manifest(
        context,
        outputs,
        [plan.escenario.nombre for plan in planificaciones],
        warnings,
        presentation=presentation,
    )

    print()
    print(f'Completado. Resultados en: {output_dir}')


if __name__ == '__main__':
    verificar_entorno()

    if len(sys.argv) < 2:
        print('Uso: python main.py <nombre-proyecto> [archivo.bc3]')
        print('Ejemplos:')
        print('  python main.py Complexe-Balear')
        print('  python main.py Complexe-Balear pressupost_revisat.bc3')
        sys.exit(1)

    project_name = sys.argv[1]
    bc3_filename = sys.argv[2] if len(sys.argv) > 2 else None
    main(project_name, bc3_filename)

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
import yaml
from pathlib import Path

from gantt.bc3.models import cargar_company
from gantt.bc3.parser import parse_bc3
from gantt.planning.analyser import duracion_total_proyecto
from gantt.reporting.styles import build_palette


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
    bc3_files = sorted(input_dir.glob('*.bc3'), key=lambda ruta: ruta.name.lower())

    if bc3_filename:
        bc3_path = input_dir / bc3_filename
        if not bc3_path.exists():
            raise FileNotFoundError(
                f'No se encontró el archivo BC3 indicado: {bc3_path}'
            )
        return bc3_path, bc3_files

    if not bc3_files:
        raise FileNotFoundError(f'No se encontró ningún .bc3 en {input_dir}')
    if len(bc3_files) > 1:
        raise ValueError(
            f'Múltiples .bc3 en {input_dir}. '
            f'Especifica el archivo: '
            f'python main.py <nombre-proyecto> <archivo.bc3>\n'
            f'Disponibles: {[f.name for f in bc3_files]}'
        )
    return bc3_files[0], bc3_files


def resolver_output_dir(project_name: str, bc3_path: Path, bc3_files: list[Path]) -> Path:
    """
    Si un proyecto tiene varias versiones BC3, separa sus salidas por fichero.
    """
    project_output_dir = Path('projects') / project_name / 'output'
    if len(bc3_files) > 1:
        return project_output_dir / bc3_path.stem
    return project_output_dir


def main(project_name: str, bc3_filename: str | None = None) -> None:
    """
    Pipeline completo para el proyecto indicado.
    El bloque de presupuesto y reporting se ejecuta siempre.
    El bloque de planificación solo si existe planificacion.yaml.
    """
    input_dir  = Path('projects') / project_name / 'input'
    bc3_path, bc3_files = resolver_bc3_path(input_dir, bc3_filename)
    output_dir = resolver_output_dir(project_name, bc3_path, bc3_files)
    output_dir.mkdir(parents=True, exist_ok=True)

    yaml_path = input_dir / 'planificacion.yaml'
    total = 4 if yaml_path.exists() else 3
    output_path = output_dir / f'{project_name}_gantt.xlsx'

    # ── BLOQUE 1: Presupuesto ──────────────────────────────────────────
    print(f'[1/{total}] Parseando BC3: {bc3_path.name}')
    presupuesto = parse_bc3(bc3_path)
    print(f'      {len(presupuesto.capitulos)} capítulos, '
          f'{len(presupuesto.recursos_mo)} recursos MO, '
          f'{len(presupuesto.recursos_mt)} recursos MT')

    config_path = input_dir / 'config.yaml'
    project_config: dict = {}
    if config_path.exists():
        project_config = yaml.safe_load(config_path.read_text(encoding='utf-8')) or {}
        presupuesto = presupuesto.model_copy(
            update={'config': presupuesto.config.model_copy(update=project_config)}
        )
        print(f'      config.yaml cargado')

    companies_dir = Path('companies')
    empresa_slug = project_config.get('empresa')
    if empresa_slug and companies_dir.exists():
        presupuesto.company = cargar_company(empresa_slug, companies_dir)
        print(f'      Empresa: {presupuesto.company.nombre} ({presupuesto.company.idioma})')
    else:
        print('      Sin empresa configurada — documentos sin branding')

    palette = build_palette(presupuesto.company)

    print(f'[2/{total}] Generando documentos de presupuesto...')
    rutas = generar_todos(presupuesto, output_dir, palette=palette)
    for ruta in rutas:
        print(f'      {ruta.name}')

    print(f'[3/{total}] Generando gráficos de reporting...')
    exportar_analisis(presupuesto, output_path, None)
    report = AnalysisChartsReport(
        excel_path=output_path,
        output_dir=output_dir / 'reporting',
    )
    report.generate()
    print(f'      {output_path.name}')
    print(f'      reporting/ ({len(list((output_dir / "reporting").rglob("*.png")))} gráficos)')

    # ── BLOQUE 2: Planificación (opcional) ────────────────────────────
    if not yaml_path.exists():
        print()
        print('No se encontró planificacion.yaml — omitiendo Gantt y diagrama de red.')
        print(f'Documentos generados en: {output_dir}')
        return

    print('[4/4] Calculando planificación y exportando...')
    parametros, escenarios, tareas, bandas = cargar_planificacion_yaml(yaml_path)

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

    exportar_analisis(presupuesto, output_path, planificaciones)
    print(f'      {output_path.name} (actualizado con planificación)')

    ruta_png = output_dir / f'{project_name}_gantt.png'
    png_bytes = generar_diagrama_red(planificaciones[0])
    ruta_png.write_bytes(png_bytes)
    print(f'      {ruta_png.name}')

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

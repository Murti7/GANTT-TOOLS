"""
Punto de entrada principal del sistema gantt-tools.

Orquesta el pipeline completo: parseo BC3, cálculo de planificación
y exportación Excel. Recibe el nombre del proyecto como argumento.

Uso:
    python main.py <nombre-proyecto> [archivo.bc3]

Ejemplos:
    python main.py Complexe-Balear
    python main.py Complexe-Balear pressupost_revisat.bc3
"""

import sys
from datetime import timedelta
from pathlib import Path

from gantt.bc3.parser import parse_bc3
from gantt.planning.calculator import calcular_planificacion
from gantt.planning.models import cargar_planificacion_yaml
from gantt.reporting.analysis_charts import AnalysisChartsReport
from gantt.reporting.excel_exporter import exportar_analisis
from gantt.reporting.presupuesto_exporter import generar_todos


def main(project_name: str, bc3_filename: str | None = None) -> None:
    """
    Ejecuta el pipeline completo para el proyecto indicado.
    Si bc3_filename se especifica, usa ese archivo BC3 concreto.
    Si no, busca el único .bc3 disponible en input/ (error si hay varios).
    """
    input_dir  = Path('projects') / project_name / 'input'
    output_dir = Path('projects') / project_name / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)

    # — Paso 1: localizar y parsear BC3
    if bc3_filename:
        bc3_path = input_dir / bc3_filename
        if not bc3_path.exists():
            raise FileNotFoundError(
                f'No se encontró el archivo BC3 indicado: {bc3_path}'
            )
    else:
        bc3_files = list(input_dir.glob('*.bc3'))
        if not bc3_files:
            raise FileNotFoundError(f'No se encontró ningún .bc3 en {input_dir}')
        if len(bc3_files) > 1:
            raise ValueError(
                f'Multiples .bc3 en {input_dir}. '
                f'Especifica el archivo: python main.py {project_name} <archivo.bc3>\n'
                f'Disponibles: {[f.name for f in bc3_files]}'
            )
        bc3_path = bc3_files[0]

    print(f'[1/5] Parseando BC3: {bc3_path.name}')
    presupuesto = parse_bc3(bc3_path)
    importe = f'{presupuesto.importe_total:,.2f}'.translate(str.maketrans(',.', '.,'))
    print(f'      {len(presupuesto.capitulos)} capítulos | '
          f'{len(presupuesto.recursos_mo)} recursos MO | '
          f'{importe} € PEM')

    # — Paso 2: cargar planificacion.yaml y calcular si existe
    planificaciones = None
    yaml_path = input_dir / 'planificacion.yaml'

    if yaml_path.exists():
        print('[2/5] Cargando planificacion.yaml y calculando escenarios...')
        parametros, escenarios, tareas = cargar_planificacion_yaml(yaml_path)
        planificaciones = [
            calcular_planificacion(presupuesto, parametros, tareas, escenario)
            for escenario in escenarios
        ]
        for plan in planificaciones:
            hito_fin = next((t for t in plan.tareas if t.es_fin_plazo), None)
            cumple = ''
            if hito_fin and hito_fin.fecha_fin:
                dias = sum(
                    1 for i in range((hito_fin.fecha_fin - parametros.fecha_inicio).days)
                    if (parametros.fecha_inicio + timedelta(days=i)).weekday()
                    < parametros.dias_semana
                )
                estado = 'CUMPLE' if dias <= parametros.plazo_contractual_dias else 'NO CUMPLE'
                cumple = (f'{dias}d habiles hasta {hito_fin.nombre} '
                          f'-> {estado} (plazo: {parametros.plazo_contractual_dias}d)')
            print(f'      Escenario [{plan.escenario.nombre}]: {cumple}')
    else:
        print('[2/5] planificacion.yaml no encontrado — solo análisis BC3')

    # — Paso 3: exportar Excel
    output_path = output_dir / f'{project_name}_gantt.xlsx'
    print(f'[3/5] Generando Excel: {output_path.name}')
    exportar_analisis(presupuesto, output_path, planificaciones)
    print(f'      Excel generado en {output_path}')

    # — Paso 4: generar gráficos de reporting
    print('[4/5] Generando gráficos de reporting...')
    report = AnalysisChartsReport(
        excel_path=output_path,
        output_dir=output_dir / 'reporting',
    )
    report.generate()
    print(f'      Gráficos generados en {output_dir / "reporting"}')

    # — Paso 5: generar documentos de presupuesto
    print('[5/5] Generando documentos de presupuesto...')
    rutas = generar_todos(presupuesto, output_dir)
    for ruta in rutas:
        print(f'      {ruta.name}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Uso: python main.py <nombre-proyecto> [archivo.bc3]')
        print('Ejemplos:')
        print('  python main.py Complexe-Balear')
        print('  python main.py Complexe-Balear pressupost_revisat.bc3')
        sys.exit(1)

    project_name = sys.argv[1]
    bc3_filename = sys.argv[2] if len(sys.argv) > 2 else None
    main(project_name, bc3_filename)

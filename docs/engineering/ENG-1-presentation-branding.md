# ENG-1 - Presentation & Branding Architecture

## Estado

Implementado un contrato explicito de presentacion para la ejecucion:

- `PresentationContext` en `gantt/reporting/presentation.py`.
- `DocumentPalette`, `ChartStyle` y `DiagramStyle` en `gantt/reporting/styles.py`.
- Validacion estricta de colores `#RRGGBB` al construir la paleta.
- `run_manifest.json` registra `presentation_theme` y `presentation_language`.

## Decisiones de arquitectura

La configuracion de empresa sigue viviendo en `companies/<slug>/company.yaml` y se carga como `BrandingConfig`. El pipeline construye una sola vez el `PresentationContext` despues de parsear el BC3 y resolver la empresa, y pasa su `palette` a:

- documentos de presupuesto;
- Excel de analisis;
- graficos PNG de reporting;
- diagrama temporal de planificacion.

Los graficos de reporting ya pueden renderizar desde `Presupuesto` mediante `BudgetAnalysisData`, sin depender del Excel como fuente intermedia. La lectura de Excel queda conservada como camino legacy en `AnalysisChartsReport`.

## Cambios funcionales

El pipeline principal usa:

```python
presentation = build_presentation_context(presupuesto.company)
palette = presentation.palette
```

Los PNG de reporting se generan con:

```python
AnalysisChartsReport.from_presupuesto(..., palette=palette).generate()
```

El diagrama temporal acepta ahora:

```python
generar_diagrama_red(planificacion, palette=palette)
```

`exportar_analisis` acepta `palette` y aplica los tokens al workbook. Internamente conserva un puente temporal para el exporter historico, que aun usa constantes globales.

## Compatibilidad

Se mantienen las llamadas existentes:

- `AnalysisChartsReport(excel_path, output_dir)` sigue leyendo Excel.
- `exportar_analisis(presupuesto, output_path, planificaciones=None)` sigue funcionando sin paleta explicita.
- `generar_diagrama_red(planificacion)` sigue funcionando con paleta por defecto.

## Validacion

Suite automatizada:

- `33 passed`.
- Tests nuevos cubren validacion de color, contexto de presentacion, graficos desde modelo, Excel con paleta y diagrama con paleta.

Validacion real:

- `python main.py "Viding Fitness Calvià" "BAF-VID-2026-001_V2.2.bc3"`
- Outputs BAFRAS en `projects/Viding Fitness Calvià/output/BAF-VID-2026-001_V2.2`.
- Outputs Murti de validacion en `projects/Viding Fitness Calvià/output/ENG1_murti_autonomo_visual`.

Nota: `bafras-engineering` y `murti-autonomo` tienen identidades fiscales distintas, pero actualmente declaran la misma paleta visual en `company.yaml`; por eso los colores resultantes coinciden.

## Deuda pendiente

El exporter de analisis Excel aun debe migrarse de forma completa a estilos inyectados por funcion. ENG-1 introduce un puente acotado para evitar una reescritura grande y preservar compatibilidad, pero el siguiente paso natural es eliminar las constantes `GANTT_*` del flujo de escritura.

Tambien queda pendiente enriquecer `company.yaml` si se quieren paletas visuales realmente diferenciadas entre BAFRAS y Murti.

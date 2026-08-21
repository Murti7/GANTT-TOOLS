# ENG-0 - Core consolidation

## Problema

`main.py` y algunos modulos auxiliares resolvian directamente rutas de proyecto,
BC3, config, empresa y output. Esto duplicaba politica de ejecucion y hacia
dificil auditar que entrada habia producido cada carpeta de salida.

Ademas, parte de la configuracion YAML circulaba como `dict` sin contrato fuerte.
En particular, `empresa` podia entrar mediante `model_copy(update=...)` aunque no
pertenecia a `ProjectConfig`.

## Decisiones

- Se introduce `gantt.application.context` como unica frontera de resolucion de
  proyecto, BC3, config, empresa, output y `run_id`.
- Se mantiene el comportamiento de salida historico para proyectos con un solo
  BC3 y se separa por `output/<bc3_stem>` cuando hay varias versiones BC3.
- `config.yaml` se valida con `ProjectInputConfig`. El campo `empresa` queda en
  el contexto de ejecucion; los campos documentales/financieros se validan y se
  aplican a `ProjectConfig`.
- `company.yaml` se valida con modelos Pydantic de bloques: identidad, fiscal,
  bancario, branding y facturacion. La salida sigue siendo `BrandingConfig` para
  no romper exporters.
- Se introduce `run_manifest.json` por ejecucion con rutas, hash SHA-256 del BC3,
  config, empresa, timestamp, version Python, planificacion, escenarios, outputs
  y warnings.
- Se introduce `gantt.planning.validation` para validar relaciones entre
  presupuesto, tareas y escenarios antes de calcular fechas.
- Los escenarios incompletos se registran como warning, no como error, para
  preservar el comportamiento historico: el calculo actual usa 1 operario por
  defecto cuando un recurso MO no aparece en el escenario.
- Se extrae `gantt.bc3.economics.calcular_resumen_financiero` para sacar la
  cascada financiera critica de los exporters y poder testearla sin Excel.

## Alternativas descartadas

- No se hizo un refactor completo de exporters. Son grandes, pero moverlos en
  bloque aumentaria el riesgo sin aportar una frontera clara inmediata.
- No se convirtio `analysis_charts.py` a modelo -> graficos en ENG-0. Sigue
  leyendo el Excel generado; queda planificado para una fase posterior.
- No se corrigio automaticamente mojibake de BC3 o documentacion. La estrategia
  elegida es hacer explicita la politica de encoding y evitar reparaciones
  heuristicas.

## Politica de encoding

- Codigo Python, YAML y JSON del proyecto: UTF-8.
- BC3: lectura actual en `latin-1`, compatible con los archivos existentes y con
  la realidad habitual FIEBDC-3. No se fuerza conversion a UTF-8.
- Si un BC3 ya contiene texto corrupto por mojibake, ENG-0 no intenta repararlo.
  Cualquier reparacion automatica debera basarse en reglas inequívocas y tests
  especificos.
- Los manifests se escriben como UTF-8 con `ensure_ascii=False`.

## Archivos modificados

- `main.py`
- `gantt/bc3/models.py`
- `gantt/bc3/economics.py`
- `gantt/application/__init__.py`
- `gantt/application/context.py`
- `gantt/application/manifest.py`
- `gantt/planning/validation.py`
- `gantt/reporting/presupuesto_exporter.py`
- `tests/test_application_context_manifest.py`
- `tests/test_bc3_economics.py`
- `tests/test_pipeline_end_to_end.py`
- `tests/test_planning_validation.py`

## Impacto esperado

- Ejecuciones mas trazables mediante `run_manifest.json`.
- Menos duplicacion en resolucion proyecto/BC3/output.
- Configuracion YAML validada y con errores legibles.
- Planificaciones invalidas fallan antes del forward pass.
- Calculos financieros criticos testeables fuera del layout Excel.

## Compatibilidad

- `python main.py <proyecto> [archivo.bc3]` sigue siendo la interfaz principal.
- Los wrappers `resolver_bc3_path` y `resolver_output_dir` siguen disponibles en
  `main.py` para compatibilidad de tests/usos existentes.
- El baseline de Complexe-Balear se mantiene:
  - `PPT_Complex_V09_Un_Ref_Sin_Act_Vaso.bc3`: base 108, optimizado 51.
  - `PPT_Complex_VFinal.bc3`: base 93, optimizado 46.

## Deuda pendiente

- `analysis_charts.py` aun depende de releer el Excel generado.
- `DocumentPalette` no es todavia la unica fuente visual de todos los outputs.
- Persisten literales y comentarios con mojibake en documentacion y codigo.
- Los CLIs auxiliares de modulos internos no se han consolidado.
- No hay `pyproject.toml`, lint, type-checking ni CI.
- Los outputs historicos dentro de `projects/*/output` no se migraron.

## Fronteras futuras

- ENG-1 Branding: unificar `DocumentPalette` en Excel de analisis, graficos y
  diagrama temporal; eliminar hardcodes visuales gradualmente.
- ENG-2 Facturacion: separar fiscalidad/facturacion real de branding y de
  documentos de presupuesto; completar modelos de empresa y cliente.

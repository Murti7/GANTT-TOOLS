# ENG-3B - Document Rendering & Layout Hardening

## Baseline

Suite inicial antes de cambios:

```text
65 passed
```

Golden documents regenerados:

- Budgeting Viding Fitness Calvia, `BAF-VID-2026-001_VC`.
- Planning Complexe-Balear, `PPT_Complex_VFinal`.
- Billing draft BAFRAS y draft Murti desde el presupuesto Viding.

Invariantes de planificacion preservados:

```text
PPT_Complex_V09_Un_Ref_Sin_Act_Vaso: base=108, optimizado=51
PPT_Complex_VFinal: base=93, optimizado=46
```

## ENG-3B.3 follow-up: PRES.03 split

La normalizacion inicial de `PRES.03_Mediciones.xlsx` se migra a dos documentos
canonicos independientes:

- `PRES.03.01_Mediciones.xlsx`: recupera el documento historico de mediciones,
  basado en el generador compartido de descompuesto con `mostrar_precios=False`.
  Mantiene capitulos, subcapitulos, partidas, descripciones, unidades,
  cantidades, detalle de recursos/descompuesto sin precios visibles y filas
  marcador `TOTAL CAPITULO` no economicas.
- `PRES.03.02_Mediciones_Ciegas.xlsx`: mantiene la variante ciega introducida en
  ENG-3B.3, construida directamente desde capitulos/partidas y sin precios,
  importes, recursos, descompuestos, totales economicos ni economia oculta.

`PRES.03_Mediciones.xlsx` queda como nombre intermedio obsoleto; no se genera en
la salida canonica.

## Auditoria visual inicial

Incidencias observadas antes de implementar:

- Budget: doble cabecera, una corporativa en filas 1-4 y otra documental legacy en filas 6-10. Clasificacion: `LAYOUT`, `CONSISTENCY`, `LEGACY`.
- Budget: logo insertado sin politica documental explicita y dependiente de dimensiones fijas. Clasificacion: `LOGO`.
- Budget: footer visible separado de la configuracion de impresion y print area calculada antes del footer. Clasificacion: `PRINTING`, `CONSISTENCY`.
- Analysis: sin cabecera documental, sin logo, sin footer y sin print area. Clasificacion: `METADATA`, `LAYOUT`, `PRINTING`.
- Planning analysis: dashboard sin cabecera documental normalizada y sin footer/print area; el diagrama se insertaba correctamente, pero el libro no parecia documento Planning. Clasificacion: `LAYOUT`, `CONSISTENCY`, `PRINTING`.
- Invoice: layout propio funcional, pero sin logo, sin footer, sin print area ni metadata documental explicita. Clasificacion: `BRANDING`, `LOGO`, `PRINTING`, `METADATA`.
- Configuracion legacy de Viding no define `client`; para golden Billing se uso cliente explicito derivado de la entidad/proyecto. Clasificacion: `LEGACY`.

Limitacion de validacion: no se genero PDF ni se abrio Excel en GUI desde el entorno. La validacion real se hizo mediante inspeccion estructural de workbooks con openpyxl y revision de PNG generados.

## Decisiones

- Se introduce `gantt.reporting.rendering` como capa pequena de render: cabecera, footer, logo, printing y adaptadores legacy.
- No se crea framework de plantillas.
- No se modifican calculos economicos ni planificacion.
- `DocumentMetadata` y `PresentationContext` pasan a ser la fuente para Analysis, Planning e Invoice.
- Budget mantiene firmas legacy, pero su cabecera se adapta internamente a `DocumentMetadata` para evitar un refactor masivo de `presupuesto_exporter.py`.
- El logo usa `LogoRenderPolicy` y `calculate_logo_size`, preservando aspect ratio y limites maximos.

## Layouts finales

Budget:

- Emisor y logo.
- Proyecto.
- Site/cliente.
- Codigo y titulo documental.
- Referencia, revision y fecha.
- Tabla funcional debajo de la cabecera.
- Footer visible e impresion normalizada.

Planning:

- Header propio con "PLANIFICACION DEL PROYECTO".
- Proyecto, site, escenario, referencia, revision y fecha.
- Dashboard y hojas tecnicas bajo la cabecera.
- Footer y print area por worksheet.

Invoice:

- Header propio con logo/emisor a la izquierda y bloque FACTURA a la derecha.
- Numero/estado/fecha/vencimiento desde metadata.
- Secciones EMISOR, CLIENTE, CONCEPTOS, RESUMEN, FORMA DE PAGO.
- Footer fiscal y configuracion de impresion.

## Exporters migrados

- `presupuesto_exporter.py`: Budget usa cabecera/footer/printing comunes mediante adaptador legacy.
- `excel_exporter.py`: acepta `metadata` y `presentation`; aplica layout Budget o Planning segun uso.
- `invoice_exporter.py`: acepta `metadata`; renderiza factura desde `Invoice + InvoiceCalculation + DocumentMetadata + PresentationContext`.
- `application/budgeting.py` y `application/planning.py`: propagan metadata ya resuelta tras aplicar fallbacks de identidad documental.
- `application/billing.py`: acepta metadata opcional para nuevas rutas de aplicacion.

## Validacion post-cambio

Suite final:

```text
69 passed
```

Golden documents revisados estructuralmente:

- `PRES.01_Cuadro_Oferta.xlsx`
- `PRES.02.01_Cuadro_Precios_1.xlsx`
- `PRES.02.03_Presupuesto_Descompuesto.xlsx`
- `PRES.02.04_Resumen_Capitulos.xlsx`
- `budget_analysis.xlsx`
- `planning_analysis.xlsx`
- factura draft BAFRAS
- factura draft Murti

Resultado:

- Budget ya no muestra doble cabecera.
- `Importado desde ficheros CSV` no gobierna el titulo principal.
- Analysis y Planning tienen cabecera, footer, print area, orientacion y freeze panes.
- Invoice tiene logo, footer, print area y bloque documental de factura.
- BAFRAS mantiene retencion cero y Murti mantiene retencion profesional.

## Tests anadidos

- Logo scaling preservando aspect ratio.
- Header Budget sin titulo importado y con printing.
- Invoice con metadata, footer, printing y fallback sin logo.
- Filename documental de invoice draft.

## Trabajo pendiente

- Validacion visual con render PDF real cuando haya LibreOffice/Excel disponible.
- Hacer que todos los documentos Budget reciban `DocumentMetadata` por firma publica explicita, eliminando el adaptador legacy.
- Propagar `DocumentMetadata` tambien a titulos/subtitulos de graficos PNG.
- Mejorar `project.yaml` en proyectos legacy para evitar cliente derivado manualmente en Billing.

## ENG-3B.1 - Visual QA & Delivery Layout Fix

### Baseline

Suite inicial:

```text
69 passed
```

Invariantes de planificacion preservados:

```text
PPT_Complex_V09_Un_Ref_Sin_Act_Vaso: base=108, optimizado=51
PPT_Complex_VFinal: base=93, optimizado=46
```

### Problemas iniciales

- Budget DELIVERY tenia demasiada altura de cabecera, duplicaba cliente/site y congelaba paneles como si fuera una hoja de trabajo.
- Analysis reutilizaba una cabecera demasiado orientada a entrega y no distinguia claramente su uso operativo.
- Invoice parecia un formulario interno: repetia `BORRADOR/DRAFT`, separaba emisor y cliente en bloques verticales, mostraba muchos `PENDIENTE` y congelaba paneles.
- El logo estaba tecnicamente insertado, pero hacia falta validar su geometria en XML porque OpenPyXL recarga el tamano original del asset.

### DELIVERY vs ANALYSIS

Se introduce `DocumentPurpose`:

- `DELIVERY`: documentos para cliente/impresion. Sin freeze panes por defecto, cabecera contractual compacta, footer y print area.
- `ANALYSIS`: documentos para trabajo en pantalla. Cabecera compacta, freeze panes semantico, filtros y navegacion.

### Budget final

Budget DELIVERY queda con:

- zona superior de identidad corporativa;
- bloque de cliente/proyecto sin duplicacion absurda cuando site y cliente coinciden;
- titulo documental con codigo;
- control documental compacto con referencia, revision y fecha;
- `freeze_panes=None`;
- `print_title_rows` apuntando al header de tabla, no a toda la cabecera corporativa.

### Invoice final

Factura redisenada como documento DELIVERY:

- `FACTURA`, `BORRADOR` y fecha en bloque superior derecho;
- EMISOR y CLIENTE en dos columnas;
- datos pendientes representados como `-`;
- bloque PROYECTO compacto;
- resumen fiscal con `Base imponible`, `IVA`, `Retencion` cuando aplica, `TOTAL FACTURA` y `A PAGAR` cuando hay retencion;
- forma de pago mostrando solo datos existentes;
- `freeze_panes=None`;
- manifest con `ready_to_issue`, `completeness_status` y `missing_required_fields`.

### Planning y Gantt

Planning Analysis usa cabecera ANALYSIS compacta:

```text
BAFRAS Engineering S.L. - PLANIFICACION
<project>
Escenario - Ref - Fecha
```

El PNG de Gantt se mantiene sin cambios de calculo. QA de asset: `5256 x 3059 RGBA`.

### Logo

`LogoRenderPolicy` ahora explicita:

- `max_width`;
- `max_height`;
- `minimum_visual_presence`;
- `clear_space`;
- preservacion de aspect ratio.

QA XML confirma escalado del logo a `cx=1609725`, `cy=381000`, aproximadamente 169 x 40 px.

### Freeze policy

| Purpose | Policy |
| ------- | ------ |
| DELIVERY | `freeze_panes=None` |
| ANALYSIS | `freeze_panes` en primera fila navegable tras cabecera compacta |

### Printing y footer

- DELIVERY: print-first, orientacion por documento, fit-to-width, area de impresion y rows-to-repeat para headers de tabla.
- ANALYSIS: landscape cuando procede, filtros, freeze y print area.
- Budget footer: emisor + referencia/revision + pagina.
- Invoice footer: emisor fiscal + referencia/revision + pagina.

### Tests

Suite final:

```text
70 passed
```

Cobertura anadida/actualizada:

- `DocumentPurpose` DELIVERY vs ANALYSIS;
- Budget DELIVERY sin freeze panes;
- Invoice sin redundancia `BORRADOR/DRAFT`;
- Invoice sin freeze panes;
- manifest de completeness documental;
- logo scaling.

### Golden outputs

Budget oficial Viding no pudo regenerarse completo porque
`PRES.02.03_Presupuesto_Descompuesto.xlsx` estaba bloqueado por el sistema
(`PermissionError: [Errno 13] Permission denied`), probablemente abierto en
Excel. Para QA se genero un set completo alternativo:

```text
projects/Viding Fitness Calvia/output/budgeting/BAF-VID-2026-001_VC_ENG3B1_QA
```

Planning e invoices si se regeneraron en sus rutas normales.

### QA matrix

| Documento | Diseno | Metadata | Logo | Freeze | Printing | Visual QA |
| --- | --- | --- | --- | --- | --- | --- |
| PRES.01 | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| PRES.02.01 | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| PRES.02.03 | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| PRES.02.04 | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| Budget Analysis | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| Planning Analysis | PASS | PASS | N/A | PASS | PASS | STRUCTURAL |
| Gantt PNG | PASS | PASS | N/A | N/A | N/A | ASSET |
| Invoice BAFRAS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| Invoice Murti | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |

### Limitaciones

No hay `soffice`, `libreoffice` ni `excel` disponibles en PATH, por lo que no
se pudo generar PDF/screenshot real headless. La QA visual se hizo con:

- inspeccion estructural OpenPyXL;
- inspeccion de XML interno de drawings para geometria de logo;
- verificacion de dimensiones de PNG con Pillow;
- golden outputs generados para revision manual.

### No modificado

No se modifico parser BC3, PEM, GG, BI, PEC, IVA, retenciones,
`InvoiceCalculation`, escenarios ni duraciones.

## ENG-3B.2 - Excel Integrity & Final Visual Polish

### Baseline

Suite inicial:

```text
70 passed
```

Suite final:

```text
72 passed
```

Baselines de planificacion preservados:

```text
PPT_Complex_V09_Un_Ref_Sin_Act_Vaso: base=108, optimizado=51
PPT_Complex_VFinal: base=93, optimizado=46
```

### P0 - Integridad XLSX

Se introdujo `Excel Integrity Gate` en `gantt.reporting.xlsx_integrity`.

Valida:

- ZIP valido;
- XML parseable;
- relaciones internas existentes;
- defined names sin `#REF!`;
- roundtrip `load_workbook -> save -> load_workbook`.

La causa raiz corregida en codigo fue el uso de `ws.auto_filter.ref =
ws.dimensions` en libros ANALYSIS. Ese rango incluia cabeceras fusionadas y
footer, una combinacion aceptada por OpenPyXL pero propensa a reparacion por
Microsoft Excel. La solucion calcula la primera fila filtrable real y limita el
autofiltro al rango de tabla:

```text
A<table_header_row>:<last_column><last_data_row>
```

Esto preserva filtros y evita incluir merges documentales.

### Archivos reparados por Excel

No fue posible ejecutar Microsoft Excel desde el entorno. `soffice`,
`libreoffice` y `excel` no estaban disponibles en PATH en la fase anterior y la
validacion automatica se limito a OOXML/OpenPyXL. Por tanto, la confirmacion
final `NO REPAIR DIALOG` queda como QA manual obligatoria antes de release.

### Golden integrity

El gate estructural se ejecuto sobre 24 XLSX:

- Budget QA completo;
- Budget Analysis;
- Planning Analysis;
- F01-F06 BAFRAS;
- F01 Murti.

Resultado:

```text
checked 24 failed 0
```

Ademas:

```text
checked 24 problems []
```

para `####`, A4, print area y fit-to-width.

### Facturas F01-F06

Se centraliza la identidad documental en `gantt.billing.presets`:

| Codigo | Titulo documental |
| --- | --- |
| F01 | Factura de honorarios profesionales |
| F02 | Factura sobre presupuesto de ejecucion material |
| F03 | Factura sobre presupuesto con beneficio industrial |
| F04 | Factura sobre presupuesto completo |
| F05 | Factura de certificacion parcial |
| F06 | Factura de liquidacion final |

`BillingPreset` sigue siendo clasificacion interna. El titulo visible sale de
`document_title` y el exporter no lo hardcodea.

### Metadata economica de factura

`Invoice` incorpora:

- `billing_preset`;
- `billing_source_type`;
- `billing_source_reference`;
- `economic_basis`.

`create_invoice` rellena `source_type` y `economic_basis` desde el preset cuando
no se pasan explicitamente. Esto no modifica fiscalidad ni calculos.

### Filename de factura

Politica:

```text
F03_DRAFT_<invoice-id>.xlsx
<invoice-number>_F03.xlsx
```

Los outputs QA usan esta politica en `output/billing/ENG3B2_QA`.

### Invoice polish

- Cabecera superior muestra el titulo especifico de factura, `BORRADOR` o numero, fecha y vencimiento si aplica.
- Campos opcionales vacios se omiten en bloques emisor/cliente.
- `PENDIENTE` no aparece visualmente en facturas QA.
- Footer de invoice omite datos pendientes y usa separador `·`.
- Formato monetario unificado: `#.##0,00 [$€-es-ES]`.
- Anchos mantienen importes razonables sin `####`.

### Budget polish

- Cabeceras estrechas recuperan reference/revision/date sin sobrescribir el claim.
- DELIVERY mantiene `freeze_panes=None`.
- ANALYSIS usa autofiltro semantico y freeze en primera fila navegable.

### A4 policy

`configure_excel_printing` fija A4 en todos los XLSX renderizados por la capa
documental:

- Invoice: A4 portrait, fit-to-width 1.
- PRES estrechos: A4 portrait.
- PRES anchos: A4 landscape cuando ya lo definia el exporter.
- Analysis/Planning: A4 landscape cuando el ancho lo requiere.

### Golden outputs ENG-3B.2

Budget:

```text
projects/Viding Fitness Calvia/output/budgeting/BAF-VID-2026-001_VC_ENG3B2_QA
```

Planning:

```text
projects/Complexe-Balear/output/planning/PPT_Complex_VFinal
```

Billing:

```text
projects/Viding Fitness Calvia/output/billing/ENG3B2_QA
```

### Acceptance matrix

| Documento | Excel Integrity | Title | Metadata | Logo | Layout | Monetary Format | Printing | Visual QA |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| PRES.01 | PASS | PASS | PASS | PASS | PASS | N/A | PASS | STRUCTURAL |
| PRES.02.01 | PASS | PASS | PASS | PASS | PASS | N/A | PASS | STRUCTURAL |
| PRES.02.03 | PASS | PASS | PASS | PASS | PASS | N/A | PASS | STRUCTURAL |
| PRES.02.04 | PASS | PASS | PASS | PASS | PASS | N/A | PASS | STRUCTURAL |
| Budget Analysis | PASS | PASS | PASS | PASS | PASS | N/A | PASS | STRUCTURAL |
| Planning Analysis | PASS | PASS | PASS | N/A | PASS | N/A | PASS | STRUCTURAL |
| F01 | PASS | PASS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| F02 | PASS | PASS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| F03 | PASS | PASS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| F04 | PASS | PASS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| F05 | PASS | PASS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| F06 | PASS | PASS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |
| Murti invoice | PASS | PASS | PASS | PASS | PASS | PASS | PASS | STRUCTURAL |

### QA Microsoft Excel

Pendiente de validacion manual fuera del entorno:

```text
Abrir cada golden XLSX en Microsoft Excel y confirmar NO REPAIR DIALOG.
```

No declarar release documental final si cualquier archivo muestra reparacion.

### No modificado

No se modifico parser BC3, cantidades, precios, PEM, GG, BI, PEC, IVA, IRPF,
`InvoiceCalculation`, `TaxPolicy`, escenarios ni duraciones.

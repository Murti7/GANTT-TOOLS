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

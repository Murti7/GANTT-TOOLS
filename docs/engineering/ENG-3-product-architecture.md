# ENG-3 - Product Architecture & Document System

## Auditoria inicial

Baseline antes de cambios:

```text
48 passed
PPT_Complex_V09_Un_Ref_Sin_Act_Vaso: base 108, optimizado 51
PPT_Complex_VFinal: base 93, optimizado 46
```

Problemas observados:

- `main.py` seguia representando conceptualmente todo el producto.
- Las salidas legacy mezclaban documentos entregables y analisis.
- No existia catalogo central de tipos documentales.
- No existia metadata documental explicita.
- `config.yaml` era la unica configuracion de proyecto, aunque ya no cubria
  bien Billing/Client/Product.
- Los documentos ENG funcionaban como historial, pero faltaba documentacion
  actual de producto.

## Decisiones

Se implementa una evolucion incremental, no un movimiento masivo de archivos.

Nuevas fronteras:

- `gantt.application.product`: capabilities.
- `gantt.application.project`: `ProjectModel`, `project.yaml`, legacy adapter.
- `gantt.application.outputs`: politica de filesystem.
- `gantt.application.budgeting`: caso de uso Budgeting.
- `gantt.application.planning`: caso de uso Planning.
- `gantt.application.billing`: Billing artifacts y salida draft/issued.
- `gantt.reporting.documents`: `DocumentType`, `DocumentMetadata`, titulos y filenames.
- `gantt.reporting.layouts`: familias Budget/Planning/Invoice.
- `gantt.billing.presets`: F01-F06 como presets.

## Project model

`project.yaml` es la fuente canonica futura. Si no existe, se adapta
`input/config.yaml` a `ProjectModel` y se registra warning trazable.

`company.yaml` sigue describiendo emisores reutilizables.

`planning.yaml` sigue limitado a planificacion.

## Output architecture

Nuevas APIs por capability escriben:

```text
output/budgeting/<source-id>/documents
output/budgeting/<source-id>/analysis
output/budgeting/<source-id>/analysis/charts
output/planning/<source-id>
output/billing/drafts/<invoice-id>
output/billing/issued/<invoice-number>
```

`main.py` mantiene la salida legacy para compatibilidad.

## Document system

`DocumentType` define identidad documental, no calculos.

`DocumentMetadata` separa:

- proyecto;
- cliente;
- documento;
- revision;
- fecha;
- source;
- metadata especifica Budget/Planning/Billing.

La politica de titulo prioriza:

```text
explicit title -> configured title -> metadata title -> default title
-> source description -> filename fallback
```

Asi, una descripcion BC3 como `Importado desde ficheros CSV` no gobierna el
titulo editorial del documento.

## Layouts

Se formalizan tres familias:

- Budget;
- Planning;
- Invoice.

Comparten `PresentationContext`, pero la metadata documental es independiente
del branding.

## F01-F06

F01-F06 se formalizan como presets/origenes economicos:

- F01: Professional Fees.
- F02: PEM.
- F03: PEM + BI.
- F04: PEM + GG + BI / PEC.
- F05: Partial Certification.
- F06: Final Settlement.

No se crean seis motores de factura.

## Documentation architecture

Se crea documentacion actual:

```text
docs/architecture
docs/capabilities
docs/configuration
docs/examples
```

`docs/engineering` queda como historial de decisiones.

## Compatibilidad

No se mueve ningun proyecto historico.

No se modifica `main.py` como CLI legacy.

No se cambian formulas de BC3, planificacion ni Billing.

## Deuda pendiente

- Migrar progresivamente exporters de presupuesto para recibir
  `DocumentMetadata` en sus firmas publicas.
- Aplicar fisicamente layouts ENG-3 en todas las cabeceras Excel de Budget y
  Planning; hoy queda formalizado y testeado como capa, con integracion
  completa en nuevos flujos de salida.
- Migrar `gantt.bc3.economics` de `float` a `Decimal` en una fase economica
  posterior.
- Implementar CLI definitivo por subcomandos.
- Implementar validacion visual automatizada mas rica si se anade soporte PDF
  o screenshots.

## No implementado deliberadamente

- Facturae/FACE.
- PDF.
- Numeracion fiscal persistente.
- Emision irreversible.
- Base de datos.
- ERP/CRM.
- Refactor completo de exporters.
- Migracion masiva de outputs historicos.

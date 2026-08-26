# Document System

ENG-3 introduces:

- `DocumentType`: central catalog of document identities.
- `DocumentMetadata`: source of document metadata.
- title resolution policy: explicit title, configured title, metadata title,
  document default, source description, filename.
- filename policy: stable, short and traceable.
- layout families: Budget, Planning and Invoice.

Source BC3 descriptions are traceability data. They do not govern editorial
document titles.

Presentation remains handled by `PresentationContext`; document identity is a
separate concern.

## Rendering Contract

ENG-3B makes the document system visible in the generated artifacts.

Excel exporters should receive document identity and presentation as resolved
inputs:

```text
domain data + DocumentMetadata + PresentationContext -> rendered document
```

Exporters must not resolve project, client, company, revision, source or
language from filesystem paths. Legacy Budget exporters still expose their old
function signatures, but their headers and footers are adapted internally to
`DocumentMetadata` through `gantt.reporting.rendering`.

## Rendering Primitives

`gantt.reporting.rendering` owns:

- logo sizing via `LogoRenderPolicy` and `calculate_logo_size`;
- Budget header rendering;
- Planning header rendering;
- Invoice header rendering;
- visible document footer;
- Excel print setup: margins, orientation, fit-to-page, print area and repeat
  rows.

Budget, Planning and Invoice use different layout families. They share tokens
from `PresentationContext`, but they do not share the same geometry.

## Current Exporter State

- `presupuesto_exporter.py`: uses the common Budget rendering layer through a
  compatibility adapter.
- `excel_exporter.py`: accepts optional `metadata` and `presentation`, and
  renders Budget Analysis or Planning Analysis accordingly.
- `invoice_exporter.py`: accepts optional `metadata` and renders from
  `Invoice + InvoiceCalculation + DocumentMetadata + PresentationContext`.

Source descriptions such as `Importado desde ficheros CSV` are traceability
data only and must not become the editorial project title or document title.

## Document Purpose

Every document type has a functional purpose:

- `DELIVERY`: print-first/client-delivery documents such as PRES documents,
  invoices, future certifications and settlements. These use compact
  contractual headers, footers, print areas and no freeze panes by default.
- `ANALYSIS`: screen-first workbooks such as budget analysis, planning analysis
  and resource analysis. These use compact operational headers, freeze panes,
  filters and table navigation.

`document_purpose(document_type)` is the public policy helper.

Budget measurements are modeled as two distinct document identities:

- `DocumentType.BLIND_MEASUREMENTS`: `PRES.03.01`, filename
  `PRES.03.01_Mediciones_Ciegas.xlsx`.
- `DocumentType.MEASUREMENTS`: `PRES.03.02`, filename
  `PRES.03.02_Mediciones.xlsx`.

Both are Budget `DELIVERY` documents generated independently from the
`Presupuesto` domain model. `PRES.03.01` is the commercially safe blind variant;
it must not be derived from another workbook by deleting or hiding columns.

Freeze panes are owned by the layout policy. Budget `DELIVERY` documents do not
configure freeze panes in their exporters. Analysis workbooks may configure
freeze panes for screen navigation.

## Excel Integrity Gate

`gantt.reporting.xlsx_integrity.validate_xlsx_integrity` validates generated
XLSX packages before client delivery:

- ZIP structure;
- XML parseability;
- internal relationship targets;
- defined names without obvious broken references;
- worksheet `sheetView` coherence;
- OpenPyXL load/save/load roundtrip.

The historical Excel repair issue came from this sequence:

```text
freeze pane -> DELIVERY print setup removes freeze pane -> orphan selection pane
```

OpenPyXL removed `<pane>` but kept a `<selection pane="bottomLeft">` in
`sheet1.xml`. Microsoft Excel reported this as a repaired worksheet view. The
gate now rejects any worksheet where a selection references a pane that does not
exist.

For Microsoft Excel release QA, manual confirmation is still required when Excel
is not available in the execution environment:

```text
NO REPAIR DIALOG
```

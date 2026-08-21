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

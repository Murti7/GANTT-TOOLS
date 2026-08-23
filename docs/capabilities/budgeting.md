# Budgeting

Input:

- BC3 file.
- Optional project configuration.
- Optional company issuer for branding.

Output in ENG-3 capability API:

```text
output/budgeting/<source-id>/
  documents/
  analysis/
    budget_analysis.xlsx
    charts/
  run_manifest.json
```

The budget domain uses `gantt.bc3.parser` and `gantt.bc3.economics`.
Charts render from domain/reporting models, not from Excel as the source of
truth.

## Documents

Budgeting generates the official PRES/JUST documents from the parsed BC3 domain
model:

- `PRES.02.03_Presupuesto_Descompuesto.xlsx`: decomposed budget with
  measurements and economic information.
- `PRES.03.01_Mediciones.xlsx`: measurements document that preserves the
  historical Budgeting representation. It includes chapters, items, technical
  descriptions, units, quantities, non-priced resource/decomposition detail and
  non-economic chapter total marker rows.
- `PRES.03.02_Mediciones_Ciegas.xlsx`: blind measurements for external exchange.
  It contains scope, technical descriptions, units and quantities only.

`PRES.03.02` is built directly from chapters and items. It must not be produced
by copying another XLSX, hiding columns, or writing prices/imports into hidden
cells, formulas, sheets, defined names or XML.

| Content | PRES.03.01 | PRES.03.02 |
| --- | --- | --- |
| Chapters | Yes | Yes |
| Items | Yes | Yes |
| Short description | Yes | Yes |
| Long description | Yes | Yes |
| Unit | Yes | Yes |
| Quantity | Yes | Yes |
| Decomposition | Non-priced resources | No |
| Price | No | No |
| Amount | No | No |
| Hidden economy | - | No |

Migration note: `PRES.03_Mediciones.xlsx` was an intermediate ENG-3B.3 filename.
The canonical outputs are now `PRES.03.01_Mediciones.xlsx` and
`PRES.03.02_Mediciones_Ciegas.xlsx`.

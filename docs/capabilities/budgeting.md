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

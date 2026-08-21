# Architecture Overview

`gantt-tools` is organized around three product capabilities:

- Budgeting: BC3 to contractual budget documents and analysis.
- Planning: BC3 plus `planning.yaml` to schedule analysis and diagrams.
- Billing: fees, contract, BC3, certification or settlement sources to invoices.

The intended dependency direction is:

```text
Domain -> Application -> Reporting / Presentation -> Files
```

`main.py` remains a legacy orchestration entry point. New product flows live in
`gantt.application.budgeting`, `gantt.application.planning` and
`gantt.application.billing`.

# Execution Flow

Budgeting:

```text
BC3 -> Presupuesto -> budget documents
                 -> budget_analysis.xlsx
                 -> charts
```

Planning:

```text
BC3 + planning.yaml -> PlanificacionProyecto -> planning_analysis.xlsx
                                         -> gantt.png
```

Billing:

```text
Billing source -> InvoiceLine -> Invoice -> InvoiceCalculation
                                      -> invoice.json
                                      -> invoice.xlsx
                                      -> invoice_manifest.json
```

The legacy CLI still runs the historical combined flow:

```bash
python main.py <project> [bc3]
```

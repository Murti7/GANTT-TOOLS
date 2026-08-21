# gantt-tools

`gantt-tools` is a technical/economic document system for engineering and
construction projects.

It currently supports three capabilities:

- **Budgeting**: BC3 files to budget documents, price tables, decomposed
  budgets, summaries, resource justification, analysis workbooks and charts.
- **Planning**: BC3 plus `planning.yaml` to schedule analysis, resource load,
  scenarios, critical path and temporal diagrams.
- **Billing**: project fees, BC3-derived bases, milestones, certifications or
  settlements to invoice domain objects, JSON snapshots, Excel and manifests.

The architecture separates:

```text
Domain -> Application -> Reporting / Presentation -> Output files
```

## Quick Start

Install dependencies in a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Run the legacy combined pipeline:

```bash
python main.py <project-name> [file.bc3]
```

Example:

```bash
python main.py "Viding Fitness Calvià" "BAF-VID-2026-001_VC.bc3"
```

The legacy CLI remains supported. New application APIs are organized by
capability in:

- `gantt.application.budgeting`
- `gantt.application.planning`
- `gantt.application.billing`

## Project Structure

Legacy projects are still supported:

```text
projects/<project>/
  input/
    config.yaml
    *.bc3
    planificacion.yaml
  output/
```

The canonical future structure is:

```text
projects/<project>/
  project.yaml
  input/
    budget/
    planning/
    billing/
  output/
    budgeting/
    planning/
    billing/
```

If `project.yaml` exists, it is preferred. If it does not exist,
`input/config.yaml` is adapted as legacy configuration and a warning is recorded.

## Documentation

- Architecture: `docs/architecture/`
- Capabilities: `docs/capabilities/`
- Configuration contracts: `docs/configuration/`
- Examples: `docs/examples/`
- Engineering history: `docs/engineering/`

Start with:

- `docs/architecture/overview.md`
- `docs/architecture/document-system.md`
- `docs/configuration/project-yaml.md`

## Tests

Run the full suite:

```bash
.venv\Scripts\python.exe -m pytest
```

## Scope Notes

ENG-3 does not implement a new CLI, PDF generation, Facturae/FACE, irreversible
invoice numbering, database persistence or an ERP workflow. Those remain future
extensions.

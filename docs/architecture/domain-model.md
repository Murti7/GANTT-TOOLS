# Domain Model

Core domains:

- `gantt.bc3`: BC3 parsing and budget economics.
- `gantt.planning`: tasks, scenarios, durations, CPM analysis and resource load.
- `gantt.billing`: invoice, client, issuer, tax policy, payment terms and JSON snapshots.

Shared application models:

- `ProjectExecutionContext`: resolved execution context.
- `ProjectModel`: future canonical project configuration.
- `DocumentMetadata`: document identity.
- `PresentationContext`: company branding and palette.

Billing can consume BC3 economics through application services, but `Invoice`
does not depend on the BC3 parser or budget exporters.

# Billing

Input:

- Manual fees.
- Contract milestones.
- BC3-derived economic bases.
- Future certification or final settlement sources.

Output:

```text
output/billing/drafts/<invoice-id>/
  invoice.xlsx
  invoice.json
  invoice_manifest.json

output/billing/issued/<invoice-number>/
  invoice.xlsx
  invoice.json
  invoice_manifest.json
```

ENG-3 keeps issued invoices conceptual. It does not implement irreversible
numbering or fiscal emission persistence.

F01-F06 are presets that describe the billing source/base, not six separate
invoice engines.

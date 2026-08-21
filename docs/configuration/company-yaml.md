# company.yaml

Company configuration describes reusable issuers:

- legal identity;
- fiscal settings;
- banking;
- billing defaults;
- branding.

It is not project-specific. A project selects an issuer by slug, for example:

```yaml
issuer: bafras-engineering
```

Billing maps company data to `Issuer`. Presentation maps company data to
`PresentationContext`.

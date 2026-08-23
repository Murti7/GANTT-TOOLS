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

## Branding Logo Asset Type

`branding.logo` points to the raster/vector asset. `branding.logo_asset_type`
describes what the asset contains:

- `symbol`: symbol only. The renderer may add company name and claim.
- `wordmark`: textual wordmark. The renderer must not duplicate the issuer name
  inside the same brand area.
- `lockup_horizontal`: symbol plus wordmark in a horizontal lock-up.
- `lockup_vertical`: symbol plus wordmark in a vertical lock-up.

When omitted, the legacy fallback is `symbol` for compatibility with existing
company files.

Example:

```yaml
branding:
  logo: logo.png
  logo_asset_type: lockup_horizontal
```

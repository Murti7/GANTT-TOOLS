"""Politicas fiscales configurables para facturacion."""

from gantt.billing.models import Client, ClientType, Issuer, TaxPolicy, TaxRate, WithholdingPolicy


def default_tax_policy(issuer: Issuer, client: Client) -> TaxPolicy:
    """Construye una politica fiscal por defecto desde emisor y cliente.

    La regla matematica de retencion vive en el calculador; esta funcion solo
    decide si la politica activa la retencion segun configuracion actual.
    """
    withholding_enabled = (
        issuer.withholding_enabled
        and issuer.default_withholding_rate > 0
        and client.client_type != ClientType.INDIVIDUAL
    )
    return TaxPolicy(
        vat=TaxRate(rate=issuer.default_vat_rate),
        withholding=WithholdingPolicy(
            enabled=withholding_enabled,
            rate=issuer.default_withholding_rate if withholding_enabled else 0,
            reason="Retencion profesional configurable" if withholding_enabled else "",
        ),
    )

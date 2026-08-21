# ENG-2 - Billing & Invoicing Domain

## Objetivo

ENG-2 introduce un dominio de facturacion independiente de Excel. La factura
existe como objeto de dominio, se calcula con `Decimal`, se persiste como JSON y
solo despues se representa como Excel.

Arquitectura resultante:

```text
Dominio Billing
    -> Application Billing
        -> Reporting / Presentation
            -> Excel / JSON / Manifest
```

`main.py` no genera facturas automaticamente. La facturacion queda como
capacidad separada.

## Entidades

- `Issuer`: emisor legal, construido desde `BrandingConfig` y `company.yaml`.
- `Client`: cliente contractual, con soporte para empresa, profesional,
  particular, administracion publica y extranjero.
- `InvoiceLine`: linea facturable con cantidad, precio, descuento y origen.
- `SourceReference`: trazabilidad a manual, BC3, hito, certificacion, anticipo,
  liquidacion o abono.
- `Invoice`: factura de dominio con estado, fechas, emisor, cliente, lineas,
  fiscalidad y pago.
- `Certification`: valoracion tecnica/economica separada de factura.
- `FinalSettlement`: liquidacion conceptual separada de factura.

## Fiscalidad

La fiscalidad se divide en:

- `TaxPolicy`: decide que aplicar a una factura concreta.
- `calculate_invoice`: calcula matematicamente importes.

El IVA se modela con `TaxRate` y `TaxTreatment`:

- `standard`;
- `reduced`;
- `exempt`;
- `not_subject`.

La retencion se modela con `WithholdingPolicy`. No es descuento comercial.

Formula:

```text
base imponible
+ IVA
= total factura
- retencion
= liquido a percibir
```

El porcentaje de IVA/retencion es configurable desde `company.yaml` o desde la
politica fiscal de la factura. No se codifican reglas legales cambiantes como
invariantes universales.

## Redondeo

Todos los calculos monetarios finales usan `Decimal`.

Politica:

```text
cantidad x precio
-> importe bruto de linea redondeado a 0.01
-> descuento de linea redondeado a 0.01
-> base imponible de linea redondeada a 0.01
-> suma de bases
-> IVA / retencion redondeados a 0.01
-> total factura / liquido redondeados a 0.01
```

Modo: `ROUND_HALF_UP`.

Python es la fuente oficial de verdad; Excel representa valores ya calculados.

## Estados y numeracion

Estados implementados:

- `draft`;
- `issued`;
- `partially_paid`;
- `paid`;
- `overdue`;
- `cancelled`;
- `credited`.

ENG-2 valida especialmente:

- `draft`: puede carecer de numero definitivo.
- `issued`: requiere numero, fecha de emision, identificacion fiscal y domicilio
  fiscal del emisor; cliente no particular requiere identificacion fiscal.

No se implementa contador irreversible ni persistencia de series. Las series
quedan preparadas mediante `Issuer.invoice_series`.

## Condiciones de pago

`PaymentTerms` representa:

- forma de pago;
- plazo;
- vencimiento;
- IBAN;
- BIC;
- titular;
- observaciones;
- referencia de pago.

Se construye por defecto desde el emisor con `payment_terms_from_issuer`.

## PEM / GG / BI / PEC

La facturacion desde BC3 no depende del parser dentro de `Invoice`.

Flujo:

```text
Presupuesto
    -> gantt.bc3.economics.calcular_resumen_financiero
        -> invoice_line_from_presupuesto
            -> InvoiceLine
```

F02-F04 se representan mediante `BillingBaseKind`:

- `PEM`;
- `PEM_BI`;
- `PEC`.

Las formulas no se duplican en billing.

## Certificaciones y liquidaciones

ENG-2 prepara los conceptos:

- `Certification`: contrato, acumulado ejecutado, certificado anterior,
  certificacion actual y porcentaje ejecutado.
- `FinalSettlement`: contrato original, modificaciones, facturado, ajustes y
  saldo final.

No se implementa gestion mensual completa, Facturae, FACE ni contabilidad.

## Persistencia y trazabilidad

Se genera `invoice.json` como snapshot reproducible:

- factura;
- calculo;
- hash SHA-256;
- timestamp;
- version Python;
- version `gantt-tools` si esta instalada.

La API de aplicacion `export_invoice_artifacts` genera:

- `invoice.xlsx`;
- `invoice.json`;
- `invoice_manifest.json`.

El manifest de factura registra datos no sensibles:

- `invoice_id`;
- `invoice_number`;
- estado;
- emisor;
- cliente;
- proyecto;
- source;
- hash;
- outputs.

## Excel

`gantt/reporting/invoice_exporter.py` recibe:

```python
Invoice
InvoiceCalculation
PresentationContext
output_path
```

No calcula impuestos. Usa la identidad visual de ENG-1.

Estructura:

- cabecera empresa / FACTURA;
- numero, fecha, proyecto y referencia;
- emisor;
- cliente;
- conceptos;
- resumen;
- forma de pago;
- notas.

## Configuracion de cliente

`ProjectConfig` acepta opcionalmente:

```yaml
client:
  legal_name:
  tax_id:
  vat_id:
  address:
  country:
  type:
  billing_email:
  contact_person:
  language:
  contract_reference:
  purchase_order:
  expediente:
```

Los proyectos existentes siguen funcionando sin `client`.

## Ejemplos

Factura sociedad:

```text
base 1000.00
IVA 21% 210.00
retencion 0.00
total factura 1210.00
liquido 1210.00
```

Factura autonomo:

```text
base 1000.00
IVA 21% 210.00
retencion 15% 150.00
total factura 1210.00
liquido 1060.00
```

Certificacion conceptual:

```text
contrato 100000.00
ejecutado acumulado 40000.00
certificado anterior 15000.00
certificacion actual 25000.00
```

Liquidacion conceptual:

```text
contrato original
+ modificaciones
+ ajustes
- importes ya facturados
= saldo final
```

## Validacion

Suite final tras ENG-2:

```text
48 passed
```

Baseline de planificacion preservada:

```text
PPT_Complex_V09_Un_Ref_Sin_Act_Vaso.bc3: base 108, optimizado 51
PPT_Complex_VFinal.bc3: base 93, optimizado 46
```

Validacion real DRAFT sobre Viding:

```text
projects/Viding Fitness Calvià/output/ENG2_billing_drafts/bafras-engineering
projects/Viding Fitness Calvià/output/ENG2_billing_drafts/murti-autonomo
```

Resultados:

```text
BAFRAS: base 11090.81, IVA 2329.07, retencion 0.00, total 13419.88, liquido 13419.88
Murti:  base 11090.81, IVA 2329.07, retencion 1663.62, total 13419.88, liquido 11756.26
```

## Limitaciones

- No hay contador persistente de facturas.
- No hay PDF fiscal definitivo.
- No hay Facturae/FACE.
- No hay base de datos ni workflow de cobros.
- No hay fiscalidad internacional completa.
- El exporter Excel es profesional pero inicial; PDF y plantillas avanzadas
  quedan para fases posteriores.
- `gantt.bc3.economics` sigue usando `float`; ENG-2 convierte a `Decimal` en la
  frontera Billing, pero una futura fase deberia migrar economia BC3 a Decimal.

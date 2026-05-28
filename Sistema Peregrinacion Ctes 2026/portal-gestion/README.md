# Sistema Peregrinacion CNC Corrientes 2026

Archivos nuevos creados sin modificar la plantilla original.

## Archivos

- `API.gs`: backend para Google Apps Script.
- `index.html`: portal web de participantes y panel organizador.

## Google Sheet esperado

El backend lee las columnas del formulario por encabezado. Esta version esta pensada para el formulario actual:

- Marca temporal
- Nombre/s
- Apellido/s
- Fecha de Nacimiento
- DNI
- Email que vas a usar para el portal de pagos
- Celular de contacto
- Telefono de contacto por emergencia
- Parroquia
- Comunidad
- Restricciones alimentarias
- Email automatico de Google Forms, si aparece

El sistema agrega al final, automaticamente, estas columnas internas:

- Aviso WhatsApp
- Fecha aviso WhatsApp
- Estado cupo
- Orden inscripcion
- Nota baja
- Reemplazo
- Vencimiento DNI

Tambien crea la hoja `Cambios Perfil` para guardar solicitudes de modificacion de datos hasta que un catequista las apruebe o rechace.

## Configurar `API.gs`

Reemplazar en `CONFIG`:

- `EMAIL_ORGANIZADOR`
- `ALIAS_TRANSFERENCIA`
- `CAPACIDAD`
- `ADMIN_PASSWORD`
- `ADMIN_VIEW_PASSWORD`
- `PORTAL_URL`

Cuando definamos el plan real, actualizar `CUOTAS`.

Plan actual cargado:

- 1ra cuota: $35.000, vence 30/06/2026
- 2da cuota: $35.000, vence 31/07/2026
- 3ra cuota: $35.000, vence 31/08/2026
- 4ta cuota: $35.000, vence 30/09/2026

## Configurar `index.html`

Reemplazar:

- `API`
- `ALIAS`
- `CVU`
- `TITULAR`
- `CLD_CLOUD`
- `CLD_PRESET`

Cuando definamos el plan real, actualizar `PLAN_PAGOS` con los mismos `key`, `label`, `monto` y `vencimiento` que `CUOTAS` en `API.gs`.

El portal muestra dos medios de pago:

- Transferencia por Mercado Pago: se debe subir comprobante con nombre, destino y numero de comprobante.
- Efectivo: se debe subir foto del recibo entregado por el responsable, indicando quien recibio el dinero. El pago queda en verificacion hasta que ese responsable entregue el dinero a los organizadores.

Reglas de pago:

- No se puede cargar una cuota posterior sin haber cargado antes el comprobante de la cuota anterior.
- Se puede declarar un monto distinto al valor de la cuota; ese monto queda asociado al comprobante y se usa para los totales reales.
- Si el comprobante se carga por efectivo, el peregrino debe indicar nombre y apellido del responsable.

## Flujo

1. El participante entra con DNI y email del portal de pagos.
2. Ve cuotas, importes, estado y alias de Mercado Pago.
3. Sube comprobantes de transferencia o foto del recibo de efectivo a Cloudinary.
4. Apps Script registra cada comprobante en la hoja `Pagos`.
5. El organizador revisa desde el panel admin.
6. Al confirmar, el sistema envia email de confirmacion al participante.

## Perfil del peregrino

El peregrino puede solicitar cambios en sus datos de contacto desde `Mis datos`. No puede cambiar DNI ni email de acceso desde el portal. Los cambios quedan pendientes y se aprueban desde `Acceso catequistas > Cambios de datos`.

## Pendiente antes de publicar

- Definir plan de pagos real.
- Configurar Cloudinary.
- Desplegar Apps Script como aplicacion web.
- Publicar `index.html` en GitHub Pages u otro hosting.
- Probar con una inscripcion de prueba.

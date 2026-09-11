# Viaje apostolico del Papa León XIV a Argentina - Inscripcion

Formulario para el viaje de los hermanos de las comunidades neocatecumenales de Corrientes, Chaco y Formosa.

URL prevista:

```text
https://maria365.online/visita-papa-leon/inscripcion-2026/
```

Google Sheet:

```text
https://docs.google.com/spreadsheets/d/12b2T7yNNkBGIbUe2XmWG4oQ1coLeKaXMHzikOsIKYPs/edit
```

## Estructura

- `index.html`: formulario publico de inscripcion.
- `API.gs`: backend para Google Apps Script.
- `README.md`: notas de despliegue.

## Datos solicitados

- Nombre
- Apellido
- DNI
- Fecha de nacimiento
- Email, necesario para ingresar al portal del peregrino
- Celular en formato internacional para WhatsApp
- Parroquia
- Comunidad
- Restricciones alimenticias
- Otra restriccion, solo si se elige `Otros`

La hoja guarda ademas `ID`, `Timestamp`, `Estado` y `Origen` como campos operativos. El backend no permite registrar dos veces el mismo DNI.

## Listas iniciales

Las parroquias se tomaron de la planilla publica del portal de peregrinacion anterior y se normalizaron para evitar duplicados de escritura. Las comunidades se normalizaron como `1ra` a `6ta` y `Otra`.

Restricciones alimenticias:

- Ninguna
- Celiaco/a
- Intolerancia a la lactosa
- Alergia a frutos secos
- Alergia al huevo
- Diabetes
- Otros

## Despliegue de Apps Script

Crear o vincular un proyecto de Apps Script para esta inscripcion, subir `API.gs` y desplegarlo como Web App.

Configuracion recomendada:

- Ejecutar como: propietario del script.
- Quien tiene acceso: cualquier usuario.

La URL `/exec` del Apps Script desplegado ya esta configurada en `index.html`.

Prueba minima:

```text
https://script.google.com/macros/s/ID_DE_DEPLOY/exec?action=health
```

Debe devolver:

```json
{"ok":true,"service":"visita-papa-leon-inscripcion-2026"}
```

URL actual del backend:

```text
https://script.google.com/macros/s/AKfycbwTawcad2Uc_MeEv_NM580i1ZlmTgogeJS-SmfuIvSw1LmDVhdF4i_P7G91Kc7NFPR0sQ/exec
```

## Pendiente

- Publicar los archivos del portal en GitHub Pages.

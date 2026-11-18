# App Comparativo Compras (FO 08-04-04)

Web app en Google Apps Script para el Comparativo Comercial FO 08-04-04 y el Informe de Adjudicación (PR-34).

- Proyecto de Apps Script: [`13xw3RNLSanWfNI-X6HO6wGeuG0LEq-2zFMWdjfbXdKzyjyPbk2ZlnSEV`](https://script.google.com/home/projects/13xw3RNLSanWfNI-X6HO6wGeuG0LEq-2zFMWdjfbXdKzyjyPbk2ZlnSEV/edit)
- URL de desarrollo (último código subido): `https://script.google.com/a/macros/ventus.global/s/AKfycbwpqptNfYMPfMKALCZWrlT6CO0PeG6OjsV3Qt76Y14o/dev`

## Estructura
| Ruta | Qué es |
|---|---|
| `src/Index.html` | **Fuente** de la app: HTML, CSS y JavaScript legibles. |
| `src/Manual.html` | **Fuente** del manual de uso (se abre con `?manual=1`). |
| `build.js` | Genera `appscript/*.html` desde `src/*.html`. |
| `appscript/` | Carpeta enlazada con CLASP (`.clasp.json`): lo que se sube a Apps Script. |
| `appscript/Código.js` | `doGet()`: sirve `Index` o `Manual` con HtmlService; `urlManual()`. |
| `appscript/Repositorio.js` | Repositorio en línea: procesos en Hojas de cálculo, perfiles y roles, flujo CMO, versiones, PDF en Drive y correos. |
| `appscript/Proveedores.js` | Lectura (con caché) de la base "BD Proveedores – Rev 1" para el compliance. |
| `appscript/appsscript.json` | Manifiesto: V8, `executeAs: USER_DEPLOYING`, `access: DOMAIN`. |

> **No edites `appscript/Index.html` ni `appscript/Manual.html`**: son archivos generados y el script va codificado. Los `.js` de `appscript/` sí se editan directamente.

## Flujo de trabajo
1. Traer lo último de Apps Script (por si alguien editó en el editor web):
   ```bash
   cd appscript && clasp pull
   ```
   Si cambiaron los `.html`, hay que actualizar `src/` antes de seguir. `node build.js --check` indica si `appscript/` y `src/` dejaron de coincidir.
2. Editar `src/*.html` o `appscript/*.js`.
3. Generar y subir:
   ```bash
   node build.js
   ```
   ```bash
   cd appscript && clasp push
   ```
4. Probar en la URL `/dev`.
5. Publicar: crear una versión nueva de la implementación (Implementar → Gestionar implementaciones, o `clasp deploy -i <deploymentId>`).
6. Hacer commit en git de `src/`, `appscript/` y `build.js`.

## Por qué existe `build.js`
`HtmlService.createHtmlOutputFromFile` procesa el JavaScript en línea y quita lo que cree que son comentarios (`//…` y `/* … */`). Ese procesador no entiende las *template literals* ni las expresiones regulares, así que corta cadenas como `'image/*'`, `'http://…'` o un base64. El script queda truncado y la app no carga: se ve el encabezado, pero no las pestañas.

`build.js` publica el script de cada página codificado en **base64url**, que no contiene `/`, `*` ni comillas. Un cargador de una línea lo decodifica y lo inserta como `<script>` normal, con el mismo alcance global, así que los `onclick="…"` siguen funcionando. Además, revisa la sintaxis antes de generar los archivos.

## Requisitos
- Node.js 18 o superior.
- CLASP 3 con una sesión iniciada en una cuenta con acceso de editor al proyecto (`clasp login`, `clasp show-authorized-user`).

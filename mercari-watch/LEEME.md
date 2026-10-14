# Vigilante de Mercari para tu Google Sheets

Cada 2 horas, GitHub (en su nube, sin tu ordenador) abre tus búsquedas de Mercari.
Si hay anuncios nuevos, la celda del enlace se pone **amarilla** y recibe una **nota**
(pasa el ratón por encima) con los títulos y enlaces de los nuevos.
La primera ejecución solo "memoriza" lo que ya hay (no avisa de nada).

Para quitar un aviso: selecciona la celda → clic derecho → *Borrar nota*, y quita el color.
No volverá a avisarte de esos mismos anuncios.

## Instalación (una sola vez, ~15 min)

### 1) Cuenta de servicio de Google (gratis)
1. Entra en https://console.cloud.google.com y crea un proyecto (ej. "mercari").
2. Menú **APIs y servicios → Biblioteca** → busca **Google Sheets API** → *Habilitar*.
3. **APIs y servicios → Credenciales → Crear credenciales → Cuenta de servicio**. Ponle un nombre y créala.
4. Entra en la cuenta creada → pestaña **Claves → Agregar clave → Crear clave nueva → JSON**. Se descarga un archivo `.json`.
5. Abre ese JSON y copia el valor de `client_email` (algo@proyecto.iam.gserviceaccount.com).
6. En tu Google Sheets pulsa **Compartir** y añade ese email como **Editor**.

> Tu archivo debe ser una hoja de Google nativa (no un .xlsx subido a Drive). Si lo es, el título no termina en ".xlsx".

### 2) GitHub
1. Crea cuenta en https://github.com y un repositorio nuevo **público** (los minutos de Actions son ilimitados en repos públicos).
2. Sube todos los archivos de esta carpeta (incluida la carpeta oculta `.github`). Lo más fácil: *Add file → Upload files* arrastrando la carpeta entera.
3. En el repo: **Settings → Secrets and variables → Actions → New repository secret**, crea dos:
   - `SPREADSHEET_ID`: el trozo de la URL de tu hoja entre `/d/` y `/edit`.
   - `GOOGLE_CREDENTIALS`: pega **todo el contenido** del archivo JSON.
   (Los secretos no son visibles para nadie, aunque el repo sea público.)

### 3) Probar
1. Pestaña **Actions → Vigilar Mercari → Run workflow**.
2. Abre la ejecución y mira el log: verás por cada enlace algo como `[5/125] LV X r5c4: 48 anuncios`.
   - Si salen anuncios → funciona. Esta primera vez es la línea base.
   - Si sale `FALLO` en casi todos → Mercari está bloqueando a GitHub (ver abajo).
3. Se creará una pestaña oculta `_mercari_estado` (la memoria del script). No la borres.
4. Para ver un aviso de prueba: espera a la siguiente ejecución o borra de esa pestaña la fila de un enlace
   y, después de una ejecución, publica/espera un anuncio nuevo.

## Ajustes
- **Frecuencia**: en `.github/workflows/mercari.yml`, línea `cron` (`0 */2 * * *` = cada 2 h; `0 */1 * * *` = cada hora).
- **Solo algunas pestañas**: descomenta la línea `SHEETS` en el mismo archivo, ej. `"LV X|Prime"`.

## Avisos importantes
- Si Mercari bloquea a GitHub, la ejecución falla y GitHub te manda un email de "workflow failed". No se pierde nada.
- GitHub desactiva los workflows programados si el repo lleva 60 días sin ninguna actividad. Te avisará por email;
  basta con pulsar *Enable workflow* (o hacer cualquier cambio pequeño en el repo).
- Se fuerza `solo en venta` + `más recientes` en las búsquedas (tu enlace de エレキブルFB no lo tenía).
- Los logs de un repo público son visibles: solo muestran palabras de búsqueda de cartas, nada privado.

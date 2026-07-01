# Auditoría del Sistema de Reservas

Revisión completa del proyecto (backend `.gs` y frontend `.html`) buscando errores, problemas de seguridad, optimizaciones de velocidad y código zombie. Al final se incluye la propuesta de **sistema de actualización para las copias** del programa.

Fecha de la revisión: 2026-07-01 · Archivos revisados: `Codigo.gs`, `AdminFunctions.gs`, `ReservasRecurrentes.gs`, `Incidencias.gs`, `Setup.gs`, `appsscript.json` y los 9 archivos HTML.

---

## 1. Errores (bugs)

### 1.1 🔴 Nombres de hoja con mayúsculas inconsistentes (`Config` vs `CONFIG`, `Incidencias` vs `INCIDENCIAS`)

`getSheetByName()` distingue mayúsculas/minúsculas. El instalador (`Setup.gs`, `DB_SCHEMA`) crea las hojas **`Config`** e **`Incidencias`**, pero varios sitios buscan otros nombres:

| Dónde | Busca | Debería buscar |
|---|---|---|
| `Codigo.gs:25` | `SHEETS.INCIDENCIAS = 'INCIDENCIAS'` | `'Incidencias'` |
| `AdminFunctions.gs:160` (`getAdminData`) | `getSheetByName('CONFIG')` | `'Config'` |
| `AdminFunctions.gs:1595` (`saveBatchConfig`) | `getSheetByName('CONFIG')` → **si no la encuentra crea una hoja `CONFIG` duplicada** | `'Config'` |
| `Incidencias.gs:329` (`enviarEmailNuevaIncidencia`) | `getSheetByName('CONFIG')` | `'Config'` |

Consecuencias en una instalación limpia: `reportarIncidencia()` falla con hoja nula (usa `SHEETS.INCIDENCIAS` = `'INCIDENCIAS'`), el panel admin no carga la configuración, y al guardar configuración se puede acabar con dos hojas de config desincronizadas.

**Arreglo:** usar siempre las constantes `SHEETS.CONFIG` / `SHEETS.INCIDENCIAS` con el valor exacto del `DB_SCHEMA` (`'Config'`, `'Incidencias'`) y eliminar todos los literales.

### 1.2 🔴 `handleAdminApproval` purga una clave de caché obsoleta

`Codigo.gs:1678`: purga `CACHE_KEYS.STATIC_DATA` (`'STATIC_DATA_V5'`), pero la caché activa es `CACHE_KEY_STATIC` (`'STATIC_DATA_V6_FULL'`). Al aprobar un usuario nuevo, el `usuariosMap` cacheado sigue viejo hasta 6 horas (el nombre del nuevo usuario no aparece en las reservas de otros). Cambiar por `cache.remove(CACHE_KEY_STATIC)`.

### 1.3 🟠 Zona horaria mal aplicada en `validarAntelacionMinima` (`Codigo.gs:1867`)

Construye la hora del tramo con `setUTCHours(horas, minutos)` cuando `hora_inicio` está en hora local de Madrid. En verano (UTC+2) el tramo de las 09:00 se trata como las 11:00 locales: la validación de antelación mínima es ~2h más permisiva de lo configurado. Construir la fecha con `Utilities.parseDate(fechaISO + ' ' + tramo.hora_inicio, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')` o calcular el offset con la zona del spreadsheet.

### 1.4 🟠 `getDatosMatrizUnificada`: mapa de usuarios con columnas intercambiadas

`AdminFunctions.gs:1777-1782`: el mapa se indexa por `u[0]` (columna A = **Nombre**) pero luego se consulta por email (`usuariosMap.get(emailUsuario)`), y además guarda `nombre: u[1]` (que es el email). El enriquecimiento nombre/área no funciona nunca. Debe ser: clave `u[1]` (Email), `nombre: u[0]`. El mismo error existe en `getSolicitudesPendientesGlobal` (`AdminFunctions.gs:1875-1883`), aunque esa función es código zombie (ver §4).

### 1.5 🟠 Celda del "modo visualización" de Cursos: D1 vs F1

- `Setup.gs:82` escribe el modo en **F1** (`sheet.getRange('F1').setValue('botones')`).
- `getStaticData` (`Codigo.gs:438`), `getAdminData` (`AdminFunctions.gs:150`) y `saveAllCursos` (`AdminFunctions.gs:539`) leen/escriben **D1**.

Además `crearDatosEjemploInterno` (`Setup.gs:113-116`) escribe 5 columnas de ejemplo en una hoja cuyo esquema declara solo 3 (`Etapa`, `Curso`, `Mostrar el curso con:`). Unificar en una celda (D1) y ajustar los datos de ejemplo a 3 columnas.

### 1.6 🟡 `saveBatchUsuarios` desalinea la columna Especialidad

`AdminFunctions.gs:1403`: limpia y reescribe solo las columnas A–D. Si el guardado reordena o elimina usuarios, la columna E (`Especialidad`) queda pegada a filas que ya no corresponden. Leer/escribir las 5 columnas o eliminar la columna del esquema si no se usa.

### 1.7 🟡 Carrera de IDs en incidencias

`reportarIncidencia` (`Incidencias.gs:8`) calcula `INC-AA-NNN` leyendo el máximo actual sin `LockService`. Dos reportes simultáneos pueden obtener el mismo ID. Envolver en `LockService.getScriptLock()` como ya se hace en `crearNuevaReserva`.

### 1.8 🟡 Email dentro del lock en `crearNuevaReserva`

`Codigo.gs:1021`: `sendConfirmationEmail()` (llamada externa lenta, ~1-2 s) se ejecuta **antes** de `lock.releaseLock()`. Todos los usuarios que intenten reservar a la vez esperan también el envío del correo. Mover el envío del email después de liberar el lock (igual que la limpieza de caché).

### 1.9 🟡 `createUsuario` invierte columnas (función zombie, pero peligrosa si se reactiva)

`AdminFunctions.gs:589`: `appendRow([email_usuario, nombre_completo, ...])` — el esquema es `[Nombre, Email, ...]`. Además escribe `'Activo'` (string) en una columna booleana, valor que `checkUserAuthorization` no reconoce como activo. Como está sin usar, lo mejor es **borrarla** (ver §4).

---

## 2. Seguridad

> Contexto importante: **cualquier persona que pueda cargar cualquier página de la webapp (incluida la pantalla de registro) puede invocar por consola cualquier función global del servidor vía `google.script.run`**. Por eso toda función sensible necesita su propia comprobación de permisos en el servidor; no basta con que el botón solo aparezca en el panel admin.

### 2.1 🔴 Aprobar/rechazar solicitudes recurrentes sin comprobación de admin

- `aprobarSolicitudRecurrente` (`ReservasRecurrentes.gs:332`) y `rechazarSolicitudRecurrente` (`:430`) **no verifican que el llamante sea admin**. Cualquier usuario (incluso no registrado, si está en el dominio) puede aprobar su propia solicitud recurrente llamando a la función por consola.
- La ruta de email `?action=aprobar_recurrente&id=...` (`doGet` → `handleAprobarRecurrenteDesdeEmail`, `Codigo.gs:678`) tampoco comprueba admin, a diferencia de `handleAdminApproval` que sí lo hace.

**Arreglo:** añadir al principio de ambas funciones y del handler de email:
```js
if (!isUserAdmin()) return { success: false, error: 'Permiso denegado' };
```

### 2.2 🔴 Cancelación por URL sin comprobar propiedad

`?action=cancel&id=<uuid>` (`handleEmailCancelation`, `Codigo.gs:1252`) cancela cualquier reserva sin comprobar sesión ni propiedad. El problema es que **los IDs de todas las reservas se envían a todos los usuarios** (`getStaticData().reservas` incluye `id_reserva` y `email_usuario` de todo el mundo), así que cualquier usuario puede cancelar reservas ajenas construyendo la URL.

Opciones (de menos a más cambio):
1. Comprobar en `handleEmailCancelation` que `Session.getActiveUser().getEmail()` coincide con el email de la reserva (o es admin). Como la webapp es `access: DOMAIN`, siempre hay sesión.
2. No exponer `id_reserva` en las reservas ajenas que devuelve `getStaticData` (el frontend solo necesita el ID para las propias, que ya van en `misReservasActivas`).

Se recomienda aplicar **ambas**.

### 2.3 🟠 `getStaticData` no exige estar autorizado

`Codigo.gs:366`: calcula `checkUserAuthorization` pero devuelve todos los datos (recursos, reservas con emails, mapa completo de usuarios) aunque `isAuthorized === false`. Un usuario del dominio sin aprobar puede leer todo llamando a la función desde la pantalla de registro. Añadir:
```js
if (!auth.isAuthorized) return { success: false, error: 'No autorizado' };
```
Lo mismo aplica a `getSolicitudesRecurrentes` (expone motivos y emails de todos; solo debería ser admin — el filtrado por usuario ya lo hace `getMisSolicitudesRecurrentes`), a `getIncidencias` (expone emails de quien reporta) y a `crearSolicitudRecurrente`/`reportarIncidencia` (deberían exigir usuario autorizado, no solo email de sesión).

### 2.4 🟠 XSS almacenado en la app de usuario (`scripts.html`)

`admin-scripts.html` tiene `escaparHTML()` y lo usa de forma bastante consistente, pero **`scripts.html` no escapa nada** (0 usos) e inyecta con `innerHTML` campos escritos por usuarios:

- `inc.Descripcion` y `inc.Notas_Admin` (líneas ~2771, ~2958-2966)
- `motivoDisplay` de recurrencias (~1433) y `notasHtml` de reservas (~1560)

Un usuario puede guardar `<img src=x onerror="...">` en la descripción de una incidencia y ejecutar JS en el navegador de cualquier otro usuario o del admin (robo de acciones admin vía `google.script.run`). Copiar el helper `escaparHTML` a `scripts.html` y aplicarlo a todo campo de origen usuario; también en los correos HTML del backend (interpolan `notas`, `motivo`, `descripcion` sin escapar).

### 2.5 🟡 `setXFrameOptionsMode(ALLOWALL)`

`Codigo.gs` lo aplica en todas las páginas. Solo es necesario si se incrusta la app en Google Sites u otro iframe; si no, usar `XFrameOptionsMode.DEFAULT` para evitar clickjacking.

### 2.6 🟡 Varios menores

- `procesarSolicitudRegistro` no tiene límite: cualquiera del dominio puede inundar de emails al admin. Un contador en `CacheService` por email (p. ej. máx. 3/día) lo mitiga.
- `cambiarURLManual` valida la URL solo con `includes('script.google.com')` — suficiente para un menú de admin, pero un regex `^https://script\.google\.com/.+/exec$` sería más robusto.
- `appsscript.json` no declara `oauthScopes` explícitos; Apps Script auto-detecta scopes amplios (Drive completo por `DriveApp` en `procesarUrlLogoDrive`). Declararlos explícitamente y valorar `drive.file` reduce lo que se pide a cada centro al autorizar.

---

## 3. Rendimiento

### 3.1 🔴 `crearNuevaReserva` lee la hoja Reservas ~5-6 veces

Cadena actual por cada reserva:

1. `validarRestriccionesConfiguracion` → `validarAntelacionMinima` → **`getStaticData()`** (lee Reservas 2 veces: `getReservasFrescas` + `getMyActiveReservationsData`) y `validarLimiteReservas` → **`getActiveReservations()`** (otra lectura completa).
2. `crearNuevaReserva` llama **otra vez** a `getStaticData()` (2 lecturas más).
3. `checkAvailability` → **`getActiveReservations()`** (otra más).

Además `checkUserAuthorization` (lectura completa de Usuarios) se ejecuta 2-3 veces por petición.

**Arreglo:** cargar una sola vez al principio (`const staticData = getStaticData(); const reservas = getActiveReservations();`) y pasar ambos como parámetros a `validarAntelacionMinima`, `validarLimiteReservas` y `checkAvailability` (esta última ya acepta `staticData`, solo falta que acepte también las reservas). Con esto la reserva pasa de ~6 lecturas de hoja a 1-2 y el lock se mantiene mucho menos tiempo.

### 3.2 🟠 `getStaticData` relee Reservas dos veces

`getReservasFrescas()` y `getMyActiveReservationsData(email)` hacen cada una su `sheetToObjects(Reservas)`. Las "mis reservas" se pueden derivar en memoria del array ya cargado:
```js
const misReservasActivas = reservas.filter(r => r.email_usuario.toLowerCase() === email.toLowerCase());
```

### 3.3 🟠 `purgarCache` regenera todo para poder purgar

`Codigo.gs:1988`: tras borrar la caché llama a `getStaticData()` (recarga todo el spreadsheet + autorización) solo para obtener los IDs de recursos y borrar sus claves de disponibilidad. Cada guardado del panel admin (`saveBatch*`) paga ese coste. Alternativa barata: leer solo la columna A de la hoja Recursos, o mantener una clave de "versión" que invalide las claves `DISP_*` sin enumerarlas.

### 3.4 🟠 `saveBatchDisponibilidad` escribe celda a celda

`AdminFunctions.gs:1044-1045`: dentro del bucle hace dos `setValue()` por cambio (2 llamadas a la API por fila). Con 50 cambios son 100 llamadas. Acumular los cambios y escribir por rangos (o leer la matriz, modificarla en memoria y hacer un único `setValues` de las columnas E-F).

### 3.5 🟡 Frontend

- `tailwind.js` (CDN runtime, ~300 KB de JS que compila CSS en el navegador) se carga en 4 páginas. Para producción, generar una CSS estática compilada (Tailwind CLI) y servirla desde el mismo GitHub Pages: menos peso, sin "flash" de estilos y sin dependencia del JIT.
- `checkUserAuthorization` se podría cachear por email (p. ej. 5 min en `CacheService`) para no leer la hoja Usuarios en cada llamada RPC.

---

## 4. Código zombie

Funciones definidas que **no se llaman ni desde el frontend ni desde otro punto del backend** (verificado con búsqueda cruzada de todas las llamadas `google.script.run` y referencias internas). Se pueden borrar; el panel admin usa las versiones `saveBatch*`:

| Archivo | Funciones |
|---|---|
| `AdminFunctions.gs` | `createRecurso`, `updateRecurso`, `deleteRecurso`, `createTramo`, `updateTramo`, `deleteTramo`, `updateDisponibilidad`, `generarDisponibilidadRecurso`, `createUsuario`, `updateUsuario`, `deleteUsuario`, `updateReservaAdmin`, `deleteReservaAdmin`, `enviarNotificacionCambioReserva`, `enviarNotificacionEliminacionReserva`, `getSolicitudesPendientesGlobal` |
| `Incidencias.gs` | `actualizarEstadoIncidencia` (el propio comentario dice "¿¿ESTO SOBRA??" — sí, sobra; la usada es `backend_actualizarIncidencia`) |
| `ReservasRecurrentes.gs` | `getMisReservasRecurrentes`, `getMisSolicitudesRecurrentes`, `getReservasDeGrupoRecurrente`, `contarSolicitudesPendientes`, `actualizarNotasRecurrencia` |

Otros restos:

- `Codigo.gs:29-38`: `CACHE_KEYS.STATIC_DATA` (V5) y `CACHE_TIMES.STATIC` ya no se usan para la caché activa (V6). Tras corregir §1.2, dejar solo `DISPONIBILIDAD` y `CONFIGURACION`.
- `Codigo.gs:737-770` (`getReservasFrescas`): la rama fallback tras `if (typeof getActiveReservations === 'function')` es inalcanzable (la función siempre existe). Reducir a `return getActiveReservations();`.
- `Codigo.gs:1995-1999` (`purgarCache`): la limpieza de `STATIC_DATA_V5/V4/V3` puede eliminarse una vez desplegado un tiempo (las claves caducan solas a las 6h).
- Duplicidad funcional: `adminCancelarReserva` (usada) vs `deleteReservaAdmin`/`updateReservaAdmin` (zombis) hacen lo mismo con distinto estilo. Nota: `adminCancelarReserva` **borra** la fila mientras el resto del sistema marca `Cancelada` — valorar unificar criterio para conservar histórico.

---

## 5. Sistema de actualización para las copias

Situación: cada centro hace una **copia de la hoja de cálculo** (con el script contenido) y la despliega. Hoy, cuando publicas mejoras, cada copia queda congelada con el código del día que se copió.

Hay tres estrategias viables en Apps Script; la recomendación es combinar la B ahora y migrar a la A como arquitectura definitiva.

### Opción A — Librería de Apps Script (actualización real y automática) ✅ recomendada a medio plazo

1. Mueves **todo** el código (`.gs` y `.html`) a un proyecto Apps Script independiente tuyo ("ReservasCore") y lo publicas como **librería** (compartida como "Cualquiera con el enlace: Lector").
2. La hoja plantilla que copian los centros solo lleva un *stub* mínimo que delega en la librería:

```js
// El script de la copia queda reducido a esto:
function doGet(e)  { return ReservasCore.doGet(e); }
function onOpen()  { ReservasCore.onOpen(); }

// google.script.run solo ve funciones del proyecto contenedor,
// así que se expone un despachador único:
function api(nombreFuncion, args) {
  return ReservasCore.api(nombreFuncion, args);
}
```

3. En la librería, `api()` valida el nombre contra una **lista blanca** de funciones públicas y las invoca. En el frontend se sustituyen las llamadas `google.script.run.crearNuevaReserva(datos)` por `google.script.run.api('crearNuevaReserva', datos)` (cambio mecánico; se puede envolver en un helper JS para tocar poco código).
4. Las copias referencian la librería **en modo desarrollo (HEAD)**: cada cambio que guardes en tu proyecto llega al instante a todas las copias, sin que los centros hagan nada. Si prefieres estabilidad, publicas versiones numeradas (pero entonces cada copia tendría que subir de versión a mano — HEAD es lo que te da el "auto-update").

- **Pros:** actualización real de lógica y de HTML para todos; un solo código que mantener; los datos (la hoja) siguen siendo de cada centro.
- **Contras:** las copias existentes tendrían que sustituir su script por el stub una vez (o rehacer la copia desde la nueva plantilla); si tu cuenta borra la librería, todas las copias caen; en modo HEAD un fallo tuyo llega a todos al momento (conviene tener una copia "beta" donde probar antes de guardar en el proyecto de la librería).
- Para los cambios de **esquema de la hoja** (columnas nuevas, etc.) mantén en la librería un `migrarEsquema()` versionado (ya tienes el patrón en `migrarIdSolicitudRecurrente`, que `getAdminData` ejecuta al vuelo): guarda `SCHEMA_VERSION` en `PropertiesService` y aplica solo las migraciones pendientes.

### Opción B — Avisador de versión (inmediato, ~50 líneas) ✅ recomendada ya

No actualiza el código, pero informa al admin de cada copia de que existe una versión nueva y cómo actualizar:

1. Declaras en el código `const APP_VERSION = '1.3.0';`.
2. Publicas en tu GitHub Pages (ya usas `maestroseb.github.io/recursos/`) un `version.json`:
   ```json
   { "version": "1.4.0", "novedades": "…", "instrucciones": "https://github.com/maestroseb/reservasrecursosysalas#actualizar" }
   ```
3. En el backend, una función cacheada 24 h:
   ```js
   function comprobarActualizacion() {
     const cache = CacheService.getScriptCache();
     let info = cache.get('VERSION_REMOTA');
     if (!info) {
       try {
         info = UrlFetchApp.fetch('https://maestroseb.github.io/recursos/version.json',
                                  { muteHttpExceptions: true }).getContentText();
         cache.put('VERSION_REMOTA', info, 86400);
       } catch (e) { return null; }
     }
     const remota = JSON.parse(info);
     return (remota.version !== APP_VERSION) ? { actual: APP_VERSION, ...remota } : null;
   }
   ```
4. El panel admin (y/o el menú `onOpen`) muestra un aviso "🔄 Nueva versión disponible (1.4.0)" con el changelog y el enlace a las instrucciones de actualización.

- **Pros:** trivial de implementar, sin riesgo, funciona con las copias tal y como se distribuyen hoy (basta con que actualicen una vez para recibir el avisador).
- **Contras:** la actualización sigue siendo manual (rehacer copia o copiar/pegar código); añade el scope de `UrlFetchApp` (un permiso más en la pantalla de autorización).

### Opción C — Auto-actualización con la API de Apps Script ❌ desaconsejada

Técnicamente el script puede reescribirse a sí mismo con `projects.updateContent` de la Apps Script API descargando los ficheros de GitHub, pero exige que **cada usuario** active la API en sus ajustes, requiere scopes muy sensibles (`script.projects`), y un fallo a mitad de escritura deja la copia rota. No compensa para este caso de uso.

### Hoja de ruta sugerida

1. **Ahora:** corregir los puntos rojos de §1 y §2, añadir la Opción B, y publicar nueva plantilla.
2. **Siguiente versión mayor:** reestructurar como librería (Opción A) — el avisador de la Opción B sirve entonces para avisar a las copias antiguas de que migren a la plantilla con stub, y a partir de ahí las actualizaciones son transparentes.

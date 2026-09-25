# Auditoría de la aplicación (sept. 2026)

## 🐞 Bucle de registro ("me vuelve a pedir que me registre")

**Causa:** `Session.getActiveUser().getEmail()` devuelve **vacío** cuando la web app (ejecutada como *Usuario que implementa*) se ha implementado desde una cuenta de **otro dominio** que el del profesorado (p.ej. una @gmail.com, o una cuenta distinta de @g.educaand.es). Entonces:
1. El profesor ve el registro con el email vacío y lo teclea a mano.
2. El admin lo aprueba y el email queda bien guardado en *Usuarios*.
3. En el siguiente acceso Google vuelve a dar email vacío → no coincide con nadie → registro otra vez. Siempre.

A ti no te pasa porque tu implementación y tu profesorado están en el mismo dominio.

**Solución para la compañera:** volver a implementar la app (Implementar > Nueva implementación) **desde su cuenta @g.educaand.es**, con *Ejecutar como: Usuario que implementa* y *Acceso: cualquier usuario de g.educaand.es*. Si el profe tiene varias cuentas abiertas en el navegador, que pruebe en incógnito.

**Cambios en código:** si el email llega vacío ya no se muestra el registro (inútil) sino una pantalla que explica la causa. El email del registro ya no se puede teclear (solo el de la sesión) y `checkUserAuthorization('')` ya no puede coincidir con una fila vacía.

## 🌐 Centros con otros dominios o cuentas Gmail
Con *Ejecutar como: Usuario que implementa*, Google **solo** facilita el email de los usuarios del **mismo dominio Workspace** que la cuenta que implementa. Por tanto:
- Centro Workspace (g.educaand.es, dominio propio…) cuyos usuarios son todos de ese dominio → funciona si se implementa desde ese dominio.
- Usuarios de otro dominio o cuentas @gmail.com → la app **no puede identificarlos** (la pantalla de diagnóstico lo explica).
- Para admitirlos haría falta otro sistema de identidad (p.ej. código de verificación por email). Pendiente de decisión.

## ✅ Fase 1 — aplicada (bajo riesgo)
| Cambio | Fichero |
|---|---|
| Pantalla de diagnóstico si Google no da el email; registro solo con email de sesión | Codigo.gs, registro.html |
| `Activo`/`Admin` aceptan "Sí" con tilde | Codigo.gs |
| Aprobar usuario borraba una clave de caché que ya no existía (V5 vs V6) | Codigo.gs |
| `purgarCache` ahora también limpia la caché de Configuración (cambios tardaban 5 min) | Codigo.gs |
| `ejecutarSetupVinculado` no se puede re-ejecutar si ya está instalado (cualquiera podía hacerse admin y sobrescribir datos) | Setup.gs |
| Aprobar/rechazar solicitud recurrente y enlace de email: solo admin (antes un profe podía auto-aprobarse) | ReservasRecurrentes.gs, Codigo.gs |
| `procesarUrlLogoDrive` solo admin (hacía públicos ficheros de Drive) | Codigo.gs |
| Incidencias: spinner infinito para no-admin | scripts.html |
| Selector de cursos ignoraba el modo "listado" para no-admin | scripts.html |
| No se pide `getAdminData` a usuarios no-admin (llamada inútil en cada carga) | admin-scripts.html |
| Buscador de cursos daba error (id `searchCursos` inexistente) y se ejecutaba 2 veces | admin-scripts.html, admin-panel.html |
| Tras guardar cursos, cambiar de pestaña restauraba la lista antigua | admin-scripts.html |
| XSS en panel de pendientes (`diasDisplay`) y `escaparHTML` no escapaba comillas | admin-scripts.html |

## ✅ Fase 2 — aplicada
| Cambio | Nota |
|---|---|
| Config: booleanos se leían como `1` → **modo mantenimiento y copia al admin ahora SÍ funcionan** | Revisar sus valores en *Config* tras desplegar |
| Funciones internas privadas (sufijo `_`): emails, `generarReservasDesdeRecurrente_`, cancelaciones masivas, migración | Ya no se pueden invocar desde el navegador |
| Enlace de cancelación del email firmado (HMAC); enlaces antiguos solo valen al dueño o a un admin | Funciona aunque Google no identifique al usuario |
| `getStaticData`, `getIncidencias`, `reportarIncidencia` exigen usuario autorizado | |
| `cantidad` validada (≥1) | |
| Antelación mínima calculada en hora local (antes UTC: 1-2 h de desfase) | Ya no se puede reservar un tramo empezado |
| Bloquear disponibilidad cancela las reservas de las recurrencias afectadas (buscaba columna inexistente) | |
| Editar tramos de una recurrencia ya no cancela reservas normales de otros usuarios | |
| Matriz de disponibilidad: nombre/área de usuario con columnas correctas | |
| `LockService` en recurrentes, cancelación de grupo, eliminar tramo, cancelación admin y guardado de usuarios | |
| Usuarios: "FALSE"/"No" ya no se reactivan; se conserva la columna *Especialidad* | |
| `checkIfAdmin` y `getAdminsEmails_` usan la misma regla que el login (respetan *Activo*, aceptan "Sí") | |
| XSS escapado en `scripts.html` (incidencias, motivo, notas, cursos, nombres) y en todos los emails HTML | |
| `USER_EMAIL`/`USER_NAME` sin comillas dobles | |

### Pendiente (cambia comportamiento, a decidir)
- Las recurrencias aprobadas días después generan reservas en fechas pasadas y no comprueban *Disponibilidad*.
- Renombrar el ID de un tramo no actualiza Reservas ni SolicitudesRecurrentes.
- Si se aprueba un usuario por email con el panel admin abierto, el siguiente "Guardar usuarios" lo borra.

## 🚀 Fase 3 — velocidad de carga
1. **No incluir `admin-panel` + `admin-scripts` (≈320 KB, 60 % de la página) para no-admin.** Requiere Fase 1 (ya hecha).
2. **Tailwind se compila en el navegador** (tailwind.js runtime): sustituir por CSS precompilado. Es la mejora de carga más grande.
3. Iconify con `defer`.
4. `crearNuevaReserva`: lee Reservas ~6 veces y Usuarios 3 dentro del lock, envía el email antes de liberar el lock y purga la caché estática sin necesidad.
5. `getAdminData` ejecuta la migración `migrarIdSolicitudRecurrente` en cada carga; `purgarCache` recalcula `getStaticData` entero.
6. `getAppConfig` sin caché en cada `doGet`; `checkUserAuthorization` se ejecuta 2 veces por carga.

## 🧟 Fase 4 — código zombie (≈1.500 líneas, borrado sin riesgo funcional)
- **AdminFunctions.gs**: `createRecurso/updateRecurso/deleteRecurso`, `generarDisponibilidadRecurso`, `createTramo/updateTramo/deleteTramo`, `updateDisponibilidad`, `createUsuario` (¡escribe columnas en orden incorrecto!), `updateUsuario/deleteUsuario`, `updateReservaAdmin/deleteReservaAdmin` + sus emails, `getSolicitudesPendientesGlobal`.
- **ReservasRecurrentes.gs**: `getMisSolicitudesRecurrentes`, `actualizarNotasRecurrencia`, `getReservasDeGrupoRecurrente`, `getMisReservasRecurrentes`, `contarSolicitudesPendientes`.
- **Incidencias.gs**: `actualizarEstadoIncidencia`. **Codigo.gs**: fallback inalcanzable en `getReservasFrescas`, `CACHE_KEYS.STATIC_DATA`, `CACHE_TIMES.STATIC`.
- **scripts.html**: `ajustarHeaderSegunIncidencias` (110 líneas), `cambiarEstadoIncidencia`, `debounce`.
- **admin-scripts.html**: lista de recurrentes sin contenedor en el HTML (~660 líneas), `aprobarSolicitudRapida`, `aprobarTodasPendientes`, `verDetalleSolicitud`, `ICONIFY_ICONS`, `getDispKey`, `setVal/getVal/...`.
- **admin-panel.html**: modales `modalRecurso` y `modalDisponibilidad` (~325 líneas).
- **styles.html**: ~450 líneas de CSS sin uso y reglas duplicadas (`.spinner`, `dialog`, `.admin-subtab`…).
- Config sin efecto: `horas_cancelacion`, `exigir_motivo`, `permitir_multitramo`, `max_tramos_simultaneos` se guardan pero ningún código las aplica.

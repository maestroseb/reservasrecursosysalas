# Auditoría de la aplicación (sept. 2026)

## 🐞 Bucle de registro ("me vuelve a pedir que me registre")

**Causa:** `Session.getActiveUser().getEmail()` devuelve **vacío** cuando la web app (ejecutada como *Usuario que implementa*) se ha implementado desde una cuenta de **otro dominio** que el del profesorado (p.ej. una @gmail.com, o una cuenta distinta de @g.educaand.es). Entonces:
1. El profesor ve el registro con el email vacío y lo teclea a mano.
2. El admin lo aprueba y el email queda bien guardado en *Usuarios*.
3. En el siguiente acceso Google vuelve a dar email vacío → no coincide con nadie → registro otra vez. Siempre.

A ti no te pasa porque tu implementación y tu profesorado están en el mismo dominio.

**Solución para la compañera:** volver a implementar la app (Implementar > Nueva implementación) **desde su cuenta @g.educaand.es**, con *Ejecutar como: Usuario que implementa* y *Acceso: cualquier usuario de g.educaand.es*. Si el profe tiene varias cuentas abiertas en el navegador, que pruebe en incógnito.

**Cambios en código:** si el email llega vacío ya no se muestra el registro (inútil) sino una pantalla que explica la causa. El email del registro ya no se puede teclear (solo el de la sesión) y `checkUserAuthorization('')` ya no puede coincidir con una fila vacía.

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

## ⚠️ Fase 2 — pendiente de tu OK (cambian comportamiento)
1. **Modo mantenimiento y "copia al admin" nunca funcionan**: los booleanos de Config se convierten a `1` y se comparan con `true`. Al arreglarlo *empezarán a funcionar*: revisa antes sus valores en Config.
2. **Seguridad RPC**: con `google.script.run` cualquier usuario puede llamar funciones internas: envío de emails arbitrarios desde tu cuenta (`sendConfirmationEmail`, `enviarEmail*`…), `generarReservasDesdeRecurrente`, `cancelarReservasFuturasDeTramos`, `procesarCancelacionesPorDisponibilidad`, `migrarIdSolicitudRecurrente`. Solución: renombrarlas con `_` final (privadas).
3. **Cancelar por enlace de email** (`?action=cancel&id=`) no comprueba quién hace clic: cualquiera puede cancelar reservas ajenas.
4. `getStaticData` / `getIncidencias` devuelven emails y datos de todos incluso a no registrados.
5. `cantidad` negativa en recursos agrupados rompe el aforo.
6. `validarAntelacionMinima` usa UTC: permite reservar un tramo ya empezado (1-2 h de desfase).
7. Bloquear disponibilidad **no cancela** las reservas recurrentes (busca columna `tipo_reserva` inexistente).
8. `cancelarReservasFuturasDeTramos` cancela también reservas normales de otros usuarios en ese tramo.
9. `getDatosMatrizUnificada`: mapa de usuarios con columnas cruzadas (nombre/email).
10. Falta `LockService` en escrituras de admin/recurrentes (riesgo de doble reserva o fila equivocada).
11. `getAdminData` marca como activo a un usuario con "FALSE"/"No" en texto y al guardar lo reactiva.
12. `saveBatchUsuarios` no reescribe la columna Especialidad (se desplaza al borrar usuarios).
13. Escapado XSS en `scripts.html` (incidencias, motivo, notas, nombres) y en emails HTML.
14. `USER_EMAIL` posiblemente llega con comillas dobles (`<?= JSON.stringify ?>`): comprueba en consola `window.USER_EMAIL`; si sale `"\"x@..\""` las marcas "Tú"/"Reservado por ti" no funcionan.

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

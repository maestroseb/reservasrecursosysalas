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

## ✅ Fase 3 — velocidad de carga (aplicada)
| Cambio | Efecto |
|---|---|
| Tailwind precompilado en `tailwind-css.html` (62 KB) en vez de `tailwind.js` compilando en el navegador | Sin bloqueo de render ni recompilación en cada cambio del DOM; sin depender de GitHub Pages/CDN |
| `admin-panel` + `admin-scripts` solo para administradores | ~250 KB menos por carga para el profesorado |
| Iconify con `defer` | No bloquea el render |
| `getStaticData`: Reservas se lee 1 vez (antes 2) | |
| `crearNuevaReserva`: Reservas 1 lectura (antes ~4), sin releer todas las hojas, email tras liberar el lock, sin purgar la caché estática | Reservas más rápidas y menos esperas con varios usuarios |
| `getAppConfig` usa la configuración cacheada | Una lectura menos en cada carga |
| `purgarCache` ya no regenera todos los datos | Guardados del panel admin más rápidos |
| Migración de `ID_Solicitud_Recurrente` solo una vez y en bloque | Apertura del panel admin más rápida |

⚠️ **Nuevo flujo de trabajo**: si añades clases de Tailwind nuevas hay que regenerar `tailwind-css.html` (README > "Regenerar el CSS de Tailwind").

## ✅ Fase 4 — código zombie (aplicada, ~2.600 líneas)
- **Servidor**: 22 funciones sin ninguna llamada (CRUD antiguos de recursos/tramos/usuarios/reservas, emails asociados, consultas de recurrentes, `actualizarEstadoIncidencia`), constantes de caché obsoletas y fallback inalcanzable.
- **Cliente**: `debounce`, `cambiarEstadoIncidencia`, `ajustarHeaderSegunIncidencias`, `hayIncidenciasPendientes`, `aprobarSolicitudRapida`, `aprobarTodasPendientes`, `renderizarSolicitudesRecurrentes`, `verDetalleSolicitud`, filtros de la lista antigua, helpers `setVal/getVal…`, `ICONIFY_ICONS`.
- **HTML**: modales `modalRecurso` y `modalDisponibilidad`.
- **CSS**: 70 reglas de 33 clases sin uso.

### Pendiente / a decidir
- **Editar tramos de una recurrencia aprobada** (`abrirModalEditarTramosRecurrencia` + `editarTramosRecurrencia`): el código existe y funciona, pero ya no hay ningún botón que lo abra. ¿Recuperarlo en la UI o eliminarlo?
- Reglas CSS duplicadas (`.spinner`, `dialog`, `.admin-subtab`): unificarlas cambia el aspecto, no lo he tocado.
- `button:focus { outline: none !important }` quita el foco visible con teclado (accesibilidad).
- Config sin efecto: `horas_cancelacion`, `exigir_motivo`, `permitir_multitramo`, `max_tramos_simultaneos` se guardan pero ningún código las aplica.

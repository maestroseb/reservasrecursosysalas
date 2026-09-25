# Historial de versiones

## v1.5.0 — Auditoría completa (septiembre 2026)

### Nuevo
- **Multitramo**: con `permitir_multitramo` activado en *Config*, al reservar se puede elegir cuántos tramos seguidos coger (hasta `max_tramos_simultaneos`). Solo se ofrecen tramos libres consecutivos; en recursos agrupados la cantidad máxima es la menor disponible entre ellos. La reserva es "todo o nada", cada tramo cuenta como una reserva para el límite, y el email incluye un enlace de cancelación por tramo.
- **Motivo obligatorio** (`exigir_motivo`): si se activa, las notas de la reserva pasan a ser obligatorias (se valida también en el servidor).
- **Tope de cancelación** (`horas_cancelacion`): un usuario no puede cancelar con menos de X horas de antelación al inicio del tramo, ni desde la app ni desde el enlace del email. Los administradores quedan exentos. No afecta a cancelar una recurrencia completa.
- **Usuarios en bloque**: botón *Importar* para pegar la lista desde Excel/Sheets/Séneca o subir un CSV (detecta sola la columna del email; "admin"/"sí" marca administrador; ignora cabecera, repetidos y los que ya existen). Botón para **vaciar la lista** de una vez (conserva siempre a los administradores). Nada se aplica hasta pulsar *Guardar*.
- Tabla de usuarios más compacta (una línea por usuario).
- **Guardado en segundo plano** en el panel admin (recursos, horarios, usuarios, cursos, tramos y configuración): ya no bloquea la pantalla; aparece "Guardando…" y un aviso al terminar. Si falla, los cambios quedan pendientes y la barra de Guardar vuelve a aparecer.
- Si el logo del centro no carga, se oculta en vez de mostrar una imagen rota.

### Corregido
- El panel de Configuración mostraba 2 h de tope de cancelación y 30 min de antelación cuando el valor guardado era 0 (al guardar se habría cambiado sin querer).
- **Bucle de registro**: cuando Google no facilita el email del usuario (app implementada desde otro dominio o cuenta @gmail.com) se muestra una pantalla que explica la causa en lugar de pedir el registro una y otra vez.
- **Modo mantenimiento** y **copia de reservas al admin** no funcionaban nunca (los valores de Config se leían mal). ⚠️ Revisa sus valores en la hoja *Config* al actualizar.
- La antelación mínima se calculaba en UTC (1-2 h de desfase): se podía reservar un tramo ya empezado.
- Bloquear un tramo en *Disponibilidad* no cancelaba las reservas recurrentes afectadas.
- Editar los tramos de una recurrencia cancelaba también reservas normales de otros usuarios.
- Usuarios marcados como "FALSE"/"No" se reactivaban al guardar desde el panel; la columna *Especialidad* se desplazaba al borrar usuarios.
- "Sí" con tilde no se reconocía en *Activo*/*Admin*.
- Incidencias: spinner infinito para usuarios no administradores.
- Selector de cursos ignoraba el modo "listado" para no administradores.
- Buscador de cursos del panel admin daba error; los cursos guardados reaparecían con la versión antigua al cambiar de pestaña.
- Los paneles de solicitudes recurrentes pendientes no se refrescaban tras aprobar/rechazar.
- Área/especialidad del usuario no aparecía en la matriz de disponibilidad.
- Al aprobar un usuario por email no se limpiaba la caché correcta.

### Seguridad
- Aprobar/rechazar recurrentes, cambiar el logo y reinstalar el sistema requieren ser administrador.
- Las funciones internas (envío de emails, generación/cancelación masiva de reservas) ya no se pueden invocar desde el navegador.
- Los enlaces de cancelación de los emails van firmados: nadie puede cancelar reservas ajenas. Los enlaces de emails antiguos solo funcionan para el dueño de la reserva o un admin.
- Los usuarios no registrados ya no pueden leer reservas ni emails de otros.
- Textos de usuario escapados en la app y en todos los emails (evita inyección de HTML).
- Cantidad reservada validada (≥ 1).
- Bloqueo (LockService) en todas las escrituras concurrentes.

### Rendimiento
- Tailwind CSS precompilado (antes se compilaba en el navegador en cada carga) y sin dependencia de CDN externo.
- El código del panel admin (~300 KB) solo se envía a administradores.
- Iconify se carga en diferido.
- Menos lecturas de hojas al cargar la página, reservar y guardar en el panel admin; el email de confirmación se envía tras liberar el bloqueo.
- La migración de datos del panel admin se ejecuta una sola vez.

### Limpieza
- Eliminadas ~2.600 líneas de código y CSS sin uso.
- Eliminado el modal "Editar tramos" de recurrencias (no era accesible y el detalle de la recurrencia ya permite quitar tramos uno a uno).
- Reglas CSS duplicadas unificadas; el foco vuelve a verse al navegar con teclado.
- Versión visible al pie de la página.

## v1.4.1
Versión estable previa a la auditoría.

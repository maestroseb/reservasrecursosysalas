/**
 * ===========================================================================
 * 🔐 ACCESO CON CÓDIGO DE VERIFICACIÓN (v1.6.0)
 * ===========================================================================
 * Si Google facilita el email del usuario (mismo dominio que quien implementa),
 * todo funciona como siempre. Si no (otro dominio, cuentas @gmail.com...), el
 * usuario se identifica con un código de 6 dígitos que recibe por email y la app
 * lo recuerda en ese navegador durante DIAS_SESION días.
 *
 * - Sesión: token firmado (HMAC) "email|caducidad", sin almacenamiento en servidor.
 *   Desactivar al usuario en la hoja Usuarios le corta el acceso al instante.
 * - Códigos: CacheService (10 min, 5 intentos) con límites de envío.
 * - Config opcional "dominios_acceso" (p. ej. "g.educaand.es, gmail.com"):
 *   si existe, solo se envían códigos a esos dominios.
 */

const DIAS_SESION = 30;
const MINUTOS_CODIGO = 10;
const MAX_INTENTOS_CODIGO = 5;
const MAX_CODIGOS_POR_EMAIL = 3;         // cada 15 minutos
const MAX_CODIGOS_REGISTRADOS_HORA = 40; // usuarios ya dados de alta
const MAX_CODIGOS_NUEVOS_HORA = 8;       // correos no registrados (solicitudes de alta): cupo aparte
const MAX_CODIGOS_NUEVOS_DIA = 25;       // así un abuso no agota el cupo diario de correo ni bloquea a los profes
const MAX_FALLOS_DIA = 10;               // códigos erróneos por correo en 24 h (anti fuerza bruta)

// Token de sesión de la petición actual (lo fija ejecutarConSesion o doGet)
let TOKEN_PETICION_ = null;
// true solo mientras se procesa un enlace de aprobación firmado (email al admin)
let ADMIN_POR_ENLACE_FIRMADO_ = false;

/**
 * Email del usuario de la petición actual: identidad de Google si existe;
 * si no, la de la sesión verificada por código.
 */
function emailActual_() {
  const email = Session.getActiveUser().getEmail();
  if (email) return email;
  if (TOKEN_PETICION_) {
    const datos = validarTokenSesion_(TOKEN_PETICION_);
    if (datos) return datos.email;
  }
  return '';
}

/* ---------- Firma (reutiliza el secreto de los enlaces de cancelación) ---------- */
function firmarTexto_(texto) {
  const firma = Utilities.computeHmacSha256Signature(String(texto), getSecretoEnlaces_());
  return Utilities.base64EncodeWebSafe(firma).replace(/=+$/, '');
}

function crearTokenSesion_(email) {
  const caduca = Date.now() + DIAS_SESION * 24 * 3600 * 1000;
  const payload = String(email).toLowerCase().trim() + '|' + caduca;
  return Utilities.base64EncodeWebSafe(payload, Utilities.Charset.UTF_8).replace(/=+$/, '') + '.' + firmarTexto_('ses|' + payload);
}

function validarTokenSesion_(token) {
  try {
    const partes = String(token || '').split('.');
    if (partes.length !== 2) return null;
    let b64 = partes[0];
    while (b64.length % 4) b64 += '=';
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(b64)).getDataAsString('UTF-8');
    if (firmarTexto_('ses|' + payload) !== partes[1]) return null;
    const partes2 = payload.split('|');
    if (partes2.length !== 2) return null;
    const email = partes2[0];
    const caduca = Number(partes2[1]);
    if (!email || !isFinite(caduca) || caduca < Date.now()) return null;
    return { email: email };
  } catch (e) {
    return null;
  }
}

/* URL de la app sin el tramo de dominio (/a/macros/<dominio>/...): la forma genérica
   funciona para usuarios de cualquier dominio o Gmail. */
function urlApp_() {
  return String(ScriptApp.getService().getUrl() || '').replace(/\/a\/macros\/[^/]+\//, '/macros/');
}

/* ---------- Tickets de un solo uso para entrar en la app (2 min) ---------- */
function crearTicket_(token) {
  const ticket = Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('TICKET_' + ticket, token, 120);
  return ticket;
}

function consumirTicket_(ticket) {
  if (!/^[a-f0-9]{32}$/i.test(String(ticket || ''))) return null;
  const cache = CacheService.getScriptCache();
  const token = cache.get('TICKET_' + ticket);
  if (!token) return null;
  cache.remove('TICKET_' + ticket);
  const datos = validarTokenSesion_(token);
  return datos ? { email: datos.email, token: token } : null;
}

/* ---------- API pública para la pantalla de acceso (acceso.html) ---------- */

function solicitarCodigoAcceso(emailEntrada) {
  const email = String(emailEntrada || '').toLowerCase().trim();
  if (!/^[^\s@|]+@[^\s@|]+\.[^\s@|]+$/.test(email)) {
    return { success: false, error: 'Escribe un correo electrónico válido.' };
  }

  const dominios = String(getConfigValue('dominios_acceso', '') || '')
    .split(',').map(d => d.trim().toLowerCase().replace(/^@/, '')).filter(Boolean);
  if (dominios.length && dominios.indexOf(email.split('@')[1]) === -1) {
    return { success: false, error: 'Este centro solo admite correos de: ' + dominios.join(', ') };
  }

  const cache = CacheService.getScriptCache();
  const registrado = checkUserAuthorization(email).isAuthorized;
  const ahora = new Date();
  const claveHora = (registrado ? 'OTPGR_' : 'OTPGN_') + Utilities.formatDate(ahora, 'UTC', 'yyyyMMddHH');
  const claveDia = 'OTPGND_' + Utilities.formatDate(ahora, 'UTC', 'yyyyMMdd');

  // Leer, comprobar y anotar los contadores con bloqueo (evita saltarse los límites con peticiones en paralelo)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'Sistema ocupado. Inténtalo en unos segundos.' };
  let codigo;
  try {
    if (Number(cache.get('OTPF_' + email) || 0) >= MAX_FALLOS_DIA) {
      return { success: false, error: 'Demasiados códigos erróneos para este correo. Inténtalo mañana o avisa al administrador.' };
    }
    const envios = Number(cache.get('OTPN_' + email) || 0);
    if (envios >= MAX_CODIGOS_POR_EMAIL) {
      return { success: false, error: 'Has pedido varios códigos seguidos. Espera unos minutos y revisa tu correo (también la carpeta de spam).' };
    }
    const enHora = Number(cache.get(claveHora) || 0);
    const enDia = Number(cache.get(claveDia) || 0);
    if (enHora >= (registrado ? MAX_CODIGOS_REGISTRADOS_HORA : MAX_CODIGOS_NUEVOS_HORA) ||
        (!registrado && enDia >= MAX_CODIGOS_NUEVOS_DIA)) {
      return { success: false, error: 'Se han enviado demasiados códigos. Inténtalo más tarde.' };
    }

    // Código de 6 dígitos a partir de un UUID aleatorio
    codigo = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 1000000).padStart(6, '0');
    cache.put('OTP_' + email, JSON.stringify({ h: firmarTexto_('otp|' + email + '|' + codigo), n: 0 }), MINUTOS_CODIGO * 60);
    cache.put('OTPN_' + email, String(envios + 1), 15 * 60);
    cache.put(claveHora, String(enHora + 1), 3600);
    if (!registrado) cache.put(claveDia, String(enDia + 1), 86400);
  } finally {
    lock.releaseLock();
  }

  const appName = (getAppConfig().appName) || 'Sistema de Reservas';
  MailApp.sendEmail({
    to: email,
    subject: `${codigo} es tu código de acceso - ${appName}`,
    htmlBody: `
      <div style="font-family: sans-serif; max-width: 420px; margin: auto; border: 1px solid #e5e7eb; border-radius: 12px; overflow: hidden;">
        <div style="background: linear-gradient(135deg, #4f46e5, #2563eb); padding: 18px; color: white; text-align: center;">
          <h2 style="margin: 0; font-size: 18px;">${escHtml_(appName)}</h2>
        </div>
        <div style="padding: 24px; text-align: center; color: #374151;">
          <p style="margin: 0 0 12px;">Tu código de acceso es:</p>
          <div style="font-size: 34px; font-weight: bold; letter-spacing: 8px; color: #1f2937;">${codigo}</div>
          <p style="margin: 16px 0 0; font-size: 13px; color: #6b7280;">Caduca en ${MINUTOS_CODIGO} minutos. Si no lo has pedido tú, ignora este correo.</p>
        </div>
      </div>`
  });
  return { success: true };
}

function verificarCodigoAcceso(emailEntrada, codigoEntrada) {
  const email = String(emailEntrada || '').toLowerCase().trim();
  const codigo = String(codigoEntrada || '').replace(/\D/g, '');
  const cache = CacheService.getScriptCache();

  // Comprobación y anotación de intentos con bloqueo: sin él, peticiones en paralelo leerían n=0 a la vez
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'Sistema ocupado. Inténtalo en unos segundos.' };
  try {
    const guardado = cache.get('OTP_' + email);
    if (!guardado) return { success: false, error: 'El código ha caducado. Pide uno nuevo.' };

    const datos = JSON.parse(guardado);
    if (datos.n >= MAX_INTENTOS_CODIGO) {
      cache.remove('OTP_' + email);
      return { success: false, error: 'Demasiados intentos. Pide un código nuevo.' };
    }
    if (firmarTexto_('otp|' + email + '|' + codigo) !== datos.h) {
      datos.n++;
      cache.put('OTP_' + email, JSON.stringify(datos), MINUTOS_CODIGO * 60);
      const fallos = Number(cache.get('OTPF_' + email) || 0) + 1;
      cache.put('OTPF_' + email, String(fallos), 86400);
      if (fallos >= MAX_FALLOS_DIA) cache.remove('OTP_' + email);
      return { success: false, error: `Código incorrecto (quedan ${Math.max(0, MAX_INTENTOS_CODIGO - datos.n)} intentos).` };
    }
    cache.remove('OTP_' + email);
  } finally {
    lock.releaseLock();
  }

  const token = crearTokenSesion_(email);
  return { success: true, token: token, ticket: crearTicket_(token), url: urlApp_() };
}

/** Sesión recordada en el navegador: devuelve un ticket para entrar en la app. */
function canjearTokenPorTicket(token) {
  if (!validarTokenSesion_(token)) return { success: false };
  return { success: true, ticket: crearTicket_(token), url: urlApp_() };
}

/**
 * Pasarela para usuarios con sesión por código: el cliente envía el nombre de la
 * función, el token y los argumentos.
 * SEGURIDAD: lista CERRADA de funciones (nunca globalThis[nombre], que permitiría
 * ejecutar eval u otras funciones internas). Si la app empieza a llamar a una
 * función nueva del servidor, hay que añadirla aquí.
 */
function funcionesPasarela_() {
  return {
    actualizarMotivoRecurrencia, adminCancelarReserva, aprobarSolicitudRecurrente,
    backend_actualizarIncidencia, backend_toggleMantenimiento, cancelarGrupoRecurrente,
    cancelarRecurrenciaAprobada, cancelarReservaCliente, cargarDisponibilidadRecurso,
    crearNuevaReserva, crearRecurrenteDirecta, crearReservasMultitramo, crearSolicitudRecurrente,
    eliminarTramoDeRecurrencia, getAdminData, getConflictosRecurrencia, getDatosMatrizUnificada,
    getIncidencias, getSolicitudesRecurrentes, getStaticData, procesarSolicitudRegistro,
    procesarUrlLogoDrive, rechazarSolicitudRecurrente, reportarIncidencia, saveAllCursos,
    saveBatchConfig, saveBatchDisponibilidadConValidacion, saveBatchRecursos, saveBatchTramos,
    saveBatchUsuarios
  };
}

function ejecutarConSesion(nombre, token, args) {
  const permitidas = funcionesPasarela_();
  if (typeof nombre !== 'string' || !Object.prototype.hasOwnProperty.call(permitidas, nombre)) {
    throw new Error('Función no permitida: ' + nombre);
  }
  TOKEN_PETICION_ = token;
  return permitidas[nombre].apply(null, Array.isArray(args) ? args : []);
}

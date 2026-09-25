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
const MAX_CODIGOS_POR_EMAIL = 3;     // cada 15 minutos
const MAX_CODIGOS_GLOBAL_HORA = 60;  // protección anti-abuso del correo

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
    const [email, caduca] = payload.split('|');
    if (!email || Number(caduca) < Date.now()) return null;
    return { email: email };
  } catch (e) {
    return null;
  }
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
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, error: 'Escribe un correo electrónico válido.' };
  }

  const dominios = String(getConfigValue('dominios_acceso', '') || '')
    .split(',').map(d => d.trim().toLowerCase().replace(/^@/, '')).filter(Boolean);
  if (dominios.length && dominios.indexOf(email.split('@')[1]) === -1) {
    return { success: false, error: 'Este centro solo admite correos de: ' + dominios.join(', ') };
  }

  const cache = CacheService.getScriptCache();
  const envios = Number(cache.get('OTPN_' + email) || 0);
  if (envios >= MAX_CODIGOS_POR_EMAIL) {
    return { success: false, error: 'Has pedido varios códigos seguidos. Espera unos minutos y revisa tu correo (también la carpeta de spam).' };
  }
  const claveGlobal = 'OTPG_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMddHH');
  const globales = Number(cache.get(claveGlobal) || 0);
  if (globales >= MAX_CODIGOS_GLOBAL_HORA) {
    return { success: false, error: 'Se han enviado demasiados códigos en la última hora. Inténtalo más tarde.' };
  }

  // Código de 6 dígitos a partir de un UUID aleatorio
  const codigo = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 1000000).padStart(6, '0');
  cache.put('OTP_' + email, JSON.stringify({ h: firmarTexto_('otp|' + email + '|' + codigo), n: 0 }), MINUTOS_CODIGO * 60);
  cache.put('OTPN_' + email, String(envios + 1), 15 * 60);
  cache.put(claveGlobal, String(globales + 1), 3600);

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
    return { success: false, error: `Código incorrecto (quedan ${MAX_INTENTOS_CODIGO - datos.n} intentos).` };
  }

  cache.remove('OTP_' + email);
  const token = crearTokenSesion_(email);
  return { success: true, token: token, ticket: crearTicket_(token), url: ScriptApp.getService().getUrl() };
}

/** Sesión recordada en el navegador: devuelve un ticket para entrar en la app. */
function canjearTokenPorTicket(token) {
  if (!validarTokenSesion_(token)) return { success: false };
  return { success: true, ticket: crearTicket_(token), url: ScriptApp.getService().getUrl() };
}

/**
 * Pasarela para usuarios con sesión por código: el cliente envía el nombre de la
 * función, el token y los argumentos. Solo se permiten las funciones que ya son
 * públicas para google.script.run (nunca las terminadas en "_").
 */
const FUNCIONES_NO_PASARELA_ = ['ejecutarConSesion', 'doGet', 'onOpen', 'include', 'ejecutarSetupVinculado',
  'diagnosticarArchivos', 'repararInstalacionYGuardarURL', 'mostrarInstruccionesSidebar', 'mostrarURLRapido', 'cambiarURLManual'];

function ejecutarConSesion(nombre, token, args) {
  if (typeof nombre !== 'string' || /_$/.test(nombre) || FUNCIONES_NO_PASARELA_.indexOf(nombre) !== -1) {
    throw new Error('Función no permitida: ' + nombre);
  }
  const fn = globalThis[nombre];
  if (typeof fn !== 'function') throw new Error('Función no encontrada: ' + nombre);
  TOKEN_PETICION_ = token;
  return fn.apply(null, Array.isArray(args) ? args : []);
}

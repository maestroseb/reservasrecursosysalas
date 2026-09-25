/**
 * SISTEMA DE RESERVAS - CÓDIGO PRINCIPAL
 * Versión: ver APP_VERSION (historial en CHANGELOG.md)
 * 
 * Este archivo contiene las funciones principales del sistema.
 * Las funciones de administración están en AdminFunctions.gs
 * Las funciones de setup están en Setup.gs
 */

// Versión de la aplicación (se muestra al pie de la página)
const APP_VERSION = '1.5.0';

function getDB() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ============================================
   CONSTANTES
   ============================================ */

const SHEETS = {
  RECURSOS: "Recursos",
  TRAMOS: "Tramos",
  DISPONIBILIDAD: "Disponibilidad",
  RESERVAS: "Reservas",
  USUARIOS: "Usuarios",
  CURSOS: "Cursos",
  INCIDENCIAS: 'INCIDENCIAS',
  CONFIG: 'Config'
};

const CACHE_KEYS = {
  DISPONIBILIDAD: 'DISP_',
  CONFIGURACION: "configuracion_v1"
};

const CACHE_TIMES = {
  DISPONIBILIDAD: 1800
};


/* ============================================
   MENÚ DINÁMICO Y ACTIVACIÓN DEL SISTEMA
   ============================================ */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const isInstalled = props.getProperty('SETUP_COMPLETED') === 'true';

  const menu = ui.createMenu('🗓️ Sistema de Reservas');

  if (!isInstalled) {
    // NO INSTALADO -> Mostrar instrucciones
    menu.addItem('🚀 Instrucciones de Instalación', 'mostrarInstruccionesSidebar');
    menu.addToUi();

    // Toast discreto
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Haz clic en "🗓️ Sistema de Reservas" para ver cómo activar el sistema',
      '⚠️ Sistema sin activar',
      15
    );

  } else {
    // YA INSTALADO -> Menú normal
    menu.addItem('🔗 Ver URL de acceso', 'mostrarURLRapido');
    menu.addSeparator();
    menu.addItem('⚙️ Cambiar URL manualmente', 'cambiarURLManual');
    menu.addToUi();
  }
}

function mostrarInstruccionesSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setTitle('🚀 Guía de Instalación')
    .setWidth(420);

  SpreadsheetApp.getUi().showSidebar(html);
}

function mostrarURLRapido() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('WEB_APP_URL');

  if (!url) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Sistema no activado',
      'Primero debes activar el sistema desde el menú.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const fechaActivacion = props.getProperty('FECHA_ACTIVACION');
  const fecha = fechaActivacion ? new Date(fechaActivacion).toLocaleDateString('es-ES') : 'Desconocida';

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
        <script src="https://code.iconify.design/iconify-icon/1.0.7/iconify-icon.min.js"></script>
        <style>
          body { font-family: 'Inter', sans-serif; margin: 0; padding: 24px; background: #f9fafb; }
          .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 16px; padding: 28px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
          h3 { margin: 0 0 12px 0; color: #2563eb; font-size: 24px; display: flex; align-items: center; gap: 10px; }
          p { color: #6b7280; font-size: 15px; margin: 0 0 20px 0; }
          .url-box { background: #f3f4f6; padding: 14px; border-radius: 10px; border: 2px solid #e5e7eb; margin-bottom: 20px; word-break: break-all; font-family: monospace; font-size: 13px; color: #374151; cursor: pointer; line-height: 1.6; }
          .url-box:hover { background: #e5e7eb; }
          .button-group { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px; }
          .btn { padding: 14px 20px; border-radius: 10px; font-weight: 600; text-decoration: none; text-align: center; cursor: pointer; border: none; font-size: 15px; transition: all 0.2s; display: flex; align-items: center; justify-content: center; gap: 8px; }
          .btn-success { background: #10b981; color: white; }
          .btn-success:hover { background: #059669; }
          .btn-primary { background: #2563eb; color: white; }
          .btn-primary:hover { background: #1d4ed8; }
          .info { margin-top: 24px; padding: 16px; background: #f9fafb; border-radius: 10px; font-size: 14px; color: #6b7280; line-height: 1.6; }
          .info strong { color: #374151; }
        </style>
      </head>
      <body>
        <div class="container">
          <h3>
            <iconify-icon icon="mdi:link-variant" width="28"></iconify-icon>
            URL del Sistema de Reservas
          </h3>
          <p>Comparte este enlace con los usuarios:</p>
          
          <div class="url-box" onclick="copiarURL()" id="urlBox">
            ${url}
          </div>

          <div class="button-group">
            <button class="btn btn-success" onclick="copiarURL()" id="btnCopiar">
              <iconify-icon icon="mdi:content-copy" width="20"></iconify-icon>
              Copiar URL
            </button>
            
            <a href="${url}" target="_blank" class="btn btn-primary">
              <iconify-icon icon="mdi:open-in-new" width="20"></iconify-icon>
              Abrir Aplicación
            </a>
          </div>

          <div class="info">
            <strong>📅 Activado el:</strong> ${fecha}<br><br>
            <strong>💡 Nota:</strong> Si actualizas el código, recuerda crear una nueva versión desde el editor de Apps Script.
          </div>
        </div>

        <script>
          function copiarURL() {
            const url = "${url}";
            
            if (navigator.clipboard) {
              navigator.clipboard.writeText(url).then(() => {
                mostrarCopiado();
              }).catch(() => {
                copiarFallback(url);
              });
            } else {
              copiarFallback(url);
            }
          }
          
          function copiarFallback(url) {
            const textarea = document.createElement('textarea');
            textarea.value = url;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();
            
            try {
              document.execCommand('copy');
              mostrarCopiado();
            } catch(e) {
              alert('Por favor, copia manualmente la URL');
            }
            
            document.body.removeChild(textarea);
          }

          function mostrarCopiado() {
            const btn = document.getElementById('btnCopiar');
            const originalHTML = btn.innerHTML;
            
            btn.innerHTML = '<iconify-icon icon="mdi:check" width="20"></iconify-icon> ¡Copiado!';
            btn.style.background = '#059669';
            
            setTimeout(() => {
              btn.innerHTML = originalHTML;
              btn.style.background = '#10b981';
            }, 2000);
          }
        </script>
      </body>
    </html>
  `).setWidth(650).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, '🔗 URL de Acceso');
}

function cambiarURLManual() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const urlActual = props.getProperty('WEB_APP_URL') || 'No configurada';

  const prompt = ui.prompt(
    '⚙️ Cambiar URL',
    `URL actual:\n${urlActual}\n\nPega la nueva URL:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() == ui.Button.OK) {
    const url = prompt.getResponseText().trim();

    if (url.includes('script.google.com') && url.includes('/exec')) {
      props.setProperty('WEB_APP_URL', url);
      props.setProperty('FECHA_ACTIVACION', new Date().toISOString());
      ui.alert('✅ URL actualizada correctamente.');
    } else {
      ui.alert('❌ URL inválida. Debe ser una URL de Apps Script que termine en /exec');
    }
  }
}


/* ============================================
   UTILIDADES GENERALES
   ============================================ */

function sheetToObjects(sheet) {
  if (!sheet) {
    Logger.log("Error en sheetToObjects: Se recibió una hoja nula.");
    return [];
  }
  const numRows = sheet.getLastRow();
  if (numRows < 2) {
    Logger.log(`Aviso en sheetToObjects: La hoja '${sheet.getName()}' tiene pocos datos.`);
    return [];
  }
  const numCols = sheet.getLastColumn();
  const data = sheet.getRange(1, 1, numRows, numCols).getValues();
  const headers = data.shift().map(h => h.toString().trim().toLowerCase());

  const ss = getDB();
  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();

  return data.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      let value = row[index];
      if (value instanceof Date) {
        if (header === 'fecha') {
          obj[header] = Utilities.formatDate(value, spreadsheetTimezone, "yyyy-MM-dd");
        } else if (header === 'hora_inicio' || header === 'hora_fin') {
          obj[header] = Utilities.formatDate(value, spreadsheetTimezone, "HH:mm");
        } else {
          obj[header] = value.toISOString();
        }
      } else {
        obj[header] = value;
      }
    });
    return obj;
  });
}

/* ===========================================
   SISTEMA DE AUTORIZACIÓN (BLINDADO)
   =========================================== */
function checkUserAuthorization(emailUser) {
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.USUARIOS);

    // Si no existe la hoja, pánico (nadie entra)
    if (!sheet) {
      console.error("ERROR CRÍTICO: No existe hoja Usuarios");
      return { isAuthorized: false, isAdmin: false, error: "Falta hoja Usuarios" };
    }

    // Sin email de sesión no se puede identificar a nadie (evita que una fila vacía "coincida")
    if (!emailUser || !String(emailUser).trim()) {
      return { isAuthorized: false, isAdmin: false, error: "EMAIL_NO_DISPONIBLE" };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { isAuthorized: false, isAdmin: false }; // Hoja vacía

    const headers = data[0].map(h => h.toString().toLowerCase().trim());

    // 1. BUSCAR COLUMNA EMAIL (Flexible)
    // Buscamos 'email', 'email_usuario', 'correo', etc.
    let colEmail = headers.indexOf('email_usuario');
    if (colEmail === -1) colEmail = headers.indexOf('email'); // <--- AQUÍ ESTABA EL FALLO
    if (colEmail === -1) colEmail = headers.indexOf('correo');

    // 2. BUSCAR COLUMNA ADMIN (Flexible)
    let colAdmin = headers.indexOf('admin');
    if (colAdmin === -1) colAdmin = headers.indexOf('administrador');
    if (colAdmin === -1) colAdmin = headers.indexOf('rol'); // Por si acaso

    // 3. BUSCAR COLUMNA ACTIVO (Flexible)
    let colActivo = headers.indexOf('activo');

    // Si no encontramos la columna Email, no podemos validar nada
    if (colEmail === -1) {
      console.error("No se encuentra la columna Email en Usuarios");
      return { isAuthorized: false, isAdmin: false, error: "Columna Email no encontrada" };
    }

    // 4. BARRIDO DE USUARIOS
    // Normalizamos tu email para evitar problemas de mayúsculas/minúsculas
    const miEmail = emailUser.toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userEmail = String(row[colEmail]).toLowerCase().trim();

      if (userEmail === miEmail) {
        // ¡TE ENCONTRÉ!

        // Verificar si estás activo (si no existe columna Activo, asumimos que sí)
        const valActivo = String(row[colActivo]).toLowerCase().trim().replace('í', 'i');
        const isActive = colActivo === -1 ? true : (row[colActivo] === true || valActivo === 'true' || valActivo === 'si');

        if (!isActive) {
          return { isAuthorized: false, isAdmin: false, error: "Usuario inactivo" };
        }

        // Verificar si eres Admin
        // Aceptamos: TRUE, true, "Si", "Yes", "Admin"
        let isAdmin = false;
        if (colAdmin !== -1) {
          const valAdmin = String(row[colAdmin]).toLowerCase().trim().replace('í', 'i');
          isAdmin = (valAdmin === 'true' || valAdmin === 'si' || valAdmin === 'yes' || valAdmin === 'admin');
        }

        // Recuperar nombre (opcional)
        const colNombre = headers.indexOf('nombre') > -1 ? headers.indexOf('nombre') : headers.indexOf('nombre_completo');
        const userName = colNombre > -1 ? row[colNombre] : userEmail;

        return {
          isAuthorized: true,
          isAdmin: isAdmin,
          userName: userName
        };
      }
    }

    // Si llegamos aquí, es que tu email no está en la lista
    return { isAuthorized: false, isAdmin: false, error: "Usuario no registrado" };

  } catch (e) {
    console.error("Error auth: " + e);
    return { isAuthorized: false, isAdmin: false, error: e.toString() };
  }
}

/* ============================================
   GESTIÓN DE CACHÉ
   ============================================ */

const CACHE_KEY_STATIC = "STATIC_DATA_V6_FULL"; // Clave única
const CACHE_TIME = 21600; // 6 Horas

/**
 * Datos estáticos (recursos, tramos, usuarios, cursos, config) desde caché o desde las hojas.
 * Separado de getStaticData para que reservar / purgar caché no relean también las reservas.
 */
function getDatosEstaticos_() {
  // 2. CACHÉ DE DATOS ESTÁTICOS
  const cache = CacheService.getScriptCache();
  const cachedJSON = cache.get(CACHE_KEY_STATIC);

  let recursos, tramos, usuariosMap, cursos, modoVisualizacionCursos, configuracion;  // ✅ MODIFICADO

  if (cachedJSON) {
    Logger.log("✅ Datos estáticos desde CACHÉ V6");
    const staticData = JSON.parse(cachedJSON);
    recursos = staticData.recursos;
    tramos = staticData.tramos;
    usuariosMap = staticData.usuariosMap;
    cursos = staticData.cursos;
    modoVisualizacionCursos = staticData.modoVisualizacionCursos;
    configuracion = staticData.configuracion || {};  // ✅ AÑADIDO
  } else {
    Logger.log("🔄 Generando datos estáticos desde Excel...");
    const ss = getDB();

    // RECURSOS (solo activos)
    const sheetRecursos = ss.getSheetByName(SHEETS.RECURSOS);
    recursos = sheetToObjects(sheetRecursos)
      .filter(r => r.estado && r.estado.toLowerCase() === 'activo');

    // TRAMOS (con normalización de campos)
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    const tramosRaw = sheetToObjects(sheetTramos);

    tramos = tramosRaw.map(t => {
      const nombreCampo = t.nombre_tramo || t.nombretramo || t['nombre tramo'] ||
        t.nombre || t.tramo || Object.values(t)[1] || 'Tramo sin nombre';

      const horainicioCampo = t.hora_inicio || t.horainicio || t['hora inicio'] ||
        t.hora_ini || t.inicio || '';

      const horafinCampo = t.hora_fin || t.horafin || t['hora fin'] ||
        t.hora_final || t.fin || '';

      return {
        id_tramo: t.id_tramo || t.idtramo || t.id || Object.values(t)[0],
        nombre_tramo: nombreCampo,
        hora_inicio: horainicioCampo,
        hora_fin: horafinCampo,
        activo: t.activo !== undefined ? t.activo : true
      };
    });

    // USUARIOS → MAPA (email → nombre)
    const sheetUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
    const allUsuarios = sheetToObjects(sheetUsuarios);
    usuariosMap = {};
    allUsuarios.forEach(u => {
      if (u.email_usuario) {
        usuariosMap[u.email_usuario.toLowerCase()] = u.nombre_completo || u.email_usuario;
      }
    });

    // CURSOS + MODO VISUALIZACIÓN
    const sheetCursos = ss.getSheetByName(SHEETS.CURSOS);
    let cursosData = { cursos: [], modoVisualizacion: 'botones' };

    if (sheetCursos) {
      const modoViz = sheetCursos.getRange('D1').getValue();
      cursosData.modoVisualizacion = modoViz && modoViz.toString().toLowerCase() === 'listado' ? 'listado' : 'botones';

      const allCursos = sheetToObjects(sheetCursos);
      cursosData.cursos = allCursos
        .map(c => ({ etapa: c.etapa || '', curso: c.curso || '' }))
        .filter(c => c.etapa && c.curso);
    }

    cursos = cursosData.cursos;
    modoVisualizacionCursos = cursosData.modoVisualizacion;

    // ✅ CONFIGURACIÓN (NUEVO)
    Logger.log("📋 Cargando configuración del sistema...");
    const sheetConfig = ss.getSheetByName(SHEETS.CONFIG);
    configuracion = {};
    
    if (sheetConfig) {
      const configData = sheetToObjects(sheetConfig);
      configData.forEach(item => {
        const clave = item.clave;
        let valor = item.valor;
        
        if (!clave) return;
        
        configuracion[clave] = parsearValorConfig_(valor);
      });
      Logger.log(`⚙️ Configuración cargada: ${Object.keys(configuracion).length} parámetros`);
    }

    // GUARDAR EN CACHÉ V6
    const dataToCache = {
      recursos,
      tramos,
      usuariosMap,
      cursos,
      modoVisualizacionCursos,
      configuracion  // ✅ AÑADIDO
    };

    try {
      cache.put(CACHE_KEY_STATIC, JSON.stringify(dataToCache), CACHE_TIME);
      Logger.log(`💾 Cache V6 guardado: ${recursos.length} recursos, ${tramos.length} tramos`);
    } catch (e) {
      Logger.log("⚠️ Error guardando caché: " + e.message);
    }
  }

  return { recursos, tramos, usuariosMap, cursos, modoVisualizacionCursos, configuracion };
}

/**
 * FUNCIÓN ÚNICA - TODO EN UNO
 * Devuelve: Datos estáticos + Reservas + Info del usuario
 */
function getStaticData() {
  Logger.log("🚀 getStaticData() - Función Única Fusionada");

  try {
    // 1. AUTORIZACIÓN (necesaria para el frontend)
    const email = Session.getActiveUser().getEmail();
    const auth = checkUserAuthorization(email);
    // Los no registrados (pantalla de registro) no deben poder leer reservas ni emails de otros
    if (!auth || !auth.isAuthorized) {
      throw new Error("No tienes acceso al sistema.");
    }
    const isAdmin = auth ? auth.isAdmin : false;
    const userName = auth ? auth.userName : email;

    // 2. DATOS ESTÁTICOS (caché)
    const { recursos, tramos, usuariosMap, cursos, modoVisualizacionCursos, configuracion } = getDatosEstaticos_();

    // 3. RESERVAS FRESCAS (SIEMPRE desde Excel)
    // (antes se leía la hoja Reservas dos veces: una para todas y otra para las mías)
    const objsReservas = leerReservasObjetos_();
    const reservas = getActiveReservations_(objsReservas);

    // 4. MIS RESERVAS ACTIVAS
    const misReservasActivas = getMyActiveReservationsData_(email, objsReservas);

    // 4b. MOTIVOS DE RECURRENCIAS (para mostrar en "Mis Reservas")
    const misRecurrencias = {};
    const idsSolicitudes = [...new Set(
      misReservasActivas
        .filter(r => r.id_solicitud_recurrente)
        .map(r => String(r.id_solicitud_recurrente))
    )];

    if (idsSolicitudes.length > 0) {
      try {
        const sheetSol = getDB().getSheetByName('SolicitudesRecurrentes');
        if (sheetSol && sheetSol.getLastRow() > 1) {
          const dataSol = sheetSol.getDataRange().getValues();
          for (let i = 1; i < dataSol.length; i++) {
            const idSol = String(dataSol[i][0] || '').trim();
            if (idsSolicitudes.includes(idSol)) {
              misRecurrencias[idSol] = {
                motivo: String(dataSol[i][10] || '').trim(),
                dias_semana: String(dataSol[i][5] || '').trim()
              };
            }
          }
        }
      } catch (e) {
        Logger.log('⚠️ Error cargando motivos recurrencias: ' + e.message);
      }
    }

    // 5. RESPUESTA COMPLETA
    return {
      success: true,
      userEmail: email,
      userName: userName,
      isAdmin: isAdmin,
      recursos: recursos,
      tramos: tramos,
      cursos: cursos,
      modoVisualizacionCursos: modoVisualizacionCursos,
      usuariosMap: usuariosMap,
      reservas: reservas,
      misReservasActivas: misReservasActivas,
      misRecurrencias: misRecurrencias,
      configuracion: configuracion  // ✅ AÑADIDO
    };

  } catch (error) {
    Logger.log(`❌ Error en getStaticData: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}


/* ============================================
   GESTIÓN DE RESERVAS
   ============================================ */

// Añade filas al final de la hoja en una sola escritura.
// A diferencia de getRange(getLastRow()+1, ...), amplía la hoja si ya no quedan filas libres
// (appendRow lo hacía solo; getRange fuera de rango lanza error).
function anadirFilas_(sheet, filas) {
  if (!filas || filas.length === 0) return;
  const inicio = sheet.getLastRow() + 1;
  const faltan = inicio + filas.length - 1 - sheet.getMaxRows();
  if (faltan > 0) sheet.insertRowsAfter(sheet.getMaxRows(), faltan);
  sheet.getRange(inicio, 1, filas.length, filas[0].length).setValues(filas);
}

// Lee la hoja Reservas como objetos (una sola lectura reutilizable dentro de la misma petición)
function leerReservasObjetos_() {
  const sheetReservas = getDB().getSheetByName(SHEETS.RESERVAS);
  if (!sheetReservas || sheetReservas.getLastRow() < 2) return [];
  return sheetToObjects(sheetReservas);
}

// todasLasReservas (opcional): resultado de leerReservasObjetos_() para no volver a leer la hoja
function getActiveReservations_(todasLasReservas) {
  if (!todasLasReservas) todasLasReservas = leerReservasObjetos_();

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const reservasActivas = todasLasReservas.filter(r => {
    if (!r.fecha || !r.estado) return false;

    // Pequeña mejora de seguridad para fechas:
    // A veces Sheets devuelve un objeto Date, a veces String. Esto lo unifica:
    let fechaStr = r.fecha;
    if (r.fecha instanceof Date) {
      fechaStr = r.fecha.toISOString().split('T')[0]; // Convierte a YYYY-MM-DD
    }

    const fechaReserva = new Date(fechaStr + "T12:00:00Z");
    return r.estado.toLowerCase() === 'confirmada' && fechaReserva >= hoy;
  });

  return reservasActivas.map(r => ({
    id_reserva: r.id_reserva,
    id_recurso: r.id_recurso,
    email_usuario: r.email_usuario,
    fecha: r.fecha, // Puedes devolver el objeto Date original o formatearlo
    id_tramo: r.id_tramo,
    cantidad: parseInt(r.cantidad, 10) || 1,
    estado: r.estado,
    notas: r.notas || '',
    curso: r.curso || '',
    id_solicitud_recurrente: r.id_solicitud_recurrente || ''
  }));
}

/* ============================================
   OBTENER RESERVAS ACTIVAS DEL USUARIO (PARA CARGA INICIAL)
   ============================================ */
function getMyActiveReservationsData_(userEmail, todasLasReservas) {
  try {
    const allReservas = todasLasReservas || leerReservasObjetos_();

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    const myReservations = allReservas.filter(r => {
      if (!r.email_usuario || !r.fecha || !r.estado) return false;
      const fechaReserva = new Date(r.fecha + "T12:00:00Z");
      return r.email_usuario.toLowerCase() === userEmail.toLowerCase() &&
        r.estado.toLowerCase() === 'confirmada' &&
        fechaReserva >= hoy;
    });

    myReservations.sort((a, b) => {
      const dateA = new Date(a.fecha);
      const dateB = new Date(b.fecha);
      return dateA - dateB;
    });

    Logger.log(`✅ Cargadas ${myReservations.length} reservas para ${userEmail}`);
    return myReservations;

  } catch (error) {
    Logger.log(`Error en getMyActiveReservationsData: ${error.message}`);
    return [];
  }
}

/* ============================================
   ENDPOINTS PRINCIPALES (CON LOGS DE DIAGNÓSTICO 🕵️‍♂️)
   ============================================ */

/* ===========================================================================
   1. CONTROLADOR DE ACCESO (doGet)
   =========================================================================== */
function doGet(e) {

  // --- FASE 0: DETECCIÓN DE INSTALACIÓN ---
  // Esta comprobación debe ir PRIMERO. Si no está instalado, no permitimos nada más.
  const props = PropertiesService.getScriptProperties();

  // Usamos la misma clave que definimos en Setup.gs ('SETUP_COMPLETED')
  const isInstalled = props.getProperty('SETUP_COMPLETED');

  // Si NO está instalado (o es distinto de 'true'), lanzamos el Instalador
  if (isInstalled !== 'true') {
    return HtmlService.createTemplateFromFile('ActivacionSistema') // Asegúrate de que el HTML se llama 'ActivacionSistema'
      .evaluate()
      .setTitle('🚀 Setup del Sistema')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // --- FASE 1: ACCIONES DIRECTAS (Vienen de emails) ---
  // Solo llegamos aquí si el sistema YA está instalado.

  // 1.1. Cancelación de reservas
  if (e.parameter.action === "cancel" && e.parameter.id) {
    try {
      const reservaId = e.parameter.id;
      return handleEmailCancelation(reservaId, e.parameter.t);
    } catch (error) {
      Logger.log(error);
      return HtmlService.createHtmlOutput(
        `<h1>Error Grave</h1><p>Se ha producido un error al procesar la cancelación: ${error.message}</p>`
      );
    }
  }

  // 1.2. Aprobación de usuarios por Admin
  if (e.parameter.action === "approve" && e.parameter.email && e.parameter.nombre) {
    return handleAdminApproval(e.parameter.email, e.parameter.nombre);
  }

  // 1.3. Aprobación de solicitud recurrente desde email
  if (e.parameter.action === "aprobar_recurrente" && e.parameter.id) {
    return handleAprobarRecurrenteDesdeEmail(e.parameter.id);
  }

  // (LA ANTIGUA FASE 2 SE ELIMINA PORQUE ERA REDUNDANTE Y PROVOCABA ERROR)

  // --- FASE 2: AUTENTICACIÓN Y AUTORIZACIÓN ---

  // Obtener usuario actual
  const userEmail = Session.getActiveUser().getEmail();

  Logger.log("--- DEBUG LOGIN ---");
  Logger.log("Email detectado: " + userEmail);

  // Verificar permisos
  const authResult = checkUserAuthorization(userEmail);
  Logger.log("¿Está autorizado?: " + authResult.isAuthorized);

  // 🛑 Google no nos da el email del usuario -> registrarse no serviría de nada
  // (ocurre si la app se implementó con una cuenta de otro dominio, p.ej. @gmail.com,
  //  o si el usuario entra con una cuenta distinta a la del centro)
  if (!userEmail) {
    return HtmlService.createHtmlOutput(buildHtmlEmailNoDisponible_())
      .setTitle("No se pudo identificar tu cuenta")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 🛑 Si NO está autorizado -> Pantalla de Registro
  if (!authResult.isAuthorized) {
    Logger.log(">> Usuario no autorizado. Mostrando Registro.");

    const template = HtmlService.createTemplateFromFile('registro');
    template.email = userEmail;

    return template.evaluate()
      .setTitle("Solicitud de Registro")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // --- FASE 3: CARGAR APLICACIÓN PRINCIPAL ---

  const appConfig = getAppConfig();

  const template = HtmlService.createTemplateFromFile('index');

  // Variables para HTML
  template.userNameForHtml = authResult.userName || userEmail;
  template.userEmailForHtml = userEmail;
  template.appName = appConfig.appName || "Sistema de Reservas";
  template.logoUrl = appConfig.logoUrl || "";
  template.appVersion = APP_VERSION;

  // Variables para Javascript (JSON stringified)
  template.userNameForJs = JSON.stringify(authResult.userName || userEmail);
  template.userEmailForJs = JSON.stringify(userEmail);
  template.isAdminForJs = authResult.isAdmin ? 'true' : 'false';

  return template.evaluate()
    .setTitle(appConfig.appName || "Sistema de Reservas")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================
   CARGA DE DATOS INICIALES (TU LÓGICA EXACTA + VELOCIDAD CACHÉ 🚀)
   ============================================ */

// Función separada para Reservas (No cacheable)
function getReservasFrescas_() {
  return getActiveReservations_();
}


function cargarDisponibilidadRecurso(recursoId) {
  try {
    if (!recursoId) {
      throw new Error("ID de recurso no proporcionado");
    }

    const disponibilidad = getDisponibilidadRecurso(recursoId);
    Logger.log(`Cargada disponibilidad para ${recursoId}: ${disponibilidad.length} registros`);

    // Incluir reservas frescas para que el cliente siempre tenga datos actualizados
    const reservas = getReservasFrescas_();

    return {
      success: true,
      disponibilidad: disponibilidad,
      reservas: reservas
    };
  } catch (e) {
    Logger.log(`Error en cargarDisponibilidadRecurso: ${e.message}`);
    return {
      success: false,
      error: e.message
    };
  }
}

/* ============================================
   GESTIÓN DE DISPONIBILIDAD Y VALIDACIÓN (CORREGIDO ✅)
   ============================================ */

function getDisponibilidadRecurso(recursoId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_KEYS.DISPONIBILIDAD + recursoId;
  const cached = cache.get(cacheKey);

  if (cached != null) {
    return JSON.parse(cached);
  }

  const ss = getDB();
  const sheet = ss.getSheetByName(SHEETS.DISPONIBILIDAD);

  // 🛑 CORRECCIÓN DE SEGURIDAD:
  // Si la hoja está vacía (solo cabecera), devolvemos array vacío y guardamos en caché
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    const emptyResult = [];
    cache.put(cacheKey, JSON.stringify(emptyResult), CACHE_TIMES.DISPONIBILIDAD);
    return emptyResult;
  }

  // Ahora es seguro leer porque sabemos que hay datos
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // ID del recurso limpio para comparar
  const targetId = String(recursoId).trim();

  const disponibilidadRecurso = data
    .filter(row => String(row[0]).trim() === targetId)
    .map(row => {
      let diaRaw = String(row[1]).trim();
      let diaNormalizado = diaRaw;

      if (diaRaw.match(/^Lunes/i)) diaNormalizado = "1";
      else if (diaRaw.match(/^Martes/i)) diaNormalizado = "2";
      else if (diaRaw.match(/^Mi[eé]rcoles/i)) diaNormalizado = "3";
      else if (diaRaw.match(/^Jueves/i)) diaNormalizado = "4";
      else if (diaRaw.match(/^Viernes/i)) diaNormalizado = "5";
      else if (diaRaw.match(/^S[aá]bado/i)) diaNormalizado = "6";
      else if (diaRaw.match(/^Domingo/i)) diaNormalizado = "7";

      const permitidoRaw = String(row[4]).toLowerCase().trim();
      const esSi = (permitidoRaw === 'si' || permitidoRaw === 'sí' || permitidoRaw === 'true' || permitidoRaw === '1' || permitidoRaw === 'yes');

      return {
        id_recurso: String(row[0]).trim(),
        dia_semana: diaNormalizado,
        id_tramo: String(row[2]).trim(),
        permitido: esSi ? 'si' : 'no',
        razon_bloqueo: String(row[5])
      };
    });

  cache.put(cacheKey, JSON.stringify(disponibilidadRecurso), CACHE_TIMES.DISPONIBILIDAD);
  return disponibilidadRecurso;
}

/* ============================================
   VALIDACIÓN DE DISPONIBILIDAD (LÓGICA PERMISIVA ✅)
   ============================================ */
function checkAvailability(recursoId, fechaISO, tramoId, cantidadPedida, recurso = null, staticData = null, reservasActivasPrevias = null) {

  // ✅ Solo cargar si no se pasaron como parámetro
  if (!staticData) {
    staticData = getDatosEstaticos_();
  }

  if (!recurso) {
    recurso = staticData.recursos.find(r => String(r.id_recurso) === String(recursoId));
    if (!recurso) throw new Error("El recurso no existe.");
  }

  // 1. Validar que el Recurso está activo
  if (recurso.estado && recurso.estado.toLowerCase() !== 'activo') {
    throw new Error("El recurso no está activo.");
  }

  // 2. Validar Bloqueos Administrativos (Disponibilidad)
  const disponibilidadRecurso = getDisponibilidadRecurso(recursoId);

  const fecha = new Date(fechaISO + "T12:00:00Z");
  let diaSemanaJS = fecha.getDay();
  let diaSemanaGoogle = (diaSemanaJS === 0) ? 7 : diaSemanaJS;
  const diaSemanaStr = diaSemanaGoogle.toString();

  const bloqueoEncontrado = disponibilidadRecurso.find(d => {
    const diaDisp = String(d.dia_semana).trim();
    const diaCoincide = diaDisp === diaSemanaStr || diaDisp.startsWith(diaSemanaStr + " ");
    return diaCoincide && String(d.id_tramo) === String(tramoId) && d.permitido === 'no';
  });

  if (bloqueoEncontrado) {
    const razon = bloqueoEncontrado.razon_bloqueo || "Tramo bloqueado por administración.";
    throw new Error(razon);
  }

  // 3. Validar Ocupación
  const reservasActivas = reservasActivasPrevias || getActiveReservations_();

  let cantidadReservada = 0;

  reservasActivas.forEach(r => {
    let fechaR = r.fecha;
    if (fechaR instanceof Date) {
      fechaR = Utilities.formatDate(fechaR, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    if (String(r.id_recurso) === String(recursoId) &&
      fechaR === fechaISO &&
      String(r.id_tramo) === String(tramoId)) {
      cantidadReservada += (parseInt(r.cantidad, 10) || 1);
    }
  });

  const esAgrupado = recurso.tipo.toLowerCase() === 'agrupado';
  const capacidadTotal = parseInt(recurso.capacidad, 10) || 1;

  if (esAgrupado) {
    const disponibles = capacidadTotal - cantidadReservada;
    if (cantidadPedida > disponibles) {
      throw new Error(`Completo: Solo quedan ${disponibles} unidades.`);
    }
  } else {
    if (cantidadReservada > 0) {
      throw new Error("La sala ya está reservada en ese horario.");
    }
  }

  return true;
}

function crearNuevaReserva(reservaData) {
  return crearReservas_(reservaData, [reservaData && reservaData.tramoId]);
}

/**
 * MULTITRAMO: reserva varios tramos SEGUIDOS del mismo recurso y día en una sola operación.
 * Requiere Config: permitir_multitramo = TRUE y max_tramos_simultaneos >= nº de tramos.
 * Es "todo o nada": si un tramo falla, no se crea ninguna reserva.
 */
function crearReservasMultitramo(reservaData, tramoIds) {
  try {
    if (!Array.isArray(tramoIds) || tramoIds.length === 0) {
      throw new Error("No se han indicado tramos.");
    }
    const ids = tramoIds.map(t => String(t).trim());
    if (new Set(ids).size !== ids.length) throw new Error("Hay tramos repetidos.");

    if (ids.length > 1) {
      if (getConfigValue('permitir_multitramo', false) !== true) {
        throw new Error("La reserva de varios tramos seguidos no está activada.");
      }
      const maxTramos = parseInt(getConfigValue('max_tramos_simultaneos', 1), 10) || 1;
      if (ids.length > maxTramos) {
        throw new Error(`Solo puedes reservar hasta ${maxTramos} tramos seguidos.`);
      }
      // Deben ser consecutivos según el orden de la hoja Tramos
      const orden = getDatosEstaticos_().tramos.map(t => String(t.id_tramo).trim());
      const posiciones = ids.map(id => orden.indexOf(id));
      if (posiciones.some(p => p === -1)) throw new Error("Algún tramo no existe.");
      posiciones.sort((a, b) => a - b);
      for (let k = 1; k < posiciones.length; k++) {
        if (posiciones[k] !== posiciones[k - 1] + 1) throw new Error("Los tramos deben ser seguidos.");
      }
      // Guardamos en el orden del horario
      ids.splice(0, ids.length, ...posiciones.map(p => orden[p]));
    }
    return crearReservas_(reservaData, ids);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Núcleo común de reserva (1 o varios tramos). Todas las validaciones se hacen
 * antes de escribir nada y con el lock cogido; las filas se escriben de una vez.
 */
function crearReservas_(reservaData, tramoIds) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: "El sistema está ocupado. Por favor, intenta de nuevo en unos segundos." };
  }

  try {
    const email = Session.getActiveUser().getEmail();
    const authResult = checkUserAuthorization(email);
    if (!authResult.isAuthorized) {
      throw new Error("Tu sesión ha caducado o ya no tienes permisos.");
    }

    const { recursoId, recursoNombre, fechaISO, tramoNombre, notas, curso } = reservaData;
    tramoIds = (tramoIds || []).map(t => String(t || '').trim()).filter(Boolean);
    if (tramoIds.length === 0) throw new Error("Debes seleccionar un tramo.");

    // Cantidad: entero >= 1 (una cantidad negativa "liberaba" aforo en recursos agrupados)
    const cantidad = parseInt(reservaData.cantidad, 10) || 1;
    if (cantidad < 1) {
      throw new Error("La cantidad debe ser al menos 1.");
    }

    if (!curso || curso.trim() === '') {
      throw new Error("Debes seleccionar un curso para realizar la reserva.");
    }

    // exigir_motivo (Config): las notas pasan a ser obligatorias
    if (getConfigValue('exigir_motivo', false) === true && !String(notas || '').trim()) {
      throw new Error("Indica en las notas para qué es la reserva.");
    }

    // ✅ VALIDACIONES DE CONFIGURACIÓN (una sola lectura de Reservas para todas)
    const reservasActivas = getActiveReservations_();
    validarModoMantenimiento();
    validarDiasVista(fechaISO);
    tramoIds.forEach(t => validarAntelacionMinima(fechaISO, t));
    validarLimiteReservas(email, reservasActivas, tramoIds.length); // cada tramo cuenta como una reserva

    const staticData = getDatosEstaticos_();
    const recurso = staticData.recursos.find(r => String(r.id_recurso) === String(recursoId));

    if (!recurso) {
      throw new Error("El recurso seleccionado no existe.");
    }

    tramoIds.forEach(t => {
      try {
        checkAvailability(recursoId, fechaISO, t, cantidad, recurso, staticData, reservasActivas);
      } catch (errDisp) {
        if (tramoIds.length === 1) throw errDisp;
        const tr = staticData.tramos.find(x => String(x.id_tramo).trim() === t);
        throw new Error(`${tr ? tr.nombre_tramo : t}: ${errDisp.message}`);
      }
    });

    const timestamp = new Date();
    const fechaReserva = new Date(fechaISO + "T12:00:00Z");

    const sheetReservas = getDB().getSheetByName(SHEETS.RESERVAS);
    const numCols = sheetReservas.getLastColumn();
    const headers = sheetReservas.getRange(1, 1, 1, numCols).getValues()[0];

    const headerMap = {};
    headers.forEach((h, i) => {
      headerMap[h.toString().trim().toLowerCase()] = i;
    });

    const filas = [];
    const nuevasReservas = [];
    const tramosTexto = [];

    tramoIds.forEach(tramoId => {
      const idReserva = Utilities.getUuid();
      const fila = new Array(numCols).fill("");
      if (headerMap['id_reserva'] !== undefined) fila[headerMap['id_reserva']] = idReserva;
      if (headerMap['id_recurso'] !== undefined) fila[headerMap['id_recurso']] = recursoId;
      if (headerMap['email_usuario'] !== undefined) fila[headerMap['email_usuario']] = email;
      if (headerMap['fecha'] !== undefined) fila[headerMap['fecha']] = fechaReserva;
      if (headerMap['id_tramo'] !== undefined) fila[headerMap['id_tramo']] = tramoId;
      if (headerMap['cantidad'] !== undefined) fila[headerMap['cantidad']] = cantidad;
      if (headerMap['estado'] !== undefined) fila[headerMap['estado']] = "Confirmada";
      if (headerMap['notas'] !== undefined) fila[headerMap['notas']] = notas;
      if (headerMap['curso'] !== undefined) fila[headerMap['curso']] = curso;
      if (headerMap['timestamp'] !== undefined) fila[headerMap['timestamp']] = timestamp;
      filas.push(fila);

      const tramo = staticData.tramos.find(t => String(t.id_tramo).trim() === tramoId) || null;
      let texto = tramoIds.length === 1 ? (tramoNombre || tramoId) : tramoId;
      if (tramo) {
        const nombreTramo = tramo.nombre_tramo || texto;
        texto = (tramo.hora_inicio && tramo.hora_fin)
          ? `${nombreTramo} (${tramo.hora_inicio} - ${tramo.hora_fin})`
          : nombreTramo;
      }
      tramosTexto.push(texto);

      nuevasReservas.push({
        id_reserva: idReserva,
        id_recurso: recursoId,
        email_usuario: email,
        fecha: fechaISO,
        id_tramo: tramoId,
        cantidad: cantidad,
        estado: 'Confirmada',
        notas: notas,
        curso: curso,
        timestamp_creacion: timestamp.toISOString()
      });
    });

    // Una sola escritura para todas las filas
    anadirFilas_(sheetReservas, filas);
    Logger.log(`[crearReservas_] ${filas.length} reserva(s) creada(s): ${nuevasReservas.map(r => r.id_reserva).join(', ')}`);

    const fechaFormateada = fechaReserva.toLocaleDateString('es-ES', {
      day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC'
    });

    // ✅ Limpiar caché de disponibilidad del recurso.
    // (La caché estática NO contiene reservas: borrarla solo obligaba a releer todas las hojas.)
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEYS.DISPONIBILIDAD + recursoId);

    // Las filas ya están escritas: liberamos el lock antes de enviar el email (lento)
    lock.releaseLock();

    sendConfirmationEmail_(email, authResult.userName, {
      idReserva: nuevasReservas[0].id_reserva,
      reservas: nuevasReservas.map((r, k) => ({ idReserva: r.id_reserva, tramo: tramosTexto[k] })),
      recursoNombre: recursoNombre,
      fechaFormateada: fechaFormateada,
      tramoNombre: tramosTexto.join(', '),
      curso: curso,
      cantidad: cantidad,
      notas: notas
    });

    return {
      success: true,
      message: nuevasReservas.length > 1
        ? `¡${nuevasReservas.length} tramos reservados! Se ha enviado un correo de confirmación.`
        : "¡Reserva confirmada! Se ha enviado un correo de confirmación.",
      nuevaReserva: nuevasReservas[0],
      nuevasReservas: nuevasReservas
    };

  } catch (error) {
    lock.releaseLock();
    Logger.log(`❌ Error en crearReservas_: ${error.message}`);
    return { success: false, message: error.message };  // ✅ Sin el prefijo "Error al crear la reserva:"
  }
}

/* ============================================
   EMAIL DE CONFIRMACIÓN DE RESERVA
   ============================================ */
// Escapa texto de usuario antes de insertarlo en HTML (emails y páginas)
function escHtml_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function sendConfirmationEmail_(email, userName, details) {
  try {
    // Generar URL(s) de cancelación (dentro del try-catch para que un fallo aquí no impida el email)
    // Multitramo: un enlace por tramo, para poder cancelar solo uno de ellos
    const reservasEmail = (details.reservas && details.reservas.length) ? details.reservas
      : [{ idReserva: details.idReserva, tramo: details.tramoNombre }];
    let urlApp = '';
    try {
      urlApp = ScriptApp.getService().getUrl();
    } catch (urlError) {
      Logger.log(`⚠️ No se pudo obtener URL de la app: ${urlError.message}`);
    }
    const urlCancelar = id => {
      try { return `${urlApp}?action=cancel&id=${id}&t=${firmarIdReserva_(id)}`; }
      catch (e) { Logger.log('⚠️ No se pudo firmar el enlace: ' + e.message); return ''; }
    };
    const estiloBoton = 'padding: 10px 15px; background-color: #d9534f; color: white; text-decoration: none; border-radius: 5px; display: inline-block; margin: 4px 0;';

    const asunto = `Reserva Confirmada: ${details.recursoNombre} - ${details.fechaFormateada}`;

    // Bloque de cancelación: solo si tenemos URL
    let bloqueCancelacion = '';
    if (urlApp && reservasEmail.length === 1 && urlCancelar(reservasEmail[0].idReserva)) {
      bloqueCancelacion = `<p>Si necesitas cancelar la reserva, puedes hacerlo desde este enlace:</p>
         <p><a href="${urlCancelar(reservasEmail[0].idReserva)}" style="${estiloBoton}">Cancelar esta Reserva</a></p>`;
    } else if (urlApp) {
      bloqueCancelacion = `<p>Si necesitas cancelar algún tramo, puedes hacerlo desde estos enlaces:</p>` +
        reservasEmail.filter(r => urlCancelar(r.idReserva))
          .map(r => `<p><a href="${urlCancelar(r.idReserva)}" style="${estiloBoton}">Cancelar ${escHtml_(r.tramo)}</a></p>`).join('');
    }

    const cuerpoHtml = `
      <p>¡Hola ${userName || ''}!</p>
      <p>Tu reserva ha sido confirmada con éxito.</p>
      <hr>
      <ul>
        <li><strong>Recurso:</strong> ${escHtml_(details.recursoNombre)}</li>
        <li><strong>Fecha:</strong> ${details.fechaFormateada}</li>
        <li><strong>${reservasEmail.length > 1 ? 'Tramos' : 'Tramo'}:</strong> ${escHtml_(details.tramoNombre)}</li>
        <li><strong>Curso:</strong> ${escHtml_(details.curso)}</li>
        ${details.cantidad > 1 ? `<li><strong>Cantidad:</strong> ${details.cantidad}</li>` : ''}
        ${details.notas ? `<li><strong>Notas:</strong> ${escHtml_(details.notas)}</li>` : ''}
      </ul>
      <hr>
      ${bloqueCancelacion}
      <p style="font-size: 0.8em; color: #777;">Puedes gestionar tus reservas desde la propia aplicación.</p>
    `;

    const emailOptions = {
      to: email,
      subject: asunto,
      htmlBody: cuerpoHtml
    };

    // Añadir BCC solo si el admin lo tiene activado y hay email configurado
    const adminEmail = getConfigValue('email_admin', '');
    const recibirCopia = getConfigValue('admin_recibir_copia_reservas', false);
    if (adminEmail && recibirCopia === true) {
      emailOptions.bcc = adminEmail;
    }

    MailApp.sendEmail(emailOptions);
    Logger.log(`✅ Correo de confirmación enviado a ${email} para reserva ${details.idReserva}`);
  } catch (e) {
    Logger.log(`❌ Error al enviar correo de confirmación a ${email}: ${e.message}\nStack: ${e.stack}`);
  }
}

/* ============================================
   GESTIÓN DE RESERVAS DEL USUARIO
   ============================================ */

function cancelarReservaCliente(reservaId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: "El sistema está ocupado. Por favor, intenta de nuevo." };
  }

  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const sheetReservas = getDB().getSheetByName(SHEETS.RESERVAS);
    if (!sheetReservas) throw new Error("No se encuentra la hoja de Reservas.");

    const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const COL_ID_INDEX = headers.indexOf("ID_Reserva");
    const COL_ESTADO_INDEX = headers.indexOf("Estado");
    const COL_EMAIL_INDEX = headers.indexOf("Email_Usuario");
    const COL_RECURSO_ID_INDEX = headers.indexOf("ID_Recurso");
    const COL_TRAMO_ID_INDEX = headers.indexOf("ID_Tramo");
    const COL_FECHA_INDEX = headers.indexOf("Fecha");

    if (COL_ID_INDEX === -1 || COL_ESTADO_INDEX === -1 || COL_EMAIL_INDEX === -1) {
      throw new Error("La hoja de 'Reservas' está mal configurada.");
    }

    const idColumnRange = sheetReservas.getRange(2, COL_ID_INDEX + 1, sheetReservas.getLastRow() - 1);
    const textFinder = idColumnRange.createTextFinder(reservaId).matchEntireCell(true);
    const celdaEncontrada = textFinder.findNext();

    if (!celdaEncontrada) {
      throw new Error("No se ha encontrado la reserva.");
    }

    const filaEncontrada = celdaEncontrada.getRow();
    const rowData = sheetReservas.getRange(filaEncontrada, 1, 1, headers.length).getValues()[0];
    const emailEnReserva = rowData[COL_EMAIL_INDEX].toString().toLowerCase();
    const estadoActual = rowData[COL_ESTADO_INDEX];

    if (emailEnReserva !== userEmail) {
      Logger.log(`Intento de cancelación ILEGÍTIMO: ${userEmail} intentó cancelar reserva de ${emailEnReserva}`);
      throw new Error("No tienes permiso para cancelar esta reserva.");
    }

    if (estadoActual === "Cancelada") {
      lock.releaseLock();
      return { success: true, message: "Esta reserva ya había sido cancelada.", canceledId: reservaId };
    }

    // Tope de cancelación (horas_cancelacion); los admin quedan exentos
    if (COL_FECHA_INDEX !== -1 && COL_TRAMO_ID_INDEX !== -1 && !checkUserAuthorization(userEmail).isAdmin) {
      const bloqueo = motivoBloqueoCancelacion_(rowData[COL_FECHA_INDEX], rowData[COL_TRAMO_ID_INDEX]);
      if (bloqueo) throw new Error(bloqueo);
    }

    sheetReservas.getRange(filaEncontrada, COL_ESTADO_INDEX + 1).setValue("Cancelada");

    const authResult = checkUserAuthorization(userEmail);

    const recursoId = rowData[COL_RECURSO_ID_INDEX];
    const tramoId = rowData[COL_TRAMO_ID_INDEX];
    const fecha = new Date(rowData[COL_FECHA_INDEX]);

    // ✅ PASO 1B: Leer solo del caché (sin cargar getStaticData completo)
    let recursoNombre = 'Recurso no disponible';
    let tramoNombre = 'Horario no disponible';

    const cache = CacheService.getScriptCache();
    const cachedStatic = cache.get(CACHE_KEY_STATIC);

    if (cachedStatic) {
      try {
        const staticData = JSON.parse(cachedStatic);

        // Buscar recurso
        const recurso = staticData.recursos.find(r =>
          String(r.id_recurso).trim() === String(recursoId).trim()
        );
        if (recurso) {
          recursoNombre = recurso.nombre;
        }

        // Buscar tramo
        const tramo = staticData.tramos.find(t =>
          String(t.id_tramo).trim() === String(tramoId).trim()
        );
        if (tramo) {
          const nombreTramo = tramo.nombre_tramo || tramo.nombretramo || tramo.nombre || 'Tramo';
          const horaInicio = tramo.hora_inicio || tramo.horainicio || '';
          const horaFin = tramo.hora_fin || tramo.horafin || '';

          if (horaInicio && horaFin) {
            tramoNombre = `${nombreTramo} (${horaInicio} - ${horaFin})`;
          } else {
            tramoNombre = nombreTramo;
          }
        }
      } catch (e) {
        Logger.log(`Error parseando caché en cancelación: ${e.message}`);
      }
    }

    const fechaFormateada = fecha.toLocaleDateString('es-ES', {
      day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC'
    });

    const details = {
      recursoNombre: recursoNombre,
      fechaFormateada: fechaFormateada,
      tramoNombre: tramoNombre
    };

    sendCancelationEmail_(userEmail, authResult.userName, details);

    cache.remove(CACHE_KEYS.DISPONIBILIDAD + recursoId);

    lock.releaseLock();
    return { success: true, message: "Reserva cancelada con éxito.", canceledId: reservaId };

  } catch (error) {
    lock.releaseLock();
    Logger.log(`Error en cancelarReservaCliente: ${error.message}`);
    return { success: false, message: error.message };
  }
}

/* ============================================
   CANCELACIÓN VÍA EMAIL
   ============================================ */

/* Firma HMAC para los enlaces de cancelación de los emails.
   Así el enlace funciona aunque Google no identifique al usuario (otras cuentas/dominios)
   y nadie puede cancelar reservas ajenas adivinando o copiando un ID. */
function getSecretoEnlaces_() {
  const props = PropertiesService.getScriptProperties();
  let secreto = props.getProperty('SECRETO_ENLACES');
  if (!secreto) {
    secreto = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SECRETO_ENLACES', secreto);
  }
  return secreto;
}

function firmarIdReserva_(idReserva) {
  const firma = Utilities.computeHmacSha256Signature(String(idReserva), getSecretoEnlaces_());
  return Utilities.base64EncodeWebSafe(firma).replace(/=+$/, '').substring(0, 24);
}

function handleEmailCancelation(reservaId, firma) {
  const sheetReservas = getDB().getSheetByName(SHEETS.RESERVAS);
  if (!sheetReservas) {
    return HtmlService.createHtmlOutput(buildHtmlCancelPage("Error", "No se encuentra la hoja de Reservas."));
  }

  const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
  const COL_ID_INDEX = headers.indexOf("ID_Reserva");
  const COL_ESTADO_INDEX = headers.indexOf("Estado");
  const COL_EMAIL_INDEX = headers.indexOf("Email_Usuario");
  const COL_RECURSO_ID_INDEX = headers.indexOf("ID_Recurso");
  const COL_TRAMO_ID_INDEX = headers.indexOf("ID_Tramo");
  const COL_FECHA_INDEX = headers.indexOf("Fecha");

  if (COL_ID_INDEX === -1 || COL_ESTADO_INDEX === -1 || COL_EMAIL_INDEX === -1) {
    return HtmlService.createHtmlOutput(buildHtmlCancelPage("Error", "La hoja de 'Reservas' está mal configurada."));
  }

  const idColumnRange = sheetReservas.getRange(2, COL_ID_INDEX + 1, sheetReservas.getLastRow() - 1);
  const textFinder = idColumnRange.createTextFinder(reservaId).matchEntireCell(true);
  const celdaEncontrada = textFinder.findNext();

  if (!celdaEncontrada) {
    return HtmlService.createHtmlOutput(buildHtmlCancelPage("Error", "No se ha encontrado la reserva."));
  }

  const filaEncontrada = celdaEncontrada.getRow();
  const rowData = sheetReservas.getRange(filaEncontrada, 1, 1, headers.length).getValues()[0];
  const estadoActual = rowData[COL_ESTADO_INDEX];

  // 🛡️ Autorización: enlace firmado (emails nuevos) o bien el propio dueño / un admin (emails antiguos)
  const firmaValida = firma && firma === firmarIdReserva_(reservaId);
  if (!firmaValida) {
    const emailActual = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
    const emailDueno = String(rowData[COL_EMAIL_INDEX] || '').toLowerCase().trim();
    const esDueno = emailActual && emailActual === emailDueno;
    if (!esDueno && !(emailActual && checkUserAuthorization(emailActual).isAdmin)) {
      return HtmlService.createHtmlOutput(buildHtmlCancelPage("Acceso denegado", "Solo quien hizo la reserva o un administrador puede cancelarla. Hazlo desde la aplicación."));
    }
  }

  // ✅ OBTENER DATOS
  const recursoId = rowData[COL_RECURSO_ID_INDEX];
  const tramoId = rowData[COL_TRAMO_ID_INDEX];
  const fecha = new Date(rowData[COL_FECHA_INDEX]);

  // ✅ PASO 1B: Leer solo del caché (sin cargar getStaticData completo)
  let recursoNombre = 'Recurso no disponible';
  let tramoNombre = 'Horario no disponible';

  const cache = CacheService.getScriptCache();
  const cachedStatic = cache.get(CACHE_KEY_STATIC);

  if (cachedStatic) {
    try {
      const staticData = JSON.parse(cachedStatic);

      // Buscar recurso
      const recurso = staticData.recursos.find(r =>
        String(r.id_recurso).trim() === String(recursoId).trim()
      );
      if (recurso) {
        recursoNombre = recurso.nombre;
      }

      // Buscar tramo
      const tramo = staticData.tramos.find(t =>
        String(t.id_tramo).trim() === String(tramoId).trim()
      );
      if (tramo) {
        const nombreTramo = tramo.nombre_tramo || tramo.nombretramo || tramo.nombre || 'Tramo';
        const horaInicio = tramo.hora_inicio || tramo.horainicio || '';
        const horaFin = tramo.hora_fin || tramo.horafin || '';

        if (horaInicio && horaFin) {
          tramoNombre = `${nombreTramo} (${horaInicio} - ${horaFin})`;
        } else {
          tramoNombre = nombreTramo;
        }
      }
    } catch (e) {
      Logger.log(`Error parseando caché en cancelación por email: ${e.message}`);
      // Valores por defecto ya asignados arriba
    }
  }

  const fechaFormateada = fecha.toLocaleDateString('es-ES', {
    day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC'
  });

  const details = {
    recursoNombre: recursoNombre,
    fechaFormateada: fechaFormateada,
    tramoNombre: tramoNombre
  };

  // ✅ SI YA ESTABA CANCELADA
  if (estadoActual === "Cancelada") {
    return HtmlService.createHtmlOutput(
      buildHtmlCancelPage("Aviso", "Esta reserva ya había sido cancelada.", details)
    );
  }

  // Tope de cancelación (horas_cancelacion); exentos solo si quien pulsa es admin identificado
  if (COL_FECHA_INDEX !== -1 && COL_TRAMO_ID_INDEX !== -1) {
    const emailClic = String(Session.getActiveUser().getEmail() || '').trim();
    if (!(emailClic && checkUserAuthorization(emailClic).isAdmin)) {
      const bloqueo = motivoBloqueoCancelacion_(rowData[COL_FECHA_INDEX], rowData[COL_TRAMO_ID_INDEX]);
      if (bloqueo) {
        return HtmlService.createHtmlOutput(buildHtmlCancelPage("No se puede cancelar", bloqueo, details));
      }
    }
  }

  // ✅ CANCELAR LA RESERVA
  sheetReservas.getRange(filaEncontrada, COL_ESTADO_INDEX + 1).setValue("Cancelada");

  const email = rowData[COL_EMAIL_INDEX];
  const authResult = checkUserAuthorization(email);
  sendCancelationEmail_(email, authResult.userName, details);

  // ✅ Reutilizar la variable cache que ya existe arriba
  cache.remove(CACHE_KEYS.DISPONIBILIDAD + recursoId);

  return HtmlService.createHtmlOutput(
    buildHtmlCancelPage("Reserva Cancelada", "Tu reserva ha sido cancelada con éxito.", details)
  );
}

/* ============================================
   UTILIDADES DE EMAIL Y HTML
   ============================================ */

function sendCancelationEmail_(email, userName, details) {
  const asunto = `Cancelación Confirmada: ${details.recursoNombre} - ${details.fechaFormateada}`;
  const cuerpoHtml = `
    <p>¡Hola ${userName || ''}!</p>
    <p>Tu reserva ha sido <b>cancelada</b> con éxito.</p>
    <hr>
    <p>Detalles de la reserva cancelada:</p>
    <ul>
      <li><strong>Recurso:</strong> ${escHtml_(details.recursoNombre)}</li>
      <li><strong>Fecha:</strong> ${details.fechaFormateada}</li>
      <li><strong>Tramo:</strong> ${escHtml_(details.tramoNombre)}</li>
    </ul>
    <hr>
    <p style="font-size: 0.9em; color: #777;">Si has cancelado por error, deberás volver a realizar la reserva desde la aplicación.</p>
  `;

  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpoHtml
    });
    Logger.log(`Correo de cancelación enviado a ${email}.`);
  } catch (e) {
    Logger.log(`Error al enviar correo de cancelación a ${email}: ${e.message}`);
  }
}

function buildHtmlCancelPage(titulo, mensaje, details) {
  let detailsHtml = '';
  if (details) {
    detailsHtml = `
      <hr style="margin: 20px 0; border: none; border-top: 1px solid #e5e7eb;">
      <ul style="text-align: left; list-style: none; padding-left: 0; margin: 0;">
        <li style="margin-bottom: 8px;"><strong>Recurso:</strong> ${escHtml_(details.recursoNombre)}</li>
        <li style="margin-bottom: 8px;"><strong>Fecha:</strong> ${details.fechaFormateada}</li>
        <li style="margin-bottom: 8px;"><strong>Tramo:</strong> ${escHtml_(details.tramoNombre)}</li>
      </ul>
      <hr style="margin: 20px 0; border: none; border-top: 1px solid #e5e7eb;">
    `;
  }

  return `
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
      <title>${titulo}</title>
      <style>
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        
        body { 
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
          display: flex;
          align-items: center;
          justify-content: center;
          min-height: 100vh;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          padding: 16px;
        }
        
        .container {
          width: 100%;
          max-width: 500px;
          padding: 2rem;
          background: white;
          border-radius: 16px;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
          text-align: center;
        }
        
        h1 { 
          color: #1f2937;
          font-size: 1.75rem;
          margin-bottom: 1rem;
          font-weight: 700;
        }
        
        p { 
          color: #4b5563;
          line-height: 1.6;
          font-size: 1rem;
          margin-bottom: 1rem;
        }
        
        hr {
          margin: 20px 0;
          border: none;
          border-top: 1px solid #e5e7eb;
        }
        
        ul {
          text-align: left;
          list-style: none;
          padding-left: 0;
          margin: 0;
        }
        
        li {
          margin-bottom: 8px;
          font-size: 0.95rem;
          color: #374151;
        }
        
        strong {
          color: #1f2937;
          font-weight: 600;
        }
        
        .emoji {
          font-size: 3rem;
          margin-bottom: 1rem;
          display: block;
        }
        
        @media (max-width: 480px) {
          .container {
            padding: 1.5rem;
          }
          
          h1 {
            font-size: 1.5rem;
          }
          
          p {
            font-size: 0.95rem;
          }
          
          li {
            font-size: 0.9rem;
          }
          
          .emoji {
            font-size: 2.5rem;
          }
        }
        
        @media (prefers-color-scheme: dark) {
          body {
            background: linear-gradient(135deg, #1e3a8a 0%, #7c3aed 100%);
          }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <span class="emoji">✅</span>
        <h1>${titulo}</h1>
        <p>${mensaje}</p>
        ${detailsHtml}
      </div>
    </body>
    </html>
  `;
}

/* ============================================
   APROBACIÓN DE SOLICITUD RECURRENTE DESDE EMAIL
   ============================================ */
function handleAprobarRecurrenteDesdeEmail(idSolicitud) {
  try {
    if (!isUserAdmin()) {
      return HtmlService.createHtmlOutput(buildHtmlCancelPage(
        "Acceso denegado",
        "Solo un administrador puede aprobar solicitudes."
      ));
    }

    // Obtener datos de la solicitud
    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheet.getDataRange().getValues();

    let filaEncontrada = -1;
    let solicitud = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idSolicitud).trim()) {
        filaEncontrada = i + 1;
        solicitud = {
          id: data[i][0],
          recurso: data[i][2],
          usuario: data[i][4],
          email: data[i][3],
          dias: data[i][5],
          tramo: data[i][7],
          estado: data[i][11]
        };
        break;
      }
    }

    if (!solicitud) {
      return HtmlService.createHtmlOutput(buildHtmlCancelPage(
        "Solicitud no encontrada",
        "No se ha encontrado la solicitud de reserva recurrente."
      ));
    }

    // Verificar que está pendiente
    if (solicitud.estado && solicitud.estado.toLowerCase() !== 'pendiente') {
      return HtmlService.createHtmlOutput(buildHtmlCancelPage(
        "Solicitud ya procesada",
        `Esta solicitud ya fue ${solicitud.estado.toLowerCase()}.`
      ));
    }

    // Aprobar la solicitud
    const resultado = aprobarSolicitudRecurrente(idSolicitud, 'Aprobada desde email');

    if (resultado && resultado.success) {
      return HtmlService.createHtmlOutput(buildHtmlCancelPage(
        "Solicitud Aprobada",
        `La solicitud de reserva recurrente de <strong>${escHtml_(solicitud.usuario)}</strong> para <strong>${escHtml_(solicitud.recurso)}</strong> ha sido aprobada correctamente.`,
        {
          recursoNombre: solicitud.recurso,
          fechaFormateada: `${solicitud.dias}`,
          tramoNombre: solicitud.tramo
        }
      ));
    } else {
      return HtmlService.createHtmlOutput(buildHtmlCancelPage(
        "Error al aprobar",
        resultado?.error || "Ha ocurrido un error al aprobar la solicitud."
      ));
    }

  } catch (error) {
    Logger.log('❌ Error en handleAprobarRecurrenteDesdeEmail: ' + error.message);
    return HtmlService.createHtmlOutput(buildHtmlCancelPage(
      "Error",
      "Ha ocurrido un error al procesar la aprobación: " + error.message
    ));
  }
}


/* ============================================
   NUEVAS FUNCIONES DE ALTA DE USUARIO 🚀
   ============================================ */

// Página mostrada cuando Session.getActiveUser() devuelve vacío
function buildHtmlEmailNoDisponible_() {
  let dominioDespliegue = '';
  try { dominioDespliegue = (Session.getEffectiveUser().getEmail() || '').split('@')[1] || ''; } catch (e) { }
  return `
    <!DOCTYPE html>
    <html>
      <body style="font-family: sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; background-color: #fef2f2; padding: 16px; box-sizing: border-box;">
        <div style="max-width: 520px; text-align: center; padding: 32px; background: white; border-radius: 15px; box-shadow: 0 10px 25px rgba(0,0,0,0.1);">
          <div style="font-size: 50px; margin-bottom: 10px;">🔒</div>
          <h1 style="color: #991b1b; margin: 0; font-size: 22px;">No se pudo identificar tu cuenta</h1>
          <p style="color: #4b5563; margin-top: 14px; line-height: 1.5;">
            Google no ha facilitado tu correo a la aplicación, por lo que no es posible comprobar tu acceso.
          </p>
          <ul style="text-align: left; color: #4b5563; line-height: 1.6; font-size: 14px;">
            <li>Entra con tu cuenta del centro${dominioDespliegue ? ' (<b>@' + dominioDespliegue + '</b>)' : ''}. Si tienes varias cuentas abiertas, prueba en una ventana de incógnito.</li>
            <li><b>Administrador/a:</b> la aplicación debe implementarse desde una cuenta del <b>mismo dominio</b> que el profesorado (p.ej. @g.educaand.es). Si se implementó desde una cuenta @gmail.com u otro dominio, Google oculta el correo de los usuarios y la aplicación les pedirá registrarse siempre.</li>
          </ul>
        </div>
      </body>
    </html>`;
}

// 1. EL USUARIO ENVÍA LA SOLICITUD (CORREGIDO)
function procesarSolicitudRegistro(nombreSolicitante, emailManual) {
  // Intentamos obtenerlo de la sesión, si falla, usamos el que escribió el usuario
  // Solo se admite el email de la sesión: un email tecleado no se puede verificar
  // y, tras aprobarlo, el usuario seguiría sin poder entrar (bucle de registro).
  const emailFinal = Session.getActiveUser().getEmail();

  if (!emailFinal) {
    throw new Error("Google no ha facilitado tu correo a la aplicación. Entra con tu cuenta del centro o avisa al administrador (la app debe implementarse desde una cuenta del mismo dominio).");
  }

  const scriptUrl = ScriptApp.getService().getUrl();
  const admins = getAdminsEmails_();
  if (admins.length === 0) {
    throw new Error("No hay administradores configurados. Contacta con el responsable del sistema.");
  }
  const emailAdmin = admins[0];

  // Enlace con los datos correctos
  const enlaceAprobar = `${scriptUrl}?action=approve&email=${encodeURIComponent(emailFinal)}&nombre=${encodeURIComponent(nombreSolicitante)}`;

  const asunto = `🔔 Nueva Solicitud: ${nombreSolicitante}`;
  const cuerpo = `
    <div style="font-family: sans-serif; max-width: 500px; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #2563EB; padding: 20px; color: white; text-align: center;">
        <h2 style="margin: 0;">Solicitud de Acceso</h2>
      </div>
      <div style="padding: 20px;">
        <p>Hola Admin,</p>
        <p>Un nuevo compañero quiere acceder al sistema:</p>
        <ul style="background: #f9f9f9; padding: 15px; border-radius: 5px; list-style: none;">
          <li>👤 <strong>Nombre:</strong> ${escHtml_(nombreSolicitante)}</li>
          <li>📧 <strong>Email:</strong> ${emailFinal}</li>
        </ul>
        <p>Haz clic abajo para darle acceso inmediato:</p>
        <div style="text-align: center; margin-top: 20px;">
          <a href="${enlaceAprobar}" style="background-color: #10B981; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">
            ✅ Aprobar y Registrar
          </a>
        </div>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: emailAdmin,
    subject: asunto,
    htmlBody: cuerpo
  });

  return { success: true };
}

// 2. EL ADMIN APRUEBA (CON LIMPIEZA DE CACHÉ Y EMAIL 📧✨)
function handleAdminApproval(emailNuevo, nombreNuevo) {
  // 1. SEGURIDAD: Verificar que quien hace clic es Admin
  const emailAdmin = Session.getActiveUser().getEmail();
  const auth = checkUserAuthorization(emailAdmin);

  if (!auth.isAdmin) {
    return HtmlService.createHtmlOutput("<h1>⛔ Acceso Denegado</h1><p>Solo un administrador puede aprobar solicitudes.</p>");
  }

  const ss = getDB();
  const sheet = ss.getSheetByName(SHEETS.USUARIOS);
  const data = sheet.getDataRange().getValues();

  // Comprobamos si ya existe para no duplicar
  const existe = data.slice(1).some(row => String(row[1]).toLowerCase() === emailNuevo.toLowerCase());

  if (!existe) {
    // A. GUARDAR EN EXCEL
    // Nombre, Email, Activo (TRUE), Admin (FALSE)
    sheet.appendRow([nombreNuevo, emailNuevo, true, false, '']); // ✅ 5 columnas

    // B. PURGAR CACHÉ (¡ESTO SOLUCIONA TU ESPERA!) 🧹
    // Obligamos al sistema a volver a leer el Excel en la próxima carga
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY_STATIC);

    // C. ENVIAR CORREO DE BIENVENIDA 📧
    try {
      const asunto = "🎉 ¡Acceso Concedido! - Sistema de Reservas";
      const cuerpo = `
        <div style="font-family: sans-serif; max-width: 500px; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
          <div style="background-color: #10B981; padding: 20px; color: white; text-align: center;">
            <h2 style="margin: 0;">¡Bienvenido/a a bordo!</h2>
          </div>
          <div style="padding: 20px;">
            <p>Hola <strong>${escHtml_(nombreNuevo)}</strong>,</p>
            <p>Tu solicitud de acceso ha sido aprobada por el administrador.</p>
            <p>Ya puedes entrar y realizar reservas.</p>
            <div style="text-align: center; margin-top: 25px;">
              <a href="${ScriptApp.getService().getUrl()}" style="background-color: #2563EB; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold;">
                🔗 Entrar al Sistema
              </a>
            </div>
          </div>
        </div>
      `;

      MailApp.sendEmail({
        to: emailNuevo,
        subject: asunto,
        htmlBody: cuerpo
      });

    } catch (e) {
      console.log("No se pudo enviar email de bienvenida: " + e.message);
    }
  }

  // D. PANTALLA DE ÉXITO PARA EL ADMIN
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <body style="font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background-color: #f0fdf4;">
        <div style="text-align: center; padding: 40px; background: white; border-radius: 15px; box-shadow: 0 10px 25px rgba(0,0,0,0.1);">
          <div style="font-size: 60px; margin-bottom: 20px;">✅</div>
          <h1 style="color: #166534; margin: 0;">Usuario Registrado</h1>
          <p style="color: #4b5563; margin-top: 10px;">
            Se ha dado de alta a <b>${escHtml_(nombreNuevo)}</b><br>
            <span style="font-size: 0.9em; color: #6b7280;">(${escHtml_(emailNuevo)})</span>
          </p>
          <p style="margin-top: 20px; color: #059669; font-weight: bold;">
            📧 Se le ha enviado un correo de aviso.
          </p>
          <button onclick="window.close()" style="margin-top: 20px; padding: 10px 25px; background: #166534; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 16px;">
            Cerrar Ventana
          </button>
        </div>
      </body>
    </html>
  `);
}

// Helper para buscar emails de admins
function getAdminsEmails_() {
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.USUARIOS);
    const data = sheet.getDataRange().getValues();
    // Asumimos Col B=Email, Col D=Admin
    // Filtramos filas donde Col D es true/si/yes
    // Mismos valores que acepta checkUserAuthorization (TRUE, Sí, yes, admin)
    const admins = data.filter((row, i) => i > 0 && row[1] &&
        (esValorVerdadero_(row[3]) || String(row[3]).toLowerCase().trim() === 'admin'))
      .map(row => row[1]);
    return admins;
  } catch (e) { return []; }
}

/* ============================================
   GESTIÓN DE CONFIGURACIÓN
   ============================================ */

/**
 * Obtiene toda la configuración del sistema
 * @returns {Object} Objeto con todas las configuraciones
 */
function getConfiguracion() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEYS.CONFIGURACION);
  
  if (cached) {
    Logger.log("✅ Configuración desde caché");
    return JSON.parse(cached);
  }
  
  Logger.log("🔄 Cargando configuración desde Excel...");
  const sheet = getDB().getSheetByName(SHEETS.CONFIG);
  
  if (!sheet) {
    Logger.log("⚠️ Hoja Config no encontrada");
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  // Parsear la tabla (columnas: CLAVE, VALOR, DESCRIPCION)
  for (let i = 1; i < data.length; i++) {
    const clave = data[i][0];
    let valor = data[i][1];
    
    if (!clave) continue;
    
    config[clave] = parsearValorConfig_(valor);
  }
  
  // Cachear por 5 minutos
  try {
    cache.put(CACHE_KEYS.CONFIGURACION, JSON.stringify(config), 300);
    Logger.log(`💾 Configuración cacheada: ${Object.keys(config).length} parámetros`);
  } catch (e) {
    Logger.log("⚠️ Error guardando caché de configuración: " + e.message);
  }
  
  return config;
}

/**
 * Convierte un valor de la hoja Config a su tipo JS.
 * Antes los booleanos reales (casillas) acababan como Number(true) = 1 y
 * las comprobaciones "=== true" nunca se cumplían.
 */
function parsearValorConfig_(valor) {
  if (typeof valor === 'boolean' || typeof valor === 'number') return valor;
  if (valor instanceof Date) return valor;
  const txt = String(valor).trim();
  if (/^(true|false)$/i.test(txt)) return txt.toLowerCase() === 'true';
  if (txt !== '' && !isNaN(txt)) return Number(txt);
  return valor;
}

/**
 * Obtiene un valor específico de configuración
 * @param {string} clave - La clave del parámetro a obtener
 * @param {*} valorPorDefecto - Valor por defecto si no se encuentra
 * @returns {*} El valor configurado o el valor por defecto
 */
function getConfigValue(clave, valorPorDefecto = null) {
  const config = getConfiguracion();
  return config[clave] !== undefined ? config[clave] : valorPorDefecto;
}

/* ============================================
   VALIDACIONES DE LÍMITES
   ============================================ */

/**
 * Valida que el sistema no esté en modo mantenimiento
 * @throws {Error} Si el sistema está en mantenimiento
 */
function validarModoMantenimiento() {
  const modoMantenimiento = getConfigValue('modo_mantenimiento', false);
  
  if (modoMantenimiento === true) {
    throw new Error("El sistema está en mantenimiento. Inténtalo más tarde.");
  }
  
  Logger.log("✅ Sistema operativo (no en mantenimiento)");
  return true;
}

/**
 * Valida que la fecha no exceda el límite de días a futuro
 * @param {string} fechaISO - Fecha en formato YYYY-MM-DD
 * @throws {Error} Si la fecha excede el límite configurado
 */
function validarDiasVista(fechaISO) {
  const diasMaximo = getConfigValue('dias_vista_maximo', 5);
  
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  const fechaReserva = new Date(fechaISO + "T00:00:00Z");
  fechaReserva.setHours(0, 0, 0, 0);
  
  const diferenciaMs = fechaReserva - hoy;
  const diasDiferencia = Math.floor(diferenciaMs / (1000 * 60 * 60 * 24));
  
  if (diasDiferencia > diasMaximo) {
    throw new Error(`No puedes reservar con más de ${diasMaximo} días de antelación.`);
  }
  
  Logger.log(`✅ Fecha válida: ${diasDiferencia} días de antelación (máx: ${diasMaximo})`);
  return true;
}

/**
 * Valida que haya suficiente antelación antes del tramo
 * @param {string} fechaISO - Fecha en formato YYYY-MM-DD
 * @param {string} tramoId - ID del tramo horario
 * @throws {Error} Si no hay suficiente antelación
 */
function validarAntelacionMinima(fechaISO, tramoId) {
  const minutosMinimos = getConfigValue('minutos_antelacion', 30);
  
  // Obtener el tramo para conocer su hora de inicio
  const staticData = getDatosEstaticos_();
  const tramo = staticData.tramos.find(t => String(t.id_tramo) === String(tramoId));
  
  if (!tramo || !tramo.hora_inicio) {
    // Si no hay hora, solo validar que no sea una fecha pasada
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const fechaReserva = new Date(fechaISO + "T00:00:00Z");
    fechaReserva.setHours(0, 0, 0, 0);
    
    if (fechaReserva < hoy) {
      throw new Error("No puedes reservar en fechas pasadas.");
    }
    
    Logger.log("⚠️ Tramo sin hora de inicio, solo validando fecha no pasada");
    return true;
  }
  
  // Construir la fecha/hora exacta del inicio del tramo
  // Hora local del centro (zona del script). Antes se usaba UTC: desfase de 1-2 h.
  const [horas, minutos] = String(tramo.hora_inicio).split(':').map(Number);
  const hhmm = ('0' + horas).slice(-2) + ':' + ('0' + (minutos || 0)).slice(-2);
  const fechaHoraTramo = Utilities.parseDate(fechaISO + ' ' + hhmm, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  
  // Obtener la hora actual
  const ahora = new Date();
  
  // Calcular diferencia en minutos
  const diferenciaMs = fechaHoraTramo - ahora;
  const diferenciaMinutos = Math.floor(diferenciaMs / (1000 * 60));
  
  if (diferenciaMinutos < minutosMinimos) {
    throw new Error(`Debes reservar con al menos ${minutosMinimos} minutos de antelación.`);
  }
  
  Logger.log(`✅ Antelación válida: ${diferenciaMinutos} minutos (mín: ${minutosMinimos})`);
  return true;
}

/**
 * horas_cancelacion (Config): antelación mínima para que un usuario cancele.
 * Devuelve el mensaje de error si NO se puede cancelar, o null si sí.
 * Los administradores quedan exentos (se comprueba fuera).
 */
function motivoBloqueoCancelacion_(fechaValor, tramoId) {
  const horas = parseFloat(getConfigValue('horas_cancelacion', 0)) || 0;
  if (horas <= 0) return null;

  const tramo = getDatosEstaticos_().tramos.find(t => String(t.id_tramo).trim() === String(tramoId).trim());
  if (!tramo || !tramo.hora_inicio) return null;

  const tz = Session.getScriptTimeZone();
  const fechaISO = fechaValor instanceof Date
    ? Utilities.formatDate(fechaValor, tz, 'yyyy-MM-dd')
    : String(fechaValor).substring(0, 10);
  const [h, m] = String(tramo.hora_inicio).split(':').map(Number);
  if (isNaN(h)) return null;
  const hhmm = ('0' + h).slice(-2) + ':' + ('0' + (m || 0)).slice(-2);
  const inicio = Utilities.parseDate(fechaISO + ' ' + hhmm, tz, 'yyyy-MM-dd HH:mm');

  if (inicio.getTime() - Date.now() < horas * 3600 * 1000) {
    return `Solo se puede cancelar con al menos ${horas} ${horas === 1 ? 'hora' : 'horas'} de antelación. Si necesitas anularla, contacta con un administrador.`;
  }
  return null;
}

/**
 * Valida que el usuario no exceda el límite de reservas activas
 * @param {string} email - Email del usuario
 * @throws {Error} Si el usuario excede el límite de reservas
 */
function validarLimiteReservas(email, reservasActivasPrevias, nuevas) {
  const limiteReservas = getConfigValue('limite_reservas', 3);

  const reservasActivas = reservasActivasPrevias || getActiveReservations_();
  const emailNorm = String(email).toLowerCase().trim();
  // Solo contar reservas manuales (excluir las generadas por recurrencias)
  const reservasUsuario = reservasActivas.filter(r =>
    String(r.email_usuario).toLowerCase().trim() === emailNorm && !r.id_solicitud_recurrente
  );

  if (reservasUsuario.length + (nuevas || 1) > limiteReservas) {
    if (nuevas > 1 && reservasUsuario.length < limiteReservas) {
      throw new Error(`Con ${nuevas} tramos superarías tu límite de ${limiteReservas} reservas activas (tienes ${reservasUsuario.length}).`);
    }
    throw new Error(`Has alcanzado el límite de ${limiteReservas} reservas activas. Cancela alguna para continuar.`);
  }

  Logger.log(`✅ Límite de reservas: ${reservasUsuario.length}/${limiteReservas}`);
  return true;
}

/* ============================================
   FUNCIÓN AUXILIAR PARA LEER LA CONFIG RÁPIDO Y ACTUALIZAR
   ============================================ */
function getAppConfig() {
  // Usa la configuración cacheada (antes se leía la hoja Config en cada carga de página)
  const cfg = getConfiguracion();
  return {
    appName: cfg.nombre_centro ? String(cfg.nombre_centro) : "Sistema de Reservas",
    logoUrl: cfg.url_logo ? convertirUrlLogoParaMostrar(String(cfg.url_logo)) : ""
  };
}

/* ============================================
   FUNCIÓN AUXILIAR PARA INCLUIR ARCHIVOS HTML
   ============================================ */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================
   UTILIDADES DE MANTENIMIENTO
   ============================================ */

function purgarCache() {
  const cache = CacheService.getScriptCache();
  // 1. Borrar la clave principal (La importante)
  cache.remove(CACHE_KEY_STATIC);

  // 2. (Opcional) Limpieza profunda de versiones antiguas por si acaso
  // Esto ayuda si cambiaste de versión y quedaron residuos
  cache.removeAll([
    'STATIC_DATA_V5',
    'STATIC_DATA_V4',
    'STATIC_DATA_V3',
    CACHE_KEYS.CONFIGURACION
  ]);

  console.log("✅ Caché purgada correctamente. La próxima carga será desde Excel.");

  // Disponibilidad por recurso: basta con los IDs (columna A de Recursos, incluidos los inactivos).
  // Antes se regeneraban todos los datos estáticos solo para obtener esta lista.
  try {
    const sheetRec = getDB().getSheetByName(SHEETS.RECURSOS);
    if (sheetRec && sheetRec.getLastRow() > 1) {
      const claves = [];
      sheetRec.getRange(2, 1, sheetRec.getLastRow() - 1, 1).getValues().forEach(r => {
        const raw = String(r[0]);
        if (!raw.trim()) return;
        claves.push(CACHE_KEYS.DISPONIBILIDAD + raw, CACHE_KEYS.DISPONIBILIDAD + raw.trim());
      });
      if (claves.length) cache.removeAll(claves);
    }
  } catch (e) {
    Logger.log('⚠️ No se pudo purgar la caché de disponibilidad: ' + e.message);
  }

  Logger.log('Cachés de disponibilidad purgadas.');
}

/* ============================================
   FUNCIONES PARA MANEJO DE IMÁGENES DE DRIVE
   ============================================ */

/**
 * Extrae el ID de archivo de una URL de Google Drive
 * Soporta varios formatos:
 * - https://drive.google.com/file/d/ID/view
 * - https://drive.google.com/open?id=ID
 * - https://drive.google.com/uc?id=ID
 * @param {string} url - URL de Google Drive
 * @returns {string|null} - ID del archivo o null si no es URL de Drive
 */
function extraerIdDeDriveUrl(url) {
  if (!url || typeof url !== 'string') return null;

  // Patrón 1: /file/d/ID/
  let match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];

  // Patrón 2: ?id=ID o &id=ID
  match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];

  // Patrón 3: /d/ID/ (formato corto)
  match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];

  return null;
}

/**
 * Convierte una URL de Google Drive al formato de imagen pública
 * @param {string} url - URL original (puede ser de Drive o URL directa)
 * @returns {string} - URL convertida para mostrar imagen
 */
function convertirUrlLogoParaMostrar(url) {
  if (!url || typeof url !== 'string') return '';

  // Si ya es una URL de lh3.googleusercontent.com, devolverla tal cual
  if (url.includes('lh3.googleusercontent.com')) return url;

  // Intentar extraer ID de Drive
  const driveId = extraerIdDeDriveUrl(url);
  if (driveId) {
    return 'https://lh3.googleusercontent.com/d/' + driveId;
  }

  // Si no es de Drive, devolver la URL original
  return url;
}

/**
 * Intenta hacer pública una imagen de Google Drive
 * @param {string} url - URL de Google Drive
 * @returns {Object} - { success: boolean, message: string, convertedUrl: string }
 */
function procesarUrlLogoDrive(url) {
  try {
    if (!isUserAdmin()) {
      return { success: false, message: 'No tienes permisos de administrador', convertedUrl: '' };
    }
    if (!url || typeof url !== 'string') {
      return { success: false, message: 'URL vacía', convertedUrl: '' };
    }

    const driveId = extraerIdDeDriveUrl(url);

    // Si no es URL de Drive, simplemente devolver éxito con la URL original
    if (!driveId) {
      return {
        success: true,
        message: 'URL directa (no es de Drive)',
        convertedUrl: url,
        isDrive: false
      };
    }

    const convertedUrl = 'https://lh3.googleusercontent.com/d/' + driveId;

    // Intentar hacer el archivo público
    try {
      const file = DriveApp.getFileById(driveId);

      // Verificar si ya tiene acceso público
      const access = file.getSharingAccess();

      if (access !== DriveApp.Access.ANYONE && access !== DriveApp.Access.ANYONE_WITH_LINK) {
        // Hacer el archivo público (solo lectura)
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        Logger.log('✅ Permisos actualizados para imagen: ' + driveId);

        return {
          success: true,
          message: 'Imagen de Drive configurada como pública correctamente',
          convertedUrl: convertedUrl,
          isDrive: true,
          permissionsChanged: true
        };
      } else {
        return {
          success: true,
          message: 'La imagen de Drive ya era pública',
          convertedUrl: convertedUrl,
          isDrive: true,
          permissionsChanged: false
        };
      }
    } catch (permError) {
      // No tiene permisos para modificar el archivo
      Logger.log('⚠️ No se pudieron cambiar permisos: ' + permError.message);
      return {
        success: true,
        message: 'URL de Drive detectada, pero no se pudieron cambiar los permisos automáticamente. Asegúrate de que la imagen sea pública.',
        convertedUrl: convertedUrl,
        isDrive: true,
        permissionsChanged: false,
        warning: true
      };
    }

  } catch (error) {
    Logger.log('❌ Error procesando URL de logo: ' + error.message);
    return {
      success: false,
      message: 'Error: ' + error.message,
      convertedUrl: url
    };
  }
}


/**
 * =========================================================================== 
 * ⚙️ SETUP.GS - GESTIÓN DE INSTALACIÓN Y DATOS INICIALES
 * ===========================================================================
 */

// --- CONFIGURACIÓN DE COLUMNAS (TU ESQUEMA CORRECTO) ---
const DB_SCHEMA = {
  'Recursos': { headers: ['ID_Recurso', 'Nombre', 'Tipo', 'Icono', 'Ubicacion', 'Capacidad', 'Descripcion', 'Estado'], color: '#4f46e5' },
  'Tramos':   { headers: ['ID_Tramo', 'Nombre_Tramo', 'Hora_Inicio', 'Hora_Fin'], color: '#d97706' },
  'Disponibilidad': { headers: ['ID_Recurso', 'Dia_Semana', 'ID_Tramo', 'Hora_Inicio', 'Permitido', 'Razon_Bloqueo'], color: '#7c3aed' },
  'Cursos':   { headers: ['Etapa', 'Curso', 'Mostrar el curso con:'], color: '#0891b2' },
  'Usuarios': { headers: ['Nombre_Completo', 'Email_Usuario', 'Activo', 'Admin', 'Especialidad'], color: '#dc2626' },
  'Config':   { headers: ['CLAVE', 'VALOR', 'DESCRIPCION'], color: '#64748b' },
  'Incidencias':   { headers: ['ID_Incidencia',	'ID_Recurso',	'Nombre_Recurso',	'Email_Usuario',	'Fecha_Reporte',	'Categoria',	'Prioridad',	'Descripcion',	'Estado',	'Notas_Admin',	'Fecha_Resolucion'], color: '#64748b' },
  'Reservas': { headers: ['ID_Reserva', 'ID_Recurso', 'Email_Usuario', 'Fecha', 'Curso', 'ID_Tramo', 'Cantidad', 'Estado', 'Notas', 'Timestamp', 'ID_Solicitud_Recurrente'], color: '#059669' },
  'SolicitudesRecurrentes': { headers: ['ID_Solicitud', 'ID_Recurso', 'Nombre_Recurso', 'Email_Usuario', 'Nombre_Usuario', 'Dias_Semana', 'ID_Tramo', 'Nombre_Tramo', 'Fecha_Inicio', 'Fecha_Fin', 'Motivo', 'Estado', 'Fecha_Solicitud', 'Fecha_Resolucion', 'Admin_Resolutor', 'Notas_Admin'], color: '#8b5cf6' }
};

// ==========================================
// 🌐 FUNCIÓN QUE LLAMA EL HTML 'ActivacionSistema'
// ==========================================
function ejecutarSetupVinculado() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentUser = Session.getActiveUser().getEmail();
    
    // 🔥 MAGIA: Obtenemos la URL actual automáticamente
    const currentUrl = ScriptApp.getService().getUrl(); 

    if (!currentUrl) {
      throw new Error("No se pudo detectar la URL de la Web App. Asegúrate de haber desplegado correctamente.");
    }

    // 1. Crear estructura (pestañas)
    crearEstructuraInterna(ss);

    // 2. Rellenar datos y GUARDAR LA URL EN CONFIG (CORREGIDO: ahora pasa currentUrl y email admin)
    crearDatosEjemploInterno(ss, currentUrl, currentUser);

    // 3. Hacer Admin al usuario que está ejecutando esto
    asegurarAdminInterno(ss, currentUser);

    // 4. Marcar como instalado (IMPORTANTE: usar SETUP_COMPLETED para consistencia)
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SETUP_COMPLETED', 'true');
    props.setProperty('INSTALL_DATE', new Date().toISOString());
    props.setProperty('WEB_APP_URL', currentUrl);
    props.setProperty('FECHA_ACTIVACION', new Date().toISOString());

    // 5. Instalar comprobación automática de actualizaciones (diaria)
    try {
      instalarTriggerActualizaciones();
    } catch (triggerErr) {
      Logger.log('⚠️ No se pudo instalar trigger de actualizaciones: ' + triggerErr.message);
    }

    // Limpieza opcional de la hoja por defecto
    const hojaDefault = ss.getSheetByName("Hoja 1");
    if (hojaDefault && hojaDefault.getLastRow() === 0) {
      ss.deleteSheet(hojaDefault);
    }

    return { success: true, url: currentUrl };

  } catch (e) {
    Logger.log("ERROR SETUP: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 🔧 FUNCIONES INTERNAS
// ==========================================

function crearEstructuraInterna(ss) {
  const sheets = ss.getSheets();
  if (sheets.length > 0) sheets[0].setName("Temp_Init");

  for (const [sheetName, config] of Object.entries(DB_SCHEMA)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    
    const r = sheet.getRange(1, 1, 1, config.headers.length);
    r.setValues([config.headers]);
    r.setBackground(config.color).setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    if(sheetName === 'Cursos') sheet.getRange('F1').setValue('botones');
  }
  
  const tempSheet = ss.getSheetByName("Temp_Init");
  if (tempSheet) ss.deleteSheet(tempSheet);
}

// --- 🌟 FUNCIÓN CORREGIDA: Ahora recibe currentUrl y adminEmail como parámetros ---
function crearDatosEjemploInterno(ss, currentUrl, adminEmail) {

  // A. RECURSOS
  const sheetRec = ss.getSheetByName('Recursos');
  const recursos = [
    ['REC-INFO', 'Sala Informática', 'Sala', 'mdi:desktop-classic', 'Planta 1', 25, 'Sala con PCs fijos', 'Activo'],
    ['REC-CARR1', 'Carro Portátiles 1', 'Agrupado', 'mdi:laptop', 'Secretaría', 30, '30 Chromebooks', 'Activo']
  ];
  sheetRec.getRange(2, 1, recursos.length, recursos[0].length).setValues(recursos);

  // B. TRAMOS
  const sheetTram = ss.getSheetByName('Tramos');
  const tramos = [
    ['T001', '1ª Hora', '09:00', '10:00'],
    ['T002', '2ª Hora', '10:00', '11:00'],
    ['T003', '3ª Hora', '11:00', '11:30'],
    ['T004', '4ª Hora', '12:00', '13:00'],
    ['T005', '5ª Hora', '13:00', '14:00']
  ];
  sheetTram.getRange(2, 1, tramos.length, tramos[0].length).setValues(tramos);

  // C. CURSOS
  const sheetCur = ss.getSheetByName('Cursos');
  const cursos = [
    ['Primaria', '6º Primaria A', 1, 'PRI-6A', '']
  ];
  sheetCur.getRange(2, 1, cursos.length, cursos[0].length).setValues(cursos);

  // D. CONFIGURACIÓN (Con email del admin instalador y opciones de notificación)
  const sheetConfig = ss.getSheetByName('Config');
  const configData = [
    ['dias_vista_maximo', 30, 'Días a futuro permitidos'],
    ['minutos_antelacion', 0, 'Minutos mínimos antes de reservar'],
    ['limite_reservas', 3, 'Máx. reservas activas por usuario'],
    ['horas_cancelacion', 0, 'Horas mínimas para poder cancelar solo'],
    ['exigir_motivo', 'FALSE', 'Obligatorio escribir para qué es'],
    ['email_admin', adminEmail || '', 'Email del administrador para notificaciones'],
    ['admin_recibir_copia_reservas', 'FALSE', 'Recibir copia oculta de confirmaciones de reservas'],
    ['modo_mantenimiento', 'FALSE', 'Bloquear nuevas reservas (Pánico)'],
    ['permitir_multitramo', 'FALSE', 'Permitir seleccionar varios tramos a la vez'],
    ['max_tramos_simultaneos', 1, 'Cuántos tramos seguidos se pueden coger de golpe'],
    ['nombre_centro', 'Sistema de Reservas', 'Nombre del centro'],
    ['url_logo', '', 'URL del logo del centro'],
    ['url_webapp', currentUrl, 'URL automática de la aplicación']
  ];
  sheetConfig.getRange(2, 1, configData.length, 3).setValues(configData);
}

function asegurarAdminInterno(ss, email) {
  const sheet = ss.getSheetByName('Usuarios');
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === email.toLowerCase()) {
      found = true;
      sheet.getRange(i+1, 3).setValue(true); // Activo
      sheet.getRange(i+1, 4).setValue(true); // Admin
      break;
    }
  }
  
  if (!found) {
    const nombre = email.split('@')[0].toUpperCase();
    sheet.appendRow([nombre, email, true, true, 'Super Admin']);
  }
}


/**
 * ===========================================================================
 * 🔧 HERRAMIENTAS DE MANTENIMIENTO Y REPARACIÓN
 * ===========================================================================
 */

/**
 * Sincroniza el estado de instalación y actualiza la URL en la configuración.
 * * ¿CUÁNDO USAR ESTA FUNCIÓN?
 * 1. Si el sistema te pide "Instalar" pero tú ya tienes la hoja configurada.
 * 2. Si has cambiado la implementación de la Web App y la URL ha cambiado.
 * 3. Si has copiado el archivo y quieres reactivarlo rápidamente.
 * * ¿QUÉ HACE?
 * - Marca internamente el sistema como 'SETUP_COMPLETED'.
 * - Detecta la URL actual de la Web App.
 * - Escribe o actualiza esa URL en la hoja 'Config' (fila 'url_webapp').
 * * @return {void} Solo imprime logs en la consola.
 */
function repararInstalacionYGuardarURL() {
  Logger.log("🔧 INICIANDO REPARACIÓN DEL SISTEMA...");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConfig = ss.getSheetByName('Config');
    
    // --- PASO 1: OBTENER URL ---
    // Nota: Requiere que el script esté implementado como Web App.
    const url = ScriptApp.getService().getUrl();
    
    if (!url) {
      Logger.log("❌ ERROR CRÍTICO: No se detecta una URL de Web App activa.");
      Logger.log("   -> Solución: Ve a 'Gestionar implementaciones' y asegúrate de que existe una versión activa.");
      return;
    }

    // --- PASO 2: ACTUALIZAR MEMORIA INTERNA (PropertiesService) ---
    // Esto es lo que consulta el doGet() para saber si mostrar el instalador.
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SETUP_COMPLETED', 'true');
    props.setProperty('WEB_APP_URL', url); // Guardamos también en memoria por redundancia
    // Opcional: Actualizar fecha de instalación/reparación
    // props.setProperty('INSTALL_DATE', new Date().toISOString()); 
    
    Logger.log("✅ Memoria del Script (PropertiesService) actualizada correctamente.");

    // --- PASO 3: ACTUALIZAR HOJA VISIBLE 'Config' ---
    if (sheetConfig) {
      const data = sheetConfig.getDataRange().getValues();
      let encontrada = false;

      // Buscamos si ya existe la clave 'url_webapp' para no duplicarla
      for (let i = 0; i < data.length; i++) {
        // Asumimos que la Columna A es la CLAVE y la Columna B es el VALOR
        if (data[i][0] && data[i][0].toString() === 'url_webapp') {
          sheetConfig.getRange(i + 1, 2).setValue(url);
          Logger.log(`✏️ URL actualizada en la fila ${i + 1} de la hoja 'Config'.`);
          encontrada = true;
          break;
        }
      }

      // Si no existe, la creamos nueva al final
      if (!encontrada) {
        // Estructura: [CLAVE, VALOR, DESCRIPCIÓN]
        sheetConfig.appendRow(['url_webapp', url, 'URL automática de la aplicación (Actualizada manualmente)']);
        Logger.log("➕ Fila 'url_webapp' añadida al final de la hoja 'Config'.");
      }
    } else {
      Logger.log("⚠️ AVISO: No se encontró la hoja 'Config'. Solo se actualizó la memoria interna.");
    }

    Logger.log("🎉 REPARACIÓN COMPLETADA.");
    Logger.log("   -> Ahora puedes recargar tu Web App y entrará directamente.");
    Logger.log("   -> URL registrada: " + url);

  } catch (e) {
    Logger.log("❌ EXCEPCIÓN: Ocurrió un error inesperado.");
    Logger.log(e.toString());
  }
}

/* ========================================================= 
   SISTEMA DE INCIDENCIAS - BACKEND
   ========================================================= */

/**
 * Reportar nueva incidencia
 */
function reportarIncidencia(datos) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("Usuario no identificado");

    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.INCIDENCIAS);

    // Generar ID único: Año corto + secuencial
    const año = new Date().getFullYear().toString().slice(-2); // "25"

    let maxNumeroAño = 0;
    if (sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const patron = new RegExp(`INC-${año}-(\\d+)`);

      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          const match = String(data[i][0]).match(patron);
          if (match) {
            const num = parseInt(match[1], 10);
            if (num > maxNumeroAño) maxNumeroAño = num;
          }
        }
      }
    }

    const nuevoId = `INC-${año}-${String(maxNumeroAño + 1).padStart(3, '0')}`;
    // Resultado: INC-25-001, INC-25-002... INC-26-001 (nuevo año)

    // Datos a guardar
    const nuevaFila = [
      nuevoId,                           // A - ID_Incidencia
      datos.id_recurso || '',            // B - ID_Recurso
      datos.nombre_recurso || '',        // C - Nombre_Recurso (cache)
      userEmail,                         // D - Email_Usuario
      new Date(),                        // E - Fecha_Reporte
      datos.categoria || 'Otro',         // F - Categoria
      datos.prioridad || 'Media',        // G - Prioridad
      datos.descripcion || '',           // H - Descripcion
      'Pendiente',                       // I - Estado
      '',                                // J - Notas_Admin (vacío)
      ''                                 // K - Fecha_Resolucion (vacío)
    ];

    // Insertar
    sheet.appendRow(nuevaFila);

    // Email al admin
    enviarEmailNuevaIncidencia({
      id: nuevoId,
      recurso: datos.nombre_recurso,
      usuario: userEmail,
      categoria: datos.categoria,
      prioridad: datos.prioridad,
      descripcion: datos.descripcion
    });

    purgarCache();

    return {
      success: true,
      message: 'Incidencia reportada correctamente',
      id: nuevoId
    };

  } catch (e) {
    Logger.log('Error reportarIncidencia: ' + e);
    return { success: false, error: e.toString() };
  }
}


/* =========================================================
   OBTENER INCIDENCIAS
   ========================================================= */

function getIncidencias() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Incidencias'); // Asegúrate que el nombre coincide

    if (!sheet) return { success: true, incidencias: [] };
    if (sheet.getLastRow() < 2) return { success: true, incidencias: [] };

    const data = sheet.getDataRange().getValues();
    const incidencias = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // Si hay ID

        // 1. LIMPIEZA FECHA REPORTE (Columna E -> Indice 4)
        let fechaReporte = '';
        if (data[i][4] && data[i][4] instanceof Date) {
          fechaReporte = data[i][4].toISOString();
        } else {
          fechaReporte = String(data[i][4] || '');
        }

        // 2. LIMPIEZA FECHA RESOLUCIÓN (Columna K -> Indice 10)
        // ESTO ES LO QUE FALLABA: Convertir fecha a String antes de enviar
        let fechaResolucion = '';
        if (data[i][10] && data[i][10] instanceof Date) {
          fechaResolucion = data[i][10].toISOString();
        } else {
          fechaResolucion = String(data[i][10] || '');
        }

        incidencias.push({
          ID_Incidencia: data[i][0],
          ID_Recurso: data[i][1],
          Nombre_Recurso: data[i][2],
          Email_Usuario: data[i][3],
          Fecha_Reporte: fechaReporte,
          Categoria: data[i][5],
          Prioridad: data[i][6],
          Descripcion: data[i][7],
          Estado: data[i][8],
          Notas_Admin: data[i][9],
          Fecha_Resolucion: fechaResolucion // <--- Ahora viaja como texto seguro
        });
      }
    }

    return { success: true, incidencias: incidencias };

  } catch (e) {
    Logger.log('Error getIncidencias: ' + e);
    return { success: false, error: e.toString() };
  }
}

/* =========================================================
   NUEVA FUNCIÓN DE GESTIÓN
   ========================================================= */

/**
 * Función flexible para manejar los botones del Frontend
 * @param {string} idIncidencia - El ID (ej: "INC-0001")
 * @param {string} accion - 'RESOLVER' o 'EDITAR_NOTA'
 * @param {string} valor - El nuevo valor (o null si es resolver simple)
 */
/* --- ACTUALIZAR BACKEND --- */

function backend_actualizarIncidencia(idIncidencia, accion, valor) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    const ss = getDB();
    const sheet = ss.getSheetByName('Incidencias');
    const data = sheet.getDataRange().getValues();

    // Buscar fila
    let fila = -1, recName = '', userEmail = '', oldNote = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idIncidencia)) {
        fila = i + 1;
        recName = data[i][2];
        userEmail = data[i][3];
        oldNote = data[i][9];
        break;
      }
    }
    if (fila === -1) throw new Error("Incidencia no encontrada");

    // --- LOGICA ACCIONES ---

    if (accion === 'PRIORIDAD') {
      sheet.getRange(fila, 7).setValue(valor);
      return { exito: true };
    }

    // ✅ AÑADIR ESTO: Cambiar estado sin resolver
    if (accion === 'ESTADO') {
      sheet.getRange(fila, 9).setValue(valor);
      return { exito: true };
    }

    if (accion === 'RESOLVER') {
      sheet.getRange(fila, 9).setValue('Resuelta');
      sheet.getRange(fila, 11).setValue(new Date());

      let notaFinal = valor || oldNote;
      if (valor) sheet.getRange(fila, 10).setValue(valor);

      enviarEmailIncidenciaResuelta({
        id: idIncidencia, recurso: recName, email: userEmail, notas: notaFinal
      });
      return { exito: true };
    }

    if (accion === 'EDITAR_NOTA') {
      sheet.getRange(fila, 10).setValue(valor);
      return { exito: true };
    }

  } catch (e) {
    return { exito: false, error: e.toString() };
  }
}

/**
 * Cambiar estado de recurso desde Incidencias (Versión corregida)
 */
function backend_toggleMantenimiento(idRecurso, nuevoEstado) {
  try {
    if (!isUserAdmin()) throw new Error("Acceso denegado");

    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.RECURSOS); // Usar constante si existe
    if (!sheet) throw new Error("Hoja 'Recursos' no encontrada");
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) throw new Error("No hay recursos en la hoja");

    // 1. Buscar índice de columnas por cabecera (más robusto)
    const cabeceras = data[0].map(c => String(c).toLowerCase().trim());
    const colIdIndex = cabeceras.findIndex(c => c === 'id_recurso' || c === 'id');
    const colEstadoIndex = cabeceras.findIndex(c => c === 'estado');

    if (colIdIndex === -1) throw new Error("No encuentro columna 'id_recurso' o 'id'");
    if (colEstadoIndex === -1) throw new Error("No encuentro columna 'estado'");

    // 2. Buscar fila del recurso
    let fila = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIdIndex]).trim() === String(idRecurso).trim()) {
        fila = i + 1; // +1 porque getRange es 1-indexed
        break;
      }
    }

    if (fila === -1) throw new Error(`Recurso '${idRecurso}' no encontrado`);

    // 3. Actualizar estado
    sheet.getRange(fila, colEstadoIndex + 1).setValue(nuevoEstado);

    // 4. Limpiar caché
    purgarCache();

    return { 
      exito: true, 
      idRecurso: idRecurso,
      nuevoEstado: nuevoEstado 
    };

  } catch (e) {
    Logger.log('Error backend_toggleMantenimiento: ' + e);
    return { exito: false, error: e.toString() };
  }
}

/**
 * Actualizar estado de incidencia (solo admin) ¿¿ESTO SOBRA??
 */
function actualizarEstadoIncidencia(idIncidencia, nuevoEstado, notasAdmin) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");

    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.INCIDENCIAS);
    const data = sheet.getDataRange().getValues();

    let filaEncontrada = -1;
    let emailUsuario = '';
    let nombreRecurso = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idIncidencia)) {
        filaEncontrada = i + 1;
        emailUsuario = data[i][3];
        nombreRecurso = data[i][2];
        break;
      }
    }

    if (filaEncontrada === -1) {
      throw new Error("Incidencia no encontrada");
    }

    // Actualizar estado (columna I = 9)
    sheet.getRange(filaEncontrada, 9).setValue(nuevoEstado);

    // Actualizar notas admin (columna J = 10)
    if (notasAdmin !== undefined) {
      sheet.getRange(filaEncontrada, 10).setValue(notasAdmin);
    }

    // Si se marca como resuelta, guardar fecha (columna K = 11)
    if (nuevoEstado === 'Resuelta') {
      sheet.getRange(filaEncontrada, 11).setValue(new Date());

      // Email al usuario
      enviarEmailIncidenciaResuelta({
        id: idIncidencia,
        recurso: nombreRecurso,
        email: emailUsuario,
        notas: notasAdmin
      });
    }

    purgarCache();

    return { success: true, message: 'Estado actualizado' };

  } catch (e) {
    Logger.log('Error actualizarEstadoIncidencia: ' + e);
    return { success: false, error: e.toString() };
  }
}




/* =========================================================
   EMAILS AUTOMÁTICOS
   ========================================================= */

function enviarEmailNuevaIncidencia(datos) {
  try {
    // Obtener email del admin desde CONFIG
    const ss = getDB();
    const configSheet = ss.getSheetByName('CONFIG');
    let emailAdmin = '';

    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      for (let i = 1; i < configData.length; i++) {
        if (configData[i][0] === 'email_admin') {
          emailAdmin = configData[i][1];
          break;
        }
      }
    }

    // Fallback: enviar al primer admin activo
    if (!emailAdmin) {
      const admins = getAdminsEmails();
      emailAdmin = admins[0] || Session.getActiveUser().getEmail();
    }

    const prioridadIcon = datos.prioridad === 'Crítica' ? '🔴' :
      datos.prioridad === 'Alta' ? '🟠' :
        datos.prioridad === 'Media' ? '🟡' : '🟢';

    const asunto = `⚠️ Nueva incidencia [${datos.prioridad}] - ${datos.recurso}`;

    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #eee; padding: 20px; border-radius: 8px;">
        <h2 style="color: #f57c00; margin-top: 0;">⚠️ Nueva Incidencia Reportada</h2>
        
        <div style="background: #fff3e0; padding: 15px; border-radius: 5px; border-left: 4px solid #ff9800; margin: 20px 0;">
          <p style="margin: 5px 0;"><strong>ID:</strong> ${datos.id}</p>
          <p style="margin: 5px 0;"><strong>Recurso:</strong> ${datos.recurso}</p>
          <p style="margin: 5px 0;"><strong>Categoría:</strong> ${datos.categoria}</p>
          <p style="margin: 5px 0;"><strong>Prioridad:</strong> ${prioridadIcon} ${datos.prioridad}</p>
          <p style="margin: 5px 0;"><strong>Reportado por:</strong> ${datos.usuario}</p>
        </div>
        
        <h3>Descripción:</h3>
        <p style="background: #f5f5f5; padding: 15px; border-radius: 5px; white-space: pre-wrap;">${datos.descripcion}</p>
        
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
        
        <p style="font-size: 0.9em; color: #666;">
          Accede al panel de administración para gestionar esta incidencia.
        </p>
      </div>
    `;

    MailApp.sendEmail({
      to: emailAdmin,
      subject: asunto,
      htmlBody: cuerpo
    });

    Logger.log(`📧 Email enviado a admin: ${emailAdmin}`);

  } catch (e) {
    Logger.log('⚠️ Error enviando email admin: ' + e);
  }
}


function enviarEmailIncidenciaResuelta(datos) {
  try {
    const asunto = `✅ Incidencia resuelta - ${datos.recurso}`;

    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #eee; padding: 20px; border-radius: 8px;">
        <h2 style="color: #4caf50; margin-top: 0;">✅ Incidencia Resuelta</h2>
        
        <p>La incidencia que reportaste ha sido marcada como <strong>resuelta</strong>.</p>
        
        <div style="background: #e8f5e9; padding: 15px; border-radius: 5px; border-left: 4px solid #4caf50; margin: 20px 0;">
          <p style="margin: 5px 0;"><strong>ID:</strong> ${datos.id}</p>
          <p style="margin: 5px 0;"><strong>Recurso:</strong> ${datos.recurso}</p>
        </div>
        
        ${datos.notas ? `
        <h3>Notas del administrador:</h3>
        <p style="background: #f5f5f5; padding: 15px; border-radius: 5px; white-space: pre-wrap;">${datos.notas}</p>
        ` : ''}
        
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
        
        <p style="font-size: 0.9em; color: #666;">
          El recurso ya está disponible para reservar nuevamente.
        </p>
      </div>
    `;

    MailApp.sendEmail({
      to: datos.email,
      subject: asunto,
      htmlBody: cuerpo
    });

    Logger.log(`📧 Email enviado a usuario: ${datos.email}`);

  } catch (e) {
    Logger.log('⚠️ Error enviando email usuario: ' + e);
  }
}


/**
 * ===========================================================================
 * 🔄 AUTO-UPDATER - SISTEMA DE ACTUALIZACIÓN AUTOMÁTICA v1.5.0
 * ===========================================================================
 *
 * Este módulo permite que las copias del sistema detecten nuevas versiones
 * y se actualicen automáticamente usando la API REST de Apps Script.
 *
 * CÓMO FUNCIONA:
 * 1. Cada copia tiene una versión local (SYSTEM_VERSION)
 * 2. Un trigger diario consulta un endpoint central (tu "Servidor de Versiones")
 * 3. Si hay una versión nueva, notifica al admin por email y toast
 * 4. El admin aplica la actualización con un clic → el código se reescribe solo
 * 5. El admin solo necesita redesplegar la Web App (nueva implementación)
 *
 * REQUISITOS PARA EL AUTOR:
 * - Desplegar el script "VersionServidor.gs" como Web App independiente
 * - Mantener el código fuente actualizado en GitHub
 * - Incrementar SYSTEM_VERSION en cada release
 *
 * REQUISITOS PARA LOS USUARIOS (una sola vez):
 * - Activar la API de Apps Script en https://script.google.com/home/usersettings
 */

/* ============================================
   CONFIGURACIÓN DE VERSIONES
   ============================================ */

/**
 * ⚠️ IMPORTANTE: Incrementar este número con cada actualización que publiques.
 * Formato: MAJOR.MINOR.PATCH (ej: "1.5.0")
 */
const SYSTEM_VERSION = '1.5.0';

/**
 * 🔗 URL del Servidor de Versiones (Web App desplegada por el autor).
 * El autor debe desplegar "VersionServidor.gs" y pegar aquí la URL resultante.
 *
 * INSTRUCCIONES PARA EL AUTOR:
 * 1. Crea un proyecto Apps Script independiente (no vinculado a ninguna hoja)
 * 2. Pega el contenido de "VersionServidor.gs"
 * 3. Despliégalo como Web App (Ejecutar como: Yo, Acceso: Cualquiera)
 * 4. Copia la URL y pégala aquí abajo
 */
const UPDATE_SERVER_URL = ''; // ← PEGAR AQUÍ la URL de tu servidor de versiones (y COMMITEARLA al repo)

/**
 * Devuelve la URL del servidor de versiones.
 * La propiedad UPDATE_SERVER_URL de ScriptProperties tiene prioridad sobre la
 * constante, de modo que una copia puede sobrescribirla sin tocar el código
 * y las actualizaciones no la borran.
 */
function getUpdateServerUrl_() {
  return PropertiesService.getScriptProperties().getProperty('UPDATE_SERVER_URL') || UPDATE_SERVER_URL;
}

/**
 * URL base del repositorio en GitHub (SIN rama/tag; el ref lo decide el
 * manifiesto que envía el servidor de versiones, con fallback a 'main').
 */
const GITHUB_REPO_RAW = 'https://raw.githubusercontent.com/maestroseb/reservasrecursosysalas';

/**
 * Lista LOCAL de archivos del sistema (fallback).
 * ⚠️ La lista autoritativa la envía el servidor de versiones en el campo
 * 'files' del JSON (ver VersionServidor.gs): así las copias antiguas siempre
 * actualizan con la lista de archivos de la versión NUEVA, no con la suya.
 * 'type': 'SERVER_JS' para .gs, 'HTML' para .html.
 * 'scriptName': nombre SIN extensión en la API de Apps Script.
 */
const UPDATABLE_FILES = [
  { name: 'Codigo.gs', type: 'SERVER_JS', scriptName: 'Codigo' },
  { name: 'AdminFunctions.gs', type: 'SERVER_JS', scriptName: 'AdminFunctions' },
  { name: 'ReservasRecurrentes.gs', type: 'SERVER_JS', scriptName: 'ReservasRecurrentes' },
  { name: 'Sistema.gs', type: 'SERVER_JS', scriptName: 'Sistema' },
  { name: 'index.html', type: 'HTML', scriptName: 'index' },
  { name: 'admin-panel.html', type: 'HTML', scriptName: 'admin-panel' },
  { name: 'admin-scripts.html', type: 'HTML', scriptName: 'admin-scripts' },
  { name: 'scripts.html', type: 'HTML', scriptName: 'scripts' },
  { name: 'styles.html', type: 'HTML', scriptName: 'styles' },
  { name: 'instalacion.html', type: 'HTML', scriptName: 'instalacion' }
];

/**
 * Devuelve el manifiesto de actualización guardado en la última comprobación:
 * { ref, files }. Si el servidor no envió manifiesto, usa los valores locales.
 * @private
 */
function getUpdateManifest_() {
  let ref = 'main';
  let files = UPDATABLE_FILES;
  try {
    const guardado = PropertiesService.getScriptProperties().getProperty('UPDATE_MANIFEST');
    if (guardado) {
      const manifest = JSON.parse(guardado);
      if (manifest.ref) ref = manifest.ref;
      if (manifest.files && manifest.files.length) files = manifest.files;
    }
  } catch (e) {
    Logger.log('⚠️ Manifiesto inválido, usando lista local: ' + e.toString());
  }
  return { ref: ref, files: files };
}


/* ============================================
   COMPROBACIÓN DE ACTUALIZACIONES
   ============================================ */

/**
 * Comprueba si hay una nueva versión disponible consultando el servidor de versiones.
 * @returns {Object} { hayActualizacion, versionLocal, versionRemota, changelog, ... }
 */
function comprobarActualizaciones() {
  try {
    const serverUrl = getUpdateServerUrl_();
    if (!serverUrl) {
      Logger.log('⚠️ UPDATE_SERVER_URL no configurada.');
      return { hayActualizacion: false, error: 'Servidor de actualizaciones no configurado' };
    }

    const response = UrlFetchApp.fetch(serverUrl, {
      muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' }
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('❌ Error contactando servidor: HTTP ' + response.getResponseCode());
      return { hayActualizacion: false, error: 'Error de conexión al servidor' };
    }

    const datos = JSON.parse(response.getContentText());
    const versionRemota = datos.version;
    const versionLocal = SYSTEM_VERSION;
    const hayActualizacion = compararVersiones_(versionRemota, versionLocal) > 0;

    const resultado = {
      hayActualizacion: hayActualizacion,
      versionLocal: versionLocal,
      versionRemota: versionRemota,
      changelog: datos.changelog || '',
      urlDescarga: datos.urlDescarga || '',
      fechaPublicacion: datos.fechaPublicacion || '',
      critica: datos.critica || false
    };

    // Guardar en propiedades para acceso rápido
    const props = PropertiesService.getScriptProperties();
    props.setProperty('LAST_UPDATE_CHECK', new Date().toISOString());
    props.setProperty('LATEST_REMOTE_VERSION', versionRemota);

    // Guardar el manifiesto remoto (ref + lista de archivos de la versión nueva)
    if (datos.files || datos.ref) {
      props.setProperty('UPDATE_MANIFEST', JSON.stringify({
        ref: datos.ref || 'main',
        files: datos.files || null
      }));
    } else {
      props.deleteProperty('UPDATE_MANIFEST');
    }

    if (hayActualizacion) {
      props.setProperty('UPDATE_AVAILABLE', 'true');
      props.setProperty('UPDATE_CHANGELOG', datos.changelog || '');
      Logger.log(`🔔 Nueva versión: ${versionLocal} → ${versionRemota}`);
    } else {
      props.setProperty('UPDATE_AVAILABLE', 'false');
      Logger.log(`✅ Sistema actualizado (v${versionLocal})`);
    }

    return resultado;

  } catch (e) {
    Logger.log('❌ Error comprobando actualizaciones: ' + e.toString());
    return { hayActualizacion: false, error: e.toString() };
  }
}

/**
 * Compara dos versiones semánticas.
 * @returns {number} 1 si a > b, -1 si a < b, 0 si iguales
 * @private
 */
function compararVersiones_(a, b) {
  const partesA = String(a).split('.').map(Number);
  const partesB = String(b).split('.').map(Number);
  for (let i = 0; i < Math.max(partesA.length, partesB.length); i++) {
    const numA = partesA[i] || 0;
    const numB = partesB[i] || 0;
    if (numA > numB) return 1;
    if (numA < numB) return -1;
  }
  return 0;
}


/* ============================================
   NOTIFICACIONES DE ACTUALIZACIÓN
   ============================================ */

/**
 * Comprueba actualizaciones y notifica al admin si hay nueva versión.
 * Esta función se ejecuta con el trigger diario.
 */
function comprobarYNotificarActualizacion() {
  const resultado = comprobarActualizaciones();
  if (!resultado.hayActualizacion) return;

  try {
    const emailAdmin = obtenerEmailAdmin_();
    if (!emailAdmin) return;

    const prioridadTexto = resultado.critica
      ? '🔴 <strong>ACTUALIZACIÓN CRÍTICA</strong> - Se recomienda aplicar inmediatamente'
      : '🟢 Actualización disponible';

    const asunto = resultado.critica
      ? `🔴 URGENTE: Actualización crítica del Sistema de Reservas (v${resultado.versionRemota})`
      : `🔄 Nueva versión del Sistema de Reservas disponible (v${resultado.versionRemota})`;

    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 25px; text-align: center;">
          <h1 style="color: white; margin: 0; font-size: 22px;">🔄 Actualización Disponible</h1>
          <p style="color: rgba(255,255,255,0.9); margin: 8px 0 0 0;">Sistema de Reservas de Recursos y Salas</p>
        </div>
        <div style="padding: 25px;">
          <table style="width:100%; margin-bottom: 20px;">
            <tr>
              <td style="background: #f5f5f5; padding: 12px 20px; border-radius: 8px; text-align: center;">
                <div style="font-size: 12px; color: #666;">Versión actual</div>
                <div style="font-size: 20px; font-weight: bold; color: #999;">${resultado.versionLocal}</div>
              </td>
              <td style="text-align: center; font-size: 24px; width: 50px;">→</td>
              <td style="background: #e8f5e9; padding: 12px 20px; border-radius: 8px; text-align: center;">
                <div style="font-size: 12px; color: #2e7d32;">Nueva versión</div>
                <div style="font-size: 20px; font-weight: bold; color: #2e7d32;">${resultado.versionRemota}</div>
              </td>
            </tr>
          </table>
          <p>${prioridadTexto}</p>
          ${resultado.changelog ? `
          <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #667eea; margin: 15px 0;">
            <h3 style="margin: 0 0 10px 0; font-size: 14px; color: #333;">📋 Cambios en esta versión:</h3>
            <div style="font-size: 14px; color: #555; white-space: pre-wrap;">${resultado.changelog}</div>
          </div>` : ''}
          <div style="background: #fff3e0; padding: 15px; border-radius: 8px; margin: 15px 0;">
            <h3 style="margin: 0 0 8px 0; font-size: 14px;">📝 Cómo actualizar:</h3>
            <ol style="margin: 0; padding-left: 20px; font-size: 14px; color: #555;">
              <li>Abre la hoja de cálculo del sistema</li>
              <li>Ve al menú <strong>"🗓️ Sistema de Reservas"</strong></li>
              <li>Haz clic en <strong>"🔄 Aplicar actualización"</strong></li>
              <li>Confirma y el código se actualizará automáticamente</li>
              <li>Solo tendrás que <strong>redesplegar</strong> la Web App</li>
            </ol>
          </div>
          ${resultado.fechaPublicacion ? `<p style="font-size: 12px; color: #999;">Publicada: ${resultado.fechaPublicacion}</p>` : ''}
        </div>
      </div>`;

    MailApp.sendEmail({ to: emailAdmin, subject: asunto, htmlBody: cuerpo });
    Logger.log(`📧 Notificación enviada a: ${emailAdmin}`);

  } catch (e) {
    Logger.log('⚠️ Error enviando notificación: ' + e.toString());
  }
}

/**
 * Obtiene el email del admin desde Config o Usuarios.
 * @private
 */
function obtenerEmailAdmin_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (configSheet) {
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'email_admin' && data[i][1]) return data[i][1];
    }
  }

  // Fallback: primer admin activo
  const usSheet = ss.getSheetByName('Usuarios');
  if (usSheet && usSheet.getLastRow() > 1) {
    const usuarios = usSheet.getDataRange().getValues();
    for (let i = 1; i < usuarios.length; i++) {
      if (usuarios[i][3] === true || String(usuarios[i][3]).toUpperCase() === 'TRUE') {
        return usuarios[i][1];
      }
    }
  }

  return null;
}


/* ============================================
   APLICAR ACTUALIZACIÓN (AUTOMÁTICA VÍA API)
   ============================================ */

/**
 * Descarga el código nuevo de GitHub y lo aplica automáticamente
 * usando la API REST de Apps Script.
 *
 * Flujo:
 * 1. Verificar permisos de admin
 * 2. Comprobar que hay actualización
 * 3. Confirmar con el usuario
 * 4. Descargar archivos de GitHub
 * 5. Reescribir los archivos del proyecto vía API
 * 6. Indicar al admin que redepliegue
 */
function aplicarActualizacion() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Verificar admin
  if (!verificarAdminParaUpdate_()) {
    ui.alert('⛔ Acceso denegado', 'Solo los administradores pueden aplicar actualizaciones.', ui.ButtonSet.OK);
    return;
  }

  // 2. Comprobar actualización
  const resultado = comprobarActualizaciones();
  if (!resultado.hayActualizacion) {
    ui.alert('✅ Sistema actualizado',
      `Tu sistema ya está en la última versión (v${SYSTEM_VERSION}).`,
      ui.ButtonSet.OK);
    return;
  }

  // 3. Confirmar
  const confirmacion = ui.alert(
    `🔄 Actualización a v${resultado.versionRemota}`,
    `Versión actual: v${resultado.versionLocal}\n` +
    `Nueva versión: v${resultado.versionRemota}\n\n` +
    `${resultado.changelog ? 'Cambios:\n' + resultado.changelog + '\n\n' : ''}` +
    `El código del proyecto se actualizará automáticamente.\n` +
    `Los datos de tu hoja de cálculo NO se tocarán.\n\n` +
    `Después solo tendrás que crear una nueva implementación\n` +
    `(Implementar → Gestionar implementaciones → Editar → Nueva versión).\n\n` +
    `¿Deseas continuar?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmacion !== ui.Button.YES) return;

  // 4. Descargar y aplicar
  ss.toast('Descargando archivos actualizados...', '🔄 Actualizando', -1);

  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('PRE_UPDATE_VERSION', SYSTEM_VERSION);
    props.setProperty('PRE_UPDATE_DATE', new Date().toISOString());

    // Descargar archivos de GitHub
    const archivosDescargados = descargarArchivosDesdeGitHub_();
    const exitosos = archivosDescargados.filter(a => !a.error);
    const errores = archivosDescargados.filter(a => a.error);

    if (exitosos.length === 0) {
      throw new Error('No se pudo descargar ningún archivo. Comprueba tu conexión a internet.');
    }

    // Intentar actualización automática vía API
    ss.toast('Aplicando actualización vía API...', '🔄 Actualizando', -1);

    const resultadoAPI = actualizarCodigoViaAPI_(exitosos);

    props.setProperty('LAST_UPDATE_APPLIED', new Date().toISOString());
    props.setProperty('UPDATE_AVAILABLE', 'false');

    ss.toast('', '', 1);

    if (resultadoAPI.success) {
      // ✅ Actualización automática exitosa
      ui.alert(
        '✅ ¡Actualización aplicada!',
        `El código se ha actualizado de v${resultado.versionLocal} a v${resultado.versionRemota}.\n\n` +
        `Archivos actualizados: ${exitosos.length}\n` +
        `${errores.length > 0 ? 'Archivos con error: ' + errores.length + '\n' : ''}` +
        `\n⚠️ PASO FINAL NECESARIO:\n` +
        `Para que los cambios surtan efecto en la Web App:\n` +
        `1. Ve a Extensiones → Apps Script\n` +
        `2. Implementar → Gestionar implementaciones\n` +
        `3. Haz clic en el lápiz (editar) de tu implementación activa\n` +
        `4. En "Versión" selecciona "Nueva versión"\n` +
        `5. Haz clic en "Implementar"\n\n` +
        `¡Listo! La Web App ya estará actualizada.`,
        ui.ButtonSet.OK
      );
    } else {
      // ⚠️ La API falló → ofrecer método manual como fallback
      Logger.log('⚠️ API falló: ' + resultadoAPI.error);

      // Guardar archivos en hoja oculta como fallback
      guardarArchivosParaRevision_(archivosDescargados);

      const htmlResumen = generarHtmlFallbackManual_(resultado, exitosos, errores, resultadoAPI.error);
      const output = HtmlService.createHtmlOutput(htmlResumen).setWidth(700).setHeight(550);
      ui.showModalDialog(output, `🔄 Actualización a v${resultado.versionRemota}`);
    }

  } catch (e) {
    ss.toast('', '❌ Error', 3);
    ui.alert('❌ Error en la actualización',
      'Se ha producido un error:\n\n' + e.toString() + '\n\n' +
      'Tus datos NO se han modificado. Puedes intentarlo de nuevo.',
      ui.ButtonSet.OK);
    Logger.log('Error aplicando actualización: ' + e.toString());
  }
}


/* ============================================
   ACTUALIZACIÓN VÍA API DE APPS SCRIPT
   ============================================ */

/**
 * Usa la API REST de Apps Script para reemplazar los archivos del proyecto.
 *
 * Endpoint: PUT https://script.googleapis.com/v1/projects/{scriptId}/content
 *
 * @param {Array} archivos - Archivos descargados con {name, type, scriptName, content}
 * @returns {Object} { success: boolean, error?: string }
 * @private
 */
function actualizarCodigoViaAPI_(archivos) {
  try {
    const scriptId = ScriptApp.getScriptId();
    const token = ScriptApp.getOAuthToken();

    if (!scriptId) {
      return { success: false, error: 'No se pudo obtener el ID del proyecto.' };
    }

    // Construir la lista de archivos en el formato de la API
    const apiFiles = [];

    // Primero: incluir el manifiesto (appsscript.json) - OBLIGATORIO
    // Descargarlo de GitHub también
    try {
      const manifestResp = UrlFetchApp.fetch(GITHUB_REPO_RAW + '/' + getUpdateManifest_().ref + '/appsscript.json', { muteHttpExceptions: true });
      if (manifestResp.getResponseCode() === 200) {
        apiFiles.push({
          name: 'appsscript',
          type: 'JSON',
          source: manifestResp.getContentText()
        });
      } else {
        // Si no se puede descargar, leer el manifiesto actual del proyecto
        const currentResp = UrlFetchApp.fetch(
          `https://script.googleapis.com/v1/projects/${scriptId}/content`,
          { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
        );
        if (currentResp.getResponseCode() === 200) {
          const currentContent = JSON.parse(currentResp.getContentText());
          const currentManifest = currentContent.files.find(f => f.name === 'appsscript');
          if (currentManifest) {
            apiFiles.push({
              name: 'appsscript',
              type: 'JSON',
              source: currentManifest.source
            });
          }
        }
      }
    } catch (e) {
      Logger.log('⚠️ Error obteniendo manifiesto: ' + e.toString());
    }

    // Después: incluir todos los archivos descargados
    for (const archivo of archivos) {
      apiFiles.push({
        name: archivo.scriptName,
        type: archivo.type,
        source: archivo.content
      });
    }

    if (apiFiles.length === 0) {
      return { success: false, error: 'No hay archivos para actualizar.' };
    }

    // Llamar a la API de Apps Script
    const apiUrl = `https://script.googleapis.com/v1/projects/${scriptId}/content`;

    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'put',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ files: apiFiles }),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      Logger.log('✅ Código actualizado vía API. Archivos: ' + apiFiles.length);
      return { success: true };
    } else {
      Logger.log('❌ API respondió con HTTP ' + responseCode + ': ' + responseText);

      // Intentar extraer mensaje de error legible
      let errorMsg = 'HTTP ' + responseCode;
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.error && errorData.error.message) {
          errorMsg = errorData.error.message;
        }
      } catch (parseErr) { /* ignorar */ }

      return { success: false, error: errorMsg };
    }

  } catch (e) {
    Logger.log('❌ Error en actualizarCodigoViaAPI_: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}


/* ============================================
   DESCARGA DE ARCHIVOS
   ============================================ */

/**
 * Descarga los archivos del repositorio de GitHub.
 * @private
 */
function descargarArchivosDesdeGitHub_() {
  const resultados = [];
  const manifest = getUpdateManifest_();

  for (const archivo of manifest.files) {
    try {
      const url = GITHUB_REPO_RAW + '/' + manifest.ref + '/' + archivo.name;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        const content = response.getContentText();
        resultados.push({
          name: archivo.name,
          type: archivo.type,
          scriptName: archivo.scriptName,
          content: content,
          size: content.length,
          error: null
        });
      } else {
        resultados.push({
          name: archivo.name, type: archivo.type, scriptName: archivo.scriptName,
          content: null, error: 'HTTP ' + response.getResponseCode()
        });
      }
    } catch (e) {
      resultados.push({
        name: archivo.name, type: archivo.type, scriptName: archivo.scriptName,
        content: null, error: e.toString()
      });
    }
  }

  return resultados;
}


/* ============================================
   FALLBACK: MÉTODO MANUAL
   ============================================ */

/**
 * Guarda archivos en hoja oculta como fallback si la API falla.
 * @private
 */
function guardarArchivosParaRevision_(archivos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('_ActualizacionPendiente');
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet('_ActualizacionPendiente');
  sheet.hideSheet();

  sheet.getRange(1, 1, 1, 4).setValues([['Archivo', 'Tipo', 'Estado', 'Contenido']]);
  sheet.getRange(1, 1, 1, 4).setBackground('#4f46e5').setFontColor('white').setFontWeight('bold');

  const filas = archivos.map(a => [
    a.name,
    a.type,
    a.error ? '❌ ' + a.error : '✅ Descargado (' + (a.size || 0) + ' bytes)',
    a.content || ''
  ]);

  if (filas.length > 0) {
    sheet.getRange(2, 1, filas.length, 4).setValues(filas);
  }

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 400);
}

/**
 * Genera HTML para el diálogo de fallback manual.
 * Se muestra cuando la API de Apps Script no está disponible.
 * @private
 */
function generarHtmlFallbackManual_(resultado, exitosos, errores, errorAPI) {
  const listaExitosos = exitosos.map(a =>
    `<li style="margin:4px 0;">✅ <code>${a.name}</code> <span style="color:#999;">(${Math.round(a.size / 1024 * 10) / 10} KB)</span></li>`
  ).join('');

  const listaErrores = errores.map(a =>
    `<li style="margin:4px 0;">❌ <code>${a.name}</code>: ${a.error}</li>`
  ).join('');

  return `
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; margin: 0; padding: 20px; color: #333; }
      .header { background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%); color: white; padding: 20px; border-radius: 12px; margin-bottom: 20px; text-align: center; }
      .section { background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 10px 0; }
      .step { background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 12px 15px; margin: 8px 0; display: flex; align-items: start; gap: 12px; }
      .step-num { background: #ff9800; color: white; width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 14px; flex-shrink: 0; }
      code { background: #e8eaf6; padding: 2px 6px; border-radius: 4px; font-size: 13px; }
      .btn { background: #ff9800; color: white; border: none; padding: 10px 24px; border-radius: 8px; font-size: 14px; cursor: pointer; }
      .btn:hover { background: #f57c00; }
      ul { list-style: none; padding: 0; }
    </style>

    <div class="header">
      <h2 style="margin: 0;">⚠️ Actualización manual necesaria</h2>
      <p style="margin: 8px 0 0 0; opacity: 0.9;">v${resultado.versionLocal} → v${resultado.versionRemota}</p>
    </div>

    <div style="background: #fff3e0; padding: 12px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #ff9800;">
      <strong>La actualización automática no pudo completarse:</strong><br>
      <span style="font-size: 13px; color: #666;">${errorAPI}</span><br><br>
      <strong>Para activar las actualizaciones automáticas:</strong><br>
      <span style="font-size: 13px;">Ve a <a href="https://script.google.com/home/usersettings" target="_blank">script.google.com/home/usersettings</a> y activa la <strong>API de Google Apps Script</strong>.</span>
    </div>

    <div class="section">
      <strong>📦 Archivos descargados (${exitosos.length}/${exitosos.length + errores.length}):</strong>
      <ul>${listaExitosos}</ul>
      ${errores.length > 0 ? `<strong>⚠️ Errores:</strong><ul>${listaErrores}</ul>` : ''}
    </div>

    <h3>📝 Actualización manual:</h3>

    <div class="step">
      <div class="step-num">1</div>
      <div>
        <strong>Abre el Editor de Apps Script</strong><br>
        <span style="color:#666;"><strong>Extensiones → Apps Script</strong></span>
      </div>
    </div>

    <div class="step">
      <div class="step-num">2</div>
      <div>
        <strong>Reemplaza cada archivo</strong><br>
        <span style="color:#666;">Los archivos están en la hoja oculta <code>_ActualizacionPendiente</code>.
        Muéstrala temporalmente, copia el contenido de la columna D y pégalo en cada archivo del editor.</span>
      </div>
    </div>

    <div class="step">
      <div class="step-num">3</div>
      <div>
        <strong>Guarda y despliega</strong><br>
        <span style="color:#666;">Guarda (Ctrl+S) y crea una nueva implementación.</span>
      </div>
    </div>

    <div style="text-align: center; margin-top: 20px;">
      <button class="btn" onclick="google.script.host.close()">Entendido</button>
    </div>`;
}

/**
 * Verifica que el usuario actual es admin.
 * @private
 */
function verificarAdminParaUpdate_() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usSheet = ss.getSheetByName('Usuarios');

  if (!usSheet || usSheet.getLastRow() <= 1) return false;

  const usuarios = usSheet.getDataRange().getValues();
  for (let i = 1; i < usuarios.length; i++) {
    if (String(usuarios[i][1]).toLowerCase() === email.toLowerCase() &&
      (usuarios[i][3] === true || String(usuarios[i][3]).toUpperCase() === 'TRUE')) {
      return true;
    }
  }
  return false;
}


/* ============================================
   GESTIÓN DE TRIGGERS AUTOMÁTICOS
   ============================================ */

/**
 * Instala un trigger diario para comprobar actualizaciones.
 * Se llama durante el setup inicial.
 */
function instalarTriggerActualizaciones() {
  eliminarTriggerActualizaciones_();

  ScriptApp.newTrigger('comprobarYNotificarActualizacion')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('✅ Trigger de actualizaciones instalado (diario a las 8:00)');
}

/**
 * Elimina triggers de comprobación de actualizaciones.
 * @private
 */
function eliminarTriggerActualizaciones_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'comprobarYNotificarActualizacion') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Desinstala el trigger de actualizaciones.
 */
function desinstalarTriggerActualizaciones() {
  eliminarTriggerActualizaciones_();
  Logger.log('🔕 Trigger de actualizaciones eliminado');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Las comprobaciones automáticas han sido desactivadas.',
    '🔕 Actualizaciones desactivadas', 5
  );
}


/* ============================================
   FUNCIONES DE MENÚ E INFORMACIÓN
   ============================================ */

/**
 * Muestra información de la versión actual.
 */
function mostrarInfoVersion() {
  const props = PropertiesService.getScriptProperties();
  const ultimaComprobacion = props.getProperty('LAST_UPDATE_CHECK') || 'Nunca';
  const updateDisponible = props.getProperty('UPDATE_AVAILABLE') === 'true';
  const versionRemota = props.getProperty('LATEST_REMOTE_VERSION') || '?';

  let mensaje = `📋 Información del Sistema\n\n`;
  mensaje += `Versión instalada: v${SYSTEM_VERSION}\n`;
  mensaje += `Última comprobación: ${ultimaComprobacion}\n\n`;

  if (updateDisponible) {
    mensaje += `🔔 ¡Nueva versión disponible! v${versionRemota}\n`;
    mensaje += `Ve al menú "🔄 Aplicar actualización" para actualizar.`;
  } else {
    mensaje += `✅ Tu sistema está actualizado.`;
  }

  SpreadsheetApp.getUi().alert('ℹ️ Versión del Sistema', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Comprobación manual de actualizaciones con feedback visual.
 */
function comprobarActualizacionesManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Contactando con el servidor de versiones...', '🔄 Comprobando...', -1);

  const resultado = comprobarActualizaciones();

  if (resultado.error) {
    ss.toast('', '', 1);
    SpreadsheetApp.getUi().alert('⚠️ Error',
      'No se pudo comprobar: ' + resultado.error,
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  ss.toast('', '', 1);

  if (resultado.hayActualizacion) {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      `🔔 ¡Nueva versión v${resultado.versionRemota}!`,
      `Tu versión: v${resultado.versionLocal}\n` +
      `Nueva versión: v${resultado.versionRemota}\n\n` +
      `${resultado.changelog ? 'Cambios:\n' + resultado.changelog + '\n\n' : ''}` +
      `¿Quieres aplicar la actualización ahora?`,
      ui.ButtonSet.YES_NO
    );
    if (resp === ui.Button.YES) aplicarActualizacion();
  } else {
    SpreadsheetApp.getUi().alert('✅ Todo al día',
      `Tu sistema está en la última versión (v${SYSTEM_VERSION}).`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Devuelve la versión actual (para el frontend).
 */
function getSystemVersion() {
  return SYSTEM_VERSION;
}

/**
 * Devuelve el estado de actualización (para el frontend).
 */
function getUpdateStatus() {
  const props = PropertiesService.getScriptProperties();
  return {
    version: SYSTEM_VERSION,
    updateAvailable: props.getProperty('UPDATE_AVAILABLE') === 'true',
    latestVersion: props.getProperty('LATEST_REMOTE_VERSION') || SYSTEM_VERSION,
    lastCheck: props.getProperty('LAST_UPDATE_CHECK') || null
  };
}

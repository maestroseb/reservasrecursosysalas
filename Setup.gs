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
  'Reservas': { headers: ['ID_Reserva', 'ID_Recurso', 'Email_Usuario', 'Fecha', 'Curso', 'ID_Tramo', 'Cantidad', 'Estado', 'Notas', 'Timestamp'], color: '#059669' }
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

    // 2. Rellenar datos y GUARDAR LA URL EN CONFIG (CORREGIDO: ahora pasa currentUrl)
    crearDatosEjemploInterno(ss, currentUrl);

    // 3. Hacer Admin al usuario que está ejecutando esto
    asegurarAdminInterno(ss, currentUser);

    // 4. Marcar como instalado (IMPORTANTE: usar SETUP_COMPLETED para consistencia)
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SETUP_COMPLETED', 'true');
    props.setProperty('INSTALL_DATE', new Date().toISOString());
    props.setProperty('WEB_APP_URL', currentUrl);
    props.setProperty('FECHA_ACTIVACION', new Date().toISOString());

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

// --- 🌟 FUNCIÓN CORREGIDA: Ahora recibe currentUrl como parámetro ---
function crearDatosEjemploInterno(ss, currentUrl) {
  
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

  // D. CONFIGURACIÓN (Ahora con currentUrl correctamente definido)
  const sheetConfig = ss.getSheetByName('Config');
  const configData = [
    ['dias_vista_maximo', 30, 'Días a futuro permitidos'],
    ['minutos_antelacion', 0, 'Minutos mínimos antes de reservar'],
    ['limite_reservas', 3, 'Máx. reservas activas por usuario'],
    ['horas_cancelacion', 0, 'Horas mínimas para poder cancelar solo'],
    ['exigir_motivo', 'FALSE', 'Obligatorio escribir para qué es'],
    ['email_admin', '', 'Email copia oculta (vacío = desactivado)'],
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
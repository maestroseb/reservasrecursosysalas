/** 
 * ADMIN FUNCTIONS
 * Funciones exclusivas para administradores del sistema
 */

/* ============================================
   VERIFICACIÓN DE PERMISOS
   ============================================ */

// Interpreta casillas / textos de la hoja: TRUE, "true", "Si", "Sí", "yes"
function esValorVerdadero_(v) {
  if (v === true) return true;
  const t = String(v).toLowerCase().trim().replace('í', 'i');
  return t === 'true' || t === 'si' || t === 'yes';
}

// Ejecuta fn con el bloqueo global del script (mismo que usa crearNuevaReserva)
function conLockScript_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return { success: false, error: "El sistema está ocupado. Inténtalo de nuevo en unos segundos." };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function isUserAdmin() {
  if (ADMIN_POR_ENLACE_FIRMADO_) return true; // solo durante una aprobación desde enlace firmado
  const email = emailActual_();
  const authResult = checkUserAuthorization(email);
  return authResult.isAdmin || false;
}



/* ============================================
   CARGA DE DATOS ADMIN (V9 - FIX FINAL DEFINITIVO)
   ============================================ */

function getAdminData() {
  try {
    const email = emailActual_();
    const authResult = checkUserAuthorization(email);

    if (!authResult || !authResult.isAdmin) {
      return { success: false, error: "No tienes permisos de administrador." };
    }

    // Migración de ID_Solicitud_Recurrente: solo hasta que se complete una vez
    // (antes se leía toda la hoja Reservas en cada apertura del panel)
    try {
      const props = PropertiesService.getScriptProperties();
      if (props.getProperty('MIGRACION_ID_SOL_REC_OK') !== 'true') {
        const resMig = conLockScript_(() => migrarIdSolicitudRecurrente_());
        if (resMig && resMig.success) props.setProperty('MIGRACION_ID_SOL_REC_OK', 'true');
      }
    } catch (migErr) {
      Logger.log('⚠️ Error en migración (no crítico): ' + migErr.message);
    }

    const ss = getDB();

    // 1. RECURSOS
    const sheetRec = ss.getSheetByName(SHEETS.RECURSOS);
    let recursos = [];
    if (sheetRec && sheetRec.getLastRow() > 1) {
      const data = sheetRec.getRange(2, 1, sheetRec.getLastRow()-1, 8).getValues();
      recursos = data.map(r => ({
        id_recurso: String(r[0]).trim(),
        nombre: String(r[1]).trim(),
        tipo: String(r[2]).trim(),
        icono: String(r[3]).trim(),
        ubicacion: String(r[4]).trim(),
        capacidad: r[5],
        descripcion: String(r[6]).trim(),
        estado: String(r[7]).trim(),
      }));
    }

    // 2. TRAMOS
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    let tramos = [];
    if (sheetTramos && sheetTramos.getLastRow() > 1) {
      const data = sheetTramos.getRange(2, 1, sheetTramos.getLastRow()-1, 4).getDisplayValues();
      tramos = data.map(r => ({
        id_tramo: String(r[0]).trim(),
        Nombre_Tramo: String(r[1]).trim(),
        hora_inicio: String(r[2]).trim(), 
        hora_fin: String(r[3]).trim()
      }));
    }

    // 3. DISPONIBILIDAD
    const sheetDisp = ss.getSheetByName(SHEETS.DISPONIBILIDAD);
    let disponibilidad = [];
    if (sheetDisp && sheetDisp.getLastRow() > 1) {
      const data = sheetDisp.getRange(2, 1, sheetDisp.getLastRow() - 1, 6).getValues();
      disponibilidad = data.map(row => {
        let diaRaw = String(row[1]).trim();
        let diaFinal = diaRaw;
        if (diaRaw.includes("Lunes")) diaFinal = "Lunes";
        else if (diaRaw.includes("Martes")) diaFinal = "Martes";
        else if (diaRaw.includes("Miércoles") || diaRaw.includes("Miercoles")) diaFinal = "Miércoles";
        else if (diaRaw.includes("Jueves")) diaFinal = "Jueves";
        else if (diaRaw.includes("Viernes")) diaFinal = "Viernes";

        let permRaw = String(row[4]).trim().toLowerCase();
        let permFinal = 'Si';
        if (permRaw === 'no' || permRaw === 'false' || permRaw === '0') permFinal = 'No';

        return {
          id_recurso: String(row[0]).trim(),
          dia_semana: diaFinal,
          id_tramo: String(row[2]).trim(),
          permitido: permFinal,
          razon_bloqueo: String(row[5]).trim()
        };
      });
    }

    // 4. USUARIOS
    const sheetUsers = ss.getSheetByName(SHEETS.USUARIOS);
    let usuarios = [];
    if (sheetUsers && sheetUsers.getLastRow() > 1) {
      const dataUsers = sheetUsers.getRange(2, 1, sheetUsers.getLastRow() - 1, 4).getValues();
      usuarios = dataUsers.map(row => ({
        Nombre: String(row[0]).trim(), 
        Email: String(row[1]).trim(), 
        Activo: esValorVerdadero_(row[2]), // Boolean('FALSE') era true y reactivaba usuarios
        Admin: esValorVerdadero_(row[3])
      }));
    }

    // 5. RESERVAS (con soporte para id_solicitud_recurrente)
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);
    let reservas = [];
    if (sheetReservas && sheetReservas.getLastRow() > 1) {
      // Leer todas las columnas dinámicamente
      const lastCol = sheetReservas.getLastColumn();
      const headers = sheetReservas.getRange(1, 1, 1, lastCol).getValues()[0];
      const headerMap = {};
      headers.forEach((h, i) => { headerMap[h.toString().toLowerCase().trim()] = i; });

      const dataReservas = sheetReservas.getRange(2, 1, sheetReservas.getLastRow() - 1, lastCol).getValues();
      reservas = dataReservas.map(row => {
        let fechaStr = '';
        const fechaCol = headerMap['fecha'] !== undefined ? headerMap['fecha'] : 3;
        try { fechaStr = Utilities.formatDate(new Date(row[fechaCol]), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch (e) { fechaStr = String(row[fechaCol]); }

        // Obtener id_solicitud_recurrente si existe
        const idSolRecCol = headerMap['id_solicitud_recurrente'];
        const idSolicitudRecurrente = idSolRecCol !== undefined ? String(row[idSolRecCol] || '').trim() : '';

        return {
          ID_Reserva: String(row[headerMap['id_reserva'] || 0]),
          ID_Recurso: String(row[headerMap['id_recurso'] || 1]).trim(),
          Email_Usuario: String(row[headerMap['email_usuario'] || 2]).trim(),
          Fecha: fechaStr,
          Curso: String(row[headerMap['curso'] || 4]),
          ID_Tramo: String(row[headerMap['id_tramo'] || 5]).trim(),
          Cantidad: Number(row[headerMap['cantidad'] || 6] || 1),
          Estado: String(row[headerMap['estado'] || 7]),
          Notas: String(row[headerMap['notas'] || 8] || ''),
          ID_Solicitud_Recurrente: idSolicitudRecurrente
        };
      });
    }

    // 6. CURSOS
    const sheetCursos = ss.getSheetByName(SHEETS.CURSOS);
    let cursosData = { cursos: [], modoVisualizacion: 'botones' };
    if (sheetCursos) {
      const modoViz = sheetCursos.getRange('D1').getValue();
      cursosData.modoVisualizacion = modoViz || 'botones';
      try {
        const dataC = sheetCursos.getRange(2, 1, sheetCursos.getLastRow()-1, 2).getValues(); // ✅ Solo 2 columnas
        cursosData.cursos = dataC.map(r => ({ etapa: r[0], curso: r[1] })); // ✅ Simple
      } catch(e) { cursosData.cursos = []; }
    }

    // 7. CONFIGURACIÓN (NUEVO BLOQUE ⚙️)
    // Buscamos la hoja por nombre 'CONFIG' (o usa SHEETS.CONFIG si la definiste en constantes)
    const sheetConfig = ss.getSheetByName('CONFIG'); 
    let configData = [];
    
    if (sheetConfig && sheetConfig.getLastRow() > 1) {
      // Leemos todo el rango de datos
      const dataConfig = sheetConfig.getDataRange().getValues();
      
      // Iteramos desde la fila 1 (índice 1) para saltar la cabecera
      for (let i = 1; i < dataConfig.length; i++) {
        // Columna A (0) es la CLAVE, Columna B (1) es el VALOR
        if (dataConfig[i][0]) {
          configData.push({
            clave: String(dataConfig[i][0]).trim(),
            valor: dataConfig[i][1] // No usamos String() aquí para respetar Booleans y Numbers
          });
        }
      }
    }

    return {
      success: true,
      recursos: recursos, 
      tramos: tramos, 
      disponibilidad: disponibilidad,
      usuarios: usuarios, 
      reservas: reservas, 
      cursos: cursosData.cursos,
      modoVisualizacion: cursosData.modoVisualizacion,
      config: configData // <--- ¡AQUÍ ENVIAMOS LA CONFIGURACIÓN AL FRONTEND!
    };

  } catch (e) { return { success: false, error: "Error getAdminData: " + e.toString() }; }
}

/* ===========================================
   GUARDADO TOTAL CURSOS (V3 - REEMPLAZO)
   =========================================== */
function saveAllCursos(data) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.CURSOS);
    
    // 1. Guardar preferencia visualización (Celda D1)
    // El cambio es añadir "|| 'botones'" para que nunca quede vacío
    const modoAGuardar = data.modoVisualizacion || 'botones'; 
    sheet.getRange('D1').setValue(modoAGuardar);
    
    // 2. Limpiar hoja (excepto cabecera)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
    }
    
    // 3. Escribir nuevos datos (Si hay)
    // Mapeamos: Etapa, Curso
    if (data.cursos && data.cursos.length > 0) {
      const filas = data.cursos.map((c, i) => [
        c.etapa || '',
        c.curso || '',
      ]);
      
      sheet.getRange(2, 1, filas.length, 2).setValues(filas);
    }
    
    purgarCache();
    return { success: true };
    
  } catch (e) { return { success: false, error: e.toString() }; }
}


/* ===========================================
   GUARDADO TOTAL RECURSOS (BATCH) 📦
   =========================================== */
function saveBatchRecursos(recursos) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.RECURSOS);
    
    // 1. Limpiar contenido actual (manteniendo cabeceras fila 1)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    }
    
    // 2. Preparar datos para volcar (Mapeo exacto de columnas A-H)
    if (recursos && recursos.length > 0) {
      const filas = recursos.map(r => [
        r.id_recurso || '',      // A
        r.nombre || '',          // B
        r.tipo || 'Sala',        // C
        r.icono || '',           // D
        r.ubicacion || '',       // E
        r.capacidad || 1,        // F
        r.descripcion || '',     // G
        r.estado || 'Activo',    // H
      ]);
      
      // 3. Escribir de golpe (Mucho más rápido)
      sheet.getRange(2, 1, filas.length, 8).setValues(filas);
    }
    
    purgarCache(); // Limpiar caché para que la app móvil se actualice
    return { success: true };
    
  } catch (e) { 
    return { success: false, error: "Error al guardar recursos: " + e.toString() }; 
  }
}

/* ===========================================
   GUARDADO TRAMOS CON EFECTO CASCADA 🌊
   Actualiza referencias en Disponibilidad
   =========================================== */
function saveBatchTramos(listaTramos) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    
    const ss = getDB();
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    const sheetDisp = ss.getSheetByName(SHEETS.DISPONIBILIDAD);
    
    // --- 1. ACTUALIZACIÓN EN CASCADA (DISPONIBILIDAD) ---
    // Buscamos tramos que hayan cambiado de ID
    listaTramos.forEach(t => {
      // Si tiene un ID original guardado Y es diferente al nuevo
      if (t._originalId && t._originalId !== t.id_tramo) {
        
        // Usamos TextFinder para buscar y reemplazar en la hoja Disponibilidad
        // matchEntireCell(true) es vital para no reemplazar "T-001" dentro de "T-0010"
        sheetDisp.createTextFinder(t._originalId)
                 .matchEntireCell(true)
                 .replaceAllWith(t.id_tramo);
                 
        Logger.log(`🔄 Cascada: ${t._originalId} -> ${t.id_tramo}`);
      }
    });

    // --- 2. GUARDADO NORMAL DE TRAMOS ---
    const lastRow = sheetTramos.getLastRow();
    if (lastRow > 1) {
      sheetTramos.getRange(2, 1, lastRow - 1, 4).clearContent();
    }
    
    if (listaTramos && listaTramos.length > 0) {
      const filas = listaTramos.map(t => [
        t.id_tramo,
        t.Nombre_Tramo || t.nombre,
        t.hora_inicio,
        t.hora_fin
      ]);
      
      sheetTramos.getRange(2, 1, filas.length, 4).setValues(filas);
    }
    
    purgarCache();
    return { success: true, message: "Tramos guardados y referencias actualizadas." };
    
  } catch (e) { return { success: false, error: e.toString() }; }
}

/* ===========================================
   GUARDADO MASIVO DISPONIBILIDAD (RESPETANDO ESTRUCTURA)
   =========================================== */
function saveBatchDisponibilidad(cambios) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    var ss = getDB();
    var sheet = ss.getSheetByName(SHEETS.DISPONIBILIDAD);
    
    var data = sheet.getDataRange().getValues();
    var mapaFilas = {};
    
    // Mapa para encontrar filas: ID_Recurso|Dia|ID_Tramo
    for (var i = 1; i < data.length; i++) {
      // Usamos String y trim para asegurar coincidencia
      var key = String(data[i][0]).trim() + "|" + String(data[i][1]).trim() + "|" + String(data[i][2]).trim();
      // Normalizamos día también en la clave del mapa
      if (key.includes("Lunes")) key = key.replace(/.*Lunes.*/, "Lunes"); // Simplificación rápida para el mapa
      // Nota: Mejor usar la clave tal cual viene del excel si la lectura ya normalizó
      
      // REHACEMOS EL MAPA MÁS SIMPLE: Clave exacta tal cual está en la hoja
      // Pero como vamos a escribir, necesitamos encontrar la fila exacta.
      // La clave debe coincidir con lo que envía el frontend.
      
      // ESTRATEGIA SEGURA: Usar los datos "crudos" de las columnas clave A,B,C para el mapa
      // Y asumir que el frontend envía los datos normalizados.
      // Para que coincidan, limpiamos los datos de la hoja al crear la clave.
      
      let dRaw = String(data[i][1]).trim();
      let diaKey = dRaw;
      if (dRaw.includes("Lunes")) diaKey = "Lunes";
      else if (dRaw.includes("Martes")) diaKey = "Martes";
      else if (dRaw.includes("Miércoles") || dRaw.includes("Miercoles")) diaKey = "Miércoles";
      else if (dRaw.includes("Jueves")) diaKey = "Jueves";
      else if (dRaw.includes("Viernes")) diaKey = "Viernes";
      
      let keySegura = String(data[i][0]).trim() + "|" + diaKey + "|" + String(data[i][2]).trim();
      mapaFilas[keySegura] = i + 1; 
    }
    
    var nuevasFilas = [];
    
    for (var j = 0; j < cambios.length; j++) {
      var c = cambios[j];
      var key = String(c.id_recurso).trim() + "|" + String(c.dia_semana).trim() + "|" + String(c.id_tramo).trim();
      var fila = mapaFilas[key];
      
      if (fila) {
        // EXISTE: Actualizamos Columna E (5) y F (6)
        sheet.getRange(fila, 5).setValue(c.permitido);
        sheet.getRange(fila, 6).setValue(c.razon_bloqueo || "");
      } else {
        // NO EXISTE: Creamos nueva fila
        // A=ID, B=Dia, C=Tramo, D=VACIO(Hora), E=Permitido, F=Razon
        nuevasFilas.push([
          c.id_recurso, 
          c.dia_semana, 
          c.id_tramo, 
          "", // Dejamos la hora vacía (el frontend la coge de Tramos)
          c.permitido, 
          c.razon_bloqueo || ""
        ]);
      }
    }
    
    if (nuevasFilas.length > 0) {
      anadirFilas_(sheet, nuevasFilas);
    }
    
    purgarCache();
    return { success: true };

  } catch (e) { return { success: false, error: e.toString() }; }
}

/* ===========================================
   DETECTAR CONFLICTOS CON RECURRENCIAS AL CAMBIAR DISPONIBILIDAD
   =========================================== */

/**
 * Detecta si los cambios de disponibilidad afectan a recurrencias aprobadas
 * @param {Array} cambios - Lista de cambios [{id_recurso, dia_semana, id_tramo, permitido, ...}]
 * @returns {Object} - {tieneConflictos: boolean, afectados: Array}
 */
function detectarConflictosDisponibilidad(cambios) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");

    // Filtrar solo los cambios que BLOQUEAN (de Sí a No)
    const bloqueos = cambios.filter(c =>
      c.permitido === 'No' || c.permitido === false || c.permitido === 'FALSE'
    );

    if (bloqueos.length === 0) {
      return { tieneConflictos: false, afectados: [] };
    }

    // Obtener recurrencias aprobadas
    const sheetRecurrentes = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheetRecurrentes.getDataRange().getValues();

    if (data.length <= 1) {
      return { tieneConflictos: false, afectados: [] };
    }

    // Mapeo de días: nombre completo -> letra
    const diaALetra = {
      'Lunes': 'L', 'Martes': 'M', 'Miércoles': 'X', 'Miercoles': 'X',
      'Jueves': 'J', 'Viernes': 'V'
    };

    const afectados = [];
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    // Revisar cada recurrencia aprobada
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const estado = String(row[COLS_SOLICITUDES.ESTADO] || '').toLowerCase();

      if (estado !== 'aprobada') continue;

      const idRecurso = String(row[COLS_SOLICITUDES.ID_RECURSO] || '').trim();
      const diasSemana = String(row[COLS_SOLICITUDES.DIAS_SEMANA] || '');
      const emailUsuario = row[COLS_SOLICITUDES.EMAIL_USUARIO] || '';
      const nombreUsuario = row[COLS_SOLICITUDES.NOMBRE_USUARIO] || emailUsuario;
      const nombreRecurso = row[COLS_SOLICITUDES.NOMBRE_RECURSO] || idRecurso;
      const fechaFin = row[COLS_SOLICITUDES.FECHA_FIN];

      // Verificar que la recurrencia siga activa
      if (fechaFin) {
        const fechaFinDate = new Date(fechaFin);
        if (fechaFinDate < hoy) continue; // Ya terminó
      }

      // Parsear días:tramos del formato "L:T001,M:T002,X:T003"
      const tramosRecurrencia = [];
      diasSemana.split(',').forEach(entry => {
        const [dia, tramo] = entry.trim().split(':');
        if (dia && tramo) {
          tramosRecurrencia.push({ dia: dia.trim(), tramo: tramo.trim() });
        }
      });

      // Verificar cada bloqueo contra esta recurrencia
      for (const bloqueo of bloqueos) {
        if (String(bloqueo.id_recurso).trim() !== idRecurso) continue;

        const letraDia = diaALetra[bloqueo.dia_semana] || bloqueo.dia_semana;
        const idTramo = String(bloqueo.id_tramo).trim();

        // Buscar si este día:tramo está en la recurrencia
        const afectado = tramosRecurrencia.find(t =>
          t.dia === letraDia && t.tramo === idTramo
        );

        if (afectado) {
          afectados.push({
            id_solicitud: row[COLS_SOLICITUDES.ID_SOLICITUD],
            id_recurso: idRecurso,
            email_usuario: emailUsuario,
            nombre_usuario: nombreUsuario,
            nombre_recurso: nombreRecurso,
            dia_semana: bloqueo.dia_semana,
            id_tramo: idTramo,
            dia_letra: letraDia
          });
        }
      }
    }

    return {
      tieneConflictos: afectados.length > 0,
      afectados: afectados
    };

  } catch (e) {
    Logger.log('Error en detectarConflictosDisponibilidad: ' + e.message);
    return { tieneConflictos: false, afectados: [], error: e.message };
  }
}

/**
 * Guarda disponibilidad con opción de forzar (cancela recurrencias afectadas)
 * @param {Array} cambios - Lista de cambios de disponibilidad
 * @param {boolean} forzar - Si true, cancela los tramos afectados de las recurrencias
 * @returns {Object}
 */
function saveBatchDisponibilidadConValidacion(cambios, forzar = false) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");

    // Primero detectar conflictos
    const conflictos = detectarConflictosDisponibilidad(cambios);

    if (conflictos.tieneConflictos && !forzar) {
      // Hay conflictos y no se fuerza: devolver para confirmación
      return {
        success: false,
        requiereConfirmacion: true,
        afectados: conflictos.afectados
      };
    }

    // Si hay conflictos y se fuerza, procesar las cancelaciones
    if (conflictos.tieneConflictos && forzar) {
      const resultadoCancelaciones = procesarCancelacionesPorDisponibilidad_(conflictos.afectados);
      if (!resultadoCancelaciones.success) {
        return resultadoCancelaciones;
      }
    }

    // Guardar los cambios de disponibilidad
    return saveBatchDisponibilidad(cambios);

  } catch (e) {
    Logger.log('Error en saveBatchDisponibilidadConValidacion: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Procesa las cancelaciones de tramos en recurrencias afectadas
 * @param {Array} afectados - Lista de tramos afectados
 * @returns {Object}
 */
function procesarCancelacionesPorDisponibilidad_(afectados) {
  try {
    const sheetRecurrentes = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheetRecurrentes.getDataRange().getValues();

    // Agrupar afectados por solicitud
    const porSolicitud = {};
    afectados.forEach(a => {
      if (!porSolicitud[a.id_solicitud]) {
        porSolicitud[a.id_solicitud] = {
          email: a.email_usuario,
          nombre: a.nombre_usuario,
          recurso: a.nombre_recurso,
          tramosAfectados: []
        };
      }
      porSolicitud[a.id_solicitud].tramosAfectados.push({
        dia: a.dia_semana,
        dia_letra: a.dia_letra,
        tramo: a.id_tramo
      });
    });

    // Procesar cada solicitud afectada
    for (let i = 1; i < data.length; i++) {
      const idSolicitud = data[i][COLS_SOLICITUDES.ID_SOLICITUD];

      if (!porSolicitud[idSolicitud]) continue;

      const diasActuales = String(data[i][COLS_SOLICITUDES.DIAS_SEMANA] || '');
      const tramosAfectados = porSolicitud[idSolicitud].tramosAfectados;

      // Eliminar los tramos afectados
      const tramosList = diasActuales.split(',').map(e => e.trim()).filter(e => e);
      const tramosRestantes = tramosList.filter(entry => {
        const [dia, tramo] = entry.split(':');
        // Mantener si NO está en los afectados
        return !tramosAfectados.some(t => t.dia_letra === dia && t.tramo === tramo);
      });

      const fila = i + 1;

      if (tramosRestantes.length === 0) {
        // Sin tramos restantes: marcar como cancelada
        sheetRecurrentes.getRange(fila, COLS_SOLICITUDES.ESTADO + 1).setValue('cancelada');
        sheetRecurrentes.getRange(fila, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue(
          'Cancelada automáticamente por cambio de disponibilidad del recurso'
        );
      } else {
        // Actualizar con los tramos restantes
        sheetRecurrentes.getRange(fila, COLS_SOLICITUDES.DIAS_SEMANA + 1).setValue(tramosRestantes.join(','));

        // Añadir nota
        const notaActual = data[i][COLS_SOLICITUDES.NOTAS_ADMIN] || '';
        const nuevaNota = notaActual + (notaActual ? '\n' : '') +
          `Tramos cancelados por disponibilidad: ${tramosAfectados.map(t => t.dia_letra + ':' + t.tramo).join(', ')}`;
        sheetRecurrentes.getRange(fila, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue(nuevaNota);
      }

      // Enviar notificación al usuario
      enviarNotificacionCancelacionDisponibilidad_(porSolicitud[idSolicitud]);
    }

    // También cancelar las reservas individuales futuras de esos tramos
    cancelarReservasFuturasPorDisponibilidad_(afectados);

    return { success: true };

  } catch (e) {
    Logger.log('Error en procesarCancelacionesPorDisponibilidad_: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Cancela reservas futuras individuales afectadas por cambio de disponibilidad
 */
function cancelarReservasFuturasPorDisponibilidad_(afectados) {
  try {
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);
    if (!sheetReservas) return;

    const data = sheetReservas.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());

    const colIdRecurso = headers.indexOf('id_recurso');
    const colFecha = headers.indexOf('fecha');
    const colIdTramo = headers.indexOf('id_tramo');
    const colEstado = headers.indexOf('estado');
    // La columna 'tipo_reserva' no existe en Reservas: se identifica la recurrencia por su ID
    const colIdSolRec = headers.indexOf('id_solicitud_recurrente');

    if (colIdRecurso < 0 || colFecha < 0 || colIdTramo < 0 || colEstado < 0 || colIdSolRec < 0) return;

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    // Mapeo de día de semana JS a letra
    const jsToLetra = { 0: 'D', 1: 'L', 2: 'M', 3: 'X', 4: 'J', 5: 'V', 6: 'S' };

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const estado = String(row[colEstado] || '').toLowerCase().trim();

      // Aceptar tanto 'activa' como 'confirmada'
      if (estado !== 'activa' && estado !== 'confirmada') continue;

      const idSolRec = String(row[colIdSolRec] || '').trim();
      if (!idSolRec) continue; // Solo reservas generadas por una recurrencia

      const fechaReserva = new Date(row[colFecha]);
      if (isNaN(fechaReserva.getTime())) continue;
      if (fechaReserva < hoy) continue;

      const idRecurso = String(row[colIdRecurso]).trim();
      const idTramo = String(row[colIdTramo]).trim();
      const diaSemana = jsToLetra[fechaReserva.getDay()];

      // Verificar si está afectada (debe coincidir recurso, día Y tramo)
      const esAfectada = afectados.some(a =>
        String(a.id_solicitud || '').trim() === idSolRec &&
        String(a.id_recurso || '').trim() === idRecurso &&
        a.dia_letra === diaSemana &&
        String(a.id_tramo).trim() === idTramo
      );

      if (esAfectada) {
        // Cancelar la reserva
        sheetReservas.getRange(i + 1, colEstado + 1).setValue('Cancelada');
      }
    }

  } catch (e) {
    Logger.log('Error en cancelarReservasFuturasPorDisponibilidad_: ' + e.message);
  }
}

/**
 * Envía notificación al usuario sobre cancelación por disponibilidad
 */
function enviarNotificacionCancelacionDisponibilidad_(info) {
  try {
    const email = info.email;
    const nombre = info.nombre || email;
    const recurso = info.recurso;
    const tramos = info.tramosAfectados;

    const tramosTexto = tramos.map(t => `${t.dia} (${t.tramo})`).join(', ');

    const asunto = `[Reservas] Cambio en tu reserva recurrente de ${recurso}`;

    const cuerpo = `
Hola ${nombre},

Te informamos que el administrador ha modificado la disponibilidad del recurso "${recurso}".

Como consecuencia, los siguientes tramos de tu reserva recurrente han sido cancelados:
${tramosTexto}

Las reservas futuras de esos tramos ya no están activas.

Si tienes alguna duda, contacta con el administrador.

Saludos,
Sistema de Reservas
    `.trim();

    MailApp.sendEmail({
      to: email,
      subject: asunto,
      body: cuerpo
    });

    Logger.log(`Email enviado a ${email} sobre cancelación de tramos`);

  } catch (e) {
    Logger.log('Error enviando email de cancelación: ' + e.message);
  }
}

/* ===========================================
   GUARDADO DE USUARIOS (Batch Save)
   =========================================== */
function saveBatchUsuarios(usuariosList) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
    var ss = getDB();
    var sheet = ss.getSheetByName(SHEETS.USUARIOS);
    var lastRow = sheet.getLastRow();

    // Conservar la Especialidad (columna E) de cada usuario por su email:
    // antes solo se reescribían A-D y la columna E quedaba desplazada al borrar usuarios.
    var especialidadPorEmail = {};
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 5).getValues().forEach(function (r) {
        var em = String(r[1]).toLowerCase().trim();
        if (em) especialidadPorEmail[em] = r[4];
      });
    }
    
    var dataToSave = [];
    
    for (var i = 0; i < usuariosList.length; i++) {
      var u = usuariosList[i];
      var email = String(u.Email || u.Email_Usuario || '').trim();
      
      dataToSave.push([
        u.Nombre || u.Nombre_Completo,          // Columna A: Nombre
        email,                                  // Columna B: Email
        esValorVerdadero_(u.Activo),            // Columna C: Activo
        esValorVerdadero_(u.Admin),             // Columna D: Admin
        especialidadPorEmail[email.toLowerCase()] || '' // Columna E: Especialidad
      ]);
    }
    
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    
    if (dataToSave.length > 0) {
      sheet.getRange(2, 1, dataToSave.length, 5).setValues(dataToSave);
    }
    } finally {
      lock.releaseLock();
    }
    
    purgarCache();
    return { success: true, message: "Usuarios actualizados correctamente" };
    
  } catch (e) {
    return { success: false, error: "Error: " + e.toString() };
  }
}

/* ==========================================================
   CANCELAR RESERVA (CON CRUCE DE DATOS Y NOMBRES REALES 🕵️‍♂️)
   ========================================================== */
// Envoltorio con LockService: evita dobles reservas / filas desplazadas si coincide con otra escritura
function adminCancelarReserva(idReserva) {
  return conLockScript_(() => adminCancelarReservaSinLock_(idReserva));
}

function adminCancelarReservaSinLock_(idReserva) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    
    const ss = getDB();
    
    // --- 1. CARGAR DATOS AUXILIARES PARA TRADUCIR CÓDIGOS ---
    
    // A) MAPA DE RECURSOS (ID -> Nombre Real)
    // Asumimos: Col A = ID, Col B = Nombre
    const sheetRecursos = ss.getSheetByName(SHEETS.RECURSOS);
    const dataRecursos = sheetRecursos.getDataRange().getValues();
    const mapRecursos = {}; 
    dataRecursos.forEach(r => { if(r[0]) mapRecursos[String(r[0])] = r[1]; });

    // B) MAPA DE TRAMOS (ID -> Nombre Real)
    // Asumimos: Col A = ID, Col B = Nombre
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    const dataTramos = sheetTramos.getDataRange().getValues();
    const mapTramos = {};
    dataTramos.forEach(t => { if(t[0]) mapTramos[String(t[0])] = t[1]; });

    // C) MAPA DE USUARIOS (Email -> Nombre Persona)
    // Asumimos: Col A = Nombre, Col B = Email
    // El usuario dijo "Nombre en Col A". Necesitamos el Email (Col B) para buscar.
    const sheetUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
    const dataUsuarios = sheetUsuarios.getDataRange().getValues();
    const mapUsuarios = {};
    dataUsuarios.forEach(u => { 
      // u[1] es Email (Clave), u[0] es Nombre (Valor)
      if(u[1]) mapUsuarios[String(u[1]).toLowerCase().trim()] = u[0]; 
    });

    // --- 2. BUSCAR LA RESERVA EN LA HOJA RESERVAS ---
    
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);
    const dataReservas = sheetReservas.getDataRange().getValues();
    
    let rowIndex = -1;
    let reservaInfo = null;

    // Empezamos en 1 para saltar cabecera
    for (let i = 1; i < dataReservas.length; i++) {
      // ID Reserva está en Columna A (0)
      if (String(dataReservas[i][0]) === String(idReserva)) {
        rowIndex = i + 1; // +1 porque filas Excel empiezan en 1
        
        // --- EXTRAER DATOS RAW (CRUDOS) ---
        const rawIdRecurso = dataReservas[i][1]; // Col B: ID_Recurso
        const emailUsuario = dataReservas[i][2]; // Col C: Email
        const rawFecha     = dataReservas[i][3]; // Col D: Fecha
        const rawIdTramo   = dataReservas[i][5]; // Col F: ID_Tramo

        // --- TRADUCIR A NOMBRES REALES ---
        const nombreRecurso = mapRecursos[rawIdRecurso] || rawIdRecurso; // Si no encuentra, usa el ID
        const nombreTramo   = mapTramos[rawIdTramo] || rawIdTramo;
        const nombreUsuario = mapUsuarios[String(emailUsuario).toLowerCase().trim()] || "Usuario";
        const fechaBonita   = formatDate(rawFecha);

        reservaInfo = {
          email: emailUsuario,
          usuario: nombreUsuario,
          recurso: nombreRecurso,
          fecha: fechaBonita,
          tramo: nombreTramo
        };
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error("Reserva no encontrada");
    }

    // --- 3. BORRAR LA FILA ---
    sheetReservas.deleteRow(rowIndex);
    
    // --- 4. ENVIAR CORREO ---
    if (reservaInfo && reservaInfo.email) {
      try {
        const asunto = `❌ Cancelación administrativa: ${reservaInfo.recurso} - ${reservaInfo.fecha}`;
        
        const cuerpoHtml = `
          <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #eee; padding: 20px; border-radius: 8px;">
            <h2 style="color: #d32f2f; margin-top: 0;">Reserva Cancelada</h2>
            
            <p>Hola <strong>${escHtml_(reservaInfo.usuario)}</strong>,</p>
            
            <p>Te informamos que <b>un administrador ha cancelado tu reserva</b>.</p>
            
            <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
            
            <p style="font-weight: bold; margin-bottom: 10px;">Detalles de la reserva eliminada:</p>
            <ul style="background-color: #fff1f0; padding: 15px 30px; border-radius: 5px; list-style-type: none; border: 1px solid #ffccc7;">
              <li style="margin-bottom: 8px;">📦 <strong>Recurso:</strong> ${escHtml_(reservaInfo.recurso)}</li>
              <li style="margin-bottom: 8px;">📅 <strong>Fecha:</strong> ${reservaInfo.fecha}</li>
              <li>⏰ <strong>Tramo:</strong> ${reservaInfo.tramo}</li>
            </ul>
            
            <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
            
            <p style="font-size: 0.85em; color: #777;">
              Si crees que se trata de un error, por favor ponte en contacto con la dirección del centro.
            </p>
          </div>
        `;

        MailApp.sendEmail({
          to: reservaInfo.email,
          subject: asunto,
          htmlBody: cuerpoHtml
        });
        
        Logger.log(`📧 Cancelación enviada a ${reservaInfo.usuario} (${reservaInfo.email})`);
        
      } catch (emailError) {
        Logger.log("⚠️ Error enviando email: " + emailError);
      }
    }

    purgarCache();
    return { success: true };

  } catch (e) {
    Logger.log("Error en adminCancelarReserva: " + e);
    return { success: false, error: e.toString() };
  }
}

// Asegúrate de tener esta función auxiliar al final del archivo
function formatDate(date) {
  if (!date) return "";
  try {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch(e) { return date; }
}


/* ===========================================
   GUARDAR CONFIGURACIÓN (BACKEND ⚙️)
   =========================================== */
function saveBatchConfig(configList) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    
    const ss = getDB();
    let sheet = ss.getSheetByName('CONFIG');
    
    // Si no existe, la crea
    if (!sheet) {
      sheet = ss.insertSheet('CONFIG');
      sheet.appendRow(['CLAVE', 'VALOR', 'DESCRIPCION']); // Cabeceras
    }

    const data = sheet.getDataRange().getValues();
    
    // Mapa: Clave -> Número de Fila
    const keyRowMap = {};
    for (let i = 1; i < data.length; i++) {
      keyRowMap[String(data[i][0])] = i + 1;
    }

    configList.forEach(item => {
      const row = keyRowMap[item.clave];
      if (row) {
        // Actualizar columna B (Valor)
        sheet.getRange(row, 2).setValue(item.valor);
      } else {
        // Nueva fila
        sheet.appendRow([item.clave, item.valor, 'Auto-generado']);
      }
    });
    
    purgarCache();
    return { success: true };
    
  } catch (e) { return { success: false, error: e.toString() }; }
}

/* ============================================
   MIGRACIÓN: ID_Solicitud_Recurrente
   ============================================ */

/**
 * Migra las reservas recurrentes existentes para añadir el ID_Solicitud_Recurrente
 * basándose en el campo Notas que contiene "Reserva recurrente: [ID]"
 */
function migrarIdSolicitudRecurrente_() {
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEETS.RESERVAS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, message: 'No hay reservas que migrar', migradas: 0 };
    }

    // Obtener headers
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => {
      headerMap[h.toString().toLowerCase().trim()] = i;
    });

    // Verificar si existe la columna ID_Solicitud_Recurrente
    let colIdSolicitud = headerMap['id_solicitud_recurrente'];
    if (colIdSolicitud === undefined) {
      // Crear la columna
      const nuevaCol = lastCol + 1;
      sheet.getRange(1, nuevaCol).setValue('ID_Solicitud_Recurrente');
      colIdSolicitud = nuevaCol - 1; // índice 0-based
      Logger.log('✅ Columna ID_Solicitud_Recurrente creada');
    }

    const colNotas = headerMap['notas'];
    if (colNotas === undefined) {
      return { success: false, error: 'No se encontró la columna Notas' };
    }

    // Leer todos los datos
    const numFilas = sheet.getLastRow() - 1;
    const data = sheet.getRange(2, 1, numFilas, sheet.getLastColumn()).getValues();

    let migradas = 0;
    const updates = [];

    data.forEach((row, idx) => {
      const notas = String(row[colNotas] || '');
      const idSolicitudActual = String(row[colIdSolicitud] || '').trim();

      // Si ya tiene valor, saltar
      if (idSolicitudActual) return;

      // Buscar patrón "Reserva recurrente: XXXXX"
      const match = notas.match(/Reserva recurrente:\s*([A-Za-z0-9_-]+)/i);
      if (match && match[1]) {
        const idSolicitud = match[1].trim();
        updates.push({
          fila: idx + 2, // +2 porque empezamos en fila 2
          valor: idSolicitud
        });
        migradas++;
      }
    });

    // Aplicar actualizaciones
    if (updates.length > 0) {
      // Una sola escritura de la columna completa en vez de un setValue por fila
      const columna = data.map(r => [r[colIdSolicitud]]);
      updates.forEach(u => { columna[u.fila - 2][0] = u.valor; });
      sheet.getRange(2, colIdSolicitud + 1, columna.length, 1).setValues(columna);
    }

    Logger.log(`✅ Migración completada: ${migradas} reservas actualizadas`);
    return { success: true, message: `Migración completada`, migradas: migradas };

  } catch (error) {
    Logger.log('❌ Error en migración: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   MATRIZ UNIFICADA: DISPONIBILIDAD + RECURRENCIAS
   ============================================ */

/**
 * Carga datos combinados de disponibilidad y recurrencias para un recurso.
 * Usado por la vista unificada en el admin-panel.
 * @param {string} idRecurso - ID del recurso
 * @returns {Object} { disponibilidad, recurrencias, solicitudesPendientes }
 */
function getDatosMatrizUnificada(idRecurso) {
  try {
    const email = emailActual_();
    const authResult = checkUserAuthorization(email);

    if (!authResult || !authResult.isAdmin) {
      return { success: false, error: "No tienes permisos de administrador." };
    }

    if (!idRecurso) {
      return { success: false, error: "ID de recurso no especificado." };
    }

    const ss = getDB();
    const idRecursoNorm = String(idRecurso).trim().toLowerCase();
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    // 1. DISPONIBILIDAD del recurso
    const sheetDisp = ss.getSheetByName(SHEETS.DISPONIBILIDAD);
    let disponibilidad = [];
    if (sheetDisp && sheetDisp.getLastRow() > 1) {
      const data = sheetDisp.getRange(2, 1, sheetDisp.getLastRow() - 1, 6).getValues();
      disponibilidad = data
        .filter(row => String(row[0]).trim().toLowerCase() === idRecursoNorm)
        .map(row => {
          let diaRaw = String(row[1]).trim();
          let diaFinal = diaRaw;
          if (diaRaw.includes("Lunes")) diaFinal = "Lunes";
          else if (diaRaw.includes("Martes")) diaFinal = "Martes";
          else if (diaRaw.includes("Miércoles") || diaRaw.includes("Miercoles")) diaFinal = "Miércoles";
          else if (diaRaw.includes("Jueves")) diaFinal = "Jueves";
          else if (diaRaw.includes("Viernes")) diaFinal = "Viernes";

          let permRaw = String(row[4]).trim().toLowerCase();
          let permFinal = 'Si';
          if (permRaw === 'no' || permRaw === 'false' || permRaw === '0') permFinal = 'No';

          return {
            id_tramo: String(row[2]).trim(),
            dia_semana: diaFinal,
            permitido: permFinal,
            razon_bloqueo: String(row[5]).trim()
          };
        });
    }

    // 2. RECURRENCIAS APROBADAS ACTIVAS del recurso
    const sheetSol = getOrCreateSheetSolicitudesRecurrentes();
    let recurrencias = [];
    let solicitudesPendientes = [];

    if (sheetSol && sheetSol.getLastRow() > 1) {
      const data = sheetSol.getRange(2, 1, sheetSol.getLastRow() - 1, 16).getValues();

      // Obtener datos de usuarios para enriquecer con área
      const sheetUsers = ss.getSheetByName(SHEETS.USUARIOS);
      const usuariosMap = new Map();
      if (sheetUsers && sheetUsers.getLastRow() > 1) {
        const usersData = sheetUsers.getRange(2, 1, sheetUsers.getLastRow() - 1, 5).getValues();
        usersData.forEach(u => {
          // Usuarios: A=Nombre, B=Email, E=Especialidad
          usuariosMap.set(String(u[1]).trim().toLowerCase(), {
            nombre: String(u[0]).trim(),
            area: String(u[4]).trim() || ''
          });
        });
      }

      data.forEach(row => {
        const idRec = String(row[COLS_SOLICITUDES.ID_RECURSO] || '').trim().toLowerCase();
        if (idRec !== idRecursoNorm) return;

        const estado = String(row[COLS_SOLICITUDES.ESTADO] || '').toLowerCase();
        const fechaFin = row[COLS_SOLICITUDES.FECHA_FIN];
        const emailUsuario = String(row[COLS_SOLICITUDES.EMAIL_USUARIO] || '').trim().toLowerCase();
        const userData = usuariosMap.get(emailUsuario) || { nombre: '', area: '' };

        const solicitud = {
          id_solicitud: String(row[COLS_SOLICITUDES.ID_SOLICITUD] || ''),
          nombre_usuario: String(row[COLS_SOLICITUDES.NOMBRE_USUARIO] || '') || userData.nombre,
          email_usuario: String(row[COLS_SOLICITUDES.EMAIL_USUARIO] || ''),
          area_usuario: userData.area,
          dias_semana: String(row[COLS_SOLICITUDES.DIAS_SEMANA] || ''),
          id_tramo: String(row[COLS_SOLICITUDES.ID_TRAMO] || ''),
          nombre_tramo: String(row[COLS_SOLICITUDES.NOMBRE_TRAMO] || ''),
          fecha_inicio: row[COLS_SOLICITUDES.FECHA_INICIO] instanceof Date ? row[COLS_SOLICITUDES.FECHA_INICIO].toISOString() : '',
          fecha_fin: fechaFin instanceof Date ? fechaFin.toISOString() : '',
          motivo: String(row[COLS_SOLICITUDES.MOTIVO] || ''),
          estado: estado
        };

        // Solicitudes pendientes (para panel de aprobación)
        if (estado === 'pendiente') {
          solicitudesPendientes.push(solicitud);
        }

        // Recurrencias aprobadas y activas (fecha_fin >= hoy)
        if (estado === 'aprobada') {
          let fechaFinDate = fechaFin instanceof Date ? fechaFin : null;
          if (!fechaFinDate && fechaFin) {
            fechaFinDate = new Date(fechaFin);
          }

          // Solo incluir si está vigente
          if (fechaFinDate && fechaFinDate >= hoy) {
            recurrencias.push(solicitud);
          }
        }
      });
    }

    return {
      success: true,
      disponibilidad: disponibilidad,
      recurrencias: recurrencias,
      solicitudesPendientes: solicitudesPendientes
    };

  } catch (error) {
    Logger.log('Error en getDatosMatrizUnificada: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   FIN DEL ARCHIVO AdminFunctions.gs
   ============================================ */
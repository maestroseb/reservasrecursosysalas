/** 
 * ADMIN FUNCTIONS
 * Funciones exclusivas para administradores del sistema
 */

/* ============================================
   VERIFICACIÓN DE PERMISOS
   ============================================ */

function isUserAdmin() {
  const email = Session.getActiveUser().getEmail();
  const authResult = checkUserAuthorization(email);
  return authResult.isAdmin || false;
}



/* ============================================
   CARGA DE DATOS ADMIN (V9 - FIX FINAL DEFINITIVO)
   ============================================ */

function getAdminData() {
  try {
    const email = Session.getActiveUser().getEmail();
    const authResult = checkUserAuthorization(email);

    if (!authResult || !authResult.isAdmin) {
      return { success: false, error: "No tienes permisos de administrador." };
    }

    // Ejecutar migración de ID_Solicitud_Recurrente si es necesario
    try {
      migrarIdSolicitudRecurrente();
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
        Activo: Boolean(row[2]), 
        Admin: Boolean(row[3])
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

/* ============================================
   GESTIÓN DE RECURSOS
   ============================================ */

function createRecurso(recursoData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { id_recurso, nombre, tipo, capacidad, ubicacion, icono, estado } = recursoData;
    
    if (!id_recurso || !nombre || !tipo) {
      throw new Error("Faltan campos obligatorios (ID, Nombre, Tipo).");
    }
    
    const ss = getDB();
    const sheetRecursos = ss.getSheetByName(SHEETS.RECURSOS);
    
    const recursos = sheetToObjects(sheetRecursos);
    if (recursos.find(r => r.id_recurso === id_recurso)) {
      throw new Error(`Ya existe un recurso con el ID: ${id_recurso}`);
    }
    
    const headers = sheetRecursos.getRange(1, 1, 1, sheetRecursos.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i);
    
    const nuevaFila = new Array(headers.length).fill("");
    if (headerMap['id_recurso'] !== undefined) nuevaFila[headerMap['id_recurso']] = id_recurso;
    if (headerMap['nombre'] !== undefined) nuevaFila[headerMap['nombre']] = nombre;
    if (headerMap['tipo'] !== undefined) nuevaFila[headerMap['tipo']] = tipo;
    if (headerMap['capacidad'] !== undefined) nuevaFila[headerMap['capacidad']] = capacidad || 1;
    if (headerMap['ubicacion'] !== undefined) nuevaFila[headerMap['ubicacion']] = ubicacion || '';
    if (headerMap['icono'] !== undefined) nuevaFila[headerMap['icono']] = icono || '';
    if (headerMap['estado'] !== undefined) nuevaFila[headerMap['estado']] = estado || 'Activo';
    
    sheetRecursos.appendRow(nuevaFila);
    purgarCache();
    
    return { success: true, message: "Recurso creado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en createRecurso: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function updateRecurso(recursoData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { id_recurso, nombre, tipo, capacidad, ubicacion, icono, estado } = recursoData;
    
    const ss = getDB();
    const sheetRecursos = ss.getSheetByName(SHEETS.RECURSOS);
    const headers = sheetRecursos.getRange(1, 1, 1, sheetRecursos.getLastColumn()).getValues()[0];
    const COL_ID = headers.indexOf("ID_Recurso");
    
    const idColumnRange = sheetRecursos.getRange(2, COL_ID + 1, sheetRecursos.getLastRow() - 1);
    const textFinder = idColumnRange.createTextFinder(id_recurso).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró el recurso.");
    }
    
    const fila = celda.getRow();
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i + 1);
    
    if (nombre && headerMap['nombre']) sheetRecursos.getRange(fila, headerMap['nombre']).setValue(nombre);
    if (tipo && headerMap['tipo']) sheetRecursos.getRange(fila, headerMap['tipo']).setValue(tipo);
    if (capacidad && headerMap['capacidad']) sheetRecursos.getRange(fila, headerMap['capacidad']).setValue(capacidad);
    if (ubicacion !== undefined && headerMap['ubicacion']) sheetRecursos.getRange(fila, headerMap['ubicacion']).setValue(ubicacion);
    if (icono !== undefined && headerMap['icono']) sheetRecursos.getRange(fila, headerMap['icono']).setValue(icono);
    if (estado && headerMap['estado']) sheetRecursos.getRange(fila, headerMap['estado']).setValue(estado);
    
    purgarCache();
    
    return { success: true, message: "Recurso actualizado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en updateRecurso: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function deleteRecurso(idRecurso) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const ss = getDB();
    const sheetRecursos = ss.getSheetByName(SHEETS.RECURSOS);
    const headers = sheetRecursos.getRange(1, 1, 1, sheetRecursos.getLastColumn()).getValues()[0];
    const COL_ID = headers.indexOf("ID_Recurso");
    
    const idColumnRange = sheetRecursos.getRange(2, COL_ID + 1, sheetRecursos.getLastRow() - 1);
    const textFinder = idColumnRange.createTextFinder(idRecurso).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró el recurso.");
    }
    
    sheetRecursos.deleteRow(celda.getRow());
    purgarCache();
    
    return { success: true, message: "Recurso eliminado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en deleteRecurso: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function generarDisponibilidadRecurso(idRecurso, configuracion) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const ss = getDB();
    const sheetDisponibilidad = ss.getSheetByName(SHEETS.DISPONIBILIDAD);
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    
    const tramos = sheetToObjects(sheetTramos);
    const { dias, permitido, razonBloqueo } = configuracion;
    
    const headers = sheetDisponibilidad.getRange(1, 1, 1, sheetDisponibilidad.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i);
    
    const filasNuevas = [];
    
    dias.forEach(dia => {
      tramos.forEach(tramo => {
        const nuevaFila = new Array(headers.length).fill("");
        if (headerMap['id_recurso'] !== undefined) nuevaFila[headerMap['id_recurso']] = idRecurso;
        if (headerMap['dia_semana'] !== undefined) nuevaFila[headerMap['dia_semana']] = dia;
        if (headerMap['id_tramo'] !== undefined) nuevaFila[headerMap['id_tramo']] = tramo.id_tramo;
        if (headerMap['permitido'] !== undefined) nuevaFila[headerMap['permitido']] = permitido || 'Si';
        if (headerMap['razon_bloqueo'] !== undefined) nuevaFila[headerMap['razon_bloqueo']] = razonBloqueo || '';
        
        filasNuevas.push(nuevaFila);
      });
    });
    
    if (filasNuevas.length > 0) {
      const rangoDestino = sheetDisponibilidad.getRange(
        sheetDisponibilidad.getLastRow() + 1, 
        1, 
        filasNuevas.length, 
        headers.length
      );
      rangoDestino.setValues(filasNuevas);
    }
    
    purgarCache();
    
    return { success: true, message: `Se generaron ${filasNuevas.length} registros de disponibilidad.` };
    
  } catch (error) {
    Logger.log(`Error en generarDisponibilidadRecurso: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/* ============================================
   GESTIÓN DE TRAMOS
   ============================================ */

function createTramo(tramoData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { id_tramo, nombre_tramo, hora_inicio, hora_fin } = tramoData;
    
    if (!id_tramo || !nombre_tramo || !hora_inicio || !hora_fin) {
      throw new Error("Faltan campos obligatorios.");
    }
    
    const ss = getDB();
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    
    const tramos = sheetToObjects(sheetTramos);
    if (tramos.find(t => t.id_tramo === id_tramo)) {
      throw new Error(`Ya existe un tramo con el ID: ${id_tramo}`);
    }
    
    const headers = sheetTramos.getRange(1, 1, 1, sheetTramos.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i);
    
    const nuevaFila = new Array(headers.length).fill("");
    if (headerMap['id_tramo'] !== undefined) nuevaFila[headerMap['id_tramo']] = id_tramo;
    if (headerMap['nombre_tramo'] !== undefined) nuevaFila[headerMap['nombre_tramo']] = nombre_tramo;
    if (headerMap['hora_inicio'] !== undefined) nuevaFila[headerMap['hora_inicio']] = hora_inicio;
    if (headerMap['hora_fin'] !== undefined) nuevaFila[headerMap['hora_fin']] = hora_fin;
    
    sheetTramos.appendRow(nuevaFila);
    purgarCache();
    
    return { success: true, message: "Tramo creado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en createTramo: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function updateTramo(tramoData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { id_tramo, nombre_tramo, hora_inicio, hora_fin } = tramoData;
    
    const ss = getDB();
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    const headers = sheetTramos.getRange(1, 1, 1, sheetTramos.getLastColumn()).getValues()[0];
    const COL_ID = headers.indexOf("ID_Tramo");
    
    const idColumnRange = sheetTramos.getRange(2, COL_ID + 1, sheetTramos.getLastRow() - 1);
    const textFinder = idColumnRange.createTextFinder(id_tramo).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró el tramo.");
    }
    
    const fila = celda.getRow();
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i + 1);
    
    if (nombre_tramo && headerMap['nombre_tramo']) sheetTramos.getRange(fila, headerMap['nombre_tramo']).setValue(nombre_tramo);
    if (hora_inicio && headerMap['hora_inicio']) sheetTramos.getRange(fila, headerMap['hora_inicio']).setValue(hora_inicio);
    if (hora_fin && headerMap['hora_fin']) sheetTramos.getRange(fila, headerMap['hora_fin']).setValue(hora_fin);
    
    purgarCache();
    
    return { success: true, message: "Tramo actualizado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en updateTramo: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function deleteTramo(idTramo) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const ss = getDB();
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    const headers = sheetTramos.getRange(1, 1, 1, sheetTramos.getLastColumn()).getValues()[0];
    const COL_ID = headers.indexOf("ID_Tramo");
    
    const idColumnRange = sheetTramos.getRange(2, COL_ID + 1, sheetTramos.getLastRow() - 1);
    const textFinder = idColumnRange.createTextFinder(idTramo).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró el tramo.");
    }
    
    sheetTramos.deleteRow(celda.getRow());
    purgarCache();
    
    return { success: true, message: "Tramo eliminado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en deleteTramo: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/* ============================================
   GESTIÓN DE DISPONIBILIDAD
   ============================================ */

function updateDisponibilidad(dispData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { id_recurso, dia_semana, id_tramo, permitido, razon_bloqueo } = dispData;
    
    const ss = getDB();
    const sheetDisponibilidad = ss.getSheetByName(SHEETS.DISPONIBILIDAD);
    
    const disponibilidad = sheetToObjects(sheetDisponibilidad);
    const index = disponibilidad.findIndex(d => 
      d.id_recurso === id_recurso && 
      d.dia_semana.toString() === dia_semana.toString() && 
      d.id_tramo === id_tramo
    );
    
    if (index === -1) {
      throw new Error("No se encontró el registro de disponibilidad.");
    }
    
    const fila = index + 2;
    const headers = sheetDisponibilidad.getRange(1, 1, 1, sheetDisponibilidad.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i + 1);
    
    if (permitido && headerMap['permitido']) sheetDisponibilidad.getRange(fila, headerMap['permitido']).setValue(permitido);
    if (razon_bloqueo !== undefined && headerMap['razon_bloqueo']) sheetDisponibilidad.getRange(fila, headerMap['razon_bloqueo']).setValue(razon_bloqueo);
    
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEYS.DISPONIBILIDAD + id_recurso);
    
    return { success: true, message: "Disponibilidad actualizada con éxito." };
    
  } catch (error) {
    Logger.log(`Error en updateDisponibilidad: ${error.message}`);
    return { success: false, error: error.message };
  }
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


/* ============================================
   GESTIÓN DE USUARIOS
   ============================================ */

function createUsuario(usuarioData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { email_usuario, nombre_completo, activo, admin } = usuarioData;
    
    if (!email_usuario || !nombre_completo) {
      throw new Error("Faltan campos obligatorios (Email, Nombre).");
    }
    
    const ss = getDB();
    const sheetUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
    
    const usuarios = sheetToObjects(sheetUsuarios);
    if (usuarios.find(u => u.email_usuario.toLowerCase() === email_usuario.toLowerCase())) {
      throw new Error(`Ya existe un usuario con el email: ${email_usuario}`);
    }
    
    sheetUsuarios.appendRow([email_usuario, nombre_completo, activo || 'Activo', admin || false]);
    purgarCache();
    
    return { success: true, message: "Usuario creado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en createUsuario: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function updateUsuario(usuarioData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { email_usuario, nombre_completo, activo, admin } = usuarioData;
    
    const ss = getDB();
    const sheetUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
    const headers = sheetUsuarios.getRange(1, 1, 1, sheetUsuarios.getLastColumn()).getValues()[0];
    const COL_EMAIL = headers.indexOf("Email_Usuario");
    
    const emailColumnRange = sheetUsuarios.getRange(2, COL_EMAIL + 1, sheetUsuarios.getLastRow() - 1);
    const textFinder = emailColumnRange.createTextFinder(email_usuario).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró el usuario.");
    }
    
    const fila = celda.getRow();
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i + 1);
    
    if (nombre_completo && headerMap['nombre_completo']) sheetUsuarios.getRange(fila, headerMap['nombre_completo']).setValue(nombre_completo);
    if (activo !== undefined && headerMap['activo']) sheetUsuarios.getRange(fila, headerMap['activo']).setValue(activo);
    if (admin !== undefined && headerMap['admin']) sheetUsuarios.getRange(fila, headerMap['admin']).setValue(admin);
    
    purgarCache();
    
    return { success: true, message: "Usuario actualizado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en updateUsuario: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function deleteUsuario(emailUsuario) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const ss = getDB();
    const sheetUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
    const headers = sheetUsuarios.getRange(1, 1, 1, sheetUsuarios.getLastColumn()).getValues()[0];
    const COL_EMAIL = headers.indexOf("Email_Usuario");
    
    const emailColumnRange = sheetUsuarios.getRange(2, COL_EMAIL + 1, sheetUsuarios.getLastRow() - 1);
    const textFinder = emailColumnRange.createTextFinder(emailUsuario).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró el usuario.");
    }
    
    sheetUsuarios.deleteRow(celda.getRow());
    purgarCache();
    
    return { success: true, message: "Usuario eliminado con éxito." };
    
  } catch (error) {
    Logger.log(`Error en deleteUsuario: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/* ============================================
   GESTIÓN DE RESERVAS (ADMIN)
   ============================================ */

function updateReservaAdmin(reservaData) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const { id_reserva, id_recurso, fecha, id_tramo, cantidad, notas, curso, estado } = reservaData;
    
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);
    const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const COL_ID = headers.indexOf("ID_Reserva");
    
    const idColumnRange = sheetReservas.getRange(2, COL_ID + 1, sheetReservas.getLastRow() - 1);
    const textFinder = idColumnRange.createTextFinder(id_reserva).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró la reserva.");
    }
    
    const fila = celda.getRow();
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim().toLowerCase()] = i + 1);
    
    const rowData = sheetReservas.getRange(fila, 1, 1, headers.length).getValues()[0];
    const emailUsuario = rowData[headers.indexOf("Email_Usuario")];
    
    const oldValues = {
      id_recurso: rowData[headers.indexOf("ID_Recurso")],
      fecha: rowData[headers.indexOf("Fecha")],
      id_tramo: rowData[headers.indexOf("ID_Tramo")]
    };
    
    if (id_recurso && headerMap['id_recurso']) sheetReservas.getRange(fila, headerMap['id_recurso']).setValue(id_recurso);
    if (fecha && headerMap['fecha']) sheetReservas.getRange(fila, headerMap['fecha']).setValue(new Date(fecha + "T12:00:00Z"));
    if (id_tramo && headerMap['id_tramo']) sheetReservas.getRange(fila, headerMap['id_tramo']).setValue(id_tramo);
    if (cantidad && headerMap['cantidad']) sheetReservas.getRange(fila, headerMap['cantidad']).setValue(cantidad);
    if (notas !== undefined && headerMap['notas']) sheetReservas.getRange(fila, headerMap['notas']).setValue(notas);
    if (curso && headerMap['curso']) sheetReservas.getRange(fila, headerMap['curso']).setValue(curso);
    if (estado && headerMap['estado']) sheetReservas.getRange(fila, headerMap['estado']).setValue(estado);
    
    if (emailUsuario && (id_recurso !== oldValues.id_recurso || fecha !== oldValues.fecha || id_tramo !== oldValues.id_tramo)) {
      enviarNotificacionCambioReserva(emailUsuario, reservaData);
    }
    
    purgarCache();
    
    return { success: true, message: "Reserva actualizada con éxito." };
    
  } catch (error) {
    Logger.log(`Error en updateReservaAdmin: ${error.message}`);
    return { success: false, error: error.message };
  }
}

function deleteReservaAdmin(idReserva) {
  try {
    if (!isUserAdmin()) {
      throw new Error("No tienes permisos de administrador.");
    }
    
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);
    const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const COL_ID = headers.indexOf("ID_Reserva");
    
    const idColumnRange = sheetReservas.getRange(2, COL_ID + 1, sheetReservas.getLastRow() - 1);
    const textFinder = idColumnRange.createTextFinder(idReserva).matchEntireCell(true);
    const celda = textFinder.findNext();
    
    if (!celda) {
      throw new Error("No se encontró la reserva.");
    }
    
    const fila = celda.getRow();
    const rowData = sheetReservas.getRange(fila, 1, 1, headers.length).getValues()[0];
    const emailUsuario = rowData[headers.indexOf("Email_Usuario")];
    
    sheetReservas.deleteRow(fila);
    
    if (emailUsuario) {
      enviarNotificacionEliminacionReserva(emailUsuario, rowData, headers);
    }
    
    purgarCache();
    
    return { success: true, message: "Reserva eliminada con éxito." };
    
  } catch (error) {
    Logger.log(`Error en deleteReservaAdmin: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/* ============================================
   NOTIFICACIONES POR EMAIL
   ============================================ */

function enviarNotificacionCambioReserva(emailUsuario, nuevosValores) {
  try {
    const staticData = getStaticData();
    const authResult = checkUserAuthorization(emailUsuario);
    
    const recurso = staticData.recursos.find(r => r.id_recurso === nuevosValores.id_recurso) || { nombre: nuevosValores.id_recurso };
    const tramo = staticData.tramos.find(t => t.id_tramo === nuevosValores.id_tramo) || { nombre_tramo: nuevosValores.id_tramo, hora_inicio: '', hora_fin: '' };
    const fecha = new Date(nuevosValores.fecha + "T12:00:00Z");
    const fechaFormateada = fecha.toLocaleDateString('es-ES', { 
      day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' 
    });
    
    const asunto = `⚠️ Cambio en tu Reserva: ${recurso.nombre}`;
    const cuerpoHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; border-radius: 10px 10px 0 0;">
          <h1 style="color: white; margin: 0;">⚠️ Cambio en tu Reserva</h1>
        </div>
        <div style="background: white; padding: 30px; border: 1px solid #e0e0e0; border-top: none; border-radius: 0 0 10px 10px;">
          <p>Hola <strong>${authResult.userName || ''}</strong>,</p>
          <p>Un administrador ha <strong>modificado</strong> una de tus reservas.</p>
          <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;">
          <h3 style="color: #333;">Nuevos datos de la reserva:</h3>
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Recurso:</strong></td>
              <td style="padding: 10px;">${recurso.nombre}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Fecha:</strong></td>
              <td style="padding: 10px;">${fechaFormateada}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Tramo:</strong></td>
              <td style="padding: 10px;">${tramo.nombre_tramo} (${tramo.hora_inicio} - ${tramo.hora_fin})</td>
            </tr>
            ${nuevosValores.curso ? `
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Curso:</strong></td>
              <td style="padding: 10px;">${nuevosValores.curso}</td>
            </tr>` : ''}
            ${nuevosValores.cantidad > 1 ? `
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Cantidad:</strong></td>
              <td style="padding: 10px;">${nuevosValores.cantidad}</td>
            </tr>` : ''}
            ${nuevosValores.notas ? `
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Notas:</strong></td>
              <td style="padding: 10px;">${nuevosValores.notas}</td>
            </tr>` : ''}
          </table>
          <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;">
          <p style="font-size: 0.9em; color: #777;">Si tienes dudas sobre este cambio, contacta con el administrador del sistema.</p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: emailUsuario,
      subject: asunto,
      htmlBody: cuerpoHtml
    });
    
    Logger.log(`Notificación de cambio enviada a ${emailUsuario}`);
    
  } catch (e) {
    Logger.log(`Error al enviar notificación de cambio: ${e.message}`);
  }
}

function enviarNotificacionEliminacionReserva(emailUsuario, rowData, headers) {
  try {
    const staticData = getStaticData();
    const authResult = checkUserAuthorization(emailUsuario);
    
    const idRecurso = rowData[headers.indexOf("ID_Recurso")];
    const idTramo = rowData[headers.indexOf("ID_Tramo")];
    const fecha = new Date(rowData[headers.indexOf("Fecha")]);
    
    const recurso = staticData.recursos.find(r => r.id_recurso === idRecurso) || { nombre: idRecurso };
    const tramo = staticData.tramos.find(t => t.id_tramo === idTramo) || { nombre_tramo: idTramo, hora_inicio: '', hora_fin: '' };
    const fechaFormateada = fecha.toLocaleDateString('es-ES', { 
      day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' 
    });
    
    const asunto = `❌ Reserva Eliminada: ${recurso.nombre}`;
    const cuerpoHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); padding: 20px; border-radius: 10px 10px 0 0;">
          <h1 style="color: white; margin: 0;">❌ Reserva Eliminada</h1>
        </div>
        <div style="background: white; padding: 30px; border: 1px solid #e0e0e0; border-top: none; border-radius: 0 0 10px 10px;">
          <p>Hola <strong>${authResult.userName || ''}</strong>,</p>
          <p>Un administrador ha <strong>eliminado</strong> una de tus reservas.</p>
          <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;">
          <h3 style="color: #333;">Datos de la reserva eliminada:</h3>
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Recurso:</strong></td>
              <td style="padding: 10px;">${recurso.nombre}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Fecha:</strong></td>
              <td style="padding: 10px;">${fechaFormateada}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f9f9f9;"><strong>Tramo:</strong></td>
              <td style="padding: 10px;">${tramo.nombre_tramo} (${tramo.hora_inicio} - ${tramo.hora_fin})</td>
            </tr>
          </table>
          <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;">
          <p style="font-size: 0.9em; color: #777;">Si tienes dudas sobre esta eliminación, contacta con el administrador del sistema.</p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: emailUsuario,
      subject: asunto,
      htmlBody: cuerpoHtml
    });
    
    Logger.log(`Notificación de eliminación enviada a ${emailUsuario}`);
    
  } catch (e) {
    Logger.log(`Error al enviar notificación de eliminación: ${e.message}`);
  }
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
      sheet.getRange(sheet.getLastRow() + 1, 1, nuevasFilas.length, 6).setValues(nuevasFilas);
    }
    
    purgarCache();
    return { success: true };
    
  } catch (e) { return { success: false, error: e.toString() }; }
}

/* ===========================================
   GUARDADO DE USUARIOS (Batch Save)
   =========================================== */
function saveBatchUsuarios(usuariosList) {
  try {
    if (!isUserAdmin()) throw new Error("Permiso denegado");
    
    var ss = getDB();
    var sheet = ss.getSheetByName(SHEETS.USUARIOS);
    
    var dataToSave = [];
    
    for (var i = 0; i < usuariosList.length; i++) {
      var u = usuariosList[i];
      
      // Normalizamos booleanos
      var esActivo = (u.Activo === true || u.Activo === "TRUE" || u.Activo === "Si");
      var esAdmin = (u.Admin === true || u.Admin === "TRUE" || u.Admin === "Si");
      
      // AQUÍ ESTÁ EL CAMBIO: Primero Nombre, luego Email
      dataToSave.push([
        u.Nombre || u.Nombre_Completo,  // Columna A: Nombre
        u.Email || u.Email_Usuario,     // Columna B: Email
        esActivo,                       // Columna C: Activo
        esAdmin                         // Columna D: Admin
      ]);
    }
    
    // Guardar...
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }
    
    if (dataToSave.length > 0) {
      sheet.getRange(2, 1, dataToSave.length, 4).setValues(dataToSave);
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
function adminCancelarReserva(idReserva) {
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
            
            <p>Hola <strong>${reservaInfo.usuario}</strong>,</p>
            
            <p>Te informamos que <b>un administrador ha cancelado tu reserva</b>.</p>
            
            <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
            
            <p style="font-weight: bold; margin-bottom: 10px;">Detalles de la reserva eliminada:</p>
            <ul style="background-color: #fff1f0; padding: 15px 30px; border-radius: 5px; list-style-type: none; border: 1px solid #ffccc7;">
              <li style="margin-bottom: 8px;">📦 <strong>Recurso:</strong> ${reservaInfo.recurso}</li>
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
function migrarIdSolicitudRecurrente() {
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
    updates.forEach(u => {
      sheet.getRange(u.fila, colIdSolicitud + 1).setValue(u.valor);
    });

    Logger.log(`✅ Migración completada: ${migradas} reservas actualizadas`);
    return { success: true, message: `Migración completada`, migradas: migradas };

  } catch (error) {
    Logger.log('❌ Error en migración: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   FIN DEL ARCHIVO AdminFunctions.gs
   ============================================ */
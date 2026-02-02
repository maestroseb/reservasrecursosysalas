/* ============================================
   SISTEMA DE RESERVAS RECURRENTES - BACKEND
   ============================================ */

/**
 * Nombre de la hoja de solicitudes recurrentes
 */
const SHEET_SOLICITUDES_RECURRENTES = 'SolicitudesRecurrentes';

/**
 * Columnas de la hoja SolicitudesRecurrentes
 */
const COLS_SOLICITUDES = {
  ID_SOLICITUD: 0,
  ID_RECURSO: 1,
  NOMBRE_RECURSO: 2,
  EMAIL_USUARIO: 3,
  NOMBRE_USUARIO: 4,
  DIAS_SEMANA: 5,
  ID_TRAMO: 6,
  NOMBRE_TRAMO: 7,
  FECHA_INICIO: 8,
  FECHA_FIN: 9,
  MOTIVO: 10,
  ESTADO: 11,
  FECHA_SOLICITUD: 12,
  FECHA_RESOLUCION: 13,
  ADMIN_RESOLUTOR: 14,
  NOTAS_ADMIN: 15
};

/**
 * Headers de la hoja SolicitudesRecurrentes
 */
const HEADERS_SOLICITUDES = [
  'ID_Solicitud',
  'ID_Recurso',
  'Nombre_Recurso',
  'Email_Usuario',
  'Nombre_Usuario',
  'Dias_Semana',
  'ID_Tramo',
  'Nombre_Tramo',
  'Fecha_Inicio',
  'Fecha_Fin',
  'Motivo',
  'Estado',
  'Fecha_Solicitud',
  'Fecha_Resolucion',
  'Admin_Resolutor',
  'Notas_Admin'
];

/* ============================================
   GESTIÓN DE LA HOJA
   ============================================ */

/**
 * Obtiene o crea la hoja de solicitudes recurrentes
 */
function getOrCreateSheetSolicitudesRecurrentes() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_SOLICITUDES_RECURRENTES);

  if (!sheet) {
    Logger.log('📝 Creando hoja SolicitudesRecurrentes...');
    sheet = ss.insertSheet(SHEET_SOLICITUDES_RECURRENTES);

    // Añadir headers
    sheet.getRange(1, 1, 1, HEADERS_SOLICITUDES.length).setValues([HEADERS_SOLICITUDES]);

    // Formato de headers
    const headerRange = sheet.getRange(1, 1, 1, HEADERS_SOLICITUDES.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f4f6');

    // Congelar primera fila
    sheet.setFrozenRows(1);

    // Ajustar anchos de columna
    sheet.setColumnWidth(1, 120);  // ID
    sheet.setColumnWidth(3, 150);  // Nombre Recurso
    sheet.setColumnWidth(5, 150);  // Nombre Usuario
    sheet.setColumnWidth(11, 200); // Motivo

    Logger.log('✅ Hoja SolicitudesRecurrentes creada correctamente');
  }

  return sheet;
}

/* ============================================
   CREAR SOLICITUD (USUARIO)
   ============================================ */

/**
 * Crea una nueva solicitud de reserva recurrente
 * @param {Object} datos - Datos de la solicitud
 * Formatos aceptados:
 *   - Nuevo: { id_recurso, selecciones: [{dia, id_tramo}], fecha_fin, motivo }
 *   - Antiguo: { id_recurso, dias_semana: ['L','M'], id_tramo, fecha_fin, motivo }
 * @returns {Object} - Resultado de la operación
 */
function crearSolicitudRecurrente(datos) {
  try {
    Logger.log('📝 Creando solicitud recurrente: ' + JSON.stringify(datos));

    // Validaciones básicas
    if (!datos.id_recurso) throw new Error('Recurso no especificado');
    if (!datos.fecha_fin) throw new Error('Fecha fin no especificada');
    if (!datos.motivo || datos.motivo.trim().length < 10) throw new Error('El motivo debe tener al menos 10 caracteres');

    // Detectar formato y normalizar
    let seleccionesNormalizadas = [];

    if (datos.selecciones && Array.isArray(datos.selecciones) && datos.selecciones.length > 0) {
      // Formato nuevo: múltiples día-tramo
      seleccionesNormalizadas = datos.selecciones;
    } else if (datos.dias_semana && datos.id_tramo) {
      // Formato antiguo: un tramo para todos los días
      const dias = Array.isArray(datos.dias_semana) ? datos.dias_semana : datos.dias_semana.split(',');
      seleccionesNormalizadas = dias.map(d => ({ dia: d.trim(), id_tramo: datos.id_tramo }));
    } else {
      throw new Error('Selecciona al menos un tramo');
    }

    if (seleccionesNormalizadas.length === 0) {
      throw new Error('Selecciona al menos un tramo');
    }

    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error('No se pudo obtener el email del usuario');

    // Obtener datos adicionales
    const ss = getDB();
    const recursos = sheetToObjects(ss.getSheetByName(SHEETS.RECURSOS));
    const recurso = recursos.find(r => String(r.id_recurso) === String(datos.id_recurso));
    if (!recurso) throw new Error('Recurso no encontrado');

    const tramos = sheetToObjects(ss.getSheetByName(SHEETS.TRAMOS));
    const tramosMap = {};
    tramos.forEach(t => { tramosMap[t.id_tramo] = t; });

    // Verificar que todos los tramos existen
    for (const sel of seleccionesNormalizadas) {
      if (!tramosMap[sel.id_tramo]) {
        throw new Error('Tramo no encontrado: ' + sel.id_tramo);
      }
    }

    // Obtener nombre del usuario
    const usuarios = sheetToObjects(ss.getSheetByName(SHEETS.USUARIOS));
    const usuario = usuarios.find(u => u.email_usuario && u.email_usuario.toLowerCase() === userEmail.toLowerCase());
    const nombreUsuario = usuario ? (usuario.nombre_completo || userEmail) : userEmail;

    // Generar ID único
    const idSolicitud = 'SOL-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // Formatear selecciones para almacenamiento: "L:T001,M:T001,X:T002"
    const seleccionesStr = seleccionesNormalizadas.map(s => `${s.dia}:${s.id_tramo}`).join(',');

    // Determinar tramo(s) para mostrar
    const tramosUnicos = [...new Set(seleccionesNormalizadas.map(s => s.id_tramo))];
    let tramoIdGuardar, tramoNombreGuardar;

    if (tramosUnicos.length === 1) {
      tramoIdGuardar = tramosUnicos[0];
      tramoNombreGuardar = tramosMap[tramosUnicos[0]].nombre_tramo;
    } else {
      tramoIdGuardar = tramosUnicos.join(';');
      tramoNombreGuardar = '(Múltiples tramos)';
    }

    // Fecha de inicio (hoy o la especificada)
    const fechaInicio = datos.fecha_inicio ? new Date(datos.fecha_inicio) : new Date();
    const fechaFin = new Date(datos.fecha_fin);

    // Crear fila
    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const nuevaFila = [
      idSolicitud,
      datos.id_recurso,
      recurso.nombre,
      userEmail,
      nombreUsuario,
      seleccionesStr,  // Ahora guarda "L:T001,M:T002" en lugar de "L,M"
      tramoIdGuardar,
      tramoNombreGuardar,
      fechaInicio,
      fechaFin,
      datos.motivo.trim(),
      'Pendiente',
      new Date(),
      '',
      '',
      ''
    ];

    sheet.appendRow(nuevaFila);

    // Formatear info de tramos para el email
    const infoTramos = seleccionesNormalizadas.map(s => {
      const t = tramosMap[s.id_tramo];
      const diasNombres = { 'L': 'Lunes', 'M': 'Martes', 'X': 'Miércoles', 'J': 'Jueves', 'V': 'Viernes' };
      return `${diasNombres[s.dia] || s.dia}: ${t.nombre_tramo} (${t.hora_inicio}-${t.hora_fin})`;
    }).join('\n');

    // Enviar email al admin
    enviarEmailNuevaSolicitudRecurrente({
      id: idSolicitud,
      recurso: recurso.nombre,
      usuario: nombreUsuario,
      email: userEmail,
      dias: seleccionesStr,
      tramo: tramoNombreGuardar,
      infoTramos: infoTramos,
      fechaInicio: fechaInicio,
      fechaFin: fechaFin,
      motivo: datos.motivo
    });

    Logger.log('✅ Solicitud recurrente creada: ' + idSolicitud);

    return {
      success: true,
      id_solicitud: idSolicitud,
      message: 'Solicitud enviada correctamente. Recibirás un email cuando sea revisada.'
    };

  } catch (error) {
    Logger.log('❌ Error en crearSolicitudRecurrente: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/* ============================================
   OBTENER SOLICITUDES
   ============================================ */

/**
 * Obtiene todas las solicitudes recurrentes (para admin)
 * @param {string} filtroEstado - Opcional: 'Pendiente', 'Aprobada', 'Rechazada', 'todas'
 */
function getSolicitudesRecurrentes(filtroEstado) {
  try {
    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, solicitudes: [] };
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS_SOLICITUDES.length).getValues();

    // Función auxiliar para serializar fechas
    const serializarFecha = (val) => {
      if (!val) return '';
      if (val instanceof Date) return val.toISOString();
      return String(val);
    };

    let solicitudes = data.map(row => ({
      id_solicitud: String(row[COLS_SOLICITUDES.ID_SOLICITUD] || ''),
      id_recurso: String(row[COLS_SOLICITUDES.ID_RECURSO] || ''),
      nombre_recurso: String(row[COLS_SOLICITUDES.NOMBRE_RECURSO] || ''),
      email_usuario: String(row[COLS_SOLICITUDES.EMAIL_USUARIO] || ''),
      nombre_usuario: String(row[COLS_SOLICITUDES.NOMBRE_USUARIO] || ''),
      dias_semana: String(row[COLS_SOLICITUDES.DIAS_SEMANA] || ''),
      id_tramo: String(row[COLS_SOLICITUDES.ID_TRAMO] || ''),
      nombre_tramo: String(row[COLS_SOLICITUDES.NOMBRE_TRAMO] || ''),
      fecha_inicio: serializarFecha(row[COLS_SOLICITUDES.FECHA_INICIO]),
      fecha_fin: serializarFecha(row[COLS_SOLICITUDES.FECHA_FIN]),
      motivo: String(row[COLS_SOLICITUDES.MOTIVO] || ''),
      estado: String(row[COLS_SOLICITUDES.ESTADO] || ''),
      fecha_solicitud: serializarFecha(row[COLS_SOLICITUDES.FECHA_SOLICITUD]),
      fecha_resolucion: serializarFecha(row[COLS_SOLICITUDES.FECHA_RESOLUCION]),
      admin_resolutor: String(row[COLS_SOLICITUDES.ADMIN_RESOLUTOR] || ''),
      notas_admin: String(row[COLS_SOLICITUDES.NOTAS_ADMIN] || '')
    })).filter(s => s.id_solicitud); // Filtrar filas vacías

    // Filtrar por estado si se especifica
    if (filtroEstado && filtroEstado !== 'todas') {
      solicitudes = solicitudes.filter(s =>
        s.estado.toLowerCase() === filtroEstado.toLowerCase()
      );
    }

    // Ordenar por fecha de solicitud (más recientes primero)
    solicitudes.sort((a, b) => new Date(b.fecha_solicitud) - new Date(a.fecha_solicitud));

    Logger.log('✅ getSolicitudesRecurrentes: Devolviendo ' + solicitudes.length + ' solicitudes');
    return { success: true, solicitudes: solicitudes };

  } catch (error) {
    Logger.log('❌ Error en getSolicitudesRecurrentes: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Obtiene las solicitudes del usuario actual
 */
function getMisSolicitudesRecurrentes() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const result = getSolicitudesRecurrentes('todas');

    if (!result.success) return result;

    const misSolicitudes = result.solicitudes.filter(
      s => s.email_usuario.toLowerCase() === userEmail.toLowerCase()
    );

    return { success: true, solicitudes: misSolicitudes };

  } catch (error) {
    Logger.log('❌ Error en getMisSolicitudesRecurrentes: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   APROBAR / RECHAZAR SOLICITUD (ADMIN)
   ============================================ */

/**
 * Aprueba una solicitud y genera las reservas
 * @param {string} idSolicitud - ID de la solicitud
 * @param {string} notasAdmin - Notas opcionales del admin
 */
function aprobarSolicitudRecurrente(idSolicitud, notasAdmin) {
  try {
    Logger.log('✅ Aprobando solicitud: ' + idSolicitud);

    const adminEmail = Session.getActiveUser().getEmail();
    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheet.getDataRange().getValues();

    // Buscar la solicitud
    let filaIndex = -1;
    let solicitud = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS_SOLICITUDES.ID_SOLICITUD] === idSolicitud) {
        filaIndex = i + 1; // +1 porque getRange es 1-indexed
        solicitud = {
          id_solicitud: data[i][COLS_SOLICITUDES.ID_SOLICITUD],
          id_recurso: data[i][COLS_SOLICITUDES.ID_RECURSO],
          nombre_recurso: data[i][COLS_SOLICITUDES.NOMBRE_RECURSO],
          email_usuario: data[i][COLS_SOLICITUDES.EMAIL_USUARIO],
          nombre_usuario: data[i][COLS_SOLICITUDES.NOMBRE_USUARIO],
          dias_semana: data[i][COLS_SOLICITUDES.DIAS_SEMANA],
          id_tramo: data[i][COLS_SOLICITUDES.ID_TRAMO],
          nombre_tramo: data[i][COLS_SOLICITUDES.NOMBRE_TRAMO],
          fecha_inicio: data[i][COLS_SOLICITUDES.FECHA_INICIO],
          fecha_fin: data[i][COLS_SOLICITUDES.FECHA_FIN],
          motivo: data[i][COLS_SOLICITUDES.MOTIVO],
          estado: data[i][COLS_SOLICITUDES.ESTADO]
        };
        break;
      }
    }

    if (!solicitud) throw new Error('Solicitud no encontrada');
    if (solicitud.estado !== 'Pendiente') throw new Error('Esta solicitud ya fue procesada');

    // Generar las reservas
    const resultado = generarReservasDesdeRecurrente(solicitud);

    if (!resultado.success) {
      throw new Error(resultado.error || 'Error al generar reservas');
    }

    // Actualizar estado de la solicitud
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ESTADO + 1).setValue('Aprobada');
    sheet.getRange(filaIndex, COLS_SOLICITUDES.FECHA_RESOLUCION + 1).setValue(new Date());
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ADMIN_RESOLUTOR + 1).setValue(adminEmail);
    sheet.getRange(filaIndex, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue(notasAdmin || '');

    // Purgar caché
    if (typeof purgarCache === 'function') purgarCache();

    // Enviar email al usuario
    enviarEmailSolicitudAprobada({
      email: solicitud.email_usuario,
      nombre: solicitud.nombre_usuario,
      recurso: solicitud.nombre_recurso,
      dias: solicitud.dias_semana,
      tramo: solicitud.nombre_tramo,
      fechaFin: solicitud.fecha_fin,
      reservasCreadas: resultado.reservasCreadas,
      fechasSaltadas: resultado.fechasSaltadas,
      notas: notasAdmin
    });

    Logger.log(`✅ Solicitud ${idSolicitud} aprobada. ${resultado.reservasCreadas} reservas creadas.`);

    return {
      success: true,
      message: `Solicitud aprobada. Se han creado ${resultado.reservasCreadas} reservas.`,
      reservasCreadas: resultado.reservasCreadas,
      fechasSaltadas: resultado.fechasSaltadas
    };

  } catch (error) {
    Logger.log('❌ Error en aprobarSolicitudRecurrente: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Rechaza una solicitud
 * @param {string} idSolicitud - ID de la solicitud
 * @param {string} motivoRechazo - Motivo del rechazo
 */
function rechazarSolicitudRecurrente(idSolicitud, motivoRechazo) {
  try {
    Logger.log('❌ Rechazando solicitud: ' + idSolicitud);

    if (!motivoRechazo || motivoRechazo.trim().length < 5) {
      throw new Error('Indica un motivo para el rechazo');
    }

    const adminEmail = Session.getActiveUser().getEmail();
    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheet.getDataRange().getValues();

    // Buscar la solicitud
    let filaIndex = -1;
    let solicitud = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS_SOLICITUDES.ID_SOLICITUD] === idSolicitud) {
        filaIndex = i + 1;
        solicitud = {
          email_usuario: data[i][COLS_SOLICITUDES.EMAIL_USUARIO],
          nombre_usuario: data[i][COLS_SOLICITUDES.NOMBRE_USUARIO],
          nombre_recurso: data[i][COLS_SOLICITUDES.NOMBRE_RECURSO],
          dias_semana: data[i][COLS_SOLICITUDES.DIAS_SEMANA],
          nombre_tramo: data[i][COLS_SOLICITUDES.NOMBRE_TRAMO],
          estado: data[i][COLS_SOLICITUDES.ESTADO]
        };
        break;
      }
    }

    if (!solicitud) throw new Error('Solicitud no encontrada');
    if (solicitud.estado !== 'Pendiente') throw new Error('Esta solicitud ya fue procesada');

    // Actualizar estado
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ESTADO + 1).setValue('Rechazada');
    sheet.getRange(filaIndex, COLS_SOLICITUDES.FECHA_RESOLUCION + 1).setValue(new Date());
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ADMIN_RESOLUTOR + 1).setValue(adminEmail);
    sheet.getRange(filaIndex, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue(motivoRechazo);

    // Enviar email al usuario
    enviarEmailSolicitudRechazada({
      email: solicitud.email_usuario,
      nombre: solicitud.nombre_usuario,
      recurso: solicitud.nombre_recurso,
      dias: solicitud.dias_semana,
      tramo: solicitud.nombre_tramo,
      motivo: motivoRechazo
    });

    Logger.log('✅ Solicitud ' + idSolicitud + ' rechazada');

    return { success: true, message: 'Solicitud rechazada' };

  } catch (error) {
    Logger.log('❌ Error en rechazarSolicitudRecurrente: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Actualiza las notas de administrador de una solicitud recurrente
 * @param {string} idSolicitud - ID de la solicitud
 * @param {string} notas - Nuevas notas
 */
function actualizarNotasRecurrencia(idSolicitud, notas) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    if (!checkIfAdmin(adminEmail)) {
      return { success: false, error: 'No tienes permisos para esta acción' };
    }

    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheet.getDataRange().getValues();

    // Buscar la solicitud
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS_SOLICITUDES.ID_SOLICITUD] === idSolicitud) {
        sheet.getRange(i + 1, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue(notas || '');
        return { success: true };
      }
    }

    return { success: false, error: 'Solicitud no encontrada' };

  } catch (error) {
    console.error('Error actualizando notas:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Actualiza el motivo de una solicitud de recurrencia
 */
function actualizarMotivoRecurrencia(idSolicitud, motivo) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    if (!checkIfAdmin(adminEmail)) {
      return { success: false, error: 'No tienes permisos para esta acción' };
    }

    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheet.getDataRange().getValues();

    // Buscar la solicitud
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS_SOLICITUDES.ID_SOLICITUD] === idSolicitud) {
        sheet.getRange(i + 1, COLS_SOLICITUDES.MOTIVO + 1).setValue(motivo || '');
        return { success: true };
      }
    }

    return { success: false, error: 'Solicitud no encontrada' };

  } catch (error) {
    Logger.log('Error actualizando motivo: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Cancela una recurrencia aprobada (admin)
 * Cambia el estado a Cancelada y cancela todas las reservas futuras
 * @param {string} idSolicitud - ID de la solicitud
 */
function cancelarRecurrenciaAprobada(idSolicitud) {
  try {
    Logger.log('🚫 Cancelando recurrencia aprobada: ' + idSolicitud);

    const adminEmail = Session.getActiveUser().getEmail();
    if (!checkIfAdmin(adminEmail)) {
      throw new Error('No tienes permisos para esta acción');
    }

    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheet.getDataRange().getValues();

    // Buscar la solicitud
    let filaIndex = -1;
    let solicitud = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS_SOLICITUDES.ID_SOLICITUD] === idSolicitud) {
        filaIndex = i + 1;
        solicitud = {
          email_usuario: data[i][COLS_SOLICITUDES.EMAIL_USUARIO],
          nombre_usuario: data[i][COLS_SOLICITUDES.NOMBRE_USUARIO],
          nombre_recurso: data[i][COLS_SOLICITUDES.NOMBRE_RECURSO],
          estado: data[i][COLS_SOLICITUDES.ESTADO]
        };
        break;
      }
    }

    if (!solicitud) throw new Error('Solicitud no encontrada');
    if (solicitud.estado !== 'Aprobada') throw new Error('Solo se pueden cancelar solicitudes aprobadas');

    // Actualizar estado a Cancelada
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ESTADO + 1).setValue('Cancelada');
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ADMIN_RESOLUTOR + 1).setValue(adminEmail);
    sheet.getRange(filaIndex, COLS_SOLICITUDES.FECHA_RESOLUCION + 1).setValue(new Date());
    sheet.getRange(filaIndex, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue('Cancelada por administrador');

    // Cancelar reservas futuras
    const resultCancelar = cancelarGrupoRecurrente(idSolicitud);

    // Purgar caché
    if (typeof purgarCache === 'function') purgarCache();

    Logger.log(`✅ Recurrencia cancelada: ${resultCancelar.canceladas || 0} reservas eliminadas`);

    // Notificar al usuario
    try {
      enviarEmailCancelacionRecurrente(solicitud, resultCancelar.canceladas || 0);
    } catch (e) {
      Logger.log('⚠️ Error enviando email de cancelación: ' + e.message);
    }

    return {
      success: true,
      message: `Recurrencia cancelada. ${resultCancelar.canceladas || 0} reservas futuras eliminadas.`,
      reservasCanceladas: resultCancelar.canceladas || 0
    };

  } catch (error) {
    Logger.log('❌ Error cancelando recurrencia: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Envía email notificando la cancelación de una recurrencia
 */
function enviarEmailCancelacionRecurrente(solicitud, numReservas) {
  const config = getConfig();
  const tituloApp = config.TITULO_APP || 'Sistema de Reservas';

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #dc2626;">Recurrencia Cancelada</h2>
      <p>Hola ${solicitud.nombre_usuario},</p>
      <p>Tu reserva recurrente ha sido <strong>cancelada</strong> por un administrador:</p>
      <div style="background: #fef2f2; padding: 15px; border-radius: 8px; margin: 15px 0;">
        <p><strong>Recurso:</strong> ${solicitud.nombre_recurso}</p>
        <p><strong>Reservas canceladas:</strong> ${numReservas}</p>
      </div>
      <p>Si tienes alguna duda, contacta con el administrador.</p>
      <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 20px 0;">
      <p style="color: #6b7280; font-size: 12px;">${tituloApp}</p>
    </div>
  `;

  MailApp.sendEmail({
    to: solicitud.email_usuario,
    subject: `${tituloApp} - Recurrencia Cancelada`,
    htmlBody: htmlBody
  });
}

/* ============================================
   GENERAR RESERVAS DESDE SOLICITUD APROBADA
   ============================================ */

/**
 * Genera las reservas individuales a partir de una solicitud aprobada
 * @param {Object} solicitud - Datos de la solicitud
 * Formatos soportados para dias_semana:
 *   - Nuevo: "L:T001,M:T002,X:T001" (cada día con su tramo)
 *   - Antiguo: "L,M,X" (todos los días con id_tramo de la solicitud)
 */
function generarReservasDesdeRecurrente(solicitud) {
  try {
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);

    // --- VERIFICAR Y CREAR COLUMNA ID_Solicitud_Recurrente SI NO EXISTE ---
    const headersActuales = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const headersLower = headersActuales.map(h => h.toString().toLowerCase().trim());

    if (!headersLower.includes('id_solicitud_recurrente')) {
      // La columna no existe, crearla
      const nuevaCol = sheetReservas.getLastColumn() + 1;
      sheetReservas.getRange(1, nuevaCol).setValue('ID_Solicitud_Recurrente');
      Logger.log('✅ Columna ID_Solicitud_Recurrente creada en posición ' + nuevaCol);
    }

    // Mapeo de días: L=1, M=2, X=3, J=4, V=5, S=6, D=0
    const mapaDiasNum = { 'L': 1, 'M': 2, 'X': 3, 'J': 4, 'V': 5, 'S': 6, 'D': 0 };
    const mapaDiasLetra = { 1: 'L', 2: 'M', 3: 'X', 4: 'J', 5: 'V', 6: 'S', 0: 'D' };

    // Parsear días y tramos según el formato
    const diasSemanaStr = String(solicitud.dias_semana || '');
    let seleccionesMap = {}; // { 'L': 'T001', 'M': 'T002', ... }

    if (diasSemanaStr.includes(':')) {
      // Formato nuevo: "L:T001,M:T002"
      diasSemanaStr.split(',').forEach(item => {
        const [dia, tramo] = item.trim().split(':');
        if (dia && tramo) {
          seleccionesMap[dia.toUpperCase()] = tramo;
        }
      });
    } else {
      // Formato antiguo: "L,M,X" con un único tramo
      diasSemanaStr.split(',').forEach(dia => {
        const d = dia.trim().toUpperCase();
        if (d && mapaDiasNum[d] !== undefined) {
          seleccionesMap[d] = solicitud.id_tramo;
        }
      });
    }

    const diasSeleccionados = Object.keys(seleccionesMap).map(d => mapaDiasNum[d]);

    const fechaInicio = new Date(solicitud.fecha_inicio);
    const fechaFin = new Date(solicitud.fecha_fin);

    // Obtener reservas existentes para verificar disponibilidad
    const reservasExistentes = getActiveReservations();

    // Obtener headers de la hoja de reservas para mapear columnas
    const headersReservas = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const headerMap = {};
    headersReservas.forEach((h, i) => {
      headerMap[h.toString().toLowerCase().replace(/ /g, '_')] = i;
    });

    let reservasCreadas = 0;
    let fechasSaltadas = [];
    const nuevasReservas = [];

    // Iterar desde fecha inicio hasta fecha fin
    let fechaActual = new Date(fechaInicio);
    fechaActual.setHours(0, 0, 0, 0);

    while (fechaActual <= fechaFin) {
      const diaSemana = fechaActual.getDay(); // 0=Domingo, 1=Lunes, etc.
      const diaLetra = mapaDiasLetra[diaSemana];

      if (diasSeleccionados.includes(diaSemana) && seleccionesMap[diaLetra]) {
        const fechaISO = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const tramoParaEsteDia = seleccionesMap[diaLetra];

        // Verificar si ya hay una reserva para ese recurso/fecha/tramo
        const yaReservado = reservasExistentes.some(r =>
          String(r.id_recurso) === String(solicitud.id_recurso) &&
          r.fecha === fechaISO &&
          String(r.id_tramo) === String(tramoParaEsteDia)
        );

        if (yaReservado) {
          fechasSaltadas.push(fechaISO + ' (' + tramoParaEsteDia + ')');
        } else {
          // Generar ID único para la reserva
          const idReserva = 'RES-' + Utilities.getUuid().substring(0, 8).toUpperCase();

          // Crear fila de reserva
          const numCols = sheetReservas.getLastColumn();
          const nuevaFila = new Array(numCols).fill('');

          if (headerMap['id_reserva'] !== undefined) nuevaFila[headerMap['id_reserva']] = idReserva;
          if (headerMap['id_recurso'] !== undefined) nuevaFila[headerMap['id_recurso']] = solicitud.id_recurso;
          if (headerMap['email_usuario'] !== undefined) nuevaFila[headerMap['email_usuario']] = solicitud.email_usuario;
          if (headerMap['fecha'] !== undefined) nuevaFila[headerMap['fecha']] = new Date(fechaActual);
          if (headerMap['curso'] !== undefined) nuevaFila[headerMap['curso']] = '';
          if (headerMap['id_tramo'] !== undefined) nuevaFila[headerMap['id_tramo']] = tramoParaEsteDia;
          if (headerMap['cantidad'] !== undefined) nuevaFila[headerMap['cantidad']] = 1;
          if (headerMap['estado'] !== undefined) nuevaFila[headerMap['estado']] = 'Confirmada';
          if (headerMap['notas'] !== undefined) nuevaFila[headerMap['notas']] = 'Reserva recurrente: ' + solicitud.id_solicitud;
          if (headerMap['timestamp'] !== undefined) nuevaFila[headerMap['timestamp']] = new Date();
          if (headerMap['id_solicitud_recurrente'] !== undefined) nuevaFila[headerMap['id_solicitud_recurrente']] = solicitud.id_solicitud;

          nuevasReservas.push(nuevaFila);
          reservasCreadas++;
        }
      }

      // Avanzar al siguiente día
      fechaActual.setDate(fechaActual.getDate() + 1);
    }

    // Insertar todas las reservas de una vez
    if (nuevasReservas.length > 0) {
      sheetReservas.getRange(
        sheetReservas.getLastRow() + 1,
        1,
        nuevasReservas.length,
        nuevasReservas[0].length
      ).setValues(nuevasReservas);
    }

    Logger.log(`✅ Generadas ${reservasCreadas} reservas. Saltadas: ${fechasSaltadas.length}`);

    return {
      success: true,
      reservasCreadas: reservasCreadas,
      fechasSaltadas: fechasSaltadas
    };

  } catch (error) {
    Logger.log('❌ Error en generarReservasDesdeRecurrente: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   CANCELAR RESERVAS RECURRENTES
   ============================================ */

/**
 * Cancela todas las reservas de un grupo recurrente
 * @param {string} idSolicitud - ID de la solicitud recurrente
 */
function cancelarGrupoRecurrente(idSolicitud) {
  try {
    Logger.log('🗑️ Cancelando grupo recurrente: ' + idSolicitud);

    const userEmail = Session.getActiveUser().getEmail();
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);

    // Obtener headers
    const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const colIdSolicitud = headers.findIndex(h => h.toString().toLowerCase().includes('id_solicitud_recurrente'));
    const colEmail = headers.findIndex(h => h.toString().toLowerCase().includes('email_usuario'));
    const colEstado = headers.findIndex(h => h.toString().toLowerCase() === 'estado');
    const colFecha = headers.findIndex(h => h.toString().toLowerCase() === 'fecha');

    if (colIdSolicitud === -1) {
      throw new Error('La columna id_solicitud_recurrente no existe en la hoja de reservas');
    }

    const data = sheetReservas.getDataRange().getValues();
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    let canceladas = 0;

    // Buscar y cancelar reservas futuras del grupo
    for (let i = data.length - 1; i >= 1; i--) {
      const idSolReserva = data[i][colIdSolicitud];
      const emailReserva = data[i][colEmail];
      const estadoReserva = data[i][colEstado];
      const fechaReserva = new Date(data[i][colFecha]);

      if (String(idSolReserva) === String(idSolicitud) &&
          estadoReserva === 'Confirmada' &&
          fechaReserva >= hoy) {

        // Verificar que el usuario es el dueño o es admin
        const esAdmin = checkIfAdmin(userEmail);
        if (emailReserva.toLowerCase() !== userEmail.toLowerCase() && !esAdmin) {
          continue; // Saltar reservas de otros usuarios si no es admin
        }

        // Cancelar la reserva
        sheetReservas.getRange(i + 1, colEstado + 1).setValue('Cancelada');
        canceladas++;
      }
    }

    // Purgar caché
    if (typeof purgarCache === 'function') purgarCache();

    Logger.log(`✅ Canceladas ${canceladas} reservas del grupo ${idSolicitud}`);

    return {
      success: true,
      message: `Se han cancelado ${canceladas} reservas futuras`,
      canceladas: canceladas
    };

  } catch (error) {
    Logger.log('❌ Error en cancelarGrupoRecurrente: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Obtiene las reservas de un grupo recurrente
 * @param {string} idSolicitud - ID de la solicitud recurrente
 */
function getReservasDeGrupoRecurrente(idSolicitud) {
  try {
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);

    const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const colIdSolicitud = headers.findIndex(h => h.toString().toLowerCase().includes('id_solicitud_recurrente'));

    if (colIdSolicitud === -1) {
      return { success: true, reservas: [] };
    }

    const data = sheetReservas.getDataRange().getValues();
    const reservas = [];
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIdSolicitud]) === String(idSolicitud)) {
        const fechaReserva = new Date(data[i][headers.findIndex(h => h.toString().toLowerCase() === 'fecha')]);
        const estado = data[i][headers.findIndex(h => h.toString().toLowerCase() === 'estado')];

        if (estado === 'Confirmada' && fechaReserva >= hoy) {
          reservas.push({
            id_reserva: data[i][headers.findIndex(h => h.toString().toLowerCase() === 'id_reserva')],
            fecha: Utilities.formatDate(fechaReserva, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            id_tramo: data[i][headers.findIndex(h => h.toString().toLowerCase() === 'id_tramo')],
            estado: estado
          });
        }
      }
    }

    // Ordenar por fecha
    reservas.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));

    return { success: true, reservas: reservas };

  } catch (error) {
    Logger.log('❌ Error en getReservasDeGrupoRecurrente: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   CREAR RESERVA RECURRENTE DIRECTAMENTE (ADMIN)
   ============================================ */

/**
 * Crea una reserva recurrente directamente sin pasar por solicitud (solo admin)
 * @param {Object} datos - Datos de la reserva
 */
function crearRecurrenteDirecta(datos) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();

    // Verificar que es admin
    if (!checkIfAdmin(adminEmail)) {
      throw new Error('No tienes permisos para esta acción');
    }

    Logger.log('📝 Admin creando reserva recurrente directa: ' + JSON.stringify(datos));

    // Crear una solicitud ya aprobada
    const ss = getDB();
    const recursos = sheetToObjects(ss.getSheetByName(SHEETS.RECURSOS));
    const recurso = recursos.find(r => String(r.id_recurso) === String(datos.id_recurso));
    if (!recurso) throw new Error('Recurso no encontrado');

    const tramos = sheetToObjects(ss.getSheetByName(SHEETS.TRAMOS));

    // Formato de dias_semana: "V:T001,V:T002,L:T001" (día:tramo)
    // Extraer el primer tramo para mostrar en la solicitud
    const diasStr = Array.isArray(datos.dias_semana) ? datos.dias_semana.join(',') : (datos.dias_semana || '');
    let primerTramoId = datos.id_tramo || '';
    let nombreTramo = 'Múltiples tramos';

    // Si el formato incluye tramos (nuevo formato: "L:T001,M:T002")
    if (diasStr.includes(':')) {
      const primerItem = diasStr.split(',')[0];
      if (primerItem && primerItem.includes(':')) {
        primerTramoId = primerItem.split(':')[1];
      }
    }

    // Buscar el tramo para obtener el nombre
    if (primerTramoId) {
      const tramo = tramos.find(t => String(t.id_tramo) === String(primerTramoId));
      if (tramo) {
        nombreTramo = tramo.nombre_tramo;
      }
    }

    // Email del usuario beneficiario (puede ser otro usuario o el admin)
    const emailBeneficiario = datos.email_usuario || adminEmail;

    // Obtener nombre del beneficiario
    const usuarios = sheetToObjects(ss.getSheetByName(SHEETS.USUARIOS));
    const usuario = usuarios.find(u => u.email_usuario && u.email_usuario.toLowerCase() === emailBeneficiario.toLowerCase());
    const nombreBeneficiario = usuario ? (usuario.nombre_completo || emailBeneficiario) : emailBeneficiario;

    // Generar ID
    const idSolicitud = 'SOL-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    const fechaInicio = datos.fecha_inicio ? new Date(datos.fecha_inicio) : new Date();
    const fechaFin = new Date(datos.fecha_fin);

    // Crear solicitud ya aprobada
    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const nuevaFila = [
      idSolicitud,
      datos.id_recurso,
      recurso.nombre,
      emailBeneficiario,
      nombreBeneficiario,
      diasStr,
      primerTramoId,
      nombreTramo,
      fechaInicio,
      fechaFin,
      datos.motivo || 'Creada directamente por administrador',
      'Aprobada',
      new Date(),
      new Date(),
      adminEmail,
      datos.notas_admin || 'Creación directa'
    ];

    sheet.appendRow(nuevaFila);

    // Generar las reservas
    const solicitud = {
      id_solicitud: idSolicitud,
      id_recurso: datos.id_recurso,
      email_usuario: emailBeneficiario,
      dias_semana: diasStr,
      id_tramo: primerTramoId, // Se usa para formato antiguo, el nuevo usa dias_semana
      fecha_inicio: fechaInicio,
      fecha_fin: fechaFin
    };

    const resultado = generarReservasDesdeRecurrente(solicitud);

    if (!resultado.success) {
      throw new Error(resultado.error);
    }

    // Purgar caché
    if (typeof purgarCache === 'function') purgarCache();

    Logger.log(`✅ Reserva recurrente directa creada: ${resultado.reservasCreadas} reservas`);

    return {
      success: true,
      id_solicitud: idSolicitud,
      reservasCreadas: resultado.reservasCreadas,
      fechasSaltadas: resultado.fechasSaltadas,
      message: `Reserva recurrente creada. ${resultado.reservasCreadas} reservas generadas.`
    };

  } catch (error) {
    Logger.log('❌ Error en crearRecurrenteDirecta: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   OBTENER CONFLICTOS PARA NUEVA RECURRENCIA
   ============================================ */

/**
 * Obtiene los conflictos para un recurso específico:
 * - Recurrencias aprobadas existentes (día:tramo ocupados)
 * - Disponibilidad del recurso (día:tramo no permitidos)
 *
 * @param {string} idRecurso - ID del recurso
 * @returns {Object} { recurrenciasOcupadas: [{dia, tramo, usuario}], nodisponibles: [{dia, tramo}] }
 */
function getConflictosRecurrencia(idRecurso) {
  try {
    const ss = getDB();
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    // 1. RECURRENCIAS APROBADAS ACTIVAS para este recurso
    const sheetSolicitudes = getOrCreateSheetSolicitudesRecurrentes();
    let recurrenciasOcupadas = [];

    if (sheetSolicitudes && sheetSolicitudes.getLastRow() > 1) {
      const solicitudes = sheetToObjects(sheetSolicitudes);

      // Filtrar solo aprobadas, del mismo recurso, y que no hayan terminado
      const activas = solicitudes.filter(sol => {
        if (String(sol.id_recurso) !== String(idRecurso)) return false;
        if ((sol.estado || '').toLowerCase() !== 'aprobada') return false;

        // Verificar que la fecha_fin sea >= hoy
        const fechaFin = sol.fecha_fin ? new Date(sol.fecha_fin) : null;
        if (fechaFin && fechaFin < hoy) return false;

        return true;
      });

      // Extraer día:tramo de cada recurrencia activa
      activas.forEach(sol => {
        const diasSemanaStr = sol.dias_semana || '';

        if (diasSemanaStr.includes(':')) {
          // Formato nuevo: "V:T001,L:T002"
          diasSemanaStr.split(',').forEach(item => {
            const [dia, tramo] = item.trim().split(':');
            if (dia && tramo) {
              recurrenciasOcupadas.push({
                dia: dia.toUpperCase(),
                tramo: tramo,
                usuario: sol.nombre_usuario || sol.email_usuario,
                idSolicitud: sol.id_solicitud
              });
            }
          });
        } else {
          // Formato antiguo: "L,M,X" con un único tramo
          const tramo = sol.id_tramo;
          diasSemanaStr.split(',').forEach(d => {
            const dia = d.trim().toUpperCase();
            if (dia && tramo) {
              recurrenciasOcupadas.push({
                dia: dia,
                tramo: tramo,
                usuario: sol.nombre_usuario || sol.email_usuario,
                idSolicitud: sol.id_solicitud
              });
            }
          });
        }
      });
    }

    // 2. DISPONIBILIDAD DEL RECURSO (qué día:tramo están bloqueados)
    const sheetDisp = ss.getSheetByName(SHEETS.DISPONIBILIDAD);
    let nodisponibles = [];

    if (sheetDisp && sheetDisp.getLastRow() > 1) {
      const disponibilidad = sheetToObjects(sheetDisp);

      // Mapeo de día de semana a letra
      const diasMap = {
        'lunes': 'L', 'martes': 'M', 'miércoles': 'X', 'miercoles': 'X',
        'jueves': 'J', 'viernes': 'V', 'sábado': 'S', 'sabado': 'S', 'domingo': 'D',
        'l': 'L', 'm': 'M', 'x': 'X', 'j': 'J', 'v': 'V', 's': 'S', 'd': 'D'
      };

      disponibilidad.forEach(d => {
        if (String(d.id_recurso) !== String(idRecurso)) return;

        // Si Permitido es false/no/0, está bloqueado
        const permitido = d.permitido;
        const esBloqueado = permitido === false ||
                           permitido === 'false' ||
                           permitido === 'FALSE' ||
                           permitido === 'No' ||
                           permitido === 'no' ||
                           permitido === 'NO' ||
                           permitido === 0 ||
                           permitido === '0';

        if (esBloqueado) {
          const diaRaw = String(d.dia_semana || '').toLowerCase().trim();
          const dia = diasMap[diaRaw] || diaRaw.toUpperCase().charAt(0);
          const tramo = d.id_tramo;

          if (dia && tramo) {
            nodisponibles.push({
              dia: dia,
              tramo: tramo,
              razon: d.razon_bloqueo || 'No disponible'
            });
          }
        }
      });
    }

    Logger.log(`📋 Conflictos para ${idRecurso}: ${recurrenciasOcupadas.length} recurrencias, ${nodisponibles.length} no disponibles`);

    return {
      success: true,
      recurrenciasOcupadas: recurrenciasOcupadas,
      nodisponibles: nodisponibles
    };

  } catch (error) {
    Logger.log('❌ Error en getConflictosRecurrencia: ' + error.message);
    return { success: false, error: error.message, recurrenciasOcupadas: [], nodisponibles: [] };
  }
}

/* ============================================
   OBTENER RESERVAS RECURRENTES DEL USUARIO
   ============================================ */

/**
 * Obtiene las reservas recurrentes activas del usuario actual
 * (agrupadas por solicitud)
 */
function getMisReservasRecurrentes() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);

    // Obtener headers
    const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => { headerMap[h.toString().toLowerCase()] = i; });

    const colIdSolicitud = headers.findIndex(h => h.toString().toLowerCase().includes('id_solicitud_recurrente'));

    if (colIdSolicitud === -1) {
      return { success: true, grupos: [] };
    }

    const data = sheetReservas.getDataRange().getValues();
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    // Agrupar por id_solicitud_recurrente
    const grupos = {};

    for (let i = 1; i < data.length; i++) {
      const email = data[i][headerMap['email_usuario']];
      const estado = data[i][headerMap['estado']];
      const idSolRecurrente = data[i][colIdSolicitud];
      const fecha = new Date(data[i][headerMap['fecha']]);

      if (!idSolRecurrente ||
          email.toLowerCase() !== userEmail.toLowerCase() ||
          estado !== 'Confirmada' ||
          fecha < hoy) {
        continue;
      }

      if (!grupos[idSolRecurrente]) {
        grupos[idSolRecurrente] = {
          id_solicitud: idSolRecurrente,
          id_recurso: data[i][headerMap['id_recurso']],
          id_tramo: data[i][headerMap['id_tramo']],
          reservas: []
        };
      }

      grupos[idSolRecurrente].reservas.push({
        id_reserva: data[i][headerMap['id_reserva']],
        fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      });
    }

    // Convertir a array y ordenar reservas dentro de cada grupo
    const gruposArray = Object.values(grupos).map(g => {
      g.reservas.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));
      g.proximaFecha = g.reservas[0]?.fecha;
      g.totalRestantes = g.reservas.length;
      return g;
    });

    // Ordenar grupos por próxima fecha
    gruposArray.sort((a, b) => new Date(a.proximaFecha) - new Date(b.proximaFecha));

    return { success: true, grupos: gruposArray };

  } catch (error) {
    Logger.log('❌ Error en getMisReservasRecurrentes: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   EMAILS DE NOTIFICACIÓN
   ============================================ */

/**
 * Email al admin cuando hay nueva solicitud
 */
function enviarEmailNuevaSolicitudRecurrente(datos) {
  try {
    const adminEmail = getConfigValue('email_admin', '');
    if (!adminEmail) {
      Logger.log('⚠️ No hay email de admin configurado');
      return;
    }

    const diasLegibles = datos.dias.split(',').map(d => {
      const mapa = { 'L': 'Lunes', 'M': 'Martes', 'X': 'Miércoles', 'J': 'Jueves', 'V': 'Viernes', 'S': 'Sábado', 'D': 'Domingo' };
      return mapa[d.trim().toUpperCase()] || d;
    }).join(', ');

    const fechaFinFormateada = Utilities.formatDate(new Date(datos.fechaFin), Session.getScriptTimeZone(), 'dd/MM/yyyy');

    // URL para aprobar directamente desde el email
    const urlApp = ScriptApp.getService().getUrl();
    const urlAprobar = `${urlApp}?action=aprobar_recurrente&id=${datos.id}`;

    const asunto = `🔄 Nueva solicitud de reserva recurrente - ${datos.recurso}`;
    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #f59e0b; margin-top: 0;">🔄 Nueva Solicitud de Reserva Recurrente</h2>

        <p><strong>${datos.usuario}</strong> (${datos.email}) ha solicitado una reserva recurrente:</p>

        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Recurso</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.recurso}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Días</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${diasLegibles}</td>
          </tr>
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Tramo</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.tramo}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Hasta</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${fechaFinFormateada}</td>
          </tr>
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Motivo</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.motivo}</td>
          </tr>
        </table>

        <div style="text-align: center; margin: 25px 0;">
          <a href="${urlAprobar}" style="display: inline-block; padding: 12px 24px; background-color: #10b981; color: white; text-decoration: none; border-radius: 8px; font-weight: bold; margin-right: 10px;">✓ Aprobar Solicitud</a>
        </div>

        <p style="color: #6b7280; font-size: 0.9em;">O accede al panel de administración para más opciones (rechazar, añadir notas, etc.)</p>
      </div>
    `;

    MailApp.sendEmail({
      to: adminEmail,
      subject: asunto,
      htmlBody: cuerpo
    });

    Logger.log('📧 Email de nueva solicitud enviado a: ' + adminEmail);

  } catch (error) {
    Logger.log('❌ Error enviando email de solicitud: ' + error.message);
  }
}

/**
 * Email al usuario cuando se aprueba su solicitud
 */
function enviarEmailSolicitudAprobada(datos) {
  try {
    const diasLegibles = datos.dias.split(',').map(d => {
      const mapa = { 'L': 'Lunes', 'M': 'Martes', 'X': 'Miércoles', 'J': 'Jueves', 'V': 'Viernes', 'S': 'Sábado', 'D': 'Domingo' };
      return mapa[d.trim().toUpperCase()] || d;
    }).join(', ');

    const fechaFinFormateada = Utilities.formatDate(new Date(datos.fechaFin), Session.getScriptTimeZone(), 'dd/MM/yyyy');

    const asunto = `✅ Solicitud de reserva recurrente aprobada - ${datos.recurso}`;
    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #10b981; margin-top: 0;">✅ Solicitud Aprobada</h2>

        <p>Hola ${datos.nombre},</p>

        <p>Tu solicitud de reserva recurrente ha sido <strong style="color: #10b981;">aprobada</strong>.</p>

        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Recurso</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.recurso}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Días</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${diasLegibles}</td>
          </tr>
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Tramo</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.tramo}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Hasta</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${fechaFinFormateada}</td>
          </tr>
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Reservas creadas</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.reservasCreadas}</td>
          </tr>
        </table>

        ${datos.fechasSaltadas && datos.fechasSaltadas.length > 0 ? `
          <p style="color: #f59e0b;">⚠️ Algunas fechas ya estaban reservadas y se han saltado: ${datos.fechasSaltadas.join(', ')}</p>
        ` : ''}

        ${datos.notas ? `<p><strong>Notas del administrador:</strong> ${datos.notas}</p>` : ''}

        <p style="color: #6b7280;">Puedes ver y gestionar tus reservas desde "Mis Reservas".</p>
      </div>
    `;

    MailApp.sendEmail({
      to: datos.email,
      subject: asunto,
      htmlBody: cuerpo
    });

    Logger.log('📧 Email de aprobación enviado a: ' + datos.email);

  } catch (error) {
    Logger.log('❌ Error enviando email de aprobación: ' + error.message);
  }
}

/**
 * Email al usuario cuando se rechaza su solicitud
 */
function enviarEmailSolicitudRechazada(datos) {
  try {
    const asunto = `❌ Solicitud de reserva recurrente rechazada - ${datos.recurso}`;
    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #ef4444; margin-top: 0;">❌ Solicitud Rechazada</h2>

        <p>Hola ${datos.nombre},</p>

        <p>Lamentamos informarte que tu solicitud de reserva recurrente ha sido <strong style="color: #ef4444;">rechazada</strong>.</p>

        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Recurso</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.recurso}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Días solicitados</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.dias}</td>
          </tr>
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Tramo</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${datos.tramo}</td>
          </tr>
        </table>

        <p><strong>Motivo del rechazo:</strong></p>
        <p style="background: #fef2f2; padding: 15px; border-radius: 8px; border-left: 4px solid #ef4444;">
          ${datos.motivo}
        </p>

        <p style="color: #6b7280;">Si tienes dudas, contacta con el administrador.</p>
      </div>
    `;

    MailApp.sendEmail({
      to: datos.email,
      subject: asunto,
      htmlBody: cuerpo
    });

    Logger.log('📧 Email de rechazo enviado a: ' + datos.email);

  } catch (error) {
    Logger.log('❌ Error enviando email de rechazo: ' + error.message);
  }
}

/* ============================================
   UTILIDADES
   ============================================ */

/**
 * Verifica si un usuario es administrador
 */
function checkIfAdmin(email) {
  try {
    const ss = getDB();
    const sheetUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
    const usuarios = sheetToObjects(sheetUsuarios);

    const usuario = usuarios.find(u =>
      u.email_usuario && u.email_usuario.toLowerCase() === email.toLowerCase()
    );

    if (!usuario) return false;

    const adminValue = String(usuario.admin || '').toLowerCase();
    return adminValue === 'si' || adminValue === 'sí' || adminValue === 'yes' || adminValue === 'true' || adminValue === '1';

  } catch (error) {
    Logger.log('Error verificando admin: ' + error.message);
    return false;
  }
}

/**
 * Contar solicitudes pendientes (para badge)
 */
function contarSolicitudesPendientes() {
  try {
    const result = getSolicitudesRecurrentes('Pendiente');
    return result.success ? result.solicitudes.length : 0;
  } catch (error) {
    return 0;
  }
}

/**
 * Edita los tramos de una recurrencia aprobada
 * @param {string} idSolicitud - ID de la solicitud
 * @param {Array} tramosRestantes - Tramos que se mantienen [{dia, tramo}, ...]
 * @param {Array} tramosEliminados - Tramos que se eliminan [{dia, tramo}, ...]
 */
function editarTramosRecurrencia(idSolicitud, tramosRestantes, tramosEliminados) {
  try {
    if (!isUserAdmin()) {
      return { success: false, error: 'Permiso denegado' };
    }

    const sheetSolicitudes = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheetSolicitudes.getDataRange().getValues();

    // Buscar la solicitud
    let filaEncontrada = -1;
    let solicitud = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS_SOLICITUDES.ID_SOLICITUD] === idSolicitud) {
        filaEncontrada = i + 1;
        solicitud = {
          id_solicitud: data[i][COLS_SOLICITUDES.ID_SOLICITUD],
          id_recurso: data[i][COLS_SOLICITUDES.ID_RECURSO],
          nombre_recurso: data[i][COLS_SOLICITUDES.NOMBRE_RECURSO],
          email_usuario: data[i][COLS_SOLICITUDES.EMAIL_USUARIO],
          nombre_usuario: data[i][COLS_SOLICITUDES.NOMBRE_USUARIO],
          dias_semana: data[i][COLS_SOLICITUDES.DIAS_SEMANA],
          notas_admin: data[i][COLS_SOLICITUDES.NOTAS_ADMIN] || ''
        };
        break;
      }
    }

    if (filaEncontrada < 0 || !solicitud) {
      return { success: false, error: 'Solicitud no encontrada' };
    }

    // Verificar que hay tramos restantes
    if (!tramosRestantes || tramosRestantes.length === 0) {
      return { success: false, error: 'Debe quedar al menos un tramo. Use "Cancelar" para eliminar la recurrencia completa.' };
    }

    // Construir nuevo string de dias_semana
    const nuevosDiasSemana = tramosRestantes.map(t => `${t.dia}:${t.tramo}`).join(',');

    // Actualizar la solicitud
    sheetSolicitudes.getRange(filaEncontrada, COLS_SOLICITUDES.DIAS_SEMANA + 1).setValue(nuevosDiasSemana);

    // Añadir nota de modificación
    const diasNombres = { 'L': 'Lunes', 'M': 'Martes', 'X': 'Miércoles', 'J': 'Jueves', 'V': 'Viernes' };
    const eliminadosTexto = tramosEliminados.map(t => `${diasNombres[t.dia] || t.dia}:${t.tramo}`).join(', ');
    const nuevaNota = solicitud.notas_admin +
      (solicitud.notas_admin ? '\n' : '') +
      `[${new Date().toLocaleDateString('es-ES')}] Tramos eliminados por admin: ${eliminadosTexto}`;
    sheetSolicitudes.getRange(filaEncontrada, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue(nuevaNota);

    // Cancelar las reservas futuras de los tramos eliminados
    const reservasCanceladas = cancelarReservasFuturasDeTramos(
      solicitud.id_recurso,
      idSolicitud,
      tramosEliminados
    );

    // Notificar al usuario
    enviarNotificacionEdicionRecurrencia(solicitud, tramosEliminados, reservasCanceladas);

    return {
      success: true,
      message: `Tramos actualizados. ${reservasCanceladas} reservas futuras canceladas.`
    };

  } catch (error) {
    Logger.log('Error en editarTramosRecurrencia: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Cancela las reservas futuras de tramos específicos de una recurrencia
 */
function cancelarReservasFuturasDeTramos(idRecurso, idSolicitudRecurrente, tramosEliminados) {
  try {
    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);
    if (!sheetReservas) return 0;

    const data = sheetReservas.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());

    const colIdRecurso = headers.indexOf('id_recurso');
    const colFecha = headers.indexOf('fecha');
    const colIdTramo = headers.indexOf('id_tramo');
    const colEstado = headers.indexOf('estado');
    const colIdSolicitudRec = headers.indexOf('id_solicitud_recurrente');

    Logger.log(`[cancelarReservasFuturasDeTramos] Buscando reservas - Recurso: ${idRecurso}, Solicitud: ${idSolicitudRecurrente}`);
    Logger.log(`[cancelarReservasFuturasDeTramos] Tramos a eliminar: ${JSON.stringify(tramosEliminados)}`);
    Logger.log(`[cancelarReservasFuturasDeTramos] Columnas - idRecurso: ${colIdRecurso}, fecha: ${colFecha}, tramo: ${colIdTramo}, estado: ${colEstado}, idSolRec: ${colIdSolicitudRec}`);

    if (colIdRecurso < 0 || colFecha < 0 || colIdTramo < 0 || colEstado < 0) return 0;

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    // Mapeo de día de semana JS a letra
    const jsToLetra = { 0: 'D', 1: 'L', 2: 'M', 3: 'X', 4: 'J', 5: 'V', 6: 'S' };

    let canceladas = 0;

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const estado = String(row[colEstado] || '').toLowerCase().trim();

      // Aceptar tanto 'activa' como 'confirmada'
      if (estado !== 'activa' && estado !== 'confirmada') continue;

      const idRecursoReserva = String(row[colIdRecurso] || '').trim();
      if (idRecursoReserva !== String(idRecurso).trim()) continue;

      // Si existe la columna de solicitud recurrente, verificar que coincida
      if (colIdSolicitudRec >= 0) {
        const idSolRec = String(row[colIdSolicitudRec] || '').trim();
        // Solo filtrar si hay un ID de solicitud en la reserva
        if (idSolRec && idSolRec !== String(idSolicitudRecurrente).trim()) continue;
      }

      const fechaReserva = new Date(row[colFecha]);
      if (isNaN(fechaReserva.getTime())) continue;
      if (fechaReserva < hoy) continue;

      const idTramo = String(row[colIdTramo] || '').trim();
      const diaSemana = jsToLetra[fechaReserva.getDay()];

      // Verificar si este tramo está en los eliminados
      const debeEliminar = tramosEliminados.some(t =>
        t.dia === diaSemana && t.tramo === idTramo
      );

      if (debeEliminar) {
        Logger.log(`[cancelarReservasFuturasDeTramos] Cancelando reserva fila ${i + 1}: ${fechaReserva.toISOString()} - ${diaSemana}:${idTramo}`);
        sheetReservas.getRange(i + 1, colEstado + 1).setValue('Cancelada');
        canceladas++;
      }
    }

    Logger.log(`[cancelarReservasFuturasDeTramos] Total canceladas: ${canceladas}`);
    return canceladas;

  } catch (error) {
    Logger.log('Error en cancelarReservasFuturasDeTramos: ' + error.message);
    return 0;
  }
}

/**
 * Envía notificación al usuario sobre edición de su recurrencia
 */
function enviarNotificacionEdicionRecurrencia(solicitud, tramosEliminados, reservasCanceladas) {
  try {
    const email = solicitud.email_usuario;
    const nombre = solicitud.nombre_usuario || email;
    const recurso = solicitud.nombre_recurso;

    const diasNombres = { 'L': 'Lunes', 'M': 'Martes', 'X': 'Miércoles', 'J': 'Jueves', 'V': 'Viernes' };
    const tramosTexto = tramosEliminados.map(t =>
      `${diasNombres[t.dia] || t.dia} (${t.tramo})`
    ).join(', ');

    const asunto = `[Reservas] Modificación en tu reserva recurrente de ${recurso}`;

    const cuerpo = `
Hola ${nombre},

Te informamos que el administrador ha modificado tu reserva recurrente del recurso "${recurso}".

Los siguientes tramos han sido eliminados:
${tramosTexto}

Se han cancelado ${reservasCanceladas} reserva(s) futura(s) de estos tramos.

El resto de tus tramos de esta recurrencia siguen activos.

Si tienes alguna duda, contacta con el administrador.

Saludos,
Sistema de Reservas
    `.trim();

    MailApp.sendEmail({
      to: email,
      subject: asunto,
      body: cuerpo
    });

    Logger.log(`Email enviado a ${email} sobre edición de recurrencia`);

  } catch (e) {
    Logger.log('Error enviando email de edición recurrencia: ' + e.message);
  }
}

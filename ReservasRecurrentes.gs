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
    enviarEmailNuevaSolicitudRecurrente_({
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

    if (!isUserAdmin()) {
      return { success: false, error: 'No tienes permisos de administrador' };
    }

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
    const resultado = generarReservasDesdeRecurrente_(solicitud);

    if (!resultado.success) {
      throw new Error(resultado.error || 'Error al generar reservas');
    }

    if (resultado.reservasCreadas === 0) {
      const saltadas = resultado.fechasSaltadas || [];
      Logger.log('⚠️ 0 reservas generadas. Fechas saltadas: ' + JSON.stringify(saltadas));
      throw new Error(
        'No se pudo generar ninguna reserva. ' +
        (saltadas.length > 0
          ? `${saltadas.length} fecha(s) ya estaban ocupadas (posiblemente por reservas previas sin cancelar correctamente). ` +
            'Revisa la hoja Reservas y elimina filas huérfanas de recurrencias anteriores.'
          : 'Verifica que el rango de fechas incluya días futuros con los tramos seleccionados.')
      );
    }

    // Actualizar estado de la solicitud
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ESTADO + 1).setValue('Aprobada');
    sheet.getRange(filaIndex, COLS_SOLICITUDES.FECHA_RESOLUCION + 1).setValue(new Date());
    sheet.getRange(filaIndex, COLS_SOLICITUDES.ADMIN_RESOLUTOR + 1).setValue(adminEmail);
    sheet.getRange(filaIndex, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue(notasAdmin || '');

    // Purgar caché
    if (typeof purgarCache === 'function') purgarCache();

    // Enviar email al usuario
    enviarEmailSolicitudAprobada_({
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
      fechasSaltadas: resultado.fechasSaltadas,
      reservas: resultado.reservas || []
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

    if (!isUserAdmin()) {
      return { success: false, error: 'No tienes permisos de administrador' };
    }

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
    enviarEmailSolicitudRechazada_({
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
          dias_semana: data[i][COLS_SOLICITUDES.DIAS_SEMANA],
          nombre_tramo: data[i][COLS_SOLICITUDES.NOMBRE_TRAMO],
          fecha_inicio: data[i][COLS_SOLICITUDES.FECHA_INICIO],
          fecha_fin: data[i][COLS_SOLICITUDES.FECHA_FIN],
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
      enviarEmailCancelacionRecurrente_(solicitud, resultCancelar.canceladas || 0);
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
 * Envía email notificando la cancelación/revocación de una recurrencia
 */
function enviarEmailCancelacionRecurrente_(solicitud, numReservas) {
  Logger.log('📧 Preparando email de revocación...');
  Logger.log('📧 Datos de solicitud: ' + JSON.stringify(solicitud));

  // Obtener título de la app de forma segura (igual que en email de aprobación)
  let tituloApp = 'Sistema de Reservas';
  try {
    const config = getConfiguracion();
    if (config && config.TITULO_APP) {
      tituloApp = config.TITULO_APP;
    }
  } catch (e) {
    Logger.log('⚠️ No se pudo obtener configuración, usando título por defecto');
  }

  // Formatear días y tramos
  const diasMap = { 'L': 'Lunes', 'M': 'Martes', 'X': 'Miércoles', 'J': 'Jueves', 'V': 'Viernes', 'S': 'Sábado', 'D': 'Domingo' };
  const diasSemanaStr = solicitud.dias_semana || '';
  let diasDisplay = '';
  let tramosDisplay = solicitud.nombre_tramo || '';

  if (diasSemanaStr.includes(':')) {
    // Formato nuevo: "L:T001,M:T002,X:T001"
    const diasSet = new Set();
    const tramosSet = new Set();

    diasSemanaStr.split(',').forEach(item => {
      const [d, t] = item.trim().split(':');
      if (d && diasMap[d.toUpperCase()]) diasSet.add(diasMap[d.toUpperCase()]);
      if (t) tramosSet.add(t);
    });

    diasDisplay = Array.from(diasSet).join(', ');

    // Obtener nombres de tramos desde la tabla
    if (tramosSet.size > 0) {
      const tramosIds = Array.from(tramosSet);
      const tramosNombres = obtenerNombresTramos(tramosIds);
      tramosDisplay = tramosNombres.join(', ');
    }
  } else {
    // Formato antiguo
    diasDisplay = diasSemanaStr.split(',').map(d => diasMap[d.trim().toUpperCase()] || d.trim()).join(', ');
  }

  // Formatear fechas
  const formatFecha = (f) => {
    if (!f) return '';
    const fecha = new Date(f);
    return fecha.toLocaleDateString('es-ES', { day: 'numeric', month: 'long', year: 'numeric' });
  };

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #dc2626;">Reserva Recurrente Revocada</h2>
      <p>Hola ${escHtml_(solicitud.nombre_usuario || 'Usuario')},</p>
      <p>Tu reserva recurrente ha sido <strong>revocada</strong> por un administrador.</p>
      <p>Todas las reservas futuras asociadas han sido eliminadas.</p>
      <div style="background: #fef2f2; padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid #dc2626;">
        <p style="margin: 5px 0;"><strong>Recurso:</strong> ${escHtml_(solicitud.nombre_recurso || '')}</p>
        <p style="margin: 5px 0;"><strong>Días:</strong> ${diasDisplay}</p>
        <p style="margin: 5px 0;"><strong>Tramos:</strong> ${tramosDisplay}</p>
        <p style="margin: 5px 0;"><strong>Periodo:</strong> ${formatFecha(solicitud.fecha_inicio)} - ${formatFecha(solicitud.fecha_fin)}</p>
        <p style="margin: 5px 0;"><strong>Reservas eliminadas:</strong> ${numReservas}</p>
      </div>
      <p>Si tienes alguna duda, contacta con el administrador.</p>
      <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 20px 0;">
      <p style="color: #6b7280; font-size: 12px;">${tituloApp}</p>
    </div>
  `;

  const destinatario = solicitud.email_usuario;
  if (!destinatario) {
    Logger.log('❌ Error: No hay email de destinatario');
    throw new Error('No hay email de destinatario');
  }

  Logger.log('📧 Enviando email a: ' + destinatario);

  MailApp.sendEmail({
    to: destinatario,
    subject: `${tituloApp} - Reserva Recurrente Revocada`,
    htmlBody: htmlBody
  });

  Logger.log('📧 Email de revocación enviado correctamente a: ' + destinatario);
}

/**
 * Obtiene los nombres de tramos a partir de sus IDs
 */
function obtenerNombresTramos(tramosIds) {
  try {
    const ss = getDB();
    const sheetTramos = ss.getSheetByName(SHEETS.TRAMOS);
    if (!sheetTramos) return tramosIds;

    const data = sheetTramos.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    const idCol = headers.indexOf('id_tramo');
    const nombreCol = headers.indexOf('nombre_tramo');

    if (idCol === -1 || nombreCol === -1) return tramosIds;

    const tramosMap = {};
    for (let i = 1; i < data.length; i++) {
      tramosMap[data[i][idCol]] = data[i][nombreCol];
    }

    return tramosIds.map(id => tramosMap[id] || id);
  } catch (e) {
    Logger.log('⚠️ Error obteniendo nombres de tramos: ' + e.message);
    return tramosIds;
  }
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
// Envoltorio con LockService: evita dobles reservas / filas desplazadas si coincide con otra escritura
function generarReservasDesdeRecurrente_(solicitud) {
  return conLockScript_(() => generarReservasDesdeRecurrenteSinLock_(solicitud));
}

function generarReservasDesdeRecurrenteSinLock_(solicitud) {
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
    let seleccionesMap = {}; // { 'L': ['T001', 'T002'], 'M': ['T003'], ... }

    if (diasSemanaStr.includes(':')) {
      // Formato nuevo: "L:T001,L:T002,M:T003"
      diasSemanaStr.split(',').forEach(item => {
        const [dia, tramo] = item.trim().split(':');
        if (dia && tramo) {
          const d = dia.toUpperCase();
          if (!seleccionesMap[d]) seleccionesMap[d] = [];
          seleccionesMap[d].push(tramo);
        }
      });
    } else {
      // Formato antiguo: "L,M,X" con un único tramo
      diasSemanaStr.split(',').forEach(dia => {
        const d = dia.trim().toUpperCase();
        if (d && mapaDiasNum[d] !== undefined) {
          seleccionesMap[d] = [solicitud.id_tramo];
        }
      });
    }

    const diasSeleccionados = Object.keys(seleccionesMap).map(d => mapaDiasNum[d]);

    const fechaInicio = new Date(solicitud.fecha_inicio);
    const fechaFin = new Date(solicitud.fecha_fin);

    // Obtener reservas existentes para verificar disponibilidad
    const reservasExistentes = getActiveReservations_();

    // Obtener headers de la hoja de reservas para mapear columnas
    const headersReservas = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const headerMap = {};
    headersReservas.forEach((h, i) => {
      headerMap[h.toString().trim().toLowerCase()] = i;
    });

    let reservasCreadas = 0;
    let fechasSaltadas = [];
    const nuevasReservas = [];
    const reservasObjetos = [];

    Logger.log(`[generarReservas] Rango: ${Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'yyyy-MM-dd')} → ${Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`);
    Logger.log(`[generarReservas] seleccionesMap: ${JSON.stringify(seleccionesMap)}`);
    Logger.log(`[generarReservas] diasSeleccionados: ${JSON.stringify(diasSeleccionados)}`);

    // Iterar desde fecha inicio hasta fecha fin
    let fechaActual = new Date(fechaInicio);
    fechaActual.setHours(0, 0, 0, 0);

    while (fechaActual <= fechaFin) {
      const diaSemana = fechaActual.getDay(); // 0=Domingo, 1=Lunes, etc.
      const diaLetra = mapaDiasLetra[diaSemana];

      if (diasSeleccionados.includes(diaSemana) && seleccionesMap[diaLetra]) {
        const fechaISO = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const tramosParaEsteDia = seleccionesMap[diaLetra]; // Array de tramos

        for (const tramoActual of tramosParaEsteDia) {
          // Verificar si ya hay una reserva para ese recurso/fecha/tramo
          const yaReservado = reservasExistentes.some(r =>
            String(r.id_recurso) === String(solicitud.id_recurso) &&
            r.fecha === fechaISO &&
            String(r.id_tramo) === String(tramoActual)
          );

          if (yaReservado) {
            Logger.log(`[generarReservas] SALTADA: ${fechaISO} ${diaLetra}:${tramoActual} (ya reservado)`);
            fechasSaltadas.push(fechaISO + ' (' + tramoActual + ')');
          } else {
            // Generar ID único para la reserva
            const idReserva = 'RES-' + Utilities.getUuid().substring(0, 8).toUpperCase();

            // Fecha como mediodía UTC (consistente con crearNuevaReserva)
            // Evita desplazamientos de día por diferencias de timezone
            const fechaReserva = new Date(fechaISO + 'T12:00:00Z');

            // Crear fila de reserva
            const numCols = sheetReservas.getLastColumn();
            const nuevaFila = new Array(numCols).fill('');

            if (headerMap['id_reserva'] !== undefined) nuevaFila[headerMap['id_reserva']] = idReserva;
            if (headerMap['id_recurso'] !== undefined) nuevaFila[headerMap['id_recurso']] = solicitud.id_recurso;
            if (headerMap['email_usuario'] !== undefined) nuevaFila[headerMap['email_usuario']] = solicitud.email_usuario;
            if (headerMap['fecha'] !== undefined) nuevaFila[headerMap['fecha']] = fechaReserva;
            if (headerMap['curso'] !== undefined) nuevaFila[headerMap['curso']] = '';
            if (headerMap['id_tramo'] !== undefined) nuevaFila[headerMap['id_tramo']] = tramoActual;
            if (headerMap['cantidad'] !== undefined) nuevaFila[headerMap['cantidad']] = 1;
            if (headerMap['estado'] !== undefined) nuevaFila[headerMap['estado']] = 'Confirmada';
            if (headerMap['notas'] !== undefined) nuevaFila[headerMap['notas']] = 'Reserva recurrente: ' + solicitud.id_solicitud;
            if (headerMap['timestamp'] !== undefined) nuevaFila[headerMap['timestamp']] = new Date();
            if (headerMap['id_solicitud_recurrente'] !== undefined) nuevaFila[headerMap['id_solicitud_recurrente']] = solicitud.id_solicitud;

            nuevasReservas.push(nuevaFila);
            // Objeto para devolver al frontend (mismo formato que getAdminData)
            reservasObjetos.push({
              ID_Reserva: idReserva,
              ID_Recurso: String(solicitud.id_recurso),
              Email_Usuario: String(solicitud.email_usuario),
              Fecha: fechaISO,
              Curso: '',
              ID_Tramo: String(tramoActual),
              Cantidad: 1,
              Estado: 'Confirmada',
              Notas: 'Reserva recurrente: ' + solicitud.id_solicitud,
              ID_Solicitud_Recurrente: String(solicitud.id_solicitud)
            });
            reservasCreadas++;
            Logger.log(`[generarReservas] CREADA: ${fechaISO} ${diaLetra}:${tramoActual} → ${idReserva}`);
          }
        }
      }

      // Avanzar al siguiente día
      fechaActual.setDate(fechaActual.getDate() + 1);
    }

    // Insertar todas las reservas de una vez
    if (nuevasReservas.length > 0) {
      anadirFilas_(sheetReservas, nuevasReservas);
    }

    Logger.log(`✅ Generadas ${reservasCreadas} reservas. Saltadas: ${fechasSaltadas.length}`);

    return {
      success: true,
      reservasCreadas: reservasCreadas,
      fechasSaltadas: fechasSaltadas,
      reservas: reservasObjetos
    };

  } catch (error) {
    Logger.log('❌ Error en generarReservasDesdeRecurrente_: ' + error.message);
    return { success: false, error: error.message };
  }
}

/* ============================================
   CANCELAR RESERVAS RECURRENTES
   ============================================ */

/**
 * Elimina todas las reservas futuras de un grupo recurrente
 * @param {string} idSolicitud - ID de la solicitud recurrente
 */
// Envoltorio con LockService: evita dobles reservas / filas desplazadas si coincide con otra escritura
function cancelarGrupoRecurrente(idSolicitud) {
  return conLockScript_(() => cancelarGrupoRecurrenteSinLock_(idSolicitud));
}

function cancelarGrupoRecurrenteSinLock_(idSolicitud) {
  try {
    Logger.log('🗑️ Eliminando reservas del grupo recurrente: ' + idSolicitud);

    const userEmail = Session.getActiveUser().getEmail();
    let esAdmin = null;
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

    // Recolectar filas a eliminar (de abajo hacia arriba para no afectar índices)
    const filasAEliminar = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const idSolReserva = data[i][colIdSolicitud];
      const emailReserva = data[i][colEmail];
      const estadoReserva = data[i][colEstado];
      const fechaReserva = new Date(data[i][colFecha]);

      const estadoNorm = String(estadoReserva || '').toLowerCase().trim();
      if (String(idSolReserva) === String(idSolicitud) &&
          estadoNorm === 'confirmada' &&
          fechaReserva >= hoy) {

        // Verificar que el usuario es el dueño o es admin
        if (esAdmin === null) esAdmin = checkIfAdmin(userEmail); // una sola lectura de Usuarios
        if (String(emailReserva || '').toLowerCase() !== String(userEmail).toLowerCase() && !esAdmin) {
          continue; // Saltar reservas de otros usuarios si no es admin
        }

        // Agregar fila a la lista de eliminación
        filasAEliminar.push(i + 1); // +1 porque los índices de Sheets son 1-based
      }
    }

    // Eliminar filas de abajo hacia arriba (ya están ordenadas así)
    let eliminadas = 0;
    for (const fila of filasAEliminar) {
      sheetReservas.deleteRow(fila);
      eliminadas++;
    }

    // Purgar caché
    if (typeof purgarCache === 'function') purgarCache();

    Logger.log(`✅ Eliminadas ${eliminadas} reservas del grupo ${idSolicitud}`);

    return {
      success: true,
      message: `Se han eliminado ${eliminadas} reservas futuras`,
      canceladas: eliminadas
    };

  } catch (error) {
    Logger.log('❌ Error en cancelarGrupoRecurrente: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Elimina un tramo específico de una recurrencia aprobada
 * Si era el último tramo, cancela la recurrencia completa
 * @param {string} idSolicitud - ID de la solicitud recurrente
 * @param {string} diaLetra - Letra del día (L, M, X, J, V, S, D)
 * @param {string} idTramo - ID del tramo a eliminar
 */
// Envoltorio con LockService: evita dobles reservas / filas desplazadas si coincide con otra escritura
function eliminarTramoDeRecurrencia(idSolicitud, diaLetra, idTramo) {
  return conLockScript_(() => eliminarTramoDeRecurrenciaSinLock_(idSolicitud, diaLetra, idTramo));
}

function eliminarTramoDeRecurrenciaSinLock_(idSolicitud, diaLetra, idTramo) {
  try {
    Logger.log(`🔧 Eliminando tramo ${diaLetra}:${idTramo} de recurrencia ${idSolicitud}`);

    const adminEmail = Session.getActiveUser().getEmail();
    if (!checkIfAdmin(adminEmail)) {
      throw new Error('No tienes permisos para esta acción');
    }

    // 1. Buscar la solicitud
    const sheet = getOrCreateSheetSolicitudesRecurrentes();
    const data = sheet.getDataRange().getValues();

    let filaIndex = -1;
    let solicitud = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COLS_SOLICITUDES.ID_SOLICITUD]) === String(idSolicitud)) {
        filaIndex = i + 1;
        solicitud = {
          dias_semana: String(data[i][COLS_SOLICITUDES.DIAS_SEMANA] || ''),
          estado: String(data[i][COLS_SOLICITUDES.ESTADO] || '')
        };
        break;
      }
    }

    if (!solicitud) throw new Error('Solicitud no encontrada');
    if (solicitud.estado !== 'Aprobada') throw new Error('Solo se pueden modificar solicitudes aprobadas');

    // 2. Parsear dias_semana y eliminar el par día:tramo
    const diasStr = solicitud.dias_semana;
    let items = [];

    if (diasStr.includes(':')) {
      items = diasStr.split(',').map(s => s.trim()).filter(s => s);
    } else {
      // Formato antiguo: convertir a formato nuevo para poder eliminar
      items = diasStr.split(',').map(d => {
        const dd = d.trim().toUpperCase();
        return dd ? `${dd}:${idTramo}` : '';
      }).filter(s => s);
    }

    const target = `${diaLetra.toUpperCase()}:${idTramo}`;
    const itemsNuevos = items.filter(item => item.toUpperCase() !== target.toUpperCase());

    if (itemsNuevos.length === items.length) {
      throw new Error(`El tramo ${diaLetra}:${idTramo} no existe en esta recurrencia`);
    }

    // 3. Eliminar reservas futuras de ese día+tramo específico
    const mapaDiasNum = { 'L': 1, 'M': 2, 'X': 3, 'J': 4, 'V': 5, 'S': 6, 'D': 0 };
    const diaSemanaNum = mapaDiasNum[diaLetra.toUpperCase()];

    const ss = getDB();
    const sheetReservas = ss.getSheetByName(SHEETS.RESERVAS);
    const headers = sheetReservas.getRange(1, 1, 1, sheetReservas.getLastColumn()).getValues()[0];
    const colIdSolicitud = headers.findIndex(h => h.toString().toLowerCase().includes('id_solicitud_recurrente'));
    const colIdTramo = headers.findIndex(h => h.toString().toLowerCase() === 'id_tramo');
    const colEstado = headers.findIndex(h => h.toString().toLowerCase() === 'estado');
    const colFecha = headers.findIndex(h => h.toString().toLowerCase() === 'fecha');

    const dataReservas = sheetReservas.getDataRange().getValues();
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    const filasAEliminar = [];
    for (let i = dataReservas.length - 1; i >= 1; i--) {
      const idSolReserva = String(dataReservas[i][colIdSolicitud] || '');
      const tramoReserva = String(dataReservas[i][colIdTramo] || '').trim();
      const estadoReserva = String(dataReservas[i][colEstado] || '').toLowerCase().trim();
      const fechaReserva = new Date(dataReservas[i][colFecha]);

      if (idSolReserva === String(idSolicitud) &&
          tramoReserva === String(idTramo) &&
          estadoReserva === 'confirmada' &&
          fechaReserva >= hoy &&
          fechaReserva.getDay() === diaSemanaNum) {
        filasAEliminar.push(i + 1);
      }
    }

    for (const fila of filasAEliminar) {
      sheetReservas.deleteRow(fila);
    }

    Logger.log(`✅ Eliminadas ${filasAEliminar.length} reservas de ${diaLetra}:${idTramo}`);

    // 4. Actualizar o cancelar la recurrencia
    let recurrenciaCancelada = false;

    if (itemsNuevos.length === 0) {
      // No quedan tramos: cancelar la recurrencia completa
      sheet.getRange(filaIndex, COLS_SOLICITUDES.ESTADO + 1).setValue('Cancelada');
      sheet.getRange(filaIndex, COLS_SOLICITUDES.ADMIN_RESOLUTOR + 1).setValue(adminEmail);
      sheet.getRange(filaIndex, COLS_SOLICITUDES.FECHA_RESOLUCION + 1).setValue(new Date());
      sheet.getRange(filaIndex, COLS_SOLICITUDES.NOTAS_ADMIN + 1).setValue('Cancelada al eliminar último tramo');
      recurrenciaCancelada = true;
      Logger.log('🚫 Recurrencia cancelada: no quedan tramos');
    } else {
      // Actualizar dias_semana con los items restantes
      sheet.getRange(filaIndex, COLS_SOLICITUDES.DIAS_SEMANA + 1).setValue(itemsNuevos.join(','));
      Logger.log(`📝 dias_semana actualizado: ${itemsNuevos.join(',')}`);
    }

    if (typeof purgarCache === 'function') purgarCache();

    return {
      success: true,
      message: recurrenciaCancelada
        ? 'Último tramo eliminado. Recurrencia cancelada.'
        : `Tramo eliminado. Quedan ${itemsNuevos.length} tramo(s).`,
      reservasEliminadas: filasAEliminar.length,
      recurrenciaCancelada: recurrenciaCancelada,
      diasSemanaActualizado: itemsNuevos.join(',')
    };

  } catch (error) {
    Logger.log('❌ Error en eliminarTramoDeRecurrencia: ' + error.message);
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

    const resultado = generarReservasDesdeRecurrente_(solicitud);

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
   EMAILS DE NOTIFICACIÓN
   ============================================ */

/**
 * Email al admin cuando hay nueva solicitud
 */
function enviarEmailNuevaSolicitudRecurrente_(datos) {
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

        <p><strong>${escHtml_(datos.usuario)}</strong> (${escHtml_(datos.email)}) ha solicitado una reserva recurrente:</p>

        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Recurso</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${escHtml_(datos.recurso)}</td>
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
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${escHtml_(datos.motivo)}</td>
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
function enviarEmailSolicitudAprobada_(datos) {
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

        <p>Hola ${escHtml_(datos.nombre)},</p>

        <p>Tu solicitud de reserva recurrente ha sido <strong style="color: #10b981;">aprobada</strong>.</p>

        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Recurso</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${escHtml_(datos.recurso)}</td>
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

        ${datos.notas ? `<p><strong>Notas del administrador:</strong> ${escHtml_(datos.notas)}</p>` : ''}

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
function enviarEmailSolicitudRechazada_(datos) {
  try {
    const asunto = `❌ Solicitud de reserva recurrente rechazada - ${datos.recurso}`;
    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #ef4444; margin-top: 0;">❌ Solicitud Rechazada</h2>

        <p>Hola ${escHtml_(datos.nombre)},</p>

        <p>Lamentamos informarte que tu solicitud de reserva recurrente ha sido <strong style="color: #ef4444;">rechazada</strong>.</p>

        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr style="background: #f3f4f6;">
            <td style="padding: 10px; border: 1px solid #e5e7eb;"><strong>Recurso</strong></td>
            <td style="padding: 10px; border: 1px solid #e5e7eb;">${escHtml_(datos.recurso)}</td>
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
          ${escHtml_(datos.motivo)}
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
  // Misma regla que el resto de la app (respeta 'Activo'; antes un admin desactivado seguía pasando)
  if (!email) return false;
  return checkUserAuthorization(email).isAdmin === true;
}

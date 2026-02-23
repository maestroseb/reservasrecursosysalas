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


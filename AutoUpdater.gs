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
const UPDATE_SERVER_URL = ''; // ← PEGAR AQUÍ la URL de tu servidor de versiones

/**
 * URL base del repositorio en GitHub (para descargar código fuente).
 * Se usa la API raw de GitHub para obtener los archivos .gs y .html
 */
const GITHUB_REPO_RAW = 'https://raw.githubusercontent.com/maestroseb/reservasrecursosysalas/main';

/**
 * Lista de archivos que componen el sistema y deben actualizarse.
 * 'name' es el nombre del archivo en GitHub (con extensión).
 * 'type' es el tipo de archivo en la API de Apps Script:
 *   - 'SERVER_JS' para archivos .gs
 *   - 'HTML' para archivos .html
 * 'scriptName' es el nombre SIN extensión que usa la API de Apps Script.
 */
const UPDATABLE_FILES = [
  { name: 'Codigo.gs', type: 'SERVER_JS', scriptName: 'Codigo' },
  { name: 'AdminFunctions.gs', type: 'SERVER_JS', scriptName: 'AdminFunctions' },
  { name: 'Setup.gs', type: 'SERVER_JS', scriptName: 'Setup' },
  { name: 'Incidencias.gs', type: 'SERVER_JS', scriptName: 'Incidencias' },
  { name: 'ReservasRecurrentes.gs', type: 'SERVER_JS', scriptName: 'ReservasRecurrentes' },
  { name: 'AutoUpdater.gs', type: 'SERVER_JS', scriptName: 'AutoUpdater' },
  { name: 'index.html', type: 'HTML', scriptName: 'index' },
  { name: 'admin-panel.html', type: 'HTML', scriptName: 'admin-panel' },
  { name: 'admin-scripts.html', type: 'HTML', scriptName: 'admin-scripts' },
  { name: 'scripts.html', type: 'HTML', scriptName: 'scripts' },
  { name: 'styles.html', type: 'HTML', scriptName: 'styles' },
  { name: 'Sidebar.html', type: 'HTML', scriptName: 'Sidebar' },
  { name: 'ActivacionSistema.html', type: 'HTML', scriptName: 'ActivacionSistema' },
  { name: 'registro.html', type: 'HTML', scriptName: 'registro' }
];


/* ============================================
   COMPROBACIÓN DE ACTUALIZACIONES
   ============================================ */

/**
 * Comprueba si hay una nueva versión disponible consultando el servidor de versiones.
 * @returns {Object} { hayActualizacion, versionLocal, versionRemota, changelog, ... }
 */
function comprobarActualizaciones() {
  try {
    if (!UPDATE_SERVER_URL) {
      Logger.log('⚠️ UPDATE_SERVER_URL no configurada.');
      return { hayActualizacion: false, error: 'Servidor de actualizaciones no configurado' };
    }

    const response = UrlFetchApp.fetch(UPDATE_SERVER_URL, {
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
      const manifestResp = UrlFetchApp.fetch(GITHUB_REPO_RAW + '/appsscript.json', { muteHttpExceptions: true });
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

  for (const archivo of UPDATABLE_FILES) {
    try {
      const url = GITHUB_REPO_RAW + '/' + archivo.name;
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

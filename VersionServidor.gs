/**
 * ===========================================================================
 * 🖥️ SERVIDOR DE VERSIONES - SCRIPT INDEPENDIENTE PARA EL AUTOR
 * ===========================================================================
 *
 * ⚠️ ESTE ARCHIVO NO SE INCLUYE EN LAS COPIAS DE LOS USUARIOS.
 * Es un script independiente que TÚ (el autor) despliegas por separado.
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 * ─────────────────────────────
 * 1. Ve a https://script.google.com y crea un proyecto nuevo (sin hoja)
 * 2. Pega SOLO el contenido de este archivo
 * 3. Haz clic en "Implementar" → "Nueva implementación"
 * 4. Tipo: "Aplicación web"
 *    - Ejecutar como: YO MISMO
 *    - Quién tiene acceso: CUALQUIERA
 * 5. Copia la URL generada
 * 6. Pega esa URL en la constante UPDATE_SERVER_URL de "AutoUpdater.gs"
 *    en tu repositorio principal
 *
 * CÓMO PUBLICAR UNA NUEVA VERSIÓN:
 * ─────────────────────────────────
 * 1. Actualiza VERSION_INFO más abajo con el nuevo número y changelog
 * 2. Guarda el script (Ctrl+S)
 * 3. IMPORTANTE: Crea una nueva implementación (o actualiza la existente)
 *    Para actualizar sin cambiar la URL:
 *    - "Implementar" → "Gestionar implementaciones" → Editar → "Nueva versión"
 * 4. Actualiza SYSTEM_VERSION en AutoUpdater.gs del repo principal
 * 5. Haz push a GitHub
 *
 * Así de simple: cada vez que publiques una versión, editas este archivo
 * y todas las copias sabrán que hay una actualización.
 */

/* ============================================
   INFORMACIÓN DE LA VERSIÓN ACTUAL
   ============================================ */

const VERSION_INFO = {
  // ⚡ Número de versión actual (incrementar con cada release)
  version: '1.5.0',

  // 📋 Descripción de cambios (se muestra en el email y diálogo de actualización)
  changelog: [
    '🔄 Sistema de actualización automática vía API de Apps Script',
    '📋 Comprobación diaria de nuevas versiones con notificación al admin',
    '⚡ Actualización con un clic desde el menú del sistema'
  ].join('\n'),

  // 📅 Fecha de publicación
  fechaPublicacion: '2026-03-02',

  // 🔴 ¿Es una actualización crítica? (true = email con prioridad alta)
  critica: false,

  // 🔗 URL de descarga o más información (opcional)
  urlDescarga: 'https://github.com/maestroseb/reservasrecursosysalas',

  // 📌 Versión mínima compatible (opcional - para avisar a los muy desactualizados)
  versionMinima: '1.0.0'
};


/* ============================================
   HISTORIAL DE VERSIONES (OPCIONAL)
   Útil para mostrar changelogs completos
   ============================================ */

const VERSION_HISTORY = [
  {
    version: '1.5.0',
    fecha: '2026-03-02',
    cambios: [
      '🔄 Sistema de actualización automática vía API de Apps Script',
      '📋 Comprobación diaria de nuevas versiones con notificación al admin',
      '⚡ Actualización con un clic desde el menú del sistema'
    ]
  },
  {
    version: '1.3.0',
    fecha: '2026-02-15',
    cambios: [
      '🔁 Sistema de reservas recurrentes',
      '📱 Mejoras de interfaz',
      '🔔 Notificaciones mejoradas'
    ]
  }
  // Añade más entradas según publiques versiones...
];


/* ============================================
   ENDPOINT WEB (NO TOCAR)
   ============================================ */

/**
 * Responde a las peticiones GET con la información de versión en JSON.
 * Las copias del sistema consultan esta URL para saber si hay actualizaciones.
 */
function doGet(e) {
  // Permitir filtrar por parámetro de acción
  const action = e && e.parameter ? e.parameter.action : 'check';

  let responseData;

  switch (action) {
    case 'history':
      // Devolver historial completo de versiones
      responseData = {
        current: VERSION_INFO,
        history: VERSION_HISTORY
      };
      break;

    case 'check':
    default:
      // Respuesta estándar: solo versión actual
      responseData = {
        version: VERSION_INFO.version,
        changelog: VERSION_INFO.changelog,
        fechaPublicacion: VERSION_INFO.fechaPublicacion,
        critica: VERSION_INFO.critica,
        urlDescarga: VERSION_INFO.urlDescarga,
        versionMinima: VERSION_INFO.versionMinima,
        timestamp: new Date().toISOString()
      };
      break;
  }

  // Devolver JSON con cabeceras CORS para permitir acceso desde cualquier origen
  return ContentService
    .createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Endpoint POST (opcional) - Para estadísticas de instalaciones.
 * Las copias pueden reportar su versión para que sepas cuántas hay.
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Registrar en una hoja de cálculo de estadísticas (opcional)
    // Si quieres trackear instalaciones, crea una hoja y descomenta:
    //
    // const ss = SpreadsheetApp.openById('TU_ID_DE_HOJA_ESTADISTICAS');
    // const sheet = ss.getSheetByName('Instalaciones') || ss.insertSheet('Instalaciones');
    // sheet.appendRow([
    //   new Date(),
    //   data.version || '?',
    //   data.domain || '?',
    //   data.action || 'ping'
    // ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

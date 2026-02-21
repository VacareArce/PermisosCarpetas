/**
 * @OnlyCurrentDoc
 *
 * Este script audita los permisos de una Unidad Compartida (Shared Drive)
 * y registra en la hoja de cálculo activa:
 * 1. Los permisos base de la Unidad Compartida (incluyendo públicos/dominio).
 * 2. Cualquier archivo o carpeta interna con permisos DIFERENTES a los de su padre.
 *
 * Formato de Registro:
 * - Ruta, Link, Tipo, Usuarios con Permisos Asignados (Rol)
 * - Los usuarios se listan en una sola celda, separados por comas.
 *
 * Ejecución:
 * - MANUAL. Límite de 20 minutos (para cuentas Workspace).
 * - El usuario debe ejecutar "Continuar Auditoría" para seguir.
 *
 * NOTA: Este script ahora usa una hoja "Queue_STATE" para manejar colas muy grandes
 * que exceden los límites de PropertiesService.
 */

// --- Constantes Globales ---
const SHEET_NAME = 'Reporte de Permisos';
const QUEUE_SHEET_NAME = 'Queue_STATE'; // Hoja para la cola
// Límite manual para cuentas Workspace es 30 min. Usamos 20 min.
const MAX_RUNTIME_MS = 20 * 60 * 1000;
// PropertiesService YA NO SE USA para la cola, solo para limpieza (opcional).
const LEGACY_STATE_PROPERTY_KEY = 'DRIVE_AUDIT_STATE';

/**
 * Crea un menú personalizado en la hoja de cálculo al abrirla.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Auditoría de Drive')
    .addItem('Iniciar Auditoría de Permisos', 'startAudit')
    .addItem('Continuar Auditoría', 'processQueue')
    .addSeparator()
    .addItem('Limpiar Estado', 'clearState')
    .addToUi();
}

/**
 * Borra el estado guardado (la hoja de cola) para reiniciar la auditoría.
 * @param {boolean} silent - Si es true, no muestra la alerta al usuario.
 */
function clearState(silent) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Borrar la hoja de cola
  const queueSheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  if (queueSheet) {
    ss.deleteSheet(queueSheet);
  }
  
  // Borrar estado antiguo de PropertiesService (por si acaso)
  PropertiesService.getScriptProperties().deleteProperty(LEGACY_STATE_PROPERTY_KEY);
  
  Logger.log('Estado limpiado.');
  if (!silent) {
    ui.alert('Estado de auditoría limpiado (Hoja de Cola eliminada). Puede iniciar una nueva.');
  }
}

/**
 * Obtiene el conjunto de permisos "simplificado" (Editores, Visualizadores, Público, Dominio)
 * para una rápida comparación de herencia.
 * @param {File|Folder} driveFile - El archivo o carpeta.
 * @return {Object} Un objeto con { editors: [], viewers: [], publicAccess: 'Ninguno', domainAccess: 'Ninguno' }.
 */
function getPermissionSet(driveFile) {
  // ... (Esta función no cambia) ...
  const editors = new Set();
  const viewers = new Set();
  let publicAccess = 'Ninguno';
  let domainAccess = 'Ninguno';

  try {
    driveFile.getEditors().forEach(user => editors.add(user.getEmail()));
  } catch (e) { /* Ignorar errores de permisos */ }
  try {
    driveFile.getViewers().forEach(user => viewers.add(user.getEmail()));
  } catch (e) { /* Ignorar errores de permisos */ }
  
  try {
    const sharingAccess = driveFile.getSharingAccess();
    const sharingPermission = driveFile.getSharingPermission();

    if (sharingAccess === DriveApp.Access.ANYONE || sharingAccess === DriveApp.Access.ANYONE_WITH_LINK) {
      if (sharingPermission === DriveApp.Permission.EDIT) {
        publicAccess = 'Editor';
      } else if (sharingPermission === DriveApp.Permission.VIEW) {
        publicAccess = 'Visualizador';
      }
    } else if (sharingAccess === DriveApp.Access.DOMAIN || sharingAccess === DriveApp.Access.DOMAIN_WITH_LINK) {
      if (sharingPermission === DriveApp.Permission.EDIT) {
        domainAccess = 'Editor';
      } else if (sharingPermission === DriveApp.Permission.VIEW) {
        domainAccess = 'Visualizador';
      }
    }
  } catch (e) { /* Ignorar errores de permisos */ }

  const finalViewers = [...viewers].filter(v => !editors.has(v));
  
  return {
    editors: [...editors],
    viewers: finalViewers,
    publicAccess: publicAccess,
    domainAccess: domainAccess
  };
}

/**
 * Obtiene el nivel de acceso simplificado de un usuario.
 * @param {string} userEmail - El email del usuario.
 * @param {Object} permissionSet - El conjunto de permisos.
 * @return {string} 'Editor', 'Visualizador', o 'Ninguno'.
 */
function getAccess(userEmail, permissionSet) {
  // ... (Esta función no cambia) ...
  if (permissionSet.editors.includes(userEmail)) return 'Editor';
  if (permissionSet.viewers.includes(userEmail)) return 'Visualizador';
  return 'Ninguno';
}

/**
 * Función principal que solicita el ID y CONFIGURA la primera ejecución.
 */
function startAudit() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Limpiar estado anterior
  clearState(true);

  // 2. Obtener el ID de la Unidad Compartida
  // ... (Esta sección no cambia) ...
  const response = ui.prompt('ID de la Unidad Compartida', 'Ingresa el ID de la Unidad Compartida a auditar:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) {
    ui.alert('Auditoría cancelada.');
    return;
  }
  const driveId = response.getResponseText().trim();

  // 3. Preparar la hoja de cálculo de resultados
  // ... (Esta sección no cambia) ...
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  const headers = ['Ruta', 'Link', 'Tipo', 'Usuarios con Permisos Asignados (Rol)'];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // 3b. Preparar la HOJA DE COLA
  let queueSheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  if (queueSheet) {
    ss.deleteSheet(queueSheet); // Limpiar por si acaso
  }
  queueSheet = ss.insertSheet(QUEUE_SHEET_NAME);
  queueSheet.hideSheet(); // Ocultar la hoja
  const queueHeaders = ['ID', 'Ruta', 'Link', 'Permisos Padre (JSON)', 'Es Raíz (bool)'];
  queueSheet.appendRow(queueHeaders);
  

  let driveName = "Unidad Compartida";
  let rootFolder;
  let rootUrl;

  // 4. Verificar acceso y registrar permisos raíz usando Drive API
  // ... (Esta sección no cambia) ...
  try {
    // Verificar acceso con Drive API y obtener el nombre
    const driveInfo = Drive.Drives.get(driveId, { useDomainAdminAccess: true });
    driveName = driveInfo.name;

    // Obtener la carpeta raíz (el ID de la unidad es el ID de su carpeta raíz)
    rootFolder = DriveApp.getFolderById(driveId);
    rootUrl = rootFolder.getUrl();

    // Obtener permisos detallados de la raíz
    const permissions = Drive.Permissions.list(driveId, {
      supportsAllDrives: true,
      useDomainAdminAccess: true,
      fields: 'permissions(id, emailAddress, role, type, domain, permissionDetails)'
    });

    sheet.appendRow([driveName, rootUrl, 'Unidad Compartida (Raíz)', '--- Permisos Base de la Unidad ---']);

    if (permissions.permissions && permissions.permissions.length > 0) {
      // Agrupar usuarios por rol
      const roles = {};
      permissions.permissions.forEach(p => {
        // Ignorar permisos heredados en el nivel raíz
        if (p.permissionDetails && p.permissionDetails[0] && p.permissionDetails[0].inherited) {
            return;
        }

        const roleName = p.role.charAt(0).toUpperCase() + p.role.slice(1); // Capitalizar (e.g., organizer)
        let userIdentifier;
        
        if (p.type === 'user') {
          userIdentifier = p.emailAddress;
        } else if (p.type === 'group') {
          userIdentifier = p.emailAddress || `Grupo (ID: ${p.id})`;
        } else if (p.type === 'anyone') {
          userIdentifier = 'Público (Cualquiera con el enlace)';
        } else if (p.type === 'domain') {
          userIdentifier = `Dominio (${p.domain || 'tu dominio'})`;
        } else {
          userIdentifier = `${p.type} (ID: ${p.id})`;
        }
        
        const userWithRole = `${userIdentifier} - ${roleName}`;
        
        if (!roles[roleName]) {
          roles[roleName] = [];
        }
        roles[roleName].push(userIdentifier);
      });

      // Escribir los permisos agrupados
      for (const role in roles) {
         const usersWithRole = roles[role].map(email => `${email} - ${role}`).join(', ');
         sheet.appendRow([driveName, rootUrl, `Permiso: ${role}`, usersWithRole]);
      }

    } else {
      sheet.appendRow([driveName, rootUrl, 'Unidad Compartida (Raíz)', 'No se encontraron permisos explícitos.']);
    }

    sheet.appendRow(['---', '---', '---', '---']); // Separador
    sheet.appendRow(['A continuación, se listan solo archivos/carpetas con permisos DIFERENTES a su padre:']);
    sheet.appendRow(['---', '---', '---', '---']); // Separador


  } catch (e) {
    Logger.log(`Error al acceder a la Unidad Compartida: ${e.message}`);
    const errorMsg = `No se pudo encontrar o acceder a la Unidad Compartida. Error: ${e.message}`;
    // Escribir el error en la hoja para diagnóstico
    sheet.appendRow([driveId, 'ERROR', 'Error de Acceso', errorMsg]);
    
    // Mostrar alerta al usuario
    if (e.message.includes("Drive API has not been used")) {
      ui.alert(`${errorMsg}\n\nAcción requerida: Habilite la 'Drive API' en la consola de Google Cloud para este proyecto.`);
    } else if (e.message.includes("Forbidden")) {
      ui.alert(`${errorMsg}\n\nAsegúrese de tener permisos de 'Administrador' sobre esta Unidad Compartida.`);
    } else {
      ui.alert(errorMsg);
    }
    return; // Detener ejecución
  }

  // 5. Preparar la cola de procesamiento inicial
  const rootPermissions = getPermissionSet(rootFolder);
  
  // Añadir la raíz a la HOJA DE COLA
  queueSheet.appendRow([
    rootFolder.getId(),
    driveName, // Usar el nombre de la unidad como raíz de la ruta
    rootUrl,
    JSON.stringify(rootPermissions), // Convertir a JSON para guardar en celda
    true // Marcar como raíz
  ]);

  // 6. Guardar estado y alertar al usuario
  // Ya no se usa PropertiesService
  
  // 7. Iniciar el primer lote de procesamiento
  ui.alert('Auditoría iniciada. El script procesará durante 20 minutos.\n\n' +
           'Si se detiene, ejecute "Continuar Auditoría" desde el menú.');
  
  processQueue(); // Ejecutar el primer lote inmediatamente
}

/**
 * Procesa la cola de carpetas (almacenada en la hoja). Se ejecuta MANUALMENTE.
 */
function processQueue() {
  const startTime = new Date().getTime();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Cargar la HOJA DE COLA
  const queueSheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  if (!queueSheet) {
    ui.alert('No se encontró la hoja de cola "Queue_STATE". Por favor, inicie una nueva auditoría.');
    return;
  }
  
  // 2. Cargar la HOJA DE REPORTE
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    ui.alert(`No se encontró la hoja de reporte "${SHEET_NAME}". Deteniendo.`);
    return;
  }
  
  const domain = Session.getEffectiveUser().getEmail().split('@')[1];

  // 2. Procesar elementos de la cola mientras haya tiempo
  // El bucle ahora comprueba si hay filas en la hoja de cola
  while (queueSheet.getLastRow() > 1) { // > 1 porque la fila 1 es de encabezados
    const currentTime = new Date().getTime();
    if (currentTime - startTime > MAX_RUNTIME_MS) {
      // Tiempo agotado
      // No necesitamos guardar nada, el estado ES la hoja de cálculo
      SpreadsheetApp.flush();
      ui.alert('Límite de tiempo alcanzado (20 min).\n\n' +
               'Se ha guardado el progreso. Por favor, vuelva a ejecutar "Continuar Auditoría" desde el menú para seguir.');
      return; // Salir de la función
    }

    // Procesar el siguiente item (Fila 2)
    const range = queueSheet.getRange(2, 1, 1, queueSheet.getLastColumn());
    const values = range.getValues()[0];
    
    let parentPerms;
    try {
      parentPerms = JSON.parse(values[3]); // Reconvertir de JSON
    } catch (e) {
      Logger.log(`Error al parsear JSON de permisos para ${values[1]}: ${e.message}`);
      queueSheet.deleteRow(2); // Borrar la fila mala
      continue;
    }

    const currentItem = {
      id: values[0],
      path: values[1],
      url: values[2],
      parentPerms: parentPerms,
      isRoot: values[4]
    };
    
    let currentFolder;
    try {
      currentFolder = DriveApp.getFolderById(currentItem.id);
      Logger.log(`Procesando: ${currentItem.path}`);
    } catch (e) {
      Logger.log(`No se pudo acceder a la carpeta ${currentItem.path} (ID: ${currentItem.id}). Omitiendo. Error: ${e}`);
      sheet.appendRow([currentItem.path, currentItem.url, 'Carpeta', `ERROR: No se pudo acceder. ${e.message}`]);
      queueSheet.deleteRow(2); // Borrar la fila mala para continuar
      continue; // Siguiente item en la cola
    }

    const currentPermissions = getPermissionSet(currentFolder);

    // 3. Registrar diferencias de la CARPETA actual (si no es la raíz)
    if (!currentItem.isRoot) {
       logDifferences(currentItem.parentPerms, currentPermissions, currentItem.path, currentItem.url, 'Carpeta', sheet, domain);
    }

    // 4. Procesar ARCHIVOS dentro de esta carpeta
    try {
      const files = currentFolder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        const filePath = `${currentItem.path}/${file.getName()}`;
        const fileUrl = file.getUrl();
        try {
          const filePermissions = getPermissionSet(file);
          logDifferences(currentPermissions, filePermissions, filePath, fileUrl, 'Archivo', sheet, domain);
        } catch (e) {
          Logger.log(`Error al procesar archivo ${filePath}: ${e.message}`);
          sheet.appendRow([filePath, fileUrl, 'Archivo', `ERROR al leer permisos: ${e.message}`]);
        }
      }
    } catch (e) {
       Logger.log(`Error al obtener archivos en ${currentItem.path}: ${e.message}`);
    }

    // 5. Agregar SUBCARPETAS a la cola (al final de la hoja)
    const newRows = [];
    try {
      const subFolders = currentFolder.getFolders();
      while (subFolders.hasNext()) {
        const subFolder = subFolders.next();
        newRows.push([
          subFolder.getId(),
          `${currentItem.path}/${subFolder.getName()}`,
          subFolder.getUrl(),
          JSON.stringify(currentPermissions), // Los permisos de la carpeta actual son los del "padre"
          false // Ya no es la raíz
        ]);
      }
      if (newRows.length > 0) {
        queueSheet.getRange(queueSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
                  .setValues(newRows);
      }
    } catch (e) {
       Logger.log(`Error al obtener carpetas en ${currentItem.path}: ${e.message}`);
    }
    
    // 6. Borrar la fila que acabamos de procesar (Fila 2)
    queueSheet.deleteRow(2);
    
  } // fin del while

  // 7. Si el loop termina = Terminamos
  if (queueSheet.getLastRow() <= 1) { // Solo quedan encabezados
    Logger.log('Procesamiento completado.');
    clearState(true); // Limpiar el estado silenciosamente (borra la hoja de cola)
    SpreadsheetApp.flush();
    ui.alert('¡Auditoría completada exitosamente!'); // Avisar al usuario
  }
}

/**
 * Compara dos conjuntos de permisos y registra las diferencias en la hoja.
 * @param {Object} parentSet - Conjunto de permisos del padre.
... (y el resto de parámetros)
 */
function logDifferences(parentSet, childSet, path, url, type, sheet, domain) {
  // ... (Esta función no cambia) ...
  // Combinar todos los usuarios de ambos conjuntos para una comparación completa
  const allUsers = new Set([
    ...parentSet.editors,
    ...parentSet.viewers,
    ...childSet.editors,
    ...childSet.viewers
  ]);

  let differencesFound = false;

  // 1. Comprobar diferencias de usuarios
  allUsers.forEach(user => {
    const parentAccess = getAccess(user, parentSet);
    const childAccess = getAccess(user, childSet);
    if (parentAccess !== childAccess) {
      differencesFound = true;
    }
  });

  // 2. Comprobar diferencias de permisos de enlace (Público/Dominio)
  if (parentSet.publicAccess !== childSet.publicAccess || parentSet.domainAccess !== childSet.domainAccess) {
    differencesFound = true;
  }

  // Si encontramos CUALQUIER diferencia, registramos TODOS los permisos de este item.
  if (differencesFound) {
    Logger.log(`-> Se encontraron diferencias de permisos en: ${path}`);
    
    const usersWithRole = [];
    
    // Añadir permisos públicos/dominio
    if (childSet.publicAccess !== 'Ninguno') {
      usersWithRole.push(`Público (Cualquiera con enlace) - ${childSet.publicAccess}`);
    }
    if (childSet.domainAccess !== 'Ninguno') {
      usersWithRole.push(`Dominio (${domain}) - ${childSet.domainAccess}`);
    }
    
    // Añadir usuarios individuales
    childSet.editors.forEach(email => {
        usersWithRole.push(`${email} - Editor`);
    });
    
    childSet.viewers.forEach(email => {
        usersWithRole.push(`${email} - Visualizador`);
    });

    const userList = usersWithRole.length > 0 ? usersWithRole.join(', ') : 'Permisos solo por herencia (pero con diferencias)';

    sheet.appendRow([path, url, type, userList]);
  }
}
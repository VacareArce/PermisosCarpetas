/**
 * Archivo: API_SalidaDrive.js
 * Propósito: Gestiona la creación de la carpeta de reportes en Drive y la paginación 
 * autónoma de libros de Google Sheets cuando exceden el límite de filas configurado.
 */

/**
 * Crea una carpeta padre en la raíz del usuario que ejecuta el script para guardar todos los reportes particionados.
 * @param {string} nombreUnidad Auditada.
 * @return {string} El ID de la nueva carpeta creada en Drive.
 */
function instanciarCarpetaMaestra(nombreUnidad) {
    const fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const nombreCarpeta = `${PREFIJO_CARPETA_AUDITORIA} ${nombreUnidad} - ${fechaStr}`;

    // Crea la carpeta en la raíz (Mi unidad)
    const nuevaCarpeta = DriveApp.createFolder(nombreCarpeta);
    return nuevaCarpeta.getId();
}

/**
 * Gestiona el volcado de datos hacia un Sheet alojado en la carpeta de auditoría.
 * Si el sheet actual excede el límite (o no existe), crea uno nuevo (Parte N) y devuelve el nuevo estado.
 * 
 * @param {Object} estadoActual - { idSheet: 'string', filaActual: int, ramaNombre: 'string', parteActual: int, idCarpetaRaiz: 'string' }
 * @param {Array} datosFila - El array con los datos a insertar `['ruta', 'url', 'tipo', 'roles']`.
 * @return {Object} Retorna el objeto `estadoActual` el cual pudo haber sido mutado si se creó una nueva partición.
 */
function volcarHallazgoAPaginacion(estadoActual, datosFila) {
    // 1. Verificar si necesitamos crear un Sheet (porque es el primero, o rebasó el límite)
    if (!estadoActual.idSheet || estadoActual.filaActual >= LIMITE_FILAS_POR_HOJA_REPORTE) {

        estadoActual.parteActual += 1; // Subir de Parte 1 a Parte 2, etc.
        estadoActual.filaActual = 1; // Reiniciar contador de filas a inyectar (fila 1 = cabeceras)

        const nombreNuevoArchivo = `Reporte - ${estadoActual.ramaNombre} (Parte ${estadoActual.parteActual})`;

        // Crear el nuevo Google Sheet en la raíz virtual
        const nuevoSpreadsheet = SpreadsheetApp.create(nombreNuevoArchivo);
        const m_sheet = nuevoSpreadsheet.getSheets()[0];

        // Grabar Cabeceras Estéticas
        m_sheet.appendRow(CABECERAS_REPORTE_TECNICO);
        m_sheet.getRange(1, 1, 1, CABECERAS_REPORTE_TECNICO.length).setFontWeight('bold');
        m_sheet.setFrozenRows(1);

        estadoActual.idSheet = nuevoSpreadsheet.getId();
        estadoActual.filaActual = 2; // Apunta a la primera fila útil a rellenar

        // Mover físicamente el archivo recién nacido desde la raíz a la Carpeta Maestra
        const archivoVirtualEnDrive = DriveApp.getFileById(estadoActual.idSheet);
        const carpetaContenedora = DriveApp.getFolderById(estadoActual.idCarpetaRaiz);
        archivoVirtualEnDrive.moveTo(carpetaContenedora);
    }

    // 2. Anexar el dato de la anomalía al Sheet Activo
    try {
        const spreadDestino = SpreadsheetApp.openById(estadoActual.idSheet);
        const hojaDestino = spreadDestino.getSheets()[0];
        hojaDestino.appendRow(datosFila);
        estadoActual.filaActual += 1; // Aumentar en 1 el peso de esta hoja
    } catch (error) {
        Logger.log(`[DriverError] Fallo al insertar fila en la página ${estadoActual.idSheet} de rama ${estadoActual.ramaNombre}: ${error.message}`);
    }

    // Devolver el estado (idSheet e iteración de filas) al ciclo principal
    return estadoActual;
}

/**
 * Archivo: UI_Menu.js
 * Propósito: Gestiona el menú superior en Google Sheets y orquesta los puntos
 * de entrada de auditoría manejando estado.
 */

/**
 * @OnlyCurrentDoc
 */

/**
 * Función reservada en Google Apps Script que se ejecuta sola al abrir el documento.
 */
function onOpen() {
    crearMenuAuditoria();
}

/**
 * Instala y despliega un menú personalizado en la hoja actual visible.
 */
function crearMenuAuditoria() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Auditoría de Drive')
        .addItem('Iniciar Auditoría de Permisos', 'iniciarAuditoria')
        .addItem('Continuar Auditoría', 'continuarAuditoria')
        .addSeparator()
        .addItem('Limpiar Estado', 'limpiarEstadoAuditoria')
        .addToUi();
}

/**
 * Borra el estado guardado (la hoja de cola) para poder reiniciar la auditoría desde cero sin errores.
 * @param {boolean} silencioso - Parámetro para forzar que el usuario no reciba notificaciones (alerts) visuales al terminar.
 */
function limpiarEstadoAuditoria(silencioso) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Borrar la hoja oculta que hace de cola local
    const hojaDeCola = ss.getSheetByName(NOMBRE_HOJA_COLA);
    if (hojaDeCola) {
        ss.deleteSheet(hojaDeCola);
    }

    // Limpiar rastros guardados de configuraciones de scripts antiguos (Legacy)
    PropertiesService.getScriptProperties().deleteProperty(CLAVE_ESTADO_LEGADO);

    Logger.log('Estado limpiado desde interfaz.');
    if (!silencioso) {
        ui.alert('Estado de auditoría limpiado con éxito (Hoja de Cola fue eliminada). Ya puede iniciar una fase nueva.');
    }
}

/**
 * Analiza la Unidad Compartida pedida y empuja la primera capa al sistema de cola (Queue).
 */
function iniciarAuditoria() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Resetear el sistema de encolado actual
    limpiarEstadoAuditoria(true);

    // 2. Pedir credenciales o IDs de la unidad objetivo
    const respuestaUsuario = ui.prompt('ID de la Unidad Compartida', 'Ingrese por favor el ID oficial de la Unidad Compartida a examinar:', ui.ButtonSet.OK_CANCEL);
    if (respuestaUsuario.getSelectedButton() !== ui.Button.OK || !respuestaUsuario.getResponseText()) {
        ui.alert('Auditoría cancelada a petición.');
        return;
    }
    const idUnidadDrive = respuestaUsuario.getResponseText().trim();

    // 3. Pre-crear o reiniciar la hoja base para el usuario visual
    let hojaReporte = ss.getSheetByName(NOMBRE_HOJA_REPORTE);
    if (hojaReporte) {
        hojaReporte.clear();
    } else {
        hojaReporte = ss.insertSheet(NOMBRE_HOJA_REPORTE);
    }
    const cabecerasReporte = ['Ruta Escaneada', 'Enlace', 'Tipo (Doc/Folder)', 'Usuarios Encontrados (Roles)'];
    hojaReporte.appendRow(cabecerasReporte);
    hojaReporte.setFrozenRows(1);
    hojaReporte.getRange(1, 1, 1, cabecerasReporte.length).setFontWeight('bold');

    // 3b. Ocultar e inyectar datos raíz para la etapa paralela del proceso en segundo plano (Hoja Cola)
    let hojaDeCola = ss.getSheetByName(NOMBRE_HOJA_COLA);
    if (hojaDeCola) {
        ss.deleteSheet(hojaDeCola); // Evitar colisiones pasadas
    }
    hojaDeCola = ss.insertSheet(NOMBRE_HOJA_COLA);
    hojaDeCola.hideSheet(); // Solo para sistema, que no estorbe visualmente
    const cabecerasDeCola = ['ID Componente', 'Ruta Virtual', 'URL Enlace', 'Caché de Permisos (JSON)', 'Raiz (bool)'];
    hojaDeCola.appendRow(cabecerasDeCola);


    let nombreDrive = "Unidad Compartida Genérica";
    let carpetaRaizDrive;
    let urlRaizDrive;

    // 4. Intentar consumir API de Google
    try {
        const infoAvanzadaDrive = Drive.Drives.get(idUnidadDrive, { useDomainAdminAccess: true });
        nombreDrive = infoAvanzadaDrive.name;

        carpetaRaizDrive = DriveApp.getFolderById(idUnidadDrive);
        urlRaizDrive = carpetaRaizDrive.getUrl();

        // Detalles crudos nativos sin cache
        const permisosNativosRaiz = Drive.Permissions.list(idUnidadDrive, {
            supportsAllDrives: true,
            useDomainAdminAccess: true,
            fields: 'permissions(id, emailAddress, role, type, domain, permissionDetails)'
        });

        hojaReporte.appendRow([nombreDrive, urlRaizDrive, 'Unidad Compartida (Raíz/Nivel 0)', '--- Base Analizada de Todos los Permisos ---']);

        if (permisosNativosRaiz.permissions && permisosNativosRaiz.permissions.length > 0) {
            // Tabulacion por cargo y tipos de entidades en la raiz inicial
            const directorioRoles = {};
            permisosNativosRaiz.permissions.forEach(permiso => {
                if (permiso.permissionDetails && permiso.permissionDetails[0] && permiso.permissionDetails[0].inherited) {
                    return; // Bloquear los que viajan en cascada externamente a este Drive (ej dominical)
                }

                const denominacionRol = permiso.role.charAt(0).toUpperCase() + permiso.role.slice(1);
                let textoRepresentativo;

                if (permiso.type === 'user') {
                    textoRepresentativo = permiso.emailAddress;
                } else if (permiso.type === 'group') {
                    textoRepresentativo = permiso.emailAddress || `Mailing List/Grupo (ID: ${permiso.id})`;
                } else if (permiso.type === 'anyone') {
                    textoRepresentativo = 'Extranjero Público (Persona anónima)';
                } else if (permiso.type === 'domain') {
                    textoRepresentativo = `Colaboradores del Dominio (${permiso.domain || 'Dominio interno'})`;
                } else {
                    textoRepresentativo = `${permiso.type} (ID Sistema: ${permiso.id})`;
                }

                if (!directorioRoles[denominacionRol]) {
                    directorioRoles[denominacionRol] = [];
                }
                directorioRoles[denominacionRol].push(textoRepresentativo);
            });

            // Plasmar cabecera analfabética 
            for (const denominacion in directorioRoles) {
                const consolidados = directorioRoles[denominacion].map(correoAsignado => `${correoAsignado} - [${denominacion}]`).join(', ');
                hojaReporte.appendRow([nombreDrive, urlRaizDrive, `Permisos Glob. ${denominacion}`, consolidados]);
            }

        } else {
            hojaReporte.appendRow([nombreDrive, urlRaizDrive, 'Unidad (Raíz Padre)', 'Drive Limpio - Carece de permisos asignados explícitamente a este nivel.']);
        }

        hojaReporte.appendRow(['---', '---', '---', '---']);
        hojaReporte.appendRow(['(Búsqueda de Profundidad) Componentes con esquemas de acceso distintos a su estructura Padre superior:']);
        hojaReporte.appendRow(['---', '---', '---', '---']);

    } catch (error) {
        Logger.log(`API Bloqueada/Falló de forma prematura en acceso central: ${error.message}`);
        const detalleFalla = `La conexión con dicha base/unidad Drive falló (ID de Drive proporcionado rechazado). Error de sistema: ${error.message}`;

        hojaReporte.appendRow([idUnidadDrive, 'ERROR DE CRITICIDAD', 'Problema de Acceso', detalleFalla]);

        // Alertas por contexto a usuario para que solucione desde consola externa si es necesario
        if (error.message.includes("Drive API has not been used")) {
            ui.alert(`${detalleFalla}\n\n-> DEBE: Habilitar "Drive API V3" vía panel GCP Console asociado a su proyecto.`);
        } else if (error.message.includes("Forbidden")) {
            ui.alert(`${detalleFalla}\n\n-> DEBE: Solicite escalamiento de Perfil de 'Administrador' o permisos adecuados para entrar a tal Unidad Compartida Drive.`);
        } else {
            ui.alert(detalleFalla);
        }
        return; // Kill proces
    }

    // 5. Arranque real del worker inyectando el Root original al motor (Fila 2 en Backend de Cola)
    const permisosGeneralesRaiz = obtenerConjuntoPermisos(carpetaRaizDrive); // Llama funcion importada de API_Permisos.js

    hojaDeCola.appendRow([
        carpetaRaizDrive.getId(),
        nombreDrive,
        urlRaizDrive,
        JSON.stringify(permisosGeneralesRaiz),
        true
    ]);

    ui.alert('Motor analítico de permisos inicializado de fondo.\nLa auditoría tiene de límite 20m.\n\nEn caso de interrupción prematura por volumen, por favor ejecute "Continuar Auditoría" del menú superior.');

    continuarAuditoria(); // Conecta ciclo #2
}

/**
 * Iterador cronometrado que lee de la base oculta Queue_STATE y aplica búsqueda
 * en anchura limitándose por tiempo definido a constantes para evadir bloqueos por exceso de cómputo.
 */
function continuarAuditoria() {
    const tiempoEmpezadoMS = new Date().getTime();
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Confirmar infraestructura oculta levantada
    const hojaDeCola = ss.getSheetByName(NOMBRE_HOJA_COLA);
    if (!hojaDeCola) {
        ui.alert('Ausencia de tabla Cola ("Queue_STATE" No detectada). Debes usar primero "Iniciar Auditoría de Permisos".');
        return;
    }

    const hojaReporte = ss.getSheetByName(NOMBRE_HOJA_REPORTE);
    if (!hojaReporte) {
        ui.alert(`Falta hoja gráfica Reportes de Salida ("${NOMBRE_HOJA_REPORTE}"). Proceso parado urgemente.`);
        return;
    }

    const dominioOrganizacion = Session.getEffectiveUser().getEmail().split('@')[1];

    // 2. Loop a nivel capa (1 elemento sacado, N hijos listados detrás)
    while (hojaDeCola.getLastRow() > 1) {
        const milisegundosMarcador = new Date().getTime();
        if (milisegundosMarcador - tiempoEmpezadoMS > TIEMPO_MAXIMO_EJECUCION_MS) {
            // Bloqueo de cortocircuito (Evasión de Límite Runtime V8 Apps Script - Timeout Previsto)
            SpreadsheetApp.flush();
            ui.alert('Pausa táctica por Límite de Tiempo por CPU GSuite (Pasaron 20 minutos completos).\n\nProgreso indexado sin pérdidas. Presione opción "Continuar Auditoría" desde el Menú para resumir extracción.');
            return;
        }

        // Extracción tipo Fila (Indice superior de cola, celda "A2" - "E2")
        const rangoActualSuperior = hojaDeCola.getRange(2, 1, 1, hojaDeCola.getLastColumn());
        const columnaMapeada = rangoActualSuperior.getValues()[0];

        let permisosAlmacenadosJSON;
        try {
            permisosAlmacenadosJSON = JSON.parse(columnaMapeada[3]);
        } catch (e) {
            Logger.log(`JSON Crash parsing en caché fila ruta: ${columnaMapeada[1]}: Detalles JSON: ${e.message}`);
            hojaDeCola.deleteRow(2); // Suprimir registro corrompido, evita DeadLock cíclico.
            continue;
        }

        const entidadDirectorioActual = {
            idNode: columnaMapeada[0],
            rutaArmada: columnaMapeada[1],
            urlVisita: columnaMapeada[2],
            permisosRecuperadosPadre: permisosAlmacenadosJSON,
            banderaRaiz: columnaMapeada[4]
        };

        let objetoCarpetaDrive;
        try {
            objetoCarpetaDrive = DriveApp.getFolderById(entidadDirectorioActual.idNode);
        } catch (errores_api) {
            Logger.log(`Punto Ciego de Carpeta ${entidadDirectorioActual.rutaArmada} (ID Cifrado: ${entidadDirectorioActual.idNode}). Fallo Inaccesible vía API: ${errores_api}`);
            hojaReporte.appendRow([entidadDirectorioActual.rutaArmada, entidadDirectorioActual.urlVisita, 'Folder Ciego', `ERROR API Carga: Restricción del propio Google sobre la ID oculta: ${errores_api.message}`]);
            hojaDeCola.deleteRow(2);
            continue;
        }

        const permisosEntidadLocal = obtenerConjuntoPermisos(objetoCarpetaDrive); // Llama a API_Permisos.js

        // 3. Revisión Delta: Identificar desviaciones de la norma base
        if (!entidadDirectorioActual.banderaRaiz) {
            // Enlaza la ejecución reportada a API_Permisos.js
            registrarDiferenciasPermisos(entidadDirectorioActual.permisosRecuperadosPadre, permisosEntidadLocal, entidadDirectorioActual.rutaArmada, entidadDirectorioActual.urlVisita, 'Carpeta Plegable', hojaReporte, dominioOrganizacion);
        }

        // 4. Extracción Lineal de los documentos simples hijos
        try {
            const listaDocumentosHijos = objetoCarpetaDrive.getFiles();
            while (listaDocumentosHijos.hasNext()) {
                const documentoSingular = listaDocumentosHijos.next();
                const senderoDocTexto = `${entidadDirectorioActual.rutaArmada}/${documentoSingular.getName()}`;
                const enlaceDocSalida = documentoSingular.getUrl();
                try {
                    const permisosDelDocumento = obtenerConjuntoPermisos(documentoSingular);
                    registrarDiferenciasPermisos(permisosEntidadLocal, permisosDelDocumento, senderoDocTexto, enlaceDocSalida, 'Documento Unitario', hojaReporte, dominioOrganizacion);
                } catch (errorlecturadoc) {
                    hojaReporte.appendRow([senderoDocTexto, enlaceDocSalida, 'Documento Unitario', `Alerta de Lectura en Permisología: ${errorlecturadoc.message}`]);
                }
            }
        } catch (errorCargaFolders) {
            Logger.log(`Problematica al tirar Folders Hijos de Root Folder ${entidadDirectorioActual.rutaArmada}: ${errorCargaFolders.message}`);
        }

        // 5. Recarga Reversa de Nueva Carpeta Hija sub encontrada directo en Cola final
        const ramificacionesNuevasACola = [];
        try {
            const foldersDebajo = objetoCarpetaDrive.getFolders();
            while (foldersDebajo.hasNext()) {
                const subSubFolder = foldersDebajo.next();
                ramificacionesNuevasACola.push([
                    subSubFolder.getId(),
                    `${entidadDirectorioActual.rutaArmada}/${subSubFolder.getName()}`,
                    subSubFolder.getUrl(),
                    JSON.stringify(permisosEntidadLocal), // Se re-heredan estos como padre temporal
                    false // Se levanta flag raíz natural
                ]);
            }

            // Encolado masivo a la AppScript Sheet
            if (ramificacionesNuevasACola.length > 0) {
                hojaDeCola.getRange(hojaDeCola.getLastRow() + 1, 1, ramificacionesNuevasACola.length, ramificacionesNuevasACola[0].length)
                    .setValues(ramificacionesNuevasACola);
            }
        } catch (ejecucionFail) { }

        // 6. Ciclo matanza - destruye de hoja Cola fila 2 despues de finalizado el arbol.
        hojaDeCola.deleteRow(2);

    }

    // 7. Vaciado limpio. Cola esta sin hijos adentro (solo headers cabecera).
    if (hojaDeCola.getLastRow() <= 1) {
        Logger.log('Procesamiento completado y analizado en su totalidad del árbol.');
        limpiarEstadoAuditoria(true);
        SpreadsheetApp.flush();
        ui.alert('¡Barrido y Auditoría Arquitectónica finalizada 100% de forma correcta con éxito total!');
    }
}

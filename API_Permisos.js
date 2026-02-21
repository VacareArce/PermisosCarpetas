/**
 * Archivo: API_Permisos.js
 * Propósito: Contiene las funciones que interactúan con DriveApp y Drive API para extraer,
 * comparar y registrar los niveles de acceso.
 */

/**
 * Obtiene el conjunto de permisos "simplificado" (Editores, Visualizadores, Público, Dominio)
 * para una rápida comparación de herencia.
 * @param {File|Folder} archivoDrive - El archivo o carpeta a auditar.
 * @return {Object} Un objeto con { editores: [], visualizadores: [], accesoPublico: 'Ninguno', accesoDominio: 'Ninguno' }.
 */
function obtenerConjuntoPermisos(archivoDrive) {
    const editores = new Set();
    const visualizadores = new Set();
    let accesoPublico = 'Ninguno';
    let accesoDominio = 'Ninguno';

    try {
        archivoDrive.getEditors().forEach(usuario => editores.add(usuario.getEmail()));
    } catch (e) { /* Ignorar errores al obtener editores */ }

    try {
        archivoDrive.getViewers().forEach(usuario => visualizadores.add(usuario.getEmail()));
    } catch (e) { /* Ignorar errores al obtener visualizadores */ }

    try {
        const accesoCompartido = archivoDrive.getSharingAccess();
        const permisoCompartido = archivoDrive.getSharingPermission();

        if (accesoCompartido === DriveApp.Access.ANYONE || accesoCompartido === DriveApp.Access.ANYONE_WITH_LINK) {
            if (permisoCompartido === DriveApp.Permission.EDIT) {
                accesoPublico = 'Editor';
            } else if (permisoCompartido === DriveApp.Permission.VIEW) {
                accesoPublico = 'Visualizador';
            }
        } else if (accesoCompartido === DriveApp.Access.DOMAIN || accesoCompartido === DriveApp.Access.DOMAIN_WITH_LINK) {
            if (permisoCompartido === DriveApp.Permission.EDIT) {
                accesoDominio = 'Editor';
            } else if (permisoCompartido === DriveApp.Permission.VIEW) {
                accesoDominio = 'Visualizador';
            }
        }
    } catch (e) { /* Ignorar errores al obtener permisos de enlaces públicos o de dominio */ }

    const visualizadoresFinales = [...visualizadores].filter(v => !editores.has(v));

    return {
        editores: [...editores],
        visualizadores: visualizadoresFinales,
        accesoPublico: accesoPublico,
        accesoDominio: accesoDominio
    };
}

/**
 * Obtiene el nivel de acceso simplificado de un usuario específico.
 * @param {string} emailUsuario - El correo electrónico del usuario.
 * @param {Object} conjuntoPermisos - Objeto con los permisos obtenido previamente de "obtenerConjuntoPermisos".
 * @return {string} 'Editor', 'Visualizador', o 'Ninguno'.
 */
function obtenerNivelAccesoUsuario(emailUsuario, conjuntoPermisos) {
    if (conjuntoPermisos.editores.includes(emailUsuario)) return 'Editor';
    if (conjuntoPermisos.visualizadores.includes(emailUsuario)) return 'Visualizador';
    return 'Ninguno';
}

/**
 * Compara dos conjuntos de permisos y formatea las diferencias detectadas para su encolamiento.
 * @param {Object} conjuntoPadre - Permisos de la carpeta/unidad base.
 * @param {Object} conjuntoHijo - Permisos del archivo/carpeta interno a evaluar.
 * @param {string} ruta - Ruta amigable del archivo o carpeta.
 * @param {string} url - Enlace URL al archivo o carpeta.
 * @param {string} tipoItem - Denominación en texto (ej. "Carpeta", "Archivo").
 * @param {string} dominioOrganizacion - El dominio principal de la cuenta con la que se ejecuta (texto).
 * @return {Array|null} Retorna el arreglo de la fila `[ruta, url, tipo, usuariosTexto]` si hubo hallazgos, o null si todo está en orden.
 */
function registrarDiferenciasPermisos(conjuntoPadre, conjuntoHijo, ruta, url, tipoItem, dominioOrganizacion) {
    // Combinar usuarios de ambos niveles para comprobar diferencias
    const todosLosUsuarios = new Set([
        ...conjuntoPadre.editores,
        ...conjuntoPadre.visualizadores,
        ...conjuntoHijo.editores,
        ...conjuntoHijo.visualizadores
    ]);

    let diferenciasDetectadas = false;

    // 1. Evaluar las diferencias a nivel de cuenta individual (usuarios directos)
    todosLosUsuarios.forEach(usuario => {
        const accesoPadre = obtenerNivelAccesoUsuario(usuario, conjuntoPadre);
        const accesoHijo = obtenerNivelAccesoUsuario(usuario, conjuntoHijo);
        if (accesoPadre !== accesoHijo) {
            diferenciasDetectadas = true;
        }
    });

    // 2. Evaluar reglas de publicacion general
    if (conjuntoPadre.accesoPublico !== conjuntoHijo.accesoPublico || conjuntoPadre.accesoDominio !== conjuntoHijo.accesoDominio) {
        diferenciasDetectadas = true;
    }

    // Escribir solo si se detectó alteración respecto a su carpeta padre
    if (diferenciasDetectadas) {
        Logger.log(`-> Se encontraron diferencias de permisos en: ${ruta}`);

        const listaUsuariosRoles = [];

        // Nivel público o de organización entera
        if (conjuntoHijo.accesoPublico !== 'Ninguno') {
            listaUsuariosRoles.push(`Público (Cualquiera con enlace) - ${conjuntoHijo.accesoPublico}`);
        }
        if (conjuntoHijo.accesoDominio !== 'Ninguno') {
            listaUsuariosRoles.push(`Dominio (${dominioOrganizacion}) - ${conjuntoHijo.accesoDominio}`);
        }

        // Nivel usuarios o grupos
        conjuntoHijo.editores.forEach(correo => {
            listaUsuariosRoles.push(`${correo} - Editor`);
        });

        conjuntoHijo.visualizadores.forEach(correo => {
            listaUsuariosRoles.push(`${correo} - Visualizador`);
        });

        const usuariosTexto = listaUsuariosRoles.length > 0 ? listaUsuariosRoles.join(', ') : 'Permisos solo por herencia (pero presenta diferencias en configuración de enlace)';

        hojaReporte.appendRow([ruta, url, tipoItem, usuariosTexto]);
    }
}

/**
 * Archivo: Cfg_Constantes.js
 * Propósito: Almacena las variables globales y configuraciones del proyecto.
 */

const NOMBRE_HOJA_REPORTE = 'Reporte de Permisos';
const NOMBRE_HOJA_COLA = 'Queue_STATE'; // Hoja usada para gestionar la cola de auditoría

// Límite manual para cuentas Workspace es de 30 minutos. Se usan 20 minutos (en milisegundos).
const TIEMPO_MAXIMO_EJECUCION_MS = 20 * 60 * 1000;

// Propiedad heredada que se mantiene por limpieza (puede ser opcional en nuevas instalaciones).
const CLAVE_ESTADO_LEGADO = 'DRIVE_AUDIT_STATE';

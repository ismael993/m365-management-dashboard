export interface OfficeAccount {
  id: number;
  email: string;
  nombre: string;
  fechaCreacion: string;
  rut: string;
  licencia: string;
  ot: string;
  activa: boolean;
  //Nuevos campos de seguridad
  mfaHabilitado?: boolean;
  diasInactivo?: number;
}

//Interfaz para los registros de Audotoria (Audit Logs)
export interface AuditLogEntry {
  id: string;
  fechaHora: string;
  usuarioAdmin: string;
  accion: 'BAJA' | 'REACTIVACION' | 'EXPORTACION' | 'INICIO_SESION';
  detalles: string;
  usuarioAfectado?: string;
  ipCliente?: string;
}

// Interfaz para el Inventario e Identificación de costos de Licencias
export interface LicenseInventory {
  skuId: string;
  nombreComercial: string;
  totalCompradas: number;
  asignadas: number;
  disponibles: number;
  porcentajeUso: number;
  estadoConsumo: 'Optimo' | 'Advertencia' | 'Critico';
}

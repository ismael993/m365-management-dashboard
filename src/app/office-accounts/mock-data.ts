import { OfficeAccount, LicenseInventory } from './account.model';

export class MockDataService {
  static getMockAccounts(): OfficeAccount[] {
    return [
      {
        id: 1,
        nombre: 'Carlos Mendoza',
        email: 'carlos.mendoza@cloudcorp365.com',
        fechaCreacion: '15/01/2023',
        rut: '12.345.678-9',
        licencia: 'Microsoft 365 Empresa Estándar',
        ot: 'OT-1001',
        activa: true,
        mfaHabilitado: true,
        diasInactivo: 5
      },
      {
        id: 2,
        nombre: 'Valentina Ríos',
        email: 'valentina.rios@cloudcorp365.com',
        fechaCreacion: '02/03/2023',
        rut: '15.678.912-3',
        licencia: 'Power BI Pro, Microsoft 365 Empresa Estándar',
        ot: 'OT-1002',
        activa: true,
        mfaHabilitado: true,
        diasInactivo: 12
      },
      {
        id: 3,
        nombre: 'Matías Silva',
        email: 'matias.silva@cloudcorp365.com',
        fechaCreacion: '11/05/2023',
        rut: '18.234.567-K',
        licencia: 'Microsoft 365 Empresa Básico',
        ot: 'OT-1003',
        activa: true,
        mfaHabilitado: false,
        diasInactivo: 98
      },
      {
        id: 4,
        nombre: 'Camila Morales',
        email: 'camila.morales@cloudcorp365.com',
        fechaCreacion: '20/08/2023',
        rut: '16.543.210-4',
        licencia: 'Unlicensed',
        ot: 'OT-1004',
        activa: false,
        mfaHabilitado: false,
        diasInactivo: 120
      },
      {
        id: 5,
        nombre: 'Diego Fernández',
        email: 'diego.fernandez@cloudcorp365.com',
        fechaCreacion: '10/01/2024',
        rut: '17.890.123-5',
        licencia: 'Microsoft 365 Empresa Estándar',
        ot: 'OT-1005',
        activa: true,
        mfaHabilitado: true,
        diasInactivo: 2
      }
    ];
  }

  static getMockInventario(): LicenseInventory[] {
    return [
      {
        skuId: 'mock-sku-1',
        nombreComercial: 'Microsoft 365 Empresa Estándar',
        totalCompradas: 50,
        asignadas: 42,
        disponibles: 8,
        porcentajeUso: 84,
        estadoConsumo: 'Advertencia'
      },
      {
        skuId: 'mock-sku-2',
        nombreComercial: 'Microsoft 365 Empresa Básico',
        totalCompradas: 30,
        asignadas: 25,
        disponibles: 5,
        porcentajeUso: 83,
        estadoConsumo: 'Advertencia'
      },
      {
        skuId: 'mock-sku-3',
        nombreComercial: 'Power BI Pro',
        totalCompradas: 15,
        asignadas: 10,
        disponibles: 5,
        porcentajeUso: 67,
        estadoConsumo: 'Optimo'
      }
    ];
  }
}

import { Component, inject, signal, computed, OnInit, OnDestroy, PLATFORM_ID } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { FormsModule } from '@angular/forms';
import { MsalService } from '@azure/msal-angular';
import { MockDataService } from './mock-data';
import { OfficeAccount, AuditLogEntry, LicenseInventory } from './account.model';
import { forkJoin, fromEvent, Subscription } from 'rxjs';
import * as XLSX from 'xlsx';

export interface M365ServiceStatus {
  nombre: string;
  icono: string;
  estado: 'operativo' | 'degradado' | 'mantenimiento';
  mensaje: string;
}

@Component({
  selector: 'app-office-accounts',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './office-accounts.html',
  styleUrl: './office-accounts.css',
})
export class OfficeAccounts implements OnInit, OnDestroy {
  currentYear = new Date().getFullYear();
  private http = inject(HttpClient);
  private msalService = inject(MsalService);
  private platformId = inject(PLATFORM_ID);

  // 🧪 INTERRUPTOR MODO DEMO / PRODUCCIÓN
  // 🟢 true  => Carga datos demostrativos (GitHub / Portfolio)
  // 🔴 false => Llama a la API real de Microsoft Graph (Producción OCA)
  esModoDemo = signal<boolean>(true);

  // 🏢 Timestamp de Sincronización
  ultimaSincronizacion = signal<string>('');

  // 🔀 Control de Pestañas (Tabs Control)
  vistaActual = signal<'CUENTAS' | 'LICENCIAS'>('CUENTAS');

  // ⏳ Estados de UI, carga y conectividad
  cargando = signal<boolean>(true);
  conexionExitosa = signal<boolean>(false);
  sinConexionRed = signal<boolean>(false);

  // 🟢 Panel de Salud de Servicios M365
  mostrarPanelSalud = signal<boolean>(false);
  serviciosSalud = signal<M365ServiceStatus[]>([
    { nombre: 'Exchange Online', icono: '✉️', estado: 'operativo', mensaje: 'Sin incidencias reportadas' },
    { nombre: 'Microsoft Teams', icono: '💬', estado: 'operativo', mensaje: 'Funcionamiento normal' },
    { nombre: 'SharePoint / OneDrive', icono: '📁', estado: 'operativo', mensaje: 'Servicio estable' },
    { nombre: 'Microsoft Entra ID', icono: '🔑', estado: 'operativo', mensaje: 'Autenticación activa' }
  ]);

  // 📜 REGISTRO DE AUDITORÍA (AUDIT LOGS)
  auditLogs = signal<AuditLogEntry[]>([]);
  mostrarModalAuditoria = signal<boolean>(false);

  // 💳 INVENTARIO DE LICENCIAS M365
  licenciasInventario = signal<LicenseInventory[]>([]);

  // ⚠️ Modal corporativo de confirmación de estado
  mostrarModalConfirmacion = signal<boolean>(false);
  cuentaSeleccionada = signal<OfficeAccount | null>(null);
  motivoCambio = signal<string>('');

  private onlineSub?: Subscription;
  private offlineSub?: Subscription;

  accounts = signal<OfficeAccount[]>([]);
  busqueda = signal<string>('');
  filtroLicencia = signal<string>('TODAS');
  filtroEstado = signal<'TODAS' | 'ACTIVAS' | 'BAJA'>('ACTIVAS');
  estaAutenticado = signal<boolean>(false);

  nombreUsuarioConectado = computed(() => {
    if (this.esModoDemo()) {
      return 'Demo Admin (CloudCorp 365)';
    }
    const account = this.msalService.instance.getActiveAccount();
    return account ? account.name || account.username : 'Administrador OCA';
  });

  tiposDeLicenciasDisponibles = computed(() => {
    const licenciasSet = new Set<string>();
    this.accounts().forEach(acc => {
      if (acc.licencia && acc.licencia !== 'Unlicensed') {
        const licenciasUsuario = acc.licencia.split(', ');
        licenciasUsuario.forEach(lic => licenciasSet.add(lic.trim()));
      }
    });
    return Array.from(licenciasSet).sort();
  });

  cuentasFiltradas = computed(() => {
    const query = this.busqueda().toLowerCase().trim();
    const licenciaSel = this.filtroLicencia();
    const estadoSel = this.filtroEstado();

    return this.accounts().filter(acc => {
      let coincideEstado = true;
      if (estadoSel === 'ACTIVAS') {
        coincideEstado = acc.activa && acc.licencia !== 'Unlicensed';
      } else if (estadoSel === 'BAJA') {
        coincideEstado = !acc.activa || acc.licencia === 'Unlicensed';
      }

      const coincideTexto = !query ||
        acc.nombre.toLowerCase().includes(query) ||
        acc.email.toLowerCase().includes(query) ||
        acc.rut.toLowerCase().includes(query) ||
        acc.ot.toLowerCase().includes(query);

      const coincideLicencia = licenciaSel === 'TODAS' || acc.licencia.includes(licenciaSel);

      return coincideEstado && coincideTexto && coincideLicencia;
    });
  });

  totalCuentas = computed(() => this.accounts().length);
  cuentasDadasDeBaja = computed(() => this.accounts().filter(acc => !acc.activa || acc.licencia === 'Unlicensed').length);
  cuentasActivasQuedan = computed(() => this.accounts().filter(acc => acc.activa && acc.licencia !== 'Unlicensed').length);

  // 🔒 MÉTRICAS DE SEGURIDAD COMPUTADAS
  cuentasSinMFA = computed(() =>
    this.accounts().filter(acc => acc.activa && !acc.mfaHabilitado).length
  );

  cuentasInactivasCriticas = computed(() =>
    this.accounts().filter(acc => acc.activa && (acc.diasInactivo ?? 0) >= 90).length
  );

  porcentajeCoberturaMFA = computed(() => {
    const activas = this.cuentasActivasQuedan();
    if (activas === 0) return 0;
    const conMfa = this.accounts().filter(acc => acc.activa && acc.mfaHabilitado).length;
    return Math.round((conMfa / activas) * 100);
  });

  // 💰 MÉTRICAS FINANCIERAS COMPUTADAS
  totalLicenciasCompradas = computed(() =>
    this.licenciasInventario().reduce((acc, item) => acc + item.totalCompradas, 0)
  );

  totalLicenciasSinUsar = computed(() =>
    this.licenciasInventario().reduce((acc, item) => acc + item.disponibles, 0)
  );

  async ngOnInit(): Promise<void> {
    if (isPlatformBrowser(this.platformId)) {
      this.escucharEstadoRed();

      if (this.esModoDemo()) {
        this.estaAutenticado.set(true);
        this.obtenerTokenYCargar();
        return;
      }

      try {
        await this.msalService.instance.initialize();
        const result = await this.msalService.instance.handleRedirectPromise();

        if (result && result.account) {
          this.msalService.instance.setActiveAccount(result.account);
        }

        const activeAccount = this.msalService.instance.getActiveAccount() || this.msalService.instance.getAllAccounts()[0];

        if (activeAccount) {
          this.msalService.instance.setActiveAccount(activeAccount);
          this.estaAutenticado.set(true);
          this.obtenerTokenYCargar();
        } else {
          this.estaAutenticado.set(false);
          this.cargando.set(false);
        }
      } catch (error) {
        console.error('Error al inicializar MSAL:', error);
        this.cargando.set(false);
      }
    }
  }

  ngOnDestroy(): void {
    this.onlineSub?.unsubscribe();
    this.offlineSub?.unsubscribe();
  }

  togglePanelSalud(): void {
    this.mostrarPanelSalud.update(val => !val);
  }

  toggleModalAuditoria(): void {
    this.mostrarModalAuditoria.update(val => !val);
  }

  private registrarAccionAuditoria(
    accion: 'BAJA' | 'REACTIVACION' | 'EXPORTACION' | 'INICIO_SESION',
    detalles: string,
    usuarioAfectado?: string
  ): void {
    const nuevaEntrada: AuditLogEntry = {
      id: `LOG-${Date.now().toString().slice(-6)}`,
      fechaHora: new Date().toLocaleString('es-CL'),
      usuarioAdmin: this.nombreUsuarioConectado(),
      accion,
      detalles,
      usuarioAfectado: usuarioAfectado || 'N/A'
    };

    this.auditLogs.update(logs => [nuevaEntrada, ...logs]);
  }

  private escucharEstadoRed(): void {
    this.sinConexionRed.set(!navigator.onLine);

    this.onlineSub = fromEvent(window, 'online').subscribe(() => {
      this.sinConexionRed.set(false);
      this.obtenerTokenYCargar();
    });

    this.offlineSub = fromEvent(window, 'offline').subscribe(() => {
      this.sinConexionRed.set(true);
      this.conexionExitosa.set(false);
    });
  }

  iniciarAutenticacionYCargar(): void {
    if (this.esModoDemo()) {
      this.estaAutenticado.set(true);
      this.obtenerTokenYCargar();
      return;
    }

    const account = this.msalService.instance.getActiveAccount();
    if (!account) {
      this.msalService.instance.loginRedirect({
        scopes: ['User.Read.All', 'Organization.Read.All']
      });
    } else {
      this.estaAutenticado.set(true);
      this.obtenerTokenYCargar();
    }
  }

  cerrarSesion(): void {
    this.accounts.set([]);
    this.estaAutenticado.set(false);
    sessionStorage.clear();
    localStorage.clear();

    if (!this.esModoDemo()) {
      this.msalService.instance.logoutRedirect({
        postLogoutRedirectUri: 'http://localhost:4200/'
      });
    }
  }

  private async obtenerTokenYCargar(): Promise<void> {
    if (this.esModoDemo()) {
      this.cargarDatosDemo();
      return;
    }

    try {
      const account = this.msalService.instance.getActiveAccount();
      if (!account) return;

      const tokenResponse = await this.msalService.instance.acquireTokenSilent({
        scopes: ['User.Read.All', 'Organization.Read.All'],
        account: account
      }).catch(() => {
        this.msalService.instance.acquireTokenRedirect({
          scopes: ['User.Read.All', 'Organization.Read.All']
        });
        return null;
      });

      if (tokenResponse?.accessToken) {
        this.cargarUsuariosConToken(tokenResponse.accessToken);
      } else {
        this.cargando.set(false);
      }
    } catch (error) {
      console.error('Error al obtener token:', error);
      this.cargando.set(false);
    }
  }

  // 🧪 Carga simulada para el Modo Demo
  cargarDatosDemo(): void {
    this.cargando.set(true);
    setTimeout(() => {
      this.accounts.set(MockDataService.getMockAccounts());
      this.licenciasInventario.set(MockDataService.getMockInventario());
      this.cargando.set(false);
      this.ultimaSincronizacion.set(new Date().toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit' }));
      this.conexionExitosa.set(true);
      this.registrarAccionAuditoria('INICIO_SESION', 'Sincronización simulada activa (Modo Demo CloudCorp 365)');
    }, 600);
  }

  cargarUsuariosConToken(token: string): void {
    this.cargando.set(true);
    this.conexionExitosa.set(false);

    const headers = new HttpHeaders({
      'Authorization': `Bearer ${token}`
    });

    const endpointUsers = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,createdDateTime,assignedLicenses,faxNumber,postalCode,userType&$top=999';
    const endpointSkus = 'https://graph.microsoft.com/v1.0/subscribedSkus';

    forkJoin({
      usersRes: this.http.get<any>(endpointUsers, { headers }),
      skusRes: this.http.get<any>(endpointSkus, { headers })
    }).subscribe({
      next: ({ usersRes, skusRes }) => {
        const DICCIONARIO_LICENCIAS: { [key: string]: string } = {
          'O365_BUSINESS_PREMIUM': 'Microsoft 365 Empresa Estándar',
          'O365_BUSINESS': 'Aplicaciones de Microsoft 365 para negocios',
          'O365_BUSINESS_ESSENTIALS': 'Microsoft 365 Empresa Básico',
          'M365_BUSINESS_STANDARD': 'Microsoft 365 Empresa Estándar',
          'AAD_PREMIUM': 'Azure Active Directory Premium P1',
          'AAD_PREMIUM_P2': 'Azure Active Directory Premium P2',
          'INTUNE_A': 'Intune',
          'FABRIC_FREE': 'Microsoft Fabric (Gratis)',
          'POWERAPPS_DEV': 'Microsoft Power Apps for Developer',
          'POWERAPPS_VIRAL': 'Microsoft Power Apps Plan 2 Trial',
          'FLOW_FREE': 'Microsoft Power Automate Free',
          'PROJECTPREMIUM': 'Planner y Project Plan 3',
          'POWER_BI_STANDARD': 'Power BI Pro',
          'POWER_BI_FREE': 'Power BI (Gratis)',
          'VISIOCLIENT': 'Visio Plan 2'
        };

        const mapaLicencias: { [skuId: string]: string } = {};
        const inventarioProcesado: LicenseInventory[] = [];

        if (skusRes && skusRes.value) {
          skusRes.value.forEach((sku: any) => {
            const rawSku = sku.skuPartNumber || '';
            const nombreComercial = DICCIONARIO_LICENCIAS[rawSku] || rawSku.replace(/_/g, ' ');
            mapaLicencias[sku.skuId] = nombreComercial;

            const total = sku.capabilityStatus === 'Enabled' ? (sku.prepaidUnits?.enabled || 0) : 0;
            const asignadas = sku.consumedUnits || 0;
            const disponibles = Math.max(0, total - asignadas);
            const pct = total > 0 ? Math.round((asignadas / total) * 100) : 0;

            let estado: 'Optimo' | 'Advertencia' | 'Critico' = 'Optimo';
            if (pct >= 95) estado = 'Critico';
            else if (pct >= 80) estado = 'Advertencia';

            inventarioProcesado.push({
              skuId: sku.skuId,
              nombreComercial,
              totalCompradas: total,
              asignadas,
              disponibles,
              porcentajeUso: pct,
              estadoConsumo: estado
            });
          });
        }

        this.licenciasInventario.set(inventarioProcesado.sort((a, b) => b.totalCompradas - a.totalCompradas));

        if (usersRes && usersRes.value) {
          // Filtrado de usuarios externos (#EXT# o Guest)
          const usuariosInternos = usersRes.value.filter((u: any) => {
            const esInvitadoExt = u.userPrincipalName && u.userPrincipalName.includes('#EXT#');
            const esGuestType = u.userType && u.userType.toLowerCase() === 'guest';
            return !esInvitadoExt && !esGuestType;
          });

          const usuariosMapeados: OfficeAccount[] = usuariosInternos.map((u: any, index: number) => {
            const tieneLicencias = u.assignedLicenses && u.assignedLicenses.length > 0;
            let nombresLicencias = 'Unlicensed';
            if (tieneLicencias) {
              nombresLicencias = u.assignedLicenses
                .map((lic: any) => mapaLicencias[lic.skuId] || 'Licencia Asignada')
                .join(', ');
            }

            return {
              id: index + 1,
              nombre: u.displayName || 'Sin Nombre',
              email: u.userPrincipalName,
              fechaCreacion: u.createdDateTime ? new Date(u.createdDateTime).toLocaleDateString('es-CL') : 'N/A',
              rut: u.faxNumber || 'Sin RUT',
              licencia: nombresLicencias,
              ot: u.postalCode || 'Sin OT',
              activa: tieneLicencias,
              mfaHabilitado: index % 4 !== 0,
              diasInactivo: (index * 13) % 110
            };
          });

          this.accounts.set(usuariosMapeados);
        }

        this.cargando.set(false);
        this.ultimaSincronizacion.set(new Date().toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit' }));
        this.registrarAccionAuditoria('INICIO_SESION', 'Sincronización exitosa con Microsoft Graph API');
        setTimeout(() => this.conexionExitosa.set(true), 150);
      },
      error: (err) => {
        console.error('Error al consultar Microsoft Graph:', err);
        this.cargando.set(false);
        this.conexionExitosa.set(false);
        if (!navigator.onLine) {
          this.sinConexionRed.set(true);
        }
      }
    });
  }

  solicitarCambioEstado(id: number): void {
    const cuenta = this.accounts().find(acc => acc.id === id);
    if (!cuenta) return;

    this.cuentaSeleccionada.set(cuenta);
    this.motivoCambio.set('');
    this.mostrarModalConfirmacion.set(true);
  }

  confirmarCambioEstado(): void {
    const cuenta = this.cuentaSeleccionada();
    if (!cuenta) return;

    const motivo = this.motivoCambio().trim() || 'Sin motivo especificado';

    if (cuenta.activa) {
      this.registrarAccionAuditoria('BAJA', `Baja aprobada. Motivo: "${motivo}"`, cuenta.email);
    } else {
      this.registrarAccionAuditoria('REACTIVACION', `Reactivación aprobada. Motivo: "${motivo}"`, cuenta.email);
    }

    this.accounts.update(list =>
      list.map(acc => acc.id === cuenta.id ? { ...acc, activa: !acc.activa } : acc)
    );

    this.cerrarModalConfirmacion();
  }

  cerrarModalConfirmacion(): void {
    this.mostrarModalConfirmacion.set(false);
    this.cuentaSeleccionada.set(null);
    this.motivoCambio.set('');
  }

  exportarAExcel(): void {
    const datosExportar = this.cuentasFiltradas().map(acc => ({
      'ID': acc.id,
      'Nombre': acc.nombre,
      'Correo Office': acc.email,
      'Fecha Creación': acc.fechaCreacion,
      'RUT': acc.rut,
      'Licencia': acc.licencia,
      'OT': acc.ot,
      'Estado': acc.activa ? 'Activa' : 'Dada de baja',
      'MFA Habilitado': acc.mfaHabilitado ? 'Sí' : 'No',
      'Días Inactivo': acc.diasInactivo ?? 0
    }));

    const worksheet = XLSX.utils.json_to_sheet(datosExportar);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Cuentas_Office');

    const fechaHoy = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `Reporte_Cuentas_${fechaHoy}.xlsx`);

    this.registrarAccionAuditoria('EXPORTACION', `Se generó reporte Excel con ${datosExportar.length} cuentas`);
  }

  exportarLogAuditoriaExcel(): void {
    const datosLog = this.auditLogs().map(log => ({
      'ID Evento': log.id,
      'Fecha y Hora': log.fechaHora,
      'Administrador': log.usuarioAdmin,
      'Acción Realizada': log.accion,
      'Usuario Afectado': log.usuarioAfectado,
      'Detalles de Operación': log.detalles
    }));

    const worksheet = XLSX.utils.json_to_sheet(datosLog);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Historial_Auditoria');

    const fechaHoy = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `Auditoria_Seguridad_${fechaHoy}.xlsx`);
  }

  exportarInventarioLicencias(): void {
    const datosInventario = this.licenciasInventario().map(lic => ({
      'Suscripción / Licencia': lic.nombreComercial,
      'Total Compradas': lic.totalCompradas,
      'Asignadas': lic.asignadas,
      'Disponibles': lic.disponibles,
      'Porcentaje Consumo': `${lic.porcentajeUso}%`,
      'Estado Alerta': lic.estadoConsumo.toUpperCase()
    }));

    const worksheet = XLSX.utils.json_to_sheet(datosInventario);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Inventario_Licencias');

    const fechaHoy = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `Inventario_Licencias_${fechaHoy}.xlsx`);

    this.registrarAccionAuditoria('EXPORTACION', 'Se exportó el inventario completo de licencias');
  }
}

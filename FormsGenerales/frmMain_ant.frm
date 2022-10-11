VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "PLANALTO - SICOP"
   ClientHeight    =   9960
   ClientLeft      =   885
   ClientTop       =   2820
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   Picture         =   "frmMain.frx":6852
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15B9E
            Key             =   "Venta"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15EB8
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":161D2
            Key             =   "Producto"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":164EC
            Key             =   "Compra"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16806
            Key             =   "SolCompra"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16A98
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      DragIcon        =   "frmMain.frx":16DB2
      Height          =   630
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1111
      ButtonWidth     =   2223
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Productos"
            Key             =   "Producto"
            Object.ToolTipText     =   "Registro de Productos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Solicita Compra"
            Key             =   "SolCompra"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Compras"
            Key             =   "Compra"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pago Proveedor"
            Key             =   "Pago"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ventas"
            Key             =   "Venta"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte Venta"
            Key             =   "ReporteV"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   9690
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18653
            Text            =   "Estado"
            TextSave        =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "03/05/2009"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "08:49"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "SPC-Bolivia"
            TextSave        =   "SPC-Bolivia"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Cry 
      Left            =   0
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu MesaEntrada 
      Caption         =   "Mesa de Entrada"
      Begin VB.Menu PlanActividades 
         Caption         =   "Plan de Actividades"
      End
      Begin VB.Menu solicitudes 
         Caption         =   "Solicitudes"
         Visible         =   0   'False
         Begin VB.Menu Importar 
            Caption         =   "Importar  de  Solicitudes"
            Visible         =   0   'False
         End
         Begin VB.Menu Exportar 
            Caption         =   "Exportar Solicitudes"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu RegistroSolicitudes 
            Caption         =   "Registro de Solicitudes"
            Visible         =   0   'False
         End
         Begin VB.Menu FormularioF12 
            Caption         =   "S02 - Servicios Básicos y otros Gastos Administrativos"
            Enabled         =   0   'False
         End
         Begin VB.Menu FormularioF02 
            Caption         =   "Conformidad de Pago............. ............F02"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu FormularioF03 
            Caption         =   "S03 - Pasajes o Viáticos"
            Enabled         =   0   'False
         End
         Begin VB.Menu FormularioF04 
            Caption         =   "S04 - Adquisiciones por Licitación"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu FormularioF05 
            Caption         =   "S05 - Contratación de Personal"
            Enabled         =   0   'False
         End
         Begin VB.Menu FormularioF06 
            Caption         =   "Formulario.................................... ...F06"
            Visible         =   0   'False
         End
         Begin VB.Menu FormularioF07 
            Caption         =   "Orden de Cambio . . . . . . . . . . . . . . . . . .F07"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu FormularioF10 
            Caption         =   "Contratación Personal de Planta RR.HH. . .F10"
            Visible         =   0   'False
         End
         Begin VB.Menu imprecepcion 
            Caption         =   "Impresión de Solicitudes"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Clasificadores 
         Caption         =   "Clasificadores"
         Begin VB.Menu Generales 
            Caption         =   "Generales"
            Begin VB.Menu mnuBeneficiarios 
               Caption         =   "Clientes / Preveedores / Beneficiarios"
            End
            Begin VB.Menu UnidadesEjecutoras 
               Caption         =   "&Unidades"
               Shortcut        =   ^U
            End
            Begin VB.Menu Entidades 
               Caption         =   "E&ntidades"
               Shortcut        =   ^N
            End
            Begin VB.Menu TipoTramite 
               Caption         =   "Tipo de Trá&mite (Formularios)"
            End
            Begin VB.Menu CBeneficiarios 
               Caption         =   "Beneficiarios"
               Visible         =   0   'False
            End
            Begin VB.Menu DepartamentosBolivia 
               Caption         =   "Departamentos de Bolivia"
            End
            Begin VB.Menu ProvinciasDepartamentos 
               Caption         =   "Provincias de Departamentos"
            End
            Begin VB.Menu TiposErrores 
               Caption         =   "Tipos de Errores en Documentos"
            End
         End
         Begin VB.Menu CPresupuesto 
            Caption         =   "Gastos"
            Enabled         =   0   'False
            Begin VB.Menu PartidasGasto 
               Caption         =   "&Partidas del Gasto"
               Shortcut        =   ^P
            End
            Begin VB.Menu EconómicosGasto 
               Caption         =   "Económicos del Gasto"
               Visible         =   0   'False
            End
            Begin VB.Menu RelacionadorGastoEco 
               Caption         =   "Relacionador Gasto Eco. Vs. Partida"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu Presupuesto 
            Caption         =   "Presupuesto"
            Enabled         =   0   'False
            Begin VB.Menu FuentesFinanciamiento 
               Caption         =   "&Fuentes de Financiamiento"
               Shortcut        =   ^F
            End
            Begin VB.Menu OrganismosFinanciadores 
               Caption         =   "&Organismos Financiadores"
               Shortcut        =   ^O
            End
            Begin VB.Menu Convenios 
               Caption         =   "&Convenios de Financiamiento"
               Shortcut        =   ^C
            End
            Begin VB.Menu CategoriaFinanciadores 
               Caption         =   "Ca&tegoria de Financiadores"
               Shortcut        =   ^T
            End
            Begin VB.Menu EstructuraProgramatica 
               Caption         =   "&Proyectos"
               Shortcut        =   ^E
            End
            Begin VB.Menu Sisin 
               Caption         =   "S&isin"
               Shortcut        =   ^I
               Visible         =   0   'False
            End
            Begin VB.Menu Deducciones 
               Caption         =   "Deducciones/Retenciones"
               Visible         =   0   'False
            End
            Begin VB.Menu RelacionadorDonacionesOrganismos 
               Caption         =   "Relacionador Donaciones - Organismos"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu CTesoreria 
            Caption         =   "Tesoreria"
            Begin VB.Menu CBancos 
               Caption         =   "Bancos"
            End
            Begin VB.Menu CCtaBancarias 
               Caption         =   "Cuentas Bancarias"
            End
            Begin VB.Menu FormasPago 
               Caption         =   "Formas de Pago"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu Ingresos 
            Caption         =   "Ingresos"
            Enabled         =   0   'False
            Begin VB.Menu Rubros 
               Caption         =   "Rubros"
            End
            Begin VB.Menu EconómicosRecursos 
               Caption         =   "Económicos de Recursos"
               Visible         =   0   'False
            End
            Begin VB.Menu RelacionadorRubroEco 
               Caption         =   "Relacionador Rubro Eco. Vs Recursos"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu Contabilidad2 
            Caption         =   "Contabilidad"
            Enabled         =   0   'False
            Begin VB.Menu PlanCuentas 
               Caption         =   "Plan de Cuentas"
            End
            Begin VB.Menu RelacionadorCuentaPartidas 
               Caption         =   "Relacionador Cuenta Vs Partidas"
               Visible         =   0   'False
            End
            Begin VB.Menu RelacionadorIngresosCuentas 
               Caption         =   "Relacionador Cuentas Vs Ingresos"
               Visible         =   0   'False
            End
            Begin VB.Menu Depreciaciones 
               Caption         =   "Depreciaciones"
            End
            Begin VB.Menu ClaseAuxiliares 
               Caption         =   "Clase de Auxiliares"
            End
            Begin VB.Menu Inversiones 
               Caption         =   "Inversiones"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu Administrativos 
            Caption         =   "Administrativos"
            Visible         =   0   'False
            Begin VB.Menu Adquisiciones 
               Caption         =   "Adquisiciones"
               Enabled         =   0   'False
            End
            Begin VB.Menu Contrataciones 
               Caption         =   "Contrataciones"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu Almacenes2 
            Caption         =   "Almacenes"
            Begin VB.Menu mnugrupos 
               Caption         =   "Grupo de Bienes/Productos"
            End
            Begin VB.Menu mnuMontador 
               Caption         =   "Sub-Grupo de Bienes/Productos"
            End
            Begin VB.Menu mnuDetalle 
               Caption         =   "Bienes/Productos"
            End
            Begin VB.Menu mnuUnidadesMedida 
               Caption         =   "Unidades de Medida"
            End
            Begin VB.Menu mnuMarcas 
               Caption         =   "Marcas"
            End
            Begin VB.Menu mnuDestinos 
               Caption         =   "Destino de Entregas"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu RecursosHumanos2 
            Caption         =   "Recursos Humanos"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu Finanzas 
      Caption         =   "Procesos Financieros"
      Visible         =   0   'False
      Begin VB.Menu Procesos 
         Caption         =   "Egresos"
         Begin VB.Menu Compromiso 
            Caption         =   "Registros"
         End
         Begin VB.Menu Ejecucion 
            Caption         =   "Reportes"
            Begin VB.Menu EjecucionPpto 
               Caption         =   "Ejecucion Presupuestaria"
            End
            Begin VB.Menu EjecucionOrganismo 
               Caption         =   "Ejecucion Por Organismo, Convenio y Categoría"
               Enabled         =   0   'False
               Visible         =   0   'False
            End
            Begin VB.Menu EjecucionProyVsPpto 
               Caption         =   "Ejecucion por Organismo, Proyecto Vs. Ppto. de Ley"
               Enabled         =   0   'False
               Visible         =   0   'False
            End
            Begin VB.Menu EjecucionUniOrgProyPar 
               Caption         =   "Ejecución por Unidad, Organismo, Proyecto y Partida"
               Enabled         =   0   'False
               Visible         =   0   'False
            End
            Begin VB.Menu mnuEjecucionPOA 
               Caption         =   "Ejecución Física-Financiera"
            End
            Begin VB.Menu RepGraf 
               Caption         =   "Reportes Gráficos"
               Enabled         =   0   'False
               Visible         =   0   'False
               Begin VB.Menu repGrafUnidad 
                  Caption         =   "Por unidad"
               End
               Begin VB.Menu repGraforga 
                  Caption         =   "Por Organismo"
               End
            End
            Begin VB.Menu mnuEjecucionPresupuestaria 
               Caption         =   "Ejecución Presupuestaria Acumulada por..."
               Enabled         =   0   'False
               Visible         =   0   'False
               Begin VB.Menu EjecucionPorUni 
                  Caption         =   "Unidad"
               End
               Begin VB.Menu mnuConvenio 
                  Caption         =   "Convenio"
               End
               Begin VB.Menu mnuOrganismo 
                  Caption         =   "Organismo"
               End
               Begin VB.Menu mnuCategoria 
                  Caption         =   "Categoría"
               End
               Begin VB.Menu mnuProyecto 
                  Caption         =   "Proyecto"
               End
            End
         End
         Begin VB.Menu MnuModPpto 
            Caption         =   "Modificaciones Presupuestarias"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu CContabilidad 
         Caption         =   "Contabilidad"
         Begin VB.Menu Comprobantes 
            Caption         =   "Registros"
         End
         Begin VB.Menu mnuDescargos 
            Caption         =   "Descargos"
            Enabled         =   0   'False
            Begin VB.Menu mnuDSCRegistro 
               Caption         =   "Registro del Descargo"
            End
            Begin VB.Menu mnuDSCEC 
               Caption         =   "Reporte de Cargos y Descargos"
               Begin VB.Menu mnuDSCECBenef 
                  Caption         =   "Por Beneficiario"
               End
               Begin VB.Menu mnuDSCECConv 
                  Caption         =   "Por Convenios"
               End
               Begin VB.Menu mnuDSCResumen 
                  Caption         =   "Resumen"
               End
            End
         End
         Begin VB.Menu ReportesC 
            Caption         =   "Reportes"
            Begin VB.Menu LibroMayor 
               Caption         =   "Libro Mayor"
            End
            Begin VB.Menu LibroMayorAuxiliar 
               Caption         =   "Libro Mayor Auxiliar"
            End
            Begin VB.Menu mnuLMGral 
               Caption         =   "Libro Mayor General"
            End
            Begin VB.Menu EstadoResultados 
               Caption         =   "Estado de Resultados"
            End
            Begin VB.Menu BalanceSumasSaldos 
               Caption         =   "Balance de Sumas y Saldos"
            End
            Begin VB.Menu BalanceGeneral 
               Caption         =   "Balance General"
            End
            Begin VB.Menu mnuRepBalApertura 
               Caption         =   "Balance de Apertura"
            End
         End
      End
      Begin VB.Menu Tesoreria 
         Caption         =   "Tesorería"
         Begin VB.Menu Gastos 
            Caption         =   "Gastos (Pagos)"
         End
         Begin VB.Menu mnucaja 
            Caption         =   "Caja"
            Begin VB.Menu OperaciónCheques 
               Caption         =   "Seguimiento de Cheques"
            End
            Begin VB.Menu mnuTriFis 
               Caption         =   "Tributos Fiscales"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu CuentasBancarias2 
            Caption         =   "Cuentas Bancarias"
            Begin VB.Menu MovCta 
               Caption         =   "Movimiento de Cuentas Bancarias"
            End
            Begin VB.Menu SaldosActuales 
               Caption         =   "Saldos Actuales"
               Visible         =   0   'False
            End
            Begin VB.Menu ic 
               Caption         =   "Traspasos Cuentas Bancarias"
            End
         End
         Begin VB.Menu Reportes_tesoreria 
            Caption         =   "Reportes"
            Begin VB.Menu ConsultasPagos 
               Caption         =   "Consultas Pagos"
               Begin VB.Menu PagosEfectuados 
                  Caption         =   "Pagos"
               End
               Begin VB.Menu PagosEfectuadosRealizar 
                  Caption         =   "Pagos Efectuados y por Realizar"
               End
               Begin VB.Menu MnuLC 
                  Caption         =   "Listado de Comprobantes"
               End
            End
            Begin VB.Menu ManejoCheques 
               Caption         =   "Impresión de Comprobantes de Pago"
               Begin VB.Menu PorColaImpresion 
                  Caption         =   "Por Cola de Impresión"
               End
               Begin VB.Menu PorSeleccionComprobantes 
                  Caption         =   "Por Selección de Comprobantes"
               End
            End
            Begin VB.Menu ct 
               Caption         =   "Impresión de Transferencias"
            End
            Begin VB.Menu MnuImpCheques 
               Caption         =   "Impresión de cheques"
            End
         End
         Begin VB.Menu mnuConcilia 
            Caption         =   "Conciliación"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPF 
            Caption         =   "Programación Financiera"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAnulacion999 
            Caption         =   "Anulación Cargos de Cuenta"
         End
      End
      Begin VB.Menu MnuIngresos 
         Caption         =   "Ingresos"
         Begin VB.Menu EjecucionPresupuestaria 
            Caption         =   "Registros"
         End
         Begin VB.Menu SOES 
            Caption         =   "Registro SOES"
            Visible         =   0   'False
         End
         Begin VB.Menu confirsoes 
            Caption         =   "Confirmación SOES"
            Visible         =   0   'False
         End
         Begin VB.Menu ReportesIngresos 
            Caption         =   "Reportes"
         End
      End
   End
   Begin VB.Menu Administracion 
      Caption         =   "Procesos Administrativos"
      Begin VB.Menu MnuAdmiCont 
         Caption         =   "Administración de Contratos"
         Visible         =   0   'False
      End
      Begin VB.Menu ComprasDirectasm 
         Caption         =   "Compras"
         Begin VB.Menu FormularioF11 
            Caption         =   "Solicitud de Compra C-01"
            Visible         =   0   'False
         End
         Begin VB.Menu FormularioF01 
            Caption         =   "Solicitudes de Compra"
         End
         Begin VB.Menu pp 
            Caption         =   "Pago a Proveedores"
         End
         Begin VB.Menu mnuRegIniCompra 
            Caption         =   "Registro de Compras"
         End
         Begin VB.Menu MnuCompras 
            Caption         =   "Reportes de Compras"
         End
         Begin VB.Menu MnuAdjCompra 
            Caption         =   "Adjudicacion"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuLiqCompra 
            Caption         =   "Liquidaciones Compras"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuConsultoria_c 
         Caption         =   "Contratacion de Consultores"
         Visible         =   0   'False
         Begin VB.Menu mnuSolNoObj_c 
            Caption         =   "Registro Inicial"
         End
         Begin VB.Menu mnuPublicacion_c 
            Caption         =   "Publicación"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuVentaPliegos_c 
            Caption         =   "Venta de Pliegos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecepcionProp_c 
            Caption         =   "Recepción de Propuestas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAperturaProp_c 
            Caption         =   "Apertura de Propuestas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAdjudicacion_c 
            Caption         =   "Adjudicación (Contratación)"
         End
         Begin VB.Menu mnugesPagos 
            Caption         =   "Gestión de pagos (Planillas)"
         End
         Begin VB.Menu mnuOC 
            Caption         =   "Ordenes de cambio"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSolicitudesF04 
         Caption         =   "Solicitudes F04"
         Visible         =   0   'False
      End
      Begin VB.Menu ProgConAdq 
         Caption         =   "Programación Contrataciones y Adquisiciones"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuVentas 
         Caption         =   "Ventas"
         Begin VB.Menu MnuVentas2 
            Caption         =   "Proceso de Ventas"
         End
         Begin VB.Menu MnuVentas3 
            Caption         =   "Ventas Pendientes Sist. Anterior"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuRepVentas 
            Caption         =   "Reportes de Ventas"
         End
      End
      Begin VB.Menu mnuALMACENES 
         Caption         =   "Almacenes"
         Begin VB.Menu mnuRegistro 
            Caption         =   "Registro"
            Begin VB.Menu mnuIngreso 
               Caption         =   "Ingreso"
               Begin VB.Menu mnuIngresoManual 
                  Caption         =   "Ingreso Manual"
               End
               Begin VB.Menu mnuIngresodecompras 
                  Caption         =   "Ingreso de Compras"
               End
            End
            Begin VB.Menu mnuEntrega 
               Caption         =   "Entrega"
            End
         End
         Begin VB.Menu mnuControl 
            Caption         =   "Control"
            Begin VB.Menu mnuInventario 
               Caption         =   "Inventario"
            End
            Begin VB.Menu mnuEstadoAlmacen 
               Caption         =   "Estado Almacen"
            End
         End
      End
   End
   Begin VB.Menu InformacionGerencial 
      Caption         =   "Informacion Gerencial"
      Visible         =   0   'False
      Begin VB.Menu mnuIGPpto 
         Caption         =   "Presupuestos"
         Begin VB.Menu mnupptoEjecxUnixConv 
            Caption         =   "Ejecución por Unidad y Convenio"
         End
         Begin VB.Menu mnupptoEjecxConxOGasto 
            Caption         =   "Ejecución Por Convenio y Objeto del Gasto"
         End
         Begin VB.Menu mnupptoEjecxOrg 
            Caption         =   "Ejecución por Organismo"
         End
      End
      Begin VB.Menu mnuIGtes 
         Caption         =   "Tesoreria"
         Begin VB.Menu mnuTInfCtaaPPendiente 
            Caption         =   "Informe de Ctas y Pagos a Realizar"
         End
      End
      Begin VB.Menu mnuConta 
         Caption         =   "Contabilidad"
         Begin VB.Menu mnuContaPresInterConv 
            Caption         =   "Prestamos Inter-Convenio"
         End
      End
      Begin VB.Menu mnuIGIng 
         Caption         =   "Ingresos"
         Begin VB.Menu mnuEjecIng 
            Caption         =   "Ejecución de Ingresos y Saldos"
         End
      End
      Begin VB.Menu mnuLic 
         Caption         =   "Licitaciones"
         Begin VB.Menu mnuLicEstLic 
            Caption         =   "Estado de las Licitaciones"
         End
      End
      Begin VB.Menu mnuCD 
         Caption         =   "Compras Directas"
         Begin VB.Menu mnuCDEstComp 
            Caption         =   "Estado de las Compras"
         End
      End
      Begin VB.Menu mnuCons 
         Caption         =   "Consultoria"
         Visible         =   0   'False
         Begin VB.Menu mnuEstCons 
            Caption         =   "Estado de las Consultorías"
         End
      End
      Begin VB.Menu mnuPrg 
         Caption         =   "Prog de Contrataciones"
         Visible         =   0   'False
         Begin VB.Menu mnuDesxUni 
            Caption         =   "Desempeño por Unidad Ejecutora"
         End
         Begin VB.Menu mnuDesConvo 
            Caption         =   "Desempeño por Tipo de Convocatoria"
         End
      End
      Begin VB.Menu mnuPoa 
         Caption         =   "Poa"
         Visible         =   0   'False
         Begin VB.Menu mnuPoaCumPoa 
            Caption         =   "Cumplimiento del POA"
         End
      End
   End
   Begin VB.Menu Regularizaciones 
      Caption         =   "Regularizaciones"
      Visible         =   0   'False
      Begin VB.Menu PagosDirectos 
         Caption         =   "Pagos Directos"
      End
      Begin VB.Menu ComisionesBancarias 
         Caption         =   "Comisiones Bancarias"
      End
   End
   Begin VB.Menu MNUSeguimiento 
      Caption         =   "Seguimiento"
   End
   Begin VB.Menu mnuAdmiSistema 
      Caption         =   "Administracion Sistema"
      Begin VB.Menu mnuCambiarClave 
         Caption         =   "Cambiar Clave"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuNivelAcceso 
         Caption         =   "Nivel de Acceso"
      End
      Begin VB.Menu mnuPrivAcceso 
         Caption         =   "Privilegios de Operación"
      End
   End
   Begin VB.Menu Salida 
      Caption         =   "Salida"
      Begin VB.Menu mnuAcercade 
         Caption         =   "A cerca de ..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepara 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir del Sistema"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------- Programas en comentario de JORGE

Private Sub CuentasBancarias_Click()
    Dim e As Long
    e = Shell(App.Path & "\Reportes Tesoreria\cnsPagados.exe", 1)
End Sub

Private Sub BalanceGeneral_Click()
  frmBalanceGral.Show
End Sub

Private Sub BalanceSumasSaldos_Click()
  frmsumsaldos.Show
End Sub

'Private Sub CBeneficiarios_Click() 'clasifica
'  frmBeneficiario.Show
'End Sub

Private Sub CCtaBancarias_Click()
  'CLFrmCtaBco.Show
  FrmCtaBco.Show
End Sub

'Private Sub Clasificadores_Click()
''  FrmCtaBco.Show
''  If UCase(GlUsuario) = "SAF" Or UCase(GlUsuario) = "-" Or UCase(GlUsuario) = "-" Then
'    Dim e As Long
'   ' e = Shell(App.Path & "\Clasificadores\clasificadores.exe", 1)
'
''  Else
''      MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
''  End If
'End Sub

Private Sub cmpte_nuevo_Click()
  FrmImprimirComprobante.Show
End Sub

Private Sub ComisionesBancarias_Click()
  Dim e As Long
  e = Shell(App.Path & "\Comisiones Bancarias\proj1.exe", 1)
End Sub

'Private Sub ComprasDirectasm_Click()
    'frmadjudicacionD.Show
'    frmComprasDirectas.Show
'End Sub

Private Sub confirsoes_Click()


    Dim e As Long
    e = Shell(App.Path & "\FormsIngresos\soes\SOES.exe " & GlUsuario & " DEDUCCIONES", 1)
End Sub

Private Sub Cotizaciones_Click()
  frmpliegosSol.Show
End Sub

Private Sub CuadroComparativo_Click()
  frmadjudicacionD.Show
End Sub

Private Sub EjecucionPorUni_Click()
'  glRepPresup = "REP004"
'  frmRepPresupuesto.Show
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP004", 1)
End Sub

Private Sub EjecucionPpto_Click()
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\ProyRepPresupuesto.exe", 1)
End Sub

Private Sub EjecucionProyVsPpto_Click()
'  glRepPresup = "REP002"
'  frmRepPresupuesto.Show
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP002", 1)
End Sub

Private Sub EjecucionUniOrgProyPar_Click()
'  glRepPresup = "REP003"
'  frmRepPresupuesto.Show
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP003", 1)
End Sub

'Private Sub Entidades_Click() 'clasifica
'  Dim e As New Entidades
'  e.Show '  Entidades.se
'End Sub

Private Sub EstadoResultados_Click()
  frmEERR.Show
End Sub

Private Sub Exportar_Click()
  Frmexporta.Show vbModal
End Sub

Private Sub FormularioF01_Click()
  FrmF01.Show 'vbModal
End Sub

Private Sub FormularioF02_Click()
'  FrmF02.Show vbModal
'  ac_conformidad.Show

'  ac_PagosMain.cmdAdicionarBeneficiario.Visible = False
'  ac_PagosMain.cmdAdicionarGrupo.Visible = False
'  ac_PagosMain.cmdAdicionarPago.Visible = False
'  ac_PagosMain.cmdBorraGrupo.Visible = False
'  ac_PagosMain.cmdEliminarBeneficiario.Visible = False
'  ac_PagosMain.cmdModificarGrupo.Visible = False
'  ac_PagosMain.cmdEliminarPago.Visible = False
'  ac_PagosMain.cmdEliminarBeneficiario.Visible = False
'  ac_PagosMain.cmdEliminarPago.Visible = False
'  ac_PagosMain.cmdModificarPago.Visible = False
'  ac_PagosMain.cmdAprobadoParaEnvio.Visible = False
'  glProceso = "CONSULTORIA"
''  glProceso = "RECURSOS HUMANOS"
'  ac_PagosMain.Show vbModal
    ac_conformidad.Show
End Sub

Private Sub FormularioF03_Click()
  FrmF03.Show 'vbModal
End Sub

Private Sub FormularioF04_Click()
  FrmF04.Show 'vbModal
End Sub

Private Sub FormularioF05_Click()
  FrmF05.Show 'vbModal
End Sub

Private Sub FormularioF06_Click()
  FrmF06.Show 'vbModal
End Sub

Private Sub FormularioF07_Click()
  FrmF07.Show 'vbModal
End Sub

Private Sub FormularioF10_Click()
    FrmF10.Show
End Sub

Private Sub FormularioF11_Click()
    FrmF11.Show
End Sub

Private Sub FormularioF12_Click()
    FrmF12.Show
End Sub

Private Sub Gastos_Click()
  'FrmCuentaBancaria.Show
  FrmListadoPagos.Show
End Sub

Private Sub gesPagos_Click()
glProceso = "CONSULTORIA"
ac_PagosMain.Show vbModal
End Sub

Private Sub ic_Click()
  ' con G--
  frmtraspasos.Show
End Sub

Private Sub Importar_Click()
  FrmImporta.Show vbModal
End Sub

Private Sub imprecepcion_Click()
  Dim rstao_solicitud_recibido As New ADODB.Recordset
  Dim sino As String
  
  Set rstao_solicitud_recibido = New ADODB.Recordset
  If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
  rstao_solicitud_recibido.Open "select * from ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
  Print rstao_solicitud_recibido.RecordCount
'-------
  '  Cry.Reset
  Cry.ReportFileName = App.Path & "\FormsMigrar\RecepMigrar.rpt"
'  Cry.SelectionFormula = "{Vi_Fo_ingresos_rep.Maquina} = '" & GlMaquina & "'"

  Cry.WindowShowPrintBtn = True
  Cry.WindowShowExportBtn = True
  Cry.WindowShowRefreshBtn = True
  Cry.WindowShowPrintSetupBtn = True
  Cry.WindowShowZoomCtl = True
  Cry.WindowState = crptMaximized
  Cry.PageZoom (200)
  IResult = Cry.PrintReport
  If IResult <> 0 Then
      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
'  sino = MsgBox("¿La impresión concluyó con EXITO?", vbQuestion + vbYesNo, "Confirmando Impresión... ")
'  If sino = vbYes Then
'  rstao_solicitud_recibido.MoveFirst
    While Not rstao_solicitud_recibido.EOF
      rstao_solicitud_recibido.Delete
      rstao_solicitud_recibido.Update
      rstao_solicitud_recibido.MoveNext
    Wend
'  End If
  If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close

End Sub

Private Sub LibroMayor_Click()
  frmLMayor.Show
End Sub

Private Sub LibroMayorAuxiliar_Click()
  FrmLMayorAux.Show
End Sub

Private Sub mbuSolNoObj_c_Click()
glProceso = "CONSULTORIA"
ac_NoObjecion_c1.Show vbModal
End Sub

Private Sub mnuabogado_Click()
 'frmaperturasobresAbogado.Show
 al_aperturasobresAbogado.Show
End Sub

Private Sub mnuAcercade_Click()
  frmAbout.Show vbModal
End Sub

Private Sub MnuAdjudica_Click()
  'frmadjudicacion.Show vbModal
  al_adjudicacion.Show vbModal
End Sub

Private Sub MnuAdjCompra_Click()
''llama al proceso principal de contratacion de licitación -Adett
    Screen.MousePointer = vbHourglass
    glProceso = "F11"
    af_ProponenteCMAdjudica_M.Show vbModal
End Sub

Private Sub mnuAdjudicacion_c_Click()
'    glProceso = "RECURSOS HUMANOS"
'    ac_Adjudicacion_c.Show vbModal
End Sub

Private Sub MnuAdmiCont_Click()
  CtFrmContratos.Show
End Sub

Private Sub mnuAnulacion999_Click()
  frmAnulacion999.Show
End Sub

Private Sub mnuAperturaProp_c_Click()
glProceso = "CONSULTORIA"
ac_Apertura_Propuestas_c.Show vbModal
End Sub

Private Sub MnuAperturaSobres_Click()
  frmaperturasobres.Show vbModal
End Sub

Private Sub mnuBeneficiarios_Click()
    frmBeneficiario.Show 'vbModal
End Sub

Private Sub mnuCambiarClave_Click()
    FrmCambiarClave.Show vbModal
    'MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub mnuCategoria_Click()
'  glRepPresup = "REP007"
'  frmRepPresupuesto.Show
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP007", 1)
End Sub

Private Sub mnuCD2_Click()
  frmComprasDirectas.Show
End Sub

Private Sub mnuCDEstComp_Click()
  arIGComDir.DataControl1.ConnectionString = db
  arIGComDir.Printer.Orientation = ddOLandscape
  arIGComDir.Show

End Sub

Private Sub mnucomision_Click()
 'frmComision.Show
 al_Comision.Show
End Sub

Private Sub mnuCompras_Click()
'    With ALFrmIngDeLici
'        .Show vbModal
'        If .QResp Then
'            With AlFrmIngresoMaterial
'                .ALPrincipal 2, IngresoDeLicitacion(ALFrmIngDeLici.NoLicitacion)
'            End With
'        End If
'    End With
'    MsgBox "Error en despliegue de pantalla, consulte con el Administrador del Sistema"
    frmComprasReportes.Show
End Sub

Private Sub mnuConcilia_Click()
    FrmExplorador.Show
End Sub

Private Sub mnuContaPresInterConv_Click()
  IgCtasInterconvenio.Show
End Sub

Private Sub mnuConvenio_Click()
'  glRepPresup = "REP005"
'  frmRepPresupuesto.Show
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP005", 1)
End Sub

Private Sub MnuCorrCheques_Click()
  FrmCorrelativos.Show
End Sub

Private Sub mnuDestinos_Click()
  ALFrmCLDestinos.Show
End Sub

Private Sub mnuDetalle_Click()
  'AlFrmCreaMaterial.ALPrincipal 0
  AlFrmCreaMaterial.Show
End Sub

Private Sub mnuDSCECBenef_Click()
'    frmDDEstadoBeneficiario.Show
    frmDDRepBeneficiario.Show
End Sub

Private Sub mnuDSCECConv_Click()
    frmDDEstadoxConvenio.Show
End Sub

Private Sub mnuDSCRegistro_Click()
    Dim e As Long
    e = Shell(App.Path & "\DESCARGOS\descargos.exe", 1)
End Sub

Private Sub mnuDSCResumen_Click()
    frmDDResumen.Show
End Sub

Private Sub mnuEjecIng_Click()

  Dim rsv_Ingreso_Convenio2 As New ADODB.Recordset
  Dim rsIG_Ing_EjePptoConvenio As New ADODB.Recordset
  Dim consulta1 As String
  consulta1 = ""
  ' ANTES CON FECHAS consulta1 = "where (fecha_registro >= '" & DTPkFechaInicio & "' and fecha_registro <= '" & DTPkFechaFin & "') "
'  If Len(Trim(DtCcodigo_convenio.Text)) > 0 Then
    'ANTES CON FECHAS
    'consulta1 = consulta1 & " and codigo_convenio = '" & Trim(DtCcodigo_convenio.Text) & "' "
    ' AHORA SOLO CONVENIO
'    consulta1 = " WHERE codigo_convenio = '" & Trim(DtCcodigo_convenio.Text) & "' "
'  Else
'  End If
  Set rsv_Ingreso_Convenio2 = New ADODB.Recordset
  If rsv_Ingreso_Convenio2.State = 1 Then rsv_Ingreso_Convenio2.Close
  rsv_Ingreso_Convenio2.Open "select * from v_Ingreso_Convenio2 " & consulta1 & " order by codigo_convenio", db, adOpenKeyset, adLockReadOnly
  Print rsv_Ingreso_Convenio2.RecordCount
  Set rsIG_Ing_EjePptoConvenio = New ADODB.Recordset
  db.Execute "DELETE FROM IG_Ing_EjePptoConvenio WHERE maquina = '" & GlMaquina & "'"
  Set rsIG_Ing_EjePptoConvenio = New ADODB.Recordset
  rsIG_Ing_EjePptoConvenio.Open "select * from IG_Ing_EjePptoConvenio where maquina = '" & GlMaquina & "'", db, adOpenKeyset, adLockOptimistic
  While Not rsv_Ingreso_Convenio2.EOF
    rsIG_Ing_EjePptoConvenio.AddNew
    rsIG_Ing_EjePptoConvenio!codigo_convenio = rsv_Ingreso_Convenio2!codigo_convenio
    rsIG_Ing_EjePptoConvenio!Cta_Codigo = rsv_Ingreso_Convenio2!Cta_Codigo
    rsIG_Ing_EjePptoConvenio!org_codigo = rsv_Ingreso_Convenio2!org_codigo
    rsIG_Ing_EjePptoConvenio!estado_recaudado = rsv_Ingreso_Convenio2!estado_recaudado
    rsIG_Ing_EjePptoConvenio!monto_dolares = rsv_Ingreso_Convenio2!monto_dolares
    rsIG_Ing_EjePptoConvenio!monto_Bolivianos = rsv_Ingreso_Convenio2!monto_Bolivianos
    rsIG_Ing_EjePptoConvenio!monto_formulado_us = rsv_Ingreso_Convenio2!monto_formulado_us
    rsIG_Ing_EjePptoConvenio!monto_vigente_us = rsv_Ingreso_Convenio2!monto_vigente_us
    rsIG_Ing_EjePptoConvenio!monto_compromiso_us = rsv_Ingreso_Convenio2!monto_compromiso_us
    rsIG_Ing_EjePptoConvenio!monto_devengado_us = rsv_Ingreso_Convenio2!monto_devengado_us
    rsIG_Ing_EjePptoConvenio!monto_pagado_us = rsv_Ingreso_Convenio2!monto_pagado_us
    rsIG_Ing_EjePptoConvenio!codigo_convenio = rsv_Ingreso_Convenio2!codigo_convenio
    rsIG_Ing_EjePptoConvenio!org_codigo_cta = rsv_Ingreso_Convenio2!org_codigo_cta
    rsIG_Ing_EjePptoConvenio!maquina = GlMaquina
    rsIG_Ing_EjePptoConvenio.Update
    rsv_Ingreso_Convenio2.MoveNext
  Wend

  Cry.WindowShowRefreshBtn = True
  Cry.ReportFileName = App.Path & "\InfGerencial\Ingresos\Rpt_Ingresos_convenio3.rpt"
'  Cry.ReportFileName = App.path & "\Rpt_Ingresos_convenio3.rpt"
  
  Cry.WindowShowPrintBtn = True
  Cry.WindowShowExportBtn = True
  Cry.WindowShowPrintSetupBtn = True
  Cry.WindowShowGroupTree = True
  Cry.WindowState = crptMaximized
  IResult = Cry.PrintReport
  If IResult <> 0 Then
      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

End Sub

Private Sub mnuEjecucionPOA_Click()
  Dim e As Long
'  e = Shell(App.Path & "\Reportes\poa\POA.exe")
End Sub

Private Sub mnuEstado_Click()
  ALFrmAlmacen.Show
End Sub

Private Sub mnuEntrega_Click()
 'AlmFrmSalidaMaterial.ALPrincipal 0
    MsgBox "Error en resolucion de pantalla, consulte con el Administrador.."
End Sub

Private Sub mnuEstadoAlmacen_Click()
    ALFrmAlmacen.Show
End Sub

Private Sub mnugesPagos_Click()
'  glProceso = "CONSULTORIA"
'  ac_PagosMain.Show vbModal
End Sub

Private Sub mnuGrupos_Click()
  AlmFrmCLGrupos.Show
End Sub

Private Sub MnuImpCheques_Click()
'    Dim e As Long
'    e = Shell("c:\SAF-2002\Reportes Tesoreria\cnsTotal.exe", 1)
    FrmChequesNuevo.Show
End Sub

Private Sub EjecucionOrganismo_Click()
'  glRepPresup = "REP001"
' frmRepPresupuesto.Show
    Dim e As Long
    e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP001", 1)
End Sub

Private Sub mnuIngreso_Click()
' ALB
  'AlFrmIngresoMaterial.ALPrincipal 0
End Sub

Private Sub mnuIngresodecompras_Click()
' With ALFrmIngDeLici
'        .Show vbModal
'        If .QResp Then
'            With AlFrmIngresoMaterial
'                .ALPrincipal 2, IngresoDeLicitacion(ALFrmIngDeLici.NoLicitacion)
'            End With
'        End If
'    End With
    MsgBox "Error en resolucin de pantalla, consulte con el Administrador.."
End Sub

Private Sub mnuIngresoManual_Click()
'ALB
'AlFrmIngresoMaterial.ALPrincipal 0
    MsgBox "Error en resolucin de pantalla, consulte con el Administrador.."
End Sub

Private Sub mnuInventario_Click()
  AlmFrmInventario.Show
End Sub

Private Sub MnuLC_Click()
  FrmPagosProyectos.Show
End Sub

Private Sub mnuLicEstLic_Click()
  arIGLicitaciones.DataControl1.ConnectionString = db
  arIGLicitaciones.Printer.Orientation = ddOLandscape
  arIGLicitaciones.Show
End Sub

Private Sub MnuLiqCompra_Click()
''llama al proceso principal de contratacion de licitación -Adett
    Screen.MousePointer = vbHourglass
    glProceso = "F11"
    af_LiquidaMain_M.Show vbModal
End Sub

Private Sub mnuLMGral_Click()
  FrmLMayorGral.Show
End Sub

Private Sub mnuMarcas_Click()
AlFrm_Marcas.Show
End Sub

Private Sub MnuModPpto_Click()

Dim e As Long
  e = Shell(App.Path & "\FORMULACION\formulacion.exe", 1)
End Sub

Private Sub mnuMontador_Click()
    ALFrm_montado.Show
End Sub

Private Sub mnuNivelAcceso_Click()
    FrmNivelesAcceso.Show
'    MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub mnuOC_Click()
glProceso = "CONSULTORIA"
ac_NoObjecion_OC.Show vbModal
End Sub

Private Sub mnuorden_Click()
 'frmOrdendePago.Show vbModal
 al_PagosMain.Show vbModal
End Sub

Private Sub mnuOrganismo_Click()
'  glRepPresup = "REP006"
'  frmRepPresupuesto.Show
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP006", 1)
End Sub

Private Sub mnuPF_Click()
    FrmPF.Show
End Sub

Private Sub mnupptoEjecxConxOGasto_Click()
  db.Execute " EXEC pto_ActualizaFormulacionGastoCOM "
  db.Execute " EXEC pto_ActualizaFormulacionGastoDEV "
  db.Execute " EXEC pto_ActualizaFormulacionGastoPAG "
  Call IGrptEjepptoConPar
  Cry.WindowShowRefreshBtn = True
  Cry.ReportFileName = App.Path & "\InfGerencial\presupuestos\IG_ejePpto_Conv_Par.rpt"
  
  Cry.WindowShowPrintBtn = True
  Cry.WindowShowExportBtn = True
  Cry.WindowShowPrintSetupBtn = True
  Cry.WindowShowGroupTree = True
  Cry.WindowState = crptMaximized
  IResult = Cry.PrintReport
  If IResult <> 0 Then
      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
End Sub

Private Sub mnupptoEjecxOrg_Click()
  db.Execute " EXEC pto_ActualizaCategoriaCOM "
  db.Execute " EXEC pto_ActualizaCategoriaDEV "
  db.Execute " EXEC pto_ActualizaCategoriaPAG "
  Cry.WindowShowRefreshBtn = True
  Cry.ReportFileName = App.Path & "\InfGerencial\presupuestos\IG_EjecPptoCat_fin.rpt"
  Cry.WindowShowPrintBtn = True
  Cry.WindowShowExportBtn = True
  Cry.WindowShowPrintSetupBtn = True
  Cry.WindowShowGroupTree = True
  Cry.WindowState = crptMaximized
  IResult = Cry.PrintReport
  If IResult <> 0 Then
      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

End Sub

Private Sub mnupptoEjecxUnixConv_Click()
  Cry.WindowShowRefreshBtn = True
  Cry.ReportFileName = App.Path & "\InfGerencial\presupuestos\IG_ejePpto_Conv_uni0.rpt"
  
  Cry.WindowShowPrintBtn = True
  Cry.WindowShowExportBtn = True
  Cry.WindowShowPrintSetupBtn = True
  Cry.WindowShowGroupTree = True
  Cry.WindowState = crptMaximized
  IResult = Cry.PrintReport
  If IResult <> 0 Then
      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

End Sub

Private Sub mnuProyecto_Click()
'  glRepPresup = "REP008"
'  frmRepPresupuesto.Show
  Dim e As Long
  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP008", 1)
End Sub

Private Sub mnuPublicacion_c_Click()
glProceso = "CONSULTORIA"
ac_Publicacion_c.Show vbModal
End Sub

Private Sub mnupublicacion_Click()
  'frmPublicacion1.Show vbModal
  al_Publicacion1.Show vbModal
End Sub

Private Sub mnuRecepcionProp_c_Click()
  glProceso = "CONSULTORIA"
  ac_RecepcionSobres_c.Show vbModal
End Sub

Private Sub MnuRecepSobres_Click()
  'frmrecepcionsobres.Show vbModal
  al_recepcionsobres.Show vbModal
End Sub

Private Sub mnuRegIniCompra_Click()
'''llama al proceso principal de contratacion de licitación
'    Screen.MousePointer = vbHourglass
'    glProceso = "F11"
'    af_IniciaCompraMenor_M.Show vbModal

    frmComprasDirectas.Show
End Sub

Private Sub mnuRepBalApertura_Click()
   cc_balapertura.Show
End Sub

Private Sub mnuSalidas_Click()
  AlmFrmSalidaMaterial.ALPrincipal 0
End Sub

Private Sub MnuRepVentas_Click()
    frmVentasReportes.Show
    'MsgBox "Error en despliegue de pantalla, consulte con el Administrador del Sistema"
End Sub

Private Sub mnuSalir_Click()
   Unload Me
   End
End Sub

Private Sub MNUSeguimiento_Click()
'frmadjudicacionD Frm_Seguimiento.Show
End Sub

Private Sub mnuSolicitudesF04_Click()
  al_SolicitudesF04.Show vbModal
End Sub

Private Sub mnuSolNoObj_c_Click()
'  glProceso = "CONSULTORIA"
'  ac_NoObjecion_c1.Show vbModal
End Sub

Private Sub MnuSolNoObjec_Click()
  'frmNoObjecion.Show vbModal
  al_NoObjecion.Show vbModal
End Sub

Private Sub mnuTInfCtaaPPendiente_Click()
  FrmPagosARealizar.Show
End Sub

Private Sub mnuTriFis_Click()
  FrmTributosFiscales.Show
End Sub

Private Sub mnuUnidadesMedida_Click()
AlFrm_Ing_UnidadMedida.Show
End Sub

Private Sub mnuUsuarios_Click()
    FrmSisUsuarios.Show
    'MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub Comprobantes_Click()
    frm_ManualConta.Show
End Sub

Private Sub Compromiso_Click()
  BuscaTipoAcceso "compromiso"
  FrmRegularizacion.Show
End Sub

Private Sub ct_Click()
    FrmTransferenciasNuevo.Show
End Sub

Private Sub dc_Click()
    If NombreUsuario = "jcc001" Or NombreUsuario = "JCC002" Or NombreUsuario = "MYB159" Then
        FrmDesactivacionCheques.Show
    Else
        MsgBox "El Usuario NO está autorizado . . . "
    End If
End Sub

'Private Sub DepartamentosBolivia_Click() 'clasifica
'    frmDepto.Show
'End Sub

Private Sub EconómicosGasto_Click()
    frmEcoGasto.Show
End Sub

Private Sub EjecucionPresupuestaria_Click()
    usuario2 = GlUsuario
    FrmIngresosabm.Show
End Sub

Private Sub MDIForm_Load()
   Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
   Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
   Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
   Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
   LoadNewDoc
End Sub

Private Sub LoadNewDoc()
   'Static lDocumentCount As Long
   'Dim frmD As frmDocument
   'lDocumentCount = lDocumentCount + 1
   'Set frmD = New frmDocument
   'frmD.Caption = "Document " & lDocumentCount
   'frmD.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If Me.WindowState <> vbMinimized Then
      SaveSetting App.Title, "Settings", "MainLeft", Me.Left
      SaveSetting App.Title, "Settings", "MainTop", Me.Top
      SaveSetting App.Title, "Settings", "MainWidth", Me.Width
      SaveSetting App.Title, "Settings", "MainHeight", Me.Height
   End If
   'gerardo
   'rsPrm.Close
'   db.Close
End Sub

Private Sub mnuDatafrmDocument_Click()
'   Dim f As New frmDocument
'   f.Show
End Sub

Private Sub PCE_Click()
Flag_TipComp = "PCE"

End Sub

Private Sub MnuVentaPliegos_Click()
  'frmpliegos.Show vbModal
  al_pliegos.Show vbModal
End Sub


Private Sub MnuVentas2_Click()
    FrmVentas.Show
End Sub

Private Sub MnuVentas3_Click()
    FrmVentas_H.Show
End Sub

Private Sub MovCta_Click()
  FrmCuentas.Show
End Sub

Private Sub OperaciónCheques_Click()
    FrmActivacionCheques.Show
End Sub

Private Sub OrdenPagoCD_Click()
  frmOrdendePagoD.Show
End Sub

Private Sub PagosDirectos_Click()
     Dim e As Long
     e = Shell(App.Path & "\pagos directos\principal.EXE", 1)
'     e = Shell("D:\Archivos de programa\gtz\PD\principal.EXE", 1)
End Sub

Private Sub PagosEfectuados_Click()
    FrmPagosRealizados.Show
End Sub

Private Sub PagosEfectuadosRealizar_Click()
    FrmPagosTotal.Show
End Sub

Private Sub PlanActividades_Click()
    FrmPOA.Show
End Sub

Private Sub PorColaImpresion_Click()
    FrmColaImpresion.Show
End Sub

Private Sub PorSeleccionComprobantes_Click()
    FrmImprimeComprobanteNuevo.Show
End Sub
Private Sub mnuPrivAcceso_Click()
    frmPrivAcceso.Show
End Sub

'Private Sub pp_Click()
'If NombreUsuario = "rag001" Or NombreUsuario = "RAG001" Or NombreUsuario = "MYB159" Then
''If NombreUsuario = "jcc001" Or NombreUsuario = "JCC001" Then
'    FrmCP.Show
'  Else
'    MsgBox "El usuario NO tiene Acceso !! ...", vbInformation + vbCritical, "Validación"
'  End If
'End Sub

Private Sub pp_Click()
    FrmCP.Show 'vbModal
End Sub

'Private Sub OrganismosFinanciadores_Click()
'   Dim f As New OrganismoFinanciador
'   f.Show
'End Sub
'
'Private Sub PartidasGasto_Click()
'   Dim f As New PartidasGasto
'   f.Show
'End Sub
'
'Private Sub PartidasGasto_Click()
'   Dim f As New PartidasGasto
'   f.Show
'End Sub

Private Sub Regularizacion_Click()
   FrmRegularizacion.Show
End Sub


Private Sub RecursosHumanos_Click()
'RRHH1.exe
  Dim e As Long
  e = Shell(App.Path & "\RRHH1\RRHH1.exe", 1)
End Sub

'Private Sub ReversionCompromiso_Click()
'   Dim f As New FrmCompromisoR
'   f.Show
'End Sub
'
'Private Sub ReversionDevengado_Click()
'   Dim f As New FrmDevengadoR
'   f.Show
'End Sub


'Private Sub ProvinciasDepartamentos_Click() 'clasifica
'    frmProv.Show
'End Sub

Private Sub RelacionadorCuentaPartidas_Click()
    FrmC_Relacionador.Show
End Sub

Private Sub RelacionadorDonacionesOrganismos_Click()
    frmRelDonOrg.Show
End Sub

Private Sub RelacionadorGastoEco_Click()
    frmRelGastoEcoPar.Show
End Sub

Private Sub RelacionadorIngresosCuentas_Click()
    frmRelingresoCta.Show
End Sub

Private Sub RelacionadorRubroEco_Click()
    frmRelRecEcoRec.Show
End Sub

Private Sub repGraforga_Click()
  glRepPresup = "repGraf002"
  frmRepPresGlobal.Show
End Sub

Private Sub repGrafUnidad_Click()
  glRepPresup = "repGraf001"
  frmRepPresGlobal.Show
End Sub

Private Sub ReportesIngresos_Click()
  Dim e As Long
  e = Shell(App.Path & "\FormsIngresos\ingresos\ProyRepIngresos.exe")
End Sub

Private Sub SaldosActuales_Click()
'  FrmSaldosReales.Show
  FrmSaldosBancarios.Show
End Sub

'Private Sub Salida_Click()
'   Unload Me
'   End
'End Sub

'Private Sub Sisin_Click()
'   Dim f As New Sisin
'   f.Show
'End Sub
'
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error Resume Next
   Select Case Button.Key
      Case "Cliente"
         frmBeneficiario.Show
      Case "Producto"
         AlFrmCreaMaterial.Show
      Case "Guardar"
         'TareasPendientes: Agregar código de botón 'Guardar'.
         MsgBox "Agregar código de botón 'Guardar'."
      Case "Imprimir"
         'TareasPendientes: Agregar código de botón 'Imprimir'.
         MsgBox "Agregar código de botón 'Imprimir'."
      Case "Cortar"
         'TareasPendientes: Agregar código de botón 'Cortar'.
         MsgBox "Agregar código de botón 'Cortar'."
      Case "Copiar"
         'TareasPendientes: Agregar código de botón 'Copiar'.
         MsgBox "Agregar código de botón 'Copiar'."
      Case "Pegar"
         'TareasPendientes: Agregar código de botón 'Pegar'.
         MsgBox "Agregar código de botón 'Pegar'."
      Case "Negrita"
         ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
         Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
      Case "Cursiva"
         ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
         Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
      Case "Subrayado"
         ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
         Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
      Case "Alinear a la izquierda"
         ActiveForm.rtfText.SelAlignment = rtfLeft
      Case "Centrar"
         ActiveForm.rtfText.SelAlignment = rtfCenter
      Case "Alinear a la derecha"
         ActiveForm.rtfText.SelAlignment = rtfRight
   End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error Resume Next
   Select Case Button.Key
      Case "Producto"
         AlFrmCreaMaterial.Show
      Case "Cliente"
         frmBeneficiario.Show
      Case "Solicita Compra"
         FrmF01.Show
      Case "Compra"
         frmComprasDirectas.Show
      Case "Pago"
         FrmPagosTotal.Show
      Case "Venta"
         FrmVentas.Show
      Case "ReporteV"
         frmVentasReportes.Show
      Case "ReporteC"
         'ActiveForm.rtfText.SelAlignment = rtfRight
   End Select
End Sub

'Private Sub TipoTramite_Click()
'   Dim f As New frmFormulario
'   f.Show
'End Sub
'
'Private Sub UnidadesEjecutoras_Click()
'   Dim f As New UnidadEjecutora
'   f.Show
'End Sub


'Private Sub TiposErrores_Click() 'clasifica
'    frmErrores.Show
'End Sub

'Private Sub TipoTramite_Click()
'    frmFormulario.Show
'End Sub
'
'Private Sub UnidadesEjecutoras_Click()
'    Unidad.Show
'End Sub

'*********************** Freddy ****************************
Public Sub NivelAcceso(vNivelAcceso As Integer)
'Subrutina que habilita o deshabilita las opciones de menu
On Error Resume Next
Dim vNombOpcMenu As String

rsNivelAcceso.Open "Select * From NivelAcceso Where IdNivelAcceso=" & vNivelAcceso, db, adOpenStatic
If rsNivelAcceso.RecordCount > 0 Then
    While Not rsNivelAcceso.EOF
    vNombOpcMenu = LCase(rsNivelAcceso!NombOpcMenu)
    If vNombOpcMenu = "clasificadores" Then Clasificadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "generales" Then generales.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)  'Presupuesto
                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "departamentosbolivia" Then DepartamentosBolivia.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "provinciasdepartamentos" Then ProvinciasDepartamentos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "tiposerrores" Then TiposErrores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)  'Presupuesto
                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "economicosgasto" Then economicosgasto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "RelacionadorGastoEco" Then RelacionadorGastoEco.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "presupuesto" Then Presupuesto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)   'Presupuesto
                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "sisin" Then Sisin.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "deducciones" Then Deducciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadordonacionesorganismos" Then RelacionadorDonacionesOrganismos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'             If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)  'Presupuesto
'                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "distribucionporcentual" Then DistribucionPorcentual.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "ctesoreria" Then CTesoreria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "cbancos" Then CBancos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "cctabancarias" Then CCtaBancarias.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formaspago" Then FormasPago.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "ingresos" Then Ingresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "rubros" Then Rubros.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadorrubroeco" Then RelacionadorRubroEco.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "contabilidad2" Then Contabilidad2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "plancuentas" Then PlanCuentas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadorcuentapartidas" Then RelacionadorCuentaPartidas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadoringresoscuentas" Then RelacionadorIngresosCuentas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "depreciaciones" Then Depreciaciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "claseauxiliares" Then ClaseAuxiliares.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "inversiones" Then Inversiones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "administrativos" Then Administrativos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "adquisiciones" Then Adquisiciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "contrataciones" Then Contrataciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "almacenes2" Then Almacenes2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "recursoshumanos2" Then RecursosHumanos2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "mesaentrada" Then MesaEntrada.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)      'Contabilidad
            If vNombOpcMenu = "registrosolicitudes" Then RegistroSolicitudes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof01" Then FormularioF01.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof02" Then FormularioF02.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof03" Then FormularioF03.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof04" Then FormularioF04.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof05" Then FormularioF05.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof06" Then FormularioF06.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof07" Then FormularioF07.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "copiarestauracion" Then CopiaRestauracion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "progconadq" Then ProgConAdq.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
    
    If vNombOpcMenu = "procesos" Then Procesos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)          'Egresos
        If vNombOpcMenu = "compromiso" Then Compromiso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Ejecucion Presupuestaria
        If vNombOpcMenu = "ejecucion" Then Ejecucion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Reportes de Ejecucion
            '''ALB If vNombOpcMenu = "ejecucionppto" Then EjecucionPpto.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
            If vNombOpcMenu = "mnuejecucionpoa" Then mnuEjecucionPOA.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            '''ALB  If vNombOpcMenu = "repgraf" Then RepGraf.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
            '''ALB  If vNombOpcMenu = "repgrafunidad" Then repGrafUnidad.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
                If vNombOpcMenu = "repgraforga" Then repGraforga.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            '''ALB If vNombOpcMenu = "mnuejecucionpresupuestaria" Then mnuEjecucionPresupuestaria.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
                If vNombOpcMenu = "ejecucionporuni" Then EjecucionPorUni.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "mnuconvenio" Then mnuConvenio.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "mnuorganismo" Then mnuOrganismo.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "mnucategoria" Then mnuCategoria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "mnuproyecto" Then mnuProyecto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        
'        If vNombOpcMenu = "ejecucion" Then Ejecucion.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)     'Reportes de Ejecucion
'            If vNombOpcMenu = "ejecucionorganismo" Then EjecucionOrganismo.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            If vNombOpcMenu = "ejecucioncomprobante" Then EjecucionComprobante.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            If vNombOpcMenu = "ejecucioncompromiso" Then EjecucionCompromiso.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            If vNombOpcMenu = "ejecucióndevengado" Then EjecuciónDevengado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            If vNombOpcMenu = "ejecuciónpagado" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
        If vNombOpcMenu = "mnumodppto" Then MnuModPpto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Modificaciones Presupuestarias
        
    If vNombOpcMenu = "tesoreria" Then Tesoreria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        If vNombOpcMenu = "pp" Then pp.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Pagos pendientes
        If vNombOpcMenu = "pagosefectuados" Then PagosEfectuados.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Activacion de Cheques
        If vNombOpcMenu = "cuentasbancarias2" Then CuentasBancarias2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Desactivacion de cheques
        If vNombOpcMenu = "manejocheques" Then ManejoCheques.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Trasnferencia
        If vNombOpcMenu = "ic" Then ic.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
        If vNombOpcMenu = "ct" Then ct.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
            If vNombOpcMenu = "mnuimpcheques" Then MnuImpCheques.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Cheques
            If vNombOpcMenu = "cuentasbancarias" Then CuentasBancarias.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Cheques
        
    If vNombOpcMenu = "ccontabilidad" Then CContabilidad.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
            If vNombOpcMenu = "comprobantes" Then Comprobantes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
            If vNombOpcMenu = "reportesc" Then ReportesC.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "libromayor" Then LibroMayor.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "libromayorauxiliar" Then LibroMayorAuxiliar.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "balancegeneral" Then BalanceGeneral.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "balancesumassaldos" Then BalanceSumasSaldos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                If vNombOpcMenu = "estadoresultados" Then EstadoResultados.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
                
    If vNombOpcMenu = "mnuingresos" Then MnuIngresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)          'Egresos
        If vNombOpcMenu = "ejecucionpresupuestaria" Then EjecucionPresupuestaria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Ejecucion Presupuestaria
        If vNombOpcMenu = "reportesingresos" Then ReportesIngresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Reportes de Ejecucion
    
''''    If vNombOpcMenu = "administracion" Then Administracion.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''        If vNombOpcMenu = "adquisicionbienes" Then AdquisicionBienes.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''                If vNombOpcMenu = "comprasdirectas" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''                If vNombOpcMenu = "licitacionesnacionales" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''                If vNombOpcMenu = "licitacionesinternacionales" Then LicitacionesInternacionales.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''        If vNombOpcMenu = "contratacionservicios" Then ContratacionServicios.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''                If vNombOpcMenu = "consultoresindividuales" Then ConsultoresIndividuales.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''                If vNombOpcMenu = "empresasconsultoras" Then EmpresasConsultoras.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''        If vNombOpcMenu = "almacenes" Then Almacenes.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''                If vNombOpcMenu = "ingresosa" Then IngresosA.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''''                If vNombOpcMenu = "salidasa" Then SalidasA.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
    
    If vNombOpcMenu = "administracion" Then Administracion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        If vNombOpcMenu = "mnuadmicont" Then MnuAdmiCont.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        
        If vNombOpcMenu = "almacenes" Then Almacenes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "ingresosa" Then IngresosA.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
              If vNombOpcMenu = "mnuingreso" Then mnuIngreso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
              If vNombOpcMenu = "mnucompras" Then MnuCompras.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnusalidas" Then mnuSalidas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnucontrol" Then mnuControl.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
              If vNombOpcMenu = "mnuinventario" Then mnuInventario.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
              If vNombOpcMenu = "mnuestado" Then mnuestado.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
              
        If vNombOpcMenu = "mnulicitac" Then mnuLicitac.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnusolnoobjec" Then MnuSolNoObjec.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnupublicacion" Then mnuPublicacion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnucomision" Then mnucomision.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuventapliegos" Then mnuVentaPliegos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnurecepsobres" Then MnuRecepSobres.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuabogado" Then mnuabogado.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuadjudica" Then mnuAdjudica.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuorden" Then mnuorden.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        
        If vNombOpcMenu = "mnuconsultoria_c" Then mnuConsultoria_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnusolnoobj_c" Then mnuSolNoObj_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnupublicacion_c" Then mnuPublicacion_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuventapliegos_c" Then mnuVentaPliegos_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnurecepcionprop_c" Then mnuRecepcionProp_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuaperturaprop_c" Then mnuAperturaProp_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuadjudicacion_c" Then mnuAdjudicacion_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnugespagos" Then mnugesPagos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "mnuoc" Then mnuOC.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        
        If vNombOpcMenu = "comprasdirectas" Then ComprasDirectas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "solicitudnoobjecioncd" Then solicitudnoobjecioncd.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "cotizaciones" Then Cotizaciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "cuadrocomparativo" Then cuadrocomparativo.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
          If vNombOpcMenu = "ordenpagocd" Then OrdenPagoCD.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        
        'If vNombOpcMenu = "progconadq" Then ProgConAdq.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        If vNombOpcMenu = "mnusolicitudesf04" Then mnuSolicitudesF04.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)

    If vNombOpcMenu = "recursoshumanos" Then RecursosHumanos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
        If vNombOpcMenu = "administracionpersonal" Then AdministracionPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        If vNombOpcMenu = "controlpersonal" Then ControlPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        If vNombOpcMenu = "capacitacionpersonal" Then CapacitacionPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
        If vNombOpcMenu = "evaluaciondesempeño" Then EvaluacionDesempeño.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
    
    If vNombOpcMenu = "informaciongerencial" Then InformacionGerencial.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad

    If vNombOpcMenu = "mnuadmisistema" Then mnuAdmiSistema.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Administracion del sistema
        If vNombOpcMenu = "mnucambiarclave" Then mnuCambiarClave.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Cambiar clave
        If vNombOpcMenu = "mnuusuarios" Then mnuUsuarios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Definicion de Usuarios
        If vNombOpcMenu = "mnunivelacceso" Then mnuNivelAcceso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Cambiar clave
        If vNombOpcMenu = "mnuprivacceso" Then mnuPrivAcceso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Privilegios de Operación
    rsNivelAcceso.MoveNext
    Wend
    rsNivelAcceso.MoveFirst
End If
End Sub

'Private Sub TipoTramite_Click() 'clasifica
'  frmFormulario.Show
'End Sub

'Private Sub UnidadesEjecutoras_Click() 'clasifica
''  Dim e As New Unidad
''  e.Show
'  Unidad.Show
'End Sub
Private Sub SOES_Click()
    Dim e As Long
    e = Shell(App.Path & "\FormsIngresos\soes\SOES.exe " & GlUsuario & " ABM_SOES", 1)
End Sub

Private Sub SolicitudNoObjecionCD_Click()
  frmNoObjecionD.Show
End Sub


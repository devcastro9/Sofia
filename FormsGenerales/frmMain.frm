VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sistema de Organización Financiera y Administrativa (SOFIA)"
   ClientHeight    =   8010
   ClientLeft      =   885
   ClientTop       =   735
   ClientWidth     =   14715
   Icon            =   "frmMain.frx":0000
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox frm_alertas 
      Align           =   1  'Align Top
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   14655
      TabIndex        =   27
      Top             =   525
      Visible         =   0   'False
      Width           =   14715
      Begin VB.Frame Frame1 
         Caption         =   "ALERTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1815
         Left            =   5880
         TabIndex        =   28
         Top             =   120
         Width           =   6855
         Begin VB.OptionButton Option3 
            Caption         =   "Cerrar sin Elegir ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   31
            Top             =   1200
            Width           =   2655
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Alertas para Mantenimientos Gratuitos (Actas de Entrega Definitiva)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   30
            Top             =   720
            Width           =   6375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Alertas para Contratos de Ventas (Vencidos o por Vencer)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   29
            Top             =   240
            Width           =   5535
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   4440
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   14715
      TabIndex        =   10
      Top             =   0
      Width           =   14715
      Begin VB.CommandButton CmdRepA 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Contraseña"
         Height          =   520
         Left            =   17280
         Picture         =   "frmMain.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cambiar Contraseña"
         Top             =   20
         Width           =   1575
      End
      Begin VB.CommandButton CmdPagos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Financiero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   16080
         Picture         =   "frmMain.frx":1404
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Reportes Financieros/Contables"
         Top             =   20
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CommandButton CmdAlmacen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Almacenes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   14640
         Picture         =   "frmMain.frx":1E06
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Inventario de Almacenes"
         Top             =   20
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CommandButton CmdRepC 
         BackColor       =   &H80000004&
         Caption         =   "Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   13200
         Picture         =   "frmMain.frx":2808
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Reportes Compras y Pagos"
         Top             =   20
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CommandButton CmdTesoreria 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rep.Tesorería"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   11760
         Picture         =   "frmMain.frx":2D92
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Reportes de Tesorería"
         Top             =   15
         Width           =   1470
      End
      Begin VB.CommandButton CmdProd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BB.ySS."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   4560
         Picture         =   "frmMain.frx":3794
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Registro de Bienes (Repuestos, Herramientas, etc.) y Servicios"
         Top             =   20
         Width           =   1170
      End
      Begin VB.CommandButton CmdEqp 
         BackColor       =   &H80000004&
         Caption         =   "Equipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   3420
         Picture         =   "frmMain.frx":4196
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Registro de Equipos (Ascensores, Escaleras y Otros Similares)"
         Top             =   20
         Width           =   1170
      End
      Begin VB.CommandButton CmdEdif 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Edificios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   2280
         Picture         =   "frmMain.frx":4B98
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Registro de Edificios"
         Top             =   20
         Width           =   1170
      End
      Begin VB.CommandButton cmd_alertas 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ALERTAS!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   5790
         Picture         =   "frmMain.frx":559A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Alertas! Contratos con Observaciones"
         Top             =   20
         Width           =   1470
      End
      Begin VB.CommandButton CmdEmp 
         BackColor       =   &H80000004&
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   1145
         Picture         =   "frmMain.frx":5F9C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Registro de Empresas"
         Top             =   20
         Width           =   1170
      End
      Begin VB.CommandButton CmdRepV 
         BackColor       =   &H80000004&
         Caption         =   "Rep.Cobranzas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   10320
         Picture         =   "frmMain.frx":6326
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Reportes de Cobranzas Facturadas"
         Top             =   15
         Width           =   1470
      End
      Begin VB.CommandButton CmdVenta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contratos.Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   8760
         Picture         =   "frmMain.frx":6D28
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Administracion de Contratos"
         Top             =   15
         Width           =   1590
      End
      Begin VB.CommandButton CmdRepGral 
         Caption         =   "Rep.Gerencial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   7320
         Picture         =   "frmMain.frx":72B2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Reportes Gerenciales"
         Top             =   20
         Width           =   1470
      End
      Begin VB.CommandButton CmdBenef 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Personas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   530
         Left            =   0
         Picture         =   "frmMain.frx":7CB4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Registro de Personas"
         Top             =   10
         Width           =   1170
      End
      Begin VB.CommandButton CmdPedido 
         Caption         =   "Factura.Elect."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   6600
         Picture         =   "frmMain.frx":86B6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Subir Facturas Electrónicas"
         Top             =   20
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton CmdAsist 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   18840
         Picture         =   "frmMain.frx":8C40
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Nuevo Registro"
         Top             =   30
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.PictureBox sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   7740
      Width           =   14715
      Begin VB.Label txtVersion 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   14040
         TabIndex        =   26
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblVersion 
         Caption         =   "Versión:"
         Height          =   255
         Left            =   13320
         TabIndex        =   32
         Top             =   0
         Width           =   735
      End
      Begin VB.Label txtHoraGl 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9600
         TabIndex        =   25
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblHoraGl 
         Caption         =   "Hora:"
         Height          =   255
         Left            =   8760
         TabIndex        =   24
         Top             =   0
         Width           =   495
      End
      Begin VB.Label txtFechaGl 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         TabIndex        =   23
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblFechaGl 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblBDGl 
         Caption         =   "Base de Datos:"
         Height          =   255
         Left            =   17040
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label txtBDGl 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   18360
         TabIndex        =   20
         Top             =   0
         Width           =   960
      End
      Begin VB.Label txtUsuarioGl 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   19
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblUsuarioGL 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport Cry 
      Left            =   0
      Top             =   9960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR07 
      Left            =   480
      Top             =   9960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Menu Mnu_Planificacion 
      Caption         =   "PLANIFICACION"
      Begin VB.Menu Mnu_ClasificadoresGral 
         Caption         =   "CLASIFICADORES DE USO GENERAL"
         Begin VB.Menu MnuClasificacionBeneficiarios 
            Caption         =   "CLASIFICACION DE BENEFICIARIOS"
         End
         Begin VB.Menu Mnu_TipoVivienda 
            Caption         =   "TIPOS DE VIVIENDAS"
         End
         Begin VB.Menu Mnu_TipoViasAcceso 
            Caption         =   "TIPOS DE VIAS DE ACCESO"
         End
         Begin VB.Menu mnu_clasificacionTramites 
            Caption         =   "CLASIFICACION DE TRAMITES"
         End
         Begin VB.Menu Mnu_ClasificacionDocumentos 
            Caption         =   "CLASIFICACION DE DOCUMENTOS"
         End
         Begin VB.Menu Div01 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_Paises 
            Caption         =   "PAISES"
         End
         Begin VB.Menu Mnu_Departamentos 
            Caption         =   "DEPARTAMENTOS DEL PAIS"
         End
         Begin VB.Menu Mnu_Provincias 
            Caption         =   "PROVINCIAS"
         End
         Begin VB.Menu Mnu_Municipios 
            Caption         =   "MUNICIPIOS"
         End
         Begin VB.Menu Mnu_Comunidades 
            Caption         =   "COMUNIDADES / LOCALIDADES"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_zonas 
            Caption         =   "ZONAS / BARRIOS"
         End
         Begin VB.Menu Mnu_ViasAcceso 
            Caption         =   "VIAS DE ACCESO (Calle, Av, etc.)"
         End
         Begin VB.Menu Div02 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_DocumentosRespaldo 
            Caption         =   "DOCUMENTOS DE RESPALDO (ISO)"
         End
         Begin VB.Menu Mnu_TiposImpuestos 
            Caption         =   "TIPOS DE IMPUESTOS"
         End
      End
      Begin VB.Menu MnuProcedimientos 
         Caption         =   "PROCEDIMIENTOS ADMINISTRATIVOS (ISO)"
         Begin VB.Menu MunClasificacionGeneral 
            Caption         =   "PROCESOS (CLASIFICACION NIVEL 1)"
         End
         Begin VB.Menu MnuClasificacionEspecifica 
            Caption         =   "SUBPROCESOS (CLASIFICACION NIVEL 2)"
         End
         Begin VB.Menu MnuEtapas 
            Caption         =   "ETAPAS (CLASIFICACION NIVEL 3)"
         End
      End
      Begin VB.Menu MnuModPpto 
         Caption         =   "PRESUPUESTO"
         Enabled         =   0   'False
         Begin VB.Menu MnuFormulacionPresupuestaria 
            Caption         =   "FORMULACION PRESUPUESTARIA"
         End
         Begin VB.Menu MnuModificacionesPresupuestarias 
            Caption         =   "MODIFICACIONES PRESUPUESTARIAS"
         End
      End
   End
   Begin VB.Menu MnuProcesosAdministrativos 
      Caption         =   "COMERCIAL"
      Begin VB.Menu MnuClasificadoresAdministrativos 
         Caption         =   "CLASIFICADORES ADMINISTRATIVOS"
         Begin VB.Menu Mnu_EdificiosInstalacion 
            Caption         =   "ORGANIZACION EDIFICIOS EN INSTALACION"
         End
         Begin VB.Menu Mnu_TareasCronoInstalacion 
            Caption         =   "TAREAS CRONOGRAMA INSTALACION"
         End
         Begin VB.Menu DivA02 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_ModalidadesContratacion 
            Caption         =   "MODALIDADES DE CONTRATACION"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_Componentes_Equipos 
            Caption         =   "CARACTERISTICAS EQUIPOS Y OTROS BIENES"
         End
         Begin VB.Menu DivA03 
            Caption         =   "-"
         End
         Begin VB.Menu MnuGruposBienes 
            Caption         =   "GRUPOS DE BIENES Y SERVICIOS"
         End
         Begin VB.Menu MnuSubgrupoBienes 
            Caption         =   "SUBGRUPO DE BIENES Y SERVICIOS"
         End
         Begin VB.Menu MnuBienesServicios 
            Caption         =   "BIENES Y SERVICIOS"
         End
         Begin VB.Menu MnuUnidadesMedida 
            Caption         =   "UNIDADES DE MEDIDA"
         End
         Begin VB.Menu DivA04 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_CostosComercializacion 
            Caption         =   "COSTOS DE COMERCIALIZACION"
         End
      End
      Begin VB.Menu MnuComercialComex 
         Caption         =   "VENTAS NUEVAS"
         Begin VB.Menu MnuIdentificacionCliente 
            Caption         =   "IDENTIFICACION DEL CLIENTE VENTA NUEVA"
         End
         Begin VB.Menu Mnu_ParametrosCalculo 
            Caption         =   "PARAMETROS DE CALCULO"
         End
         Begin VB.Menu MnuCotizacionesEquipos 
            Caption         =   "COTIZACION DE EQUIPOS VENTA NUEVA"
         End
         Begin VB.Menu Mnu_ProcesoVentas 
            Caption         =   "CONTRATO VENTAS NUEVAS"
         End
         Begin VB.Menu Mnu_ProcesoCompras 
            Caption         =   "SEGUIMIENTO COMERCIAL"
         End
         Begin VB.Menu Mnu_Compras 
            Caption         =   "COMPRAS Y CONTRATACIONES"
            Visible         =   0   'False
            Begin VB.Menu Mnu_SolicitudCompra 
               Caption         =   "COMPRA SERVICIOS INSTALACION"
            End
            Begin VB.Menu Mnu_AdjudicacionCompra 
               Caption         =   "ADJUDICACION COMPRA"
               Visible         =   0   'False
            End
            Begin VB.Menu Mnu_OrdenPago 
               Caption         =   "ORDENES DE PAGO"
               Visible         =   0   'False
            End
         End
      End
      Begin VB.Menu Mnu_ImportacionEquipos 
         Caption         =   "COMEX (IMPORTACION DE EQUIPOS)"
         Begin VB.Menu Mnu_ProveedoresEquipos 
            Caption         =   "PROVISION DE EQUIPOS / REPUESTOS"
         End
         Begin VB.Menu Mnu_Transporte 
            Caption         =   "TRANSPORTE"
         End
         Begin VB.Menu Mnu_Nacionalizacion 
            Caption         =   "NACIONALIZACION"
         End
         Begin VB.Menu Mnu_Descarguio 
            Caption         =   "DESCARGUIO"
         End
         Begin VB.Menu Mnu_SeguimientoPago 
            Caption         =   "SEGUIMIENTO COMEX"
         End
      End
      Begin VB.Menu mnu_instalaciones 
         Caption         =   "INSTALACIONES"
         Begin VB.Menu Mnu_CronogramaInstalaciones 
            Caption         =   "ELABORACION CRONOGRAMA INSTALACIONES"
         End
         Begin VB.Menu Mnu_ContratacionTecnicos 
            Caption         =   "CONTRATACION TECNICOS"
         End
         Begin VB.Menu Mnu_IdentificacionClienteInstalacion 
            Caption         =   "IDENTIFICACION DEL CLIENTE INSTALACIONES"
         End
         Begin VB.Menu Mnu_ProcesoInstalaciones 
            Caption         =   "VENTA SERVICIO DE INSTALACIONES"
         End
         Begin VB.Menu Mnu_EjecucionInstalaciones 
            Caption         =   "EJECUCION CRONOGRAMA INSTALACIONES"
         End
         Begin VB.Menu Mnu_SeguimientoInstalaciones 
            Caption         =   "ACTA DE ENTREGA DEFINITIVA"
         End
         Begin VB.Menu Mnu_BitacoraInstalaciones 
            Caption         =   "BITACORA DE INSTALACIONES"
         End
      End
      Begin VB.Menu Mnu_AdminitracionContratos 
         Caption         =   "ADMINISTRACION DE CONTRATOS"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ActivosFijos 
         Caption         =   "ACTIVOS FIJOS"
         Visible         =   0   'False
         Begin VB.Menu mnu_AsignacionActivos 
            Caption         =   "ASIGNACION ACTIVOS FIJOS"
         End
         Begin VB.Menu mnu_TransferenciasActivos 
            Caption         =   "TRANSFERENCIAS DE ACTIVOS"
         End
      End
   End
   Begin VB.Menu mnu_gerenciaTecnica 
      Caption         =   "TECNICO"
      Begin VB.Menu Mnu_ClasificadoresAreaTecnica 
         Caption         =   "CLASIFICADORES AREA TECNICA"
         Begin VB.Menu Mnu_Definicion_zonas 
            Caption         =   "ORGANIZACION DE ZONAS PILOTO"
         End
         Begin VB.Menu Mnu_CalendarioZonas 
            Caption         =   "CALENDARIO POR ZONAS"
            Enabled         =   0   'False
         End
         Begin VB.Menu MnuAlmacenesFisicos 
            Caption         =   "ALMACENES FISICOS"
         End
      End
      Begin VB.Menu mnu_mantenimiento 
         Caption         =   "MANTENIMIENTO"
         Begin VB.Menu Mnu_IdentificacionClienteMantenimiento 
            Caption         =   "IDENTIFICACION DEL CLIENTE MANTENIMIENTO"
         End
         Begin VB.Menu Mnu_ProcesoMantenimiento 
            Caption         =   "VENTA SERVICIO DE MANTENIMIENTO"
         End
         Begin VB.Menu Mnu_CronogramaMantenimiento 
            Caption         =   "CRONOGRAMA POR CONTRATO DE MANTENIMIENTO"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_AsignacionTareasM 
            Caption         =   "CRONOGRAMA MENSUAL DE MANTENIMIENTO"
         End
         Begin VB.Menu Mnu_Ejecucion_Servicio_M 
            Caption         =   "EJECUCION DEL SERVICIO (Certificados)"
         End
         Begin VB.Menu Mnu_SeguimientoMantenimiento 
            Caption         =   "BITACORA DE MANTENIMIENTO"
         End
      End
      Begin VB.Menu mnu_reparaciones 
         Caption         =   "REPARACIONES"
         Begin VB.Menu Mnu_IdentificacionClienteReparacion 
            Caption         =   "IDENTIFICACION DEL CLIENTE REPARACIONES"
         End
         Begin VB.Menu Mnu_ProcesoReparacion 
            Caption         =   "VENTA SERVICIO DE REPARACIONES"
         End
         Begin VB.Menu Mnu_CronogramaReparacion 
            Caption         =   "CRONOGRAMA MENSUAL DE REPARACIONES"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_SeguimientoReparacion 
            Caption         =   "BITACORA DE REPARACIONES"
         End
      End
      Begin VB.Menu mnu_emergencias 
         Caption         =   "EMERGENCIAS"
         Begin VB.Menu Mnu_IdentificacionClienteEmergencia 
            Caption         =   "IDENTIFICACION DEL CLIENTE EMERGENCIAS"
         End
         Begin VB.Menu Mnu_ProcesoEmergencia 
            Caption         =   "VENTA DE SERVICIOS POR EMERGENCIAS"
         End
         Begin VB.Menu Mnu_CronogramaEmergencia 
            Caption         =   "CRONOGRAMA POR SERVICIO DE EMERGENCIAS"
         End
         Begin VB.Menu Mnu_AsignacionTareasE 
            Caption         =   "REGISTRO DE TAREAS DE EMERGENCIAS"
         End
         Begin VB.Menu Mnu_SeguimientoEmergencia 
            Caption         =   "BITACORA DE EMERGENCIAS"
         End
      End
      Begin VB.Menu mnu_modernizacion 
         Caption         =   "MODERNIZACION"
         Begin VB.Menu Mnu_IdentificacionClienteModernizacion 
            Caption         =   "IDENTIFICACION DEL CLIENTE MODERNIZACION"
         End
         Begin VB.Menu Mnu_AsignacionTareasD 
            Caption         =   "PARAMETROS DE CALCULO"
         End
         Begin VB.Menu Mnu_CronogramaModernizacion 
            Caption         =   "COTIZACION DE MODERNIZACION"
         End
         Begin VB.Menu Mnu_ProcesoModernizacion 
            Caption         =   "VENTAS MODERNIZACIONES"
         End
         Begin VB.Menu Mnu_SeguimientoModernizacion 
            Caption         =   "BITACORA DE MODERNIZACION"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Mnu_AlmacenInsumos 
         Caption         =   "ALMACEN DE INSUMOS Y MATERIALES"
         Begin VB.Menu Mnu_solicitudInsumos 
            Caption         =   "SOLICITUD DE INSUMOS Y MATERIALES"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_IngresosAlmacen 
            Caption         =   "INGRESOS ALMACEN INSUMOS (Pago Posterior)"
         End
         Begin VB.Menu mnu_IngresosAlmacen2 
            Caption         =   "INGRESOS ALMACEN INSUMOS (Pago Anticipado)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_SalidaAlmacen 
            Caption         =   "SALIDA ALMACEN INSUMOS (MANTENIMIENTO)"
         End
         Begin VB.Menu mnu_SalidaAlmacenOtro 
            Caption         =   "SALIDA/TRANSFERENCIA ALMACEN INSUMOS"
         End
         Begin VB.Menu mnu_InventarioAlmacen 
            Caption         =   "INVENTARIOS INSUMOS"
         End
      End
      Begin VB.Menu Mnu_AlmacenRepuestos 
         Caption         =   "ALMACEN DE REPUESTOS"
         Begin VB.Menu Mnu_solicitudRepuestos 
            Caption         =   "SOLICITUD DE REPUESTOS"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_IngresosAlmacenRep 
            Caption         =   "INGRESOS ALMACEN REPUESTOS (Pago Posterior)"
         End
         Begin VB.Menu mnu_IngresosAlmacenRep2 
            Caption         =   "INGRESOS ALMACEN REPUESTOS (Pago Anticipado)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_SalidaAlmacenRep 
            Caption         =   "SALIDA/TRANSFERENCIA ALMACEN REPUESTOS"
         End
         Begin VB.Menu mnu_InventarioAlmacenRep 
            Caption         =   "INVENTARIOS REPUESTOS"
         End
      End
      Begin VB.Menu Mnu_AlmacenHerramientas 
         Caption         =   "ALMACEN DE HERRAMIENTAS"
         Begin VB.Menu Mnu_SolicitudHerramientas 
            Caption         =   "SOLICITUD DE HERRAMIENTAS"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_IngresosAlmacenHerr 
            Caption         =   "INGRESOS ALMACEN HERRAMIENTAS (Pago Posterior)"
         End
         Begin VB.Menu mnu_IngresosAlmacenHerr2 
            Caption         =   "INGRESOS ALMACEN HERRAMIENTAS (Pago Anticipado)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_SalidaAlmacenHerr 
            Caption         =   "SALIDA/TRANSFERENCIA ALMACEN HERRAMIENTAS"
         End
         Begin VB.Menu mnu_InventarioAlmacenHerr 
            Caption         =   "INVENTARIOS HERRAMIENTAS"
         End
      End
      Begin VB.Menu Mnu_RecibosOficialesEgresos2 
         Caption         =   "ORDEN DE CANCELACION (EGRESOS)"
      End
      Begin VB.Menu Mnu_saldosinicialesalmacenes 
         Caption         =   "SALDOS INICIALES ALMACENES"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_RRHH 
      Caption         =   "RECURSOS_HUMANOS"
      Begin VB.Menu Mnu_ClasificadoresRRHH 
         Caption         =   "CLASIFICADORES RRHH"
         Begin VB.Menu mnu_GerenciaGeneral 
            Caption         =   "GERENCIA GENERAL"
         End
         Begin VB.Menu mnu_GerenciasOperativas 
            Caption         =   "GERENCIAS OPERATIVAS"
         End
         Begin VB.Menu Mnu_UnidadesEjecutoras 
            Caption         =   "UNIDADES EJECUTORAS"
         End
         Begin VB.Menu Mnu_CargosFuncionales 
            Caption         =   "CARGOS FUNCIONALES"
         End
         Begin VB.Menu Mnu_PuestosOrganizacionales 
            Caption         =   "PUESTOS ORGANIZACIONALES"
         End
         Begin VB.Menu Div1_RRHH 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_NivelesEducacion 
            Caption         =   "NIVELES DE EDUCACION"
         End
         Begin VB.Menu MnuProfesionesOcupaciones 
            Caption         =   "PROFESIONES / OCUPACIONES"
         End
         Begin VB.Menu mnuRgistroPersonas 
            Caption         =   "REGISTRO DE PERSONAL"
         End
         Begin VB.Menu Mnu_AsignaResponsableUnidad 
            Caption         =   "ASIGNA RESPONSABLE DE UNIDAD"
         End
         Begin VB.Menu Mnu_Parentesco 
            Caption         =   "PARENTESCO"
            Visible         =   0   'False
         End
         Begin VB.Menu Div3_RRHH 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_CalendarioLaboral 
            Caption         =   "CALENDARIO LABORAL"
         End
         Begin VB.Menu Mnu_HorariosLaborales 
            Caption         =   "HORARIOS LABORALES"
         End
         Begin VB.Menu MNU_ANTIGUEDAD 
            Caption         =   "PARAMETROS DE ANTIGUEDAD"
         End
         Begin VB.Menu Mnu_Motivos_Procesos 
            Caption         =   "MOTIVOS POR PROCESOS"
         End
      End
      Begin VB.Menu mnu_AdministracionPersonal 
         Caption         =   "PROCESOS RECURSOS HUMANOS"
         Begin VB.Menu menuPostulante 
            Caption         =   "POSTULANTES"
         End
         Begin VB.Menu mnu_ProcesoContratacion 
            Caption         =   "CONTRATACION DE PERSONAL"
         End
         Begin VB.Menu Mnu_ImportarRegistroAsistencia 
            Caption         =   "IMPORTAR REGISTRO ASISTENCIA"
         End
         Begin VB.Menu Mnu_FichaAdministracionPersonal 
            Caption         =   "FICHA DEL PERSONAL"
         End
         Begin VB.Menu Mnu_PrestamosAPersonal 
            Caption         =   "PRESTAMOS A PERSONAL"
         End
      End
      Begin VB.Menu mnu_ControlPersonal 
         Caption         =   "CONTROL DE PERSONAL"
         Visible         =   0   'False
         Begin VB.Menu Mnu_FileControlPersonal 
            Caption         =   "FILE CONTROL DE PERSONAL"
         End
      End
      Begin VB.Menu Mnu_PlanillasPagosPersonal 
         Caption         =   "PLANILLAS Y PAGOS PERSONAL"
         Begin VB.Menu Mnu_PlanillasGrupos 
            Caption         =   "DEFINICION DE PLANILLAS (GRUPOS )"
         End
         Begin VB.Menu Mnu_PlanillasSubGrupos 
            Caption         =   "DEFINICION DE SUBPLANILLAS (SUBGRUPOS)"
         End
         Begin VB.Menu Mnu_ProcesoPlanillas 
            Caption         =   "PROCESO DE PLANILLAS"
         End
         Begin VB.Menu menuBoletaPagos 
            Caption         =   "DEFINICION DE BOLETAS DE PAGO"
         End
      End
      Begin VB.Menu Mnu_Reportes_RRHH 
         Caption         =   "REPORTES RRHH"
      End
   End
   Begin VB.Menu Mnu_ProcesosFinancierso 
      Caption         =   "FINANCIERO"
      Begin VB.Menu mnu_ClasificadoresFinancieros 
         Caption         =   "CLASIFICADORES FINANCIEROS"
         Begin VB.Menu mnu_FuentesFinanciamiento 
            Caption         =   "FUENTES DE FINANCIAMIENTO"
         End
         Begin VB.Menu mnu_Financiadores 
            Caption         =   "FINANCIADORES"
         End
         Begin VB.Menu mnu_ProgramasProyectos 
            Caption         =   "PROYECTOS O ACTIVIDADES"
         End
         Begin VB.Menu mnu_GruposIngresos 
            Caption         =   "GRUPOS DE INGRESOS"
         End
         Begin VB.Menu mnu_RubrosIngresos 
            Caption         =   "RUBROS DE INGRESOS"
         End
         Begin VB.Menu mnu_PartidasGasto 
            Caption         =   "PARTIDAS POR OBJETO DEL GASTOS"
         End
         Begin VB.Menu Mnu_PlanCuentas 
            Caption         =   "PLAN DE CUENTAS"
         End
         Begin VB.Menu Mnu_RelacionadorIngresos 
            Caption         =   "RELACIONADOR INGRESOS"
         End
         Begin VB.Menu Mnu_Relacionador_Gastos 
            Caption         =   "RELACIONADOR GASTOS"
         End
         Begin VB.Menu Div_CF2 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_dosificacion_facturas 
            Caption         =   "DOSIFICACION DE FACTURAS"
         End
         Begin VB.Menu Mnu_Bancos 
            Caption         =   "BANCOS Y ENTIDADES FINANCIERAS"
         End
         Begin VB.Menu Mnu_CuentasBancarias 
            Caption         =   "CUENTAS BANCARIAS"
         End
      End
      Begin VB.Menu Mnu_EjecucionIngresos 
         Caption         =   "EJECUCION DE INGRESOS"
         Visible         =   0   'False
         Begin VB.Menu mnu_registroIngresos 
            Caption         =   "REGISTROS DE INGRESOS"
         End
      End
      Begin VB.Menu mnu_EjecucionGasto 
         Caption         =   "EJECUCION DE EGRESOS"
         Begin VB.Menu Mnu_SolicitudEgresos 
            Caption         =   "SOLICITUD DE EGRESOS"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_RegistroPagos 
            Caption         =   "COMPRAS REGULARES"
         End
         Begin VB.Menu Mnu_ServiciosBasicos 
            Caption         =   "PAGOS DE SERVICIOS BASICOS"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_RegistroGastos 
            Caption         =   "EJECUCION PRESUPUESTARIA"
         End
      End
      Begin VB.Menu mnu_Contabilidad 
         Caption         =   "CONTABILIDAD"
         Enabled         =   0   'False
         Begin VB.Menu Mnu_RegistroDiario 
            Caption         =   "MODULO CONTABILIDAD"
         End
         Begin VB.Menu BalanceApertura 
            Caption         =   "BALANCE DE APERTURA"
         End
         Begin VB.Menu Mnu_MayorAuxiliar 
            Caption         =   "MAYOR AUXLIAR"
         End
         Begin VB.Menu Mnu_BalanceGeneral 
            Caption         =   "BALANCE GENERAL"
         End
      End
      Begin VB.Menu Mnu_FaturacionCobranza 
         Caption         =   "FACTURACION Y COBRANZA"
         Begin VB.Menu Mnu_Facturacion 
            Caption         =   "FACTURACION SOFIA NUEVO"
         End
         Begin VB.Menu Mnu_FacturacionAntes 
            Caption         =   "FACTURAS ANTIGUAS"
         End
         Begin VB.Menu mnu_NotaCreditoDebito 
            Caption         =   "ORDEN DE COBRO"
         End
         Begin VB.Menu Mnu_Cobranzas 
            Caption         =   "REGISTRO DE COBRANZAS"
         End
         Begin VB.Menu Mnu_SeguimientoCobranzas 
            Caption         =   "SEGUIMIENTO COBRANZAS"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu_Tesoreria 
         Caption         =   "TESORERIA"
         Begin VB.Menu Mnu_RecibosOficiales 
            Caption         =   "RECIBOS OFICIALES INGRESOS"
         End
         Begin VB.Menu CuentasBancarias2 
            Caption         =   "TRASPASOS INGRESOS CUENTAS BANCARIAS"
         End
         Begin VB.Menu mnusepara10 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_RecibosOficialesEgresos 
            Caption         =   "ORDEN DE CANCELACION (EGRESOS)"
         End
         Begin VB.Menu Mnu_TraspasosEgresos 
            Caption         =   "TRASPASOS EGRESOS CUENTAS BANCARIAS"
         End
         Begin VB.Menu mnusepara11 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_SeguimientoCheques 
            Caption         =   "CONCILIACION BANCARIA"
         End
      End
      Begin VB.Menu Mnu_Descargos 
         Caption         =   "FONDOS A RENDIR"
         Begin VB.Menu Mnu_cargoCuenta 
            Caption         =   "CARGOS DE CUENTA"
         End
         Begin VB.Menu Mnu_DescargosFondosViajes 
            Caption         =   "FONDOS PARA VIAJES"
         End
         Begin VB.Menu Mnu_DecargosCajaChica 
            Caption         =   "CAJA CHICA"
         End
      End
      Begin VB.Menu Mnu_ReportesCobranzas 
         Caption         =   "REPORTES FINANCIEROS"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_InformacionGerencial 
      Caption         =   "INFORMACION_GERENCIAL"
      Visible         =   0   'False
      Begin VB.Menu mnuPoa 
         Caption         =   "PLANIFICACION"
      End
      Begin VB.Menu Mnu_ReportesAdministrativos 
         Caption         =   "ADMINISTRATIVOS"
      End
      Begin VB.Menu Mnu_ReportesFinancieros 
         Caption         =   "FINANCIEROS"
      End
      Begin VB.Menu Mnu_ReportesAreaTecnica 
         Caption         =   "AREA TECNICA"
      End
      Begin VB.Menu Mnu_ReportesRRHH 
         Caption         =   "RECURSOS HUMANOS"
      End
   End
   Begin VB.Menu mnu_AdministracionSistema 
      Caption         =   "ADMINISTRACION_SISTEMA"
      Begin VB.Menu mnu_AdministracionUsuarios 
         Caption         =   "ADMINISTRACION DE USUARIOS"
      End
      Begin VB.Menu mnu_ControlAccesos 
         Caption         =   "CONTROL DE ACCESOS"
      End
      Begin VB.Menu mnu_PrivilegiosOperacion 
         Caption         =   "CONTROL DE PRIVILEGIOS DE OPERACION"
      End
   End
   Begin VB.Menu mnu_Salida 
      Caption         =   "SALIDA"
      Begin VB.Menu mnuAcercade 
         Caption         =   "A cerca de ..."
      End
      Begin VB.Menu MnuAyuda 
         Caption         =   "Ayuda"
         HelpContextID   =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnusepara 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SalirSistema 
         Caption         =   "SALIR DEL SISTEMA"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------- Programas de JORGE
Dim rs_datos1 As New ADODB.Recordset

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Sub BalanceApertura_Click()
    fw_balance_de_apertura.lbl_titulo = BalanceApertura.Caption
    fw_balance_de_apertura.FraNavega = BalanceApertura.Caption
    fw_balance_de_apertura.lbl_titulo2 = BalanceApertura.Caption
    fw_balance_de_apertura.Show
    'fw_balance_de_apertura.Show
End Sub

Private Sub cmd_alertas_Click()
    frm_alertas.Visible = True
    
'    Cry.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_servicio_regional_Alerta.rpt"
'        titulo2 = "CONTRATOS DE VENTAS"
'        subtitulo2 = "VIGENCIA VENCIDA O A DIAS DE VENCER"
'        Cry.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        Cry.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        Cry.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
'        iResult = Cry.PrintReport
'        If iResult <> 0 Then
'            MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'    Cry.WindowState = crptMaximized
''    Timer1.Enabled = False
'
''    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "DOLMOS" Then
''        frm_rc_puestos.lbl_titulo = Mnu_PuestosOrganizacionales.Caption
''        frm_rc_puestos.FraNavega = Mnu_PuestosOrganizacionales.Caption
''        frm_rc_puestos.lbl_titulo2 = Mnu_PuestosOrganizacionales.Caption
''        frm_rc_puestos.Show
''    Else
''        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
''    End If
End Sub

Private Sub CmdBenef_Click()
    Glaux = "0"
    gw_p_gc_beneficiario_persona.lbl_titulo = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.FraNavega = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.lbl_titulo2 = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.Show
End Sub

Private Sub CmdEdif_Click()
    gw_edificaciones.lbl_titulo = "Edificaciones"
    gw_edificaciones.FraNavega = "Edificaciones"    'mnu_proyecto_edificacion.Caption
    gw_edificaciones.lbl_titulo2 = "Edificaciones"  'mnu_proyecto_edificacion.Caption
    gw_edificaciones.Show
End Sub

Private Sub CmdEmp_Click()
    gw_p_gc_beneficiario_empresa.lbl_titulo = "REGISTRO DE EMPRESAS"    'MnuRegistroEmpresas.Caption
    gw_p_gc_beneficiario_empresa.FraNavega = "REGISTRO DE EMPRESAS"    'MnuRegistroEmpresas.Caption
    gw_p_gc_beneficiario_empresa.lbl_titulo2 = "REGISTRO DE EMPRESAS"    'MnuRegistroEmpresas.Caption
    gw_p_gc_beneficiario_empresa.Show
End Sub

Private Sub CmdEqp_Click()
    frm_ac_bienes_eqp.lbl_titulo = "Equipos (Ascensores, Escaleras, etc.)"      'MnuEquipos.Caption
    frm_ac_bienes_eqp.FraNavega = "Equipos (Ascensores, Escaleras, etc.)"       'MnuEquipos.Caption
    frm_ac_bienes_eqp.lbl_titulo2 = "Equipos (Ascensores, Escaleras, etc.)"     'MnuEquipos.Caption
    frm_ac_bienes_eqp.Show
End Sub

Private Sub CmdPagos_Click()
    Form1.Show
End Sub

Private Sub CmdPedido_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "SQUISPE" Or glusuario = "RCUELA" Or glusuario = "CSALINAS" Then
        mw_importar_facturas_eletronicas.Show
    Else
        MsgBox "El usuario no tiene acceso !, Consulte con el Administrador del Sistema...", vbInformation + vbOKOnly
    End If
'    gw_edificaciones.lbl_titulo = mnu_proyecto_edificacion.Caption
'    gw_edificaciones.FraNavega = mnu_proyecto_edificacion.Caption
End Sub

Private Sub CmdProd_Click()
    aw_bienes.lbl_titulo = MnuBienesServicios.Caption
    aw_bienes.FraNavega = MnuBienesServicios.Caption
    aw_bienes.lbl_titulo2 = MnuBienesServicios.Caption
    aw_bienes.Show
End Sub

Private Sub CmdRepA_Click()
    FrmCambiarClave.Show
End Sub

Private Sub CmdRepC_Click()
    Fw_ReportesCompras.lbl_titulo = "REPORTES DE COMPRAS Y PAGOS"      'Mnu_ReportesCobranzas.Caption
    Fw_ReportesCompras.Show
End Sub

Private Sub CmdRepGral_Click()
    'gw_reportes_gerenciales.Show
    gw_rep_generales.Show
End Sub

Private Sub CmdRepV_Click()
    If glusuario = "JORAQUENI" Then
        MsgBox "El Usuario No tiene acceso, Consulte con el Administrador del Sistema ...", , "Atención"
        Exit Sub
    End If
    Fw_ReportesCobranzas.lbl_titulo = "REPORTES DE COBRANZAS"      'Mnu_ReportesCobranzas.Caption
    Fw_ReportesCobranzas.Show
'    rm_VentasReportes.lbl_titulo = "REPORTES DE VENTAS Y COBRANZAS"      'Mnu_ReportesCobranzas.Caption
'    frm_VentasReportes.Show
End Sub

Private Sub CmdTesoreria_Click()
    If glusuario = "JORAQUENI" Then
        MsgBox "El Usuario No tiene acceso, Consulte con el Administrador del Sistema ...", , "Atención"
        Exit Sub
    End If
    Fw_ReportesTesoreria.lbl_titulo = "REPORTES DE TESORERIA"
    Fw_ReportesTesoreria.Show
End Sub

Private Sub CmdVenta_Click()
    If glusuario = "JORAQUENI" Then
        MsgBox "El Usuario No tiene acceso, Consulte con el Administrador del Sistema ...", , "Atención"
        Exit Sub
    End If
    Aux = "DCOBR"
    'aw_seguimiento_contratos.lbl_titulo = Mnu_SeguimientoCobranzas.Caption
    fw_seguimiento_ventas.lbl_titulo = "ADMINISTRACION DE CONTRATOS"
    fw_seguimiento_ventas.Show
End Sub

Private Sub CuentasBancarias2_Click()
    Aux = "R-641"
    'fw_ventas_cobranzas.lbl_titulo = Mnu_Cobranzas.Caption
    fw_traspaso_bancos.FraNavega = CuentasBancarias2.Caption
    fw_traspaso_bancos.lbl_titulo = CuentasBancarias2.Caption
    fw_traspaso_bancos.Show
End Sub

Private Sub LblAlertaContratos_Click()
    'Timer1.Interval = 250
    'Timer1.Enabled = True
    Cry.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_servicio_regional_Alerta.rpt"
        titulo2 = "CONTRATOS DE VENTAS"
        subtitulo2 = "VIGENCIA VENCIDA O A DIAS DE VENCER"
        Cry.Formulas(2) = "Titulo = '" & titulo2 & "'"
        Cry.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        Cry.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = Cry.PrintReport
        If iResult <> 0 Then
            MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    Cry.WindowState = crptMaximized
'    Timer1.Enabled = False
End Sub

Private Sub menuBoletaPagos_Click()
    FrmBoletaPagos.Show
End Sub

Private Sub menuPostulante_Click()
    FrmPostulantes.Show
End Sub

Private Sub mnu_AdministracionUsuarios_Click()
'    FrmCambiarClave.lbl_titulo = mnu_AdministracionUsuarios.Caption
'    FrmCambiarClave.FraNavega = mnu_AdministracionUsuarios.Caption
'    FrmCambiarClave.lbl_titulo2 = mnu_AdministracionUsuarios.Caption
    FrmCambiarClave.Show
End Sub

Private Sub MNU_ANTIGUEDAD_Click()
    frm_rc_antiguedad.lbl_titulo = MNU_ANTIGUEDAD.Caption
    frm_rc_antiguedad.FraNavega = MNU_ANTIGUEDAD.Caption
    frm_rc_antiguedad.lbl_titulo2 = MNU_ANTIGUEDAD.Caption
    frm_rc_antiguedad.Show
End Sub

Private Sub Mnu_AsignacionTareasD_Click()
    Aux = "DNMOD"
    'mw_solicitud_calculo_trafico_mod
    mw_solicitud_calculo_trafico_mod.lbl_titulo = Mnu_ParametrosCalculo.Caption
    mw_solicitud_calculo_trafico_mod.FraNavega = Mnu_ParametrosCalculo.Caption
    mw_solicitud_calculo_trafico_mod.lbl_titulo2 = Mnu_ParametrosCalculo.Caption
    mw_solicitud_calculo_trafico_mod.Show
    
'    aw_p_ao_solicitud_calculo_trafico.lbl_titulo = Mnu_ParametrosCalculo.Caption
'    aw_p_ao_solicitud_calculo_trafico.FraNavega = Mnu_ParametrosCalculo.Caption
'    aw_p_ao_solicitud_calculo_trafico.lbl_titulo2 = Mnu_ParametrosCalculo.Caption
'    aw_p_ao_solicitud_calculo_trafico.Show
End Sub

Private Sub Mnu_AsignacionTareasM_Click()
    If glusuario = "ADMIN" Or glusuario = "CSALINAS" Or glusuario = "JSAAVEDRA" Or glusuario = "ACASTRO" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "KGARCIA" Or glusuario = "VMEJIA" Or glusuario = "FFLORES" Or glusuario = "PMAJLUF" Or glusuario = "PRODAS" Or glusuario = "CESCALANTE" Or glusuario = "TCRUZ" Or glusuario = "NPAREDES" Or glusuario = "ARODRIGUEZ" Or glusuario = "MARTEAGA" Or glusuario = "LVEDIA" Then
        Aux = "DNMAN"
        frm_to_cronograma_mensual.lbl_titulo = Mnu_AsignacionTareasM.Caption
        frm_to_cronograma_mensual.FraNavega = Mnu_AsignacionTareasM.Caption
        frm_to_cronograma_mensual.lbl_titulo2 = Mnu_AsignacionTareasM.Caption
        frm_to_cronograma_mensual.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_AsignaResponsableUnidad_Click()
    rw_unidad_vs_responsable.lbl_titulo = Mnu_AsignaResponsableUnidad.Caption
    rw_unidad_vs_responsable.FraNavega = Mnu_AsignaResponsableUnidad.Caption
    rw_unidad_vs_responsable.lbl_titulo2 = Mnu_AsignaResponsableUnidad.Caption
    rw_unidad_vs_responsable.Show
End Sub

Private Sub Mnu_Bancos_Click()
    fw_bancos.lbl_titulo = Mnu_Bancos.Caption
    fw_bancos.FraNavega = Mnu_Bancos.Caption
    fw_bancos.lbl_titulo2 = Mnu_Bancos.Caption
    fw_bancos.Show
End Sub

Private Sub Mnu_BitacoraInstalaciones_Click()
    Aux = "DNINS"
    tw_tecnico_bitacora.lbl_titulo = Mnu_BitacoraInstalaciones.Caption
    tw_tecnico_bitacora.FraNavega = Mnu_BitacoraInstalaciones.Caption
    tw_tecnico_bitacora.lbl_titulo2 = Mnu_BitacoraInstalaciones.Caption
    tw_tecnico_bitacora.Show
End Sub

Private Sub Mnu_CalendarioLaboral_Click()
    FrmRc_Calendario.lbl_titulo = Mnu_CalendarioLaboral.Caption
    'FrmRc_Calendario.FraNavega = Mnu_CalendarioLaboral.Caption
    'FrmRc_Calendario.lbl_titulo2 = Mnu_CalendarioLaboral.Caption
    FrmRc_Calendario.Show
End Sub

Private Sub Mnu_CalendarioZonas_Click()
    If glusuario = "JORAQUENI" Or glusuario = "ADMIN" Or glusuario = "KGARCIA" Or glusuario = "DOLMOS" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "VMEJIA" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Then
        Aux = "DNMAN"
        tw_calendario_zonas.lbl_titulo = Mnu_CalendarioZonas.Caption
        tw_calendario_zonas.FraNavega = Mnu_CalendarioZonas.Caption
        tw_calendario_zonas.lbl_titulo2 = Mnu_CalendarioZonas.Caption
        tw_calendario_zonas.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_cargoCuenta_Click()
    Aux = "30"
    fw_compras_fondos.lbl_titulo = Mnu_cargoCuenta.Caption
    fw_compras_fondos.FraNavega = Mnu_cargoCuenta.Caption
    fw_compras_fondos.lbl_titulo2 = Mnu_cargoCuenta.Caption
    fw_compras_fondos.Show
End Sub

Private Sub Mnu_CargosFuncionales_Click()
    frm_rc_cargos.lbl_titulo = Mnu_CargosFuncionales.Caption
    frm_rc_cargos.FraNavega = Mnu_CargosFuncionales.Caption
    frm_rc_cargos.lbl_titulo2 = Mnu_CargosFuncionales.Caption
    frm_rc_cargos.Show
End Sub

Private Sub Mnu_ClasificacionDocumentos_Click()
    gw_p_gc_documentos_clasificacion.lbl_titulo = Mnu_ClasificacionDocumentos.Caption
    gw_p_gc_documentos_clasificacion.FraNavega = Mnu_ClasificacionDocumentos.Caption
    gw_p_gc_documentos_clasificacion.lbl_titulo2 = Mnu_ClasificacionDocumentos.Caption
    gw_p_gc_documentos_clasificacion.Show
End Sub

Private Sub mnu_clasificacionTramites_Click()
    gw_p_gc_tipo_solicitud.lbl_titulo = mnu_clasificacionTramites.Caption
    gw_p_gc_tipo_solicitud.FraNavega = mnu_clasificacionTramites.Caption
    gw_p_gc_tipo_solicitud.lbl_titulo2 = mnu_clasificacionTramites.Caption
    gw_p_gc_tipo_solicitud.Show
End Sub

Private Sub Mnu_Cobranzas_Click()
'    Timer1.Enabled = False
    Aux = "DCOBR"
    fw_ventas_cobranzas.lbl_titulo1 = Mnu_Cobranzas.Caption
    fw_ventas_cobranzas.FraNavega2 = Mnu_Cobranzas.Caption
    fw_ventas_cobranzas.Show
End Sub

Private Sub Mnu_Componentes_Equipos_Click()
    aw_Componentes_Equipos.lbl_titulo = Mnu_Componentes_Equipos.Caption
    aw_Componentes_Equipos.Show
End Sub

Private Sub Mnu_ContratacionTecnicos_Click()
    Aux = "COMEX"
    Glaux = "CONTR"
'    mw_opcion_importacion.Show modal
    fw_compras_comex.lbl_titulo = Mnu_ContratacionTecnicos.Caption
    fw_compras_comex.FraNavega = Mnu_ContratacionTecnicos.Caption
    fw_compras_comex.lbl_titulo2 = Mnu_ContratacionTecnicos.Caption
    fw_compras_comex.Show
End Sub

Private Sub mnu_ControlAccesos_Click()
    FrmNivelesAcceso.Show
End Sub

Private Sub Mnu_cotizacion_servicio4_Click()
    Aux = "DNMAN"
    frm_to_solicitud_cotiza_venta.lbl_titulo = Mnu_cotizacion_servicio4.Caption
    frm_to_solicitud_cotiza_venta.FraNavega = Mnu_cotizacion_servicio4.Caption
    frm_to_solicitud_cotiza_venta.lbl_titulo2 = Mnu_cotizacion_servicio4.Caption
    frm_to_solicitud_cotiza_venta.Show
End Sub

Private Sub Mnu_CostosComercializacion_Click()
    frm_ac_costos_comercializacion.lbl_titulo = Mnu_CostosComercializacion.Caption
    frm_ac_costos_comercializacion.FraNavega = Mnu_CostosComercializacion.Caption
    frm_ac_costos_comercializacion.lbl_titulo2 = Mnu_CostosComercializacion.Caption
    frm_ac_costos_comercializacion.Show
End Sub

Private Sub Mnu_CronogramaInstalaciones_Click()
    If glusuario = "AURBINA" Or glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "CSALINAS" Or glusuario = "VPAREDES" Or glusuario = "CPAREDES" Or glusuario = "ACLAROS" Or glusuario = "AACOSTA" Or glusuario = "CURDININEA" Or glusuario = "FCABRERA" Or glusuario = "BRAMOS" Or glusuario = "JMAMANI" Or glusuario = "EVILLALOBOS" Or glusuario = "RMORA" Or glusuario = "BINFANTE" Or glusuario = "MRODRIGUEZ" Or glusuario = "AFLORES" Or glusuario = "RBUSTILLOS" Then
    '    Santa Cruz:
    'Cesar Paredes, Ariel Claros , Angel Acosta, Carlos Urdininea.
    '    'Cochabamba:
    'Franco Cabrera, Basilio Ramos, Juan carlos Mamani.
    '    'Sucre -Tarija:
    'Esteban Villalobos, Rolando Mora
    '    'La Paz-El Alto-Oruro:
    'Alvaro Urbina, Boris Infante, Mauricio Rodriguez, Alvaro Flores, Rodrigo Bustillos, Wilfredo Plata, Dulfredo Terceros.

        Aux = "DNINS"
        'tw_cronograma_mensual_inst
        tw_cronograma_mensual_inst.lbl_titulo = Mnu_CronogramaInstalaciones.Caption
        tw_cronograma_mensual_inst.FraNavega = Mnu_CronogramaInstalaciones.Caption
        'tw_cronograma_mensual_inst.lbl_titulo2 = Mnu_CronogramaInstalaciones.Caption
        tw_cronograma_mensual_inst.Show
        
'        tw_tecnico_cronograma.lbl_titulo = Mnu_CronogramaInstalaciones.Caption
'        tw_tecnico_cronograma.FraNavega = Mnu_CronogramaInstalaciones.Caption
'        tw_tecnico_cronograma.lbl_titulo2 = Mnu_CronogramaInstalaciones.Caption
'        tw_tecnico_cronograma.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_CronogramaMantenimiento_Click()
    If glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "ADMIN" Or glusuario = "FFLORES" Or glusuario = "PMAJLUF" Or glusuario = "JSAAVEDRA" Or glusuario = "ACASTRO" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "KGARCIA" Or glusuario = "VMEJIA" Or glusuario = "JMAMANI" Or glusuario = "CESCALANTE" Or glusuario = "PRODAS" Or glusuario = "CESCALANTE" Or glusuario = "CSALINAS" Or glusuario = "NPAREDES" Or glusuario = "ARODRIGUEZ" Or glusuario = "MARTEAGA" Then
        Aux = "DNMAN"
        Frm_to_tecnico_cronograma.lbl_titulo = Mnu_CronogramaMantenimiento.Caption
        Frm_to_tecnico_cronograma.FraNavega = Mnu_CronogramaMantenimiento.Caption
        Frm_to_tecnico_cronograma.lbl_titulo2 = Mnu_CronogramaMantenimiento.Caption
        Frm_to_tecnico_cronograma.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_CronogramaModernizacion_Click()
    Aux = "DNMOD"

    mw_solicitud_cotiza_venta.lbl_titulo = MnuCotizacionesEquipos.Caption
    mw_solicitud_cotiza_venta.FraNavega0 = MnuCotizacionesEquipos.Caption
'    mw_solicitud_cotiza_venta.lbl_titulo2 = MnuCotizacionesEquipos.Caption
    mw_solicitud_cotiza_venta.Show
End Sub

Private Sub Mnu_CuentasBancarias_Click()
    fw_cuenta_bancaria.lbl_titulo = Mnu_CuentasBancarias.Caption
    fw_cuenta_bancaria.FraNavega = Mnu_CuentasBancarias.Caption
    fw_cuenta_bancaria.lbl_titulo2 = Mnu_CuentasBancarias.Caption
    fw_cuenta_bancaria.Show
End Sub

Private Sub Mnu_DecargosCajaChica_Click()
    Aux = "28"
    fw_compras_fondos.lbl_titulo = Mnu_DecargosCajaChica.Caption
    fw_compras_fondos.FraNavega = Mnu_DecargosCajaChica.Caption
    fw_compras_fondos.lbl_titulo2 = Mnu_DecargosCajaChica.Caption
    fw_compras_fondos.Show
End Sub

Private Sub Mnu_Definicion_zonas_Click()
    'Dim e As Long
    'e = Shell(App.Path & "\SOFIA2012\TECNICO\MANTENIMIENTO\ZONAS_PILOTO\CAPA_PRESENTACION.exe", 1)
    ''C:\Jorge\SIGCGI\SOFIA2012\TECNICO\MANTENIMIENTO\ZONAS_PILOTO
'    tw_cronograma_zonas.lbl_titulo = Mnu_Definicion_zonas.Caption
'    tw_cronograma_zonas.FraNavega = Mnu_Definicion_zonas.Caption
'    tw_cronograma_zonas.lbl_titulo2 = Mnu_Definicion_zonas.Caption
'    tw_cronograma_zonas.Show
    If glusuario = "ADMIN" Or glusuario = "JMAMANI" Or glusuario = "VMEJIA" Or glusuario = "ACASTRO" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "KGARCIA" Or glusuario = "FFLORES" Or glusuario = "PMAJLUF" Or glusuario = "PRODAS" Or glusuario = "CESCALANTE" Or glusuario = "CSALINAS" Or glusuario = "NPAREDES" Or glusuario = "ARODRIGUEZ" Or glusuario = "MARTEAGA" Or glusuario = "LVEDIA" Then
        Aux = "DNMAN"
        tw_organizacion_zonas.lbl_titulo = Mnu_Definicion_zonas.Caption
        tw_organizacion_zonas.FraNavega = Mnu_Definicion_zonas.Caption
        tw_organizacion_zonas.lbl_titulo2 = Mnu_Definicion_zonas.Caption
        tw_organizacion_zonas.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_Departamentos_Click()
    gw_p_gc_departamento.lbl_titulo = Mnu_Departamentos.Caption
    gw_p_gc_departamento.FraNavega = Mnu_Departamentos.Caption
    gw_p_gc_departamento.lbl_titulo2 = Mnu_Departamentos.Caption
    gw_p_gc_departamento.Show
End Sub

Private Sub Mnu_DescargosFondosViajes_Click()
    Aux = "29"
    fw_compras_fondos.lbl_titulo = Mnu_DescargosFondosViajes.Caption
    fw_compras_fondos.FraNavega = Mnu_DescargosFondosViajes.Caption
    fw_compras_fondos.lbl_titulo2 = Mnu_DescargosFondosViajes.Caption
    fw_compras_fondos.Show
End Sub

Private Sub Mnu_Descarguio_Click()
    Aux = "COMEX"
    Glaux = "DESCA"
'    mw_opcion_importacion.Show modal
    fw_compras_comex.lbl_titulo = Mnu_Descarguio.Caption
    fw_compras_comex.FraNavega = Mnu_Descarguio.Caption
    fw_compras_comex.lbl_titulo2 = Mnu_Descarguio.Caption
    fw_compras_comex.Show
End Sub

Private Sub Mnu_DocumentosRespaldo_Click()
    gw_p_gc_documentos_respaldo.lbl_titulo = Mnu_DocumentosRespaldo.Caption
    gw_p_gc_documentos_respaldo.FraNavega = Mnu_DocumentosRespaldo.Caption
    gw_p_gc_documentos_respaldo.lbl_titulo2 = Mnu_DocumentosRespaldo.Caption
    gw_p_gc_documentos_respaldo.Show
End Sub

Private Sub mnu_dosificacion_facturas_Click()
    frm_fc_dosificacion_docs.lbl_titulo = mnu_dosificacion_facturas.Caption
    frm_fc_dosificacion_docs.FraNavega = mnu_dosificacion_facturas.Caption
    frm_fc_dosificacion_docs.lbl_titulo2 = mnu_dosificacion_facturas.Caption
    frm_fc_dosificacion_docs.Show
End Sub

Private Sub Mnu_EdificiosInstalacion_Click()
    If glusuario = "AURBINA" Or glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "CSALINAS" Or glusuario = "VPAREDES" Or glusuario = "CPAREDES" Or glusuario = "ACLAROS" Or glusuario = "AACOSTA" Or glusuario = "CURDININEA" Or glusuario = "FCABRERA" Or glusuario = "BRAMOS" Or glusuario = "JMAMANI" Or glusuario = "EVILLALOBOS" Or glusuario = "RMORA" Or glusuario = "BINFANTE" Or glusuario = "MRODRIGUEZ" Or glusuario = "AFLORES" Or glusuario = "RBUSTILLOS" Then
        Aux = "DNINS"
        tw_organizacion_zonas_inst.lbl_titulo = Mnu_EdificiosInstalacion.Caption
        tw_organizacion_zonas_inst.FraNavega = Mnu_EdificiosInstalacion.Caption
        tw_organizacion_zonas_inst.lbl_titulo2 = Mnu_EdificiosInstalacion.Caption
        tw_organizacion_zonas_inst.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_Ejecucion_Servicio_M_Click()
    If glusuario = "ADMIN" Or glusuario = "APALACIOS" Or glusuario = "JCASTRO" Or glusuario = "LVEDIA" Or glusuario = "JSAAVEDRA" Or glusuario = "KGARCIA" Or glusuario = "TCRUZ" Or glusuario = "EMACHICADO" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "VMEJIA" Or glusuario = "SQUISPE" Or glusuario = "FCABRERA" Or glusuario = "VPEÑA" Or glusuario = "TCASTILLO" Or glusuario = "GFLORES" Or glusuario = "BMONTAÑO" Or glusuario = "JCHIPANA" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "MMENACHO" Or glusuario = "CLEDEZMA" Or glusuario = "HMARIN" Or glusuario = "RCUELA" Or glusuario = "CSALINAS" Or glusuario = "EVILLALOBOS" Or glusuario = "FFLORES" Or glusuario = "PMAJLUF" Or glusuario = "PRODAS" Or glusuario = "CESCALANTE" Or glusuario = "ULEDEZMA" Or glusuario = "DVEGA" Or glusuario = "LVASQUEZ" Or glusuario = "RLAVAYEN" Or glusuario = "NPAREDES" Or glusuario = "ARODRIGUEZ" Or glusuario = "MARTEAGA" Or glusuario = "RPRIETO" Then
        Aux = "DNMAN"
'        frm_to_cronograma_certifica.lbl_titulo = Mnu_Ejecucion_Servicio_M.Caption
'        frm_to_cronograma_certifica.FraNavega = Mnu_Ejecucion_Servicio_M.Caption
'        frm_to_cronograma_certifica.lbl_titulo2 = Mnu_Ejecucion_Servicio_M.Caption
'        frm_to_cronograma_certifica.Show
        tw_cronograma_certifica.lbl_titulo = Mnu_Ejecucion_Servicio_M.Caption
        tw_cronograma_certifica.FraNavega = Mnu_Ejecucion_Servicio_M.Caption
        tw_cronograma_certifica.lbl_titulo2 = Mnu_Ejecucion_Servicio_M.Caption
        tw_cronograma_certifica.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_EjecucionInstalaciones_Click()
    If glusuario = "AURBINA" Or glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "CSALINAS" Or glusuario = "VPAREDES" Or glusuario = "CPAREDES" Or glusuario = "ACLAROS" Or glusuario = "AACOSTA" Or glusuario = "CURDININEA" Or glusuario = "FCABRERA" Or glusuario = "BRAMOS" Or glusuario = "JMAMANI" Or glusuario = "EVILLALOBOS" Or glusuario = "RMORA" Or glusuario = "BINFANTE" Or glusuario = "MRODRIGUEZ" Or glusuario = "AFLORES" Or glusuario = "RBUSTILLOS" Then
        Aux = "DNINS"
        tw_cronograma_certifica_inst.lbl_titulo = Mnu_EjecucionInstalaciones.Caption
        tw_cronograma_certifica_inst.FraNavega = Mnu_EjecucionInstalaciones.Caption
        tw_cronograma_certifica_inst.lbl_titulo2 = Mnu_EjecucionInstalaciones.Caption
        tw_cronograma_certifica_inst.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_Facturacion_Click()
    Aux = "DCOBR"
    fw_facturacion.lbl_titulo = Mnu_Facturacion.Caption
    fw_facturacion.FraNavega = Mnu_Facturacion.Caption
    'Fw_facturacion.lbl_titulo2 = Mnu_Cobranzas.Caption
    fw_facturacion.Show
End Sub

Private Sub Mnu_FacturacionAntes_Click()
    Aux = "DCOBR"
    Frm_ao_ventas_cobranzas.lbl_titulo = Mnu_FacturacionAntes.Caption
    Frm_ao_ventas_cobranzas.FraNavega = Mnu_FacturacionAntes.Caption
    'Frm_ao_ventas_cobranzas.lbl_titulo2 = Mnu_Cobranzas.Caption
    Frm_ao_ventas_cobranzas.Show
End Sub

'Private Sub Mnu_EsteticaCabina_Click()
'    aw_p_ac_bienes_equipo_cabina_estetica.lbl_titulo = Mnu_EsteticaCabina.Caption
'    aw_p_ac_bienes_equipo_cabina_estetica.FraNavega = Mnu_EsteticaCabina.Caption
'    aw_p_ac_bienes_equipo_cabina_estetica.lbl_titulo2 = Mnu_EsteticaCabina.Caption
'    aw_p_ac_bienes_equipo_cabina_estetica.Show
'End Sub

Private Sub Mnu_FichaAdministracionPersonal_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        rw_ficha_rrhh.lbl_titulo = Mnu_FichaAdministracionPersonal.Caption
        'frmBeneficiario_Admin.FraNavega = Mnu_FichaAdministracionPersonal.Caption
        rw_ficha_rrhh.lbl_titulo2 = Mnu_FichaAdministracionPersonal.Caption
        rw_ficha_rrhh.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_FileControlPersonal_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "DVARGAS" Or glusuario = "DOLMOS" Or glusuario = "CSALINAS" Then
        frmBeneficiario_Control.lbl_titulo = Mnu_FileControlPersonal.Caption
        frmBeneficiario_Control.FraNavega = Mnu_FileControlPersonal.Caption
        frmBeneficiario_Control.lbl_titulo2 = Mnu_FileControlPersonal.Caption
        frmBeneficiario_Control.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub mnu_Financiadores_Click()
    fw_Organismo_Financiador.lbl_titulo = mnu_Financiadores.Caption
    fw_Organismo_Financiador.FraNavega = mnu_Financiadores.Caption
    fw_Organismo_Financiador.lbl_titulo2 = mnu_Financiadores.Caption
    fw_Organismo_Financiador.Show
End Sub

Private Sub mnu_FuentesFinanciamiento_Click()
    fw_fuente_financiamiento.lbl_titulo = mnu_FuentesFinanciamiento.Caption
    fw_fuente_financiamiento.FraNavega = mnu_FuentesFinanciamiento.Caption
    fw_fuente_financiamiento.lbl_titulo2 = mnu_FuentesFinanciamiento.Caption
    fw_fuente_financiamiento.Show
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

'Private Sub EjecucionPorUni_Click()
''  glRepPresup = "REP004"
''  frmRepPresupuesto.Show
'  Dim e As Long
'  e = Shell(App.Path & "\reportes\presupuesto\reppresupuesto.exe " & GlUsuario & " REP004", 1)
'End Sub

'Private Sub Exportar_Click()
'  Frmexporta.Show vbModal
'End Sub
'
'Private Sub Importar_Click()
'  FrmImporta.Show vbModal
'End Sub
'
'Private Sub imprecepcion_Click()
'  Dim rstao_solicitud_recibido As New ADODB.Recordset
'  Dim sino As String
'
'  Set rstao_solicitud_recibido = New ADODB.Recordset
'  If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'  rstao_solicitud_recibido.Open "select * from ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
'  Print rstao_solicitud_recibido.RecordCount
''-------
'  '  Cry.Reset
'  Cry.ReportFileName = App.Path & "\FormsMigrar\RecepMigrar.rpt"
''  Cry.SelectionFormula = "{Vi_Fo_ingresos_rep.Maquina} = '" & GlMaquina & "'"
'
'  Cry.WindowShowPrintBtn = True
'  Cry.WindowShowExportBtn = True
'  Cry.WindowShowRefreshBtn = True
'  Cry.WindowShowPrintSetupBtn = True
'  Cry.WindowShowZoomCtl = True
'  Cry.WindowState = crptMaximized
'  Cry.PageZoom (200)
'  iResult = Cry.PrintReport
'  If iResult <> 0 Then
'      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
'  End If
''  sino = MsgBox("¿La impresión concluyó con EXITO?", vbQuestion + vbYesNo, "Confirmando Impresión... ")
''  If sino = vbYes Then
''  rstao_solicitud_recibido.MoveFirst
'    While Not rstao_solicitud_recibido.EOF
'      rstao_solicitud_recibido.Delete
'      rstao_solicitud_recibido.Update
'      rstao_solicitud_recibido.MoveNext
'    Wend
''  End If
'  If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'
'End Sub

Private Sub mnu_GerenciaGeneral_Click()
    gw_p_gc_direccion_general.lbl_titulo = mnu_GerenciaGeneral.Caption
    gw_p_gc_direccion_general.FraNavega = mnu_GerenciaGeneral.Caption
    gw_p_gc_direccion_general.lbl_titulo2 = mnu_GerenciaGeneral.Caption
    gw_p_gc_direccion_general.Show
End Sub

Private Sub mnu_GerenciasOperativas_Click()
    gw_p_gc_direccion_administrativa.lbl_titulo = mnu_GerenciasOperativas.Caption
    gw_p_gc_direccion_administrativa.FraNavega = mnu_GerenciasOperativas.Caption
    gw_p_gc_direccion_administrativa.lbl_titulo2 = mnu_GerenciasOperativas.Caption
    gw_p_gc_direccion_administrativa.Show
End Sub

'Private Sub Mnu_GrupoCoches_Click()
'    aw_p_ac_bienes_equipo_grupo_coches.lbl_titulo = Mnu_GrupoCoches.Caption
'    aw_p_ac_bienes_equipo_grupo_coches.FraNavega = Mnu_GrupoCoches.Caption
'    aw_p_ac_bienes_equipo_grupo_coches.lbl_titulo2 = Mnu_GrupoCoches.Caption
'    aw_p_ac_bienes_equipo_grupo_coches.Show
'End Sub

Private Sub mnu_GruposIngresos_Click()
    fw_ingresos_grupo.lbl_titulo = mnu_GruposIngresos.Caption
    fw_ingresos_grupo.FraNavega = mnu_GruposIngresos.Caption
    fw_ingresos_grupo.lbl_titulo2 = mnu_GruposIngresos.Caption
    fw_ingresos_grupo.Show
End Sub

Private Sub Mnu_HorariosLaborales_Click()
    FrmRc_Horarios.lbl_titulo = Mnu_HorariosLaborales.Caption
    'FrmRc_Horarios.FraNavega = Mnu_HorariosLaborales.Caption
    'FrmRc_Horarios.lbl_titulo2 = Mnu_HorariosLaborales.Caption
    FrmRc_Horarios.Show
End Sub

Private Sub Mnu_IdentificacionClienteAjustes_Click()
    'Aux = "DNAJS"
    Aux = "DNINS"
    tw_identificacion_cliente.lbl_titulo = Mnu_IdentificacionClienteAjustes.Caption
    tw_identificacion_cliente.FraNavega = Mnu_IdentificacionClienteAjustes.Caption
    tw_identificacion_cliente.lbl_titulo2 = Mnu_IdentificacionClienteAjustes.Caption
    tw_identificacion_cliente.Show
End Sub

Private Sub Mnu_IdentificacionClienteEmergencia_Click()
    Aux = "DNEME"
    tw_identificacion_cliente.lbl_titulo = Mnu_IdentificacionClienteEmergencia.Caption
    tw_identificacion_cliente.FraNavega = Mnu_IdentificacionClienteEmergencia.Caption
    tw_identificacion_cliente.lbl_titulo2 = Mnu_IdentificacionClienteEmergencia.Caption
    tw_identificacion_cliente.Show
    
'    frm_to_id_emergencia.lbl_titulo = Mnu_IdentificacionClienteEmergencia.Caption
'    frm_to_id_emergencia.FraNavega = Mnu_IdentificacionClienteEmergencia.Caption
'    frm_to_id_emergencia.lbl_titulo2 = Mnu_IdentificacionClienteEmergencia.Caption
'    frm_to_id_emergencia.Show
End Sub

Private Sub Mnu_IdentificacionClienteInstalacion_Click()
    Aux = "DNINS"
    tw_identificacion_cliente.lbl_titulo = Mnu_IdentificacionClienteInstalacion.Caption
    tw_identificacion_cliente.FraNavega = Mnu_IdentificacionClienteInstalacion.Caption
    tw_identificacion_cliente.lbl_titulo2 = Mnu_IdentificacionClienteInstalacion.Caption
    tw_identificacion_cliente.Show
End Sub

Private Sub Mnu_IdentificacionClienteMantenimiento_Click()
    If glusuario = "JORAQUENI" Then
        MsgBox "El Usuario No tiene acceso, Consulte con el Administrador del Sistema ...", , "Atención"
        Exit Sub
    End If
    Aux = "DNMAN"
    tw_identificacion_cliente.lbl_titulo = Mnu_IdentificacionClienteMantenimiento.Caption
    tw_identificacion_cliente.FraNavega = Mnu_IdentificacionClienteMantenimiento.Caption
    tw_identificacion_cliente.lbl_titulo2 = Mnu_IdentificacionClienteMantenimiento.Caption
    tw_identificacion_cliente.Show
End Sub

Private Sub Mnu_IdentificacionClienteModernizacion_Click()
    Aux = "DNMOD"
    mw_solicitud.lbl_titulo = Mnu_IdentificacionClienteModernizacion.Caption
    mw_solicitud.FraNavega = Mnu_IdentificacionClienteModernizacion.Caption
    mw_solicitud.lbl_titulo2 = Mnu_IdentificacionClienteModernizacion.Caption
    mw_solicitud.Show
End Sub

Private Sub Mnu_IdentificacionClienteReparacion_Click()
    Aux = "DNREP"
    tw_identificacion_cliente.lbl_titulo = Mnu_IdentificacionClienteReparacion.Caption
    tw_identificacion_cliente.FraNavega = Mnu_IdentificacionClienteReparacion.Caption
    tw_identificacion_cliente.lbl_titulo2 = Mnu_IdentificacionClienteReparacion.Caption
    tw_identificacion_cliente.Show
End Sub

Private Sub Mnu_ImportarRegistroAsistencia_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        Dim e As Long
        e = Shell(App.Path & "\Asistencia\sofiaNET_6.exe", 1)
'        rw_importar_registro_asistencia.lbl_titulo = Mnu_ImportarRegistroAsistencia.Caption
'        rw_importar_registro_asistencia.FraNavega = Mnu_ImportarRegistroAsistencia.Caption
'        rw_importar_registro_asistencia.lbl_titulo2 = Mnu_ImportarRegistroAsistencia.Caption
'        rw_importar_registro_asistencia.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub mnu_IngresosAlmacen_Click()
    Aux = "UALMI"       'INSUMOS Y MATERIALES
    Glaux = "UALMI"
    fw_compras_gral.lbl_titulo = mnu_IngresosAlmacen.Caption
    fw_compras_gral.FraNavega = mnu_IngresosAlmacen.Caption
    fw_compras_gral.lbl_titulo2 = mnu_IngresosAlmacen.Caption
    fw_compras_gral.Show
'    fw_compras_gral.lbl_titulo = mnu_IngresosAlmacen.Caption
'    fw_compras_gral.FraNavega = mnu_IngresosAlmacen.Caption
'    fw_compras_gral.lbl_titulo2 = mnu_IngresosAlmacen.Caption
'    fw_compras_gral.Show
End Sub

Private Sub mnu_IngresosAlmacen2_Click()
    Aux = "30"
    Glaux = "UALMI"
    fw_compras_fondos.lbl_titulo = Mnu_cargoCuenta.Caption
    fw_compras_fondos.FraNavega = Mnu_cargoCuenta.Caption
    fw_compras_fondos.lbl_titulo2 = Mnu_cargoCuenta.Caption
    fw_compras_fondos.Show
End Sub

Private Sub mnu_IngresosAlmacenHerr_Click()
    Aux = "UALMH"       'HERRAMIENTAS
    Glaux = "UALMH"
    fw_compras_gral.lbl_titulo = mnu_IngresosAlmacenHerr.Caption
    fw_compras_gral.FraNavega = mnu_IngresosAlmacenHerr.Caption
    fw_compras_gral.lbl_titulo2 = mnu_IngresosAlmacenHerr.Caption
    fw_compras_gral.Show
End Sub

Private Sub mnu_IngresosAlmacenHerr2_Click()
    Aux = "30"
    Glaux = "UALMH"
    fw_compras_fondos.lbl_titulo = Mnu_cargoCuenta.Caption
    fw_compras_fondos.FraNavega = Mnu_cargoCuenta.Caption
    fw_compras_fondos.lbl_titulo2 = Mnu_cargoCuenta.Caption
    fw_compras_fondos.Show
End Sub

Private Sub mnu_IngresosAlmacenRep_Click()
    Aux = "UALMR"       'REPUESTOS
    Glaux = "UALMR"
    fw_compras_gral.lbl_titulo = mnu_IngresosAlmacenRep.Caption
    fw_compras_gral.FraNavega = mnu_IngresosAlmacenRep.Caption
    fw_compras_gral.lbl_titulo2 = mnu_IngresosAlmacenRep.Caption
    fw_compras_gral.Show
End Sub

Private Sub mnu_IngresosAlmacenRep2_Click()
    Aux = "30"
    Glaux = "UALMR"
    fw_compras_fondos.lbl_titulo = Mnu_cargoCuenta.Caption
    fw_compras_fondos.FraNavega = Mnu_cargoCuenta.Caption
    fw_compras_fondos.lbl_titulo2 = Mnu_cargoCuenta.Caption
    fw_compras_fondos.Show
End Sub

Private Sub mnu_InventarioAlmacen_Click()
    Aux = "I"
    aw_almacen_inventario.lbl_titulo = mnu_InventarioAlmacen.Caption
    'aw_almacen_inventario.FraNavega = Mnu_LineasEquipos.Caption
    'aw_almacen_inventario.lbl_titulo2 = Mnu_LineasEquipos.Caption
    aw_almacen_inventario.Show
End Sub

Private Sub mnu_InventarioAlmacenHerr_Click()
     Aux = "H"
    aw_almacen_inventario.lbl_titulo = mnu_InventarioAlmacenHerr.Caption
    'aw_almacen_inventario.FraNavega = Mnu_LineasEquipos.Caption
    'aw_almacen_inventario.lbl_titulo2 = Mnu_LineasEquipos.Caption
    aw_almacen_inventario.Show
End Sub

Private Sub mnu_InventarioAlmacenRep_Click()
    Aux = "R"
    aw_almacen_inventario.lbl_titulo = mnu_InventarioAlmacenRep.Caption
    'aw_almacen_inventario.FraNavega = Mnu_LineasEquipos.Caption
    'aw_almacen_inventario.lbl_titulo2 = Mnu_LineasEquipos.Caption
    aw_almacen_inventario.Show
End Sub

'Private Sub Mnu_LineasEquipos_Click()
'    frm_ac_bienes_tecnologia_linea.lbl_titulo = Mnu_LineasEquipos.Caption
'    frm_ac_bienes_tecnologia_linea.FraNavega = Mnu_LineasEquipos.Caption
'    frm_ac_bienes_tecnologia_linea.lbl_titulo2 = Mnu_LineasEquipos.Caption
'    frm_ac_bienes_tecnologia_linea.Show
'End Sub

Private Sub Mnu_MayorAuxiliar_Click()
    Fw_Mayor_Auxiliar.lbl_titulo = Mnu_MayorAuxiliar.Caption
    'frm_rc_modalidad_contratacion.FraNavega = Mnu_ModalidadesContratacion.Caption
    'frm_rc_modalidad_contratacion.lbl_titulo2 = Mnu_ModalidadesContratacion.Caption
    Fw_Mayor_Auxiliar.Show
End Sub

Private Sub Mnu_ModalidadesContratacion_Click()
    frm_rc_modalidad_contratacion.lbl_titulo = Mnu_ModalidadesContratacion.Caption
    frm_rc_modalidad_contratacion.FraNavega = Mnu_ModalidadesContratacion.Caption
    frm_rc_modalidad_contratacion.lbl_titulo2 = Mnu_ModalidadesContratacion.Caption
    frm_rc_modalidad_contratacion.Show
End Sub

'Private Sub Mnu_modelos_Click()
'    frm_ac_bienes_modelos.lbl_titulo = Mnu_modelos.Caption
'    frm_ac_bienes_modelos.FraNavega = Mnu_modelos.Caption
'    frm_ac_bienes_modelos.lbl_titulo2 = Mnu_modelos.Caption
'    frm_ac_bienes_modelos.Show
'End Sub

Private Sub Mnu_Motivos_Procesos_Click()
    frm_rc_motivos_procesos.lbl_titulo = Mnu_Motivos_Procesos.Caption
    frm_rc_motivos_procesos.FraNavega = Mnu_Motivos_Procesos.Caption
    frm_rc_motivos_procesos.lbl_titulo2 = Mnu_Motivos_Procesos.Caption
    frm_rc_motivos_procesos.Show
End Sub

Private Sub Mnu_Municipios_Click()
    frm_gc_municipio.lbl_titulo = Mnu_Municipios.Caption
    frm_gc_municipio.FraNavega = Mnu_Municipios.Caption
    frm_gc_municipio.lbl_titulo2 = Mnu_Municipios.Caption
    frm_gc_municipio.Show
End Sub

Private Sub Mnu_Nacionalizacion_Click()
    Aux = "COMEX"
    Glaux = "ADUAN"
    ' mw_opcion_importacion.Show modal
    fw_compras_comex.lbl_titulo = Mnu_Nacionalizacion.Caption
    fw_compras_comex.FraNavega = Mnu_Nacionalizacion.Caption
    fw_compras_comex.lbl_titulo2 = Mnu_Nacionalizacion.Caption
    fw_compras_comex.Show
End Sub

Private Sub Mnu_NivelesEducacion_Click()
    frm_rc_nivel_educacion.lbl_titulo = Mnu_NivelesEducacion.Caption
    frm_rc_nivel_educacion.FraNavega = Mnu_NivelesEducacion.Caption
    frm_rc_nivel_educacion.lbl_titulo2 = Mnu_NivelesEducacion.Caption
    frm_rc_nivel_educacion.Show
End Sub

Private Sub mnu_NotaCreditoDebito_Click()
    Aux = "DCOBR"
    fw_orden_cobranza.FraNavega = mnu_NotaCreditoDebito.Caption
    fw_orden_cobranza.lbl_titulo = mnu_NotaCreditoDebito.Caption
    fw_orden_cobranza.Show
End Sub

Private Sub Mnu_Paises_Click()
    gw_p_gc_pais.lbl_titulo = Mnu_Paises.Caption
    gw_p_gc_pais.FraNavega = Mnu_Paises.Caption
    gw_p_gc_pais.lbl_titulo2 = Mnu_Paises.Caption
    gw_p_gc_pais.Show
End Sub

Private Sub Mnu_ParametrosCalculo_Click()
    Aux = "DVTA"
    aw_p_ao_solicitud_calculo_trafico.lbl_titulo = Mnu_ParametrosCalculo.Caption
    aw_p_ao_solicitud_calculo_trafico.FraNavega = Mnu_ParametrosCalculo.Caption
    aw_p_ao_solicitud_calculo_trafico.lbl_titulo2 = Mnu_ParametrosCalculo.Caption
    aw_p_ao_solicitud_calculo_trafico.Show
End Sub

Private Sub Mnu_Parentesco_Click()
    frm_rc_personal_parentesco.lbl_titulo = Mnu_Parentesco.Caption
    frm_rc_personal_parentesco.FraNavega = Mnu_Parentesco.Caption
    frm_rc_personal_parentesco.lbl_titulo2 = Mnu_Parentesco.Caption
    frm_rc_personal_parentesco.Show
End Sub

Private Sub mnu_PartidasGasto_Click()
    fw_partida_gasto.lbl_titulo = mnu_PartidasGasto.Caption
    fw_partida_gasto.FraNavega = mnu_PartidasGasto.Caption
    fw_partida_gasto.lbl_titulo2 = mnu_PartidasGasto.Caption
    fw_partida_gasto.Show
    
End Sub

Private Sub Mnu_PlanCuentas_Click()
    fw_plan_cuentas.lbl_titulo = Mnu_PlanCuentas.Caption
    fw_plan_cuentas.FraNavega = Mnu_PlanCuentas.Caption
    fw_plan_cuentas.lbl_titulo2 = Mnu_PlanCuentas.Caption
    fw_plan_cuentas.Show
End Sub

Private Sub Mnu_PlanillasGrupos_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        rw_planilla_grupo.lbl_titulo = Mnu_PlanillasGrupos.Caption
        rw_planilla_grupo.FraNavega = Mnu_PlanillasGrupos.Caption
        rw_planilla_grupo.lbl_titulo2 = Mnu_PlanillasGrupos.Caption
        rw_planilla_grupo.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_PlanillasSubGrupos_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        rw_planilla_sub_grupo.lbl_titulo = Mnu_PlanillasSubGrupos.Caption
        rw_planilla_sub_grupo.FraNavega = Mnu_PlanillasSubGrupos.Caption
        rw_planilla_sub_grupo.lbl_titulo2 = Mnu_PlanillasSubGrupos.Caption
        rw_planilla_sub_grupo.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_PrestamosAPersonal_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        rw_prestamos.lbl_titulo = Mnu_PrestamosAPersonal.Caption
        rw_prestamos.FraNavega = Mnu_PrestamosAPersonal.Caption
        rw_prestamos.lbl_titulo2 = Mnu_PrestamosAPersonal.Caption
        rw_prestamos.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_ProcesoAjustes_Click()
    'Aux = "DNAJS"
    Aux = "DNINS"
    tw_tecnico_venta.lbl_titulo = Mnu_ProcesoAjustes.Caption
    tw_tecnico_venta.FraNavega = Mnu_ProcesoAjustes.Caption
    tw_tecnico_venta.lbl_titulo2 = Mnu_ProcesoAjustes.Caption
    tw_tecnico_venta.Show
End Sub

Private Sub Mnu_ProcesoCompras_Click()
    Aux = "DVTA"
    mw_ventas_seguimiento.lbl_titulo = Mnu_ProcesoCompras.Caption
    mw_ventas_seguimiento.FraNavega = Mnu_ProcesoCompras.Caption
    mw_ventas_seguimiento.lbl_titulo2 = Mnu_ProcesoCompras.Caption
    mw_ventas_seguimiento.Show
    
'    frm_ao_compra_proceso.lbl_titulo = Mnu_ProcesoCompras.Caption
'    frm_ao_compra_proceso.FraNavega = Mnu_ProcesoCompras.Caption
'    frm_ao_compra_proceso.lbl_titulo2 = Mnu_ProcesoCompras.Caption
'    frm_ao_compra_proceso.Show
End Sub

Private Sub mnu_ProcesoContratacion_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        Aux = "DRRHH"
        rw_contratacion_personal.lbl_titulo = mnu_ProcesoContratacion.Caption
        rw_contratacion_personal.FraNavega = mnu_ProcesoContratacion.Caption
        rw_contratacion_personal.lbl_titulo2 = mnu_ProcesoContratacion.Caption
        rw_contratacion_personal.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_ProcesoEmergencia_Click()
    Aux = "DNEME"
    tw_tecnico_venta.lbl_titulo = Mnu_ProcesoEmergencia.Caption
    tw_tecnico_venta.FraNavega = Mnu_ProcesoEmergencia.Caption
    tw_tecnico_venta.lbl_titulo2 = Mnu_ProcesoEmergencia.Caption
    tw_tecnico_venta.Show
End Sub

Private Sub Mnu_ProcesoInstalaciones_Click()
    Aux = "DNINS"
    tw_tecnico_venta.lbl_titulo = Mnu_ProcesoInstalaciones.Caption
    tw_tecnico_venta.FraNavega = Mnu_ProcesoInstalaciones.Caption
    tw_tecnico_venta.lbl_titulo2 = Mnu_ProcesoInstalaciones.Caption
    tw_tecnico_venta.Show
End Sub

Private Sub Mnu_ProcesoMantenimiento_Click()
    If glusuario = "JORAQUENI" Then
        MsgBox "El Usuario No tiene acceso, Consulte con el Administrador del Sistema ...", , "Atención"
        Exit Sub
    End If
    Aux = "DNMAN"
    tw_tecnico_venta.lbl_titulo = Mnu_ProcesoMantenimiento.Caption
    tw_tecnico_venta.FraNavega = Mnu_ProcesoMantenimiento.Caption
    tw_tecnico_venta.lbl_titulo2 = Mnu_ProcesoMantenimiento.Caption
    tw_tecnico_venta.Show
End Sub

Private Sub Mnu_ProcesoModernizacion_Click()
    Aux = "DNMOD"
    mw_ventas_cabecera.lbl_titulo = Mnu_ProcesoModernizacion.Caption
    mw_ventas_cabecera.FraNavega = Mnu_ProcesoModernizacion.Caption
    mw_ventas_cabecera.lbl_titulo2 = Mnu_ProcesoModernizacion.Caption
    mw_ventas_cabecera.Show
End Sub

Private Sub Mnu_ProcesoPlanillas_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
'        frm_ro_pagos_grupos_principal.lbl_titulo = Mnu_ProcesoPlanillas.Caption
'        frm_ro_pagos_grupos_principal.FraNavega = Mnu_ProcesoPlanillas.Caption
'        frm_ro_pagos_grupos_principal.lbl_titulo2 = Mnu_ProcesoPlanillas.Caption
'        frm_ro_pagos_grupos_principal.Show
        rw_planillas_procesos.lbl_titulo = Mnu_ProcesoPlanillas.Caption
        rw_planillas_procesos.FraNavega = Mnu_ProcesoPlanillas.Caption
        rw_planillas_procesos.lbl_titulo2 = Mnu_ProcesoPlanillas.Caption
        rw_planillas_procesos.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_ProcesoReparacion_Click()
    Aux = "DNREP"
    tw_tecnico_venta.lbl_titulo = Mnu_ProcesoReparacion.Caption
    tw_tecnico_venta.FraNavega = Mnu_ProcesoReparacion.Caption
    tw_tecnico_venta.lbl_titulo2 = Mnu_ProcesoReparacion.Caption
    tw_tecnico_venta.Show
End Sub

Private Sub Mnu_ProcesoVentas_Click()
    Aux = "DVTA"
    mw_ventas_cabecera.lbl_titulo = Mnu_ProcesoVentas.Caption
    mw_ventas_cabecera.FraNavega = Mnu_ProcesoVentas.Caption
    mw_ventas_cabecera.lbl_titulo2 = Mnu_ProcesoVentas.Caption
    mw_ventas_cabecera.Show
End Sub

Private Sub mnu_ProgramasProyectos_Click()
    fw_Estructura_Programatica.lbl_titulo = mnu_ProgramasProyectos.Caption
    fw_Estructura_Programatica.FraNavega = mnu_ProgramasProyectos.Caption
    fw_Estructura_Programatica.lbl_titulo2 = mnu_ProgramasProyectos.Caption
    fw_Estructura_Programatica.Show
End Sub

Private Sub Mnu_ProveedoresEquipos_Click()
    Aux = "COMEX"
    Glaux = "PROVI"
'    mw_opcion_importacion.Show 'vbModal
    fw_compras_comex.lbl_titulo = Mnu_ProveedoresEquipos.Caption
    fw_compras_comex.FraNavega = Mnu_ProveedoresEquipos.Caption
    fw_compras_comex.lbl_titulo2 = Mnu_ProveedoresEquipos.Caption
    fw_compras_comex.Show
End Sub

Private Sub Mnu_Provincias_Click()
    frm_gc_provincia.lbl_titulo = Mnu_Provincias.Caption
    frm_gc_provincia.FraNavega = Mnu_Provincias.Caption
    frm_gc_provincia.lbl_titulo2 = Mnu_Provincias.Caption
    frm_gc_provincia.Show
End Sub

Private Sub mnu_proyecto_edificacion_Click()
'    gw_edificaciones.lbl_titulo = mnu_proyecto_edificacion.Caption
'    gw_edificaciones.FraNavega = mnu_proyecto_edificacion.Caption
'    gw_edificaciones.lbl_titulo2 = mnu_proyecto_edificacion.Caption
'    gw_edificaciones.Show
End Sub

Private Sub Mnu_PuestosOrganizacionales_Click()
    frm_rc_puestos.lbl_titulo = Mnu_PuestosOrganizacionales.Caption
    frm_rc_puestos.FraNavega = Mnu_PuestosOrganizacionales.Caption
    frm_rc_puestos.lbl_titulo2 = Mnu_PuestosOrganizacionales.Caption
    frm_rc_puestos.Show
End Sub

Private Sub Mnu_RecibosOficiales_Click()
    Aux = "R-640"
    'fw_ventas_cobranzas.lbl_titulo = Mnu_Cobranzas.Caption
    fw_recibos_oficiales.FraNavega = Mnu_RecibosOficiales.Caption + " - " + mnu_Tesoreria.Caption
    fw_recibos_oficiales.lbl_titulo = Mnu_RecibosOficiales.Caption + " - " + mnu_Tesoreria.Caption
    fw_recibos_oficiales.Show
End Sub

Private Sub Mnu_RecibosOficialesEgresos_Click()
    Aux = "R-643"
    'fw_ventas_cobranzas.lbl_titulo = Mnu_Cobranzas.Caption
    fw_recibos_oficiales_egresos.FraNavega = Mnu_RecibosOficialesEgresos.Caption + " - " + mnu_Tesoreria.Caption
    fw_recibos_oficiales_egresos.lbl_titulo = Mnu_RecibosOficialesEgresos.Caption + " - " + mnu_Tesoreria.Caption
    'fw_recibos_oficiales_egresos.lbl_titulo2 = "ORIGEN - " + Mnu_RecibosOficialesEgresos.Caption + " - " + mnu_Tesoreria.Caption
    fw_recibos_oficiales_egresos.Show
End Sub

Private Sub Mnu_RecibosOficialesEgresos2_Click()
    Aux = "R-643"
    Glaux = "UALM"
    fw_recibos_oficiales_egresos.FraNavega = Mnu_RecibosOficialesEgresos.Caption + " - " + mnu_Tesoreria.Caption
    fw_recibos_oficiales_egresos.lbl_titulo = Mnu_RecibosOficialesEgresos.Caption + " - " + mnu_Tesoreria.Caption
    fw_recibos_oficiales_egresos.Show
End Sub

Private Sub Mnu_RegistroDiario_Click()
    fw_contab_diario.lbl_titulo = Mnu_RegistroDiario.Caption
    fw_contab_diario.FraNavega = Mnu_RegistroDiario.Caption
    fw_contab_diario.lbl_titulo2 = Mnu_RegistroDiario.Caption
    fw_contab_diario.Show
End Sub

Private Sub Mnu_RegistroGastos_Click()
    ff_egresos.Show
End Sub

Private Sub mnu_registroIngresos_Click()
'    frm_fo_ingresos.lbl_titulo = mnu_registroIngresos.Caption
'    frm_fo_ingresos.FraNavega = mnu_registroIngresos.Caption
'    frm_fo_ingresos.lbl_titulo2 = mnu_registroIngresos.Caption
'    frm_fo_ingresos.Show
End Sub

Private Sub Mnu_RegistroPersonas1_Click()
    Glaux = "0"
    gw_p_gc_beneficiario_persona.lbl_titulo = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.FraNavega = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.lbl_titulo2 = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.Show
End Sub

Private Sub Mnu_RegistroPagos_Click()
    Aux = "DCONT"
    Glaux = "GADM"
    fw_compras_gral.lbl_titulo = Mnu_RegistroPagos.Caption
    fw_compras_gral.FraNavega = Mnu_RegistroPagos.Caption
    fw_compras_gral.lbl_titulo2 = Mnu_RegistroPagos.Caption
    fw_compras_gral.Show
'    fw_gastos_detalle.lbl_titulo = Mnu_RegistroPagos.Caption
'    fw_gastos_detalle.FraNavega = Mnu_RegistroPagos.Caption
'    fw_gastos_detalle.lbl_titulo2 = Mnu_RegistroPagos.Caption
'    fw_gastos_detalle.Show
End Sub

Private Sub Mnu_Relacionador_Gastos_Click()
    fw_Relacionador_Gastos.lbl_titulo = Mnu_Relacionador_Gastos.Caption
    fw_Relacionador_Gastos.FraNavega = Mnu_Relacionador_Gastos.Caption
    fw_Relacionador_Gastos.lbl_titulo2 = Mnu_Relacionador_Gastos.Caption
    fw_Relacionador_Gastos.Show
End Sub

Private Sub Mnu_RelacionadorIngresos_Click()
    fw_Relacionador_Ingresos.lbl_titulo = mnu_RubrosIngresos.Caption
    fw_Relacionador_Ingresos.FraNavega = mnu_RubrosIngresos.Caption
    fw_Relacionador_Ingresos.lbl_titulo2 = mnu_RubrosIngresos.Caption
    fw_Relacionador_Ingresos.Show
End Sub

Private Sub Mnu_Reportes_RRHH_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "RCUELA" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        frm_ReportesRRH.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_ReportesCobranzas_Click()
    Fw_ReportesCobranzas.lbl_titulo = Mnu_ReportesCobranzas.Caption
'    frm_fc_ingresos_rubro.FraNavega = Mnu_ReportesCobranzas.Caption
'    frm_fc_ingresos_rubro.lbl_titulo2 = Mnu_ReportesCobranzas.Caption
    Fw_ReportesCobranzas.Show
End Sub

Private Sub mnu_RubrosIngresos_Click()
    fw_ingresos_rubro.lbl_titulo = mnu_RubrosIngresos.Caption
    fw_ingresos_rubro.FraNavega = mnu_RubrosIngresos.Caption
    'fw_ingresos_rubro.lbl_titulo2 = mnu_RubrosIngresos.Caption
    fw_ingresos_rubro.Show
End Sub

Private Sub Mnu_saldosinicialesalmacenes_Click()
    aw_almacen_saldo_inicial.lbl_titulo = Mnu_saldosinicialesalmacenes.Caption
    aw_almacen_saldo_inicial.FraNavega = Mnu_saldosinicialesalmacenes.Caption
    aw_almacen_saldo_inicial.lbl_titulo2 = Mnu_saldosinicialesalmacenes.Caption
    aw_almacen_saldo_inicial.Show
End Sub

Private Sub mnu_SalidaAlmacen_Click()
    aw_salida_almacen_mant.lbl_titulo = mnu_SalidaAlmacen.Caption
    aw_salida_almacen_mant.FraNavega = mnu_SalidaAlmacen.Caption
    aw_salida_almacen_mant.lbl_titulo2 = mnu_SalidaAlmacen.Caption
    aw_salida_almacen_mant.Show
End Sub

Private Sub mnu_SalidaAlmacenHerr_Click()
    Aux = "UALMH"
    aw_almacen_salida.lbl_titulo = mnu_SalidaAlmacenHerr.Caption
    aw_almacen_salida.FraNavega = mnu_SalidaAlmacenHerr.Caption
    aw_almacen_salida.lbl_titulo2 = mnu_SalidaAlmacenHerr.Caption
    aw_almacen_salida.Show
'    Aux = "UALMH"
'    aw_almacen_salida_rep.lbl_titulo = mnu_SalidaAlmacenHerr.Caption
'    aw_almacen_salida_rep.FraNavega = mnu_SalidaAlmacenHerr.Caption
'    aw_almacen_salida_rep.lbl_titulo2 = mnu_SalidaAlmacenHerr.Caption
'    aw_almacen_salida_rep.Show
End Sub

Private Sub mnu_SalidaAlmacenOtro_Click()
    Aux = "UALMI"
    aw_almacen_salida.lbl_titulo = mnu_SalidaAlmacenOtro.Caption
    aw_almacen_salida.FraNavega = mnu_SalidaAlmacenOtro.Caption
    aw_almacen_salida.lbl_titulo2 = mnu_SalidaAlmacenOtro.Caption
    aw_almacen_salida.Show
    
'    aw_almacen_salida_rep.lbl_titulo = mnu_SalidaAlmacenOtro.Caption
'    aw_almacen_salida_rep.FraNavega = mnu_SalidaAlmacenOtro.Caption
'    aw_almacen_salida_rep.lbl_titulo2 = mnu_SalidaAlmacenOtro.Caption
'    aw_almacen_salida_rep.Show

End Sub

Private Sub mnu_SalidaAlmacenRep_Click()
    Aux = "UALMR"
'    aw_almacen_salida_rep.lbl_titulo = mnu_SalidaAlmacenRep.Caption
'    aw_almacen_salida_rep.FraNavega = mnu_SalidaAlmacenRep.Caption
'    aw_almacen_salida_rep.lbl_titulo2 = mnu_SalidaAlmacenRep.Caption
'    aw_almacen_salida_rep.Show

    aw_almacen_salida.lbl_titulo = mnu_SalidaAlmacenRep.Caption
    aw_almacen_salida.FraNavega = mnu_SalidaAlmacenRep.Caption
    aw_almacen_salida.lbl_titulo2 = mnu_SalidaAlmacenRep.Caption
    aw_almacen_salida.Show

End Sub

Private Sub mnu_SalirSistema_Click()
   Unload Me
   End
End Sub

Private Sub Mnu_SeguimientoCheques_Click()
    fw_conciliacion_bancaria.lbl_titulo = Mnu_SeguimientoCheques.Caption
    fw_conciliacion_bancaria.Show
End Sub

Private Sub Mnu_SeguimientoCobranzas_Click()
    'Timer1.Enabled = False
    Aux = "DCOBR"
    aw_seguimiento_cobranzas.lbl_titulo = Mnu_SeguimientoCobranzas.Caption
   '  aw_seguimiento_cobranzas.FraNavega = Mnu_SeguimientoCobranzas.Caption
'    aw_seguimiento_cobranzas.lbl_titulo2 = Mnu_SeguimientoCobranzas.Caption
    aw_seguimiento_cobranzas.Show
    
'    Frm_ao_ventas_seguimiento.lbl_titulo = Mnu_SeguimientoCobranzas.Caption
'    Frm_ao_ventas_seguimiento.FraNavega = Mnu_SeguimientoCobranzas.Caption
'    Frm_ao_ventas_seguimiento.lbl_titulo2 = Mnu_SeguimientoCobranzas.Caption
'    Frm_ao_ventas_seguimiento.Show
End Sub

Private Sub Mnu_SeguimientoInstalaciones_Click()
    Aux = "DVTA"
    mw_ventas_alcance_acta.lbl_titulo = Mnu_SeguimientoInstalaciones.Caption
    mw_ventas_alcance_acta.FraNavega = Mnu_SeguimientoInstalaciones.Caption
    mw_ventas_alcance_acta.lbl_titulo2 = Mnu_SeguimientoInstalaciones.Caption
    mw_ventas_alcance_acta.Show
End Sub

Private Sub Mnu_SeguimientoMantenimiento_Click()
    Aux = "DNMAN"
    tw_tecnico_bitacora.lbl_titulo = Mnu_SeguimientoMantenimiento.Caption
    tw_tecnico_bitacora.FraNavega = Mnu_SeguimientoMantenimiento.Caption
    tw_tecnico_bitacora.lbl_titulo2 = Mnu_SeguimientoMantenimiento.Caption
    tw_tecnico_bitacora.Show
End Sub

Private Sub Mnu_SeguimientoPago_Click()
    aw_seguimiento_comex.lbl_titulo = Mnu_SeguimientoPago.Caption
    aw_seguimiento_comex.FraNavega = Mnu_SeguimientoPago.Caption
    aw_seguimiento_comex.lbl_titulo2 = Mnu_SeguimientoPago.Caption
    aw_seguimiento_comex.Show
End Sub

Private Sub Mnu_SeguimientoReparacion_Click()
    Aux = "DNREP"
    tw_tecnico_bitacora.lbl_titulo = Mnu_SeguimientoReparacion.Caption
    tw_tecnico_bitacora.FraNavega = Mnu_SeguimientoReparacion.Caption
    tw_tecnico_bitacora.lbl_titulo2 = Mnu_SeguimientoReparacion.Caption
    tw_tecnico_bitacora.Show
End Sub

Private Sub Mnu_ServiciosBasicos_Click()
    Aux = "DCONT"
    fw_solicitud_servicio_basico.lbl_titulo = Mnu_ServiciosBasicos.Caption
    fw_solicitud_servicio_basico.FraNavega = Mnu_ServiciosBasicos.Caption
    fw_solicitud_servicio_basico.lbl_titulo2 = Mnu_ServiciosBasicos.Caption
    fw_solicitud_servicio_basico.Show
End Sub

'Private Sub Mnu_SistemaPuertas_Click()
'    aw_p_ac_bienes_equipo_sistema_puertas.lbl_titulo = Mnu_SistemaPuertas.Caption
'    aw_p_ac_bienes_equipo_sistema_puertas.FraNavega = Mnu_SistemaPuertas.Caption
'    aw_p_ac_bienes_equipo_sistema_puertas.lbl_titulo2 = Mnu_SistemaPuertas.Caption
'    aw_p_ac_bienes_equipo_sistema_puertas.Show
'End Sub

Private Sub Mnu_SolicitudCompra_Click()
    Aux = "DNINS"
    fw_compras_gral.lbl_titulo = Mnu_SolicitudCompra.Caption
    fw_compras_gral.FraNavega = Mnu_SolicitudCompra.Caption
    fw_compras_gral.lbl_titulo2 = Mnu_SolicitudCompra.Caption
    fw_compras_gral.Show
End Sub

Private Sub mnu_SolicitudContratacionPersonal_Click()
    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "EHALKYER" Or glusuario = "DOLMOS" Or glusuario = "DVARGAS" Or glusuario = "CSALINAS" Then
        FrmPostulantes.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_SolicitudEgresos_Click()
    Aux = "DCONT"
    fw_solicitud_compras.lbl_titulo = Mnu_SolicitudEgresos.Caption
    fw_solicitud_compras.FraNavega = Mnu_SolicitudEgresos.Caption
    fw_solicitud_compras.lbl_titulo2 = Mnu_SolicitudEgresos.Caption
    fw_solicitud_compras.Show
End Sub

Private Sub Mnu_SolicitudHerramientas_Click()
    Aux = "UALMH"
    frm_ao_requerimiento_compra.lbl_titulo = Mnu_solicitudRepuestos.Caption
    frm_ao_requerimiento_compra.FraNavega = Mnu_solicitudRepuestos.Caption
    frm_ao_requerimiento_compra.lbl_titulo2 = Mnu_solicitudRepuestos.Caption
    frm_ao_requerimiento_compra.Show
End Sub


'Private Sub Mnu_solicitudInsumos_Click()
'    Aux = "UALMI"
'    aw_requerimiento_compra.lbl_titulo = Mnu_solicitudInsumos.Caption
'    aw_requerimiento_compra.FraNavega = Mnu_solicitudInsumos.Caption
'    aw_requerimiento_compra.lbl_titulo2 = Mnu_solicitudInsumos.Caption
'    aw_requerimiento_compra.Show
'End Sub

Private Sub Mnu_solicitudRepuestos_Click()
    Aux = "UALMR"
    frm_ao_requerimiento_compra.lbl_titulo = Mnu_solicitudRepuestos.Caption
    frm_ao_requerimiento_compra.FraNavega = Mnu_solicitudRepuestos.Caption
    frm_ao_requerimiento_compra.lbl_titulo2 = Mnu_solicitudRepuestos.Caption
    frm_ao_requerimiento_compra.Show
End Sub

Private Sub Mnu_TareasCronoInstalacion_Click()
    If glusuario = "AURBINA" Or glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "CSALINAS" Or glusuario = "VPAREDES" Or glusuario = "CPAREDES" Or glusuario = "ACLAROS" Or glusuario = "AACOSTA" Or glusuario = "CURDININEA" Or glusuario = "FCABRERA" Or glusuario = "BRAMOS" Or glusuario = "JMAMANI" Or glusuario = "EVILLALOBOS" Or glusuario = "RMORA" Or glusuario = "BINFANTE" Or glusuario = "MRODRIGUEZ" Or glusuario = "AFLORES" Or glusuario = "RBUSTILLOS" Then
        tw_tareas_crono_instalacion.lbl_titulo = Mnu_TareasCronoInstalacion.Caption
        tw_tareas_crono_instalacion.FraNavega = Mnu_TareasCronoInstalacion.Caption
        tw_tareas_crono_instalacion.lbl_titulo2 = Mnu_TareasCronoInstalacion.Caption
        tw_tareas_crono_instalacion.Show
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Mnu_TiposImpuestos_Click()
    frm_gc_impuestos.lbl_titulo = Mnu_TiposImpuestos.Caption
    frm_gc_impuestos.FraNavega = Mnu_TiposImpuestos.Caption
    frm_gc_impuestos.lbl_titulo2 = Mnu_TiposImpuestos.Caption
    frm_gc_impuestos.Show
End Sub

Private Sub Mnu_TipoViasAcceso_Click()
    gw_p_gc_calle_tipo.lbl_titulo = Mnu_TipoViasAcceso.Caption
    gw_p_gc_calle_tipo.FraNavega = Mnu_TipoViasAcceso.Caption
    gw_p_gc_calle_tipo.lbl_titulo2 = Mnu_TipoViasAcceso.Caption
    gw_p_gc_calle_tipo.Show
End Sub

Private Sub Mnu_TipoVivienda_Click()
    gw_p_gc_edificacion_tipo.lbl_titulo = Mnu_TipoVivienda.Caption
    gw_p_gc_edificacion_tipo.FraNavega = Mnu_TipoVivienda.Caption
    gw_p_gc_edificacion_tipo.lbl_titulo2 = Mnu_TipoVivienda.Caption
    gw_p_gc_edificacion_tipo.Show
End Sub

Private Sub mnu_transferenciasAlmacenes_Click()
    Aux = "UALMI"
    aw_almacen_traspaso.lbl_titulo = mnu_transferenciasAlmacenes.Caption
    aw_almacen_traspaso.FraNavega = mnu_transferenciasAlmacenes.Caption
    aw_almacen_traspaso.lbl_titulo2 = mnu_transferenciasAlmacenes.Caption
    aw_almacen_traspaso.Show
End Sub

Private Sub Mnu_Transporte_Click()
    Aux = "COMEX"
    Glaux = "TRANS"
'    mw_opcion_importacion.Show modal
    fw_compras_comex.lbl_titulo = Mnu_Transporte.Caption
    fw_compras_comex.FraNavega = Mnu_Transporte.Caption
    fw_compras_comex.lbl_titulo2 = Mnu_Transporte.Caption
    fw_compras_comex.Show
End Sub

Private Sub Mnu_TraspasosEgresos_Click()
    Aux = "R-644"
    fw_traspaso_bancos_egresos.FraNavega = Mnu_TraspasosEgresos.Caption
    fw_traspaso_bancos_egresos.lbl_titulo = "DESTINO - " + Mnu_TraspasosEgresos.Caption
    fw_traspaso_bancos_egresos.Show
End Sub

Private Sub Mnu_UnidadesEjecutoras_Click()
    frm_gc_unidad_ejecutora.lbl_titulo = Mnu_UnidadesEjecutoras.Caption
    frm_gc_unidad_ejecutora.FraNavega = Mnu_UnidadesEjecutoras.Caption
    frm_gc_unidad_ejecutora.lbl_titulo2 = Mnu_UnidadesEjecutoras.Caption
    frm_gc_unidad_ejecutora.Show
End Sub

'Private Sub Mnu_VelocidadEquipo_Click()
'    aw_p_ac_bienes_equipo_velocidad.lbl_titulo = Mnu_VelocidadEquipo.Caption
'    aw_p_ac_bienes_equipo_velocidad.FraNavega = Mnu_VelocidadEquipo.Caption
'    aw_p_ac_bienes_equipo_velocidad.lbl_titulo2 = Mnu_VelocidadEquipo.Caption
'    aw_p_ac_bienes_equipo_velocidad.Show
'End Sub

Private Sub Mnu_ViasAcceso_Click()
    frm_gc_calles.lbl_titulo = Mnu_ViasAcceso.Caption
    frm_gc_calles.FraNavega = Mnu_ViasAcceso.Caption
    frm_gc_calles.lbl_titulo2 = Mnu_ViasAcceso.Caption
    frm_gc_calles.Show
End Sub

Private Sub Mnu_Vivienda_Click()
    gw_edificaciones.lbl_titulo = Mnu_Vivienda.Caption
    gw_edificaciones.FraNavega = Mnu_Vivienda.Caption
    gw_edificaciones.lbl_titulo2 = Mnu_Vivienda.Caption
    gw_edificaciones.Show
End Sub

Private Sub Mnu_zonas_Click()
    frm_gc_zonas.lbl_titulo = Mnu_zonas.Caption
    frm_gc_zonas.FraNavega = Mnu_zonas.Caption
    frm_gc_zonas.lbl_titulo2 = Mnu_zonas.Caption
    frm_gc_zonas.Show
End Sub

Private Sub mnuAcercade_Click()
'  frmAbout.Show vbModal
End Sub

Private Sub MnuActividades_Click()
    frm_po_actividad.lbl_titulo = MnuActividades.Caption
    frm_po_actividad.FraNavega = MnuActividades.Caption
    frm_po_actividad.lbl_titulo2 = MnuActividades.Caption
    frm_po_actividad.Show
End Sub

Private Sub MnuAlmacenesFisicos_Click()
    frm_ac_almacenes.lbl_titulo = MnuAlmacenesFisicos.Caption
    frm_ac_almacenes.FraNavega = MnuAlmacenesFisicos.Caption
    frm_ac_almacenes.lbl_titulo2 = MnuAlmacenesFisicos.Caption
    frm_ac_almacenes.Show
End Sub

Private Sub MnuAyuda_Click()
    Dim e As Long
    'e = Shell(App.Path & "\PERSONAL\MANUAL_USUARIO_GRAL.htm" & GlUsuario & " AYUDA ", 1)
    'e = Shell(App.Path & "\PERSONAL\MANUAL_USUARIO_GRAL.PDF")
End Sub

Private Sub MnuBienesServicios_Click()
    aw_bienes.lbl_titulo = MnuBienesServicios.Caption
    aw_bienes.FraNavega = MnuBienesServicios.Caption
    aw_bienes.lbl_titulo2 = MnuBienesServicios.Caption
    aw_bienes.Show
End Sub

Private Sub mnuCambiarClave_Click()
    FrmCambiarClave.Show vbModal
    'MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub MnuClasificacionBeneficiarios_Click()
    gw_p_gc_tipo_beneficiario.lbl_titulo = MnuClasificacionBeneficiarios.Caption
    gw_p_gc_tipo_beneficiario.FraNavega = MnuClasificacionBeneficiarios.Caption
    gw_p_gc_tipo_beneficiario.lbl_titulo2 = MnuClasificacionBeneficiarios.Caption
    gw_p_gc_tipo_beneficiario.Show
End Sub

Private Sub MnuClasificacionEspecifica_Click()
    pw_p_pc_proceso_nivel2.lbl_titulo = MnuClasificacionEspecifica.Caption
    pw_p_pc_proceso_nivel2.FraNavega = MnuClasificacionEspecifica.Caption
    pw_p_pc_proceso_nivel2.lbl_titulo2 = MnuClasificacionEspecifica.Caption
    pw_p_pc_proceso_nivel2.Show
End Sub

Private Sub MnuCotizacionesEquipos_Click()
    Aux = "DVTA"
'    frm_ao_solicitud_cotiza_venta.lbl_titulo = MnuCotizacionesEquipos.Caption
'    frm_ao_solicitud_cotiza_venta.FraNavega0 = MnuCotizacionesEquipos.Caption
''    mw_solicitud_cotiza_venta.lbl_titulo2 = MnuCotizacionesEquipos.Caption
'    frm_ao_solicitud_cotiza_venta.Show
    
    mw_solicitud_cotiza_venta.lbl_titulo = MnuCotizacionesEquipos.Caption
    mw_solicitud_cotiza_venta.FraNavega0 = MnuCotizacionesEquipos.Caption
'    mw_solicitud_cotiza_venta.lbl_titulo2 = MnuCotizacionesEquipos.Caption
    mw_solicitud_cotiza_venta.Show
End Sub

Private Sub MnuEquipos_Click()
'    frm_ac_bienes_eqp.lbl_titulo = MnuEquipos.Caption
'    frm_ac_bienes_eqp.FraNavega = MnuEquipos.Caption
'    frm_ac_bienes_eqp.lbl_titulo2 = MnuEquipos.Caption
'    frm_ac_bienes_eqp.Show
End Sub

Private Sub MnuEtapas_Click()
    frm_pc_proceso_nivel3.lbl_titulo = MnuEtapas.Caption
    frm_pc_proceso_nivel3.FraNavega = MnuEtapas.Caption
    frm_pc_proceso_nivel3.lbl_titulo2 = MnuEtapas.Caption
    frm_pc_proceso_nivel3.Show
End Sub

Private Sub MnuFormulacionPresupuestaria_Click()
    Frm_fo_ppto.lbl_titulo = MnuFormulacionPresupuestaria.Caption
    Frm_fo_ppto.FraNavega = MnuFormulacionPresupuestaria.Caption
    Frm_fo_ppto.lbl_titulo2 = MnuFormulacionPresupuestaria.Caption
    Frm_fo_ppto.Show
End Sub

Private Sub MnuGruposBienes_Click()
    frm_ac_bienes_grupos.lbl_titulo = MnuGruposBienes.Caption
    frm_ac_bienes_grupos.FraNavega = MnuGruposBienes.Caption
    frm_ac_bienes_grupos.lbl_titulo2 = MnuGruposBienes.Caption
    frm_ac_bienes_grupos.Show
End Sub

'Private Sub mnuEjecIng_Click()
'
'  Dim rsv_Ingreso_Convenio2 As New ADODB.Recordset
'  Dim rsIG_Ing_EjePptoConvenio As New ADODB.Recordset
'  Dim consulta1 As String
'  consulta1 = ""
'  ' ANTES CON FECHAS consulta1 = "where (fecha_registro >= '" & DTPkFechaInicio & "' and fecha_registro <= '" & DTPkFechaFin & "') "
''  If Len(Trim(DtCcodigo_convenio.Text)) > 0 Then
'    'ANTES CON FECHAS
'    'consulta1 = consulta1 & " and codigo_convenio = '" & Trim(DtCcodigo_convenio.Text) & "' "
'    ' AHORA SOLO CONVENIO
''    consulta1 = " WHERE codigo_convenio = '" & Trim(DtCcodigo_convenio.Text) & "' "
''  Else
''  End If
'  Set rsv_Ingreso_Convenio2 = New ADODB.Recordset
'  If rsv_Ingreso_Convenio2.State = 1 Then rsv_Ingreso_Convenio2.Close
'  rsv_Ingreso_Convenio2.Open "select * from v_Ingreso_Convenio2 " & consulta1 & " order by codigo_convenio", db, adOpenKeyset, adLockReadOnly
'  Print rsv_Ingreso_Convenio2.RecordCount
'  Set rsIG_Ing_EjePptoConvenio = New ADODB.Recordset
'  db.Execute "DELETE FROM IG_Ing_EjePptoConvenio WHERE maquina = '" & GlMaquina & "'"
'  Set rsIG_Ing_EjePptoConvenio = New ADODB.Recordset
'  rsIG_Ing_EjePptoConvenio.Open "select * from IG_Ing_EjePptoConvenio where maquina = '" & GlMaquina & "'", db, adOpenKeyset, adLockOptimistic
'  While Not rsv_Ingreso_Convenio2.EOF
'    rsIG_Ing_EjePptoConvenio.AddNew
'    rsIG_Ing_EjePptoConvenio!codigo_convenio = rsv_Ingreso_Convenio2!codigo_convenio
'    rsIG_Ing_EjePptoConvenio!Cta_Codigo = rsv_Ingreso_Convenio2!Cta_Codigo
'    rsIG_Ing_EjePptoConvenio!org_codigo = rsv_Ingreso_Convenio2!org_codigo
'    rsIG_Ing_EjePptoConvenio!estado_recaudado = rsv_Ingreso_Convenio2!estado_recaudado
'    rsIG_Ing_EjePptoConvenio!monto_dolares = rsv_Ingreso_Convenio2!monto_dolares
'    rsIG_Ing_EjePptoConvenio!monto_Bolivianos = rsv_Ingreso_Convenio2!monto_Bolivianos
'    rsIG_Ing_EjePptoConvenio!monto_formulado_us = rsv_Ingreso_Convenio2!monto_formulado_us
'    rsIG_Ing_EjePptoConvenio!monto_vigente_us = rsv_Ingreso_Convenio2!monto_vigente_us
'    rsIG_Ing_EjePptoConvenio!monto_compromiso_us = rsv_Ingreso_Convenio2!monto_compromiso_us
'    rsIG_Ing_EjePptoConvenio!monto_devengado_us = rsv_Ingreso_Convenio2!monto_devengado_us
'    rsIG_Ing_EjePptoConvenio!monto_pagado_us = rsv_Ingreso_Convenio2!monto_pagado_us
'    rsIG_Ing_EjePptoConvenio!codigo_convenio = rsv_Ingreso_Convenio2!codigo_convenio
'    rsIG_Ing_EjePptoConvenio!org_codigo_cta = rsv_Ingreso_Convenio2!org_codigo_cta
'    rsIG_Ing_EjePptoConvenio!maquina = GlMaquina
'    rsIG_Ing_EjePptoConvenio.Update
'    rsv_Ingreso_Convenio2.MoveNext
'  Wend
'
'  Cry.WindowShowRefreshBtn = True
'  Cry.ReportFileName = App.Path & "\InfGerencial\Ingresos\Rpt_Ingresos_convenio3.rpt"
''  Cry.ReportFileName = App.path & "\Rpt_Ingresos_convenio3.rpt"
'
'  Cry.WindowShowPrintBtn = True
'  Cry.WindowShowExportBtn = True
'  Cry.WindowShowPrintSetupBtn = True
'  Cry.WindowShowGroupTree = True
'  Cry.WindowState = crptMaximized
'  iResult = Cry.PrintReport
'  If iResult <> 0 Then
'      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
'  End If
'
'End Sub

Private Sub MnuIdentificacionCliente_Click()
    Aux = "DVTA"
    mw_solicitud.lbl_titulo = MnuIdentificacionCliente.Caption
    mw_solicitud.FraNavega = MnuIdentificacionCliente.Caption
    mw_solicitud.lbl_titulo2 = MnuIdentificacionCliente.Caption
    mw_solicitud.Show
End Sub

Private Sub mnuNivelAcceso_Click()
    FrmNivelesAcceso.Show
'    MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub MnuInsumos_Click()
    frm_po_insumo.lbl_titulo = MnuInsumos.Caption
    frm_po_insumo.FraNavega = MnuInsumos.Caption
    frm_po_insumo.lbl_titulo2 = MnuInsumos.Caption
    frm_po_insumo.Show
End Sub

'Private Sub MnuMarcas_Click()
'    aw_p_ac_bienes_marcas.lbl_titulo = MnuMarcas.Caption
'    aw_p_ac_bienes_marcas.FraNavega = MnuMarcas.Caption
'    aw_p_ac_bienes_marcas.lbl_titulo2 = MnuMarcas.Caption
'    aw_p_ac_bienes_marcas.Show
'End Sub

Private Sub MnuObjetivoEspecifico_Click()
    frm_po_objetivo_especifico.lbl_titulo = MnuObjetivoEspecifico.Caption
    frm_po_objetivo_especifico.FraNavega = MnuObjetivoEspecifico.Caption
    frm_po_objetivo_especifico.lbl_titulo2 = MnuObjetivoEspecifico.Caption
    frm_po_objetivo_especifico.Show
End Sub

Private Sub MnuObjetivoGeneral_Click()
    frm_po_objetivo_general.lbl_titulo = MnuObjetivoGeneral.Caption
    frm_po_objetivo_general.FraNavega = MnuObjetivoGeneral.Caption
    frm_po_objetivo_general.lbl_titulo2 = MnuObjetivoGeneral.Caption
    frm_po_objetivo_general.Show
End Sub

Private Sub MnuPlanConsolidado_Click()
    frm_po_poa.lbl_titulo = MnuPlanConsolidado.Caption
    frm_po_poa.FraNavega = MnuPlanConsolidado.Caption
    frm_po_poa.lbl_titulo2 = MnuPlanConsolidado.Caption
    frm_po_poa.Show
End Sub

Private Sub MnuProfesionesOcupaciones_Click()
    gw_p_gc_ocupacion_profesion.lbl_titulo = MnuProfesionesOcupaciones.Caption
    gw_p_gc_ocupacion_profesion.FraNavega = MnuProfesionesOcupaciones.Caption
    gw_p_gc_ocupacion_profesion.lbl_titulo2 = MnuProfesionesOcupaciones.Caption
    gw_p_gc_ocupacion_profesion.Show
End Sub

Private Sub MnuRegistroEmpresas_Click()
    gw_p_gc_beneficiario_empresa.lbl_titulo = MnuRegistroEmpresas.Caption
    gw_p_gc_beneficiario_empresa.FraNavega = MnuRegistroEmpresas.Caption
    gw_p_gc_beneficiario_empresa.lbl_titulo2 = MnuRegistroEmpresas.Caption
    gw_p_gc_beneficiario_empresa.Show
End Sub

Private Sub mnuRepBalApertura_Click()
   cc_balapertura.Show
End Sub

Private Sub MnuRepVentas_Click()
    frmVentasReportes.Show
    'MsgBox "Error en despliegue de pantalla, consulte con el Administrador del Sistema"
End Sub

Private Sub mnuRgistroPersonas_Click()
    Glaux = "1"
    gw_p_gc_beneficiario_persona.lbl_titulo = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.FraNavega = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.lbl_titulo2 = mnuRgistroPersonas.Caption
    gw_p_gc_beneficiario_persona.Show
End Sub

Private Sub MnuSolicitudCotizacionVenta_Click()
    Aux = "DVTA"
    mw_solicitud.lbl_titulo = MnuSolicitudCotizacionVenta.Caption
    mw_solicitud.FraNavega = MnuSolicitudCotizacionVenta.Caption
    mw_solicitud.lbl_titulo2 = MnuSolicitudCotizacionVenta.Caption
    mw_solicitud.Show
End Sub

Private Sub MnuSubgrupoBienes_Click()
    frm_ac_bienes_subgrupo.lbl_titulo = MnuSubgrupoBienes.Caption
    frm_ac_bienes_subgrupo.FraNavega = MnuSubgrupoBienes.Caption
    frm_ac_bienes_subgrupo.lbl_titulo2 = MnuSubgrupoBienes.Caption
    frm_ac_bienes_subgrupo.Show
End Sub

Private Sub MnuTareas_Click()
    frm_po_tarea.lbl_titulo = MnuTareas.Caption
    frm_po_tarea.FraNavega = MnuTareas.Caption
    frm_po_tarea.lbl_titulo2 = MnuTareas.Caption
    frm_po_tarea.Show
End Sub

Private Sub MnuTraspasosPresupuestarios_Click()
    FrmModPresup.lbl_titulo = MnuFormulacionPresupuestaria.Caption
    FrmModPresup.FraNavega = MnuFormulacionPresupuestaria.Caption
    FrmModPresup.lbl_titulo2 = MnuFormulacionPresupuestaria.Caption
    FrmModPresup.Show
End Sub

Private Sub MnuUnidadesMedida_Click()
    aw_p_ac_bienes_unidad_medida.lbl_titulo = MnuUnidadesMedida.Caption
    aw_p_ac_bienes_unidad_medida.FraNavega = MnuUnidadesMedida.Caption
    aw_p_ac_bienes_unidad_medida.lbl_titulo2 = MnuUnidadesMedida.Caption
    aw_p_ac_bienes_unidad_medida.Show
End Sub

Private Sub mnuUsuarios_Click()
    FrmSisUsuarios.Show
    'MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub MDIForm_Load()
    ' Esta función no elimina el botón "X" de un MDI, pero sí lo deja inactivo
    Dim hSysmenu As Long
    hSysmenu = GetSystemMenu(Me.hwnd, 0)
    RemoveMenu hSysmenu, 6, &H400&

   Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
   Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
   Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
   Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)

   With sbStatusBar
'    'Agregamos el Panel1 y mostramos la hora con sbrTime (propiedad style)
'    .Panels.Add , "Hora", , sbrTime
    ''Agregamos el Panel2 y mostramos la Fecha con sbrdate (propiedad style)
    '.Panels.Add , "Fecha", , sbrDate
'    'Agregamos el Panel3 y mostramos un texto cualquiera con una imagen
'    .Panels.Add , "Impresion", "Imprimiendo trabajo.....", sbrText, LoadPicture(App.Path & "\imagen1.ico")
'    .Panels.Add , "BD", GlBaseDatos     '"ADMIN_EMPRESA"
'    .Panels.Add , "Usuario", glusuario
   End With

   txtUsuarioGl.Caption = glusuario
   txtBDGl.Caption = GlBaseDatos
   txtFechaGl.Caption = Format(Date, "dd/mm/yyyy")
   txtHoraGl.Caption = Format(Time, "hh:mm:ss")
   txtVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
   If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "MPAREDES" Or glusuario = "APALACIOS" Or glusuario = "JCHIPANA" Or glusuario = "VBELLIDO" Then
        CmdRepGral.Visible = True
   Else
        CmdRepGral.Visible = False
   End If
'   LoadNewDoc
    Call SeguridadSet(Me)
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

Private Sub mnuPrivAcceso_Click()
    frmPrivAcceso.Show
End Sub

'Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'   On Error Resume Next
'   Select Case Button.Key
'      Case "Cliente"
'         frmBeneficiario.Show
'      Case "Producto"
'         AlFrmCreaMaterial.Show
'      Case "Guardar"
'         'TareasPendientes: Agregar código de botón 'Guardar'.
'         MsgBox "Agregar código de botón 'Guardar'."
'      Case "Imprimir"
'         'TareasPendientes: Agregar código de botón 'Imprimir'.
'         MsgBox "Agregar código de botón 'Imprimir'."
'      Case "Cortar"
'         'TareasPendientes: Agregar código de botón 'Cortar'.
'         MsgBox "Agregar código de botón 'Cortar'."
'      Case "Copiar"
'         'TareasPendientes: Agregar código de botón 'Copiar'.
'         MsgBox "Agregar código de botón 'Copiar'."
'      Case "Pegar"
'         'TareasPendientes: Agregar código de botón 'Pegar'.
'         MsgBox "Agregar código de botón 'Pegar'."
'      Case "Negrita"
'         ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
'         Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
'      Case "Cursiva"
'         ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
'         Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
'      Case "Subrayado"
'         ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
'         Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
'      Case "Alinear a la izquierda"
'         ActiveForm.rtfText.SelAlignment = rtfLeft
'      Case "Centrar"
'         ActiveForm.rtfText.SelAlignment = rtfCenter
'      Case "Alinear a la derecha"
'         ActiveForm.rtfText.SelAlignment = rtfRight
'   End Select
'End Sub

'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'   On Error Resume Next
'   Select Case Button.Key
'      Case "Producto"
'         AlFrmCreaMaterial.Show
'      Case "cliente"
'         frmBeneficiario.Show
'      Case "SolCompra"
'         frmBeneficiarioEmp.Show
'      Case "Compra"
'         frmComprasDirectas.Show
'      Case "Pago"
'         FrmPagosTotal.Show
'      Case "Venta"
'         FrmVentas.Show
'      Case "ReporteV"
'         frmVentasReportes.Show
'      Case "ReporteC"
'         frmComprasReportes.Show
'         'ActiveForm.rtfText.SelAlignment = rtfRight
'   End Select
'End Sub

Public Sub NivelAcceso(vNivelAcceso As Integer)
'Subrutina que habilita o deshabilita las opciones de menu
On Error Resume Next
Dim vNombOpcMenu As String

'INI WWWWWWWWWWWWWWWWWWWWWWWWW
rsNivelAcceso.Open "Select * From gc_acceso_asignacion Where IdNivelAcceso=" & vNivelAcceso, db, adOpenStatic
If rsNivelAcceso.RecordCount > 0 Then
    rsNivelAcceso.MoveFirst
    While Not rsNivelAcceso.EOF
        Set rs_datos1 = New ADODB.Recordset
        If rs_datos1.State = 1 Then rs_datos1.Close
        rs_datos1.Open "Select * from gc_menu_sistema where menu_codigo = '" & rsNivelAcceso!menu_codigo & "' ", db, adOpenStatic
        If rs_datos1.RecordCount > 0 Then
            'vNombOpcMenu = LCase(rs_datos1!menu_name)
            vNombOpcMenu = LTrim(rs_datos1!menu_name)
        Else
            vNombOpcMenu = "10000"
        End If
        If vNombOpcMenu = "Mnu_Planificacion" Then Mnu_Planificacion.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                                'Planificacion
            If vNombOpcMenu = "Mnu_ClasificadoresGral" Then Mnu_ClasificadoresGral.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                      'ClasificadoresGral
            
        If vNombOpcMenu = "MnuProcesosAdministrativos" Then MnuProcesosAdministrativos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)
            If vNombOpcMenu = "MnuClasificadoresAdministrativos" Then MnuClasificadoresAdministrativos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)  'ClasificadoresAdministrativos
            
            If vNombOpcMenu = "MnuComercialComex" Then MnuComercialComex.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'Comercial
                If vNombOpcMenu = "MnuIdentificacionCliente" Then MnuIdentificacionCliente.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)              'IdentificacionCliente
                If vNombOpcMenu = "Mnu_ParametrosCalculo" Then Mnu_ParametrosCalculo.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                    'Mnu_ParametrosCalculo
                If vNombOpcMenu = "MnuCotizacionesEquipos" Then MnuCotizacionesEquipos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                  'MnuCotizacionesEquipos
                If vNombOpcMenu = "Mnu_ProcesoVentas" Then Mnu_ProcesoVentas.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'ProcesoVentas Nuevas
                If vNombOpcMenu = "mnu_instalaciones" Then mnu_instalaciones.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'instalaciones
                
            If vNombOpcMenu = "Mnu_ImportacionEquipos" Then Mnu_ImportacionEquipos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                  'ImportacionEquipos COMEX
            
            If vNombOpcMenu = "mnu_ActivosFijos" Then mnu_ActivosFijos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                              'ActivosFijos
    
        If vNombOpcMenu = "mnu_gerenciaTecnica" Then mnu_gerenciaTecnica.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                        'gerenciaTecnica
            If vNombOpcMenu = "Mnu_ClasificadoresAreaTecnica" Then Mnu_ClasificadoresAreaTecnica.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)    'ClasificadoresAreaTecnica
            
            If vNombOpcMenu = "mnu_mantenimiento" Then mnu_mantenimiento.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'mantenimiento
            
            If vNombOpcMenu = "mnu_reparaciones" Then mnu_reparaciones.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                              'reparaciones
            
            If vNombOpcMenu = "mnu_emergencias" Then mnu_emergencias.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                                'emergencias
            
            If vNombOpcMenu = "mnu_modernizacion" Then mnu_modernizacion.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'modernizacion
            
            If vNombOpcMenu = "Mnu_AlmacenInsumos" Then Mnu_AlmacenInsumos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                          'AlmacenInsumos
            
            If vNombOpcMenu = "Mnu_AlmacenRepuestos" Then Mnu_AlmacenRepuestos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                      'AlmacenRepuestos
            
            If vNombOpcMenu = "Mnu_AlmacenHerramientas" Then Mnu_AlmacenHerramientas.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                'AlmacenHerramientas
            
        If vNombOpcMenu = "Mnu_RRHH" Then Mnu_RRHH.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                                              'RRHH
            If vNombOpcMenu = "Mnu_ClasificadoresRRHH" Then Mnu_ClasificadoresRRHH.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                  'ClasificadoresRRHH

            If vNombOpcMenu = "mnu_AdministracionPersonal" Then mnu_AdministracionPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)          'AdministracionPersonal
            
            If vNombOpcMenu = "Mnu_PlanillasPagosPersonal" Then Mnu_PlanillasPagosPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)          'PlanillasPagosPersonal
            
            If vNombOpcMenu = "Mnu_Reportes_RRHH" Then Mnu_Reportes_RRHH.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'Reportes_RRHH
        
        If vNombOpcMenu = "Mnu_ProcesosFinancierso" Then Mnu_ProcesosFinancierso.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                'ProcesosFinancierso
            If vNombOpcMenu = "mnu_ClasificadoresFinancieros" Then mnu_ClasificadoresFinancieros.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)    'ClasificadoresFinancieros
            If vNombOpcMenu = "Mnu_EjecucionIngresos" Then Mnu_EjecucionIngresos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                    'EjecucionIngreso
            
            If vNombOpcMenu = "mnu_EjecucionGasto" Then mnu_EjecucionGasto.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                          'EjecucionGasto
            
            If vNombOpcMenu = "mnu_Contabilidad" Then mnu_Contabilidad.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                              'Contabilidad
            
            If vNombOpcMenu = "Mnu_FaturacionCobranza" Then Mnu_FaturacionCobranza.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                  'FaturacionCobranza
                If vNombOpcMenu = "Mnu_FacturacionAntes" Then Mnu_FacturacionAntes.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                  'Mnu_FacturacionAntes
                If vNombOpcMenu = "Mnu_Facturacion" Then Mnu_Facturacion.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'Facturacion
                If vNombOpcMenu = "Mnu_Cobranzas" Then Mnu_Cobranzas.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                                'Cobranzas  'Mnu_Cobranzas
                If vNombOpcMenu = "Mnu_SeguimientoCobranzas" Then Mnu_SeguimientoCobranzas.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)          'SeguimientoCobranzas
            If vNombOpcMenu = "mnu_Tesoreria" Then mnu_Tesoreria.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                                    'Tesoreria
                If vNombOpcMenu = "Mnu_RegistroPagos" Then Mnu_RegistroPagos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                            'RegistroPagos
                If vNombOpcMenu = "mnu_NotaCreditoDebito" Then mnu_NotaCreditoDebito.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                    'NotaCreditoDebito
            
            If vNombOpcMenu = "Mnu_Descargos" Then Mnu_Descargos.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                                    'Descargos
            
            If vNombOpcMenu = "Mnu_ReportesCobranzas" Then Mnu_ReportesCobranzas.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)                    'ReportesCobranzas

        If vNombOpcMenu = "Mnu_InformacionGerencial" Then Mnu_InformacionGerencial.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)              'InformacionGerencial
        
        If vNombOpcMenu = "mnu_AdministracionSistema" Then mnu_AdministracionSistema.Enabled = IIf(rsNivelAcceso!habilitado = "SI", True, False)            'AdministracionSistema
        rsNivelAcceso.MoveNext
    Wend
    rsNivelAcceso.MoveFirst
End If

    
'    If vNombOpcMenu = "clasificadores" Then Clasificadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "Mnu_ClasificadoresGral" Then Mnu_ClasificadoresGral.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)  'Presupuesto
'                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "departamentosbolivia" Then DepartamentosBolivia.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "provinciasdepartamentos" Then ProvinciasDepartamentos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "tiposerrores" Then TiposErrores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)  'Presupuesto
'                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "economicosgasto" Then economicosgasto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "RelacionadorGastoEco" Then RelacionadorGastoEco.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "presupuesto" Then Presupuesto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)   'Presupuesto
'                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "sisin" Then Sisin.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "deducciones" Then Deducciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadordonacionesorganismos" Then RelacionadorDonacionesOrganismos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
''             If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)  'Presupuesto
''                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "distribucionporcentual" Then DistribucionPorcentual.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'            If vNombOpcMenu = "ctesoreria" Then CTesoreria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "cbancos" Then CBancos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "cctabancarias" Then CCtaBancarias.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formaspago" Then FormasPago.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "ingresos" Then Ingresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "rubros" Then Rubros.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadorrubroeco" Then RelacionadorRubroEco.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "contabilidad2" Then Contabilidad2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "plancuentas" Then PlanCuentas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadorcuentapartidas" Then RelacionadorCuentaPartidas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadoringresoscuentas" Then RelacionadorIngresosCuentas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "depreciaciones" Then Depreciaciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "claseauxiliares" Then ClaseAuxiliares.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "inversiones" Then Inversiones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "administrativos" Then Administrativos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "adquisiciones" Then Adquisiciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "contrataciones" Then Contrataciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "almacenes2" Then Almacenes2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "recursoshumanos2" Then RecursosHumanos2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "mesaentrada" Then MesaEntrada.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)      'Contabilidad
'            If vNombOpcMenu = "registrosolicitudes" Then RegistroSolicitudes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof01" Then FormularioF01.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof02" Then FormularioF02.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof03" Then FormularioF03.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof04" Then FormularioF04.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof05" Then FormularioF05.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof06" Then FormularioF06.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof07" Then FormularioF07.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "copiarestauracion" Then CopiaRestauracion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "progconadq" Then ProgConAdq.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "procesos" Then Procesos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)          'Egresos
'        If vNombOpcMenu = "compromiso" Then Compromiso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Ejecucion Presupuestaria
'        If vNombOpcMenu = "ejecucion" Then Ejecucion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Reportes de Ejecucion
'            '''ALB If vNombOpcMenu = "ejecucionppto" Then EjecucionPpto.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            If vNombOpcMenu = "mnuejecucionpoa" Then mnuEjecucionPOA.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            '''ALB  If vNombOpcMenu = "repgraf" Then RepGraf.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            '''ALB  If vNombOpcMenu = "repgrafunidad" Then repGrafUnidad.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'                If vNombOpcMenu = "repgraforga" Then repGraforga.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            '''ALB If vNombOpcMenu = "mnuejecucionpresupuestaria" Then mnuEjecucionPresupuestaria.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'                If vNombOpcMenu = "ejecucionporuni" Then EjecucionPorUni.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnuconvenio" Then mnuConvenio.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnuorganismo" Then mnuOrganismo.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnucategoria" Then mnuCategoria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnuproyecto" Then mnuProyecto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
''        If vNombOpcMenu = "ejecucion" Then Ejecucion.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)     'Reportes de Ejecucion
''            If vNombOpcMenu = "ejecucionorganismo" Then EjecucionOrganismo.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecucioncomprobante" Then EjecucionComprobante.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecucioncompromiso" Then EjecucionCompromiso.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecucióndevengado" Then EjecuciónDevengado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecuciónpagado" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'        If vNombOpcMenu = "mnumodppto" Then MnuModPpto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Modificaciones Presupuestarias
'
'    If vNombOpcMenu = "tesoreria" Then Tesoreria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "pp" Then pp.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Pagos pendientes
'        If vNombOpcMenu = "pagosefectuados" Then PagosEfectuados.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Activacion de Cheques
'        If vNombOpcMenu = "cuentasbancarias2" Then CuentasBancarias2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Desactivacion de cheques
'        If vNombOpcMenu = "manejocheques" Then ManejoCheques.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Trasnferencia
'        If vNombOpcMenu = "ic" Then ic.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
'        If vNombOpcMenu = "ct" Then ct.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
'            If vNombOpcMenu = "mnuimpcheques" Then MnuImpCheques.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Cheques
'            If vNombOpcMenu = "cuentasbancarias" Then CuentasBancarias.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Cheques
'
'    If vNombOpcMenu = "ccontabilidad" Then CContabilidad.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
'            If vNombOpcMenu = "comprobantes" Then Comprobantes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "reportesc" Then ReportesC.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "libromayor" Then LibroMayor.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "libromayorauxiliar" Then LibroMayorAuxiliar.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "balancegeneral" Then BalanceGeneral.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "balancesumassaldos" Then BalanceSumasSaldos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "estadoresultados" Then EstadoResultados.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "mnuingresos" Then MnuIngresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)          'Egresos
'        If vNombOpcMenu = "ejecucionpresupuestaria" Then EjecucionPresupuestaria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Ejecucion Presupuestaria
'        If vNombOpcMenu = "reportesingresos" Then ReportesIngresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Reportes de Ejecucion
'
'''''    If vNombOpcMenu = "administracion" Then Administracion.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''        If vNombOpcMenu = "adquisicionbienes" Then AdquisicionBienes.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "comprasdirectas" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "licitacionesnacionales" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "licitacionesinternacionales" Then LicitacionesInternacionales.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''        If vNombOpcMenu = "contratacionservicios" Then ContratacionServicios.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "consultoresindividuales" Then ConsultoresIndividuales.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "empresasconsultoras" Then EmpresasConsultoras.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''        If vNombOpcMenu = "almacenes" Then Almacenes.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "ingresosa" Then IngresosA.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "salidasa" Then SalidasA.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'
'    If vNombOpcMenu = "administracion" Then Administracion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "mnuadmicont" Then MnuAdmiCont.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "almacenes" Then Almacenes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "ingresosa" Then IngresosA.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnuingreso" Then mnuIngreso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnucompras" Then MnuCompras.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnusalidas" Then mnuSalidas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnucontrol" Then MnuControl.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnuinventario" Then mnuInventario.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnuestado" Then mnuestado.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "mnulicitac" Then mnuLicitac.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnusolnoobjec" Then MnuSolNoObjec.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnupublicacion" Then mnuPublicacion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnucomision" Then mnucomision.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuventapliegos" Then mnuVentaPliegos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnurecepsobres" Then MnuRecepSobres.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuabogado" Then mnuabogado.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuadjudica" Then mnuAdjudica.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuorden" Then mnuorden.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "mnuconsultoria_c" Then mnuConsultoria_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnusolnoobj_c" Then mnuSolNoObj_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnupublicacion_c" Then mnuPublicacion_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuventapliegos_c" Then mnuVentaPliegos_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnurecepcionprop_c" Then mnuRecepcionProp_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuaperturaprop_c" Then mnuAperturaProp_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuadjudicacion_c" Then mnuAdjudicacion_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnugespagos" Then mnugesPagos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuoc" Then mnuOC.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "comprasdirectas" Then ComprasDirectas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "solicitudnoobjecioncd" Then solicitudnoobjecioncd.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "cotizaciones" Then Cotizaciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "cuadrocomparativo" Then cuadrocomparativo.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "ordenpagocd" Then OrdenPagoCD.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        'If vNombOpcMenu = "progconadq" Then ProgConAdq.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "mnusolicitudesf04" Then mnuSolicitudesF04.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "recursoshumanos" Then RecursosHumanos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
'        If vNombOpcMenu = "administracionpersonal" Then AdministracionPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "controlpersonal" Then ControlPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "capacitacionpersonal" Then CapacitacionPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "evaluaciondesempeño" Then EvaluacionDesempeño.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "informaciongerencial" Then InformacionGerencial.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
'
'    If vNombOpcMenu = "mnuadmisistema" Then mnuAdmiSistema.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Administracion del sistema
'        If vNombOpcMenu = "mnucambiarclave" Then mnuCambiarClave.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Cambiar clave
'        If vNombOpcMenu = "mnuusuarios" Then mnuUsuarios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Definicion de Usuarios
'        If vNombOpcMenu = "mnunivelacceso" Then mnuNivelAcceso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Cambiar clave
'        If vNombOpcMenu = "mnuprivacceso" Then mnuPrivAcceso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Privilegios de Operación
'    rsNivelAcceso.MoveNext
'    Wend
'    rsNivelAcceso.MoveFirst
'End If

'FIN WWWWWWWWWWWWWWWWWWWWWWWWW
'rsNivelAcceso.Open "Select * From gc_nivelacceso Where IdNivelAcceso=" & vNivelAcceso, db, adOpenStatic
'If rsNivelAcceso.RecordCount > 0 Then
'    While Not rsNivelAcceso.EOF
'    vNombOpcMenu = LCase(rsNivelAcceso!NombOpcMenu)
'    If vNombOpcMenu = "clasificadores" Then Clasificadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "Mnu_ClasificadoresGral" Then Mnu_ClasificadoresGral.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)  'Presupuesto
'                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "departamentosbolivia" Then DepartamentosBolivia.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "provinciasdepartamentos" Then ProvinciasDepartamentos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "tiposerrores" Then TiposErrores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)  'Presupuesto
'                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "economicosgasto" Then economicosgasto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "RelacionadorGastoEco" Then RelacionadorGastoEco.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "presupuesto" Then Presupuesto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)   'Presupuesto
'                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "sisin" Then Sisin.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "deducciones" Then Deducciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadordonacionesorganismos" Then RelacionadorDonacionesOrganismos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
''             If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)  'Presupuesto
''                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "distribucionporcentual" Then DistribucionPorcentual.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
''                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'            If vNombOpcMenu = "ctesoreria" Then CTesoreria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "cbancos" Then CBancos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "cctabancarias" Then CCtaBancarias.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formaspago" Then FormasPago.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "ingresos" Then Ingresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "rubros" Then Rubros.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadorrubroeco" Then RelacionadorRubroEco.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "contabilidad2" Then Contabilidad2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "plancuentas" Then PlanCuentas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadorcuentapartidas" Then RelacionadorCuentaPartidas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "relacionadoringresoscuentas" Then RelacionadorIngresosCuentas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "depreciaciones" Then Depreciaciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "claseauxiliares" Then ClaseAuxiliares.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "inversiones" Then Inversiones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "administrativos" Then Administrativos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "adquisiciones" Then Adquisiciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "contrataciones" Then Contrataciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "almacenes2" Then Almacenes2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "recursoshumanos2" Then RecursosHumanos2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Tesoreria
'                If vNombOpcMenu = "mesaentrada" Then MesaEntrada.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)      'Contabilidad
'            If vNombOpcMenu = "registrosolicitudes" Then RegistroSolicitudes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof01" Then FormularioF01.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof02" Then FormularioF02.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof03" Then FormularioF03.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof04" Then FormularioF04.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof05" Then FormularioF05.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof06" Then FormularioF06.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "formulariof07" Then FormularioF07.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "copiarestauracion" Then CopiaRestauracion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "progconadq" Then ProgConAdq.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "procesos" Then Procesos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)          'Egresos
'        If vNombOpcMenu = "compromiso" Then Compromiso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Ejecucion Presupuestaria
'        If vNombOpcMenu = "ejecucion" Then Ejecucion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Reportes de Ejecucion
'            '''ALB If vNombOpcMenu = "ejecucionppto" Then EjecucionPpto.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            If vNombOpcMenu = "mnuejecucionpoa" Then mnuEjecucionPOA.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            '''ALB  If vNombOpcMenu = "repgraf" Then RepGraf.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'            '''ALB  If vNombOpcMenu = "repgrafunidad" Then repGrafUnidad.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'                If vNombOpcMenu = "repgraforga" Then repGraforga.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            '''ALB If vNombOpcMenu = "mnuejecucionpresupuestaria" Then mnuEjecucionPresupuestaria.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'                If vNombOpcMenu = "ejecucionporuni" Then EjecucionPorUni.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnuconvenio" Then mnuConvenio.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnuorganismo" Then mnuOrganismo.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnucategoria" Then mnuCategoria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "mnuproyecto" Then mnuProyecto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
''        If vNombOpcMenu = "ejecucion" Then Ejecucion.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)     'Reportes de Ejecucion
''            If vNombOpcMenu = "ejecucionorganismo" Then EjecucionOrganismo.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecucioncomprobante" Then EjecucionComprobante.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecucioncompromiso" Then EjecucionCompromiso.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecucióndevengado" Then EjecuciónDevengado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
''            If vNombOpcMenu = "ejecuciónpagado" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'        If vNombOpcMenu = "mnumodppto" Then MnuModPpto.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Modificaciones Presupuestarias
'
'    If vNombOpcMenu = "tesoreria" Then Tesoreria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "pp" Then pp.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Pagos pendientes
'        If vNombOpcMenu = "pagosefectuados" Then PagosEfectuados.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Activacion de Cheques
'        If vNombOpcMenu = "cuentasbancarias2" Then CuentasBancarias2.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Desactivacion de cheques
'        If vNombOpcMenu = "manejocheques" Then ManejoCheques.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Trasnferencia
'        If vNombOpcMenu = "ic" Then ic.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
'        If vNombOpcMenu = "ct" Then ct.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
'            If vNombOpcMenu = "mnuimpcheques" Then MnuImpCheques.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Cheques
'            If vNombOpcMenu = "cuentasbancarias" Then CuentasBancarias.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Impresion de Cheques
'
'    If vNombOpcMenu = "ccontabilidad" Then CContabilidad.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
'            If vNombOpcMenu = "comprobantes" Then Comprobantes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'            If vNombOpcMenu = "reportesc" Then ReportesC.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "libromayor" Then LibroMayor.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "libromayorauxiliar" Then LibroMayorAuxiliar.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "balancegeneral" Then BalanceGeneral.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "balancesumassaldos" Then BalanceSumasSaldos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'                If vNombOpcMenu = "estadoresultados" Then EstadoResultados.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "mnuingresos" Then MnuIngresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)          'Egresos
'        If vNombOpcMenu = "ejecucionpresupuestaria" Then EjecucionPresupuestaria.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Ejecucion Presupuestaria
'        If vNombOpcMenu = "reportesingresos" Then ReportesIngresos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Reportes de Ejecucion
'
'''''    If vNombOpcMenu = "administracion" Then Administracion.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''        If vNombOpcMenu = "adquisicionbienes" Then AdquisicionBienes.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "comprasdirectas" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "licitacionesnacionales" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "licitacionesinternacionales" Then LicitacionesInternacionales.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''        If vNombOpcMenu = "contratacionservicios" Then ContratacionServicios.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "consultoresindividuales" Then ConsultoresIndividuales.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "empresasconsultoras" Then EmpresasConsultoras.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''        If vNombOpcMenu = "almacenes" Then Almacenes.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "ingresosa" Then IngresosA.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'''''                If vNombOpcMenu = "salidasa" Then SalidasA.Enabled = IIf(rsNivelAcceso!HABILITADO = "Si", True, False)
'
'    If vNombOpcMenu = "administracion" Then Administracion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "mnuadmicont" Then MnuAdmiCont.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "almacenes" Then Almacenes.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "ingresosa" Then IngresosA.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnuingreso" Then mnuIngreso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnucompras" Then MnuCompras.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnusalidas" Then mnuSalidas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnucontrol" Then MnuControl.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnuinventario" Then mnuInventario.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'              If vNombOpcMenu = "mnuestado" Then mnuestado.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "mnulicitac" Then mnuLicitac.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnusolnoobjec" Then MnuSolNoObjec.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnupublicacion" Then mnuPublicacion.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnucomision" Then mnucomision.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuventapliegos" Then mnuVentaPliegos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnurecepsobres" Then MnuRecepSobres.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuabogado" Then mnuabogado.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuadjudica" Then mnuAdjudica.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuorden" Then mnuorden.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "mnuconsultoria_c" Then mnuConsultoria_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnusolnoobj_c" Then mnuSolNoObj_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnupublicacion_c" Then mnuPublicacion_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuventapliegos_c" Then mnuVentaPliegos_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnurecepcionprop_c" Then mnuRecepcionProp_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuaperturaprop_c" Then mnuAperturaProp_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuadjudicacion_c" Then mnuAdjudicacion_c.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnugespagos" Then mnugesPagos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "mnuoc" Then mnuOC.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        If vNombOpcMenu = "comprasdirectas" Then ComprasDirectas.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "solicitudnoobjecioncd" Then solicitudnoobjecioncd.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "cotizaciones" Then Cotizaciones.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "cuadrocomparativo" Then cuadrocomparativo.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'          If vNombOpcMenu = "ordenpagocd" Then OrdenPagoCD.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'        'If vNombOpcMenu = "progconadq" Then ProgConAdq.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "mnusolicitudesf04" Then mnuSolicitudesF04.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "recursoshumanos" Then RecursosHumanos.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
'        If vNombOpcMenu = "administracionpersonal" Then AdministracionPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "controlpersonal" Then ControlPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "capacitacionpersonal" Then CapacitacionPersonal.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'        If vNombOpcMenu = "evaluaciondesempeño" Then EvaluacionDesempeño.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)
'
'    If vNombOpcMenu = "informaciongerencial" Then InformacionGerencial.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Contabilidad
'
'    If vNombOpcMenu = "mnuadmisistema" Then mnuAdmiSistema.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)     'Administracion del sistema
'        If vNombOpcMenu = "mnucambiarclave" Then mnuCambiarClave.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Cambiar clave
'        If vNombOpcMenu = "mnuusuarios" Then mnuUsuarios.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False)    'Definicion de Usuarios
'        If vNombOpcMenu = "mnunivelacceso" Then mnuNivelAcceso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Cambiar clave
'        If vNombOpcMenu = "mnuprivacceso" Then mnuPrivAcceso.Enabled = IIf(rsNivelAcceso!habilitado = "Si", True, False) 'Privilegios de Operación
'    rsNivelAcceso.MoveNext
'    Wend
'    rsNivelAcceso.MoveFirst
'End If
End Sub

Private Sub MunClasificacionGeneral_Click()
    pw_p_pc_proceso_nivel1.lbl_titulo = MunClasificacionGeneral.Caption
    pw_p_pc_proceso_nivel1.FraNavega = MunClasificacionGeneral.Caption
    pw_p_pc_proceso_nivel1.lbl_titulo2 = MunClasificacionGeneral.Caption
    pw_p_pc_proceso_nivel1.Show
End Sub

'Private Sub Timer1_Timer()
'    If cmd_puestos.Visible = True Then 'checking whether it is visible
'        cmd_puestos.Visible = False 'if visible then, make it invisible
'    ElseIf cmd_puestos.Visible = False Then 'checking whether it is invisible
'        cmd_puestos.Visible = True 'if invisible then, make it visible
'    End If
'End Sub

Private Sub Option1_Click()
        Dim iResult As Integer
        Cry.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_servicio_regional_Alerta.rpt"
        titulo2 = "ALERTAS CONTRATOS DE VENTAS"
        subtitulo2 = "VIGENCIA VENCIDA O A DIAS DE VENCER"
        Cry.Formulas(2) = "Titulo = '" & titulo2 & "'"
        Cry.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        Cry.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = Cry.PrintReport
        If iResult <> 0 Then
            MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    Cry.WindowState = crptMaximized
    frm_alertas.Visible = False
End Sub

Private Sub Option2_Click()
        Dim iResult As Integer
        CR07.ReportFileName = App.Path & "\reportes\comercial\ar_lista_actas_entrega_definitiva_alerta.rpt"
        CR07.WindowShowPrintSetupBtn = True
        CR07.WindowShowRefreshBtn = True
        iResult = CR07.PrintReport
        If iResult <> 0 Then MsgBox CR07.LastErrorNumber & " : " & CR07.LastErrorString, vbCritical, "Error de impresión"
        CR07.WindowState = crptMaximized
    frm_alertas.Visible = False
End Sub

Private Sub Option3_Click()
    frm_alertas.Visible = False
End Sub

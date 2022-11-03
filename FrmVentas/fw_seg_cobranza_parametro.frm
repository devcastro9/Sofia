VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_seg_cobranza_parametro 
   BackColor       =   &H00000000&
   Caption         =   "Reportes Seguimiento Cobranza"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   Icon            =   "fw_seg_cobranza_parametro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox BtnImprimir2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4200
      Picture         =   "fw_seg_cobranza_parametro.frx":0A02
      ScaleHeight     =   615
      ScaleWidth      =   1395
      TabIndex        =   12
      ToolTipText     =   "Por Edificio/Cliente"
      Top             =   3000
      Width           =   1400
   End
   Begin VB.PictureBox BtnGenerar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1320
      Picture         =   "fw_seg_cobranza_parametro.frx":12CF
      ScaleHeight     =   615
      ScaleWidth      =   1395
      TabIndex        =   11
      ToolTipText     =   "Por Gestion-Unidad-Depto y/o Edificio"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Elija los Parametros, luego click en el botón Imprimir ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9060
      Begin VB.ComboBox cmb_gestion 
         Height          =   315
         ItemData        =   "fw_seg_cobranza_parametro.frx":1B9C
         Left            =   2160
         List            =   "fw_seg_cobranza_parametro.frx":1BC1
         TabIndex        =   8
         Top             =   480
         Width           =   1515
      End
      Begin MSDataListLib.DataCombo cmb_unidad 
         Bindings        =   "fw_seg_cobranza_parametro.frx":1C07
         DataField       =   "unidad_codigo"
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmb_departamento 
         Bindings        =   "fw_seg_cobranza_parametro.frx":1C20
         DataField       =   "depto_codigo"
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1440
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "depto_descripcion"
         BoundColumn     =   "depto_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmb_edificio 
         Bindings        =   "fw_seg_cobranza_parametro.frx":1C39
         DataField       =   "edif_codigo"
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   1920
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmb_codigoedificio 
         Bindings        =   "fw_seg_cobranza_parametro.frx":1C52
         DataField       =   "edif_codigo"
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   1920
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   ""
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblFuente 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport CryReporte 
      Left            =   9120
      Top             =   975
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryVsLey 
      Left            =   9120
      Top             =   1830
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryDetalle 
      Left            =   9135
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryUnidad 
      Left            =   9120
      Top             =   1395
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryRep002_financiador 
      Left            =   9135
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   3600
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2400
      Top             =   3720
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4680
      Top             =   3600
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   5880
      Top             =   3720
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblTipoReporte 
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "fw_seg_cobranza_parametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iResult As Integer
Public vProg As String
Public vSubProg As String
Public vProy As String
Public vActi As String
Public glRepPresup As String
Public conDetalle As Boolean
Dim rs_proveedor, rs_cliente, rs_vendedor, rs_cobrador As New ADODB.Recordset
Dim rs_tipo, rs_tipoBenef, rs_ciudad As New ADODB.Recordset
Dim rs_meses, rs_producto As New ADODB.Recordset

Public Sub inicio(Usuario, Proceso As String)
  glRepPresup = Proceso
  Call llena_datos
'  dtpFecha1.Value = Format("01/01/2016", "dd/mm/yyyy")
'  dtpFecha2.Value = Format(Date, "dd/mm/yyyy")
  'dtpFecha2.Value = Date
'  frmRepPresupuesto.Show
End Sub

Private Sub BtnImprimir_Click()
    'LISTADO GENERAL DE VENTAS
'  If optRep001.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep001.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep001.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
'  ElseIf optRep001.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS ACUMULADAS POR MES
'  ElseIf optRep002.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep002.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep002.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep002.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS POR PROVEEDOR Y LINEA
'  ElseIf optRep003.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep003.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep003.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep003.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS Y COBRANZAS POR CLIENTE (Detalle)
'  ElseIf optRep004.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep004.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep004.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep004.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS Y COBRANZAS POR CLIENTE (Totales)
'  ElseIf optrep005.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optrep005.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optrep005.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optrep005.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'COMISIONES POR VENTAS Y COBRANZAS
'  ElseIf optrep006.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optrep006.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optrep006.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optrep006.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'SEGUIMIENTO DE VENTAS POR PRODUCTO
'  ElseIf optRep007.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep007.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep007.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep007.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
''  ElseIf optRep008.Value = True Then
''    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST_cli.rpt ")
''  ElseIf optRep009.Value = True Then
''    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST.rpt ")
''  ElseIf optRep0010.Value = True Then
''    '
'''  ElseIf optRep0011.Value = True Then
''    '
''  'End If
'
'  'LISTADO GENERAL DE VENTAS
'  ElseIf optRep001.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep001.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep001.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
'  ElseIf optRep001.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'  'LISTADO GENERAL DE COBRANZAS
'  ElseIf optRep010.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep010.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep010.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
'  ElseIf optRep010.Value = True And opt_4.Value = True Then
'    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'    CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'
'  'LIBRO DE VENTAS
'  ElseIf optRep011.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep011.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep011.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep011.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_libro_ventas.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'  'End If
'
'  'COBRANZAS POR FACTURA
'  ElseIf optRep012.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep012.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep012.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep012.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'
'  'COBRANZAS POR COBRADOR
'  ElseIf optRep015.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep015.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep015.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep015.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'
'        'COBRANZAS POR RECIBO
'  ElseIf optRep016.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep016.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep016.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep016.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'
'    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'  End If

End Sub

Private Sub BtnImprimir2_Click()
'    If Ado_datos02.Recordset.RecordCount > 0 Then
'        Monto_Bs = Ado_datos02.Recordset!cobranza_total_bs
'        montoLiteral = Literal(CStr(Monto_Bs)) + " Bolivianos"
'            Dim iResult As Variant  ', i%, y%
      CryReporte.ReportFileName = App.Path & "\reportes\ventas\fr_seguimiento_cobranza_kardex.rpt"
      CryReporte.WindowShowRefreshBtn = True
        If cmb_codigoedificio.Text = "" Then
              CryReporte.StoredProcParam(0) = "%"
        Else
              CryReporte.StoredProcParam(0) = cmb_codigoedificio.Text
        End If
'      CryReporte.StoredProcParam(0) = cmb_codigoedificio.Text
'      CryReporte.StoredProcParam(1) = Ado_datos02.Recordset!cobranza_codigo
'      CryReporte.Formulas(1) = "literalcobro = '" & montoLiteral & "' "
'      CryReporte.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_codigo & "' "
'      '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
      iResult = CryReporte.PrintReport
      If iResult <> 0 Then MsgBox CryReporte.LastErrorNumber & " : " & CryReporte.LastErrorString, vbCritical, "Error de impresión"
'    Else
'      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'    End If
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub butEstProg_Click()
'  frmListaEstProg.Show
End Sub

Private Sub cmdAcepta_Click()
    'LISTADO GENERAL DE VENTAS
'  If optRep001.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep001.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep001.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
'  ElseIf optRep001.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS ACUMULADAS POR MES
'  ElseIf optRep002.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep002.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep002.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep002.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS POR PROVEEDOR Y LINEA
'  ElseIf optRep003.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep003.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep003.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep003.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS Y COBRANZAS POR CLIENTE (Detalle)
'  ElseIf optRep004.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep004.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep004.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep004.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS Y COBRANZAS POR CLIENTE (Totales)
'  ElseIf optrep005.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optrep005.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optrep005.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optrep005.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'COMISIONES POR VENTAS Y COBRANZAS
'  ElseIf optrep006.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optrep006.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optrep006.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optrep006.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'SEGUIMIENTO DE VENTAS POR PRODUCTO
'  ElseIf optRep007.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep007.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep007.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep007.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'  ElseIf optRep008.Value = True Then
'    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST_cli.rpt ")
'  ElseIf optRep009.Value = True Then
'    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST.rpt ")
'  ElseIf optRep0010.Value = True Then
'    '
''  ElseIf optRep0011.Value = True Then
'    '
'  End If
End Sub

'Private Sub RepUnidad(tipoRep As String, ArchRep As String)
Private Sub RepUnidad(tipoRep As String, ArchRep As String, titulo1 As String)
'  CryUnidad.ReportFileName = App.Path & ArchRep
'  CryUnidad.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryUnidad.StoredProcParam(0) = tipoRep
''ini reporte
'  If DtcProvCod.Text = "" Then
'        CryUnidad.StoredProcParam(2) = "%"
'  Else
'        CryUnidad.StoredProcParam(2) = DtcProvCod.Text
'  End If
'  If DtcCliCod.Text = "" Then
'        CryUnidad.StoredProcParam(3) = "%"
'  Else
'        CryUnidad.StoredProcParam(3) = DtcCliCod.Text
'  End If
'  If DtcVenCod.Text = "" Then
'        CryUnidad.StoredProcParam(4) = "%"
'  Else
'        CryUnidad.StoredProcParam(4) = DtcVenCod.Text
'  End If
'  If DtcCbrCod.Text = "" Then
'        CryUnidad.StoredProcParam(5) = "%"
'  Else
'        CryUnidad.StoredProcParam(5) = DtcCbrCod.Text
'  End If
''  If DtcTipo.Text = "" Then
''        CryUnidad.StoredProcParam(6) = "%"
''  Else
''        CryUnidad.StoredProcParam(6) = DtcTipo.Text
''  End If
'  CryUnidad.StoredProcParam(6) = tipoRep
'  If optRep007.Value = True Then
'    If DtcProdC.Text = "" Then
'        CryUnidad.StoredProcParam(7) = "%"
'    Else
'        CryUnidad.StoredProcParam(7) = DtcProdC.Text
'    End If
'  End If
''fin reporte
''  Call setParametros(CryUnidad)
'  CryUnidad.Formulas(0) = "FFInicio ='" & dtpFecha1.Value & "'"
'  CryUnidad.Formulas(1) = "FFFinal ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryUnidad.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
''  If ArchRep = "\rep002.rpt" Then
''     CryUnidad.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
''  End If
'  iResult = CryUnidad.PrintReport
'  If iResult <> 0 Then
'    MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'  End If
End Sub

Private Sub Rep001(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte.ReportFileName = App.Path & ArchRep
'  CryReporte.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryReporte.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryReporte.StoredProcParam(2) = tipoRep
'  Call setParametros(CryReporte)
'  CryReporte.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
'  CryReporte.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryReporte.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryReporte.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If
  iResult = CryReporte.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte.LastErrorNumber & " : " & CryReporte.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub


Private Sub RepDetalle(tipoRep As String, ArchRep As String, titulo1 As String)
'  CryDetalle.ReportFileName = App.Path & ArchRep
'  CryDetalle.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryDetalle.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryDetalle.StoredProcParam(2) = tipoRep
'  Call setParametros(CryDetalle)
'  CryDetalle.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
'  CryDetalle.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryDetalle.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryDetalle.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If
'  iResult = CryDetalle.PrintReport
'  If iResult <> 0 Then
'    MsgBox CryDetalle.LastErrorNumber & " : " & CryDetalle.LastErrorString, vbCritical + vbOKOnly, "Error..."
'  End If
End Sub

Private Sub setParametros(objCryRep As Object)
'  If dcmFte_codigo.Text = "" Then
'    objCryRep.StoredProcParam(3) = "%"
'  Else
'    objCryRep.StoredProcParam(3) = dcmFte_codigo.BoundText
'  End If
'  If cdmOrganismo.Text = "" Then
'    objCryRep.StoredProcParam(4) = "%"
'  Else
'    objCryRep.StoredProcParam(4) = cdmOrganismo.BoundText
'  End If
'  If dtcboconvenio.Text = "" Then
'    objCryRep.StoredProcParam(5) = "%"
'  Else
'    objCryRep.StoredProcParam(5) = dtcboconvenio.BoundText
'  End If
'  If txtProg.Text = "" Then
'    objCryRep.StoredProcParam(6) = "%"
'  Else
'    objCryRep.StoredProcParam(6) = txtProg.Text
'  End If
'  If txtSubProg.Text = "" Then
'    objCryRep.StoredProcParam(7) = "%"
'  Else
'    objCryRep.StoredProcParam(7) = txtSubProg.Text
'  End If
'  If TxtProy.Text = "" Then
'    objCryRep.StoredProcParam(8) = "%"
'  Else
'    objCryRep.StoredProcParam(8) = TxtProy.Text
'  End If
'  If txtAct.Text = "" Then
'    objCryRep.StoredProcParam(9) = "%"
'  Else
'    objCryRep.StoredProcParam(9) = txtAct.Text
'  End If
'  If txtpartida.Text = "" Then
'    objCryRep.StoredProcParam(10) = "%"
'  Else
'    objCryRep.StoredProcParam(10) = txtpartida.Text
'  End If
End Sub

Private Sub Command1_Click()
'ok = frmListaEstProg.getcodigo(valor, valor)
'frmListaEstProg.Show
End Sub

Private Sub llena_datos()
'  Set tFc_fuente_financiamiento = New ADODB.Recordset
'  If tFc_fuente_financiamiento.State = 1 Then tFc_fuente_financiamiento.Close
'    tFc_fuente_financiamiento.Open "SELECT fte_codigo, fte_codigo + '  ' + fte_descripcion_larga as fte_descripcion_larga FROM fc_fuente_financiamiento order by fte_codigo ", db, adOpenDynamic, adLockOptimistic
'  Set frmRepPresupuesto.Adodc_p.Recordset = tFc_fuente_financiamiento

    
'    Set rs_proveedor = New ADODB.Recordset
'    If rs_proveedor.State = 1 Then rs_proveedor.Close
'    rs_proveedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=3 OR tipoben_codigo=22) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    'rs_proveedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=2 OR tipoben_codigo=22) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_proveedor.Recordset = rs_proveedor
'    Ado_proveedor.Refresh
'
'    Set rs_cliente = New ADODB.Recordset
'    If rs_cliente.State = 1 Then rs_cliente.Close
'    rs_cliente.Open "select * from gc_beneficiario WHERE (tipoben_codigo <> 1 AND tipoben_codigo <> 23) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    'rs_cliente.Open "select * from gc_beneficiario WHERE (tipoben_codigo <> 2 AND tipoben_codigo <> 22)  ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    Set ado_Cliente.Recordset = rs_cliente
'    ado_Cliente.Refresh
'
'    Set rs_vendedor = New ADODB.Recordset
'    If rs_vendedor.State = 1 Then rs_vendedor.Close
'    'rs_vendedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=6 OR tipoben_codigo=10) and (beneficiario_deudor = 'SI') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    rs_vendedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=1 OR tipoben_codigo=0) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_vendedor.Recordset = rs_vendedor
'    Ado_vendedor.Refresh
'
'    Set rs_cobrador = New ADODB.Recordset
'    If rs_cobrador.State = 1 Then rs_cobrador.Close
'    'rs_cobrador.Open "select * from gc_beneficiario WHERE (tipoben_codigo=7 OR tipoben_codigo=10) and (beneficiario_deudor = 'SI') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    rs_cobrador.Open "select * from gc_beneficiario WHERE (tipoben_codigo=1 OR tipoben_codigo=0) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_Cobrador.Recordset = rs_cobrador
'    Ado_Cobrador.Refresh
'
'    Set rs_tipo = New ADODB.Recordset
'    If rs_tipo.State = 1 Then rs_tipo.Close
'    rs_tipo.Open "select venta_tipo, venta_tipo_descripcion from ac_tipo_compra_venta WHERE estado_codigo='APR' ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_Tipo.Recordset = rs_tipo
'    Ado_Tipo.Refresh
'
'    Set rs_tipoBenef = New ADODB.Recordset
'    If rs_tipoBenef.State = 1 Then rs_tipoBenef.Close
'    rs_tipoBenef.Open "select tipoben_codigo, tipoben_Descripcion from gc_tipo_beneficiario WHERE (ESTADO_codigo='APR') ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_TipoBenef.Recordset = rs_tipoBenef
'    Ado_TipoBenef.Refresh
'
'    Set rs_ciudad = New ADODB.Recordset
'    If rs_ciudad.State = 1 Then rs_ciudad.Close
'    'rs_ciudad.Open "select Depto AS procedencia, municipio AS lugar_procedencia from gc_beneficiario WHERE (tipoben_codigo<>'B' and tipoben_codigo<>'O' and tipoben_codigo<>'P') and (activo = 'S') group BY Depto, municipio ", DB, adOpenKeyset, adLockReadOnly
'    rs_ciudad.Open "select Depto_codigo , munic_codigo from gc_beneficiario WHERE (tipoben_codigo <>0 ) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') group BY Depto_codigo, munic_codigo ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_Ciudad.Recordset = rs_ciudad
'    Ado_Ciudad.Refresh
'
''    Set rs_meses = New ADODB.Recordset
''    If rs_meses.State = 1 Then rs_meses.Close
''    rs_meses.Open "select * from gc_periodos WHERE (estado_registro = 'S') ", db, adOpenKeyset, adLockReadOnly
''    Set Ado_Meses.Recordset = rs_meses
''    Ado_Meses.Refresh
'
'    Set rs_producto = New ADODB.Recordset
'    If rs_producto.State = 1 Then rs_producto.Close
'    rs_producto.Open "select bien_codigo, concepto_venta from ao_ventas_detalle group BY bien_codigo, concepto_venta ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_Producto.Recordset = rs_producto
'    Ado_Producto.Refresh
'
'    DtcProvCod.Enabled = False
'    DtcProvDes.Enabled = False
'    DtcCliCod.Enabled = True
'    DtcCliDes.Enabled = True
'    DtcVenCod.Enabled = True
'    DtcVenDes.Enabled = True
'    DtcCbrCod.Enabled = False
'    DtcCbrDes.Enabled = False
'    DtcMes.Enabled = False
'    DtcMesC.Enabled = False
'    DtcProd.Enabled = False
'    DtcProdC.Enabled = False
End Sub
Private Sub showEtiquetas(mostrar As Boolean)
'  If mostrar Then
'    lblFuente.Visible = True
'    lblOrg.Visible = True
'    lblConv.Visible = True
'    lblEstr.Visible = True
'    lblPartida.Visible = True
''    dcmFte_codigo.Visible = True
''    cdmOrganismo.Visible = True
''    dtcboconvenio.Visible = True
'    txtProg.Visible = True
'    txtSubProg.Visible = True
'    txtProy.Visible = True
'    txtAct.Visible = True
'    butEstProg.Visible = True
'    txtPartida.Visible = True
'  Else
'    lblFuente.Visible = False
'    lblOrg.Visible = False
'    lblConv.Visible = False
'    lblEstr.Visible = False
'    lblPartida.Visible = False
''    dcmFte_codigo.Visible = False
''    cdmOrganismo.Visible = False
''    dtcboconvenio.Visible = False
'    txtProg.Visible = False
'    txtSubProg.Visible = False
'    txtProy.Visible = False
'    txtAct.Visible = False
'    butEstProg.Visible = False
'    txtPartida.Visible = False
'  End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub DtcCbrCod_Click(Area As Integer)
'    DtcCbrDes.BoundText = DtcCbrCod.BoundText
End Sub

Private Sub BtnGenerar_Click()
   Dim vgestion, vunidad, vdepartamento, vedificio As String
   vgestion = Cmb_gestion.Text
   vunidad = cmb_unidad.BoundText
   vdepartamento = cmb_departamento.BoundText
   vedificio = cmb_edificio.BoundText
   If Trim(Cmb_gestion.Text) = "" Then vgestion = " "
   If Trim(cmb_unidad.BoundText) = "" Then vunidad = " "
   If Trim(cmb_departamento.BoundText) = "" Then vdepartamento = " "
   If Trim(cmb_edificio.BoundText) = "" Then vedificio = " "
   Dim iResult As Integer
  CryDetalle.WindowShowPrintSetupBtn = True
  CryDetalle.WindowShowRefreshBtn = True
  CryDetalle.StoredProcParam(0) = vgestion
  CryDetalle.StoredProcParam(1) = vunidad
  CryDetalle.StoredProcParam(2) = vdepartamento
  CryDetalle.StoredProcParam(3) = vedificio
  
  ' Verifica tipo de reporte.
  If lblTipoReporte.Caption <> "" Then
        CryDetalle.ReportFileName = App.Path & "\REPORTES\Comex\fr_seguimiento_pago.rpt"
  Else
        CryDetalle.ReportFileName = App.Path & "\REPORTES\Comex\fr_seguimiento_cobranza.rpt"
  End If
  
  iResult = CryDetalle.PrintReport
  If iResult <> 0 Then
      MsgBox CryDetalle.LastErrorNumber & " : " & CryDetalle.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  CryDetalle.WindowState = crptMaximized
   
   
End Sub

Private Sub DtcCbrDes_Click(Area As Integer)
    
End Sub

Private Sub DtcCliCod_Click(Area As Integer)
    
End Sub

Private Sub DtcCliDes_Click(Area As Integer)
    
End Sub

Private Sub DtcCiu_Click(Area As Integer)
    
End Sub

Private Sub DtcDepto_Click(Area As Integer)
   
End Sub

Private Sub DtcProvCod_Click(Area As Integer)
   
End Sub

Private Sub DtcProvDes_Click(Area As Integer)
   
End Sub

Private Sub DtcTipo_Click(Area As Integer)
   
End Sub

Private Sub dtctipoDes_Click(Area As Integer)
   
End Sub

Private Sub DtcTipoCli_Click(Area As Integer)
   
End Sub

Private Sub DtcTipoCliDes_Click(Area As Integer)
   
End Sub

Private Sub DtcVenCod_Click(Area As Integer)
   
End Sub

Private Sub DtcVenDes_Click(Area As Integer)
   
End Sub


Private Sub cmb_codigoedificio_Click(Area As Integer)
   cmb_edificio.BoundText = cmb_codigoedificio.BoundText
End Sub

Private Sub cmb_edificio_Click(Area As Integer)
  cmb_codigoedificio.BoundText = cmb_edificio.BoundText
End Sub

Private Sub Form_Load()
    Call CargarControles
    
	Call SeguridadSet(Me)
End Sub

Private Sub CargarControles()
   Dim rs_UNIDAD As ADODB.Recordset
   Dim rs_departamento As ADODB.Recordset
   Dim rs_edificio As ADODB.Recordset
   
   Set rs_UNIDAD = New ADODB.Recordset
    If rs_UNIDAD.State = 1 Then rs_proveedor.Close
    rs_UNIDAD.Open " select * from gc_unidad_ejecutora ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos1.Recordset = rs_UNIDAD
    Ado_datos1.Refresh
    
     Set rs_departamento = New ADODB.Recordset
    If rs_departamento.State = 1 Then rs_departamento.Close
    rs_departamento.Open " select * from gc_departamento ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos2.Recordset = rs_departamento
    Ado_datos2.Refresh
    
     Set rs_edificio = New ADODB.Recordset
    If rs_edificio.State = 1 Then rs_edificio.Close
    rs_edificio.Open " select * from gc_edificaciones ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos3.Recordset = rs_edificio
    Ado_datos3.Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Private Sub opt_rep009_Click()
  Call SetControles(False, True)
End Sub

Private Sub optRep001_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = False
'  DtcProvDes.Enabled = False
'  DtcCliCod.Enabled = True
'  DtcCliDes.Enabled = True
'  DtcVenCod.Enabled = True
'  DtcVenDes.Enabled = True
'  DtcCbrCod.Enabled = False
'  DtcCbrDes.Enabled = False
'  DtcTipo.Enabled = True
'  DtcTipoDes.Enabled = True
'  DtcTipoCliDes.Enabled = True
'  DtcCiu.Enabled = True
'  DtcMes.Enabled = False
'  DtcMesC.Enabled = False
'  DtcProd.Enabled = False
'  DtcProdC.Enabled = False
End Sub

Private Sub optRep002_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = False
'  DtcProvDes.Enabled = False
'  DtcCliCod.Enabled = False
'  DtcCliDes.Enabled = False
'  DtcVenCod.Enabled = False
'  DtcVenDes.Enabled = False
'  DtcCbrCod.Enabled = False
'  DtcCbrDes.Enabled = False
'  DtcTipo.Enabled = False
'  DtcTipoDes.Enabled = False
'  DtcTipoCliDes.Enabled = True
'  DtcCiu.Enabled = True
'  DtcMes.Enabled = True
'  DtcMesC.Enabled = True
'  DtcProd.Enabled = False
'  DtcProdC.Enabled = False
End Sub

Private Sub optRep003_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = True
'  DtcProvDes.Enabled = True
'  DtcCliCod.Enabled = True
'  DtcCliDes.Enabled = True
'  DtcVenCod.Enabled = False
'  DtcVenDes.Enabled = False
'  DtcCbrCod.Enabled = False
'  DtcCbrDes.Enabled = False
'  DtcTipo.Enabled = True
'  DtcTipoDes.Enabled = True
'  DtcTipoCliDes.Enabled = True
'  DtcCiu.Enabled = True
'  DtcMes.Enabled = False
'  DtcMesC.Enabled = False
'  DtcProd.Enabled = False
'  DtcProdC.Enabled = False
End Sub

Private Sub optRep004_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = False
'  DtcProvDes.Enabled = False
'  DtcCliCod.Enabled = True
'  DtcCliDes.Enabled = True
'  DtcVenCod.Enabled = True
'  DtcVenDes.Enabled = True
'  DtcCbrCod.Enabled = False
'  DtcCbrDes.Enabled = False
'  DtcTipo.Enabled = True
'  DtcTipoDes.Enabled = True
'  DtcTipoCliDes.Enabled = True
'  DtcCiu.Enabled = True
'  DtcMes.Enabled = False
'  DtcMesC.Enabled = False
'  DtcProd.Enabled = False
'  DtcProdC.Enabled = False
End Sub

Private Sub optRep005_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = False
'  DtcProvDes.Enabled = False
'  DtcCliCod.Enabled = True
'  DtcCliDes.Enabled = True
'  DtcVenCod.Enabled = True
'  DtcVenDes.Enabled = True
'  DtcCbrCod.Enabled = False
'  DtcCbrDes.Enabled = False
'  DtcTipoCliDes.Enabled = True
'  DtcCiu.Enabled = True
'  DtcMes.Enabled = False
'  DtcMesC.Enabled = False
'  DtcProd.Enabled = False
'  DtcProdC.Enabled = False
End Sub

Private Sub optRep006_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = False
'  DtcProvDes.Enabled = False
'  DtcCliCod.Enabled = True
'  DtcCliDes.Enabled = True
'  DtcVenCod.Enabled = True
'  DtcVenDes.Enabled = True
'  DtcCbrCod.Enabled = True
'  DtcCbrDes.Enabled = True
'  DtcTipo.Enabled = True
'  DtcTipoDes.Enabled = True
'  DtcTipoCliDes.Enabled = True
'  DtcCiu.Enabled = True
'  DtcMes.Enabled = False
'  DtcMesC.Enabled = False
'  DtcProd.Enabled = False
'  DtcProdC.Enabled = False
End Sub

Private Sub optRep007_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = False
'  DtcProvDes.Enabled = False
'  DtcCliCod.Enabled = False
'  DtcCliDes.Enabled = False
'  DtcVenCod.Enabled = False
'  DtcVenDes.Enabled = False
'  DtcCbrCod.Enabled = False
'  DtcCbrDes.Enabled = False
'  DtcTipo.Enabled = False
'  DtcTipoDes.Enabled = False
'  DtcTipoCliDes.Enabled = False
'  DtcCiu.Enabled = False
'  DtcMes.Enabled = False
'  DtcMesC.Enabled = False
'  DtcProd.Enabled = True
'  DtcProdC.Enabled = True
End Sub

Private Sub optRep008_Click()
  Call SetControles(False, False)
End Sub
Private Sub optRep0010_Click()
'  FrameTipo.Visible = False
End Sub
Private Sub optRep0011_Click()
  Call SetControles(False, True)
End Sub
Private Sub optRep002Finanzas_Click()
  Call SetControles(False, True)
End Sub
Private Sub SetControles(tipo, conDet As Boolean)
'  FrameTipo.Visible = tipo
'  FrameConDet.Visible = conDet
End Sub

Private Sub optRep011_Click()
'  Call SetControles(True, False)
'  DtcProvCod.Enabled = False
'  DtcProvDes.Enabled = False
'  DtcCliCod.Enabled = True
'  DtcCliDes.Enabled = True
'  DtcVenCod.Enabled = True
'  DtcVenDes.Enabled = True
'  DtcCbrCod.Enabled = False
'  DtcCbrDes.Enabled = False
'  DtcTipo.Enabled = True
'  DtcTipoDes.Enabled = True
'  DtcTipoCliDes.Enabled = True
'  DtcCiu.Enabled = True
'  DtcMes.Enabled = False
'  DtcMesC.Enabled = False
'  DtcProd.Enabled = False
'  DtcProdC.Enabled = False
'  BtnImprimir2.Visible = True
End Sub


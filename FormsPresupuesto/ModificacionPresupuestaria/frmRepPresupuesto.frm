VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepPresupuesto 
   BackColor       =   &H00000000&
   Caption         =   "Reporte de Ejecucion Presupuestaria"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   Icon            =   "frmRepPresupuesto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmRepPresupuesto.frx":058A
      Height          =   315
      Left            =   2715
      TabIndex        =   9
      Top             =   2985
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "fte_codigo"
      BoundColumn     =   "fte_descripcion_larga"
      Text            =   "DataCombo1"
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frmRepPresupuesto.frx":05A0
      ScaleHeight     =   915
      ScaleWidth      =   7680
      TabIndex        =   43
      Top             =   120
      Width           =   7740
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   1560
         Picture         =   "frmRepPresupuesto.frx":6C5D2
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   360
         Picture         =   "frmRepPresupuesto.frx":6C7DC
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Imprime Lista de Personas"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte Ejecución Presupuestaria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   2445
         TabIndex        =   44
         Top             =   300
         Width           =   5175
      End
   End
   Begin VB.Frame FrameConDet 
      Caption         =   "Con Detalle"
      Height          =   600
      Left            =   6600
      TabIndex        =   39
      Top             =   6360
      Visible         =   0   'False
      Width           =   1800
      Begin VB.OptionButton optSi 
         Caption         =   "Si"
         Height          =   225
         Left            =   105
         TabIndex        =   41
         Top             =   250
         Width           =   705
      End
      Begin VB.OptionButton optNo 
         Caption         =   "No"
         Height          =   195
         Left            =   945
         TabIndex        =   40
         Top             =   250
         Value           =   -1  'True
         Width           =   600
      End
   End
   Begin VB.Frame FrameTipo 
      Caption         =   "Comprobantes"
      Height          =   1455
      Left            =   4065
      TabIndex        =   27
      Top             =   5310
      Width           =   2895
      Begin VB.OptionButton Opt_comp_poa 
         Caption         =   "Solo Comprometidos + POA"
         Height          =   255
         Left            =   180
         TabIndex        =   42
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton opt_pag 
         Caption         =   "Solo Pagados"
         Height          =   285
         Left            =   180
         TabIndex        =   30
         Top             =   780
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton Opt_comp 
         Caption         =   "Comprometidos no Dev. ni Pag."
         Height          =   240
         Left            =   180
         TabIndex        =   29
         Top             =   255
         Width           =   2670
      End
      Begin VB.OptionButton opt_dev 
         Caption         =   "Devengados no Pagados"
         Height          =   285
         Left            =   180
         TabIndex        =   28
         Top             =   510
         Width           =   2430
      End
   End
   Begin VB.Frame ConProy00 
      Caption         =   "Con Proyecto 00 "
      Height          =   960
      Left            =   6600
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   1725
      Begin VB.OptionButton OptProyXX 
         Caption         =   "No"
         Height          =   285
         Left            =   150
         TabIndex        =   26
         Top             =   540
         Width           =   945
      End
      Begin VB.OptionButton OptProy00 
         Caption         =   "Si"
         Height          =   240
         Left            =   180
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   930
      End
   End
   Begin VB.TextBox txtPartida 
      Height          =   315
      Left            =   2730
      TabIndex        =   18
      Top             =   3345
      Width           =   1095
   End
   Begin VB.TextBox txtAct 
      Height          =   315
      Left            =   3915
      TabIndex        =   17
      Top             =   3015
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtProy 
      Height          =   315
      Left            =   3405
      TabIndex        =   16
      Top             =   3015
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtSubProg 
      Height          =   315
      Left            =   2910
      TabIndex        =   15
      Top             =   3015
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtProg 
      Height          =   315
      Left            =   2370
      TabIndex        =   14
      Top             =   3015
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton butEstProg 
      Caption         =   "..."
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      Top             =   3015
      Visible         =   0   'False
      Width           =   450
   End
   Begin MSAdodcLib.Adodc AdoFcConvenios 
      Height          =   330
      Left            =   4590
      Top             =   2625
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoFc_organismo_financiamiento 
      Height          =   330
      Left            =   4545
      Top             =   2325
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc1"
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
   Begin MSDataListLib.DataCombo cdmOrganismo 
      Bindings        =   "frmRepPresupuesto.frx":6CD99
      Height          =   315
      Left            =   2730
      TabIndex        =   11
      Top             =   2310
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "org_descripcion"
      BoundColumn     =   "org_codigo"
      Text            =   "Todos"
   End
   Begin MSAdodcLib.Adodc Adodc_p 
      Height          =   330
      Left            =   4575
      Top             =   1980
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Caption         =   "Adodc1"
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
   Begin VB.Frame fmrTipoReporte 
      BackColor       =   &H00000000&
      Caption         =   "Tipo de Reporte"
      ForeColor       =   &H00FFFF80&
      Height          =   4110
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   6495
      Begin VB.OptionButton optRep002_financiero 
         BackColor       =   &H00000000&
         Caption         =   "Por Organismo Financiador Vs. Presupuesto por Proyecto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   150
         TabIndex        =   38
         Top             =   915
         Width           =   6045
      End
      Begin VB.OptionButton optRep002Finanzas 
         BackColor       =   &H00000000&
         Caption         =   "Por Organismo Financiador, Proyecto Vs. Presupuesto Institucional"
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   150
         TabIndex        =   37
         Top             =   3600
         Width           =   5610
      End
      Begin VB.OptionButton opt_rep001_comp_dev 
         BackColor       =   &H00000000&
         Caption         =   $"frmRepPresupuesto.frx":6CDC6
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         TabIndex        =   31
         Top             =   3210
         Width           =   6180
      End
      Begin VB.OptionButton optRep008 
         BackColor       =   &H00000000&
         Caption         =   "Ejecucion Acumulada por Proyecto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   150
         TabIndex        =   23
         Top             =   2640
         Width           =   3195
      End
      Begin VB.OptionButton optRep007 
         BackColor       =   &H00000000&
         Caption         =   "Ejecucion Acumulada por Categoria"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   150
         TabIndex        =   22
         Top             =   2910
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.OptionButton optRep006 
         BackColor       =   &H00000000&
         Caption         =   "Ejecucion Acumulada por Organismo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   150
         TabIndex        =   21
         Top             =   2235
         Width           =   3660
      End
      Begin VB.OptionButton optRep005 
         BackColor       =   &H00000000&
         Caption         =   "Ejecucion Acumulada por Convenio"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   150
         TabIndex        =   20
         Top             =   1905
         Width           =   3660
      End
      Begin VB.OptionButton optRep003 
         BackColor       =   &H00000000&
         Caption         =   "Por Unidad Productiva, Organismo Financiador, Proyecto y Partida"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         TabIndex        =   19
         Top             =   1215
         Width           =   6180
      End
      Begin VB.OptionButton optRep001 
         BackColor       =   &H00000000&
         Caption         =   "Detalle de Ejecucion Presupuestaria por Financiador (Comprobantes pagados)"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   6180
      End
      Begin VB.OptionButton optRep004 
         BackColor       =   &H00000000&
         Caption         =   "Ejecucion Acumulada por Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   150
         TabIndex        =   7
         Top             =   1545
         Width           =   3660
      End
      Begin VB.OptionButton optRep002 
         BackColor       =   &H00000000&
         Caption         =   "Por Organismo Financiador, Proyecto Vs. Presupuesto Institucional"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   150
         TabIndex        =   6
         Top             =   585
         Width           =   6045
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Fecha de Registro de Comprobante "
      ForeColor       =   &H00FFFF80&
      Height          =   675
      Left            =   645
      TabIndex        =   0
      Top             =   1200
      Width           =   6210
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83951617
         CurrentDate     =   40909
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83951617
         CurrentDate     =   41274
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Fecha Final :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Fecha Inicial :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   285
         Width           =   1110
      End
   End
   Begin Crystal.CrystalReport CryReporte 
      Left            =   6645
      Top             =   5295
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSDataListLib.DataCombo dcmFte_codigo 
      Bindings        =   "frmRepPresupuesto.frx":6CE51
      Height          =   315
      Left            =   2730
      TabIndex        =   10
      Top             =   1980
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "fte_descripcion_larga"
      BoundColumn     =   "fte_codigo"
      Text            =   "Todos"
   End
   Begin MSDataListLib.DataCombo dtcboconvenio 
      Bindings        =   "frmRepPresupuesto.frx":6CE67
      Height          =   315
      Left            =   2730
      TabIndex        =   12
      Top             =   2640
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Denominacion_convenio"
      BoundColumn     =   "codigo_convenio"
      Text            =   "Todos"
   End
   Begin Crystal.CrystalReport CryVsLey 
      Left            =   6600
      Top             =   4470
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
      Left            =   6645
      Top             =   5730
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
      Left            =   6600
      Top             =   4875
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryRep002_financiador 
      Left            =   6600
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lblFuente 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuente de Financiamiento :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   645
      TabIndex        =   36
      Top             =   2025
      Width           =   1935
   End
   Begin VB.Label lblOrg 
      BackStyle       =   0  'Transparent
      Caption         =   "Organismo Financiamiento :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   645
      TabIndex        =   35
      Top             =   2385
      Width           =   1980
   End
   Begin VB.Label lblConv 
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   645
      TabIndex        =   34
      Top             =   2745
      Width           =   855
   End
   Begin VB.Label lblEstr 
      BackStyle       =   0  'Transparent
      Caption         =   "Estructura Programatica :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   645
      TabIndex        =   33
      Top             =   3105
      Width           =   1935
   End
   Begin VB.Label lblPartida 
      BackStyle       =   0  'Transparent
      Caption         =   "Partida :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   645
      TabIndex        =   32
      Top             =   3435
      Width           =   855
   End
End
Attribute VB_Name = "frmRepPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IResult As Integer
Public vProg As String
Public vSubProg As String
Public vProy As String
Public vActi As String
Public glRepPresup As String
Public conDetalle As Boolean

Public Sub inicio(Usuario, Proceso As String)
  glRepPresup = Proceso
  Call llena_datos
  dtpFecha2.Value = Date
  frmRepPresupuesto.Show
End Sub

Private Sub butEstProg_Click()
  frmListaEstProg.Show
End Sub

Private Sub BtnImprimir_Click()
  If optRep001.Value = True And opt_pag.Value = True Then
    Call Rep001("REP001", "\rep001.rpt", "")
  ElseIf optRep001.Value = True And opt_dev.Value = True Then
    Call Rep001("REP001_DEV", "\rep001_dev.rpt", "")
  ElseIf optRep001.Value = True And Opt_comp.Value = True Then
    Call Rep001("REP001_PAG", "\rep001_comp.rpt", "")
  ElseIf optRep001.Value = True And Opt_comp_poa.Value = True Then
    Call Rep001("REP001_PAG", "\rep001_comp_poa.rpt", "")
  ElseIf optRep002.Value = True Then  'vs. presupueto de ley
    Call RepVsLey("REP002", "\rep002.rpt", "")
  ElseIf optRep002_financiero.Value = True Then   'vs. presupueto de ley del Financiador
    Call RepVsLeyFinanciador("REP002_FINANCIERO", "\rep002_financiero.rpt", "")
  ElseIf optRep003.Value = True Then
    Call RepUnidad("REP003", "\rep003.rpt", "")
  ElseIf optRep004.Value = True Then
    Call RepVsLey("REP004", "\rep004.rpt", "Ejecución Presupuestaria Acumulada por Unidad")
  ElseIf optRep005.Value = True Then
    Call RepVsLey("REP004", "\rep005.rpt", "Ejecución Presupuestaria Acumulada por Convenio")
  ElseIf optRep006.Value = True Then
    Call RepVsLey("REP004", "\rep006.rpt", "Ejecución Presupuestaria Acumulada por Organismo")
  ElseIf optRep007.Value = True Then
    Call RepVsLey("REP004_CAT", "\rep007.rpt", "Ejecución Presupuestaria Acumulada por Categoria")
  ElseIf optRep008.Value = True Then
    Call RepVsLey("REP004", "\rep008.rpt", "Ejecución Presupuestaria Acumulada por Proyecto")
  ElseIf opt_rep001_comp_dev.Value = True Then
    Call RepDetalle("REP001_COMP_DEV", "\rep001_comp_dev.rpt", "")
  ElseIf optRep002Finanzas.Value = True Then
    Call RepVsLey("REP002", "\Rep002Finanzas.rpt", "")
  End If
End Sub

Private Sub Rep001(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte.ReportFileName = App.Path & ArchRep
  CryReporte.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryReporte.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryReporte.StoredProcParam(2) = tipoRep
  Call setParametros(CryReporte)
  CryReporte.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
  CryReporte.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  If titulo1 <> "" Then
    CryReporte.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
  End If
  
  If ArchRep = "\rep002.rpt" Then
     CryReporte.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
  End If
  
  IResult = CryReporte.PrintReport
  If IResult <> 0 Then
    MsgBox CryReporte.LastErrorNumber & " : " & CryReporte.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub RepVsLey(tipoRep As String, ArchRep As String, titulo1 As String)
  CryVsLey.ReportFileName = App.Path & ArchRep
  CryVsLey.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryVsLey.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryVsLey.StoredProcParam(2) = tipoRep
  Call setParametros(CryVsLey)
  CryVsLey.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
  CryVsLey.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  If titulo1 <> "" Then
    CryVsLey.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
  End If
  
  If ArchRep = "\rep002.rpt" Or ArchRep = "\Rep002Finanzas.rpt" Then
     CryVsLey.Formulas(2) = "conDetalle = " & IIf(optSi, "true", "false")
  End If
  
  IResult = CryVsLey.PrintReport
  If IResult <> 0 Then
    MsgBox CryVsLey.LastErrorNumber & " : " & CryVsLey.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub RepUnidad(tipoRep As String, ArchRep As String, titulo1 As String)
  CryUnidad.ReportFileName = App.Path & ArchRep
  CryUnidad.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryUnidad.StoredProcParam(2) = tipoRep
  Call setParametros(CryUnidad)
  CryUnidad.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
  CryUnidad.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  If titulo1 <> "" Then
    CryUnidad.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
  End If
  
  If ArchRep = "\rep002.rpt" Then
     CryUnidad.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
  End If
  
  IResult = CryUnidad.PrintReport
  If IResult <> 0 Then
    MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub RepDetalle(tipoRep As String, ArchRep As String, titulo1 As String)
  CryDetalle.ReportFileName = App.Path & ArchRep
  CryDetalle.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryDetalle.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryDetalle.StoredProcParam(2) = tipoRep
  Call setParametros(CryDetalle)
  CryDetalle.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
  CryDetalle.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  If titulo1 <> "" Then
    CryDetalle.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
  End If
  
  If ArchRep = "\rep002.rpt" Then
     CryDetalle.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
  End If
  
  IResult = CryDetalle.PrintReport
  If IResult <> 0 Then
    MsgBox CryDetalle.LastErrorNumber & " : " & CryDetalle.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub setParametros(objCryRep As Object)
  If dcmFte_codigo.Text = "" Then
    objCryRep.StoredProcParam(3) = "%"
  Else
    objCryRep.StoredProcParam(3) = dcmFte_codigo.BoundText
  End If
  If cdmOrganismo.Text = "" Then
    objCryRep.StoredProcParam(4) = "%"
  Else
    objCryRep.StoredProcParam(4) = cdmOrganismo.BoundText
  End If
  If dtcboconvenio.Text = "" Then
    objCryRep.StoredProcParam(5) = "%"
  Else
    objCryRep.StoredProcParam(5) = dtcboconvenio.BoundText
  End If
  If txtProg.Text = "" Then
    objCryRep.StoredProcParam(6) = "%"
  Else
    objCryRep.StoredProcParam(6) = txtProg.Text
  End If
  
  If txtSubProg.Text = "" Then
    objCryRep.StoredProcParam(7) = "%"
  Else
    objCryRep.StoredProcParam(7) = txtSubProg.Text
  End If
  
  If txtProy.Text = "" Then
    objCryRep.StoredProcParam(8) = "%"
  Else
    objCryRep.StoredProcParam(8) = txtProy.Text
  End If
  
  If txtAct.Text = "" Then
    objCryRep.StoredProcParam(9) = "%"
  Else
    objCryRep.StoredProcParam(9) = txtAct.Text
  End If
  
  If txtPartida.Text = "" Then
    objCryRep.StoredProcParam(10) = "%"
  Else
    objCryRep.StoredProcParam(10) = txtPartida.Text
  End If
End Sub

Private Sub Command1_Click()
'ok = frmListaEstProg.getcodigo(valor, valor)
frmListaEstProg.Show
End Sub


Private Sub BtnSalir_Click()
  Unload Me
End Sub

'Private Sub DataCombo1_Click(Area As Integer)
'  DataCombo2.Text = DataCombo1.BoundText
'End Sub

'Private Sub DataCombo2_Click(Area As Integer)
'Print DataCombo2.BoundText
'End Sub

Private Sub llena_datos()
  Set tFc_fuente_financiamiento = New ADODB.Recordset
  If tFc_fuente_financiamiento.State = 1 Then tFc_fuente_financiamiento.Close
    tFc_fuente_financiamiento.Open "SELECT fte_codigo, fte_codigo + '  ' + fte_descripcion_larga as fte_descripcion_larga FROM fc_fuente_financiamiento ", db, adOpenDynamic, adLockOptimistic
  Set frmRepPresupuesto.Adodc_p.Recordset = tFc_fuente_financiamiento
  
  Set tFc_organismo_financiamiento = New ADODB.Recordset
  If tFc_organismo_financiamiento.State = 1 Then tFc_organismo_financiamiento.Close
    tFc_organismo_financiamiento.Open "SELECT org_codigo, org_codigo + '  ' + org_descripcion as org_descripcion FROM Fc_organismo_financiamiento ", db, adOpenDynamic, adLockOptimistic
  Set frmRepPresupuesto.adoFc_organismo_financiamiento.Recordset = tFc_organismo_financiamiento
  
  Set tFc_convenios = New ADODB.Recordset
  If tFc_convenios.State = 1 Then tFc_convenios.Close
    tFc_convenios.Open "SELECT codigo_convenio, codigo_convenio + ' => ' + Denominacion_convenio as Denominacion_convenio FROM Fc_convenios ", db, adOpenDynamic, adLockReadOnly
  Set frmRepPresupuesto.AdoFcConvenios.Recordset = tFc_convenios
  
  Set tFc_estructura_programatica = New ADODB.Recordset
  If tFc_estructura_programatica.State = 1 Then tFc_estructura_programatica.Close
   tFc_estructura_programatica.Open "SELECT * FROM Fc_estructura_programatica ", db, adOpenDynamic, adLockReadOnly
  Set frmListaEstProg.adoEstrProg.Recordset = tFc_estructura_programatica
End Sub

Private Sub showEtiquetas(mostrar As Boolean)
  If mostrar Then
    lblFuente.Visible = True
    lblOrg.Visible = True
    lblConv.Visible = True
    lblEstr.Visible = True
    lblPartida.Visible = True
    dcmFte_codigo.Visible = True
    cdmOrganismo.Visible = True
    dtcboconvenio.Visible = True
    txtProg.Visible = True
    txtSubProg.Visible = True
    txtProy.Visible = True
    txtAct.Visible = True
    butEstProg.Visible = True
    txtPartida.Visible = True
  Else
    lblFuente.Visible = False
    lblOrg.Visible = False
    lblConv.Visible = False
    lblEstr.Visible = False
    lblPartida.Visible = False
    dcmFte_codigo.Visible = False
    cdmOrganismo.Visible = False
    dtcboconvenio.Visible = False
    txtProg.Visible = False
    txtSubProg.Visible = False
    txtProy.Visible = False
    txtAct.Visible = False
    butEstProg.Visible = False
    txtPartida.Visible = False
  End If
End Sub

Private Sub opt_rep001_comp_dev_Click()
  Call SetControles(False, True)
End Sub

Private Sub optRep001_Click()
  Call SetControles(True, False)
End Sub

Private Sub optRep002_Click()
  Call SetControles(False, True)
End Sub

Private Sub optRep002b_Click()
  FrameTipo.Visible = False
End Sub

Private Sub optRep002_financiero_Click()
  Call SetControles(False, True)
End Sub

Private Sub optRep002Finanzas_Click()
  Call SetControles(False, True)
End Sub

Private Sub optRep003_Click()
  Call SetControles(False, False)
End Sub

Private Sub optRep004_Click()
  Call SetControles(False, False)
End Sub

Private Sub optRep005_Click()
  Call SetControles(False, False)
End Sub

Private Sub optRep006_Click()
  Call SetControles(False, False)
End Sub

Private Sub optRep007_Click()
  Call SetControles(False, False)
End Sub

Private Sub optRep008_Click()
  Call SetControles(False, False)
End Sub

Private Sub RepVsLeyFinanciador(tipoRep As String, ArchRep As String, titulo1 As String)
  CryRep002_financiador.ReportFileName = App.Path & ArchRep
  CryRep002_financiador.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryRep002_financiador.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryRep002_financiador.StoredProcParam(2) = tipoRep
  Call setParametros(CryRep002_financiador)
  CryRep002_financiador.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
  CryRep002_financiador.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  CryRep002_financiador.Formulas(2) = "conDetalle = " & IIf(optSi.Value = True, "true", "false")

  IResult = CryRep002_financiador.PrintReport
  If IResult <> 0 Then
    MsgBox CryRep002_financiador.LastErrorNumber & " : " & CryRep002_financiador.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub SetControles(tipo, conDet As Boolean)
  FrameTipo.Visible = tipo
  FrameConDet.Visible = conDet
End Sub

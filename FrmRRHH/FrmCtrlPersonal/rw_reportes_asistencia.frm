VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rw_reportes_asistencia 
   BackColor       =   &H00000000&
   Caption         =   "SOFIA"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000006&
      Height          =   780
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   8400
      TabIndex        =   13
      Top             =   0
      Width           =   8460
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "rw_reportes_asistencia.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   15
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   60
         Width           =   1455
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1800
         Picture         =   "rw_reportes_asistencia.frx":08CD
         ScaleHeight     =   615
         ScaleWidth      =   1365
         TabIndex        =   14
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   60
         Width           =   1365
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REPORTES ASISTENCIA"
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
         Left            =   4410
         TabIndex        =   16
         Top             =   180
         Width           =   3705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   8100
      Begin VB.OptionButton optrep005 
         BackColor       =   &H00000000&
         Caption         =   "RESUMEN ASISTENCIA "
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   510
         Width           =   3420
      End
      Begin VB.OptionButton optRep004 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA DETALLE ASISTENCIA"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   3285
      End
      Begin Crystal.CrystalReport CryReporte2 
         Left            =   7680
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8100
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "TODAS LAS PLANILLAS"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   5880
         TabIndex        =   12
         Top             =   960
         Width           =   2115
      End
      Begin VB.ComboBox cbo_mes_rep 
         Height          =   315
         ItemData        =   "rw_reportes_asistencia.frx":108F
         Left            =   4320
         List            =   "rw_reportes_asistencia.frx":10B7
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmb_gestion_rep 
         Height          =   315
         ItemData        =   "rw_reportes_asistencia.frx":1120
         Left            =   1920
         List            =   "rw_reportes_asistencia.frx":1145
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   630
      End
      Begin MSDataListLib.DataCombo dtc_rep_det 
         Bindings        =   "rw_reportes_asistencia.frx":118B
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_descripcion"
         BoundColumn     =   "planilla_codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_rep_cod 
         Bindings        =   "rw_reportes_asistencia.frx":11A7
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_codigo"
         BoundColumn     =   "planilla_codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos_rep 
         Height          =   330
         Left            =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "Ado_datos_rep"
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
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Caption         =   "GESTIÓN"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   380
         Width           =   735
      End
      Begin VB.Label Label33 
         BackColor       =   &H00000000&
         Caption         =   "MES"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Label34 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   975
         Width           =   855
      End
   End
End
Attribute VB_Name = "rw_reportes_asistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnImprimir_Click()
If optRep004.Value = True Or optrep005.Value = True Then
If cmb_gestion_rep.Text = "" Or cbo_mes_rep.Text = "" Or dtc_rep_cod.Text = "" Or dtc_rep_det.Text = "" Then
 sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
Else
    CryReporte2.Reset
    CryReporte2.WindowState = crptMaximized
    CryReporte2.WindowShowSearchBtn = True
    CryReporte2.WindowShowRefreshBtn = True
    CryReporte2.WindowShowPrintSetupBtn = True
   If optRep004.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_asistencia.rpt")
   End If
    'PLANILLA ASISTENCIA RESUMEN
   If optrep005.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_asistencia_totales.rpt")
   End If
End If
Else
sino = MsgBox("Elija un REPORTE por favor", vbCritical, "Atención")
End If
End Sub

Private Sub BtnSalir_Click()
  Unload Me
End Sub


Private Sub cbo_mes_rep_Click()
 txt_mes.Text = cbo_mes_rep.ListIndex
    txt_mes.Text = Val(txt_mes.Text) + 1
End Sub

Private Sub dtc_rep_cod_Click(Area As Integer)
   dtc_rep_det.BoundText = dtc_rep_cod.BoundText
    Option1.Value = False
End Sub

Private Sub dtc_rep_det_Click(Area As Integer)
 dtc_rep_cod.BoundText = dtc_rep_det.BoundText
    Option1.Value = False
End Sub

Private Sub Form_Load()
Call llenar_datos
cmb_gestion_rep.Text = Year(Date)
cbo_mes_rep.Text = UCase(MonthName(Month(Date)))
txt_mes.Text = Month(Date)

End Sub
Private Sub llenar_datos()
  Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "SELECT * FROM rc_planilla_grupo", db, adOpenStatic
    Set Ado_datos_rep.Recordset = rs_aux7
    dtc_rep_det.BoundText = dtc_rep_cod.BoundText
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
dtc_rep_cod.Text = "%"
dtc_rep_det.Text = "TODAS LAS PLANILLAS"
Else
dtc_rep_cod.Text = ""
dtc_rep_det.Text = ""
End If
End Sub

Private Sub Reportes2(ArchRep As String)
''Private Sub Reportes(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte2.ReportFileName = App.Path & ArchRep
'  CryReporte2.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryReporte2.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryReporte2.StoredProcParam(0) = tipoRep

   CryReporte2.StoredProcParam(0) = cmb_gestion_rep.Text
  If dtc_rep_cod.Text = "" Then
        CryReporte2.StoredProcParam(1) = "%"
  Else
        CryReporte2.StoredProcParam(1) = dtc_rep_cod.Text
  End If
   CryReporte2.StoredProcParam(2) = txt_mes.Text

  iResult = CryReporte2.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte2.LastErrorNumber & " : " & CryReporte2.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub


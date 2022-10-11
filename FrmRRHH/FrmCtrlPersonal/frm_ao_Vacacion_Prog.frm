VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_Vacacion_Prog 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Personal - File Funcionario - Programacion de Vacaciones"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9225
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ao_Vacacion_Prog.frx":0000
   ScaleHeight     =   2805
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_dias_vac 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7560
      MaxLength       =   80
      TabIndex        =   38
      Top             =   1920
      Width           =   855
   End
   Begin VB.PictureBox Frame2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_Vacacion_Prog.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   8955
      TabIndex        =   28
      Top             =   120
      Width           =   9015
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   2760
         Picture         =   "frm_ao_Vacacion_Prog.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Ver Contrato PDF"
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cargar"
         Height          =   680
         Left            =   1920
         Picture         =   "frm_ao_Vacacion_Prog.frx":D67D8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Carga Contrato en PDF"
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "frm_ao_Vacacion_Prog.frx":D6B60
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1080
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_Vacacion_Prog.frx":D6D6A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROGRAMACION DE VACACIONES"
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
         Left            =   3585
         TabIndex        =   31
         Top             =   240
         Width           =   5265
      End
   End
   Begin VB.Frame FraProyecto 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
      Begin VB.TextBox txt_empresa 
         Height          =   285
         Left            =   0
         MaxLength       =   80
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt_num_mes 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   5280
         MaxLength       =   80
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   3720
         MaxLength       =   80
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Txt02 
         Height          =   315
         ItemData        =   "frm_ao_Vacacion_Prog.frx":D6F74
         Left            =   7200
         List            =   "frm_ao_Vacacion_Prog.frx":D6FB4
         TabIndex        =   23
         Text            =   "1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt10 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         MaxLength       =   80
         TabIndex        =   22
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt09 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   21
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt08 
         Height          =   285
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   20
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         Height          =   315
         ItemData        =   "frm_ao_Vacacion_Prog.frx":D7008
         Left            =   1320
         List            =   "frm_ao_Vacacion_Prog.frx":D7294
         TabIndex        =   19
         Top             =   840
         Width           =   900
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   840
         MaxLength       =   80
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   795
         MaxLength       =   80
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox txt01 
         Height          =   315
         ItemData        =   "frm_ao_Vacacion_Prog.frx":D77A4
         Left            =   3120
         List            =   "frm_ao_Vacacion_Prog.frx":D77CC
         TabIndex        =   1
         Text            =   "ENERO"
         Top             =   840
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   45678593
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txt03 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   45678593
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txt05 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   45678594
         CurrentDate     =   0.333333333333333
         MinDate         =   4.16666666666667E-02
      End
      Begin MSComCtl2.DTPicker txt04 
         Height          =   315
         Left            =   3480
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   45678593
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker Txt06 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   45678594
         CurrentDate     =   0.770833333333333
         MaxDate         =   0.999305555555556
         MinDate         =   4.16666666666667E-02
      End
      Begin MSComCtl2.DTPicker DtcFec_Fin 
         Height          =   315
         Left            =   6720
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   45678593
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         DataSource      =   "frmBeneficiario.AdoPermiso"
         Height          =   315
         Left            =   6600
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         DataSource      =   "frmBeneficiario.AdoPermiso"
         Height          =   315
         Left            =   6600
         TabIndex        =   9
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker txt07 
         Height          =   315
         Left            =   6720
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   45678594
         CurrentDate     =   0.967326388888889
         MaxDate         =   0.999988425925926
         MinDate         =   4.16666666666667E-02
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes:"
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
         Height          =   240
         Index           =   3
         Left            =   2640
         TabIndex        =   36
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Días Programados"
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
         Height          =   240
         Index           =   1
         Left            =   5640
         TabIndex        =   35
         Top             =   840
         Width           =   1710
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado"
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
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas Programadas:                                     Minutos Programados:"
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
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   3600
         Visible         =   0   'False
         Width           =   5580
      End
      Begin VB.Label lblARCH 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   6600
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "                                                       Aprobado                              Nombre Archivo:"
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
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   6240
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Benef"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "SW"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   1920
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión:"
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
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Fecha:                                            Hasta Fecha:                                       Fecha Reincorporacion: "
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
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   8415
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   $"frm_ao_Vacacion_Prog.frx":D7835
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
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Visible         =   0   'False
         Width           =   8280
      End
   End
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   120
      Top             =   5280
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Clasificador"
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
End
Attribute VB_Name = "frm_ao_Vacacion_Prog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_Clasificador As New ADODB.Recordset
Dim rs_correlativo As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim nomb2 As String
Dim hora01, hora02, hora03, hora04 As String
Dim fecha1 As String
Dim MES_IN As Integer
Dim ANO_IN As Integer
Dim DIA_HOY As Integer
Dim DIA_IN As Integer
Dim VAR_GES As Integer
Public sel As Integer


Private Sub cmdCancel_Click()
    'cancela la edicion de datos
    Para_Aceptado = "N"
    Unload Me
    'Me.Hide
End Sub

Private Sub cmdOk_Click()
 
'TxtGestion.Text = Year(txt03.Value)
'DTPFec_Inicio.Value = txt03.Value
'Txt01.Text = UCase(MonthName(Month(txt03.Value)))
 'acepta las modificaciones realizadas
' Dim NoDias, NoHoras, NoMin As Integer
' Dim DifHr1, DifHr2 As Integer
' If ValidaMontos Then
'   Dim SQLS As String
'   SQLS = ""
'   If txtSW = "ADD" Then
'      'hora01 = Format(txt05.Value, "HH:mm:ss")
'      'hora02 = Format(Txt06.Value, "HH:mm:ss")
'      'hora03 = Format(txt07.Value, "HH:mm:ss")
'      'hora04 = Format(txt08.Value, "HH:mm:ss")
'      'DB.Execute "Insert INTO ro_ControlAsistencia (beneficiario_codigo, Fecha_control, mes_control, dia_control, FechaDesde, FechaHasta, fecha_reincorporacion, horadesde, horahasta, Hora_reincorporacion, ges_gestion, dias_permiso, horas_permiso, minutos_permiso, estado_codigo, fecha_registro, usr_usuario) "
'      'Values ('" & txtBenef.Text & "', '" & DTPFec_Inicio.Value & "', '" & txt01 & "', '" & Txt02 & "', '" & txt03 & "', '" & txt04 & "', '" & DtcFec_Fin & "', '" & hora01 & "', '" & hora02 & "', '" & hora03 & "', '" & TxtGestion & "', '" & txt08 & "', '" & txt09 & "', '" & txt10 & "', 'N', '" & Date & "', '" & GlUsuario & "') "
'       rw_ficha_rrhh.Ado_VacacionesProg.Recordset.AddNew
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("beneficiario_codigo").Value = txtBenef.Text
'      'rw_ficha_rrhh.Ado_VacacionesProg.Recordset("gestion").Value = TxtGestion.Text
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value = TxtGestion.Text
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("mes_control") = txt01.Text
'      Set rs_correlativo = New ADODB.Recordset
'      rs_correlativo.Open "select * from ro_vacaciones_programadas WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "'  ", db, adOpenKeyset, adLockOptimistic
'      If rs_correlativo.RecordCount > 0 Then
'            rw_ficha_rrhh.Ado_VacacionesProg.Recordset!CORREL = rs_correlativo.RecordCount
'      Else
'            rw_ficha_rrhh.Ado_VacacionesProg.Recordset!CORREL = 1
'      End If
'      'rw_ficha_rrhh.Ado_VacacionesProg.Recordset!ARCHIVO = "Cargar_Archivo"
'      'rw_ficha_rrhh.Ado_VacacionesProg.Recordset!ARCHIVO_NOMB = Trim(rw_ficha_rrhh.adoLista.Recordset!beneficiario_beneficiario_iniciales) & "_Licencias_" & rw_ficha_rrhh.AdoPermiso.Recordset!CORREL & ".pdf"
'      txtEstado.Text = "REG"
'   End If
'      'rw_ficha_rrhh.Ado_VacacionesProg.Recordset("TipoPermiso").Value = Dtc_Par.Text
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_registro").Value = DTPFec_Inicio.Value
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("dias_Programados").Value = Txt02.Text
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_ini_Prog").Value = txt03.Value
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_fin_Prog").Value = txt04.Value
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_reincoporacion").Value = DtcFec_Fin.Value
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horadesde").Value = Format(txt05.Value, "HH:mm:ss")
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horahasta").Value = Format(Txt06.Value, "HH:mm:ss")
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("Hora_reincorporacion").Value = Format(txt07.Value, "HH:mm:ss")
'      NoDias = DateDiff("d", rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_ini_Prog").Value, rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_fin_Prog").Value)
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("dias_Programados").Value = Txt02
'      GlHora1 = "08:00"
'      DifHr1 = DateDiff("h", CDate(GlHora1), rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horadesde").Value)
'      'DifHr1 = DateDiff("h", CDate("08:00"), rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horadesde").Value)
'      GlHora2 = "14:30"
'      'DifHr2 = 4 - DateDiff("h", CDate(GlHora2), rw_ficha_rrhh.AdoPermiso.Recordset("horahasta").Value)
'      DifHr2 = 4 - DateDiff("h", CDate(GlHora2), CDate(GlHora2))
'      If DifHr1 > 0 Then
'         DifHr1 = DifHr1
'      Else
'         DifHr1 = 0
'      End If
'      If DifHr2 > 0 Then
'         DifHr2 = DifHr2
'      Else
'         DifHr2 = 0
'      End If
'      If NoDias < 0 Then NoDias = NoDias * (-1)
'      NoHoras = (NoDias * 8) - (DifHr1 + DifHr2)
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horas_Programadas").Value = NoHoras     'txt09.Text
'      NoMin = NoHoras / 60
'
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("minutos_programados").Value = NoMin     'txt10.Text
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("numero_memoranda").Value = rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value + "/" + CStr(rw_ficha_rrhh.Ado_VacacionesProg.Recordset("CORREL").Value)
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "REG", txtEstado.Text)
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_registro") = Date
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("usr_usuario").Value = glusuario
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset("dias_utilizados") = "0"
'      rw_ficha_rrhh.Ado_VacacionesProg.Recordset.Update
'   'End If
'   Para_Aceptado = "S"
'   rw_ficha_rrhh.opciones
'   Unload Me
'   'Me.Hide
' End If

'NUEVO

'sino = MsgBox("¿Desea que el sistema genere autamaticamente Las planillas?", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
Select Case txt01.Text
    Case "ENERO"
        txt_num_mes.Text = "1"
    Case "FEBRERO"
        txt_num_mes.Text = "2"
    Case "MARZO"
        txt_num_mes.Text = "3"
    Case "ABRIL"
        txt_num_mes.Text = "4"
    Case "MAYO"
        txt_num_mes.Text = "5"
    Case "JUNIO"
        txt_num_mes.Text = "6"
    Case "JULIO"
        txt_num_mes.Text = "7"
    Case "AGOSTO"
        txt_num_mes.Text = "8"
    Case "SEPTIEMBRE"
        txt_num_mes.Text = "9"
    Case "OCTUBRE"
        txt_num_mes.Text = "10"
    Case "NOVIEMBRE"
        txt_num_mes.Text = "11"
    Case "DICIEMBRE"
        txt_num_mes.Text = "12"
    Case Else
        MsgBox ("EL MES" & txt01.Text & " no existe")
    Exit Sub
End Select

If sel = 1 Then
        
       If rs_aux6.State = 1 Then rs_aux6.Close
       rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE estado_codigo <> 'ANL' AND beneficiario_codigo = '" & txtBenef.Text & "' and codigo_empresa = '" & txt_empresa.Text & "'", db, adOpenStatic
       rw_ficha_rrhh.ProgressBar1.Visible = True
       If rs_aux6.RecordCount = 0 Then
       sino = MsgBox("Existe un error con esta persona, no deberia estar en la lista", vbCritical, "Error")
       Exit Sub
       End If
       With rw_ficha_rrhh.ProgressBar1
        .Max = rs_aux6.RecordCount
        .Min = 0
        .Value = 0
       End With
      'ProgressBar1.Max =
       rw_ficha_rrhh.ProgressBar1.Value = rw_ficha_rrhh.ProgressBar1.Value + 1
       If rs_aux5.State = 1 Then rs_aux5.Close
       rs_aux5.Open "select * from ro_vacaciones_programadas where ges_gestion = '" & TxtGestion.Text & "' AND beneficiario_codigo = '" & txtBenef.Text & "' and codigo_empresa = " & txt_empresa.Text & "", db, adOpenKeyset, adLockOptimistic, adCmdText
       VAR_GES = DateDiff("yyyy", rs_aux6!fecha_ingreso, Date)
       DIA_IN = Day(rs_aux6!fecha_ingreso)
       MES_IN = Month(rs_aux6!fecha_ingreso)
       ANO_IN = Year(rs_aux6!fecha_ingreso)
       
        If txt_num_mes.Text < MES_IN Then
            VAR_GES = VAR_GES - 1
        End If
        'CAMBIO JQ-2016-OCT-26
        If (VAR_GES * 12) < DateDiff("m", rs_aux6!fecha_ingreso, Date) Then
            VAR_GES = VAR_GES + 1
        End If
        sino = MsgBox("Elija SI, para que el Sistema calcule los Días Programados..." & vbCrLf & "Elija NO, para guardar Dias Programados registrados... ", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            If rs_aux8.State = 1 Then rs_aux8.Close
            rs_aux8.Open "select * from rc_vacaciones_parametro where parametro_inicio <= " & VAR_GES & " and parametro_fin >= " & VAR_GES & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
            If rs_aux8.RecordCount > 0 Then
                   txt_dias_vac.Text = rs_aux8!dias_vacacion
            Else
                    txt_dias_vac.Text = "0"
            End If
        End If
       If rs_aux5.RecordCount = 0 Then
      
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset.AddNew
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("beneficiario_codigo").Value = txtBenef.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value = TxtGestion.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("mes_control") = txt01.Text
        
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_ini_Prog").Value = "01/01/" & Year(Date)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_fin_Prog").Value = "01/" & Txt02.Text & "/" & Year(Date)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("numero_memoranda").Value = rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value + "/" + CStr(rw_ficha_rrhh.Ado_VacacionesProg.Recordset("CORREL").Value)
         rw_ficha_rrhh.Ado_VacacionesProg.Recordset("dias_Programados").Value = txt_dias_vac.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_reincoporacion").Value = "01/" & Val(Txt02.Text) + 1 & "/" & Year(Date)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horadesde").Value = "08:00:00"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horahasta").Value = "08:00:00"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("Hora_reincorporacion").Value = "08:00:00"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("codigo_empresa").Value = txt_empresa.Text
        
        
        
         Set rs_correlativo = New ADODB.Recordset
         rs_correlativo.Open "select * from ro_vacaciones_programadas WHERE beneficiario_codigo = '" & txtBenef.Text & "' and codigo_empresa = " & txt_empresa.Text & "", db, adOpenKeyset, adLockOptimistic
        If rs_correlativo.RecordCount > 0 Then
                    rw_ficha_rrhh.Ado_VacacionesProg.Recordset!CORREL = rs_correlativo.RecordCount
              Else
                    rw_ficha_rrhh.Ado_VacacionesProg.Recordset!CORREL = 1
        End If
        
        txtEstado.Text = "REG"
  
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("minutos_programados").Value = "0"     'txt10.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("numero_memoranda").Value = rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value + "/" + CStr(rw_ficha_rrhh.Ado_VacacionesProg.Recordset("CORREL").Value)
        
        
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "REG", txtEstado.Text)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_registro") = Date
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("usr_usuario").Value = glusuario
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("dias_utilizados") = "0"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("Dias_Pendientes") = "0"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("codigo_empresa") = txt_empresa.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset.Update
            
       Else
       
        If rw_ficha_rrhh.Ado_VacacionesProg.Recordset("estado_codigo") = "REG" Or glusuario = "VPAREDES" Then
            db.Execute "UPDATE ro_vacaciones_programadas SET dias_Programados = " & txt_dias_vac.Text & " , Dias_Pendientes = " & Val(txt_dias_vac.Text) - rs_aux5!dias_utilizados & ", Mes_control = '" & txt01.Text & "' where beneficiario_codigo = '" & txtBenef.Text & "' and ges_gestion = '" & TxtGestion.Text & "'"
        End If
       
       End If

   rw_ficha_rrhh.ProgressBar1.Visible = False
   Para_Aceptado = "S"
   rw_ficha_rrhh.opciones
   Call rw_ficha_rrhh.abrirtabla
   Unload Me
Else
       rw_ficha_rrhh.ProgressBar1.Visible = True
       If rs_aux6.State = 1 Then rs_aux6.Close
       rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE estado_codigo <> 'ANL' and codigo_empresa = " & txt_empresa.Text & ", db, adOpenStatic 'order by beneficiario_denominacion"
      'rs_aux6.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic
       rs_aux6.MoveFirst
       With rw_ficha_rrhh.ProgressBar1
        .Max = rs_aux6.RecordCount
        .Min = 0
        .Value = 0
       End With
      'ProgressBar1.Max =
       While Not rs_aux6.EOF
       rw_ficha_rrhh.ProgressBar1.Value = rw_ficha_rrhh.ProgressBar1.Value + 1
       If rs_aux5.State = 1 Then rs_aux5.Close
       rs_aux5.Open "select * from ro_vacaciones_programadas where ges_gestion = '" & TxtGestion.Text & "' AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' and codigo_empresa = " & txt_empresa.Text & ", db, adOpenKeyset, adLockOptimistic, adCmdText"
        VAR_GES = DateDiff("yyyy", rs_aux6!fecha_ingreso, Date)
       DIA_IN = Day(rs_aux6!fecha_ingreso)
       MES_IN = Month(rs_aux6!fecha_ingreso)
       ANO_IN = Year(rs_aux6!fecha_ingreso)
        If txt_num_mes.Text < MES_IN Then
                VAR_GES = VAR_GES - 1
        End If
        sino = MsgBox("Elija SI, para que el Sistema calcule los Días Programados..." & vbCrLf & "Elija NO, para guardar Dias Programados registrados... ", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            If rs_aux8.State = 1 Then rs_aux8.Close
            rs_aux8.Open "select * from rc_vacaciones_parametro where parametro_inicio <= " & VAR_GES & " and parametro_fin >= " & VAR_GES & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
            If rs_aux8.RecordCount > 0 Then
               txt_dias_vac.Text = rs_aux8!dias_vacacion
            Else
                txt_dias_vac.Text = "0"
            End If
        End If
       If rs_aux5.RecordCount = 0 Then
      
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset.AddNew
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("beneficiario_codigo").Value = rs_aux6!beneficiario_codigo
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value = TxtGestion.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("mes_control") = txt01.Text
        
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_ini_Prog").Value = "01/01/" & Year(Date)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_fin_Prog").Value = "01/" & Txt02.Text & "/" & Year(Date)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("numero_memoranda").Value = rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value + "/" + CStr(rw_ficha_rrhh.Ado_VacacionesProg.Recordset("CORREL").Value)
         rw_ficha_rrhh.Ado_VacacionesProg.Recordset("dias_Programados").Value = txt_dias_vac.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_reincoporacion").Value = "01/" & Val(Txt02.Text) + 1 & "/" & Year(Date)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horadesde").Value = "08:00:00"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("horahasta").Value = "08:00:00"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("Hora_reincorporacion").Value = "08:00:00"
 
        Set rs_correlativo = New ADODB.Recordset
        rs_correlativo.Open "select * from ro_vacaciones_programadas WHERE beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        If rs_correlativo.RecordCount > 0 Then
                    rw_ficha_rrhh.Ado_VacacionesProg.Recordset!CORREL = rs_correlativo.RecordCount
              Else
                    rw_ficha_rrhh.Ado_VacacionesProg.Recordset!CORREL = 1
        End If

        txtEstado.Text = "REG"
  
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("minutos_programados").Value = "0"     'txt10.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("numero_memoranda").Value = rw_ficha_rrhh.Ado_VacacionesProg.Recordset("ges_gestion").Value + "/" + CStr(rw_ficha_rrhh.Ado_VacacionesProg.Recordset("CORREL").Value)
  
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "REG", txtEstado.Text)
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("fecha_registro") = Date
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("usr_usuario").Value = glusuario
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("dias_utilizados") = "0"
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("Dias_Pendientes") = txt_dias_vac.Text 'Dias_Pendientes
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset("codigo_empresa") = txt_empresa.Text
        rw_ficha_rrhh.Ado_VacacionesProg.Recordset.Update
      
       Else
        If rw_ficha_rrhh.Ado_VacacionesProg.Recordset("estado_codigo") = "REG" Or glusuario = "VPAREDES" Then
            db.Execute "UPDATE ro_vacaciones_programadas SET dias_Programados = " & txt_dias_vac.Text & ", Dias_Pendientes = " & Val(txt_dias_vac.Text) - rs_aux5!dias_utilizados & " , mes_control = '" & txt01.Text & "' where beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' and ges_gestion = '" & TxtGestion.Text & "'  and codigo_empresa = " & txt_empresa.Text & ""
        End If
       End If
       rs_aux6.MoveNext
      Wend
 
   rw_ficha_rrhh.ProgressBar1.Visible = False
   Para_Aceptado = "S"
   rw_ficha_rrhh.opciones
   Call rw_ficha_rrhh.abrirtabla
   Unload Me
   
End If
 
End Sub

Function ValidaMontos()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
ValidaMontos = True
'If Val(Me.mskMonto) > Val(Me.mskMonto_pendiente) Then
'    ValidaMontos = False
'    MsgBox "El monto indicado sobrepasa el monto pendiente de pago", vbInformation
'    Me.mskMonto.SelStart = 0
'    Me.mskMonto.SelLength = Len(Me.mskMonto)
'    Me.mskMonto.SetFocus
'End If
    If txt01 = "" Then
        ValidaMontos = False
    End If
    If Txt02 = "" Then
        ValidaMontos = False
    End If
    If txt03 = "" Then
        ValidaMontos = False
    End If
    If txt04 = "" Then
        ValidaMontos = False
    End If
End Function


Private Sub cmdRefresh_Click()
' If lblARCH.Caption = "Cargar_Archivo" Then
'    MsgBox ("No Existe el Archivo Asociado al Contrato, debe Cargarlo ...")
' Else
'    'If GlServidor <> GlMaquina Then      ' "-" Then
'    If GlServidor = "SRVPRO" Then
'        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    Else
'        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    End If
' End If
End Sub

Private Sub CmdVerDisco_Click()
'  On Error GoTo Error_Sub
'
'  If rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo" Then
'     NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
'     Frmexporta.DirDestino.Path = NombreCarpeta
'     GlArch = "PRM"
'      'If GlServidor <> GlMaquina Then      ' "-" Then
'      If GlServidor = "SRVPRO" Then
'         DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
'      Else
'         DirCto = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = DirCto
'     Frmexporta.Show vbModal
'  Else
''    MsgBox ""
'     sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'     If sino = vbYes Then
'        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
'        Frmexporta.DirDestino.Path = NombreCarpeta
'        GlArch = "PRM"
'        'If GlServidor <> GlMaquina Then      ' "-" Then
'        If GlServidor = "SRVPRO" Then
'           DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
'        Else
'           DirCto = NombreCarpeta
'        End If
'        Frmexporta.DirDestino2.Path = DirCto
'        Frmexporta.Show vbModal
'     End If
'  End If
'
'  Exit Sub
'Error_Sub:
'  MsgBox Err.Description, vbCritical

End Sub

Private Sub Dtc_Par_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Par.BoundText
End Sub

Private Sub Dtc_ParDes_Click(Area As Integer)
    Dtc_Par.BoundText = Dtc_ParDes.BoundText
End Sub

Private Sub Form_Load()
txt03.Value = Date
txt04.Value = Date
txt01.Text = UCase(MonthName(Month(Date)))
DtcFec_Fin.Value = Date
If glProceso = "CONSULTORIA" Then
    Me.Caption = "Consultoría - Captura de datos personales"
Else
    Me.Caption = "Recursos Humanos - Captura de datos personales"
End If
Para_Aceptado = "N"
'LOS DATOS PERSONALES SE CARGAN EN EL FORMULARIO QUE LO LLAMA
'AQUI SE JALA LOS MONTOS REGISTRADOS EN AO_ADJUDICA_C
Dim Xmbe As Double, Xmde As Double, Xmbn As Double, Xmdn As Double
Dim XAbe As Double, XAde As Double, XAbn As Double, XAdn As Double
'With ac_Adjudicacion_c.adoSec.Recordset
'    Me.labTipoMoneda = !tipo_moneda
'    DE.dbo_edCmprSumaMontosLimiteBen1 !ges_gestion, !codigo_unidad, !codigo_solicitud, !numero_consultoria, Xmbe, Xmde, Xmbn, Xmdn, XAbe, XAde, XAbn, XAdn
'    If !tipo_moneda = "$US" Then
'        Me.mskMonto = Round(!monto_dolares_ext + !monto_dolares_nal, 2)
'        Me.mskMonto_ext = !monto_dolares_ext
'        Me.mskMonto_nal = !monto_dolares_nal
'        Me.mskMonto_limite = Xmde + Xmdn
'        Me.mskMonto_pendiente = Round(Xmde + Xmdn - XAde - XAdn + Val(Me.mskMonto), 2)
'        Me.labPorcExt = CStr(Format(Xmde / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        Me.labPorcNal = CStr(Format(Xmdn / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        Me.mskMonto = Round(!monto_dolares_ext + !monto_dolares_nal, 2)
'    Else
'        Me.mskMonto = Round(!monto_bolivianos_ext + !monto_bolivianos_nal)
'        Me.mskMonto_ext = !monto_bolivianos_ext
'        Me.mskMonto_nal = !monto_bolivianos_nal
'        Me.mskMonto_limite = Xmbe + Xmbn
'        Me.mskMonto_pendiente = Xmbe + Xmbn - XAbe - XAbn + Val(Me.mskMonto)
'        If Val(Me.mskMonto_limite) = 0 Then
'            Me.labPorcExt = "0 %"
'            Me.labPorcNal = "0 %"
'        Else
'            Me.labPorcExt = CStr(Format(Xmbe / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'            Me.labPorcNal = CStr(Format(Xmbn / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        End If
'        Me.mskMonto = Round(!monto_bolivianos_ext + !monto_bolivianos_nal)
'    End If
'End With
    
    Set rs_Clasificador = New ADODB.Recordset
    rs_Clasificador.Open "SELECT * FROM rc_TipoPermiso ORDER BY descripcion ", db, adOpenStatic
    Set Ado_Clasificador.Recordset = rs_Clasificador
    

'mskMonto.SetFocus
End Sub

'Private Sub mskMonto_KeyPress(KeyAscii As Integer)
'If Val(Chr(KeyAscii)) <> 0 Or Chr(KeyAscii) = "." Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "0" Or KeyAscii = 8 Then
'    'asdfasdf
'Else
'    KeyAscii = 0
'End If
'End Sub

Private Sub txt01_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub

Private Sub TxtGestion_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
'Txt01.Text = ""
'Txt01.Text = UCase(MonthName(Month(Date)))
KeyAscii = 0
Else
Exit Sub
End If
End Sub

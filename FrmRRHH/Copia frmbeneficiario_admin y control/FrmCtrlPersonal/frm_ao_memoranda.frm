VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_memoranda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Personal - File Funcionario - Memorandas"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ao_memoranda.frx":0000
   ScaleHeight     =   4800
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_memoranda.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   7635
      TabIndex        =   19
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   2680
         Picture         =   "frm_ao_memoranda.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ver Contrato PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cargar"
         Height          =   680
         Left            =   1920
         Picture         =   "frm_ao_memoranda.frx":D67D8
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Carga Contrato en PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "frm_ao_memoranda.frx":D6B60
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1080
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_memoranda.frx":D6D6A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MEMORANDAS"
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
         Left            =   4470
         TabIndex        =   22
         Top             =   240
         Width           =   2295
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
      ForeColor       =   &H00008000&
      Height          =   3705
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
      Begin VB.OptionButton Optsi 
         BackColor       =   &H0000FF00&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   3240
         Width           =   615
      End
      Begin VB.OptionButton Optno 
         BackColor       =   &H000000FF&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   34
         Top             =   3255
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.TextBox txt_mes_afectado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         MaxLength       =   80
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cbo_dias 
         DataField       =   "ges_gestion"
         Height          =   315
         ItemData        =   "frm_ao_memoranda.frx":D6F74
         Left            =   240
         List            =   "frm_ao_memoranda.frx":D6FDB
         TabIndex        =   31
         Text            =   "0"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txt_mes 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         MaxLength       =   80
         TabIndex        =   30
         Text            =   "0"
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_correl 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         MaxLength       =   80
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox TxtGestion2 
         DataField       =   "ges_gestion"
         Height          =   315
         ItemData        =   "frm_ao_memoranda.frx":D7063
         Left            =   240
         List            =   "frm_ao_memoranda.frx":D7079
         TabIndex        =   26
         Text            =   "2016"
         Top             =   4080
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   3720
         MaxLength       =   80
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6480
         MaxLength       =   80
         TabIndex        =   15
         Text            =   "0"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt09 
         Height          =   1485
         Left            =   2040
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2040
         Width           =   5415
      End
      Begin VB.TextBox txt08 
         Height          =   285
         Left            =   240
         MaxLength       =   80
         TabIndex        =   13
         Text            =   "0"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         Height          =   315
         ItemData        =   "frm_ao_memoranda.frx":D70A1
         Left            =   6480
         List            =   "frm_ao_memoranda.frx":D70B7
         TabIndex        =   12
         Text            =   "2016"
         Top             =   3240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   840
         MaxLength       =   80
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   2280
         MaxLength       =   80
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         MaxLength       =   80
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox txt01 
         Height          =   315
         ItemData        =   "frm_ao_memoranda.frx":D70DF
         Left            =   2880
         List            =   "frm_ao_memoranda.frx":D7107
         TabIndex        =   1
         Text            =   "ENERO"
         Top             =   4080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   99614721
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DtcFec_Fin 
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   99614721
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         Bindings        =   "frm_ao_memoranda.frx":D7170
         DataField       =   "tipo_memo"
         DataSource      =   "frmBeneficiario_control.Ado_Memo"
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
         ListField       =   "tipo_memo"
         BoundColumn     =   "tipo_memo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         Bindings        =   "frm_ao_memoranda.frx":D718F
         DataField       =   "tipo_memo"
         DataSource      =   "frmBeneficiario_control.Ado_Memo"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "tipo_memo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_estado 
         Bindings        =   "frm_ao_memoranda.frx":D71AE
         DataField       =   "tipo_memo"
         DataSource      =   "frmBeneficiario_control.Ado_Memo"
         Height          =   315
         Left            =   2400
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
         ListField       =   "estado_baja"
         BoundColumn     =   "tipo_memo"
         Text            =   ""
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Descontar en Planilla"
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
         Left            =   120
         TabIndex        =   33
         Top             =   3000
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Días Sanción"
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
         TabIndex        =   28
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label txt_memo 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFF80&
         Height          =   310
         Left            =   3980
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblARCH 
         Alignment       =   1  'Right Justify
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
         Height          =   310
         Left            =   5520
         TabIndex        =   17
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Memoranda                                          Nro. Memo               Nombre de Archivo"
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
         TabIndex        =   16
         Top             =   360
         Width           =   7140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Elaboracion                      Fecha de Ejecución o Baja                                  "
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
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   6930
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión a Penalizar                            Mes a Penalizar                              Minutos Penalizar"
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
         Left            =   960
         TabIndex        =   7
         Top             =   3840
         Visible         =   0   'False
         Width           =   7365
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Sanción          Observaciones"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   3165
      End
   End
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   120
      Top             =   4800
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
Attribute VB_Name = "frm_ao_memoranda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_Clasificador As New ADODB.Recordset
Dim rs_correlativo As New ADODB.Recordset

Dim nomb2 As String
Dim hora01, hora02, hora03, hora04 As String
Dim fecha1 As String
Dim total As Double

Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset





Private Sub cbo_dias_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub

Private Sub cmdCancel_Click()
    'cancela la edicion de datos
    Para_Aceptado = "N"
    Unload Me
    'Me.Hide
End Sub

Private Sub cmdOk_Click()
 'acepta las modificaciones realizadas
 Dim NoDias, NoHoras, NoMin As Integer
 If Optsi.Value = True Then
 If txt08.Text = "0" And cbo_dias.Text = "0" Then
    sino = MsgBox("Debe llenar el MONTO o los DIAS a DESCONTAR", vbCritical, "ERROR")
    Exit Sub
 End If
 
  If txt08.Text = "" Or cbo_dias.Text = "" Then
    sino = MsgBox("Debe llenar el MONTO o los DIAS a DESCONTAR", vbCritical, "ERROR")
    Exit Sub
 End If
 
 
 End If
 TxtGestion = Year(DtcFec_Fin.Value)
 TxtGestion2.Text = TxtGestion.Text
 txt01 = Month(DtcFec_Fin.Value)
 txt_mes = Month(DtcFec_Fin.Value)
 txt_mes_afectado = Month(DtcFec_Fin.Value) + 1
 
 
 If ValidaMontos Then
   Dim SQLS As String
   SQLS = ""
   If txtSW = "ADD" Then
   rw_ficha_rrhh.Ado_Memo.Recordset.AddNew
      'hora01 = Format(txt05.Value, "HH:mm:ss")
      'hora02 = Format(Txt06.Value, "HH:mm:ss")
      'hora03 = Format(txt07.Value, "HH:mm:ss")
      'hora04 = Format(txt08.Value, "HH:mm:ss")
      'DB.Execute "Insert INTO ro_ControlAsistencia (codigo_beneficiario, Fecha_control, mes_control, dia_control, FechaDesde, FechaHasta, fecha_reincorporacion, horadesde, horahasta, Hora_reincorporacion, ges_gestion, dias_permiso, horas_permiso, minutos_permiso, estado_registro, fecha_registro, usr_usuario) "
      'Values ('" & txtBenef.Text & "', '" & DTPFec_Inicio.Value & "', '" & txt01 & "', '" & Txt02 & "', '" & txt03 & "', '" & txt04 & "', '" & DtcFec_Fin & "', '" & hora01 & "', '" & hora02 & "', '" & hora03 & "', '" & TxtGestion & "', '" & txt08 & "', '" & txt09 & "', '" & txt10 & "', 'N', '" & Date & "', '" & GlUsuario & "') "
      rw_ficha_rrhh.Ado_Memo.Recordset("beneficiario_codigo").Value = txtBenef.Text
      rw_ficha_rrhh.Ado_Memo.Recordset("ges_gestion").Value = TxtGestion.Text
      rw_ficha_rrhh.Ado_Memo.Recordset("mes_descuento") = txt01.Text
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ro_memorandas WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "'  ", db, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            rw_ficha_rrhh.Ado_Memo.Recordset!CORREL = rs_correlativo.RecordCount + 1
      Else
            rw_ficha_rrhh.Ado_Memo.Recordset!CORREL = 1
      End If
      rw_ficha_rrhh.Ado_Memo.Recordset!ARCHIVO = "Cargar_Archivo"
'     rw_ficha_rrhh.Ado_Memo.Recordset!ARCHIVO_MEM = Trim(rw_ficha_rrhh.Ado_datos.Recordset!iniciales) & "_Memos_" & rw_ficha_rrhh.Ado_Memo.Recordset!CORREL & ".pdf"
      txtEstado.Text = "REG"
      
      End If
      
      rw_ficha_rrhh.Ado_Memo.Recordset("tipo_memo").Value = Dtc_Par.Text
      
      If txt09.Text = "" Then
      rw_ficha_rrhh.Ado_Memo.Recordset("Observaciones").Value = Dtc_ParDes.Text
      Else
      rw_ficha_rrhh.Ado_Memo.Recordset("Observaciones").Value = UCase(txt09.Text)
      End If
      
      rw_ficha_rrhh.Ado_Memo.Recordset("fecha_memo").Value = DTPFec_Inicio.Value
      rw_ficha_rrhh.Ado_Memo.Recordset("fecha_aprobacion").Value = DtcFec_Fin.Value
      rw_ficha_rrhh.Ado_Memo.Recordset("monto").Value = IIf(txt08.Text = "", "0", txt08.Text)
      rw_ficha_rrhh.Ado_Memo.Recordset("minutos").Value = IIf(txt10.Text = "", "0", txt10.Text)
      rw_ficha_rrhh.Ado_Memo.Recordset("dias").Value = IIf(cbo_dias.Text = "", "0", cbo_dias.Text)
      rw_ficha_rrhh.Ado_Memo.Recordset("gestion_descuento").Value = TxtGestion2.Text
      rw_ficha_rrhh.Ado_Memo.Recordset("mes_descuento").Value = txt01.Text
      
      
      
      rw_ficha_rrhh.Ado_Memo.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "REG", txtEstado.Text)
      rw_ficha_rrhh.Ado_Memo.Recordset("fecha_registro") = Date
      rw_ficha_rrhh.Ado_Memo.Recordset("usr_codigo").Value = glusuario
      rw_ficha_rrhh.Ado_Memo.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
      If Optsi.Value = True Then
      rw_ficha_rrhh.Ado_Memo.Recordset("descuento_pla").Value = "SI"
      Else
      rw_ficha_rrhh.Ado_Memo.Recordset("descuento_pla").Value = "NO"
      End If
      rw_ficha_rrhh.Ado_Memo.Recordset.Update
      
   'End If
   Para_Aceptado = "S"
'   If dtc_estado.Text = "S" Then
'   db.Execute "update ro_personal_contratado set estado_codigo = 'ANL' WHERE beneficiario_codigo = '" & txtBenef.Text & "'"
'   End If
'
'    If Dtc_Par.Text = "SAD" Then
'    total = 0
'    If rs_datos.State = 1 Then rs_datos.Close
'     rs_datos.Open "select * from ro_pagos_cronograma_Detalle where ges_gestion = '" & TxtGestion.Text & "' AND mes_grupo = " & txt_mes.Text & " AND beneficiario_codigo = '" & txtBenef.Text & "'", db, adOpenKeyset, adLockOptimistic
'     If rs_datos.RecordCount <> 0 Then
'     If txt08.Text > 0 Then
'     total = rs_datos!otros_dsctos + txt08.Text
''     rs_datos!otros_dsctos = total
''     rs_datos!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!rciva + rs_datos2!otros_dsctos
'     End If
'
'
'
'     If cbo_dias.Text > 0 Then
'     If rs_datos1.State = 1 Then rs_datos1.Close
'     rs_datos1.Open "select * from ro_personal_contratado where beneficiario_codigo = '" & txtBenef.Text & "'", db, adOpenKeyset, adLockOptimistic
'     total = total + ((rs_datos1!beneficiario_haber_mensual / 30) * cbo_dias.Text)
'     total = total + rs_datos!otros_dsctos
''     rs_datos!otros_dsctos = total
''     rs_datos!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!rciva + rs_datos2!otros_dsctos
'     End If
'
'     If total > 0 Then
'     rs_datos!otros_dsctos = total
'     rs_datos!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!rciva + rs_datos2!otros_dsctos
'     End If
'
'     'db.Execute "update ro_pagos_cronograma_Detalle set otros_dsctos = " & total & "WHERE beneficiario_codigo = '" & txtBenef.Text & "' AND mes_grupo = " & txt_mes.Text & " AND ges_gestion = '" & TxtGestion.Text & "'"
'
 rw_ficha_rrhh.opciones
     End If
'
    
     
   
   
   Unload Me
   
   'Me.Hide
 
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
'    If Txt02 = "" Then
'        ValidaMontos = False
'    End If
'    If txt03 = "" Then
'        ValidaMontos = False
'    End If
'    If txt04 = "" Then
'        ValidaMontos = False
'    End If
End Function


Private Sub cmdRefresh_Click()
 If lblARCH.Caption = "Cargar_Archivo" Then
    MsgBox ("No Existe el Archivo Asociado al Contrato, debe Cargarlo ...")
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SRVPRO" Then
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!codigo_beneficiario) & "\MEMOS\" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!codigo_beneficiario) & "\MEMOS\" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub

Private Sub CmdVerDisco_Click()
  On Error GoTo Error_Sub
  
  If rw_ficha_rrhh.Ado_Memo.Recordset!ARCHIVO = "Cargar_Archivo" Then
     NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!codigo_beneficiario) & "\MEMOS\"
     Frmexporta.DirDestino.Path = NombreCarpeta
     GlArch = "PRM"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
         DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!codigo_beneficiario) & "\MEMOS\"
      Else
         DirCto = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = DirCto
     Frmexporta.Show vbModal
  Else
'    MsgBox ""
     sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
     If sino = vbYes Then
        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!codigo_beneficiario) & "\MEMOS\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        GlArch = "PRM"
        'If GlServidor <> GlMaquina Then      ' "-" Then
        If GlServidor = "SRVPRO" Then
           DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.Ado_Memo.Recordset!codigo_beneficiario) & "\MEMOS\"
        Else
           DirCto = NombreCarpeta
        End If
        Frmexporta.DirDestino2.Path = DirCto
        Frmexporta.Show vbModal
     End If
  End If

  Exit Sub
Error_Sub:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub DataCombo1_Click(Area As Integer)
Dtc_ParDes.BoundText = Dtc_Par.BoundText
End Sub

Private Sub dtc_estado_Click(Area As Integer)
  Dtc_ParDes.BoundText = dtc_estado.BoundText
  Dtc_Par.BoundText = dtc_estado.BoundText
End Sub

Private Sub Dtc_Par_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Par.BoundText
    dtc_estado.BoundText = Dtc_Par.BoundText
End Sub

Private Sub Dtc_ParDes_Click(Area As Integer)
    Dtc_Par.BoundText = Dtc_ParDes.BoundText
    dtc_estado.BoundText = Dtc_ParDes.BoundText
End Sub

Private Sub Dtc_ParDes_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()
txt01.Text = ""
DtcFec_Fin.Value = Date


DTPFec_Inicio.Value = Date
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
    rs_Clasificador.Open "SELECT * FROM rc_tipo_memoranda WHERE uso = 'A' ORDER BY descripcion ", db, adOpenStatic
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


Private Sub txt01_Click()
txt_mes.Text = txt01.ListIndex
txt_mes.Text = Val(txt_mes.Text) + 1
End Sub


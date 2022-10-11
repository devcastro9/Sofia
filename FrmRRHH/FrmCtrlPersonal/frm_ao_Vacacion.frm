VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_Vacacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Personal - File Funcionario - Vacaciones Utilizadas"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_Vacacion.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   7755
      TabIndex        =   29
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   2680
         Picture         =   "frm_ao_Vacacion.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ver Contrato PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cargar"
         Height          =   680
         Left            =   1920
         Picture         =   "frm_ao_Vacacion.frx":6C3BA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Carga Contrato en PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "frm_ao_Vacacion.frx":6C742
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1080
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_Vacacion.frx":6C94C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VACACIONES UTILIZADAS"
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
         Left            =   3630
         TabIndex        =   32
         Top             =   240
         Width           =   3975
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
      Height          =   4185
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   3720
         MaxLength       =   80
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Txt02 
         Height          =   315
         ItemData        =   "frm_ao_Vacacion.frx":6CB56
         Left            =   5640
         List            =   "frm_ao_Vacacion.frx":6CB6F
         TabIndex        =   24
         Text            =   "LUNES"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txt10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6600
         MaxLength       =   80
         TabIndex        =   23
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt09 
         Height          =   285
         Left            =   4680
         MaxLength       =   80
         TabIndex        =   22
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt08 
         Height          =   285
         Left            =   2880
         MaxLength       =   80
         TabIndex        =   21
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         Height          =   315
         ItemData        =   "frm_ao_Vacacion.frx":6CBAF
         Left            =   960
         List            =   "frm_ao_Vacacion.frx":6CBBF
         TabIndex        =   20
         Text            =   "2011"
         Top             =   3600
         Width           =   900
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   840
         MaxLength       =   80
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   2280
         MaxLength       =   80
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         MaxLength       =   80
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox txt01 
         Height          =   315
         ItemData        =   "frm_ao_Vacacion.frx":6CBDB
         Left            =   240
         List            =   "frm_ao_Vacacion.frx":6CC03
         TabIndex        =   1
         Text            =   "ENERO"
         Top             =   1320
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   42336257
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txt03 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   42336257
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txt05 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   42336258
         CurrentDate     =   0.333333333333333
         MinDate         =   4.16666666666667E-02
      End
      Begin MSComCtl2.DTPicker txt04 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   42336257
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker Txt06 
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   42336258
         CurrentDate     =   0.770833333333333
         MaxDate         =   0.999305555555556
         MinDate         =   4.16666666666667E-02
      End
      Begin MSComCtl2.DTPicker DtcFec_Fin 
         Height          =   315
         Left            =   5520
         TabIndex        =   7
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   42336257
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         Bindings        =   "frm_ao_Vacacion.frx":6CC6C
         DataField       =   "TipoPermiso"
         DataSource      =   "frmBeneficiario.AdoPermiso"
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
         ListField       =   "TipoPermiso"
         BoundColumn     =   "TipoPermiso"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         Bindings        =   "frm_ao_Vacacion.frx":6CC8B
         DataField       =   "TipoPermiso"
         DataSource      =   "frmBeneficiario.AdoPermiso"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "TipoPermiso"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker txt07 
         Height          =   315
         Left            =   5520
         TabIndex        =   14
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   42336258
         CurrentDate     =   0.896759259259259
         MaxDate         =   0.999988425925926
         MinDate         =   4.16666666666667E-02
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Días:                 Nro.Horas:                  Nro.Minutos:"
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
         Left            =   2040
         TabIndex        =   27
         Top             =   3600
         Width           =   4500
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
         Height          =   195
         Left            =   7320
         TabIndex        =   26
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Permiso:                                                  Aprobado                  Nombre de Archivo:"
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
         TabIndex        =   25
         Top             =   360
         Width           =   7305
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Benef"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "SW"
         Height          =   195
         Index           =   10
         Left            =   1920
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes de Control:                                    Fecha Solicitud                       Día Solicitud"
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
         TabIndex        =   13
         Top             =   1080
         Width           =   6555
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Fecha:                            Hasta Fecha:                              Fecha Reincorporacion: "
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
         TabIndex        =   12
         Top             =   1920
         Width           =   7290
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Hora:                               Hasta Hora:                                 Hora Reincorporación:"
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
         TabIndex        =   11
         Top             =   2760
         Width           =   7155
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
         Index           =   10
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   120
      Top             =   5040
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
Attribute VB_Name = "frm_ao_Vacacion"
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

Private Sub cmdCancel_Click()
    'cancela la edicion de datos
    Para_Aceptado = "N"
    Unload Me
    'Me.Hide
End Sub

Private Sub cmdOk_Click()
 'acepta las modificaciones realizadas
 Dim NoDias, NoHoras, NoMin As Integer
 If ValidaMontos Then
   Dim SQLS As String
   SQLS = ""
   If txtSW = "ADD" Then
      'hora01 = Format(txt05.Value, "HH:mm:ss")
      'hora02 = Format(Txt06.Value, "HH:mm:ss")
      'hora03 = Format(txt07.Value, "HH:mm:ss")
      'hora04 = Format(txt08.Value, "HH:mm:ss")
      'DB.Execute "Insert INTO ro_ControlAsistencia (codigo_beneficiario, Fecha_control, mes_control, dia_control, FechaDesde, FechaHasta, fecha_reincorporacion, horadesde, horahasta, Hora_reincorporacion, ges_gestion, dias_permiso, horas_permiso, minutos_permiso, estado_registro, fecha_registro, usr_usuario) "
      'Values ('" & txtBenef.Text & "', '" & DTPFec_Inicio.Value & "', '" & txt01 & "', '" & Txt02 & "', '" & txt03 & "', '" & txt04 & "', '" & DtcFec_Fin & "', '" & hora01 & "', '" & hora02 & "', '" & hora03 & "', '" & TxtGestion & "', '" & txt08 & "', '" & txt09 & "', '" & txt10 & "', 'N', '" & Date & "', '" & GlUsuario & "') "
      frmBeneficiario_Control.AdoPermiso.Recordset("codigo_beneficiario").Value = txtBenef.Text
      frmBeneficiario_Control.AdoPermiso.Recordset("ges_gestion").Value = TxtGestion.Text
      frmBeneficiario_Control.AdoPermiso.Recordset("mes_control") = Txt01.Text
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ro_Permisos WHERE codigo_beneficiario = '" & Trim(txtBenef.Text) & "'  ", db, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            frmBeneficiario_Control.AdoPermiso.Recordset!CORREL = rs_correlativo.RecordCount
      Else
            frmBeneficiario_Control.AdoPermiso.Recordset!CORREL = 1
      End If
      frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo"
      frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO_NOMB = Trim(frmBeneficiario_Control.Ado_datos.Recordset!iniciales) & "_Licencias_" & frmBeneficiario_Control.AdoPermiso.Recordset!CORREL & ".pdf"
      txtEstado.Text = "NO"
   End If
      frmBeneficiario_Control.AdoPermiso.Recordset("TipoPermiso").Value = Dtc_Par.Text
      frmBeneficiario_Control.AdoPermiso.Recordset("Fecha_control").Value = DTPFec_Inicio.Value
      frmBeneficiario_Control.AdoPermiso.Recordset("dia_control").Value = Txt02.Text
      frmBeneficiario_Control.AdoPermiso.Recordset("FechaDesde").Value = txt03.Value
      frmBeneficiario_Control.AdoPermiso.Recordset("FechaHasta").Value = txt04.Value
      frmBeneficiario_Control.AdoPermiso.Recordset("fecha_reincorporacion").Value = DtcFec_Fin.Value
      frmBeneficiario_Control.AdoPermiso.Recordset("horadesde").Value = Format(txt05.Value, "HH:mm:ss")
      frmBeneficiario_Control.AdoPermiso.Recordset("horahasta").Value = Format(txt06.Value, "HH:mm:ss")
      frmBeneficiario_Control.AdoPermiso.Recordset("Hora_reincorporacion").Value = Format(txt07.Value, "HH:mm:ss")
      NoDias = DateDiff("d", frmBeneficiario_Control.AdoPermiso.Recordset("FechaHasta").Value, frmBeneficiario_Control.AdoPermiso.Recordset("FechaDesde").Value)
      frmBeneficiario_Control.AdoPermiso.Recordset("dias_permiso").Value = NoDias   'txt08.Text
      NoHoras = DateDiff("h", frmBeneficiario_Control.AdoPermiso.Recordset("FechaHasta").Value, frmBeneficiario_Control.AdoPermiso.Recordset("FechaDesde").Value)
      frmBeneficiario_Control.AdoPermiso.Recordset("horas_permiso").Value = NoHoras     'txt09.Text
      NoMin = DateDiff("n", frmBeneficiario_Control.AdoPermiso.Recordset("FechaHasta").Value, frmBeneficiario_Control.AdoPermiso.Recordset("FechaDesde").Value)
      frmBeneficiario_Control.AdoPermiso.Recordset("minutos_permiso").Value = NoMin     'txt10.Text
      frmBeneficiario_Control.AdoPermiso.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "NO", txtEstado.Text)
      frmBeneficiario_Control.AdoPermiso.Recordset("fecha_registro") = Date
      frmBeneficiario_Control.AdoPermiso.Recordset("usr_usuario").Value = glusuario
      frmBeneficiario_Control.AdoPermiso.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
      frmBeneficiario_Control.AdoPermiso.Recordset.Update
   'End If
   Para_Aceptado = "S"
   Unload Me
   'Me.Hide
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
    If Txt01 = "" Then
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
 If lblARCH.Caption = "Cargar_Archivo" Then
    MsgBox ("No Existe el Archivo Asociado al Contrato, debe Cargarlo ...")
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SRVPRO" Then
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub

Private Sub CmdVerDisco_Click()
  On Error GoTo Error_Sub
  
  If frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo" Then
     NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\"
     Frmexporta.DirDestino.Path = NombreCarpeta
     GlArch = "PRM"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
         DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\"
      Else
         DirCto = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = DirCto
     Frmexporta.Show vbModal
  Else
'    MsgBox ""
     sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
     If sino = vbYes Then
        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        GlArch = "PRM"
        'If GlServidor <> GlMaquina Then      ' "-" Then
        If GlServidor = "SRVPRO" Then
           DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\"
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

Private Sub Dtc_Par_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Par.BoundText
End Sub

Private Sub Dtc_ParDes_Click(Area As Integer)
    Dtc_Par.BoundText = Dtc_ParDes.BoundText
End Sub

Private Sub Form_Load()


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
    rs_Clasificador.Open "SELECT * FROM rc_TipoPermiso WHERE estado_registro = 'NO' ORDER BY descripcion ", db, adOpenStatic
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


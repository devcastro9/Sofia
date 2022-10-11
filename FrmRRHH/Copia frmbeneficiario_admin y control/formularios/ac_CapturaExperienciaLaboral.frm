VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ac_CapturaExperienciaLaboral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Personal - Ficha Personal - Experiencia Laboral"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7620
   ControlBox      =   0   'False
   Icon            =   "ac_CapturaExperienciaLaboral.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ac_CapturaExperienciaLaboral.frx":31F48
   ScaleHeight     =   4905
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   30
      Picture         =   "ac_CapturaExperienciaLaboral.frx":9DF7A
      ScaleHeight     =   915
      ScaleWidth      =   7515
      TabIndex        =   30
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "ac_CapturaExperienciaLaboral.frx":109FAC
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "ac_CapturaExperienciaLaboral.frx":10A1B6
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXPERIENCIA LABORAL"
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
         Left            =   2910
         TabIndex        =   33
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00FFFF80&
      Height          =   3855
      Left            =   30
      TabIndex        =   11
      Top             =   980
      Width           =   7575
      Begin VB.ComboBox txt07 
         Height          =   315
         ItemData        =   "ac_CapturaExperienciaLaboral.frx":10A3C0
         Left            =   2640
         List            =   "ac_CapturaExperienciaLaboral.frx":10A3CD
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3015
         Width           =   1335
      End
      Begin VB.TextBox Txt06 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "0"
         Top             =   3015
         Width           =   735
      End
      Begin VB.TextBox Txt02 
         Height          =   285
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   3
         Top             =   1320
         Width           =   5535
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   600
         MaxLength       =   80
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         MaxLength       =   80
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt08 
         Height          =   285
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txt03 
         Height          =   525
         Left            =   1800
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1680
         Width           =   5535
      End
      Begin VB.TextBox txt04 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   5
         Top             =   2280
         Width           =   5535
      End
      Begin VB.TextBox txt05 
         Height          =   285
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2640
         Width           =   5535
      End
      Begin VB.TextBox txt01 
         Height          =   285
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   2
         Top             =   960
         Width           =   5535
      End
      Begin VB.ComboBox cboTDoc 
         Height          =   315
         ItemData        =   "ac_CapturaExperienciaLaboral.frx":10A3EA
         Left            =   6480
         List            =   "ac_CapturaExperienciaLaboral.frx":10A3F4
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3000
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtcFec_Fin 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   405
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   91029505
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   405
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   91029505
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         Bindings        =   "ac_CapturaExperienciaLaboral.frx":10A400
         DataField       =   "tipo_institucion_codigo"
         DataSource      =   "frmBeneficiario_admin.Ado_Laborales"
         Height          =   315
         Left            =   6480
         TabIndex        =   10
         Top             =   3375
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "tipo_institucion_codigo"
         BoundColumn     =   "tipo_institucion_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         Bindings        =   "ac_CapturaExperienciaLaboral.frx":10A41F
         DataField       =   "tipo_institucion_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Laborales"
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   3375
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "tipo_institucion_descripcion"
         BoundColumn     =   "tipo_institucion_codigo"
         Text            =   ""
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nombre Institución"
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
         Index           =   4
         Left            =   60
         TabIndex        =   13
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Presento Docs.de Respaldo ?"
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
         Index           =   12
         Left            =   3660
         TabIndex        =   28
         Top             =   3015
         Width           =   2715
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "SW"
         Height          =   195
         Index           =   10
         Left            =   1320
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Benef"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Duración"
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
         Left            =   885
         TabIndex        =   25
         Top             =   3015
         Width           =   810
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cargo que Ocupó"
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
         Index           =   11
         Left            =   60
         TabIndex        =   24
         Top             =   1320
         Width           =   1590
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Institución"
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
         Left            =   75
         TabIndex        =   21
         Top             =   3405
         Width           =   1620
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Finalización "
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
         Height          =   195
         Index           =   7
         Left            =   3570
         TabIndex        =   20
         Top             =   465
         Width           =   2040
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Aprobado"
         Height          =   195
         Index           =   6
         Left            =   2400
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "País"
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
         Left            =   1290
         TabIndex        =   16
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cuidad"
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
         Left            =   1050
         TabIndex        =   15
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Función Principal"
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
         Left            =   165
         TabIndex        =   14
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Inicio"
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
         Index           =   5
         Left            =   330
         TabIndex        =   12
         Top             =   435
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   0
      Top             =   4680
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
Attribute VB_Name = "ac_CapturaExperienciaLaboral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_Clasificador As New ADODB.Recordset

Dim nomb2 As String

Private Sub BtnCancelar_Click()
    'cancela la edicion de datos
    Para_Aceptado = "N"
    Unload Me
    'Me.Hide
End Sub

Private Sub BtnGrabar_Click()
 'acepta las modificaciones realizadas
 If ValidaMontos Then
   Dim SQLS As String
   SQLS = ""
   If txtSW = "ADD" Then
      db.Execute "Insert INTO ro_experiencia_laboral (beneficiario_codigo, denominacion_institucion, cargo, funcion_general, tipo_institucion, Tiempo_Meses, tiempo_dmy, pais, ciudad, fecha_inicio, fecha_fin, presento_documento, estado_codigo, fecha_registro, usr_usuario) Values ('" & txtBenef.Text & "', '" & Txt01.Text & "', '" & Txt02.Text & "', '" & txt03.Text & "', '" & Dtc_Par.Text & "', " & txt06.Text & ", '" & txt07.Text & "', '" & txt04.Text & "', '" & txt05.Text & "', '" & DTPFec_Inicio.Value & "', '" & DtcFec_Fin.Value & "', '" & cboTDoc.Text & "', '" & txtEstado.Text & "', '" & Date & "', '" & glusuario & "')"
      'db.Execute "Insert INTO rc_experiencia_laboral (beneficiario_codigo, denominacion_institucion, cargo, funcion_general, tipo_institucion, Tiempo_Meses, tiempo_dmy, pais, ciudad, fecha_inicio, fecha_fin, presento_documento, estado_codigo, fecha_registro, usr_usuario) Values ('" & txtBenef.Text & "', '" & Txt01.Text & "', '" & txt02.Text & "', '" & txt03.Text & "', '" & Dtc_Par.Text & "', " & txt06.Text & ", '" & txt07.Text & "', '" & txt04.Text & "', '" & txt05.Text & "', '" & DTPFec_Inicio.Value & "', '" & DtcFec_Fin.Value & "', '" & cboTDoc.Text & "', '" & Txtestado.Text & "', '" & Date & "', '" & glusuario & "')"
      rw_ficha_rrhh.abrirtabla
   Else
      rw_ficha_rrhh.Ado_Laborales.Recordset("beneficiario_codigo").Value = txtBenef.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("denominacion_institucion") = Txt01.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("cargo").Value = Txt02.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("funcion_general").Value = txt03.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("tipo_institucion").Value = Dtc_Par.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("Tiempo_Meses").Value = txt06.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("tiempo_dmy").Value = txt07.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("pais").Value = txt04.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("ciudad").Value = txt05.Text
      rw_ficha_rrhh.Ado_Laborales.Recordset("fecha_inicio").Value = DTPFec_Inicio.Value
      rw_ficha_rrhh.Ado_Laborales.Recordset("fecha_fin").Value = DtcFec_Fin.Value
      rw_ficha_rrhh.Ado_Laborales.Recordset("presento_documento").Value = cboTDoc.Text '.List(1)
      rw_ficha_rrhh.Ado_Laborales.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "REG", txtEstado.Text)
      rw_ficha_rrhh.Ado_Laborales.Recordset("fecha_registro") = Date
      rw_ficha_rrhh.Ado_Laborales.Recordset("usr_usuario").Value = glusuario
      rw_ficha_rrhh.Ado_Laborales.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
      rw_ficha_rrhh.Ado_Laborales.Recordset.Update
       rw_ficha_rrhh.abrirtabla
   End If
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


Private Sub Dtc_Par_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Par.BoundText
End Sub

Private Sub Dtc_ParDes_Click(Area As Integer)
    Dtc_Par.BoundText = Dtc_ParDes.BoundText
End Sub

Private Sub Form_Load()
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
    rs_Clasificador.Open "SELECT * FROM rc_tipo_institucion ORDER BY tipo_institucion_descripcion ", db, adOpenStatic
    Set Ado_Clasificador.Recordset = rs_Clasificador
    
'If Val(Me.mskMonto_limite) = 0 Then
'    Me.labPorcExt = "0%"
'    Me.labPorcNal = "0%"
'End If
'mskMonto.SetFocus
End Sub

'Private Sub mskMonto_Change()
'    Call DivideXFte
'End Sub

'Sub DivideXFte()
''divide el monto total en montos correspondientes alos porcentajes
''externo y contraparte nacional
'Me.mskMonto_ext = Round(Val(Me.mskMonto) * Val(Left(Me.labPorcExt, Len(Me.labPorcExt) - 1)) / 100, 2)
'Me.mskMonto_nal = Round(Val(Me.mskMonto) - Val(Me.mskMonto_ext), 2)
'End Sub

'Private Sub mskMonto_ext_GotFocus()
'Me.mskMonto.SetFocus
'End Sub

Private Sub mskMonto_GotFocus()
mskMonto.SelStart = 0
mskMonto.SelLength = Len(mskMonto)
End Sub

Private Sub mskMonto_KeyPress(KeyAscii As Integer)
If Val(Chr(KeyAscii)) <> 0 Or Chr(KeyAscii) = "." Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "0" Or KeyAscii = 8 Then
    'asdfasdf
Else
    KeyAscii = 0
End If
End Sub

'Private Sub mskMonto_limite_GotFocus()
'Me.mskMonto.SetFocus
'End Sub

'Private Sub mskMonto_nal_GotFocus()
'Me.mskMonto.SetFocus
'End Sub

'Private Sub mskMonto_pendiente_GotFocus()
'Me.mskMonto.SetFocus
'End Sub


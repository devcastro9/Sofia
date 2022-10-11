VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form aw_p_ao_solicitud_item_costos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización Venta - Crear Nuevos Items para la Hoja de Costos"
   ClientHeight    =   5970
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   Icon            =   "aw_p_ao_solicitud_item_costos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "aw_p_ao_solicitud_item_costos.frx":0A02
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   8
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00404040&
         Height          =   675
         Left            =   720
         Picture         =   "aw_p_ao_solicitud_item_costos.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00C0C0C0&
         Height          =   675
         Left            =   2160
         MaskColor       =   &H00000000&
         Picture         =   "aw_p_ao_solicitud_item_costos.frx":6D20A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUEVOS ITEMS PARA HOJA DE COSTOS"
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
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   6195
      End
   End
   Begin VB.Frame Fra_datos99 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   10695
      Begin VB.TextBox txt_porc 
         DataField       =   "costo_porcentaje"
         DataSource      =   "Ado_datos10"
         Height          =   285
         Left            =   3000
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txt_monto 
         DataField       =   "costo_monto"
         DataSource      =   "Ado_datos10"
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "EUROPA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   7560
         TabIndex        =   17
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ASIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   16
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AMERICA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   15
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txt_descripcion 
         DataField       =   "costo_descripcion"
         DataSource      =   "Ado_datos10"
         Height          =   645
         Left            =   3000
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1200
         Width           =   7335
      End
      Begin VB.Label lbl_porc 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje (Valor Numérico)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lbl_monto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Costo Fijo a Aplicar en Bs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   2160
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Se aplicará en ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   1440
         TabIndex        =   18
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "codigo_costo"
         DataSource      =   "Ado_datos10"
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
         Height          =   300
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl_des 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Denominación del Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   8
         Left            =   840
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   1080
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11025
      TabIndex        =   0
      Top             =   5970
      Width           =   11025
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   120
      Top             =   5520
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "Ado_datos10"
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
Attribute VB_Name = "aw_p_ao_solicitud_item_costos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos10 As New ADODB.Recordset
Attribute rs_datos10.VB_VarHelpID = -1
'BUSCADOR

'OTROS
Dim var_cod As String
Dim VAR_VAL As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error GoTo AddErr
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        aw_p_ao_solicitud_item_costos.Ado_datos10.Recordset.CancelUpdate
        Unload Me
    End If
     Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
         rs_datos10!costo_descripcion = Txt_descripcion.Text
         rs_datos10!costo_monto = txt_monto
         rs_datos10!costo_porcentaje = txt_porc.Text
         If Check1.Value = 1 Then
            rs_datos10!costo_tipo = "B"
         Else
            rs_datos10!costo_tipo = "X"
         End If
         If Check2.Value = 1 Then
            rs_datos10!costo_tipoA = "B"
         Else
            rs_datos10!costo_tipoA = "X"
         End If
         If Check3.Value = 1 Then
            rs_datos10!costo_tipoE = "B"
         Else
            rs_datos10!costo_tipoE = "X"
         End If
         rs_datos10!fecha_registro = Date     'no cambia
         rs_datos10!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
         rs_datos10.Update    'Batch 'adAffectAll
  End If
  Unload Me
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
    If (Txt_descripcion = "") Then
    MsgBox "Debe registrar la Denominación del Item ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_monto.Text = "") Then
    MsgBox "Debe registrar el Costo Fijo del Item ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_porc.Text = "") Then
    MsgBox "Debe registrar el Porcentaje del Item ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    rs_datos10.AddNew
    mbDataChanged = False
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from ac_costos_comercializacion ", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos10.Recordset = rs_datos10
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


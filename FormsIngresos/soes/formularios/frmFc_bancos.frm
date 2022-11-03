VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFc_bancos 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   Icon            =   "frmFc_bancos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Banco "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   39
      Top             =   30
      Width           =   7785
      Begin VB.TextBox txtGes_gestion 
         DataField       =   "Ges_gestion"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   5490
         TabIndex        =   41
         Top             =   225
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txtBco_codigo 
         DataField       =   "Bco_codigo"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   855
         MaxLength       =   3
         TabIndex        =   40
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ges_gestion:"
         Height          =   255
         Index           =   0
         Left            =   3645
         TabIndex        =   43
         Top             =   270
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   42
         Top             =   210
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   945
      Left            =   75
      TabIndex        =   35
      Top             =   3945
      Width           =   7830
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   6720
         MousePointer    =   4  'Icon
         Picture         =   "frmFc_bancos.frx":324A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   165
         Width           =   1005
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   5715
         MousePointer    =   4  'Icon
         Picture         =   "frmFc_bancos.frx":3554
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   165
         Width           =   1005
      End
      Begin VB.CommandButton cdmAnular 
         Caption         =   "Anular"
         Enabled         =   0   'False
         Height          =   720
         Left            =   4710
         Picture         =   "frmFc_bancos.frx":385E
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Anula el comprobante de Ingreso"
         Top             =   165
         Width           =   1005
      End
      Begin MSAdodcLib.Adodc adoFc_bancos 
         Height          =   330
         Left            =   195
         Top             =   195
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   "Nuevo Registro"
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   7800
      Begin VB.TextBox txtBco_descripcion_larga 
         DataField       =   "Bco_descripcion_larga"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   17
         Top             =   150
         Width           =   3375
      End
      Begin VB.TextBox txtBco_sigla 
         DataField       =   "Bco_sigla"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   16
         Top             =   435
         Width           =   1650
      End
      Begin VB.TextBox txtBco_ciudad 
         DataField       =   "Bco_ciudad"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   15
         Top             =   1005
         Width           =   3300
      End
      Begin VB.TextBox txtBco_Estado_Depto 
         DataField       =   "Bco_Estado_Depto"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   14
         Top             =   1290
         Width           =   3300
      End
      Begin VB.TextBox txtBco_direccion 
         DataField       =   "Bco_direccion"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   13
         Top             =   1575
         Width           =   3375
      End
      Begin VB.TextBox txtBco_Codigo_postal 
         DataField       =   "Bco_Codigo_postal"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   12
         Top             =   1860
         Width           =   3300
      End
      Begin VB.TextBox txtBco_Reserva_Federal 
         DataField       =   "Bco_Reserva_Federal"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   7050
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox txtBco_intermediario 
         DataField       =   "Bco_intermediario"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   10
         Top             =   2145
         Width           =   3375
      End
      Begin VB.TextBox txtBco_Observaciones 
         DataField       =   "Bco_Observaciones"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   9
         Top             =   2430
         Width           =   3375
      End
      Begin VB.TextBox txtBco_activo 
         DataField       =   "Bco_activo"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   5490
         TabIndex        =   8
         Top             =   660
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.TextBox txtRepresentante 
         DataField       =   "Representante"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   7
         Top             =   2715
         Width           =   3375
      End
      Begin VB.TextBox txtCargo 
         DataField       =   "Cargo"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   6
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtfecha_registro 
         DataField       =   "fecha_registro"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   4950
         TabIndex        =   5
         Top             =   1155
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.TextBox txthora_registro 
         DataField       =   "hora_registro"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   4950
         TabIndex        =   4
         Top             =   1530
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtusr_usuario 
         DataField       =   "usr_usuario"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   4950
         TabIndex        =   3
         Top             =   1920
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.TextBox txtbco_creado 
         DataField       =   "bco_creado"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   5610
         TabIndex        =   2
         Top             =   2310
         Visible         =   0   'False
         Width           =   165
      End
      Begin MSDataListLib.DataCombo dc_paises 
         Bindings        =   "frmFc_bancos.frx":3F48
         Height          =   315
         Left            =   1845
         TabIndex        =   1
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_pais"
         BoundColumn     =   "codigo_pais"
         Text            =   "DataCombo1"
      End
      Begin MSAdodcLib.Adodc ado_paises 
         Height          =   330
         Left            =   5775
         Top             =   765
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
         Caption         =   "Ado_paises"
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         Height          =   195
         Index           =   2
         Left            =   885
         TabIndex        =   34
         Top             =   195
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
         Height          =   195
         Index           =   3
         Left            =   1380
         TabIndex        =   33
         Top             =   480
         Width           =   390
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   4
         Left            =   1425
         TabIndex        =   32
         Top             =   765
         Width           =   345
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   5
         Left            =   1230
         TabIndex        =   31
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado/Depto:"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   30
         Top             =   1335
         Width           =   1050
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Index           =   7
         Left            =   1050
         TabIndex        =   29
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal:"
         Height          =   195
         Index           =   8
         Left            =   750
         TabIndex        =   28
         Top             =   1905
         Width           =   1020
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bco_Reserva_Federal:"
         Height          =   255
         Index           =   9
         Left            =   5205
         TabIndex        =   27
         Top             =   465
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bco_intermediario:"
         Height          =   255
         Index           =   10
         Left            =   270
         TabIndex        =   26
         Top             =   2190
         Width           =   1500
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Index           =   11
         Left            =   660
         TabIndex        =   25
         Top             =   2475
         Width           =   1110
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bco_activo:"
         Height          =   255
         Index           =   12
         Left            =   3645
         TabIndex        =   24
         Top             =   705
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Representante:"
         Height          =   255
         Index           =   13
         Left            =   465
         TabIndex        =   23
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cargo:"
         Height          =   255
         Index           =   14
         Left            =   375
         TabIndex        =   22
         Top             =   3045
         Width           =   1395
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "fecha_registro:"
         Height          =   255
         Index           =   15
         Left            =   3105
         TabIndex        =   21
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "hora_registro:"
         Height          =   255
         Index           =   16
         Left            =   3105
         TabIndex        =   20
         Top             =   1575
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "usr_usuario:"
         Height          =   255
         Index           =   17
         Left            =   3105
         TabIndex        =   19
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "bco_creado:"
         Height          =   255
         Index           =   18
         Left            =   3105
         TabIndex        =   18
         Top             =   2340
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmFc_bancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim accion As String
Public bco_codigo_ret As String

Public Sub fc_bancos_procesar(proceso, bco_codigo As String)
  accion = proceso
  bco_codigo_ret = ""
  If proceso = "SELECT_UNO" Then
    Caption = "Banco..."
    fc_bancos_select proceso, bco_codigo
    llena_fc_bancos bco_codigo
  ElseIf proceso = "INSERT" Then
    Caption = "Inserción de un Nuevo Banco"
    llena_fc_bancos ""
  ElseIf proceso = "SELECT" Then
    Caption = "Lista de Bancos"
    fc_bancos_select proceso, ""
  End If
  Show vbModal
End Sub

Private Sub llena_fc_bancos(bco_codigo As String)
If bco_codigo = "" Then
  Me.txtBco_codigo = ""
  Me.txtGes_gestion = ""
  Me.txtBco_descripcion_larga = ""
  Me.txtBco_sigla = ""
  dc_paises.BoundText = ""
  Me.txtBco_ciudad = ""
  Me.txtBco_Estado_Depto = ""
  Me.txtBco_direccion = ""
  Me.txtBco_Codigo_postal = ""
  Me.txtBco_Reserva_Federal = ""
  Me.txtBco_intermediario = ""
  Me.txtBco_Observaciones = ""
  Me.txtBco_activo = ""
  Me.txtRepresentante = ""
  Me.txtCargo = ""
  Me.txtfecha_registro = ""
  Me.txthora_registro = ""
  Me.txtusr_usuario = ""
  Me.txtbco_creado = ""
Else
  Me.txtGes_gestion = adoFc_bancos.Recordset!ges_gestion
  Me.txtBco_codigo = adoFc_bancos.Recordset!bco_codigo
  Me.txtBco_descripcion_larga = adoFc_bancos.Recordset!bco_descripcion_larga
  Me.txtBco_sigla = IIf(IsNull(adoFc_bancos.Recordset!bco_sigla), "", adoFc_bancos.Recordset!bco_sigla)
  dc_paises.BoundText = adoFc_bancos.Recordset!codigo_pais
  Me.txtBco_ciudad = IIf(IsNull(adoFc_bancos.Recordset!bco_ciudad), "", adoFc_bancos.Recordset!bco_ciudad)
  Me.txtBco_Estado_Depto = IIf(IsNull(adoFc_bancos.Recordset!bco_estado_depto), "", adoFc_bancos.Recordset!bco_estado_depto)
  Me.txtBco_direccion = IIf(IsNull(adoFc_bancos.Recordset!bco_direccion), "", adoFc_bancos.Recordset!bco_direccion)
  Me.txtBco_Codigo_postal = IIf(IsNull(adoFc_bancos.Recordset!bco_codigo_postal), "", adoFc_bancos.Recordset!bco_codigo_postal)
  Me.txtBco_Reserva_Federal = IIf(IsNull(adoFc_bancos.Recordset!bco_reserva_federal), "", adoFc_bancos.Recordset!bco_reserva_federal)
  Me.txtBco_intermediario = IIf(IsNull(adoFc_bancos.Recordset!Bco_intermediario), "", adoFc_bancos.Recordset!Bco_intermediario)
  Me.txtBco_Observaciones = IIf(IsNull(adoFc_bancos.Recordset!Bco_Observaciones), "", adoFc_bancos.Recordset!Bco_Observaciones)
  Me.txtBco_activo = IIf(IsNull(adoFc_bancos.Recordset!Bco_activo), "", adoFc_bancos.Recordset!Bco_activo)
  Me.txtRepresentante = IIf(IsNull(adoFc_bancos.Recordset!Representante), "", adoFc_bancos.Recordset!Representante)
  Me.txtCargo = IIf(IsNull(adoFc_bancos.Recordset!Cargo), "", adoFc_bancos.Recordset!Cargo)
  Me.txtfecha_registro = CStr(adoFc_bancos.Recordset!fecha_registro)
  Me.txthora_registro = IIf(IsNull(adoFc_bancos.Recordset!hora_registro), "", adoFc_bancos.Recordset!hora_registro)
  Me.txtusr_usuario = IIf(IsNull(adoFc_bancos.Recordset!usr_usuario), "", adoFc_bancos.Recordset!usr_usuario)
  Me.txtbco_creado = IIf(IsNull(adoFc_bancos.Recordset!bco_creado), "", adoFc_bancos.Recordset!bco_creado)
End If
End Sub

Private Sub fc_bancos_select(tipo, bco_codigo As String)
Dim fecha As Date
  Datos.dbo_so_fc_bancos tipo, bco_codigo, "", "", "", "", "", "", "", "", "", "", "", "", "", "", fecha, "", "", ""
  With Datos.rsdbo_so_fc_bancos
    Set adoFc_bancos.Recordset = .Clone
    .Close
  End With
End Sub

Private Sub adoFc_bancos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If Not (adoFc_bancos.Recordset.EOF Or adoFc_bancos.Recordset.BOF) Then
    llena_fc_bancos adoFc_bancos.Recordset!bco_codigo
    adoFc_bancos.Caption = CStr(adoFc_bancos.Recordset.Bookmark) & " de " & CStr(adoFc_bancos.Recordset.RecordCount)
  Else
    adoFc_bancos.Caption = "0 de 0"
  End If
End Sub

Private Sub cdmAnular_Click()
Dim fecha As Date
  Datos.dbo_so_fc_bancos "DELETE", Me.txtBco_codigo, "", "", "", "", "", "", "", "", "", "", "", "", "", "", fecha, "", "", ""
  fc_bancos_select "SELECT", ""
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
  bco_codigo_ret = ""
End Sub

Private Sub CmdGrabar_Click()
  If accion = "INSERT" Then
    If valida_registro Then
      Datos.dbo_so_fc_bancos "INSERT", Me.txtBco_codigo, Me.txtGes_gestion, Me.txtBco_descripcion_larga, Me.txtBco_sigla, dc_paises.BoundText, Me.txtBco_ciudad, Me.txtBco_Estado_Depto, Me.txtBco_direccion, Me.txtBco_Codigo_postal, Me.txtBco_Reserva_Federal, Me.txtBco_intermediario, Me.txtBco_Observaciones, Me.txtBco_activo, Me.txtRepresentante, Me.txtCargo, Date, Me.txtfecha_registro, Me.txtusr_usuario, Me.txtbco_creado
      bco_codigo_ret = Me.txtBco_codigo.Text
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Load()
  Set tpaises = New ADODB.Recordset
  If tpaises.State = 1 Then tpaises.Close
    tpaises.Open "SELECT codigo_pais, denominacion_pais FROM paises ", db, adOpenDynamic, adLockReadOnly
  Set ado_paises.Recordset = tpaises
	Call SeguridadSet(Me)
End Sub

Function valida_registro() As Boolean
Dim ok As Boolean
  ok = True
  If ok And Me.txtBco_codigo = "" Then
    MsgBox "Ingrese Codigo del Banco"
    ok = False
  End If
  If ok And Me.txtBco_descripcion_larga = "" Then
    MsgBox "Ingrese Descripción o nombre del Banco"
    ok = False
  End If
  If ok And Me.txtBco_sigla = "" Then
    MsgBox "Ingrese Sigla del Banco"
    ok = False
  End If
  If ok And Me.dc_paises.Text = "" Then
    MsgBox "Ingrese Pais de origen"
    ok = False
  End If
  If ok And Me.txtBco_ciudad = "" Then
    MsgBox "Ingrese Ciudad"
    ok = False
  End If
  If ok And Me.txtBco_direccion = "" Then
    MsgBox "Ingrese Direccion"
    ok = False
  End If
'  If ok And = "" Then
'    MsgBox "Ingrese "
'    ok = False
'  End If
  valida_registro = ok
End Function

Private Sub txtBco_codigo_Validate(Cancel As Boolean)
  If txtBco_codigo.Text = "" Then
    Cancel = False
  Else
    If mod_librerias.GetValor("fc_bancos", "bco_codigo", "bco_codigo", Me.txtBco_codigo) = Me.txtBco_codigo Then
      Cancel = True
      MsgBox Me.txtBco_codigo + " ya esta registrado"
    End If
  End If
End Sub

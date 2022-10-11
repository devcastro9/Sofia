VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rw_personal_cuenta_banco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha Personal - Cuentas Bancarias"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9900
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SEGUNDA CUENTA BANCARIA PERSONAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2325
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   9660
      Begin VB.TextBox DtcCtaNom2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "beneficiario_denominacion"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   23
         Top             =   1800
         Width           =   6255
      End
      Begin VB.TextBox DtcCta2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "cta_codigo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1320
         Width           =   3405
      End
      Begin VB.ComboBox DtcCtaTip2 
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "rw_personal_cuenta_banco.frx":0000
         Left            =   3000
         List            =   "rw_personal_cuenta_banco.frx":000A
         TabIndex        =   15
         Text            =   "CUENTA CORRIENTE"
         Top             =   840
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DtcBanco2 
         Bindings        =   "rw_personal_cuenta_banco.frx":0030
         DataField       =   "bco_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   8400
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483628
         ListField       =   "bco_codigo"
         BoundColumn     =   "bco_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcBancoDes2 
         Bindings        =   "rw_personal_cuenta_banco.frx":0045
         DataField       =   "bco_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   18
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "bco_descripcion"
         BoundColumn     =   "bco_codigo"
         Text            =   ""
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad Financiera"
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
         Left            =   360
         TabIndex        =   22
         Top             =   345
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta"
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
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Bancaria"
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
         Left            =   360
         TabIndex        =   20
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominacion de la Cuenta"
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
         Left            =   360
         TabIndex        =   19
         Top             =   1800
         Width           =   2475
      End
   End
   Begin VB.Frame FraBco 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PRIMERA CUENTA BANCARIA PERSONAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2325
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   9660
      Begin VB.ComboBox DtcCtaTip 
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "rw_personal_cuenta_banco.frx":005A
         Left            =   3000
         List            =   "rw_personal_cuenta_banco.frx":0064
         TabIndex        =   5
         Text            =   "CUENTA CORRIENTE"
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox DtcCtaNom 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "beneficiario_denominacion"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1800
         Width           =   6255
      End
      Begin VB.TextBox DtcCta 
         BackColor       =   &H00FFFFFF&
         DataField       =   "cta_codigo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1320
         Width           =   3405
      End
      Begin MSDataListLib.DataCombo DtcBanco 
         Bindings        =   "rw_personal_cuenta_banco.frx":008A
         DataField       =   "bco_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   8400
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483628
         ListField       =   "bco_codigo"
         BoundColumn     =   "bco_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcBancoDes 
         Bindings        =   "rw_personal_cuenta_banco.frx":009F
         DataField       =   "bco_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "bco_descripcion"
         BoundColumn     =   "bco_codigo"
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominacion de la Cuenta"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   2475
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Bancaria"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta"
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
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad Financiera"
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
         Left            =   360
         TabIndex        =   8
         Top             =   345
         Width           =   1680
      End
   End
   Begin VB.PictureBox Frame2 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1680
         Picture         =   "rw_personal_cuenta_banco.frx":00B4
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "rw_personal_cuenta_banco.frx":09A0
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   9
         Top             =   120
         Width           =   1280
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTAS BANCARIAS PERSONALES"
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
         Left            =   3750
         TabIndex        =   1
         Top             =   240
         Width           =   5685
      End
   End
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
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
   Begin MSAdodcLib.Adodc AdoPermisoDetalle 
      Height          =   330
      Left            =   120
      Top             =   6000
      Visible         =   0   'False
      Width           =   9645
      _ExtentX        =   17013
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
      BackColor       =   12632319
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
      Caption         =   " <--- Detalle de Permisos --->"
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
   Begin MSAdodcLib.Adodc Ado_Clasificador2 
      Height          =   330
      Left            =   2640
      Top             =   6360
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
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
      Caption         =   "Ado_Clasificador2"
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
   Begin MSAdodcLib.Adodc Ado_Clasificador3 
      Height          =   330
      Left            =   5160
      Top             =   6360
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
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
      Caption         =   "Ado_Clasificador3"
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
   Begin MSAdodcLib.Adodc Ado_Clasificador4 
      Height          =   330
      Left            =   7680
      Top             =   6360
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
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
      Caption         =   "Ado_Clasificador4"
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
Attribute VB_Name = "rw_personal_cuenta_banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Para_Aceptado As String
Dim rs_Clasificador As New ADODB.Recordset
Dim rs_correlativo As New ADODB.Recordset
Dim rs_correl_vac As New ADODB.Recordset
Dim rs_Permiso_detalle As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim sqlAux As String
Dim nomb2 As String
Dim hora01, hora02, hora03, hora04 As String
Dim fecha1 As String
Dim DirLic, DirVac As String
Dim totHrs, totMin, totVac As Integer
Dim numminutosTT As Integer


Private Sub BtnCancelar_Click()
    Unload Me
End Sub

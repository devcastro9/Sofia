VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form gw_beneficiario_cuenta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Bancarias de Personas"
   ClientHeight    =   7455
   ClientLeft      =   420
   ClientTop       =   1830
   ClientWidth     =   9270
   Icon            =   "gw_beneficiario_cuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9270
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox fraDatos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7275
      ScaleWidth      =   9045
      TabIndex        =   2
      Top             =   120
      Width           =   9105
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tercera Cuenta Bancaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1935
         Left            =   180
         TabIndex        =   23
         Top             =   4200
         Width           =   8670
         Begin VB.TextBox txt_campo3 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cta_codigo3"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   24
            Top             =   1440
            Width           =   3645
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "gw_beneficiario_cuenta.frx":0A02
            DataField       =   "bco_codigo3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7800
            TabIndex        =   25
            Top             =   480
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "bco_codigo"
            BoundColumn     =   "bco_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "gw_beneficiario_cuenta.frx":0A1B
            DataField       =   "bco_codigo3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   26
            Top             =   480
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
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "gw_beneficiario_cuenta.frx":0A34
            DataField       =   "cta_tipo3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6120
            TabIndex        =   27
            Top             =   960
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "cta_tipo"
            BoundColumn     =   "cta_tipo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "gw_beneficiario_cuenta.frx":0A4D
            DataField       =   "cta_tipo3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   28
            Top             =   960
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "cta_tipo_descripcion"
            BoundColumn     =   "cta_tipo"
            Text            =   "Todos"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Entidad Financiera 3"
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
            Left            =   480
            TabIndex        =   31
            Top             =   480
            Width           =   1830
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cuenta 3"
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
            Left            =   480
            TabIndex        =   30
            Top             =   960
            Width           =   1530
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria 3"
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
            Left            =   480
            TabIndex        =   29
            Top             =   1440
            Width           =   1635
         End
      End
      Begin VB.PictureBox FraGrabarCancelar 
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   30
         ScaleHeight     =   915
         ScaleWidth      =   8940
         TabIndex        =   5
         Top             =   6220
         Width           =   9000
         Begin VB.CommandButton BtnGrabar 
            BackColor       =   &H80000015&
            Height          =   675
            Left            =   2640
            Picture         =   "gw_beneficiario_cuenta.frx":0A66
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   120
            Width           =   1365
         End
         Begin VB.CommandButton BtnCancelar 
            BackColor       =   &H80000015&
            Height          =   675
            Left            =   4200
            MaskColor       =   &H00000000&
            Picture         =   "gw_beneficiario_cuenta.frx":123C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Cancelar"
            Top             =   120
            Width           =   1365
         End
         Begin VB.Label lbl_titulo2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   435
            Left            =   10425
            TabIndex        =   6
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta Bancaria Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1905
         Left            =   180
         TabIndex        =   3
         Top             =   105
         Width           =   8670
         Begin VB.TextBox txt_campo1 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   14
            Top             =   1440
            Width           =   3645
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "gw_beneficiario_cuenta.frx":1B28
            DataField       =   "bco_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7800
            TabIndex        =   8
            Top             =   480
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "bco_codigo"
            BoundColumn     =   "bco_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "gw_beneficiario_cuenta.frx":1B41
            DataField       =   "bco_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   9
            Top             =   480
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
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "gw_beneficiario_cuenta.frx":1B5A
            DataField       =   "cta_tipo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6120
            TabIndex        =   12
            Top             =   960
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "cta_tipo"
            BoundColumn     =   "cta_tipo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "gw_beneficiario_cuenta.frx":1B73
            DataField       =   "cta_tipo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   13
            Top             =   960
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "cta_tipo_descripcion"
            BoundColumn     =   "cta_tipo"
            Text            =   "Todos"
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria 1"
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
            Left            =   480
            TabIndex        =   11
            Top             =   1440
            Width           =   1635
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cuenta 1"
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
            Left            =   480
            TabIndex        =   10
            Top             =   960
            Width           =   1530
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Entidad Financiera 1"
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
            Left            =   480
            TabIndex        =   7
            Top             =   480
            Width           =   1830
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Segunda Cuenta Bancaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1935
         Left            =   180
         TabIndex        =   4
         Top             =   2160
         Width           =   8670
         Begin VB.TextBox txt_campo2 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   22
            Top             =   1440
            Width           =   3645
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "gw_beneficiario_cuenta.frx":1B8C
            DataField       =   "bco_codigo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7800
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "bco_codigo"
            BoundColumn     =   "bco_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "gw_beneficiario_cuenta.frx":1BA5
            DataField       =   "bco_codigo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   19
            Top             =   480
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
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "gw_beneficiario_cuenta.frx":1BBE
            DataField       =   "cta_tipo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6120
            TabIndex        =   20
            Top             =   960
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "cta_tipo"
            BoundColumn     =   "cta_tipo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "gw_beneficiario_cuenta.frx":1BD7
            DataField       =   "cta_tipo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   21
            Top             =   960
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "cta_tipo_descripcion"
            BoundColumn     =   "cta_tipo"
            Text            =   "Todos"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria 2"
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
            Left            =   480
            TabIndex        =   17
            Top             =   1440
            Width           =   1635
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cuenta 2"
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
            Left            =   480
            TabIndex        =   16
            Top             =   960
            Width           =   1530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Entidad Financiera 2"
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
            Left            =   480
            TabIndex        =   15
            Top             =   480
            Width           =   1830
         End
      End
   End
   Begin Crystal.CrystalReport CR01 
      Left            =   11040
      Top             =   7560
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
   End
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   6480
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_datos6"
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   7320
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   8640
      Top             =   7680
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
      Caption         =   "Ado_datos7"
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
      Left            =   2160
      Top             =   7320
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
      Caption         =   "Ado_datos2"
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   0
      Top             =   7680
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
      Caption         =   "Ado_datos8"
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
      Left            =   4320
      Top             =   7320
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
      Caption         =   "Ado_datos3"
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   2160
      Top             =   7680
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
      Caption         =   "Ado_datos9"
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
      Left            =   6480
      Top             =   7320
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
      Caption         =   "Ado_datos4"
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   4320
      Top             =   7680
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   8640
      Top             =   7320
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
      Caption         =   "Ado_datos5"
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   0
      Top             =   8040
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
      Caption         =   "Ado_datos11"
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
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   4680
      Top             =   8040
      Width           =   3465
      _ExtentX        =   6112
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
      BackColor       =   16777152
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
      Caption         =   ""
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
Attribute VB_Name = "gw_beneficiario_cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento de Beneficiarios
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
'Dim rs_datos7 As New ADODB.Recordset
'Dim rs_datos8 As New ADODB.Recordset
'Dim rs_datos9 As New ADODB.Recordset
'Dim rs_datos10 As New ADODB.Recordset
'Dim rs_datos11 As New ADODB.Recordset

'OTROS
Dim VAR_PAIS As String
Dim VAR_VAL As String
Dim VAR_SW, VAR_AUX As String
Dim NombreCarpeta, e As String
Dim SQL_FOR As String
Dim RUTA1 As String
Dim VAR_PWD As String
Dim CodBenef As String
Dim sino As String
Dim queryinicial As String

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  If Ado_datos.Recordset.EOF Or Ado_datos.Recordset.BOF Then
'      BtnModificar.Enabled = False
'     ' BtnEliminar.Enabled = False
'      'TxtTipo.Text = Empty
'      txtCodigo.Text = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
'      txtDenominacion.Text = Empty
'      Exit Sub
'  End If
  If Ado_datos.Recordset.RecordCount > 0 Then
'    Select Case Ado_datos.Recordset.EditMode
'      Case adEditInProgress
'        Frame2.Enabled = False            'Verif. Nombre Proveedor JQA NOV-2009
'
'      Case adEditNone
'      Case adEditDelete
'      Case adEditAdd
'        Frame2.Enabled = True            'Verif. Nombre Proveedor JQA NOV-2009
'    End Select

    'If VAR_SW = "ADD" Then
      txt_campo1.Visible = True
      txt_campo2.Visible = True
      txt_campo3.Visible = True
    'Else
    '  txt_campo1.Visible = False
    '  txt_campo2.Visible = False
    '  txt_campo3.Visible = False
    'End If
    'Ado_datos.Caption = Ado_datos.Recordset!beneficiario_codigo + " - " + CStr(Ado_datos.Recordset!calle_codigo)
    'Ado_datos.Caption = CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & CStr(Ado_datos.Recordset.RecordCount)
    '  <-- Inicio                   Viviendas - Edificaciones                   Fin -->
  End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     db.Execute "update ro_personal_contratado set bco_codigo = '" & dtc_codigo1.Text & "', cta_tipo = '" & dtc_codigo2.Text & "', cta_codigo = '" & txt_campo1.Text & "', bco_codigo2 = '" & dtc_codigo3.Text & "', cta_tipo2 = '" & dtc_codigo4.Text & "', cta_codigo2 = '" & txt_campo2.Text & "', bco_codigo3 = '" & dtc_codigo5.Text & "', cta_tipo3 = '" & dtc_codigo6.Text & "', cta_codigo3 = '" & txt_campo3.Text & "' where beneficiario_codigo = '" & glBenef & "'"
  End If
     FraGrabarCancelar.Visible = False
     txt_campo1.Visible = False
     txt_campo2.Visible = False
     txt_campo3.Visible = False

  Unload Me
  Exit Sub
UpdateErr:
  MsgBox Err.Description
    
End Sub

Private Sub valida_campos()
    'Entidad Financiera 1
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  'Tipo de Cuenta 1
  If dtc_codigo2.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  'Cuenta Bancaria 1
  If txt_campo1.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
'  If dtc_codigo1.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If dtc_codigo2.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If dtc_codigo3.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
        Call ABRIR_TABLAS_AUX
        'Call ABRIR_TABLA
        rs_datos.MoveFirst
        FraGrabarCancelar.Visible = False
        txt_campo1.Visible = False
        txt_campo2.Visible = False
        txt_campo3.Visible = False
    End If
    
      Unload Me
End Sub

Private Sub pnivel5(codigo7 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo7 & "' order by zona_denominacion"
'   Set dtc_codigo8.RowSource = Nothing
'   Set dtc_codigo8.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_codigo8.ReFill
'   dtc_codigo8.BoundText = Empty
'
'   Set dtc_desc8.RowSource = Nothing
'   Set dtc_desc8.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_desc8.ReFill
'   dtc_desc8.BoundText = Empty
End Sub

Private Sub pnivel7(codigo9 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_edificaciones where munic_codigo = '" & codigo9 & "' order by edif_descripcion"
'   Set dtc_codigo10.RowSource = Nothing
'   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_codigo10.ReFill
'   dtc_codigo10.BoundText = Empty
'
'   Set dtc_desc10.RowSource = Nothing
'   Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_desc10.ReFill
'   dtc_desc10.BoundText = Empty
End Sub

Private Sub pnivel6(codigo8 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_calles where zona_codigo = '" & codigo8 & "' order by calle_denominacion"
'   Set dtc_codigo9.RowSource = Nothing
'   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_codigo9.ReFill
'   dtc_codigo9.BoundText = Empty
'
'   Set dtc_desc9.RowSource = Nothing
'   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_desc9.ReFill
'   dtc_desc9.BoundText = Empty
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    'If glPersNew = "FICHA" Then
        Call ABRIR_TABLA
    'Else
    '    Ado_datos.Recordset.AddNew
    'End If
    VAR_SW = "ADD"
    'VAR_PAIS = "BOL"
    fraDatos.Enabled = True
    FraGrabarCancelar.Visible = True
    txt_campo1.Visible = True
    txt_campo2.Visible = True
    txt_campo3.Visible = True

  Exit Sub
'AddErr:
'  MsgBox Err.Description
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
   'If glPersNew = "FICHA" Then
        Set rs_datos = New ADODB.Recordset
        If rs_datos.State = 1 Then rs_datos.Close
        'queryinicial = "select * from gc_beneficiario WHERE beneficiario_codigo = '" & glBenef & "'"
        'rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_datos.Open "select * from ro_personal_contratado WHERE beneficiario_codigo = '" & glBenef & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_datos.Sort = "beneficiario_denominacion"
        If rs_datos.RecordCount > 0 Then
        
        End If
        Set Ado_datos.Recordset = rs_datos
        'Set dg_datos.DataSource = Ado_datos.Recordset
        'Ado_datos.Recordset.MoveFirst
   'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If glPersNew = "NEWC" Then
'        Set rs_aux1 = New ADODB.Recordset
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & glBenef & "' ", db, adOpenStatic
'        Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
'        mw_solicitud.txt_ci = txt_codigo
'        mw_solicitud.txt_nombre.Visible = True
'        mw_solicitud.txt_nombre.Text = rs_aux1!beneficiario_denominacion
'        'Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
'        mw_solicitud.dtc_codigo4.Text = txt_codigo
'        mw_solicitud.dtc_desc4.BoundText = mw_solicitud.dtc_codigo4.BoundText
'        mw_solicitud.txt_obs = txt_codigo.Text + " - " + rs_aux1!beneficiario_denominacion + " - Telef. " + IIf(IsNull(rs_aux1!beneficiario_telefono_fijo), "0", rs_aux1!beneficiario_telefono_Cel)
'     End If
'  glPersNew = "N"
   
   If (rs_datos.State = adStateClosed) Then rs_datos.Close
   'Set rs_datos = Nothing
End Sub

Private Sub ABRIR_TABLAS_AUX()
    ' Banco 1
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "SELECT * FROM fc_bancos WHERE estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    ' Tipo de Cuenta 1
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "SELECT * FROM fc_cuenta_tipo WHERE estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    ' Banco 2
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "SELECT * FROM fc_bancos WHERE estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    ' Tipo de Cuenta 2
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "SELECT * FROM fc_cuenta_tipo WHERE estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    ' Banco 3
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "SELECT * FROM fc_bancos WHERE estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    ' Tipo de Cuenta 2
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "SELECT * FROM fc_cuenta_tipo WHERE estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
End Sub

Private Sub txt_campo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_campo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function ExisteBenef(CodBenef As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo_resp = '" & CodBenef & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteBenef = rs!Cuantos > 0
End Function

Private Function ExisteBenef2(CodBenef As String) As Boolean
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo = '" & CodBenef & "'"
    rs2.Open GlSqlAux, db, adOpenStatic
    ExisteBenef2 = rs2!Cuantos > 0
End Function

Private Sub pnivel2(codigo4 As String)
   Dim strConsultaF As String
     
   strConsultaF = "select * from gc_departamento where pais_codigo = '" & codigo4 & "'"
   Set dtc_codigo5.RowSource = Nothing
   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_codigo5.ReFill
   dtc_codigo5.BoundText = Empty
   
   Set dtc_desc5.RowSource = Nothing
   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_desc5.ReFill
   dtc_desc5.BoundText = Empty

End Sub

Private Sub pnivel3(codigo5 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo5 & "'"
'   Set dtc_codigo6.RowSource = Nothing
'   Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_codigo6.ReFill
'   dtc_codigo6.BoundText = Empty
'
'   Set dtc_desc6.RowSource = Nothing
'   Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_desc6.ReFill
'   dtc_desc6.BoundText = Empty
End Sub

Private Sub pnivel4(codigo6 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo6 & "'"
'   Set dtc_codigo7.RowSource = Nothing
'   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_codigo7.ReFill
'   dtc_codigo7.BoundText = Empty
'
'   Set dtc_desc7.RowSource = Nothing
'   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_desc7.ReFill
'   dtc_desc7.BoundText = Empty
End Sub


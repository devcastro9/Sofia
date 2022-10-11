VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fw_adjudica_gral 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Compra de Servicios - Adjudicación"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11010
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11115
      TabIndex        =   90
      Top             =   8160
      Width           =   11175
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "fw_adjudica_gral.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   92
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5640
         Picture         =   "fw_adjudica_gral.frx":07D6
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   91
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame fra_provedor 
      BackColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   360
      TabIndex        =   60
      Top             =   720
      Visible         =   0   'False
      Width           =   10215
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   9915
         TabIndex        =   87
         Top             =   3240
         Width           =   9975
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5040
            Picture         =   "fw_adjudica_gral.frx":10C2
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   89
            Top             =   120
            Width           =   1455
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3480
            Picture         =   "fw_adjudica_gral.frx":19AE
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   88
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtAutorizacionNew 
         Height          =   285
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   82
         Top             =   2520
         Width           =   2415
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   9915
         TabIndex        =   63
         Top             =   240
         Width           =   9975
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRO DE NUEVO PROVEEDOR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   2775
            TabIndex        =   64
            Top             =   360
            Width           =   4245
         End
      End
      Begin VB.TextBox txt_denominacion_new 
         Height          =   285
         Left            =   2520
         MaxLength       =   100
         TabIndex        =   17
         Top             =   1800
         Width           =   7455
      End
      Begin VB.TextBox txt_nit_new 
         Height          =   285
         Left            =   480
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "No.de Autorización:"
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
         TabIndex        =   81
         Top             =   2535
         Width           =   1740
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "DENOMINACION (Razon Social)"
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
         Left            =   2520
         TabIndex        =   62
         Top             =   1440
         Width           =   2925
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "NIT"
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
         TabIndex        =   61
         Top             =   1440
         Width           =   330
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif6 
      Height          =   330
      Left            =   4680
      Top             =   8760
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif6"
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
   Begin VB.PictureBox FraTitulo 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11115
      TabIndex        =   50
      Top             =   0
      Width           =   11175
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE FACTURA DEL PROVEEDOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   3060
         TabIndex        =   51
         Top             =   120
         Width           =   5085
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   360
      Top             =   9120
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif1"
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
   Begin MSAdodcLib.Adodc Ado_clasif2 
      Height          =   330
      Left            =   2520
      Top             =   9120
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif2"
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
   Begin MSAdodcLib.Adodc Ado_clasif3 
      Height          =   330
      Left            =   4680
      Top             =   9120
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif3"
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
   Begin MSAdodcLib.Adodc Ado_clasif4 
      Height          =   330
      Left            =   360
      Top             =   8760
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif4"
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
   Begin MSAdodcLib.Adodc Ado_clasif5 
      Height          =   330
      Left            =   2520
      Top             =   8760
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif5"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   7500
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   10695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK"
         Height          =   300
         Left            =   9840
         TabIndex        =   80
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "nro_autorizacion"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         MaxLength       =   50
         TabIndex        =   79
         Text            =   "%"
         Top             =   3600
         Width           =   270
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "nro_autorizacion"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   8760
         MaxLength       =   50
         TabIndex        =   78
         Text            =   "0"
         Top             =   3600
         Width           =   750
      End
      Begin VB.Frame fra_factura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "       FACTURA        "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   8280
         TabIndex        =   75
         Top             =   2520
         Width           =   2175
         Begin VB.OptionButton opt_no 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   1200
            TabIndex        =   77
            Top             =   320
            Width           =   555
         End
         Begin VB.OptionButton opt_si 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   360
            TabIndex        =   76
            Top             =   320
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.TextBox txt_13 
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   8760
         MaxLength       =   15
         TabIndex        =   73
         Text            =   "0"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txt_tipo_cambio 
         DataField       =   "nro_nota_remision"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   72
         Top             =   2760
         Width           =   1335
      End
      Begin VB.OptionButton opt_bs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   2640
         TabIndex        =   70
         Top             =   2760
         Width           =   555
      End
      Begin VB.OptionButton opt_usd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "USD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1920
         TabIndex        =   68
         Top             =   2760
         Width           =   675
      End
      Begin VB.Frame fra_almacen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ALMACEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   750
         Left            =   600
         TabIndex        =   66
         Top             =   5760
         Visible         =   0   'False
         Width           =   8685
         Begin MSDataListLib.DataCombo dtc_desc_alm 
            Bindings        =   "fw_adjudica_gral.frx":2184
            DataField       =   "almacen_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle2"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   305
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_descripcion"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_cod_alm 
            Bindings        =   "fw_adjudica_gral.frx":219E
            DataField       =   "almacen_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle2"
            Height          =   315
            Left            =   6720
            TabIndex        =   67
            Top             =   300
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_codigo"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
      End
      Begin VB.TextBox txt_descuentos 
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   6840
         MaxLength       =   15
         TabIndex        =   10
         Text            =   "0"
         Top             =   4320
         Width           =   1700
      End
      Begin VB.TextBox txt_importe_no_fiscal 
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   9
         Text            =   "0"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txt_nro_dui 
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "0"
         Top             =   3600
         Width           =   1700
      End
      Begin VB.TextBox txt_cod_control 
         DataField       =   "codigo_control"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   8760
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "0"
         Top             =   4845
         Width           =   1700
      End
      Begin VB.TextBox txt_autorizacion 
         DataField       =   "nro_autorizacion"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   6840
         MaxLength       =   50
         TabIndex        =   6
         Top             =   3600
         Width           =   1700
      End
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         DataField       =   "mes_grupo"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "0"
         Top             =   6960
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txt_total_bs 
         DataField       =   "compra_monto_bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   600
         MaxLength       =   20
         TabIndex        =   7
         Top             =   4320
         Width           =   1695
      End
      Begin VB.ComboBox cmd_unimed2 
         DataField       =   "unimed_codigo_pag"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   315
         ItemData        =   "fw_adjudica_gral.frx":21B8
         Left            =   6840
         List            =   "fw_adjudica_gral.frx":21CB
         TabIndex        =   15
         Top             =   6975
         Width           =   1875
      End
      Begin VB.TextBox txtCantCuota 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "cantidad_cuotas_pag"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         DataSource      =   "fw_compras_gral.ado_detalle2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Text            =   "1"
         Top             =   6975
         Width           =   1785
      End
      Begin VB.ComboBox cmb_mes_ini 
         DataField       =   "mes_inicio_crono"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   315
         ItemData        =   "fw_adjudica_gral.frx":21ED
         Left            =   720
         List            =   "fw_adjudica_gral.frx":2215
         TabIndex        =   13
         Top             =   6960
         Width           =   1620
      End
      Begin VB.TextBox txt_pais 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         MaxLength       =   80
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_Nota 
         DataField       =   "nro_nota_remision"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   285
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   4
         Top             =   3600
         Width           =   1700
      End
      Begin VB.TextBox txt_total_dol 
         DataField       =   "compra_monto_dol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   8
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4635
         MaxLength       =   80
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PROVEEDOR"
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   10245
         Begin VB.TextBox Text6 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   7550
            TabIndex        =   85
            Top             =   1095
            Visible         =   0   'False
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "fw_adjudica_gral.frx":227E
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle2"
            Height          =   315
            Left            =   2400
            TabIndex        =   2
            Top             =   480
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.CommandButton CmdAdd4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8760
            Picture         =   "fw_adjudica_gral.frx":2298
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Nuevo Registro"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   3105
            TabIndex        =   40
            Top             =   1095
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_aux4 
            Bindings        =   "fw_adjudica_gral.frx":2822
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "beneficiario_telefono_Cel"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9585
            TabIndex        =   41
            Top             =   1095
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "fw_adjudica_gral.frx":283C
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   1
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_aux5 
            Bindings        =   "fw_adjudica_gral.frx":2856
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle2"
            Height          =   315
            Left            =   3480
            TabIndex        =   29
            Top             =   1080
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "beneficiario_domicilio_legal"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_auto5 
            Bindings        =   "fw_adjudica_gral.frx":2870
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle2"
            Height          =   315
            Left            =   4320
            TabIndex        =   84
            Top             =   1080
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "comun_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Denominacion Proveedor"
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
            Left            =   2520
            TabIndex        =   56
            Top             =   240
            Width           =   2310
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ultima Autorización"
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
            Height          =   195
            Index           =   1
            Left            =   6720
            TabIndex        =   46
            Top             =   840
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.Label lblprov 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "NIT/CI Proveedor"
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
            Left            =   285
            TabIndex        =   39
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Teléfonos"
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
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   38
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dirección"
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
            Index           =   6
            Left            =   3525
            TabIndex        =   30
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         MaxLength       =   80
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtSW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5280
         MaxLength       =   80
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtFecha 
         DataField       =   "fecha_inicio_contrato"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   315
         Left            =   2520
         TabIndex        =   0
         Top             =   5640
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109379585
         CurrentDate     =   42248
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFecha2 
         DataField       =   "fecha_fin_contrato"
         DataSource      =   "ffw_compras_gral.ado_detalle2"
         Height          =   315
         Left            =   5280
         TabIndex        =   34
         Top             =   5640
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109379585
         CurrentDate     =   42248
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFecha3 
         DataField       =   "fecha_envio_proveedor"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   315
         Left            =   7995
         TabIndex        =   35
         Top             =   5640
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109379585
         CurrentDate     =   42248
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker txtfecha_compra 
         DataField       =   "fecha_compra"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Top             =   3600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109379585
         CurrentDate     =   42248
         MinDate         =   2
      End
      Begin VB.Label LblFactura 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   86
         Top             =   3345
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "No.Tramite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   83
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "13%"
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
         Height          =   195
         Left            =   9000
         TabIndex        =   74
         Top             =   4080
         Width           =   1425
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo cambio"
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
         Index           =   9
         Left            =   4320
         TabIndex        =   71
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Moneda"
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
         Height          =   240
         Index           =   8
         Left            =   600
         TabIndex        =   69
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuentos, bonificaciones"
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
         Height          =   435
         Left            =   6720
         TabIndex        =   65
         Top             =   3860
         Width           =   1785
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exento"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   58
         Top             =   4035
         Width           =   1425
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nro. DUI"
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
         Index           =   7
         Left            =   4800
         TabIndex        =   57
         Top             =   3345
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Control"
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
         Left            =   7320
         TabIndex        =   55
         Top             =   4845
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nro. Autorización"
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
         Left            =   6960
         TabIndex        =   54
         Top             =   3345
         Width           =   1515
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Factura/DUI"
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
         Index           =   5
         Left            =   600
         TabIndex        =   52
         Top             =   3345
         Width           =   1695
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha.Salida.de.Fabrica"
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
         Index           =   4
         Left            =   8040
         TabIndex        =   49
         Top             =   5400
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha.Fin.Fabricacion"
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
         Index           =   3
         Left            =   5280
         TabIndex        =   48
         Top             =   5400
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha.Inicio.Fabricacion"
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
         Left            =   2520
         TabIndex        =   47
         Top             =   5400
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   10920
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodicidad.de.Pago"
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
         Left            =   6735
         TabIndex        =   45
         Top             =   6720
         Width           =   1995
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Cuotas"
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
         Left            =   3810
         TabIndex        =   44
         Top             =   6705
         Width           =   1020
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes.Inicio.Pago"
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
         Left            =   720
         TabIndex        =   43
         Top             =   6705
         Width           =   1440
      End
      Begin VB.Label lbl_adjudica 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "adjudica_codigo"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9240
         TabIndex        =   37
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "No.Compra"
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
         Left            =   7680
         TabIndex        =   36
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe Dolares"
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
         Left            =   2640
         TabIndex        =   32
         Top             =   4035
         Width           =   1440
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe Bs."
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
         Left            =   600
         TabIndex        =   31
         Top             =   4035
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "No.Adjudica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   9195
         TabIndex        =   25
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label txtCodigo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "compra_codigo"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7680
         TabIndex        =   24
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incremento al Total"
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
         Left            =   8760
         TabIndex        =   23
         Top             =   3345
         Width           =   1710
      End
      Begin VB.Label lbl_campo_des 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidad Ejecutora"
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
         Left            =   2025
         TabIndex        =   22
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "fw_compras_gral.ado_detalle2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   21
         Top             =   495
         Width           =   1140
      End
      Begin VB.Label Txt_descripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2040
         TabIndex        =   20
         Top             =   495
         Width           =   5055
      End
   End
End
Attribute VB_Name = "fw_adjudica_gral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Para_Aceptado As String
Dim rs_clasif1 As New ADODB.Recordset
Dim rs_clasif2 As New ADODB.Recordset
Dim rs_clasif3 As New ADODB.Recordset
Dim rs_clasif4 As New ADODB.Recordset
Dim rs_clasif5 As New ADODB.Recordset
Dim rs_clasif6 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim Cadena As String
Dim Caracter(50, 50) As String

Dim monto_cuota, porcentaje_tot As Double
Dim VAR_OCUP, VAR_MED2, MControl As String
Dim mes_grupo, gestion, dia, fecha_pago As String

Dim VAR_COMPRA, CONT_MED, corrprog As Integer
Dim VAR_MES2, CONT3, CONT4, VAR_COBR2, ctrl  As Integer

Dim CUOTA, DOL, BS As Double
Dim FControl, FInicio As Date

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
    Para_Aceptado = "N"
    fw_compras_gral.Ado_detalle2.Recordset.CancelBatch
'    txtSW = "0"
    Unload Me
End Sub

Private Sub BtnGrabar_Click() ''acepta las modificaciones realizadas
'If txt_total_bs.Text = "" And txt_total_dol.Text = "" Then
'sino = MsgBox("Debe insertar el monto", vbCritical, Error)
'Exit Sub
'End If

If Valida Then
    Dim SQLS As String
    SQLS = ""
   'If txtSW = "ADD" Then
   '             '    fecha_recibe_almacen, almacen_codigo, poa_codigo, usr_codigo_aprueba, fecha_aprueba
   If swnuevo = 1 Then
      'DB.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, beneficiario_denominacion, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & dtc_codigo1.Text & ", '" & dtc_desc1.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      ''" & txtBenef.Caption & "',
       'DB.Execute "Insert INTO ao_solicitud_persona (ges_gestion, unidad_codigo, solicitud_codigo, benef_primer_apellido, benef_segundo_apellido, benef_nombres, benef_direccion_domicilio, benef_telefonos_ref, benef_codigo, puesto_codigo, ocup_codigo, munic_codigo, nivel_educ_codigo, observaciones, benef_fecha, estado_codigo, fecha_registro, usr_codigo) Values ('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
       '('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
     fw_compras_gral.Ado_detalle2.Recordset("ges_gestion") = glGestion
     fw_compras_gral.Ado_detalle2.Recordset("unidad_codigo") = Txt_campo1.Text
     fw_compras_gral.Ado_detalle2.Recordset("solicitud_codigo") = txt_codigo
     fw_compras_gral.Ado_detalle2.Recordset("compra_codigo").Value = fw_compras_gral.Ado_datos.Recordset!compra_codigo
     'fw_compras_gral.Ado_detalle2.Recordset("adjudica_codigo") = lbl_adjudica.Caption
     
      VAR_COMPRA = fw_compras_gral.Ado_datos.Recordset!compra_codigo
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      'fecha_registro  hora_registro usr_usuario
      VAR_COMPRA = fw_compras_gral.Ado_detalle2.Recordset("compra_codigo")
   End If
   
   Select Case Txt_campo1.Text
   
        Case "COMEX"
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "CMX"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "CMX-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "CMX-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-223"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "4.1.1"
             
        Case "DCONT"    'SOLO COMPRAS BB y SS   'FIN-03-01
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "FIN"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "FIN-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "FIN-03-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "ADM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-113"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "4.2.3"          'dtc_codigo10.Text
            'fw_compras_gral.Ado_detalle2.Recordset!solicitud_observaciones = dtc_desc2.Text + " - " + dtc_desc4.Text       ' txt_obs.Text
        
        Case "DVTA"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "COM-01-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-234"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
           ' fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.1.1"
        
        Case "DNINS", "DINSB", "DINSC", "DINSS"
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-362"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.2"
        
        Case "DNAJS", "DAJSB", "DAJSC", "DAJSS"
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-362"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.6"
        
        Case "DNMAN", "DMANB", "DMANC", "DMANS"
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-362"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.3"
            
        Case "DNREP", "DREPB", "DREPC", "DREPS"
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-362"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.4"
            
        Case "DNEME", "DEMEB", "DEMEC", "DEMES"
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-362"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.1"
            
        Case "DNMOD", "DMODB", "DMODC", "DMODS"
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-362"           'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.7"
        
        Case "UALMI", "ALMIB", "ALMIC", "ALMIS" 'INSUMOS
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "TEC"
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "TEC-06"
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "TEC-06-01"
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-306"
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.8"
            
        Case "UALMR", "ALMRB", "ALMRC", "ALMRS" 'REPUESTOS
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "TEC"
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "TEC-07"
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "TEC-07-01"
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-306"
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.5"
        
        Case "UALMH", "ALMB", "ALMC", "ALMS" 'HERRAMIENTAS
            fw_compras_gral.Ado_detalle2.Recordset!proceso_codigo = "TEC"
            fw_compras_gral.Ado_detalle2.Recordset!subproceso_codigo = "TEC-08"
            fw_compras_gral.Ado_detalle2.Recordset!clasif_codigo = "TEC"
            fw_compras_gral.Ado_detalle2.Recordset!etapa_codigo = "TEC-08-01"
            fw_compras_gral.Ado_detalle2.Recordset!doc_codigo = "R-306"
            'fw_compras_gral.Ado_detalle2.Recordset!poa_codigo = "3.2.9"
        
        End Select
'   fw_compras_gral.Ado_detalle2.Recordset("adjudica_fecha").Value = Format(Date, "dd/mm/yyyy")
'   fw_compras_gral.Ado_detalle2.Recordset("proceso_codigo") = "CMX"
'   fw_compras_gral.Ado_detalle2.Recordset("subproceso_codigo") = "CMX-01"
'   fw_compras_gral.Ado_detalle2.Recordset("etapa_codigo").Value = "CMX-01-01"
'
'   fw_compras_gral.Ado_detalle2.Recordset("clasif_codigo").Value = "CMX"
   
   fw_compras_gral.Ado_detalle2.Recordset!fecha_inicio_contrato = txtFecha.Value
   fw_compras_gral.Ado_detalle2.Recordset!fecha_fin_contrato = txtFecha2.Value
   fw_compras_gral.Ado_detalle2.Recordset!fecha_envio_proveedor = txtFecha3.Value
   
   fw_compras_gral.Ado_detalle2.Recordset("beneficiario_codigo").Value = dtc_codigo5.Text
   VAR_BENEF = fw_compras_gral.Ado_detalle2.Recordset!beneficiario_codigo

   fw_compras_gral.Ado_detalle2.Recordset("adjudica_descripcion").Value = dtc_desc5.Text         'frm_ao_compra_servicio.Ado_datos.Recordset!compra_descripcion
   fw_compras_gral.Ado_detalle2.Recordset("adjudica_cantidad_total").Value = fw_compras_gral.Ado_datos.Recordset!compra_cantidad_total
    
   'fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs") = txt_total_bs.Text
   
   fw_compras_gral.Ado_detalle2.Recordset("tipo_cambio").Value = txt_tipo_cambio.Text
   
   fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs").Value = CDbl(txt_total_bs.Text)
   fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_dol").Value = CDbl(txt_total_dol.Text)
   
   If opt_bs.Value = True Then
        fw_compras_gral.Ado_detalle2.Recordset("tipo_moneda").Value = "BOB"
   End If
   
   If opt_usd.Value = True Then
        fw_compras_gral.Ado_detalle2.Recordset("tipo_moneda").Value = "USD"
   End If
   
   fw_compras_gral.Ado_detalle2.Recordset("nro_nota_remision") = txt_Nota.Text
   fw_compras_gral.Ado_detalle2.Recordset("fecha_inicio_contrato").Value = txtFecha.Value
   fw_compras_gral.Ado_detalle2.Recordset("fecha_fin_contrato").Value = txtFecha2.Value
   fw_compras_gral.Ado_detalle2.Recordset("fecha_envio_proveedor") = txtFecha3.Value
    
   fw_compras_gral.Ado_detalle2.Recordset("mes_inicio_crono") = cmb_mes_ini
   fw_compras_gral.Ado_detalle2.Recordset("cantidad_cuotas_pag") = txtCantCuota
   fw_compras_gral.Ado_detalle2.Recordset("unimed_codigo_pag") = cmd_unimed2
    
   fw_compras_gral.Ado_detalle2.Recordset("usr_codigo") = glusuario
   fw_compras_gral.Ado_detalle2.Recordset("fecha_registro") = Date
   fw_compras_gral.Ado_detalle2.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
   fw_compras_gral.Ado_detalle2.Recordset("fecha_compra") = txtfecha_compra.Value
   fw_compras_gral.Ado_detalle2.Recordset("compra_codigo_det") = fw_compras_gral.Ado_detalle1.Recordset("compra_codigo_det")
  ' fw_compras_gral.Ado_detalle2.Recordset("almacen_codigo") = dtc_cod_alm.Text
   'fw_compras_gral.Ado_detalle1.Recordset("almacen_codigo") = dtc_cod_alm.Text

   'fw_compras_gral.Ado_detalle1.Recordset.Update
'   db.Execute "UPDATE ao_compra_detalle set estado_codigo = 'APR' WHERE compra_codigo = " & fw_compras_gral.Ado_datos.Recordset!compra_codigo & " AND compra_codigo_det = " & fw_compras_gral.Ado_detalle1.Recordset!compra_codigo_det & ""
'    sino = MsgBox("Desea APROBAR el Registro ? (Ya no podrá modificarlo)", vbYesNo + vbQuestion, "Atención")
'     If sino = vbYes Then
'
'       fw_compras_gral.Ado_detalle2.Recordset("estado_codigo") = "APR"
'       fw_compras_gral.Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
'       fw_compras_gral.Ado_detalle2.Recordset("fecha_aprueba") = Date
'       fw_compras_gral.Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
'       'db.Execute "update ao_compra_cabecera set estado_codigo_eqp = 'APR' WHERE compra_codigo = " & fw_compras_gral.Ado_detalle2.Recordset!compra_codigo & " "
'
'        db.Execute "update ac_bienes set bien_stock_ingreso =  av_compra_detalle_suma.compra_cantidad from ac_bienes, av_compra_detalle_suma where ac_bienes.bien_codigo = av_compra_detalle_suma.bien_codigo"
'        db.Execute "update ac_bienes set bien_stock_actual = bien_stock_inicial + bien_stock_ingreso - bien_stock_salida - bien_stock_salida_mant"
'
'         Set rs_aux7 = New ADODB.Recordset
'        If rs_aux7.State = 1 Then rs_aux7.Close
'        rs_aux7.Open "SELECT * FROM ac_almacenes WHERE almacen_codigo = " & dtc_cod_alm.Text & "", db, adOpenKeyset, adLockOptimistic
'        rs_aux7!correl_ing = IIf(IsNull(rs_aux7!correl_ing), 1, rs_aux7!correl_ing + 1)
'        fw_compras_gral.Ado_datos.Recordset("doc_numero") = rs_aux7!correl_ing
'        fw_compras_gral.Ado_datos.Recordset.Update
'        'db.Execute "UPDATE ao_compra_cabecera SET doc_numero = " & rs_aux7!correl_ing + 1 & " WHERE compra_codigo = " & fw_compras_gral.Ado_detalle2.Recordset!compra_codigo & ""
'        rs_aux7.Update
'
'        Set rs_aux6 = New ADODB.Recordset
'        If rs_aux6.State = 1 Then rs_aux6.Close
'        rs_aux6.Open "SELECT * FROM ac_bienes WHERE bien_codigo ='" & fw_compras_gral.Ado_detalle1.Recordset("bien_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'        rs_aux6!bien_total_compra_bs = rs_aux6!bien_total_compra_bs + fw_compras_gral.Ado_detalle1.Recordset("compra_cantidad")
'        rs_aux6.Update
'
'        Set rs_aux6 = New ADODB.Recordset
'        If rs_aux6.State = 1 Then rs_aux6.Close
'        rs_aux6.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo =" & dtc_cod_alm.Text & " AND bien_codigo = '" & fw_compras_gral.Ado_detalle1.Recordset("bien_codigo") & "'", db, adOpenStatic
'        If rs_aux6.RecordCount > 0 Then
'        db.Execute "update ao_almacen_totales set stock_ingreso  =" & Val(fw_compras_gral.Ado_detalle1.Recordset("compra_cantidad")) + rs_aux6!stock_ingreso & ", total_compra_bs =" & Val(fw_compras_gral.Ado_detalle1.Recordset("compra_cantidad") + rs_aux6!total_compra_bs) & ", stock_actual = " & Val(fw_compras_gral.Ado_detalle1.Recordset("compra_cantidad") + rs_aux6!stock_actual) & "WHERE almacen_codigo =" & dtc_cod_alm.Text & " AND bien_codigo = '" & fw_compras_gral.Ado_detalle1.Recordset("bien_codigo") & "'"
'        Else
'        db.Execute "INSERT INTO ao_almacen_totales (        almacen_codigo,                                                    bien_codigo,                                                   stock_ingreso,    stock_salida,                                                 stock_actual,                                                  total_compra_bs, total_venta_bs, utilidad_Bs,                                    total_compra_dol,                             total_venta_dol, utilidad_dol,estado_codigo, fecha_registro, usr_codigo)" & _
'                                                 "Values(" & dtc_cod_alm.Text & ", '" & fw_compras_gral.Ado_detalle1.Recordset("bien_codigo") & "', " & fw_compras_gral.Ado_detalle1.Recordset("compra_cantidad") & ", 0" & ", " & fw_compras_gral.Ado_detalle1.Recordset("compra_cantidad") & ", " & fw_compras_gral.Ado_detalle1.Recordset("compra_precio_total_bs") & ", 0, 0, " & fw_compras_gral.Ado_detalle1.Recordset("compra_precio_total_bs") / GlTipoCambioOficial & ", 0, 0, 'REG', " & Date & ", '" & glusuario & "')"
'        End If
'        db.Execute "INSERT INTO ao_almacen_ingresos (        ges_gestion,                                                         almacen_codigo,                                            doc_numero,                                                 bien_codigo,                                                edif_codigo,                                                  compra_codigo,           beneficiario_codigo,            fecha_ingreso,                                                   cantidad_ingreso,                                                   importe_compra_bs,     importe_compra_dol, estado_codigo, fecha_registro, usr_codigo)" & _
'                                                 "Values( " & fw_compras_gral.Ado_detalle1.Recordset("ges_gestion") & ", " & dtc_cod_alm.Text & ", " & fw_compras_gral.Ado_datos.Recordset!doc_numero & ", '" & fw_compras_gral.Ado_detalle1.Recordset("bien_codigo") & "', '" & fw_compras_gral.Ado_datos.Recordset!edif_codigo & "', " & fw_compras_gral.Ado_detalle1.Recordset("compra_codigo") & ", '" & dtc_codigo5.Text & "', " & txtfecha_compra.Value & ", " & fw_compras_gral.Ado_detalle1.Recordset("compra_cantidad") & ", " & fw_compras_gral.Ado_detalle1.Recordset("compra_precio_total_bs") & ", " & Val(fw_compras_gral.Ado_detalle1.Recordset("compra_precio_total_bs") / GlTipoCambioOficial) & ", 'REG', " & Date & ", '" & glusuario & "')"
'    Else
'       fw_compras_gral.Ado_detalle2.Recordset("estado_codigo") = "REG"
'    End If
   fw_compras_gral.Ado_detalle2.Recordset("estado_codigo") = "REG"
   fw_compras_gral.Ado_detalle2.Recordset("nro_autorizacion") = txt_autorizacion.Text
   
   fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs_87") = Val(txt_total_bs.Text) * 0.87
   fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_dol_87") = fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs_87") / GlTipoCambioOficial
   
   fw_compras_gral.Ado_detalle2.Recordset("nro_dui") = IIf(txt_nro_dui.Text = "", "0", txt_nro_dui.Text)
   fw_compras_gral.Ado_detalle2.Recordset("importe_no_credito_fisc") = IIf(txt_importe_no_fiscal.Text = "", "0", txt_importe_no_fiscal.Text)
   fw_compras_gral.Ado_detalle2.Recordset("sub_total") = Val(fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs") - fw_compras_gral.Ado_detalle2.Recordset("importe_no_credito_fisc"))
   fw_compras_gral.Ado_detalle2.Recordset("descuento") = IIf(txt_descuentos.Text = "", "0", txt_descuentos.Text)
   fw_compras_gral.Ado_detalle2.Recordset("importe_cred_fisc") = fw_compras_gral.Ado_detalle2.Recordset("sub_total") - fw_compras_gral.Ado_detalle2.Recordset("descuento")
   fw_compras_gral.Ado_detalle2.Recordset("credito_fiscal_13") = fw_compras_gral.Ado_detalle2.Recordset("importe_cred_fisc") * 0.13
   fw_compras_gral.Ado_detalle2.Recordset("tipo_compra") = "1"
   fw_compras_gral.Ado_detalle2.Recordset("literal") = Literal(fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs").Value)
   
   If fw_compras_gral.Ado_detalle1.Recordset("bien_codigo") = "479" Or fw_compras_gral.Ado_detalle1.Recordset("bien_codigo") = "3410007" Then
        fw_compras_gral.Ado_detalle2.Recordset("literal_neto") = Literal(fw_compras_gral.Ado_detalle2.Recordset("importe_cred_fisc").Value)
   Else
        fw_compras_gral.Ado_detalle2.Recordset("literal_neto") = Literal(fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs_87").Value)
   End If
   
   If opt_si.Value = True Then
        fw_compras_gral.Ado_detalle2.Recordset("factura") = "SI"
   Else
        fw_compras_gral.Ado_detalle2.Recordset("factura") = "NO"
   End If
   
'For i = 1 To Len(txt_cod_control.Text)
'Caracter(i, 1) = Mid(txt_cod_control.Text, i, 1)
'Next i
'
'Cadena = ""
'ctrl = 1
'For i = 1 To Len(txt_cod_control.Text)
'    If Caracter(i, 1) <> "-" Then
'        If ctrl Mod 2 = 0 Then
'            If i = Len(txt_cod_control.Text) Then
'                Cadena = Cadena & Caracter(i, 1)
'            Else
'                Cadena = Cadena & Caracter(i, 1) & "-"
'            End If
'        Else
'            Cadena = Cadena & Caracter(i, 1)
'
'        End If
'        ctrl = ctrl + 1
'    End If
'Next i
    fw_compras_gral.Ado_detalle2.Recordset("codigo_control") = IIf(txt_cod_control.Text = "", 0, txt_cod_control.Text)
   'fw_compras_gral.Ado_detalle2.Recordset("tipo_compra") = ""
   
   fw_compras_gral.Ado_detalle2.Recordset.Update
   lbl_adjudica.Caption = fw_compras_gral.Ado_detalle2.Recordset!adjudica_codigo
'    Set rs_aux7 = New ADODB.Recordset
'    If rs_aux7.State = 1 Then rs_aux7.Close
'    rs_aux7.Open "SELECT * FROM ao_compra_adjudica WHERE compra_codigo = " & VAR_COMPRA & "", db, adOpenKeyset, adLockOptimistic
'        rs_aux7!correl_ing = IIf(IsNull(rs_aux7!correl_ing), 1, rs_aux7!correl_ing + 1)
'        fw_compras_gral.Ado_datos.Recordset("doc_numero") = rs_aux7!correl_ing
'        fw_compras_gral.Ado_datos.Recordset.Update

    Set rs_aux7 = New Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "Select MAX(adjudica_codigo) AS CORREL from ao_compra_adjudica WHERE compra_codigo = " & VAR_COMPRA & " ", db, adOpenStatic
    If rs_aux7!CORREL <> "NULL" Then
       lbl_adjudica.Caption = rs_aux7!CORREL
    Else
       lbl_adjudica.Caption = "1"
    End If

   Para_Aceptado = "S"
   If Val(txt_total_bs.Text) > 0 Then
        Call CRONO_PAGO
   Else
        db.Execute "DELETE FROM ao_compra_planilla_pagos where adjudica_codigo = '" & lbl_adjudica.Caption & "' AND compra_codigo = " & VAR_COMPRA & ""
   End If
   db.Execute "update gc_beneficiario set comun_codigo = '" & txt_autorizacion.Text & "' where beneficiario_codigo = '" & dtc_codigo5.Text & "' "
   'frm_ao_solicitud_rrhh.ado_detalle2.Refresh '.Recordset.Requery
'   txtSW = "0"
   Unload Me
End If
End Sub

Private Sub CRONO_PAGO()

    Select Case RTrim(cmb_mes_ini.Text)
        Case "ENERO"
            txt_mes.Text = 1
        Case "FEBRERO"
            txt_mes.Text = 2
        Case "MARZO"
            txt_mes.Text = 3
        Case "ABRIL"
            txt_mes.Text = 4
        Case "MAYO"
            txt_mes.Text = 5
        Case "JUNIO"
            txt_mes.Text = 6
        Case "JULIO"
            txt_mes.Text = 7
        Case "AGOSTO"
            txt_mes.Text = 8
        Case "SEPTIEMBRE"
            txt_mes.Text = 9
        Case "OCTUBRE"
            txt_mes.Text = 10
        Case "NOVIEMBRE"
            txt_mes.Text = 11
        Case "DICIEMBRE"
            txt_mes.Text = 12
      End Select
    db.Execute "DELETE ao_compra_planilla_pagos where adjudica_codigo = '" & fw_compras_gral.Ado_detalle2.Recordset!adjudica_codigo & "' AND compra_codigo = " & VAR_COMPRA & ""

    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "select * from ao_compra_planilla_pagos", db, adOpenKeyset, adLockOptimistic
    mes_grupo = txt_mes.Text
    gestion = Year(txtfecha_compra.Value)
    CUOTA = 0
    monto_cuota = fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs").Value / Val(txtCantCuota.Text)
'gestion = Month(txtfecha_compra)
    If mes_grupo < Val(Month(txtfecha_compra)) Then
        gestion = gestion + 1
    End If
    dia = Day(txtfecha_compra.Value)

    While (txtCantCuota.Text <> CUOTA)
        fecha_pago = CDate(dia & "/" & mes_grupo & "/" & gestion)
        rs_aux2.AddNew

'fecha_pago = dia & "/" & mes_grupo & "/" & gestion

'If Weekday(fecha_pago, vbMonday) = 6 Then  'Pregunto si es Domingo
''dia = dia + 1
'fecha_pago = (dia + 1) & "/" & mes_grupo & "/" & gestion
'rs_aux2!pago_fecha_prog = DateAdd("m", mes_grupo, CDate(fecha_pago))
'ElseIf Weekday(fecha_pago, vbMonday) = 7 Then 'Pregunto si es Sabado
''dia = dia + 1
'fecha_pago = (dia + 2) & "/" & mes_grupo & "/" & gestion
'rs_aux2!pago_fecha_prog = DateAdd("m", mes_grupo, CDate(fecha_pago))
'Else
'rs_aux2!pago_fecha_prog = DateAdd("m", mes_grupo, CDate(fecha_pago))
'End If

        CUOTA = CUOTA + 1
        rs_aux2!ges_gestion = gestion
        rs_aux2!pago_codigo = CUOTA
        'rs_aux2!compra_codigo = fw_compras_gral.Ado_datos.Recordset!compra_codigo
        rs_aux2!compra_codigo = VAR_COMPRA
        rs_aux2!adjudica_codigo = lbl_adjudica.Caption
        rs_aux2!beneficiario_codigo = dtc_codigo5.Text
'rs_aux2!pago_emite_factura = monto_cuota / GlTipoCambioOficial
        rs_aux2!pago_fecha_prog = fecha_pago        'Format(CDate("29/02/2018"), "dd/mm/yyyy")    'CDate(Format(fecha_pago, "dd/mm/yyyy"))
    'rs_aux2!pago_fecha_efectiva = "0"
        rs_aux2!pago_monto_bs = monto_cuota
        rs_aux2!pago_monto_dol = monto_cuota / GlTipoCambioOficial
        rs_aux2!pago_descuento_bs = "0"
        rs_aux2!pago_total_bs = monto_cuota                         'fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs").Value
        rs_aux2!pago_total_dol = monto_cuota / GlTipoCambioOficial      'fw_compras_gral.Ado_detalle2.Recordset("adjudica_monto_bs").Value / GlTipoCambioOficial
    'rs_aux2!pago_nro_cmpbte_factura = txt_Nota.Text
    'rs_aux2!pago_nro_autorizacion = monto_cuota        '
    'rs_aux2!pago_respaldos = monto_cuota / GlTipoCambioOficial
        rs_aux2!Literal = ""
        rs_aux2!pago_descripcion = "CUOTA Nro." + Str(CUOTA) + " PAGO A: " + dtc_desc5.Text
        rs_aux2!estado_codigo = "REG"
        rs_aux2!usr_codigo = glusuario
        rs_aux2!Fecha_Registro = Date
        rs_aux2!hora_registro = Format(Time, "hh:mm:ss")
        If fw_compras_gral.Ado_detalle2.Recordset!factura = "SI" Then
            rs_aux2!pago_emite_factura = "S"
        Else
            rs_aux2!pago_emite_factura = "N"
        End If
        If fw_compras_gral.Ado_detalle1.Recordset!grupo_codigo = "20000" Then
            rs_aux2!bien_o_servicio = "S"
        Else
            rs_aux2!bien_o_servicio = "B"
        End If
        Select Case cmd_unimed2
            Case "MES"
                mes_grupo = mes_grupo + 1
            Case "BMES"
              mes_grupo = mes_grupo + 2
            Case "TMES"
              mes_grupo = mes_grupo + 3
            Case "SMES"
              mes_grupo = mes_grupo + 6
            Case "ANUAL"
               mes_grupo = mes_grupo + 12
        End Select

        If mes_grupo > 12 Then
            mes_grupo = mes_grupo - 12
            gestion = gestion + 1
        End If

        rs_aux2.Update
    Wend
'rs_datos.Update
  
End Sub

Public Function Literal(Cadena As String) As String
Dim SW As Integer
Dim sw1 As Integer
Dim swc As Integer
Dim VEC(20) As Long
SW = 0
      '*********PARTE DECIMAL*********
            If Cadena < 0 Then Cadena = Cadena * (-1)
            Cadena = Round(Cadena, 2)
             x = Len(Cadena)
              For k = 1 To x
                  Z = Mid(Cadena, k, 1)
                  If (Z = ".") Or SW = 1 Then
                    d = d + Mid(Cadena, k, 1)
                    SW = 1
                  End If
              Next k
              
              d = Mid(d, 2, Len(d))
              
              'Para la parte decimal del monto
              If d = "00" Or d = "" Then
                 d = d & " 00/100"
              Else
                 If d >= 0 And d <= 9 And Len(d) = 1 Then
                    d = " " & d & "0" & "/100"
                 Else
                    d = " " & d & "/100 "
                 End If
              End If
      '*********PARTE ENTERA*********
 If Cadena <> "" Then
    Cadena = Int(Cadena)
 Else
    MsgBox "No existe monto"
 End If
   s = ""
   Z = ""
   c = 0
   k = 0
   sw1 = 0
   swc = 0
   
   
   x = Len(Cadena)
   For i = 1 To x
       a = Mid(Cadena, i, 1)
       VEC(i) = Mid(Cadena, i, 1)
   Next i
j = x
While j <> 0
k = k + 1
If k <> 8 Then
  If c <> 3 Then
       c = c + 1
      
       If c = 1 And (VEC(j - 1) <> 1 And VEC(j - 1) <> 2) Then
            Select Case VEC(j)
                Case 0: s = " " + s
                Case 1:
                   If sw1 <> 1 Then
                      s = "UNO " + Z + s
                   End If
                   If sw1 = 1 Then
                      s = "UN " + Z + s
                   End If
                   
                Case 2: s = "DOS " + Z + s
                Case 3: s = "TRES " + Z + s
                Case 4: s = "CUATRO " + Z + s
                Case 5: s = "CINCO " + Z + s
                Case 6: s = "SEIS " + Z + s
                Case 7: s = "SIETE " + Z + s
                Case 8: s = "OCHO " + Z + s
                Case 9: s = "NUEVE " + Z + s
          End Select
          
           'If J + 1 <> "" And sw1 <> 1 And VEC(J - 1) <> 0 And VEC(J) <> 0 Then
           If VEC(j - 1) <> 0 And VEC(j) <> 0 Then
                 s = "Y " + s
           End If
           
        End If
        
         If c = 2 And VEC(j) = 1 Then
               swc = 1
                Select Case VEC(j + 1)
                      Case 0: s = "DIEZ " + Z + s
                      Case 1: s = "ONCE " + Z + s
                      Case 2: s = "DOCE " + Z + s
                      Case 3: s = "TRECE " + Z + s
                      Case 4: s = "CATORCE " + Z + s
                      Case 5: s = "QUINCE " + Z + s
                      Case 6: s = "DIECISEIS " + Z + s
                      Case 7: s = "DIECISIETE " + Z + s
                      Case 8: s = "DIECIOCHO " + Z + s
                      Case 9: s = "DIECINUEVE " + Z + s
                End Select
          End If
          
          If c = 2 And VEC(j) = 2 Then
                Select Case VEC(j + 1)
                      Case 0: s = "VEINTE " + Z + s
                      Case 1: s = "VEINTIUNO " + Z + s
                      Case 2: s = "VEINTIDOS " + Z + s
                      Case 3: s = "VEINTITRES " + Z + s
                      Case 4: s = "VEINTICUATRO " + Z + s
                      Case 5: s = "VEINTICINCO " + Z + s
                      Case 6: s = "VEINTISEIS " + Z + s
                      Case 7: s = "VEINTISIETE " + Z + s
                      Case 8: s = "VEINTIOCHO " + Z + s
                      Case 9: s = "VEINTINUEVE " + Z + s
                End Select
          End If
   
        If c = 2 Then
            Select Case VEC(j)
                Case 3: s = "TREINTA " + Z + s
                Case 4: s = "CUARENTA " + Z + s
                Case 5: s = "CINCUENTA " + Z + s
                Case 6: s = "SESENTA " + Z + s
                Case 7: s = "SETENTA " + Z + s
                Case 8: s = "OCHENTA " + Z + s
                Case 9: s = "NOVENTA " + Z + s
            End Select
            
        End If
        
        If c = 3 Then
            Select Case VEC(j)
                Case 1:
                If j = 1 Then
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                Else
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                       'S = "CIENTO " + z + S
                End If
                Case 2: s = "DOSCIENTOS " + Z + s
                Case 3: s = "TRESCIENTOS " + Z + s
                Case 4: s = "CUATROCIENTOS " + Z + s
                Case 5: s = "QUINIENTOS " + Z + s
                Case 6: s = "SEISCIENTOS " + Z + s
                Case 7: s = "SETECIENTOS " + Z + s
                Case 8: s = "OCHOCIENTOS " + Z + s
                Case 9: s = "NOVECIENTOS " + Z + s
            End Select
        End If
   Else
     If j >= 3 Then
            If VEC(j) = 0 And VEC(j - 1) = 0 And VEC(j - 2) = 0 Then
            Else
              s = "MIL " + s
            End If
    Else
              s = "MIL " + s
    End If
        j = j + 1
        c = 0
        sw1 = 1
   End If
 Else
    If VEC(j) <> 1 Then
      s = "MILLONES " + s
    Else
'      If K > 7 Then
      If k <> 8 Then
        s = "MILLONES " + s
      Else
        s = "MILLON " + s
      End If
    End If
      j = j + 1
      c = 0
      sw1 = 1
 End If
   j = j - 1
   
Wend

Literal = s + d
End Function

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
    Valida = True
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
    If (txt_Nota.Text = "") Then
        MsgBox "Debe registrar ... el Nro.Proforma/Factura", vbCritical + vbExclamation, "Validación de datos"
        Valida = False
    End If
'If (dtc_cod_alm.Text = "") Then
'    MsgBox "Debe registrar ... el ALMACEN", vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'
'  End If
  If (dtc_codigo5.Text = "") Then
    MsgBox "Debe registrar ... " + lblprov.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
  If opt_si.Value = True Then
    If txt_total_bs.Text = "" Or txt_total_bs.Text = "0" Then
      sino = MsgBox("Debe insertar el monto", vbCritical, Error)
      Valida = False
    End If
    If txt_total_bs.Text = "0" And txt_total_dol.Text = "0" Then
      sino = MsgBox("Debe insertar el monto", vbCritical, Error)
      Valida = False
    End If
  End If
  If txt_autorizacion = "" Then
    sino = MsgBox("Debe introducir Nro. De Autorizacion", vbCritical, Error)
    Valida = False
  End If
  
'  If txt_cod_control = "" Then
'    sino = MsgBox("Debe introducir el Código de Control", vbCritical, Error)
'    Valida = False
'  End If
  
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If (dtc_codigo4.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If txtPat = "" Then
'        Valida = False
'    End If
'    If txtNom = "" Then
'        Valida = False
'    End If
End Function

Private Sub cmb_mes_ini_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub

Private Sub CmdAdd4_Click()
    fra_provedor.Visible = True
    Frame1.Enabled = False
    dtc_codigo5.Text = ""
    FraGrabarCancelar.Visible = False
End Sub

Private Sub Command1_Click()
porcentaje_tot = 0
If opt_usd.Value = True Then
    If Text1.Text <> "" Then
        porcentaje_tot = CDbl((txt_total_dol.Text * Text1.Text) / 100)
        txt_total_dol.Text = Format(CDbl(txt_total_dol.Text + porcentaje_tot), "###,###,##0.00")
    End If
End If

If opt_bs.Value = True Then
    If Text1.Text <> "" Then
        porcentaje_tot = CDbl((txt_total_bs.Text * Text1.Text) / 100)
        txt_total_bs.Text = Format(CDbl(txt_total_bs.Text + porcentaje_tot), "###,###,##0.00")
    End If
End If

End Sub

Private Sub dtc_auto5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_auto5.BoundText
    dtc_aux4.BoundText = dtc_auto5.BoundText
    dtc_aux5.BoundText = dtc_auto5.BoundText
    dtc_desc5.BoundText = dtc_auto5.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux4.BoundText
    dtc_desc5.BoundText = dtc_aux4.BoundText
    dtc_aux5.BoundText = dtc_aux4.BoundText
    dtc_auto5.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux5.BoundText
    dtc_desc5.BoundText = dtc_aux5.BoundText
    dtc_aux4.BoundText = dtc_aux5.BoundText
    dtc_auto5.BoundText = dtc_aux5.BoundText
End Sub

Private Sub dtc_cod_alm_Click(Area As Integer)
    dtc_desc_alm.BoundText = dtc_cod_alm.BoundText
End Sub

Private Sub dtc_codigo5_Change()
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
    dtc_auto5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    dtc_aux4.BoundText = dtc_codigo5.BoundText
'    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo5_LostFocus()
If dtc_codigo5.Text <> "" Then
    If dtc_desc5.Text = "" Then
        sino = MsgBox("Este proveedor no existe, registre por favor", vbInformation, "SOFIA")
        txt_nit_new.Text = dtc_codigo5.Text
        fra_provedor.Visible = True
        Frame1.Enabled = False
    End If
End If
End Sub

Private Sub dtc_desc_alm_Change()
dtc_cod_alm.BoundText = dtc_desc_alm.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
    dtc_auto5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc5_LostFocus()
    txt_autorizacion.Text = dtc_auto5.Text
End Sub

Private Sub Form_Activate()
'    Set rs_clasif5 = New ADODB.Recordset
'    If rs_clasif5.State = 1 Then rs_clasif5.Close
''   Select Case Glaux
'    rs_clasif5.Open "SELECT * FROM gc_beneficiario ORDER BY beneficiario_denominacion ", db, adOpenStatic
''        Case "PROVI"    'PROVISION DE EQUIPOS
''            rs_clasif5.Open "SELECT * FROM gc_beneficiario where pais_codigo= '" & txt_pais.Text & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
''        Case "TRANS"    'TRANSPORTE
''            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
''        Case "ADUAN"    'DESADUANIZACION
''            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
''        Case "DESCA"    'DESCARGUIO Y OTROS
''            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
''    End Select
'    Set Ado_clasif5.Recordset = rs_clasif5
    DOL = txt_total_dol.Text
    BS = txt_total_bs.Text
    
    If parametro = "COMEX" Then
        opt_usd.Value = True
    End If
End Sub

Private Sub Form_Load()

    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    'Select Case Glaux
    rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5

'Set rs_clasif6 = New ADODB.Recordset
'    If rs_clasif6.State = 1 Then rs_clasif6.Close
'    'Select Case Glaux
'    rs_clasif6.Open "SELECT * FROM ac_almacenes where beneficiario_codigo = " & fw_compras_gral.dtc_codigo11.Text & " ORDER BY almacen_descripcion ", db, adOpenStatic
'    Set Ado_clasif6.Recordset = rs_clasif6
'
If parametro <> "COMEX" Then
    Set rs_clasif6 = New ADODB.Recordset
    If rs_clasif6.State = 1 Then rs_clasif6.Close
    'Select Case Glaux
    rs_clasif6.Open "SELECT * FROM ac_almacenes where beneficiario_codigo = " & fw_compras_gral.dtc_codigo11.Text & " ORDER BY almacen_descripcion ", db, adOpenStatic
     Set Ado_clasif6.Recordset = rs_clasif6
     'dtc_desc_alm.Enabled = True
     Text2.Visible = False
     Text1.Visible = False
     Command1.Visible = False
     lblLabels(0).Visible = False
Else
     Set rs_clasif6 = New ADODB.Recordset
     If rs_clasif6.State = 1 Then rFs_clasif6.Close
    'Select Case Glaux
     rs_clasif6.Open "SELECT * FROM ac_almacenes where almacen_codigo = 1", db, adOpenStatic
     Set Ado_clasif6.Recordset = rs_clasif6
     'dtc_desc_alm.Enabled = False
     dtc_desc_alm.BoundText = rs_clasif6!almacen_codigo
     dtc_cod_alm.Text = rs_clasif6!almacen_codigo
     Text2.Visible = True
     Text1.Visible = True
     Command1.Visible = True
     lblLabels(0).Visible = True
    
End If

    'txtSW = "0"
'    Set rs_clasif1 = New ADODB.Recordset
'    If rs_clasif1.State = 1 Then rs_clasif1.Close
'    rs_clasif1.Open "SELECT * FROM ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
'    Set Ado_clasif1.Recordset = rs_clasif1
'
'    Set rs_clasif2 = New ADODB.Recordset
'    If rs_clasif2.State = 1 Then rs_clasif2.Close
'    rs_clasif2.Open "SELECT * FROM gc_ocupacion_profesion ORDER BY ocup_descripcion ", db, adOpenStatic
'    Set Ado_clasif2.Recordset = rs_clasif2
'
'    Set rs_clasif3 = New ADODB.Recordset
'    If rs_clasif3.State = 1 Then rs_clasif3.Close
'    rs_clasif3.Open "SELECT * FROM rc_nivel_educacional ORDER BY nivel_educ_descripcion ", db, adOpenStatic
'    Set Ado_clasif3.Recordset = rs_clasif3
'
'    Set rs_clasif4 = New ADODB.Recordset
'    If rs_clasif4.State = 1 Then rs_clasif4.Close
'    rs_clasif4.Open "SELECT * FROM gc_municipio where region_codigo = 'SI' ORDER BY munic_descripcion ", db, adOpenStatic
'    Set Ado_clasif4.Recordset = rs_clasif4

End Sub

Private Sub opt_bs_Click()
    txt_total_dol.Enabled = False
    If txt_total_dol.Text <= "0" Or txt_total_dol.Text = "" Then
        txt_total_dol.Text = "0"
    End If
    txt_total_bs.Enabled = True
End Sub

Private Sub opt_no_Click()
    lblbien(5).Visible = True
    lblbien(5).Caption = "Fecha"
    txtfecha_compra.Visible = True
    'lblbien(0).Visible = False
    LblFactura.Caption = "Nro. Recibo"
    txt_Nota.Visible = True
    txt_Nota.Text = "0"
    lblbien(7).Visible = False
    txt_nro_dui.Visible = False
    Label2.Visible = False
    txt_autorizacion.Visible = False
    Label5.Visible = False
    txt_importe_no_fiscal.Visible = False
    Label9.Visible = False
    txt_descuentos.Visible = False
    Label11.Visible = False
    txt_13.Visible = False
    Label3.Visible = False
    txt_cod_control.Visible = False
    txt_autorizacion.Text = "0"
End Sub

Private Sub opt_si_Click()
    lblbien(5).Visible = True
    lblbien(5).Caption = "Fecha Factura/DUI"
    txtfecha_compra.Visible = True
    'lblbien(0).Visible = True
    LblFactura.Caption = "Nro. Factura"
    txt_Nota.Visible = True
    txt_Nota.Text = ""
    lblbien(7).Visible = True
    txt_nro_dui.Visible = True
    Label2.Visible = True
    txt_autorizacion.Visible = True
    txt_autorizacion.Text = ""
    Label5.Visible = True
    txt_importe_no_fiscal.Visible = True
    Label9.Visible = True
    txt_descuentos.Visible = True
    Label11.Visible = True
    txt_13.Visible = True
    Label3.Visible = True
    txt_cod_control.Visible = True
End Sub

Private Sub opt_usd_Click()
    txt_total_bs.Enabled = False
    If txt_total_bs.Text <= "0" Or txt_total_bs.Text = "" Then
    txt_total_bs.Text = "0"
    End If
    txt_total_dol.Enabled = True
End Sub

Private Sub Picture2_Click()
    fra_provedor.Visible = False
    Frame1.Enabled = True
    FraGrabarCancelar.Visible = True
End Sub

Private Sub Picture3_Click()
On Error GoTo UpdateErr
    db.Execute "INSERT INTO gc_beneficiario (beneficiario_codigo,      tipoben_codigo, tipodoc_codigo, beneficiario_nit,            beneficiario_denominacion,          comun_codigo,            estado_codigo, fecha_registro, usr_codigo)" & _
               "VALUES ('" & txt_nit_new.Text & "', '22',      '" & "NIT" & "', '" & txt_nit_new.Text & "', '" & txt_denominacion_new.Text & "', '" & TxtAutorizacionNew.Text & "', 'REG',     '" & Date & "', '" & glusuario & "')"

    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    'Select Case Glaux
    rs_clasif5.Open "SELECT * FROM gc_beneficiario ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5
    txt_autorizacion.Text = TxtAutorizacionNew.Text
    fra_provedor.Visible = False
    Frame1.Enabled = True
    dtc_desc5.BoundText = txt_nit_new.Text
    FraGrabarCancelar.Visible = True
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub txt_cod_control_Change()
For i = 1 To Len(txt_cod_control.Text)
Caracter(i, 1) = Mid(txt_cod_control.Text, i, 1)
Next i

Cadena = ""
ctrl = 1
For i = 1 To Len(txt_cod_control.Text)
    If Caracter(i, 1) <> "-" Then
        If ctrl Mod 2 = 0 Then
            If i = Len(txt_cod_control.Text) Then
                Cadena = Cadena & Caracter(i, 1)
            Else
                Cadena = Cadena & Caracter(i, 1) & "-"
            End If
        Else
            Cadena = Cadena & Caracter(i, 1)
        
        End If
        ctrl = ctrl + 1
    End If
Next i
txt_cod_control.Text = Cadena
txt_cod_control.SelStart = Len(txt_cod_control)
End Sub

Private Sub txt_cod_control_KeyPress(KeyAscii As Integer)
If KeyAscii <> 45 Then
    If KeyAscii <> 32 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
Else
    KeyAscii = 0
End If
End Sub



Private Sub txt_importe_no_fiscal_Change()

txt_13.Text = (CDbl(IIf(txt_total_bs.Text = "", 0, txt_total_bs.Text)) - CDbl(IIf(txt_importe_no_fiscal.Text = "", 0, txt_importe_no_fiscal.Text))) * 0.13
End Sub

Private Sub txt_tipo_cambio_Change()
On Error GoTo UpdateErr

If opt_bs.Value = True Then
    If txt_total_bs.Text <> "" And txt_total_bs.Text <> "," Then
        If txt_tipo_cambio.Text <> "" And txt_tipo_cambio.Text <> "," Then
            txt_total_dol.Text = CDbl(txt_total_bs.Text) * CDbl(txt_tipo_cambio.Text)
        Else
            txt_total_bs.Text = "0"
            txt_total_dol.Text = "0"
        End If
    Else
          txt_total_bs.Text = "0"
          txt_total_dol.Text = "0"
    End If
End If

If opt_usd.Value = True Then
    If txt_total_dol.Text <> "" And txt_total_dol.Text <> "," Then
        If txt_tipo_cambio.Text <> "" And txt_tipo_cambio.Text <> "," Then
            txt_total_bs.Text = CDbl(txt_total_dol.Text) * CDbl(txt_tipo_cambio.Text)
        Else
            txt_total_bs.Text = "0"
            txt_total_dol.Text = "0"
        End If
    Else
          txt_total_bs.Text = "0"
          txt_total_dol.Text = "0"
    End If
End If

Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub txt_tipo_cambio_KeyPress(KeyAscii As Integer)
If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub
'sino = VAR_SW
Private Sub txt_total_bs_Change()
 On Error GoTo UpdateErr
 
 
 If opt_bs.Value = True Then
    If txt_total_bs.Text <> "" And txt_total_bs.Text <> "," Then
        If txt_tipo_cambio.Text <> "" And txt_tipo_cambio.Text <> "," Then
            txt_total_dol.Text = Format(CDbl(txt_total_bs.Text) / CDbl(txt_tipo_cambio.Text), "###,###,##0.00")
        Else
          txt_total_dol.Text = "0"
        End If
    Else
          txt_total_dol.Text = "0"
    End If
End If

'If txt_total_bs.Text = "" Then
'txt_total_bs.Text = 0
'End If
'
'If txt_importe_no_fiscal.Text = "" Then
'txt_importe_no_fiscal.Text = 0
'End If
'
'txt_13.Text = (CDbl(txt_total_bs.Text) - CDbl(txt_importe_no_fiscal.Text)) * 0.13
If opt_usd.Value = True Then
'txt_total_dol.Text = CDbl(txt_total.Text / txt_tipo_cambio)

End If

If txt_total_bs > "0" And txt_total_dol > "0" Then
Label21.Visible = True
cmb_mes_ini.Visible = True

Label12.Visible = True
txtCantCuota.Visible = True
Label18.Visible = True
cmd_unimed2.Visible = True
Else
Label21.Visible = False
cmb_mes_ini.Visible = False
Label12.Visible = False
txtCantCuota.Visible = False
Label18.Visible = False
cmd_unimed2.Visible = False
End If

txt_13.Text = (CDbl(IIf(txt_total_bs.Text = "", 0, txt_total_bs.Text)) - CDbl(IIf(txt_importe_no_fiscal.Text = "", 0, txt_importe_no_fiscal.Text))) * 0.13
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub txt_total_bs_KeyPress(KeyAscii As Integer)

If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  

  '? . , 09
  ',.01234856789
End Sub


Private Sub txt_total_dol_Change()
On Error GoTo UpdateErr
''porcentaje_tot = IIf(txt_total_dol.Text = 0 Or txt_total_dol.Text = "", 0, txt_total_dol.Text)
'If opt_usd.Value = True Then
''txt_total_bs.Text = CDbl(txt_total_dol.Text * txt_tipo_cambio)
'End If
'If txt_total_bs.Text <> "0" Then
'txt_13.Text = (CDbl(IIf(txt_total_bs.Text = "", 0, txt_total_bs.Text)) - CDbl(IIf(txt_importe_no_fiscal.Text = "", 0, txt_importe_no_fiscal.Text))) * 0.13
'End If
If opt_usd.Value = True Then
    If txt_total_dol.Text <> "" And txt_total_dol.Text <> "," Then
        If txt_tipo_cambio.Text <> "" And txt_tipo_cambio.Text <> "," Then
            txt_total_bs.Text = Format(CDbl(txt_total_dol.Text) * CDbl(txt_tipo_cambio.Text), "###,###,##0.00")
        Else
            txt_total_bs.Text = "0"
        End If
    Else
          txt_total_bs.Text = "0"
    End If
End If
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub txt_total_dol_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
    Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub txt_total_dol_LostFocus()
'    If txt_total_dol.Text = "" Then
'        txt_total_dol.Text = "0"
'    End If
'    txt_total_bs.Text = CDbl(txt_total_dol) * GlTipoCambioOficial
End Sub


Private Sub txtfecha_compra_LostFocus()
    cmb_mes_ini.Text = UCase(MonthName(Month(txtfecha_compra.Value)))
    txt_mes.Text = Month(txtfecha_compra.Value)
    txtFecha.Value = txtfecha_compra.Value
    txtFecha2.Value = txtfecha_compra.Value
    txtFecha3.Value = txtfecha_compra.Value
End Sub

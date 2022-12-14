VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fw_adjudica_comex 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Adjudicación de Servicios"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   14550
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDUI 
      BackColor       =   &H00E0E0E0&
      Height          =   4575
      Left            =   1560
      TabIndex        =   119
      Top             =   1560
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox TxtDIM 
         Height          =   285
         Left            =   2400
         MaxLength       =   200
         TabIndex        =   131
         Top             =   2880
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "fw_adjudica_comex.frx":0000
         DataField       =   "nro_dui"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   315
         Left            =   4200
         TabIndex        =   129
         Top             =   1920
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin VB.ComboBox Cmbgestion 
         DataField       =   "ges_gestion"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "fw_adjudica_comex.frx":001A
         Left            =   2400
         List            =   "fw_adjudica_comex.frx":003F
         TabIndex        =   127
         Text            =   "2022"
         Top             =   1440
         Width           =   1275
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   10755
         TabIndex        =   123
         Top             =   3480
         Width           =   10815
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5520
            Picture         =   "fw_adjudica_comex.frx":0085
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   125
            Top             =   120
            Width           =   1455
         End
         Begin VB.PictureBox BtnGrabar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3960
            Picture         =   "fw_adjudica_comex.frx":0971
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   124
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   10755
         TabIndex        =   121
         Top             =   240
         Width           =   10815
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRO CODIGO DUI / DIM"
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
            Left            =   3915
            TabIndex        =   122
            Top             =   360
            Width           =   3405
         End
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   120
         Top             =   2400
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "fw_adjudica_comex.frx":115F
         DataField       =   "nro_dui"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   315
         Left            =   2400
         TabIndex        =   128
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Correlativo DUI / DIM"
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
         Left            =   345
         TabIndex        =   133
         Top             =   2880
         Width           =   1875
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Identificador"
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
         Left            =   1125
         TabIndex        =   132
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recinto Aduanero"
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
         TabIndex        =   130
         Top             =   1920
         Width           =   1620
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestion"
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
         Left            =   1440
         TabIndex        =   126
         Top             =   1440
         Width           =   810
      End
   End
   Begin VB.TextBox TxtMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      MaxLength       =   80
      TabIndex        =   104
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame FraQR 
      BackColor       =   &H00E0E0E0&
      Height          =   3735
      Left            =   240
      TabIndex        =   89
      Top             =   1560
      Visible         =   0   'False
      Width           =   14055
      Begin VB.TextBox TxtTexto 
         Height          =   285
         Left            =   240
         MaxLength       =   200
         TabIndex        =   95
         Top             =   1800
         Width           =   13455
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   13755
         TabIndex        =   93
         Top             =   240
         Width           =   13815
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRO MEDIANTE EL LECTOR DE QR"
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
            Left            =   4380
            TabIndex        =   94
            Top             =   360
            Width           =   4875
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   13755
         TabIndex        =   90
         Top             =   2520
         Width           =   13815
         Begin VB.PictureBox BtnGrabar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5400
            Picture         =   "fw_adjudica_comex.frx":1179
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   92
            Top             =   120
            Width           =   1335
         End
         Begin VB.PictureBox BtnCancelar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6960
            Picture         =   "fw_adjudica_comex.frx":1967
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   91
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lee el código QR de la FACTURA..."
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
         Left            =   240
         TabIndex        =   96
         Top             =   1440
         Width           =   3210
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   14235
      TabIndex        =   88
      Top             =   7200
      Width           =   14295
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5640
         Picture         =   "fw_adjudica_comex.frx":2253
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   29
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7320
         Picture         =   "fw_adjudica_comex.frx":2A29
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   30
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame fra_provedor 
      BackColor       =   &H00E0E0E0&
      Height          =   3855
      Left            =   240
      TabIndex        =   67
      Top             =   3240
      Visible         =   0   'False
      Width           =   14055
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   13755
         TabIndex        =   85
         Top             =   2640
         Width           =   13815
         Begin VB.PictureBox BtnCancelar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6840
            Picture         =   "fw_adjudica_comex.frx":3315
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   87
            Top             =   120
            Width           =   1455
         End
         Begin VB.PictureBox BtnGrabar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5280
            Picture         =   "fw_adjudica_comex.frx":3C01
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   86
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtAutorizacionNew 
         Height          =   285
         Left            =   10680
         MaxLength       =   50
         TabIndex        =   80
         Top             =   1800
         Width           =   2655
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   13755
         TabIndex        =   70
         Top             =   240
         Width           =   13815
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
            Left            =   4695
            TabIndex        =   71
            Top             =   360
            Width           =   4245
         End
      End
      Begin VB.TextBox txt_denominacion_new 
         Height          =   285
         Left            =   2520
         MaxLength       =   100
         TabIndex        =   33
         Top             =   1800
         Width           =   7935
      End
      Begin VB.TextBox txt_nit_new 
         Height          =   285
         Left            =   480
         MaxLength       =   50
         TabIndex        =   32
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
         Left            =   10680
         TabIndex        =   79
         Top             =   1455
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
         TabIndex        =   69
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
         TabIndex        =   68
         Top             =   1440
         Width           =   330
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif6 
      Height          =   330
      Left            =   4680
      Top             =   8040
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
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   14235
      TabIndex        =   60
      Top             =   -720
      Width           =   14295
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
         Left            =   4620
         TabIndex        =   61
         Top             =   120
         Width           =   5085
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   360
      Top             =   8400
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
      Top             =   8400
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
      Top             =   8400
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
      Top             =   8040
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
      Top             =   8040
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
      Height          =   7140
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   14295
      Begin VB.TextBox txt_total_eur 
         DataField       =   "adjudica_monto_eur"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7600
         MaxLength       =   20
         TabIndex        =   135
         Text            =   "0"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txt_tipo_cambio 
         DataField       =   "tipo_cambio"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   12480
         MaxLength       =   50
         TabIndex        =   115
         Top             =   5535
         Width           =   1575
      End
      Begin VB.CommandButton CmdCalcula 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Verificar Cálculos -->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "---------------------- TIPOS DE REGISTROS -----------------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   615
         Left            =   8880
         TabIndex        =   105
         Top             =   840
         Width           =   5175
         Begin VB.OptionButton Opt_DUI 
            BackColor       =   &H00C0C0C0&
            Caption         =   "DUI / DIM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   2280
            TabIndex        =   117
            Top             =   320
            Width           =   1275
         End
         Begin VB.OptionButton opt_gas 
            BackColor       =   &H00C0C0C0&
            Caption         =   "COMBUSTIBLE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   360
            TabIndex        =   7
            Top             =   320
            Width           =   1755
         End
         Begin VB.OptionButton opt_normal 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NORMAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   3720
            TabIndex        =   8
            Top             =   320
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.TextBox txt_87 
         DataField       =   "adjudica_monto_bs_87"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   12480
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   25
         Text            =   "0"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "--- TIPO DE MONEDA ----"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   6000
         TabIndex        =   101
         Top             =   840
         Width           =   2655
         Begin VB.OptionButton opt_eur 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Eur"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   1800
            TabIndex        =   134
            Top             =   320
            Width           =   675
         End
         Begin VB.OptionButton opt_usd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "USD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   240
            TabIndex        =   5
            Top             =   320
            Width           =   795
         End
         Begin VB.OptionButton opt_bs 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bs."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   1080
            TabIndex        =   6
            Top             =   320
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.TextBox txtSubTotal 
         DataField       =   "sub_total"
         DataSource      =   "fw_compras_comex.ado_detalle2"
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
         Height          =   285
         Left            =   12480
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   21
         Text            =   "0"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox Txt_tasa0 
         DataField       =   "grabado_tasa_cero"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   7600
         MaxLength       =   15
         TabIndex        =   19
         Text            =   "0"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Txt_Tasas 
         DataField       =   "tasas_ice_iehd"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   7600
         MaxLength       =   15
         TabIndex        =   18
         Text            =   "0"
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox txt_importe_no_fiscal 
         DataField       =   "importe_no_credito_fisc"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   7600
         MaxLength       =   15
         TabIndex        =   20
         Text            =   "0"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox TxtNIT_CGI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "nit_empresa"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "0"
         Top             =   6360
         Width           =   2055
      End
      Begin VB.CommandButton BtnQR 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         Picture         =   "fw_adjudica_comex.frx":43D7
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Nuevo Registro"
         Top             =   750
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   78
         Text            =   "%"
         Top             =   6600
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "nro_autorizacion"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   77
         Text            =   "0"
         Top             =   6600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Frame fra_factura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-------------------- FACTURA para: -----------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   615
         Left            =   240
         TabIndex        =   76
         Top             =   840
         Width           =   3975
         Begin VB.OptionButton opt_otro 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Terceros"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   1480
            TabIndex        =   114
            Top             =   320
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton opt_no 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sin/Fac"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   2760
            TabIndex        =   4
            Top             =   320
            Width           =   1035
         End
         Begin VB.OptionButton opt_si 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CGI-CGE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   320
            Width           =   1155
         End
      End
      Begin VB.TextBox txt_13 
         DataField       =   "credito_fiscal_13"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   12480
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   24
         Text            =   "0"
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Frame fra_almacen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fechas de Ejecución del Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   870
         Left            =   120
         TabIndex        =   73
         Top             =   6000
         Width           =   9765
         Begin MSDataListLib.DataCombo dtc_desc_alm 
            Bindings        =   "fw_adjudica_comex.frx":55E8
            DataField       =   "almacen_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   6000
            TabIndex        =   31
            Top             =   180
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_descripcion"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_cod_alm 
            Bindings        =   "fw_adjudica_comex.frx":5602
            DataField       =   "almacen_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   7440
            TabIndex        =   74
            Top             =   180
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_codigo"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtFecha 
            DataField       =   "fecha_inicio_contrato"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   2835
            TabIndex        =   108
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   111673345
            CurrentDate     =   44470
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker txtFecha2 
            DataField       =   "fecha_fin_contrato"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   7845
            TabIndex        =   109
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   111673345
            CurrentDate     =   44470
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker txtFecha3 
            DataField       =   "fecha_envio_proveedor"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   7875
            TabIndex        =   110
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   111673345
            CurrentDate     =   44470
            MinDate         =   32874
         End
         Begin VB.Label lblbien 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha Fin de de Proceso:"
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
            Left            =   5325
            TabIndex        =   112
            Top             =   360
            Width           =   2310
         End
         Begin VB.Label lblbien 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha Inicio de Proceso:"
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
            Left            =   495
            TabIndex        =   111
            Top             =   360
            Width           =   2220
         End
      End
      Begin VB.TextBox txt_CreditoFiscal 
         DataField       =   "importe_cred_fisc"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   12480
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   23
         Text            =   "0"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txt_descuentos 
         DataField       =   "descuento"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   12480
         MaxLength       =   15
         TabIndex        =   22
         Text            =   "0"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txt_nro_dui 
         DataField       =   "nro_dui"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2720
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "0"
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox txt_cod_control 
         DataField       =   "codigo_control"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   525
         Left            =   315
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "0"
         Top             =   5280
         Width           =   4215
      End
      Begin VB.TextBox txt_autorizacion 
         DataField       =   "nro_autorizacion"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   2720
         MaxLength       =   50
         TabIndex        =   14
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         DataField       =   "mes_grupo"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0"
         Top             =   6600
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txt_total_bs 
         DataField       =   "adjudica_monto_bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_comex.ado_detalle2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   7600
         MaxLength       =   20
         TabIndex        =   16
         Text            =   "0"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ComboBox cmd_unimed2 
         DataField       =   "unimed_codigo_pag"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   315
         ItemData        =   "fw_adjudica_comex.frx":561C
         Left            =   12120
         List            =   "fw_adjudica_comex.frx":562F
         TabIndex        =   28
         Top             =   6600
         Visible         =   0   'False
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
         DataSource      =   "fw_compras_comex.ado_detalle2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   27
         Text            =   "1"
         Top             =   6600
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox cmb_mes_ini 
         DataField       =   "mes_inicio_crono"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   315
         ItemData        =   "fw_adjudica_comex.frx":5651
         Left            =   2720
         List            =   "fw_adjudica_comex.frx":5679
         TabIndex        =   26
         Top             =   6600
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.TextBox txt_pais 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6000
         MaxLength       =   80
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_Nota 
         DataField       =   "nro_nota_remision"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   285
         Left            =   2720
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "0"
         Top             =   3600
         Width           =   1700
      End
      Begin VB.TextBox txt_total_dol 
         DataField       =   "adjudica_monto_dol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7600
         MaxLength       =   20
         TabIndex        =   17
         Text            =   "0"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4635
         MaxLength       =   80
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PROVEEDOR"
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   240
         TabIndex        =   43
         Top             =   1560
         Width           =   13845
         Begin VB.TextBox Text6 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   8865
            TabIndex        =   83
            Top             =   975
            Visible         =   0   'False
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "fw_adjudica_comex.frx":56E2
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   5160
            TabIndex        =   1
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.CommandButton CmdAdd4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   11880
            Picture         =   "fw_adjudica_comex.frx":56FC
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Nuevo Proveedor"
            Top             =   160
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   3945
            TabIndex        =   54
            Top             =   855
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_aux4 
            Bindings        =   "fw_adjudica_comex.frx":60FE
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   1200
            TabIndex        =   44
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
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
            Left            =   13305
            TabIndex        =   55
            Top             =   855
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "fw_adjudica_comex.frx":6118
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   1200
            TabIndex        =   0
            Top             =   120
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
            Bindings        =   "fw_adjudica_comex.frx":6132
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   5520
            TabIndex        =   45
            Top             =   840
            Width           =   8055
            _ExtentX        =   14208
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
            Bindings        =   "fw_adjudica_comex.frx":614C
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   5640
            TabIndex        =   82
            Top             =   960
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
         Begin MSDataListLib.DataCombo dtc_Nit5 
            Bindings        =   "fw_adjudica_comex.frx":6166
            DataField       =   "beneficiario_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle2"
            Height          =   315
            Left            =   1200
            TabIndex        =   118
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ForeColor       =   0
            ListField       =   "beneficiario_nit"
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
            Caption         =   "Denominacion"
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
            Left            =   3780
            TabIndex        =   65
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label lblprov 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "NIT ó CI "
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
            TabIndex        =   53
            Top             =   360
            Width           =   765
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
            TabIndex        =   52
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
            Left            =   4575
            TabIndex        =   46
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         MaxLength       =   80
         TabIndex        =   42
         Top             =   0
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
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtfecha_compra 
         DataField       =   "adjudica_fecha"
         DataSource      =   "fw_compras_comex.ado_detalle2"
         Height          =   315
         Left            =   2720
         TabIndex        =   11
         Top             =   3120
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111673345
         CurrentDate     =   44466
         MinDate         =   2
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equivalente en Euros:"
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
         Left            =   5595
         TabIndex        =   136
         Top             =   4080
         Width           =   1950
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo cambio:"
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
         Index           =   9
         Left            =   11160
         TabIndex        =   116
         Top             =   5520
         Width           =   1185
      End
      Begin VB.Label LblFechaFac 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de la Factura o DUI:"
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
         Left            =   100
         TabIndex        =   107
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Lbl_NitCgi 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NIT de la Empresa:"
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
         Height          =   255
         Left            =   9960
         TabIndex        =   106
         Top             =   6360
         Width           =   2055
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "87%:"
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
         Left            =   10960
         TabIndex        =   103
         Top             =   5040
         Width           =   1425
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. de DUI:"
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
         Left            =   240
         TabIndex        =   102
         Top             =   4080
         Width           =   2385
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe Neto p/Credito Fiscal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   9280
         TabIndex        =   100
         Top             =   4080
         Width           =   3180
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuentos, Bonificaciones:"
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
         Height          =   315
         Left            =   9800
         TabIndex        =   99
         Top             =   3600
         Width           =   2625
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sub TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   10880
         TabIndex        =   98
         Top             =   3120
         Width           =   1545
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ICE - IEHD - TASAS:"
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
         Left            =   5520
         TabIndex        =   97
         Top             =   4560
         Width           =   2025
      End
      Begin VB.Label LblFactura 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. de Factura:"
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
         Left            =   240
         TabIndex        =   84
         Top             =   3585
         Width           =   2385
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
         Left            =   240
         TabIndex        =   81
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "IVA 13%:"
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
         Left            =   10960
         TabIndex        =   75
         Top             =   4560
         Width           =   1425
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grabado Tasa cero:"
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
         Height          =   315
         Left            =   5640
         TabIndex        =   72
         Top             =   5025
         Width           =   1905
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exportacion u Operacion Exenta:"
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
         Height          =   315
         Left            =   4440
         TabIndex        =   66
         Top             =   5505
         Width           =   3105
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código de Control:"
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
         Left            =   960
         TabIndex        =   64
         Top             =   5025
         Width           =   1665
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nro. de  Autorización:"
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
         TabIndex        =   63
         Top             =   4545
         Width           =   2010
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
         Left            =   9975
         TabIndex        =   59
         Top             =   6600
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. de Cuotas"
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
         Left            =   6045
         TabIndex        =   58
         Top             =   6600
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes de Inicio de Pago"
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
         Left            =   510
         TabIndex        =   57
         Top             =   6600
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label lbl_adjudica 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "adjudica_codigo"
         DataSource      =   "fw_compras_comex.ado_detalle2"
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
         Left            =   12960
         TabIndex        =   51
         Top             =   255
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
         Left            =   9480
         TabIndex        =   50
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lbl_campo3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe Total en Dolares:"
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
         Left            =   5280
         TabIndex        =   48
         Top             =   3600
         Width           =   2265
      End
      Begin VB.Label lbl_campo2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total de la Factura en Bs.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   4755
         TabIndex        =   47
         Top             =   3105
         Width           =   2745
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
         Left            =   11805
         TabIndex        =   41
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
         DataSource      =   "fw_compras_comex.ado_detalle2"
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
         Left            =   10560
         TabIndex        =   40
         Top             =   255
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
         Left            =   2760
         TabIndex        =   39
         Top             =   6105
         Visible         =   0   'False
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
         Left            =   2600
         TabIndex        =   38
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
         DataSource      =   "fw_compras_comex.ado_detalle2"
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
         Left            =   1320
         TabIndex        =   37
         Top             =   255
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
         Left            =   4200
         TabIndex        =   36
         Top             =   255
         Width           =   5175
      End
   End
End
Attribute VB_Name = "fw_adjudica_comex"
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

Dim VAR_PROC, VAR_SUB, VAR_TAREA, VAR_CLASIF, VAR_POA As String
Dim VAR_OCUP, VAR_MED2, MControl As String
Dim mes_grupo, gestion, dia, fecha_pago As String
Dim VAR_DOC, VAR_GLOSA, VAR_MONEDA, FAC As String
Dim var_literal, VAR_LITERALN As String
Dim ES_QR, VAR_TRANS As String
Dim VAR_TIPOSOL, VAR_TRANSF As String
Dim usuario2, VAR_BENEF, VAR_BENEF2, VAR_DA As String

Dim VAR_COMPRA, CONT_MED, corrprog As Integer
Dim VAR_MES2, CONT3, CONT4, VAR_COBR2, ctrl  As Integer
Dim VAR_CANT, VAR_DET As Integer

Dim monto_cuota, porcentaje_tot As Double
Dim CUOTA, DOL, BS, VAR_PLANILLA As Double
Dim Bs87, Dol87, Bs13, VAR_SUBTOT, VAR_CREDFIS As Double
Dim VAR_PLANILLA_DOL As Double

Dim FControl, FInicio, FCompra As Date

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
    Para_Aceptado = "N"
    fw_compras_comex.Ado_detalle2.Recordset.CancelBatch
'    txtSW = "0"
    Unload Me
End Sub

Private Sub BtnCancelar3_Click()
    FraDUI.Visible = False
End Sub

Private Sub BtnGrabar_Click() ''acepta las modificaciones realizadas
    If (opt_si.Value = True) And (txt_total_bs.Text = "") Then      'And txt_total_dol.Text = ""
        sino = MsgBox("Debe registrar el monto", vbCritical, Error)
        Exit Sub
    End If

If Valida Then
   Dim SQLS As String
   SQLS = ""
   VAR_TRANS = "22"
   VAR_TIPOSOL = "1"
   VAR_TRANSF = "36"
   'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW JCQA2022
   Select Case txt_campo1.Text
    Case "COMEX"
        VAR_PROC = "CMX"
        Select Case Glaux
            Case "PROVI"
                VAR_SUB = "CMX-01"
                VAR_TAREA = "CMX-01-01"
            Case "TRANS"
                VAR_SUB = "CMX-02"
                VAR_TAREA = "CMX-21-01"
            Case "ADUAN"
                VAR_SUB = "CMX-03"
                VAR_TAREA = "CMX-03-01"
            Case "DESCA"
                VAR_SUB = "CMX-04"
                VAR_TAREA = "CMX-04-01"
            Case Else
                VAR_SUB = "CMX-05"
                VAR_TAREA = "CMX-05-01"
        End Select
        VAR_CLASIF = "COM"
        VAR_POA = "0"            '   "4.1.1"
    Case "DCONT"    'SOLO COMPRAS BB y SS   'FIN-03-01
        VAR_PROC = "FIN"
        VAR_SUB = "FIN-03"
        VAR_TAREA = "FIN-03-02"
        VAR_CLASIF = "ADM"
        VAR_POA = "0"            '   "4.2.3"
        'fw_compras_comex.Ado_detalle2.Recordset!solicitud_observaciones = dtc_desc2.Text + " - " + dtc_desc4.Text       ' txt_obs.Text
    Case "DVTA", "DCOMS", "DCOMB", "DCOMC"    ' COMPRA-VENTA BB Y SS - COMERCIAL
        VAR_PROC = "CMX"
        Select Case Glaux
            Case "PROVI"
                VAR_SUB = "CMX-01"
                VAR_TAREA = "CMX-01-01"
            Case "TRANS"
                VAR_SUB = "CMX-02"
                VAR_TAREA = "CMX-21-01"
            Case "ADUAN"
                VAR_SUB = "CMX-03"
                VAR_TAREA = "CMX-03-01"
            Case "DESCA"
                VAR_SUB = "CMX-04"
                VAR_TAREA = "CMX-04-01"
            Case Else
                VAR_SUB = "CMX-05"
                VAR_TAREA = "CMX-05-01"
        End Select
        VAR_CLASIF = "COM"
        VAR_POA = "0"            '   "4.1.1"
'        VAR_PROC = "COM"
'        VAR_SUB = "COM-01"
'        VAR_TAREA = "COM-01-02"
'        VAR_CLASIF = "COM"
'        VAR_POA = "0"            '   "3.1.1"
    Case "DNINS", "DINSB", "DINSC", "DINSS"
        VAR_PROC = "COM"
        VAR_SUB = "COM-03"
        VAR_TAREA = "COM-03-01"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.2"
    Case "DNAJS", "DAJSB", "DAJSC", "DAJSS"
        VAR_PROC = "COM"
        VAR_SUB = "COM-03"
        VAR_TAREA = "COM-03-01"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.6"
    Case "DNMAN", "DMANB", "DMANC", "DMANS"
        VAR_PROC = "TEC"
        VAR_SUB = "TEC-02"
        VAR_TAREA = "TEC-02-02"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.3"
    Case "DNREP", "DREPB", "DREPC", "DREPS"
        VAR_PROC = "TEC"
        VAR_SUB = "TEC-03"
        VAR_TAREA = "TEC-03-02"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.4"
    Case "DNEME", "DEMEB", "DEMEC", "DEMES"
        VAR_PROC = "TEC"
        VAR_SUB = "TEC-04"
        VAR_TAREA = "TEC-04-04"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.1"
    Case "DNMOD", "DMODB", "DMODC", "DMODS"
        VAR_PROC = "TEC"
        VAR_SUB = "TEC-05"
        VAR_TAREA = "TEC-05-04"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.7"
    Case "UALMI", "ALMIB", "ALMIC", "ALMIS" 'INSUMOS
        VAR_PROC = "TEC"
        VAR_SUB = "TEC-06"
        VAR_TAREA = "TEC-06-01"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.8"
    Case "UALMR", "ALMRB", "ALMRC", "ALMRS" 'REPUESTOS
        VAR_PROC = "TEC"
        VAR_SUB = "TEC-07"
        VAR_TAREA = "TEC-07-01"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.5"
    Case "UALMH", "ALMB", "ALMC", "ALMS" 'HERRAMIENTAS
        VAR_PROC = "TEC"
        VAR_SUB = "TEC-08"
        VAR_TAREA = "TEC-08-01"
        VAR_CLASIF = "TEC"
        VAR_POA = "0"            '   "3.2.9"
   End Select
   
   If opt_usd.Value = True Then
        VAR_MONEDA = "USD"
   Else
        VAR_MONEDA = "BOB"
   End If

   VAR_COMPRA = fw_compras_comex.Ado_datos.Recordset!compra_codigo
   FCompra = fw_compras_comex.Ado_datos.Recordset!compra_fecha
   VAR_BENEF = IIf(IsNull(fw_compras_comex.Ado_detalle2.Recordset!beneficiario_codigo), "0", fw_compras_comex.Ado_detalle2.Recordset!beneficiario_codigo)
   VAR_GLOSA = "SERVICIO DE " + TxtMenu.Text + " - Proveedor: " + RTrim(dtc_desc5.Text)
   VAR_CANT = fw_compras_comex.Ado_datos.Recordset!compra_cantidad_total
   VAR_DET = fw_compras_comex.Ado_detalle1.Recordset!compra_codigo_det
   
   If opt_gas.Value = True Then
        VAR_LITERALN = Literal(Round(CDbl(txt_CreditoFiscal.Text), 2))
        VAR_TRANS = "23"
   End If
   If opt_normal.Value = True Then
        VAR_LITERALN = Literal(Round(CDbl(txt_87.Text), 2))
        VAR_TRANS = "22"
   End If
   If Opt_DUI.Value = True Then
        VAR_LITERALN = Literal(Round(CDbl(txt_CreditoFiscal.Text), 2))
        VAR_TRANS = "24"
   End If
   
   If ES_QR = "SI" Then
        VAR_SUBTOT = Round(CDbl(txt_total_bs.Text), 2)
        VAR_CREDFIS = Round(CDbl(txt_CreditoFiscal.Text), 2)
        Bs13 = Round(CDbl(txt_13.Text), 2)
        Bs87 = Round(CDbl(VAR_CREDFIS) - CDbl(Bs13), 2)
        Dol87 = Round(Bs87 / GlTipoCambioOficial, 2)
        VAR_PLANILLA = VAR_SUBTOT
        VAR_PLANILLA_DOL = Round(VAR_SUBTOT / GlTipoCambioOficial, 2)
   Else
        If Opt_DUI.Value = True Then
            VAR_SUBTOT = Round(CDbl(txt_total_bs.Text), 2)
            VAR_CREDFIS = Round(CDbl(txt_CreditoFiscal.Text), 2)
            Bs13 = Round(CDbl(txt_13.Text), 2)
            Bs87 = Round(CDbl(txt_87.Text), 2)
            Dol87 = Round(Bs87 / GlTipoCambioOficial, 2)
            VAR_PLANILLA = Bs13
            VAR_PLANILLA_DOL = Round(Bs13 / GlTipoCambioOficial, 2)
            VAR_TRANS = "24"
            VAR_TIPOSOL = "1"
            VAR_TRANSF = "33"
        Else
            VAR_SUBTOT = Round(CDbl(txt_total_bs.Text) - CDbl(Txt_Tasas.Text) - CDbl(Txt_tasa0.Text) - CDbl(txt_importe_no_fiscal.Text), 2)
            VAR_CREDFIS = Round(VAR_SUBTOT - CDbl(txt_descuentos), 2)
            Bs13 = Round(CDbl(VAR_CREDFIS) * 0.13, 2)
            Bs87 = Round(CDbl(VAR_CREDFIS) - CDbl(Bs13), 2)
            Dol87 = Round(Bs87 / GlTipoCambioOficial, 2)
            VAR_PLANILLA = VAR_SUBTOT
            VAR_PLANILLA_DOL = Round(VAR_SUBTOT / GlTipoCambioOficial, 2)
        End If
   End If

   If opt_si.Value = True Then
        FAC = "SI"
        VAR_DOC = "R-101"    'FACTURA
        VAR_TRANSF = "33"
   Else
        FAC = "NO"
        VAR_DOC = "RE-402"    'Factura Comercial, Proforma, Purchase Order (Orden de Compra)  ' RECIBO
        VAR_TRANSF = "36"
   End If
   'If fw_compras_comex.Ado_detalle1.Recordset("bien_codigo") = "479" Or fw_compras_comex.Ado_detalle1.Recordset("bien_codigo") = "3410007" Then
   'LITERAL
   var_literal = Literal(CDbl(txt_total_bs))
   VAR_BENEF2 = IIf(fw_compras_comex.Ado_datos.Recordset!beneficiario_codigo_resp = "0", "4828818", fw_compras_comex.Ado_datos.Recordset!beneficiario_codigo_resp)
   If swnuevo = 1 Then

        db.Execute "Insert INTO ao_compra_adjudica (ges_gestion, compra_codigo, unidad_codigo, solicitud_codigo, fecha_compra, adjudica_fecha, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, " & _
                   " nro_nota_remision, beneficiario_codigo, adjudica_descripcion, adjudica_cantidad_total, adjudica_monto_bs, tipo_moneda, adjudica_monto_dol, fecha_inicio_contrato, fecha_fin_contrato, fecha_envio_proveedor, " & _
                   " fecha_recibe_almacen, almacen_codigo, poa_codigo, mes_inicio_crono, cantidad_cuotas_pag, unimed_codigo_pag, correl_pagos_prog, compra_codigo_det, observaciones, nro_autorizacion, codigo_control, nro_dui, " & _
                   " tasas_ice_iehd, grabado_tasa_cero, importe_no_credito_fisc, sub_total, descuento, importe_cred_fisc, credito_fiscal_13, adjudica_monto_bs_87, adjudica_monto_dol_87, tipo_compra, tipo_cambio, Literal, literal_neto, factura, " & _
                   " doc_codigo_alm, doc_numero_alm, estado_almacen, estado_codigo, usr_codigo, fecha_registro, hora_registro, usr_codigo_aprueba, fecha_aprueba, nit_empresa, trans_codigo, importe_planilla_bs, importe_planilla_dol, nit_beneficiario, beneficiario_codigo_resp, solicitud_tipo, trans_codigo_fac )  " & _
        " VALUES ('" & glGestion & "', " & VAR_COMPRA & ",  '" & txt_campo1.Text & "', " & Val(txt_codigo.Caption) & ", '" & FCompra & "', '" & txtfecha_compra.Value & "', '" & VAR_PROC & "', '" & VAR_SUB & "', '" & VAR_TAREA & "', '" & VAR_CLASIF & "', '" & VAR_DOC & "', '0', " & _
              " '" & txt_Nota & "', '" & dtc_codigo5.Text & "', '" & VAR_GLOSA & "', " & VAR_CANT & ", " & CDbl(txt_total_bs.Text) & ", '" & VAR_MONEDA & "', " & CDbl(txt_total_dol.Text) & ", '" & txtFecha.Value & "', '" & txtFecha2.Value & "', '" & txtFecha3.Value & "', " & _
              " '" & Date & "', '1', '" & VAR_POA & "', '" & cmb_mes_ini & "', " & txtCantCuota & ", '" & cmd_unimed2 & "', '1', " & Val(VAR_DET) & ", '" & RTrim(dtc_desc5.Text) & "', '" & txt_autorizacion.Text & "', '" & IIf(txt_cod_control.Text = "", "0", txt_cod_control.Text) & "', '" & txt_nro_dui.Text & "', " & _
              " '0', '0', " & CDbl(txt_importe_no_fiscal.Text) & ", " & VAR_SUBTOT & ", " & CDbl(txt_descuentos.Text) & ", " & VAR_CREDFIS & ", " & Bs13 & ", " & Bs87 & ", " & Dol87 & ", '1', " & GlTipoCambioOficial & ", '" & var_literal & "', '" & VAR_LITERALN & "', '" & FAC & "', " & _
              " '0', '0', 'REG', 'REG', '" & glusuario & "', '" & Date & "', '', '" & glusuario & "', '" & Date & "', '" & TxtNIT_CGI.Text & "', '" & VAR_TRANS & "', " & VAR_PLANILLA & ", " & VAR_PLANILLA_DOL & ", '" & dtc_Nit5.Text & "', '" & VAR_BENEF2 & "', " & VAR_TIPOSOL & ", '" & VAR_TRANSF & "' )"
        
   Else
        'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
       VAR_COMPRA = fw_compras_comex.Ado_detalle2.Recordset!compra_codigo
       
       fw_compras_comex.Ado_detalle2.Recordset("beneficiario_codigo").Value = dtc_codigo5.Text
       VAR_BENEF = fw_compras_comex.Ado_detalle2.Recordset!beneficiario_codigo

       fw_compras_comex.Ado_detalle2.Recordset("adjudica_descripcion").Value = VAR_GLOSA    'dtc_desc5.Text
       'fw_compras_comex.Ado_detalle2.Recordset("adjudica_cantidad_total").Value = fw_compras_comex.Ado_datos.Recordset!compra_cantidad_total
       fw_compras_comex.Ado_detalle2.Recordset("tipo_cambio").Value = txt_tipo_cambio.Text
       fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_bs").Value = CDbl(txt_total_bs.Text)
       fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_dol").Value = CDbl(txt_total_dol.Text)
       fw_compras_comex.Ado_detalle2.Recordset("nro_nota_remision") = txt_Nota.Text
       fw_compras_comex.Ado_detalle2.Recordset("fecha_inicio_contrato").Value = txtFecha.Value
       fw_compras_comex.Ado_detalle2.Recordset("fecha_fin_contrato").Value = txtFecha2.Value
       fw_compras_comex.Ado_detalle2.Recordset("fecha_envio_proveedor") = txtFecha3.Value
        fw_compras_comex.Ado_detalle2.Recordset!adjudica_fecha = txtfecha_compra.Value
       fw_compras_comex.Ado_detalle2.Recordset("mes_inicio_crono") = cmb_mes_ini
       fw_compras_comex.Ado_detalle2.Recordset("cantidad_cuotas_pag") = txtCantCuota
       fw_compras_comex.Ado_detalle2.Recordset("unimed_codigo_pag") = cmd_unimed2
        
       fw_compras_comex.Ado_detalle2.Recordset("usr_codigo") = glusuario
       fw_compras_comex.Ado_detalle2.Recordset("fecha_registro") = Date
       fw_compras_comex.Ado_detalle2.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
       fw_compras_comex.Ado_detalle2.Recordset!beneficiario_codigo_resp = VAR_BENEF2
       'fw_compras_comex.Ado_detalle2.Recordset("compra_codigo_det") = fw_compras_comex.Ado_detalle1.Recordset("compra_codigo_det")
       fw_compras_comex.Ado_detalle2.Recordset!solicitud_tipo = VAR_TIPOSOL
       fw_compras_comex.Ado_detalle2.Recordset!trans_codigo_fac = VAR_TRANSF
       'fw_compras_comex.Ado_detalle2.Recordset("estado_codigo") = "REG"
       fw_compras_comex.Ado_detalle2.Recordset("nro_autorizacion") = txt_autorizacion.Text
       
       fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_bs_87") = Bs87       'Val(txt_total_bs.Text) * 0.87
       fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_dol_87") = Dol87     'fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_bs_87") / GlTipoCambioOficial
       
       fw_compras_comex.Ado_detalle2.Recordset("nro_dui") = IIf(txt_nro_dui.Text = "", "0", txt_nro_dui.Text)
       fw_compras_comex.Ado_detalle2.Recordset("importe_no_credito_fisc") = IIf(txt_importe_no_fiscal.Text = "", "0", txt_importe_no_fiscal.Text)
       fw_compras_comex.Ado_detalle2.Recordset("sub_total") = VAR_SUBTOT
       fw_compras_comex.Ado_detalle2.Recordset("descuento") = IIf(txt_descuentos.Text = "", "0", txt_descuentos.Text)
       fw_compras_comex.Ado_detalle2.Recordset("importe_cred_fisc") = VAR_CREDFIS       'fw_compras_comex.Ado_detalle2.Recordset("sub_total") - fw_compras_comex.Ado_detalle2.Recordset("descuento")
       fw_compras_comex.Ado_detalle2.Recordset("credito_fiscal_13") = Bs13              'fw_compras_comex.Ado_detalle2.Recordset("importe_cred_fisc") * 0.13
       'fw_compras_comex.Ado_detalle2.Recordset("tipo_compra") = "1"
       fw_compras_comex.Ado_detalle2.Recordset("importe_planilla_bs") = VAR_PLANILLA
       fw_compras_comex.Ado_detalle2.Recordset("importe_planilla_dol") = VAR_PLANILLA_DOL
       fw_compras_comex.Ado_detalle2.Recordset("literal") = var_literal
       fw_compras_comex.Ado_detalle2.Recordset("literal_neto") = VAR_LITERALN
       fw_compras_comex.Ado_detalle2.Recordset!nit_beneficiario = dtc_Nit5.Text
       fw_compras_comex.Ado_detalle2.Recordset!factura = FAC
       'If fw_compras_comex.Ado_detalle1.Recordset("bien_codigo") = "479" Or fw_compras_comex.Ado_detalle1.Recordset("bien_codigo") = "3410007" Then
       '     fw_compras_comex.Ado_detalle2.Recordset("literal_neto") = Literal(fw_compras_comex.Ado_detalle2.Recordset("importe_cred_fisc").Value)
       'Else
       '     fw_compras_comex.Ado_detalle2.Recordset("literal_neto") = Literal(fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_bs_87").Value)
       'End If
       fw_compras_comex.Ado_detalle2.Recordset("codigo_control") = IIf(txt_cod_control.Text = "", 0, txt_cod_control.Text)
       fw_compras_comex.Ado_detalle2.Recordset.Update
   End If
   Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select max(adjudica_codigo) as Codigo from ao_compra_adjudica where compra_codigo = " & VAR_COMPRA & " ", db, adOpenStatic
    If Not rs_aux6.EOF Then
        lbl_adjudica.Caption = IIf(IsNull(rs_aux6!Codigo), 1, rs_aux6!Codigo)
    Else
        lbl_adjudica.Caption = 1
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
    db.Execute "DELETE ao_compra_planilla_pagos where adjudica_codigo = " & Val(lbl_adjudica.Caption) & " AND compra_codigo = " & VAR_COMPRA & " "

    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "select * from ao_compra_planilla_pagos", db, adOpenKeyset, adLockOptimistic
    mes_grupo = txt_mes.Text
    gestion = Year(txtfecha_compra.Value)
    CUOTA = 0
    monto_cuota = CDbl(txt_total_bs.Text) / Val(txtCantCuota.Text)
'gestion = Month(txtfecha_compra)
    ' fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_bs").Value
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
        'rs_aux2!compra_codigo = fw_compras_comex.Ado_datos.Recordset!compra_codigo
        rs_aux2!compra_codigo = VAR_COMPRA
        rs_aux2!adjudica_codigo = lbl_adjudica.Caption
        rs_aux2!beneficiario_codigo = dtc_codigo5.Text
    'rs_aux2!pago_emite_factura = monto_cuota / GlTipoCambioOficial
        rs_aux2!pago_fecha_prog = fecha_pago        'Format(CDate("29/02/2018"), "dd/mm/yyyy")    'CDate(Format(fecha_pago, "dd/mm/yyyy"))
    'rs_aux2!pago_fecha_efectiva = "0"
        rs_aux2!pago_monto_bs = monto_cuota
        rs_aux2!pago_monto_dol = monto_cuota / GlTipoCambioOficial
        rs_aux2!pago_descuento_bs = "0"
        rs_aux2!pago_total_bs = monto_cuota                         'fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_bs").Value
        rs_aux2!pago_total_dol = monto_cuota / GlTipoCambioOficial      'fw_compras_comex.Ado_detalle2.Recordset("adjudica_monto_bs").Value / GlTipoCambioOficial
    'rs_aux2!pago_nro_cmpbte_factura = txt_Nota.Text
    'rs_aux2!pago_nro_autorizacion = monto_cuota        '
    'rs_aux2!pago_respaldos = monto_cuota / GlTipoCambioOficial
        rs_aux2!Literal = ""
        rs_aux2!pago_descripcion = "CUOTA Nro." + Str(CUOTA) + " PAGO A: " + dtc_desc5.Text
        rs_aux2!estado_codigo = "REG"
        rs_aux2!usr_codigo = glusuario
        rs_aux2!fecha_registro = Date
        rs_aux2!hora_registro = Format(Time, "hh:mm:ss")
        If fw_compras_comex.Ado_detalle2.Recordset!factura = "SI" Then
            rs_aux2!pago_emite_factura = "S"
        Else
            rs_aux2!pago_emite_factura = "N"
        End If
        If fw_compras_comex.Ado_detalle1.Recordset!grupo_codigo = "20000" Then
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
Private Sub GRABA_FICHA()
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "SELECT * FROM ro_rrhh_apertura_sobres where rrhh_codigo = " & frm_ao_compra_servicio.Ado_datos.Recordset!rrhh_codigo & "  ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        VAR_OCUP = rs_aux3!ocup_codigo
    Else
        VAR_OCUP = "0"
    End If
    
''    db.Execute "Insert INTO ro_personal_contratado_new (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & frm_ao_compra_servicio.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
''    db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & frm_ao_compra_servicio.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
'
'    Set rs_aux2 = New ADODB.Recordset
'    If rs_aux2.State = 1 Then rs_aux2.Close
'    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
'    rs_aux2.Open "SELECT * FROM rc_puestos where puesto_codigo = '" & GlPuesto & "'  ", db, adOpenStatic
'    If rs_aux2.RecordCount > 0 Then
'        db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, unidad_codigo, cargo_codigo, fecha_ingreso, fecha_expiracion, ocup_codigo, beneficiario_haber_mensual, estado_codigo, usr_codigo, fecha_registro) Values ('" & frm_ao_compra_servicio.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "', '" & rs_aux2!unidad_codigo & "',  '" & rs_aux2!cargo_codigo & "',  '" & frm_ao_compra_servicio.Ado_detalle2.Recordset!beneficiario_fecha_inicio & "', '" & frm_ao_compra_servicio.Ado_detalle2.Recordset!beneficiario_fecha_fin & "', '" & VAR_OCUP & "', " & frm_ao_compra_servicio.Ado_detalle2.Recordset!beneficiario_monto_adjudica_dol & ", 'REG', '" & glusuario & "',  '" & Date & "')"
'        'db.Execute "Insert INTO ro_personal_contratado_NEW (rrhh_codigo, beneficiario_codigo, puesto_codigo, unidad_codigo, cargo_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & frm_ao_compra_servicio.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "', '" & rs_aux2!unidad_codigo & "',  '" & rs_aux2!cargo_codigo & "',  'REG', '" & glusuario & "',  '" & Date & "')"
'    Else
'        db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & frm_ao_compra_servicio.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
'    End If
'    'Set Ado_clasif1.Recordset = rs_aux2

End Sub

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
    Valida = True
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
If (txt_Nota.Text = "") Then
    MsgBox "Debe registrar ... el Nro.Factura o Recibo", vbCritical + vbExclamation, "Validación de datos"
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
  If txt_autorizacion = "" And opt_si.Value = True Then
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

Private Sub BtnGrabar2_Click()
    Dim VAR_POSICION As Integer
    Dim VAR_BUSCA As Integer
    Dim VAR_EXTRAER As String
    
    Dim VAR_NIT_PROV  As String
    Dim VAR_FAC As String
    Dim VAR_AUTORIZA As String
    Dim VAR_FECHA As String
    Dim VAR_TOTAL As Double
    Dim VAR_BsCredFiscal As Double
    Dim VAR_CONTROL As String
    Dim VAR_NIT_CGI As String
    Dim VAR_MONTO1 As Double
    Dim VAR_MONTO2 As Double
    Dim VAR_CredFiscal As Double
    Dim VAR_MONTO3 As Double
    Dim VAR_MONTO4 As Double
    
'    opt_bs.Value = True
    
    '2325709012|78293|277405100157455|29/06/2021|20.00|14.00|DA-C3-2C-CA|1018533029|0|0|6.00|0
    '154094027|3265|387401100142890|07/07/2021|15.00|15.00|C2-64-45-2A-61|101853029|0|0|0|0
    'NIT_PROV  FAC   AUTOR   FECHA   total   Impor.Cred.Fis  COD_CONTROL NIT_CGI 0 0 Excento.Cred.Fis Dscto
    
    VAR_BUSCA = InStr(1, TxtTexto.Text, "|", 1)
    VAR_NIT_PROV = Mid(TxtTexto.Text, 1, VAR_BUSCA - 1)
    dtc_codigo5.Text = VAR_NIT_PROV
    
    VAR_POSICION = VAR_BUSCA + 1
    VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
    VAR_FAC = Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION))
    txt_Nota.Text = VAR_FAC
    
    VAR_POSICION = VAR_BUSCA + 1
    VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
    VAR_AUTORIZA = Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION))
    txt_autorizacion.Text = VAR_AUTORIZA
    
    VAR_POSICION = VAR_BUSCA + 1
    VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
    VAR_FECHA = Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION))
    txtfecha_compra.Value = Format(VAR_FECHA, "dd/mm/yyyy")
    
    VAR_POSICION = VAR_BUSCA + 1
    VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
    VAR_TOTAL = CDbl(Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION)))
    txt_total_bs.Text = VAR_TOTAL
        
    VAR_POSICION = VAR_BUSCA + 1
    VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
    VAR_BsCredFiscal = CDbl(Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION)))
    txt_CreditoFiscal.Text = VAR_BsCredFiscal
    
    VAR_POSICION = VAR_BUSCA + 1
    VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
    VAR_CONTROL = Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION))
    txt_cod_control.Text = VAR_CONTROL
    
    VAR_POSICION = VAR_BUSCA + 1
    VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
    VAR_NIT_CGI = Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION))
    TxtNIT_CGI.Text = VAR_NIT_CGI
    
    If (VAR_NIT_CGI <> "1018533029" And VAR_NIT_CGI <> "125887020") And opt_si.Value = True Then
        MsgBox "El NIT de la Empresa es incorrecto, NO se copiarán los datos de la FACTURA. Verifique y vuelva a intentar... ", vbInformation, "SOFIA"
        dtc_codigo5.Text = "0"
        txt_Nota.Text = "0"
        txt_autorizacion.Text = "0"
        txt_total_bs.Text = "0"
        txt_CreditoFiscal.Text = "0"
        txt_cod_control.Text = ""
    End If
    
    If (VAR_NIT_CGI = "1018533029" Or VAR_NIT_CGI = "125887020") And opt_otro.Value = True Then
        MsgBox "La Factura es válida para Crédito Fiscal, por lo que debe elegir la opción: <FACTURA para ...Fac.CGI>. Verifique y vuelva a intentar... ", vbInformation, "SOFIA"
        dtc_codigo5.Text = "0"
        txt_Nota.Text = "0"
        txt_autorizacion.Text = "0"
        txt_total_bs.Text = "0"
        txt_CreditoFiscal.Text = "0"
        txt_cod_control.Text = ""
    End If
    
        VAR_POSICION = VAR_BUSCA + 1
        VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
        VAR_MONTO1 = CDbl(Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION)))
        Txt_Tasas.Text = VAR_MONTO1
        
        VAR_POSICION = VAR_BUSCA + 1
        VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
        VAR_MONTO2 = CDbl(Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION)))
        Txt_tasa0.Text = VAR_MONTO2
        
        VAR_POSICION = VAR_BUSCA + 1
        VAR_BUSCA = InStr(VAR_POSICION, TxtTexto.Text, "|", 1)
        VAR_MONTO3 = CDbl(Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION)))
        txt_importe_no_fiscal.Text = VAR_MONTO3
        
        VAR_POSICION = VAR_BUSCA + 1
        VAR_BUSCA = Len(TxtTexto.Text) + 1
        VAR_MONTO4 = CDbl(Mid(TxtTexto.Text, VAR_POSICION, (VAR_BUSCA - VAR_POSICION)))
        txt_descuentos.Text = VAR_MONTO4
        
        'SUBTOTAL
        VAR_SUBTOT = Round(CDbl(txt_total_bs.Text) - CDbl(Txt_Tasas.Text) - CDbl(Txt_tasa0.Text) - CDbl(txt_importe_no_fiscal.Text), 2)
        txtSubTotal.Text = VAR_SUBTOT
        
        Bs13 = Round(CDbl(VAR_BsCredFiscal) * 0.13, 2)
        Bs87 = Round(CDbl(VAR_BsCredFiscal) - CDbl(Bs13), 2)
        Dol87 = Round(Bs87 / GlTipoCambioOficial, 2)
        txt_13.Text = Bs13
        txt_87.Text = Bs87
    
    'Mid(Texto, Pos.inicial, nro.carac)
    'Instr(Pos.inicial, Texto, Carac.buscado, 1)
    FraQR.Visible = False
    FraGrabarCancelar.Visible = True
End Sub

Private Sub BtnGrabar3_Click()
    txt_nro_dui.Text = Cmbgestion.Text + dtc_codigo1.Text + Text4.Text + TxtDIM.Text
    FraDUI.Visible = False
End Sub

Private Sub BtnQR_Click()
    FraQR.Visible = True
    FraGrabarCancelar.Visible = False
    TxtTexto.SetFocus
    CmdCalcula.Visible = False
    ES_QR = "SI"
End Sub

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

Private Sub CmdCalcula_Click()
   If Txt_Tasas.Text = "" Then
        Txt_Tasas.Text = "0"
   End If
   If Txt_tasa0.Text = "" Then
        Txt_tasa0.Text = "0"
   End If
   If txt_importe_no_fiscal.Text = "" Then
        txt_importe_no_fiscal.Text = "0"
   End If
   If txt_descuentos.Text = "" Then
        txt_descuentos.Text = "0"
   End If
   If txt_total_bs.Text = "" Then
        txt_total_bs.Text = "0"
   End If
   If txt_tipo_cambio.Text = "" Or txt_tipo_cambio.Text = "0" Then
        txt_tipo_cambio.Text = GlTipoCambioOficial
   End If
   If opt_bs.Value = True Then          'BOB
        txt_total_bs.Enabled = True
        txt_total_dol.Enabled = False
        txt_total_eur.Enabled = False
        txt_total_dol.Text = Round(CDbl(txt_total_bs.Text) / CDbl(txt_tipo_cambio.Text), 2)
        txt_total_eur.Text = Round(CDbl(txt_total_bs.Text) / CDbl(txt_tipo_cambio.Text), 2)
   End If
   If opt_usd.Value = True Then          'USD
        txt_total_bs.Enabled = False
        txt_total_dol.Enabled = True
        txt_total_eur.Enabled = False
        txt_total_bs.Text = Round(CDbl(txt_total_dol.Text) * CDbl(txt_tipo_cambio.Text), 2)
        txt_total_eur.Text = Round(CDbl(txt_total_bs.Text) / CDbl(txt_tipo_cambio.Text), 2)
   End If
   If opt_eur.Value = True Then          'EUR
        txt_total_bs.Enabled = False
        txt_total_dol.Enabled = False
        txt_total_eur.Enabled = True
        txt_total_bs.Text = Round(CDbl(txt_total_eur.Text) * GlTipoCambioEuro, 2)
        txt_total_dol.Text = Round(CDbl(txt_total_bs.Text) / CDbl(txt_tipo_cambio.Text), 2)
   End If

   If Opt_DUI.Value = True Then
        txt_13.Text = CDbl(txt_total_bs.Text)
        VAR_SUBTOT = Round(CDbl(txt_total_bs.Text) / 0.13, 2)
        VAR_CREDFIS = Round(VAR_SUBTOT, 2)
        txt_87.Text = "0"
        txtSubTotal.Text = VAR_SUBTOT
        txt_CreditoFiscal.Text = VAR_CREDFIS
        txt_total_bs.Text = VAR_CREDFIS
        txt_total_dol.Text = Round(VAR_SUBTOT / CDbl(txt_tipo_cambio.Text), 2)
        txt_autorizacion.Text = "3"
   Else
        VAR_SUBTOT = Round(CDbl(txt_total_bs.Text) - CDbl(Txt_Tasas.Text) - CDbl(Txt_tasa0.Text) - CDbl(txt_importe_no_fiscal.Text), 2)
        VAR_CREDFIS = Round(VAR_SUBTOT - CDbl(txt_descuentos.Text), 2)
        Bs13 = Round(CDbl(VAR_CREDFIS) * 0.13, 2)
        Bs87 = Round(CDbl(VAR_CREDFIS) - CDbl(Bs13), 2)
        Dol87 = Round(Bs87 / GlTipoCambioOficial, 2)
        txtSubTotal.Text = VAR_SUBTOT
        txt_CreditoFiscal.Text = VAR_CREDFIS
        txt_13.Text = Bs13
        txt_87.Text = Bs87
        If opt_bs.Value = True Then
            txt_total_dol.Text = Round(VAR_SUBTOT / CDbl(txt_tipo_cambio.Text), 2)
            txt_total_eur.Text = Round(CDbl(VAR_SUBTOT) / GlTipoCambioEuro, 2)
        End If
   End If
   
'porcentaje_tot = 0
'If opt_usd.Value = True Then
'    If Text1.Text <> "" Then
'        porcentaje_tot = CDbl((txt_total_dol.Text * Text1.Text) / 100)
'        txt_total_dol.Text = Format(CDbl(txt_total_dol.Text + porcentaje_tot), "###,###,##0.00")
'    End If
'End If
'
'If opt_bs.Value = True Then
'    If Text1.Text <> "" Then
'        porcentaje_tot = CDbl((txt_total_bs.Text * Text1.Text) / 100)
'        txt_total_bs.Text = Format(CDbl(txt_total_bs.Text + porcentaje_tot), "###,###,##0.00")
'    End If
'End If
End Sub

Private Sub dtc_auto5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_auto5.BoundText
    dtc_aux4.BoundText = dtc_auto5.BoundText
    dtc_aux5.BoundText = dtc_auto5.BoundText
    dtc_desc5.BoundText = dtc_auto5.BoundText
    dtc_Nit5.BoundText = dtc_auto5.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux4.BoundText
    dtc_desc5.BoundText = dtc_aux4.BoundText
    dtc_aux5.BoundText = dtc_aux4.BoundText
    dtc_auto5.BoundText = dtc_aux4.BoundText
    dtc_Nit5.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux5.BoundText
    dtc_desc5.BoundText = dtc_aux5.BoundText
    dtc_aux4.BoundText = dtc_aux5.BoundText
    dtc_auto5.BoundText = dtc_aux5.BoundText
    dtc_Nit5.BoundText = dtc_aux5.BoundText
End Sub

Private Sub dtc_cod_alm_Click(Area As Integer)
    dtc_desc_alm.BoundText = dtc_cod_alm.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo5_Change()
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
    dtc_auto5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
    dtc_Nit5.BoundText = dtc_codigo5.BoundText
    dtc_auto5.BoundText = dtc_codigo5.BoundText
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

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
    dtc_auto5.BoundText = dtc_desc5.BoundText
    dtc_Nit5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc5_LostFocus()
    txt_autorizacion.Text = dtc_auto5.Text
End Sub

Private Sub dtc_Nit5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_Nit5.BoundText
    dtc_aux4.BoundText = dtc_Nit5.BoundText
    dtc_aux5.BoundText = dtc_Nit5.BoundText
    dtc_codigo5.BoundText = dtc_Nit5.BoundText
    dtc_auto5.BoundText = dtc_Nit5.BoundText
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
    'DOL = txt_total_dol.Text
    'BS = txt_total_bs.Text
    ES_QR = "NO"
    LblFactura.Caption = "Nro.Recibo/Cmpbte."
    CmdCalcula.Visible = True
    If parametro = "COMEX" Then
        opt_usd.Value = True
    End If
    FraDUI.Visible = False
End Sub

Private Sub Form_Load()
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_BENEF2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "SPAREDES"
        VAR_BENEF2 = "4828818"
        VAR_DA = "1.3"
    End If
    
    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    'Select Case Glaux
    rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5

'    Set rs_clasif6 = New ADODB.Recordset
'    If rs_clasif6.State = 1 Then rs_clasif6.Close
'    'Select Case Glaux
'    rs_clasif6.Open "SELECT * FROM ac_almacenes where beneficiario_codigo = " & fw_compras_comex.dtc_codigo11.Text & " ORDER BY almacen_descripcion ", db, adOpenStatic
'    Set Ado_clasif6.Recordset = rs_clasif6

    fw_adjudica_comex.Caption = "Adjudicación - " + fw_compras_comex.lbl_titulo
    If parametro <> "COMEX" Then
        Set rs_clasif6 = New ADODB.Recordset
        If rs_clasif6.State = 1 Then rs_clasif6.Close
        'Select Case Glaux
        rs_clasif6.Open "SELECT * FROM ac_almacenes where beneficiario_codigo = " & IIf(fw_compras_comex.dtc_codigo11.Text = "", "0", fw_compras_comex.dtc_codigo11.Text) & " ORDER BY almacen_descripcion ", db, adOpenStatic
        Set Ado_clasif6.Recordset = rs_clasif6
         'dtc_desc_alm.Enabled = True
        Text2.Visible = False
        Text1.Visible = False
         'Command1.Visible = False
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
    '     Text2.Visible = True
    '     Text1.Visible = True
    '     Command1.Visible = True
    '     lblLabels(0).Visible = True
    End If
        Call SeguridadSet(Me)
End Sub

Private Sub opt_bs_Click()
    txt_total_dol.Enabled = False
    If txt_total_dol.Text <= "0" Or txt_total_dol.Text = "" Then
        txt_total_dol.Text = "0"
    End If
    txt_total_eur.Enabled = False
    If txt_total_eur.Text <= "0" Or txt_total_eur.Text = "" Then
        txt_total_eur.Text = "0"
    End If
    txt_total_bs.Enabled = True
End Sub

Private Sub Opt_DUI_Click()
    Call opt_otro_Click
    FraDUI.Visible = True
    txt_autorizacion.Text = "3"
    Text4.Text = "C"
    'RECINTOS ADUANEROS
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    rs_clasif1.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo='APR' AND beneficiario_iniciales= 'ADUAN' ", db, adOpenStatic
    Set Ado_clasif1.Recordset = rs_clasif1
End Sub

Private Sub opt_eur_Click()
    txt_total_dol.Enabled = False
    If txt_total_dol.Text <= "0" Or txt_total_dol.Text = "" Then
        txt_total_dol.Text = "0"
    End If
    txt_total_bs.Enabled = False
    If txt_total_bs.Text <= "0" Or txt_total_bs.Text = "" Then
        txt_total_bs.Text = "0"
    End If
    txt_total_eur.Enabled = True
End Sub

Private Sub opt_gas_Click()
    FraDUI.Visible = False
End Sub

Private Sub opt_no_Click()
    ES_QR = "NO"
    BtnQR.Visible = False
    LblFechaFac.Visible = True
    LblFechaFac.Caption = "Fecha Recibo/Cmpbte."
    txtfecha_compra.Visible = True
    LblFactura.Caption = "Nro.Recibo/Cmpbte."
    lbl_campo2.Caption = "Total del Recibo en Bs.:"
    txt_Nota.Visible = True
    sino = MsgBox("Se Limpiarán los datos, desea continuar ? (Si ya registró, debe volver a hacerlo...)", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
        txt_Nota.Text = "0"
        Label22.Visible = False
        txt_nro_dui.Text = "0"
        txt_nro_dui.Visible = False
        Label2.Visible = False
        txt_autorizacion.Text = "0"
        txt_autorizacion.Visible = False
        Label3.Visible = False
        txt_cod_control.Text = ""
        txt_cod_control.Visible = False
        Lbl_NitCgi.Visible = False
        TxtNIT_CGI.Text = ""
        TxtNIT_CGI.Visible = False
        Label15.Visible = False
        Txt_Tasas.Text = "0"
        Txt_Tasas.Visible = False
        Label9.Visible = False
        Txt_tasa0 = "0"
        Txt_tasa0.Visible = False
        Label5.Visible = False
        txt_importe_no_fiscal.Text = "0"
        txt_importe_no_fiscal.Visible = False
        Label9.Visible = False
        txt_descuentos.Text = "0"
        txt_descuentos.Visible = False
        Label20.Visible = False
        txt_CreditoFiscal.Text = "0"
        txt_CreditoFiscal.Visible = False
        Label11.Visible = False
        txt_13.Text = "0"
        txt_13.Visible = False
        Label23.Visible = False
        txt_87.Text = "0"
        txt_87.Visible = False
    End If
End Sub

Private Sub opt_normal_Click()
    FraDUI.Visible = False
End Sub

Private Sub opt_otro_Click()
    BtnQR.Visible = True
    CmdCalcula.Visible = True
    LblFechaFac.Visible = True
    LblFechaFac.Caption = "Fecha Factura/DUI"
    txtfecha_compra.Visible = True
    LblFactura.Caption = "Nro. Factura"
    lbl_campo2.Caption = "Total de la Factura en Bs.:"
    sino = MsgBox("Se Limpiarán los datos de la Factura, si ya registró deberá volver a hacerlo. Desea continuar ? ...", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
'        txt_Nota.Text = "0"
'        Label22.Visible = False
'        txt_nro_dui.Text = "0"
'        txt_nro_dui.Visible = False
'        Label2.Visible = False
'        txt_autorizacion.Text = "0"
'        txt_autorizacion.Visible = False
'        Label3.Visible = False
'        txt_cod_control.Text = ""
'        txt_cod_control.Visible = False
'        Lbl_NitCgi.Visible = False
'        TxtNIT_CGI.Text = ""
'        TxtNIT_CGI.Visible = False
'        Label15.Visible = False
'        Txt_Tasas.Text = "0"
'        Txt_Tasas.Visible = False
'        Label9.Visible = False
'        Txt_tasa0 = "0"
'        Txt_tasa0.Visible = False
'        Label5.Visible = False
'        txt_importe_no_fiscal.Text = "0"
'        txt_importe_no_fiscal.Visible = False
'        Label9.Visible = False
'        txt_descuentos.Text = "0"
'        txt_descuentos.Visible = False
'        Label20.Visible = False
'        txt_CreditoFiscal.Text = "0"
'        txt_CreditoFiscal.Visible = False
'        Label11.Visible = False
'        txt_13.Text = "0"
'        txt_13.Visible = False
'        Label23.Visible = False
'        txt_87.Text = "0"
'        txt_87.Visible = False
    
        txt_Nota.Visible = True
        txt_Nota.Text = "0"
        Label22.Visible = True
        txt_nro_dui.Text = "0"
        txt_nro_dui.Visible = True
        Label2.Visible = True
        txt_autorizacion.Text = "0"
        txt_autorizacion.Visible = True
        Label3.Visible = True
        txt_cod_control.Text = ""
        txt_cod_control.Visible = True
        Lbl_NitCgi.Visible = True
        TxtNIT_CGI.Text = ""
        TxtNIT_CGI.Visible = True
        Label15.Visible = True
        Txt_Tasas.Text = "0"
        Txt_Tasas.Visible = True
        Label9.Visible = True
        Txt_tasa0 = "0"
        Txt_tasa0.Visible = True
        Label5.Visible = True
        txt_importe_no_fiscal.Text = "0"
        txt_importe_no_fiscal.Visible = True
        Label9.Visible = True
        txt_descuentos.Text = "0"
        txt_descuentos.Visible = True
        Label20.Visible = True
        txt_CreditoFiscal.Text = "0"
        txt_CreditoFiscal.Visible = True
        Label11.Visible = True
        txt_13.Text = "0"
        txt_13.Visible = True
        Label23.Visible = True
        txt_87.Text = "0"
        txt_87.Visible = True
    End If
End Sub

Private Sub opt_si_Click()
    BtnQR.Visible = True
    CmdCalcula.Visible = True
    LblFechaFac.Visible = True
    LblFechaFac.Caption = "Fecha Factura/DUI"
    txtfecha_compra.Visible = True
    LblFactura.Caption = "Nro. Factura"
    lbl_campo2.Caption = "Total de la Factura en Bs.:"
    sino = MsgBox("Se Limpiarán los datos de la Factura, si ya registró deberá volver a hacerlo. Desea continuar ? ...", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
        txt_Nota.Visible = True
        txt_Nota.Text = "0"
        Label22.Visible = True
        txt_nro_dui.Text = "0"
        txt_nro_dui.Visible = True
        Label2.Visible = True
        txt_autorizacion.Text = "0"
        txt_autorizacion.Visible = True
        Label3.Visible = True
        txt_cod_control.Text = ""
        txt_cod_control.Visible = True
        Lbl_NitCgi.Visible = True
        TxtNIT_CGI.Text = ""
        TxtNIT_CGI.Visible = True
        Label15.Visible = True
        Txt_Tasas.Text = "0"
        Txt_Tasas.Visible = True
        Label9.Visible = True
        Txt_tasa0 = "0"
        Txt_tasa0.Visible = True
        Label5.Visible = True
        txt_importe_no_fiscal.Text = "0"
        txt_importe_no_fiscal.Visible = True
        Label9.Visible = True
        txt_descuentos.Text = "0"
        txt_descuentos.Visible = True
        Label20.Visible = True
        txt_CreditoFiscal.Text = "0"
        txt_CreditoFiscal.Visible = True
        Label11.Visible = True
        txt_13.Text = "0"
        txt_13.Visible = True
        Label23.Visible = True
        txt_87.Text = "0"
        txt_87.Visible = True
    End If
End Sub

Private Sub opt_usd_Click()
    txt_total_bs.Enabled = False
    If txt_total_bs.Text <= "0" Or txt_total_bs.Text = "" Then
        txt_total_bs.Text = "0"
    End If
    txt_total_eur.Enabled = False
    If txt_total_eur.Text <= "0" Or txt_total_eur.Text = "" Then
        txt_total_eur.Text = "0"
    End If
    txt_total_dol.Enabled = True
End Sub

Private Sub BtnCancelar1_Click()
    fra_provedor.Visible = False
    Frame1.Enabled = True
    FraGrabarCancelar.Visible = True
End Sub

Private Sub BtnGrabar1_Click()
On Error GoTo UpdateErr
    db.Execute "INSERT INTO gc_beneficiario (beneficiario_codigo,      tipoben_codigo, tipodoc_codigo, beneficiario_nit,            beneficiario_denominacion,          comun_codigo,            estado_codigo, fecha_registro, usr_codigo)" & _
               "VALUES ('" & txt_nit_new.Text & "', '22',      '" & "NIT" & "', '" & txt_nit_new.Text & "', '" & txt_denominacion_new.Text & "', '" & TxtAutorizacionNew.Text & "', 'APR',     '" & Date & "', '" & glusuario & "')"

    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    'Select Case Glaux
    rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo='APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
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

Private Sub BtnCancelar2_Click()
    FraQR.Visible = False
    FraGrabarCancelar.Visible = True
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

'Private Sub txt_importe_no_fiscal_Change()
'    txt_13.Text = (CDbl(IIf(txt_total_bs.Text = "", 0, txt_total_bs.Text)) - CDbl(IIf(txt_importe_no_fiscal.Text = "", 0, txt_importe_no_fiscal.Text))) * 0.13
'End Sub

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
' On Error GoTo UpdateErr
' If opt_bs.Value = True Then
'    If txt_total_bs.Text <> "" And txt_total_bs.Text <> "," Then
'        If txt_tipo_cambio.Text <> "" And txt_tipo_cambio.Text <> "," Then
'            txt_total_dol.Text = Format(CDbl(txt_total_bs.Text) / CDbl(txt_tipo_cambio.Text), "###,###,##0.00")
'        Else
'          txt_total_dol.Text = "0"
'        End If
'    Else
'          txt_total_dol.Text = "0"
'    End If
' End If
'
''If txt_total_bs.Text = "" Then
''txt_total_bs.Text = 0
''End If
''
''If txt_importe_no_fiscal.Text = "" Then
''txt_importe_no_fiscal.Text = 0
''End If
''
''txt_13.Text = (CDbl(txt_total_bs.Text) - CDbl(txt_importe_no_fiscal.Text)) * 0.13
' If opt_usd.Value = True Then
''txt_total_dol.Text = CDbl(txt_total.Text / txt_tipo_cambio)
'
' End If
'
' If txt_total_bs > "0" And txt_total_dol > "0" Then
'    Label21.Visible = True
'    cmb_mes_ini.Visible = True
'    Label12.Visible = True
'    txtCantCuota.Visible = True
'    Label18.Visible = True
'    cmd_unimed2.Visible = True
' Else
'    Label21.Visible = False
'    cmb_mes_ini.Visible = False
'    Label12.Visible = False
'    txtCantCuota.Visible = False
'    Label18.Visible = False
'    cmd_unimed2.Visible = False
' End If
' If Opt_DUI.Value <> True Then
'    txt_13.Text = (CDbl(IIf(txt_total_bs.Text = "", 0, txt_total_bs.Text)) - CDbl(IIf(txt_importe_no_fiscal.Text = "", 0, txt_importe_no_fiscal.Text))) * 0.13
' End If
'Exit Sub
'UpdateErr:
'  MsgBox Err.Description
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
'On Error GoTo UpdateErr
'''porcentaje_tot = IIf(txt_total_dol.Text = 0 Or txt_total_dol.Text = "", 0, txt_total_dol.Text)
''If opt_usd.Value = True Then
'''txt_total_bs.Text = CDbl(txt_total_dol.Text * txt_tipo_cambio)
''End If
''If txt_total_bs.Text <> "0" Then
''txt_13.Text = (CDbl(IIf(txt_total_bs.Text = "", 0, txt_total_bs.Text)) - CDbl(IIf(txt_importe_no_fiscal.Text = "", 0, txt_importe_no_fiscal.Text))) * 0.13
''End If
'If opt_usd.Value = True Then
'    If txt_total_dol.Text <> "" And txt_total_dol.Text <> "," Then
'        If txt_tipo_cambio.Text <> "" And txt_tipo_cambio.Text <> "," Then
'            txt_total_bs.Text = Format(CDbl(txt_total_dol.Text) * CDbl(txt_tipo_cambio.Text), "###,###,##0.00")
'        Else
'            txt_total_bs.Text = "0"
'        End If
'    Else
'          txt_total_bs.Text = "0"
'    End If
'End If
'Exit Sub
'UpdateErr:
'  MsgBox Err.Description
End Sub

Private Sub txt_total_dol_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
    Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub txtfecha_compra_LostFocus()
    cmb_mes_ini.Text = UCase(MonthName(Month(txtfecha_compra.Value)))
    txt_mes.Text = Month(txtfecha_compra.Value)
    txtFecha.Value = txtfecha_compra.Value
    txtFecha2.Value = txtfecha_compra.Value
    txtFecha3.Value = txtfecha_compra.Value
End Sub

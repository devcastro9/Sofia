VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCuentaBancaria 
   Caption         =   "Cuenta Bancaria"
   ClientHeight    =   8385
   ClientLeft      =   270
   ClientTop       =   135
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "MOVIMIENTO DE PAGOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   4335
         TabIndex        =   5
         Top             =   150
         Width           =   3855
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   4
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   2
         Top             =   690
         Width           =   2460
      End
      Begin VB.Label Label8 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   75
         TabIndex        =   1
         Top             =   675
         Width           =   1110
      End
   End
   Begin Crystal.CrystalReport CryMov 
      Left            =   225
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame FraTotal 
      Caption         =   "TOTAL HASTA LA FECHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1485
      Left            =   7395
      TabIndex        =   71
      Top             =   5910
      Visible         =   0   'False
      Width           =   4500
      Begin VB.TextBox TxtSaldoInicial 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   75
         Top             =   480
         Width           =   1890
      End
      Begin VB.TextBox TxtSaldoActualBol 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   74
         Top             =   480
         Width           =   1860
      End
      Begin VB.TextBox TxtSaldoActualDol 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   73
         Top             =   1020
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox TxtMovimiento 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   72
         Top             =   1020
         Width           =   1890
      End
      Begin VB.Label Label11 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Left            =   105
         TabIndex        =   79
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label12 
         Caption         =   "Saldo Actual en Bolivianos"
         Height          =   210
         Left            =   2100
         TabIndex        =   78
         Top             =   270
         Width           =   2070
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Actual en Dolares"
         Height          =   210
         Left            =   2130
         TabIndex        =   77
         Top             =   780
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label Label13 
         Caption         =   "Movimiento"
         Height          =   210
         Left            =   135
         TabIndex        =   76
         Top             =   780
         Width           =   1005
      End
   End
   Begin VB.Frame FraFecha 
      Caption         =   "Impresi�n Por "
      Height          =   825
      Left            =   7380
      TabIndex        =   52
      Top             =   1035
      Width           =   2190
      Begin VB.CheckBox ChkDanida 
         Caption         =   "ORGANISMO 999"
         Height          =   405
         Left            =   2220
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton OptFechaPago 
         Caption         =   "Fecha de Pago"
         Height          =   300
         Left            =   165
         TabIndex        =   54
         Top             =   165
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptFechaImpresion 
         Caption         =   "Fecha de Impresi�n"
         Height          =   315
         Left            =   180
         TabIndex        =   53
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame FraMesFecha 
      Height          =   750
      Left            =   7395
      TabIndex        =   66
      Top             =   1815
      Width           =   2190
      Begin VB.OptionButton OptFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   225
         TabIndex        =   69
         Top             =   435
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha"
         CausesValidation=   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   68
         Top             =   -240
         Width           =   1935
      End
      Begin VB.OptionButton OptMes 
         Caption         =   "Mes"
         Height          =   330
         Left            =   210
         TabIndex        =   67
         Top             =   135
         Width           =   1110
      End
   End
   Begin VB.Frame FraOpcionesCuenta 
      Height          =   1530
      Left            =   9600
      TabIndex        =   61
      Top             =   1035
      Width           =   2400
      Begin VB.OptionButton OptTodasCuentas 
         Caption         =   "X Todas las Cuentas"
         Height          =   315
         Left            =   105
         TabIndex        =   63
         Top             =   615
         Width           =   2010
      End
      Begin VB.OptionButton OptUnaCuenta 
         Caption         =   "Por Cuenta"
         Height          =   330
         Left            =   105
         TabIndex        =   62
         Top             =   210
         Width           =   1830
      End
   End
   Begin VB.Frame FraFechas 
      Height          =   945
      Left            =   7395
      TabIndex        =   55
      Top             =   2520
      Visible         =   0   'False
      Width           =   4620
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   375
         Left            =   90
         TabIndex        =   56
         Top             =   420
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24707073
         CurrentDate     =   36413
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   375
         Left            =   2220
         TabIndex        =   57
         Top             =   435
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24707073
         CurrentDate     =   36413
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio"
         Height          =   240
         Left            =   105
         TabIndex        =   59
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin"
         Height          =   240
         Left            =   2235
         TabIndex        =   58
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.Frame FraCuenta 
      Caption         =   "Cuenta"
      Height          =   1320
      Left            =   7380
      TabIndex        =   48
      Top             =   3450
      Visible         =   0   'False
      Width           =   4620
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "FrmCuentaBancaria.frx":0000
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   105
         TabIndex        =   49
         Top             =   225
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDescripcion 
         Bindings        =   "FrmCuentaBancaria.frx":0018
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   105
         TabIndex        =   50
         Top             =   930
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCTgn 
         Bindings        =   "FrmCuentaBancaria.frx":0030
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   105
         TabIndex        =   51
         Top             =   585
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
   End
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   390
      Left            =   1320
      Top             =   7065
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   688
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
      Caption         =   "Cuenta Bancaria"
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
   Begin VB.Frame FraMes 
      Height          =   1095
      Left            =   7395
      TabIndex        =   12
      Top             =   4740
      Visible         =   0   'False
      Width           =   4635
      Begin VB.ComboBox CmbMes 
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   105
         TabIndex        =   64
         Top             =   420
         Width           =   3810
      End
      Begin VB.ComboBox CmbAnio 
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   2085
         TabIndex        =   17
         Text            =   "2000"
         Top             =   405
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen1 
         Bindings        =   "FrmCuentaBancaria.frx":0048
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   90
         TabIndex        =   19
         Top             =   3810
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "FrmCuentaBancaria.frx":0060
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   90
         TabIndex        =   20
         Top             =   4515
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "FrmCuentaBancaria.frx":0078
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   75
         TabIndex        =   21
         Top             =   4170
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Label LblMes 
         Caption         =   "Mes"
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   105
         TabIndex        =   65
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "A�o"
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   2085
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin MSDataGridLib.DataGrid DtGCuentaBancaria 
      Height          =   5895
      Left            =   1335
      TabIndex        =   29
      Top             =   1095
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.Animation AVI 
      Height          =   1140
      Left            =   10875
      TabIndex        =   42
      Top             =   7515
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2011
      _Version        =   393216
      FullWidth       =   64
      FullHeight      =   76
   End
   Begin MSDataGridLib.DataGrid DtGPagosDetalle 
      Height          =   5880
      Left            =   1350
      TabIndex        =   16
      Top             =   1095
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   10372
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PAGOS POR MES"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DtGCuenta 
      Height          =   5850
      Left            =   1410
      TabIndex        =   13
      Top             =   1095
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   10319
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CUENTA BANCARIA"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cta_codigo"
         Caption         =   "CODIGO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cta_descripcion_larga"
         Caption         =   "DESCRIPCION "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cta_codigo_tgn"
         Caption         =   "TGN CODIGO "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraPorCuenta 
      Caption         =   "POR CUENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1530
      Left            =   7395
      TabIndex        =   22
      Top             =   5895
      Visible         =   0   'False
      Width           =   4665
      Begin VB.TextBox TxtMovimientoCuenta 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   27
         Top             =   1050
         Width           =   1890
      End
      Begin VB.TextBox TxtSICuenta 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   24
         Top             =   480
         Width           =   1890
      End
      Begin VB.TextBox TxtSACuenta 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2430
         TabIndex        =   23
         Top             =   480
         Width           =   1860
      End
      Begin VB.Label Label14 
         Caption         =   "Movimiento"
         Height          =   210
         Left            =   135
         TabIndex        =   28
         Top             =   825
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Left            =   105
         TabIndex        =   26
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label15 
         Caption         =   "Saldo Actual en Bolivianos"
         Height          =   210
         Left            =   2445
         TabIndex        =   25
         Top             =   270
         Width           =   1890
      End
   End
   Begin VB.Frame FraTodasCuentas 
      Caption         =   "TODAS LAS CUENTAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1485
      Left            =   7380
      TabIndex        =   31
      Top             =   5880
      Visible         =   0   'False
      Width           =   4485
      Begin VB.TextBox TxtSaldoInicialTotal 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   35
         Top             =   480
         Width           =   1890
      End
      Begin VB.TextBox TxtSaldoTodasCuentas 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   34
         Top             =   480
         Width           =   1860
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   33
         Top             =   1020
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox TxtMovimientoTodos 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   32
         Top             =   1020
         Width           =   1890
      End
      Begin VB.Label Label20 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Left            =   105
         TabIndex        =   39
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label19 
         Caption         =   "Saldo Actual en Bolivianos"
         Height          =   210
         Left            =   2100
         TabIndex        =   38
         Top             =   270
         Width           =   2070
      End
      Begin VB.Label Label18 
         Caption         =   "Saldo Actual en Dolares"
         Height          =   210
         Left            =   2130
         TabIndex        =   37
         Top             =   780
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label Label17 
         Caption         =   "Movimiento"
         Height          =   210
         Left            =   135
         TabIndex        =   36
         Top             =   780
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6435
      Left            =   45
      TabIndex        =   43
      Top             =   1035
      Width           =   1245
      Begin VB.CommandButton CmdTodasCtas 
         Caption         =   "Todas las Cuentas  Bancarias"
         Height          =   1035
         Left            =   120
         Picture         =   "FrmCuentaBancaria.frx":0090
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   945
         Width           =   930
      End
      Begin VB.CommandButton CmdEjecutar 
         Caption         =   "Ejecutar"
         Height          =   720
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "FrmCuentaBancaria.frx":06FA
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1980
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimirMovimiento 
         Caption         =   "Imprimir "
         Height          =   705
         Left            =   120
         Picture         =   "FrmCuentaBancaria.frx":109C
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   225
         Width           =   930
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Restaurar"
         Height          =   705
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "FrmCuentaBancaria.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2700
         Width           =   930
      End
      Begin VB.CommandButton CmdPorCuenta 
         Caption         =   "Salir"
         Height          =   705
         Left            =   105
         TabIndex        =   44
         Top             =   3390
         Width           =   945
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   6435
      Left            =   45
      TabIndex        =   6
      Top             =   1035
      Width           =   1260
      Begin VB.CommandButton CmdTesoreria 
         Caption         =   "Tesoreria Actualiza"
         Height          =   690
         Left            =   180
         TabIndex        =   41
         Top             =   5610
         Width           =   930
      End
      Begin VB.CommandButton CmdTodosRegistros 
         Caption         =   "Todos x Fecha"
         Height          =   705
         Left            =   180
         TabIndex        =   40
         Top             =   4110
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimeGrid 
         Caption         =   "Imprime grid"
         Height          =   735
         Left            =   180
         TabIndex        =   30
         Top             =   2340
         Width           =   930
      End
      Begin VB.CommandButton Cmdanio 
         Caption         =   "Anual"
         Height          =   705
         Left            =   180
         TabIndex        =   11
         Top             =   1635
         Width           =   930
      End
      Begin VB.CommandButton CmdRangoFecha 
         Caption         =   "Movimiento por fecha"
         Height          =   735
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   930
      End
      Begin VB.CommandButton CmdmovimientoMes 
         Caption         =   "Movimiento por mes"
         Height          =   705
         Left            =   180
         TabIndex        =   9
         Top             =   195
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   180
         Picture         =   "FrmCuentaBancaria.frx":20A8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4815
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Todas las Cuentas  Bancarias"
         Height          =   1035
         Left            =   180
         Picture         =   "FrmCuentaBancaria.frx":24EA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3075
         Width           =   930
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha Inicio"
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1590
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha Fin"
      Height          =   240
      Left            =   30
      TabIndex        =   14
      Top             =   825
      Width           =   1590
   End
End
Attribute VB_Name = "FrmCuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  SAF-2000
' M�dulo:                   Movimiento de Cuenta Bancaria
' Base de Datos:            SQL SERVER 7.0 (espa�ol)
' Formulario :              FrmCuentaBancaria.frm
' Descipci�n :              Movimientos de cuentas bancaria por mes,
'                           por fecha, por a�o, por cuenta, por todas
'                           las cuentas, etc.
' Formularios relacionados: Main.frm (Padre)
'                           CryPagos, CryCtaBancaria
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creaci�n         01/Mare/ 2000
' Fecha �ltima modificaci�n 20/May/ 2000
' Versi�n:                  2.0
'========================================================================================

Dim rsCuenta As New ADODB.Recordset
Dim rsPAgoDetalle As New ADODB.Recordset
Dim rsCB As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim Monto_Actual_Bolivianos As Long
Dim Monto_Actual_Dolares As Long
Dim NoRegistros As Long
'Para actualizar saldos
Dim rsCBria As New ADODB.Recordset
Dim rsPg As New ADODB.Recordset

'Para fines de impresi�n en mdulo se definen
'Dim swMes As Integer
'Dim swFecha As Integer


Private Sub Cmdanio_Click()

'*******
db.Execute "DELETE FROM to_movimiento"

    If DtCCuentaOrigen <> "" Then
        MsgBox "Calculo por cuenta"
        Proceso_Cuenta "GESTION"
        Exit Sub
    End If
    
    Monto_Actual_Bolivianos = 0
    Monto_Actual_Dolares = 0

If OptUnaCuenta.Value = True Then
    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
''    rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion='" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                         " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPAgoDetalle.RecordCount > 0 Then
    While Not rsPAgoDetalle.EOF
                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', '" & rsPAgoDetalle!monto_bolivianos & "','" & rsPAgoDetalle!codigo_pago & "','" & rsPAgoDetalle!monto_Dolares & "','" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                rsPAgoDetalle.MoveNext
    Wend
        TxtSaldoInicial.Text = Monto_Actual_Bolivianos
        'TxtSaldoActual.Text = Monto_Actual_Bolivianos
    End If
    
    If rsPAgoDetalle.RecordCount > 0 Then
            Set DtGPagosDetalle.DataSource = rsPAgoDetalle
            DtGCuenta.Refresh
    End If
End If

If OptTodasCuentas.Value = True Then
    Set rsCB = New ADODB.Recordset
    If rsCB.State = 1 Then rsCB.Close
    rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    If rsCB.RecordCount > 0 Then
       While Not rsCB.EOF
            If rsCB.RecordCount > 0 Then
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion='" & CmbAnio.Text & "' and cta_codigo='" & rsCB("Cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                While Not rsPAgoDetalle.EOF
                            If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                            If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                            db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                            rsPAgoDetalle.MoveNext
                Wend
                    TxtSaldoInicial.Text = Monto_Actual_Bolivianos
                    'TxtSaldoActual.Text = Monto_Actual_Bolivianos
                End If
                rsCB.MoveNext
         End If
        Wend
        resfresca_grid_cuenta_bancaria
    End If
End If
End Sub

Private Sub CmdBuscar_Click()
    Set rsC = New ADODB.Recordset
    If rsC.State = 1 Then rsCuenta.Close
    rsC.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                  "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    If rsC.RecordCount > 0 Then
        NoRegistros = rsC.RecordCount
        AdoCuenta.Caption = NoRegistros
        Set DtGCuentaBancaria.DataSource = rsC
        Set AdoCuenta.Recordset = rsC
    End If
End Sub

Private Sub CmdEjecutar_Click()
Dim c As Long
Dim vectorBs(100) As Double
Dim vectorSus(100) As Double

Dim dia As Variant
Dim mes As Variant
Dim anio As Variant
Dim mes_numeral As Integer
Dim condicion As String

'MOVIMIENTO DE PAGOS POR MES
If OptFechaImpresion.Value = False And OptFechaPago.Value = False Then
        MsgBox "Elija si el reporte sera por fecha de pago o de impresion", vbInformation + vbCritical, "Validaci�n de datos"
        Exit Sub
End If

If OptMes.Value = True Then
                
                If OptFechaPago.Value = falso And OptFechaImpresion.Value = falso And OptMes.Value = False Then
                    MsgBox "Elija una opci�n por fecha o mes", vbInformation + vbCritical, "Vaidaci�n de datos"
                    Exit Sub
                End If
                
                    swMes = 1
                    swFecha = 0
                
                If DtCCuentaOrigen.Text = "" And OptUnaCuenta.Value = True Then
                    MsgBox "Lista por cuenta bancatia elija cuenta bancaria", vbInformation + vbCritical
                    AVI.Stop
                    Exit Sub
                End If
                '*********
                db.Execute "DELETE FROM to_movimiento"
                db.Execute "DELETE FROM to_cta_Bancaria"
                ''db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                If DtCCuentaOrigen <> "" Then
                    MsgBox "Calculo por cuenta"
                    Proceso_Cuenta "MES"
                    AVI.Stop
                    Exit Sub
                End If
                
                TxtSaldoActualBol.Text = ""
                'Detalle de cuenta bancaria por mes
                    CmbMes.Enabled = True
                    LblMes.ForeColor = &H80&
                    
                If CmbMes.Text = "" Then
                    MsgBox "No existe dato de mes....", vbCritical + vbInformation, "Validaci�n de datos"
                    AVI.Stop
                    Exit Sub
                End If
                            Select Case CmbMes.Text
                                Case "ENERO"
                                    mes_numeral = 1
                                Case "FEBRERO"
                                    mes_numeral = 2
                                Case "MARZO"
                                    mes_numeral = 3
                                Case "ABRIL"
                                    mes_numeral = 4
                                Case "MAYO"
                                    mes_numeral = 5
                                Case "JUNIO"
                                    mes_numeral = 6
                                Case "JULIO"
                                    mes_numeral = 7
                                Case "AGOSTO"
                                    mes_numeral = 8
                                Case "SEPTIEMBRE"
                                    mes_numeral = 9
                                Case "OCTUBRE"
                                    mes_numeral = 10
                                Case "NOVIEMBRE"
                                    mes_numeral = 11
                                Case "DICIEMBRE"
                                    mes_numeral = 12
                            End Select
                    
                    Monto_Actual_Bolivianos = 0
                    Monto_Actual_Dolares = 0
                    
                    If CmbAnio.Text = "" Then
                        MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
                        Exit Sub
                    End If
                
                If OptTodasCuentas.Value = True Then
                            'Determinando si es por fecha de pago o fecha de impresion
                            If OptFechaPago.Value = True Then
                                condicion = "month(pago_detalle.fecha_pago)='" & mes_numeral & "'"
                            End If
                            If OptFechaImpresion.Value = True Then
                                condicion = "month(pago_detalle.fecha_impresion_cheque)='" & mes_numeral & "'"
                            End If
                            'month( & condicion &)='" & mes_numeral & "'
                            If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                             rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                                                " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE   " & condicion & " and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
                                If rsPAgoDetalle.RecordCount > 0 Then
                                    Set DtGPagosDetalle.DataSource = rsPAgoDetalle
                                    DtGCuenta.Refresh
                                    While Not rsPAgoDetalle.EOF
                                        If Not IsNull(rsPAgoDetalle("monto_dolares")) And Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then
                                            Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                            Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                            db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                            '********Buscar la cuenta para acumular
                                           Select Case rsPAgoDetalle("cta_codigo")
                                               Case "0869"
                                                     '"4.41.1.1.1.402.208.11-2"
                                                     vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                                               Case "0870"
                                                    '"4.41.1.1.1.402.208.12-1"
                                                     vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                                               Case "0872"
                                                    '"4.41.1.1.1.402.208.14-0"
                                                     vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                                               Case "0873"
                                                    '"4.41.1.1.1.402.208.16-8"
                                                     vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                                               Case "2676"
                                                    '"4.41.1.1.1.402.208.18-6"
                                                     vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                                               Case "0922"
                                                     '"4.41.1.1.1.402.254.01-7"
                                                     vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                                               Case "0921"
                                                     '"4.41.1.1.1.402.254.02-6"
                                                     vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297792"
                                                     vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297809"
                                                     vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297841"
                                                     vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297867"
                                                     vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297875"
                                                     vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297883"
                                                     vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297891"
                                                     vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297916"
                                                     vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297924"
                                                     vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297932"
                                                     vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297940"
                                                     vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297958"
                                                     vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                                               Case "1-301973"
                                                     vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                                               Case "1-301999"
                                                     vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                                               Case "1-302731"
                                                     vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                                               Case "1-303515"
                                                     vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                                               Case "1-306379"
                                                     vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                                               Case "1-302731"
                                                     vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                                            End Select
                                            '**************
                                        ' rsPagoDetalle.MoveNext
                                            'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_Bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!Cta_descripcion_larga & "','" & rsPagoDetalle!Cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                                        End If
                                        rsPAgoDetalle.MoveNext
                                    Wend
                                Else
                                    Set DtGPagosDetalle.DataSource = rsNada
                                    MsgBox "No existen registros", vbInformation
                                End If
                                'Asignacion de acumulado
                                TxtMovimientoTodos.Text = Monto_Actual_Bolivianos
                                TxtSaldoTodasCuentas.Text = Val(TxtSaldoInicialTotal) - Val(Monto_Actual_Bolivianos)
                                
                                
                                'Reiniciar para la siguiente sumatoria
                                Monto_Actual_Bolivianos = 0
                                Monto_Actual_Dolares = 0
                                resfresca_grid_cuenta_bancaria
                            
                    
                    '*****  Insertando datos a to_cta_bancaria
                       'Abriendo tabla de cuenta bancaria para las cuentas
                       Set rsCuentaImprimir = New ADODB.Recordset
                       If rsCuentaImprimir.State = 1 Then rsCuentaImprimir.Close
                       rsCuentaImprimir.Open "SELECT * FROM to_cta_bancaria", db, adOpenKeyset, adLockOptimistic
                            Set rsCB = New ADODB.Recordset
                            If rsCB.State = 1 Then rsCB.Close
                            rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
                            While Not rsCB.EOF
                                 rsCuentaImprimir.AddNew
                                 rsCuentaImprimir("ges_gestion") = rsCB("ges_gestion")
                                 rsCuentaImprimir("bco_codigo") = rsCB("bco_codigo")
                                 rsCuentaImprimir("cta_codigo") = rsCB("cta_codigo")
                                 rsCuentaImprimir("cta_codigo_tgn") = rsCB("cta_codigo_tgn")
                                 rsCuentaImprimir("cta_descripcion_larga") = rsCB("cta_descripcion_larga")
                                 rsCuentaImprimir("cta_saldo_inicial") = rsCB("cta_saldo_inicial")
                             
                                           Select Case rsCB("cta_codigo")
                                               Case "0869"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(1)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(1)
                                               Case "0870"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(2)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(2)
                                               Case "0872"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(3)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(3)
                                               Case "0873"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(4)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(4)
                                               Case "2676"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(5)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(5)
                                               Case "0922"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(6)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(6)
                                               Case "0921"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(7)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(7)
                                               Case "1-297792"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(8)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(8)
                                               Case "1-297809"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(9)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(9)
                                               Case "1-297841"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(10)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(10)
                                               Case "1-297867"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(11)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(11)
                                               Case "1-297875"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(12)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(12)
                                               Case "1-297883"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(13)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(13)
                                               Case "1-297891"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(14)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(14)
                                               Case "1-297916"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(15)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(15)
                                               Case "1-297924"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(16)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(16)
                                               Case "1-297932"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(17)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(17)
                                               Case "1-297940"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(18)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(18)
                                               Case "1-297958"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(19)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(19)
                                               Case "1-301973"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(20)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(20)
                                               Case "1-301999"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(21)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(21)
                                               Case "1-302731"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(22)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(22)
                                               Case "1-303515"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(23)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(23)
                                               Case "1-306379"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(24)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(24)
                                               Case "1-302731"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(25)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(25)
                                            End Select
                                 rsCuentaImprimir.Update
                                rsCB.MoveNext
                            Wend
                    MsgBox "Ya termin�"
                End If
                AVI.Stop

End If

'MOVIMIENTO POR FECHA
If OptFecha.Value = True Then
            
            AVI.Open "c:\AVIS\search.avi"
            AVI.Play
            
            'Validaci�n de fecha
            If DTPFechaInicio.Value > DTPFechaFin.Value Or DTPFechaFin.Value < DTPFechaInicio.Value Then
                 MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
                 AVI.Stop
                 Exit Sub
            End If
            
            'Para fines de impresi�n
            swMes = 0
            swFecha = 1
            
            
            '*********
            db.Execute "DELETE FROM to_movimiento"
            
            If DtCCuentaOrigen.Text = "" And OptUnaCuenta.Value = True Then
                MsgBox "Lista por cuenta elija cuenta bancaria", vbInformation + vbCritical
                Exit Sub
            End If
                If DtCCuentaOrigen <> "" Then
                    MsgBox "Calculo por cuenta"
                    Proceso_Cuenta "RANGO"
                    Exit Sub
                End If
                
                Monto_Actual_Bolivianos = 0
                Monto_Actual_Dolares = 0
                TxtSaldoActualBol.Text = ""
                If CmbAnio.Text = "" Then
                    MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
                    Exit Sub
                End If
                
            If OptUnaCuenta.Value = True Then
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                
                rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                                   " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                     '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' ", db, adOpenKeyset, adLockOptimistic
                ''rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
                ''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                ''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                    While Not rsPAgoDetalle.EOF
                        If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                                
                        End If
                        rsPAgoDetalle.MoveNext
                        c = c + 1
                    Wend
                Else
                    MsgBox "No existen registros", vbInformation
                    Set DtGPagosDetalle.DataSource = rsNada
                End If
                
                
                    'Filtrando los datos
                    If c = 0 Then
                        MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
                    Else
                        Set rsFecha = New ADODB.Recordset
                        If rsFecha.State = 1 Then rsFecha.Close
                        rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
                        If rsFecha.RecordCount > 0 Then
                            Set DtGPagosDetalle.DataSource = rsFecha
                            Set AdoCuenta.Recordset = rsFecha
                            NoRegistros = c
                            AdoCuenta.Caption = NoRegistros
                           Exit Sub
                        End If
                    End If
            
                TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
                TxtMovimiento.Text = Monto_Actual_Bolivianos
                TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - TxtMovimiento.Text
                If rsPAgoDetalle.RecordCount > 0 Then
                NoRegistros = rsPAgoDetalle.RecordCount
                        AdoCuenta.Caption = NoRegistros
                        Set DtGPagosDetalle.DataSource = rsPAgoDetalle
                        DtGCuenta.Refresh
                End If
            End If
            
            
            If OptTodasCuentas.Value = True Then
                          If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                          rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_impresion_cheque" & _
                                             " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                            ''***                   rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.codigo_pago, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_registro " & _
                                            ''***                   "FROM (pago_detalle RIGHT JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo ORDER BY pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                                    '"Where (((pago_detalle.fecha_pago) = #5/11/2000#))"
                                                    '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                                                    'rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "' and cta_codigo='" & rsCB("Cta_codigo") & "' ORDER BY fecha_pago", db, adOpenKeyset, adLockOptimistic
                            If rsPAgoDetalle.RecordCount > 0 Then
                                While Not rsPAgoDetalle.EOF
                                    If OptFechaPago.Value = True Then 'Fecha de pago
                                        If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                                c = c + 1
                                        
                                                                '********Buscar la cuenta para acumular
                                        
                                       Select Case rsPAgoDetalle("cta_codigo")
                                           Case "0869"
                                                 '"4.41.1.1.1.402.208.11-2"
                                                 vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                                           Case "0870"
                                                '"4.41.1.1.1.402.208.12-1"
                                                 vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                                           Case "0872"
                                                '"4.41.1.1.1.402.208.14-0"
                                                 vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                                           Case "0873"
                                                '"4.41.1.1.1.402.208.16-8"
                                                 vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                                           Case "2676"
                                                '"4.41.1.1.1.402.208.18-6"
                                                 vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                                           Case "0922"
                                                 '"4.41.1.1.1.402.254.01-7"
                                                 vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                                           Case "0921"
                                                 '"4.41.1.1.1.402.254.02-6"
                                                 vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297792"
                                                 vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297809"
                                                 vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297841"
                                                 vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297867"
                                                 vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297875"
                                                 vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297883"
                                                 vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297891"
                                                 vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297916"
                                                 vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297924"
                                                 vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297932"
                                                 vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297940"
                                                 vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297958"
                                                 vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301973"
                                                 vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301999"
                                                 vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                                           Case "1-303515"
                                                 vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                                           Case "1-306379"
                                                 vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                                        End Select
                                        '**************
                                        End If
                                   End If 'Fin por fecha de pago
                                   
                                    If OptFechaImpresion.Value = True Then   '''''''esto pregunta p�r fecha de impresion
                                        If rsPAgoDetalle("fecha_impresion_cheque") >= DTPFechaInicio And rsPAgoDetalle("fecha_impresion_cheque") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, fecha_impresion) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' , '" & rsPAgoDetalle!Fecha_Impresion_Cheque & "') "
                                                c = c + 1
                                        
                                        '********Buscar la cuenta para acumular
                                        
                                       Select Case rsPAgoDetalle("cta_codigo")
                                           Case "0869"
                                                 '"4.41.1.1.1.402.208.11-2"
                                                 vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                                           Case "0870"
                                                '"4.41.1.1.1.402.208.12-1"
                                                 vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                                           Case "0872"
                                                '"4.41.1.1.1.402.208.14-0"
                                                 vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                                           Case "0873"
                                                '"4.41.1.1.1.402.208.16-8"
                                                 vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                                           Case "2676"
                                                '"4.41.1.1.1.402.208.18-6"
                                                 vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                                           Case "0922"
                                                 '"4.41.1.1.1.402.254.01-7"
                                                 vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                                           Case "0921"
                                                 '"4.41.1.1.1.402.254.02-6"
                                                 vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297792"
                                                 vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297809"
                                                 vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297841"
                                                 vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297867"
                                                 vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297875"
                                                 vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297883"
                                                 vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297891"
                                                 vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297916"
                                                 vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297924"
                                                 vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297932"
                                                 vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297940"
                                                 vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297958"
                                                 vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301973"
                                                 vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301999"
                                                 vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                                           Case "1-303515"
                                                 vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                                           Case "1-306379"
                                                 vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                                        End Select
                                        '**************
                                        End If
                                   End If  '''''''esto pregunta p�r fecha de impresion
                                        rsPAgoDetalle.MoveNext
                              Wend
                              
                              
                                'NoRegistros = c
                                'AdoCuenta.Caption = NoRegistros
                                MsgBox c
                              
                            End If
                            
                            
                            'Filtrando los datos
                             If c = 0 Then
                                 MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
                             Else
                                 Set rsFecha = New ADODB.Recordset
                                 If rsFecha.State = 1 Then rsFecha.Close
                                 rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
                                 If rsFecha.RecordCount > 0 Then
                                     Set DtGPagosDetalle.DataSource = rsFecha
                                     Set AdoCuenta.Recordset = rsFecha
                                     NoRegistros = c
                                     AdoCuenta.Caption = NoRegistros
                                    Exit Sub
                                 End If
                             End If
                             
                     
                            'Asignacion de acumulado
                            TxtMovimientoTodos.Text = Monto_Actual_Bolivianos
                            TxtSaldoTodasCuentas.Text = TxtSaldoInicialTotal - Monto_Actual_Bolivianos
            
                            'TxtSACuenta.Text = Monto_Actual_Bolivianos
                            'rsCB("Cta_Acumulado") = Monto_Actual_Bolivianos
                            'rsCB.Update
                            'rsCB.MoveNext
                            'Monto_Actual_Bolivianos = 0
                            'Monto_Actual_Dolares = 0
                    'Wend
                   'resfresca_grid_cuenta_bancaria
            '   End If
            
                '*****  Insertando datos a to_cta_bancaria
                   'Abriendo tabla de cuenta bancaria para las cuentas
                   Set rsCuentaImprimir = New ADODB.Recordset
                   If rsCuentaImprimir.State = 1 Then rsCuentaImprimir.Close
                   rsCuentaImprimir.Open "SELECT * FROM to_cta_bancaria", db, adOpenKeyset, adLockOptimistic
                        Set rsCB = New ADODB.Recordset
                        If rsCB.State = 1 Then rsCB.Close
                        rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
                        While Not rsCB.EOF
                             rsCuentaImprimir.AddNew
                             rsCuentaImprimir("ges_gestion") = rsCB("ges_gestion")
                             rsCuentaImprimir("bco_codigo") = rsCB("bco_codigo")
                             rsCuentaImprimir("cta_codigo") = rsCB("cta_codigo")
                             rsCuentaImprimir("cta_codigo_tgn") = rsCB("cta_codigo_tgn")
                             'rsCuentaImprimir("cta_descripcion_larga") = rsCB("cta_descripcion_larga")
                             rsCuentaImprimir("cta_saldo_inicial") = rsCB("cta_saldo_inicial")
                                       Select Case rsCB("cta_codigo")
                                           Case "0869"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(1)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(1)
                                           Case "0870"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(2)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(2)
                                           Case "0872"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(3)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(3)
                                           Case "0873"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(4)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(4)
                                           Case "2676"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(5)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(5)
                                           Case "0922"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(6)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(6)
                                           Case "0921"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(7)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(7)
                                           Case "1-297792"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(8)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(8)
                                           Case "1-297809"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(9)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(9)
                                           Case "1-297841"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(10)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(10)
                                           Case "1-297867"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(11)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(11)
                                           Case "1-297875"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(12)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(12)
                                           Case "1-297883"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(13)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(13)
                                           Case "1-297891"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(14)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(14)
                                           Case "1-297916"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(15)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(15)
                                           Case "1-297924"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(16)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(16)
                                           Case "1-297932"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(17)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(17)
                                           Case "1-297940"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(18)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(18)
                                           Case "1-297958"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(19)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(19)
                                           Case "1-301973"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(20)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(20)
                                           Case "1-301999"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(21)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(21)
                                           Case "1-302731"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(22)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(22)
                                           Case "1-303515"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(23)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(23)
                                           Case "1-306379"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(24)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(24)
                                           Case "1-302731"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(25)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(25)
                                        End Select
                             
                             rsCuentaImprimir.Update
                            rsCB.MoveNext
                            
                        Wend
                MsgBox "Termin�"
                
            End If
            AVI.Stop
End If



'TODOS LOS REGISTROS POR FECHA
If OptTodasCuentas.Value = True Then
            
            db.Execute "DELETE FROM to_movimiento"
            
            If DtCCuentaOrigen.Text = "" And OptUnaCuenta.Value = True Then
                MsgBox "Lista por cuenta elija cuenta bancaria", vbInformation + vbCritical
                Exit Sub
            End If
                If DtCCuentaOrigen <> "" Then
                    MsgBox "Calculo por cuenta"
                    Proceso_Cuenta "RANGO"
                    Exit Sub
                End If
                
                Monto_Actual_Bolivianos = 0
                Monto_Actual_Dolares = 0
                TxtSaldoActualBol.Text = ""
                If CmbAnio.Text = "" Then
                    MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
                    Exit Sub
                End If
                
            If OptTodasCuentas.Value = True Then
                            rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, pago_detalle.monto_bolivianos,pago_detalle.monto_dolares,fc_beneficiario.denominacion_beneficiario, pago_detalle.codigo_pago, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_registro " & _
                            "FROM (pago_detalle RIGHT JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo ORDER BY pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                            If rsPAgoDetalle.RecordCount > 0 Then
                                While Not rsPAgoDetalle.EOF
                                        If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin Then
                                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','', '" & rsPAgoDetalle!denominacion_beneficiario & "', 0,'" & rsPAgoDetalle!codigo_pago & "',0,'" & rsPAgoDetalle!tipo_cambio & "','','0','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                                c = c + 1
                                        End If
                                       If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                       If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                       If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                                c = c + 1
                                       End If
                                                
                                  '********Buscar la cuenta para acumular
                                  If Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                       Select Case rsPAgoDetalle("cta_codigo")
                                           Case "0869"
                                                 vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                                           Case "0870"
                                                 vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                                           Case "0872"
                                                 vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                                           Case "0873"
                                                 vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                                           Case "2676"
                                                 vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                                           Case "0922"
                                                 vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                                           Case "0921"
                                                 vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297792"
                                                 vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297809"
            '                                     vectorBs(9) = vectorBs(9) + rsPagoDetalle("monto_bolivianos")
            '                                     vectorSus(9) = vectorSus(9) + rsPagoDetalle("monto_dolares")
                                           Case "1-297841"
                                                 vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297867"
                                                 vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297875"
                                                 vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297883"
                                                 vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297891"
                                                 vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297916"
                                                 vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297924"
                                                 vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297932"
                                                 vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297940"
                                                 vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297958"
                                                 vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301973"
                                                 vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301999"
                                                 vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
            '                                     vectorBs(22) = vectorBs(22) + rsPagoDetalle("monto_bolivianos")
            '                                     vectorSus(22) = vectorSus(22) + rsPagoDetalle("monto_dolares")
                                           Case "1-303515"
                                                 vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                                           Case "1-306379"
                                                 vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                                        End Select
                                        '**************
                                        End If
                                rsPAgoDetalle.MoveNext
                              Wend
                                MsgBox c
                            End If
                            
                            
                            'Filtrando los datos
                             If c = 0 Then
                                 MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
                             Else
                                 Set rsFecha = New ADODB.Recordset
                                 If rsFecha.State = 1 Then rsFecha.Close
                                 rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
                                 If rsFecha.RecordCount > 0 Then
                                     Set DtGPagosDetalle.DataSource = rsFecha
                                     Set AdoCuenta.Recordset = rsFecha
                                     NoRegistros = c
                                     AdoCuenta.Caption = NoRegistros
                                    Exit Sub
                                 End If
                             End If
            
            
                '*****  Insertando datos a to_cta_bancaria
                   'Abriendo tabla de cuenta bancaria para las cuentas
                   Set rsCuentaImprimir = New ADODB.Recordset
                   If rsCuentaImprimir.State = 1 Then rsCuentaImprimir.Close
                   rsCuentaImprimir.Open "SELECT * FROM to_cta_bancaria", db, adOpenKeyset, adLockOptimistic
                        Set rsCB = New ADODB.Recordset
                        If rsCB.State = 1 Then rsCB.Close
                        rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
                        While Not rsCB.EOF
                             rsCuentaImprimir.AddNew
                             rsCuentaImprimir("ges_gestion") = rsCB("ges_gestion")
                             rsCuentaImprimir("bco_codigo") = rsCB("bco_codigo")
                             rsCuentaImprimir("cta_codigo") = rsCB("cta_codigo")
                             rsCuentaImprimir("cta_codigo_tgn") = rsCB("cta_codigo_tgn")
                             'rsCuentaImprimir("cta_descripcion_larga") = rsCB("cta_descripcion_larga")
                             rsCuentaImprimir("cta_saldo_inicial") = rsCB("cta_saldo_inicial")
                                       Select Case rsCB("cta_codigo")
                                           Case "0869"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(1)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(1)
                                           Case "0870"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(2)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(2)
                                           Case "0872"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(3)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(3)
                                           Case "0873"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(4)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(4)
                                           Case "2676"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(5)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(5)
                                           Case "0922"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(6)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(6)
                                           Case "0921"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(7)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(7)
                                           Case "1-297792"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(8)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(8)
                                           Case "1-297809"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(9)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(9)
                                           Case "1-297841"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(10)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(10)
                                           Case "1-297867"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(11)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(11)
                                           Case "1-297875"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(12)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(12)
                                           Case "1-297883"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(13)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(13)
                                           Case "1-297891"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(14)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(14)
                                           Case "1-297916"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(15)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(15)
                                           Case "1-297924"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(16)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(16)
                                           Case "1-297932"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(17)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(17)
                                           Case "1-297940"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(18)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(18)
                                           Case "1-297958"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(19)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(19)
                                           Case "1-301973"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(20)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(20)
                                           Case "1-301999"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(21)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(21)
                                           Case "1-302731"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(22)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(22)
                                           Case "1-303515"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(23)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(23)
                                           Case "1-306379"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(24)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(24)
                                           Case "1-302731"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(25)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(25)
                                        End Select
                             
                             rsCuentaImprimir.Update
                            rsCB.MoveNext
                        Wend
                MsgBox "Ya termin�"
        End If
    End If

End Sub

Private Sub CmdImprimeGrid_Click()
''''Dim I As Integer
''''Dim AUXILIAR As String
'''''On Error GoTo temporal:
''''    'Imprimir datos
''''''    Set rsCuentas = New ADODB.Recordset
''''''    If rsCuentas.State = 1 Then rsCuentas.Close
''''''    rsCuentas.Open "SELECT * FROM to_movimiento", db, adOpenStatic, adLockOptimistic
''''''    If rsCuentas.RecordCount > 0 Then
''''''        While Not rsCuentas.EOF
''''''            rsCuentas.Delete
''''''            rsCuentas.MoveNext
''''''        Wend
''''''    End If
''''
''''    Set rsCuentas = New ADODB.Recordset
''''    If rsCuentas.State = 1 Then rsCuentas.Close
''''    rsCuentas.Open "SELECT * FROM to_movimiento", db, adOpenStatic, adLockOptimistic
''''    I = 0
''''    MsgBox NoRegistros
''''    For K = 0 To NoRegistros - 1
''''        I = I + 1
'''''        If I = 20 Then
'''''                MsgBox DtGPagosDetalle.Columns(0)
'''''        End If
''''        rsCuentas.AddNew
''''        DtGPagosDetalle.Row = I
''''        If Not IsNull(DtGPagosDetalle.Columns(0)) Then rsCuentas("fecha_pago") = DtGPagosDetalle.Columns(0)
''''        If Not IsNull(DtGPagosDetalle.Columns(1)) Then rsCuentas("numero_cheque_trf") = DtGPagosDetalle.Columns(1)
''''        If Not IsNull(DtGPagosDetalle.Columns(2)) Then rsCuentas("denominacion_beneficiario") = DtGPagosDetalle.Columns(2)
''''        If Not IsNull(DtGPagosDetalle.Columns(3)) Then rsCuentas("monto_bolivianos") = DtGPagosDetalle.Columns(3)
''''        If Not IsNull(DtGPagosDetalle.Columns(4)) Then rsCuentas("codigo_pago") = DtGPagosDetalle.Columns(4)
''''        If Not IsNull(DtGPagosDetalle.Columns(5)) Then rsCuentas("monto_dolares") = DtGPagosDetalle.Columns(5)
''''        If Not IsNull(DtGPagosDetalle.Columns(6)) Then rsCuentas("tipo_cambio") = DtGPagosDetalle.Columns(6)
''''        If Not IsNull(DtGPagosDetalle.Columns(7)) Then rsCuentas("cta_descripcion_larga") = DtGPagosDetalle.Columns(7)
''''        If Not IsNull(DtGPagosDetalle.Columns(8)) Then rsCuentas("cta_codigo") = DtGPagosDetalle.Columns(8)
''''        If Not IsNull(DtGPagosDetalle.Columns(9)) Then rsCuentas("org_codigo") = DtGPagosDetalle.Columns(9)
''''        rsCuentas.Update
''''    Next K
    'Cry.PaperOrientation = crLandscape
    RepMovi.Show
    
    
 Exit Sub
'temporal:
'    Set rsCuentas = New ADODB.Recordset
'    If rsCuentas.State = 1 Then rsCuentas.Close
'    rsCuentas.Open "SELECT * FROM to_movimiento", db, adOpenDynamic, adLockOptimistic
'    Resume
 
End Sub

Private Sub CmdImprimir_Click()
    RepCtaBancaria.Show
End Sub

Private Sub CmdImprimirMovimiento_Click()
    'RepMovi.Show
'    If OptFechaPago.Value = True Then
'    End If
'    If OptFecha.Value = True Then
'        CryMov.Formulas(5) = "FFechaFin='" & FrmCuentaBancaria.DTPFechaInicio.Value & "'"
'        CryMov.Formulas(6) = "FFechaFin='" & FrmCuentaBancaria.DTPFechaFin.Value & "'"
'    End If
'    If OptMes.Value = True Then
'       CryMov.Formulas(7) = "FMes='" & FrmCuentaBancaria.CmbMes.Text & "'"
'    End If
'    CryMov.ReportFileName = "C:\SAF-2000\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\Rpt_CtaBancaria.rpt"
'    IResult = CryMov.PrintReport
'    If IResult <> 0 Then
'       MsgBox CryMov.LastErrorNumber & " : " & CryMov.LastErrorString, vbCritical + vbOKOnly, "Error..."
'     End If
    'C:\SAF-2000\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones
Dim De As String
Dim A As String
Dim TC As String
Dim Cadena As String
If FrmCuentaBancaria.OptFechaPago.Value = True Then
     Cadena = "REPORTE POR FECHA DE PAGO"
Else
     Cadena = "REPORTE POR FECHA DE IMPRESION"
End If

If FrmCuentaBancaria.OptFechaPago.Value = True Then CryMov.Formulas(4) = "Fecha_Pago_Impresion='" & Cadena & "'"
If FrmCuentaBancaria.OptFechaImpresion.Value = True Then CryMov.Formulas(4) = "Fecha_Pago_Impresion='" & Cadena & "'"

If FrmCuentaBancaria.OptUnaCuenta.Value = True Then
    CryMov.Formulas(1) = "FCodigo_Cuenta='" & FrmCuentaBancaria.DtCCuentaOrigen.Text & "'"
    CryMov.Formulas(2) = "FDescripcion_Cuenta='" & FrmCuentaBancaria.DtCDescripcion.Text & "'"
End If
If FrmCuentaBancaria.OptTodasCuentas.Value = True Then
    TC = "Todas las cuentas"
    CryMov.Formulas(9) = "FTodasCuentas='" & TC & "'"
End If
If FrmCuentaBancaria.OptMes.Value = True Then
    CryMov.Formulas(7) = "FMes='" & FrmCuentaBancaria.CmbMes.Text & "'"
Else
    CryMov.Formulas(7) = "FMes='" & " " & "'"
End If
If OptFecha.Value = True Then
    De = "De"
    A = "A"
    CryMov.Formulas(6) = "FFechaInicio='" & FrmCuentaBancaria.DTPFechaInicio.Value & "'"
    CryMov.Formulas(5) = "FFechaFin='" & FrmCuentaBancaria.DTPFechaFin.Value & "'"
    CryMov.Formulas(2) = "FDe='" & De & "' "
    CryMov.Formulas(0) = "Fa='" & A & "' "
End If

    CryMov.ReportFileName = "C:\SAF-2000\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\Rpt_CtaBancaria.rpt"
    iresult = CryMov.PrintReport
    If iresult <> 0 Then
       MsgBox CryMov.LastErrorNumber & " : " & CryMov.LastErrorString, vbCritical + vbOKOnly, "Error..."
     End If

End Sub

Private Sub CmdmovimientoMes_Click()
Dim c As Long
Dim vectorBs(100) As Double
Dim vectorSus(100) As Double

Dim dia As Variant
Dim mes As Variant
Dim anio As Variant
Dim mes_numeral As Integer
'Dim vectorBs(100) As Double
'Dim vectorSus(100) As Double
Dim condicion As String
'Dim CONDICION As String

'MOVIMIENTO DE PAGOS POR MES

If OptMes.Value = True Then
                
                
                If OptFechaPago.Value = falso And OptFechaImpresion.Value = falso And OptMes.Value = False Then
                    MsgBox "Elija una opci�n por fecha o mes", vbInformation + vbCritical, "Vaidaci�n de datos"
                    Exit Sub
                End If
                
                    swMes = 1
                    swFecha = 0
                
                If DtCCuentaOrigen.Text = "" And OptUnaCuenta.Value = True Then
                    MsgBox "Lista por cuenta bancatia elija cuenta bancaria", vbInformation + vbCritical
                    AVI.Stop
                    Exit Sub
                End If
                '*********
                db.Execute "DELETE FROM to_movimiento"
                db.Execute "DELETE FROM to_cta_Bancaria"
                ''db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                If DtCCuentaOrigen <> "" Then
                    MsgBox "Calculo por cuenta"
                    Proceso_Cuenta "MES"
                    AVI.Stop
                    Exit Sub
                End If
                
                TxtSaldoActualBol.Text = ""
                'Detalle de cuenta bancaria por mes
                    CmbMes.Enabled = True
                    LblMes.ForeColor = &H80&
                    
                If CmbMes.Text = "" Then
                    MsgBox "No existe dato de mes....", vbCritical + vbInformation, "Validaci�n de datos"
                    AVI.Stop
                    Exit Sub
                End If
                            Select Case CmbMes.Text
                                Case "ENERO"
                                    mes_numeral = 1
                                Case "FEBRERO"
                                    mes_numeral = 2
                                Case "MARZO"
                                    mes_numeral = 3
                                Case "ABRIL"
                                    mes_numeral = 4
                                Case "MAYO"
                                    mes_numeral = 5
                                Case "JUNIO"
                                    mes_numeral = 6
                                Case "JULIO"
                                    mes_numeral = 7
                                Case "AGOSTO"
                                    mes_numeral = 8
                                Case "SEPTIEMBRE"
                                    mes_numeral = 9
                                Case "OCTUBRE"
                                    mes_numeral = 10
                                Case "NOVIEMBRE"
                                    mes_numeral = 11
                                Case "DICIEMBRE"
                                    mes_numeral = 12
                            End Select
                    
                    Monto_Actual_Bolivianos = 0
                    Monto_Actual_Dolares = 0
                    
                    If CmbAnio.Text = "" Then
                        MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
                        Exit Sub
                    End If
                '''''''If OptUnaCuenta.Value = True Then
                '''''''    Monto_Actual_Bolivianos = 0
                '''''''    Monto_Actual_Dolares = 0
                '''''''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                '''''''    ''''rsPagoDetalle.Open "SELECT * FROM pago_detalle where month(fecha_pago)='" & mes_numeral & "' and estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                '''''''    rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                '''''''                       " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE month(pago_detalle.fecha_pago)='" & mes_numeral & "' ", db, adOpenKeyset, adLockOptimistic
                '''''''                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                '''''''    Set DtGPagosDetalle.DataSource = rsPagoDetalle
                '''''''    DtGCuenta.Refresh
                '''''''        If rsPagoDetalle.RecordCount > 0 Then
                '''''''        NoRegistros = rsPagoDetalle.RecordCount
                '''''''            While Not rsPagoDetalle.EOF
                '''''''                If Not IsNull(rsPagoDetalle("monto_dolares")) And Not IsNull(rsPagoDetalle("monto_bolivianos")) Then
                '''''''                    Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                '''''''                    Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                '''''''                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                '''''''                End If
                '''''''                rsPagoDetalle.MoveNext
                '''''''            Wend
                '''''''        Else
                '''''''          MsgBox "No existen registros", vbInformation
                '''''''        End If
                '''''''        TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
                '''''''        TxtMovimiento.Text = Monto_Actual_Bolivianos
                '''''''        TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - TxtMovimiento.Text
                '''''''End If
                
                If OptTodasCuentas.Value = True Then
                            'Determinando si es por fecha de pago o fecha de impresion
                            If OptFechaPago.Value = True Then
                                condicion = "month(pago_detalle.fecha_pago)='" & mes_numeral & "'"
                            End If
                            If OptFechaImpresion.Value = True Then
                                condicion = "month(pago_detalle.fecha_impresion_cheque)='" & mes_numeral & "'"
                            End If
                            'month( & condicion &)='" & mes_numeral & "'
                            If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                             rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                                                " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE   " & condicion & " and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
                                If rsPAgoDetalle.RecordCount > 0 Then
                                    Set DtGPagosDetalle.DataSource = rsPAgoDetalle
                                    DtGCuenta.Refresh
                                    While Not rsPAgoDetalle.EOF
                                        If Not IsNull(rsPAgoDetalle("monto_dolares")) And Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then
                                            Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                            Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                            db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                            '********Buscar la cuenta para acumular
                                           Select Case rsPAgoDetalle("cta_codigo")
                                               Case "0869"
                                                     '"4.41.1.1.1.402.208.11-2"
                                                     vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                                               Case "0870"
                                                    '"4.41.1.1.1.402.208.12-1"
                                                     vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                                               Case "0872"
                                                    '"4.41.1.1.1.402.208.14-0"
                                                     vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                                               Case "0873"
                                                    '"4.41.1.1.1.402.208.16-8"
                                                     vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                                               Case "2676"
                                                    '"4.41.1.1.1.402.208.18-6"
                                                     vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                                               Case "0922"
                                                     '"4.41.1.1.1.402.254.01-7"
                                                     vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                                               Case "0921"
                                                     '"4.41.1.1.1.402.254.02-6"
                                                     vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297792"
                                                     vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297809"
                                                     vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297841"
                                                     vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297867"
                                                     vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297875"
                                                     vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297883"
                                                     vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297891"
                                                     vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297916"
                                                     vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297924"
                                                     vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297932"
                                                     vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297940"
                                                     vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                                               Case "1-297958"
                                                     vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                                               Case "1-301973"
                                                     vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                                               Case "1-301999"
                                                     vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                                               Case "1-302731"
                                                     vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                                               Case "1-303515"
                                                     vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                                               Case "1-306379"
                                                     vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                                               Case "1-302731"
                                                     vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                                     vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                                            End Select
                                            '**************
                                        ' rsPagoDetalle.MoveNext
                                            'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_Bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!Cta_descripcion_larga & "','" & rsPagoDetalle!Cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                                        End If
                                        rsPAgoDetalle.MoveNext
                                    Wend
                                Else
                                    Set DtGPagosDetalle.DataSource = rsNada
                                    MsgBox "No existen registros", vbInformation
                                End If
                                'Asignacion de acumulado
                                TxtMovimientoTodos.Text = Monto_Actual_Bolivianos
                                TxtSaldoTodasCuentas.Text = Val(TxtSaldoInicialTotal) - Val(Monto_Actual_Bolivianos)
                                
                                
                                'Reiniciar para la siguiente sumatoria
                                Monto_Actual_Bolivianos = 0
                                Monto_Actual_Dolares = 0
                                resfresca_grid_cuenta_bancaria
                            
                    
                    '*****  Insertando datos a to_cta_bancaria
                       'Abriendo tabla de cuenta bancaria para las cuentas
                       Set rsCuentaImprimir = New ADODB.Recordset
                       If rsCuentaImprimir.State = 1 Then rsCuentaImprimir.Close
                       rsCuentaImprimir.Open "SELECT * FROM to_cta_bancaria", db, adOpenKeyset, adLockOptimistic
                            Set rsCB = New ADODB.Recordset
                            If rsCB.State = 1 Then rsCB.Close
                            rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
                            While Not rsCB.EOF
                                 rsCuentaImprimir.AddNew
                                 rsCuentaImprimir("ges_gestion") = rsCB("ges_gestion")
                                 rsCuentaImprimir("bco_codigo") = rsCB("bco_codigo")
                                 rsCuentaImprimir("cta_codigo") = rsCB("cta_codigo")
                                 rsCuentaImprimir("cta_codigo_tgn") = rsCB("cta_codigo_tgn")
                    'CELIA  cta_descripcion_larga ESTA EN COMENTARIO !!!!!!!!!!!!!!!
                                 rsCuentaImprimir("cta_descripcion_larga") = rsCB("cta_descripcion_larga")
                                 rsCuentaImprimir("cta_saldo_inicial") = rsCB("cta_saldo_inicial")
                             
                                           Select Case rsCB("cta_codigo")
                                               Case "0869"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(1)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(1)
                                               Case "0870"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(2)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(2)
                                               Case "0872"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(3)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(3)
                                               Case "0873"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(4)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(4)
                                               Case "2676"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(5)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(5)
                                               Case "0922"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(6)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(6)
                                               Case "0921"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(7)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(7)
                                               Case "1-297792"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(8)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(8)
                                               Case "1-297809"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(9)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(9)
                                               Case "1-297841"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(10)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(10)
                                               Case "1-297867"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(11)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(11)
                                               Case "1-297875"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(12)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(12)
                                               Case "1-297883"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(13)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(13)
                                               Case "1-297891"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(14)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(14)
                                               Case "1-297916"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(15)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(15)
                                               Case "1-297924"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(16)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(16)
                                               Case "1-297932"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(17)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(17)
                                               Case "1-297940"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(18)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(18)
                                               Case "1-297958"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(19)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(19)
                                               Case "1-301973"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(20)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(20)
                                               Case "1-301999"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(21)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(21)
                                               Case "1-302731"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(22)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(22)
                                               Case "1-303515"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(23)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(23)
                                               Case "1-306379"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(24)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(24)
                                               Case "1-302731"
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(25)
                                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(25)
                                            End Select
                                 rsCuentaImprimir.Update
                                rsCB.MoveNext
                            Wend
                    MsgBox "Ya termin�"
                End If
                AVI.Stop

End If

'MOVIMIENTO POR FECHA
If OptFecha.Value = True Then
            
            
            
            AVI.Open "c:\AVIS\search.avi"
            AVI.Play
            
            'Validaci�n de fecha
            If DTPFechaInicio.Value > DTPFechaFin.Value Or DTPFechaFin.Value < DTPFechaInicio.Value Then
                 MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
                 AVI.Stop
                 Exit Sub
            End If
            
            'Para fines de impresi�n
            swMes = 0
            swFecha = 1
            
            
            '*********
            db.Execute "DELETE FROM to_movimiento"
            
            If DtCCuentaOrigen.Text = "" And OptUnaCuenta.Value = True Then
                MsgBox "Lista por cuenta elija cuenta bancaria", vbInformation + vbCritical
                Exit Sub
            End If
                If DtCCuentaOrigen <> "" Then
                    MsgBox "Calculo por cuenta"
                    Proceso_Cuenta "RANGO"
                    Exit Sub
                End If
                
                Monto_Actual_Bolivianos = 0
                Monto_Actual_Dolares = 0
                TxtSaldoActualBol.Text = ""
                If CmbAnio.Text = "" Then
                    MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
                    Exit Sub
                End If
                
            If OptUnaCuenta.Value = True Then
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                
                rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                                   " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                     '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' ", db, adOpenKeyset, adLockOptimistic
                ''rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
                ''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                ''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                    While Not rsPAgoDetalle.EOF
                        If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                                
                        End If
                        rsPAgoDetalle.MoveNext
                        c = c + 1
                    Wend
                Else
                    MsgBox "No existen registros", vbInformation
                    Set DtGPagosDetalle.DataSource = rsNada
                End If
                
                
                    'Filtrando los datos
                    If c = 0 Then
                        MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
                    Else
                        Set rsFecha = New ADODB.Recordset
                        If rsFecha.State = 1 Then rsFecha.Close
                        rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
                        If rsFecha.RecordCount > 0 Then
                            Set DtGPagosDetalle.DataSource = rsFecha
                            Set AdoCuenta.Recordset = rsFecha
                            NoRegistros = c
                            AdoCuenta.Caption = NoRegistros
                           Exit Sub
                        End If
                    End If
            
                TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
                TxtMovimiento.Text = Monto_Actual_Bolivianos
                TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - TxtMovimiento.Text
                If rsPAgoDetalle.RecordCount > 0 Then
                NoRegistros = rsPAgoDetalle.RecordCount
                        AdoCuenta.Caption = NoRegistros
                        Set DtGPagosDetalle.DataSource = rsPAgoDetalle
                        DtGCuenta.Refresh
                End If
            End If
            
            
            If OptTodasCuentas.Value = True Then
                          If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                          rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_impresion_cheque" & _
                                             " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                            ''***                   rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.codigo_pago, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_registro " & _
                                            ''***                   "FROM (pago_detalle RIGHT JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo ORDER BY pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                                    '"Where (((pago_detalle.fecha_pago) = #5/11/2000#))"
                                                    '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                                                    'rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "' and cta_codigo='" & rsCB("Cta_codigo") & "' ORDER BY fecha_pago", db, adOpenKeyset, adLockOptimistic
                            If rsPAgoDetalle.RecordCount > 0 Then
                                While Not rsPAgoDetalle.EOF
                                    If OptFechaPago.Value = True Then 'Fecha de pago
                                        If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                                c = c + 1
                                        
                                                                '********Buscar la cuenta para acumular
                                        
                                       Select Case rsPAgoDetalle("cta_codigo")
                                           Case "0869"
                                                 '"4.41.1.1.1.402.208.11-2"
                                                 vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                                           Case "0870"
                                                '"4.41.1.1.1.402.208.12-1"
                                                 vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                                           Case "0872"
                                                '"4.41.1.1.1.402.208.14-0"
                                                 vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                                           Case "0873"
                                                '"4.41.1.1.1.402.208.16-8"
                                                 vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                                           Case "2676"
                                                '"4.41.1.1.1.402.208.18-6"
                                                 vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                                           Case "0922"
                                                 '"4.41.1.1.1.402.254.01-7"
                                                 vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                                           Case "0921"
                                                 '"4.41.1.1.1.402.254.02-6"
                                                 vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297792"
                                                 vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297809"
                                                 vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297841"
                                                 vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297867"
                                                 vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297875"
                                                 vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297883"
                                                 vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297891"
                                                 vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297916"
                                                 vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297924"
                                                 vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297932"
                                                 vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297940"
                                                 vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297958"
                                                 vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301973"
                                                 vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301999"
                                                 vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                                           Case "1-303515"
                                                 vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                                           Case "1-306379"
                                                 vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                                        End Select
                                        '**************
                                        End If
                                   End If 'Fin por fecha de pago
                                   
                                    If OptFechaImpresion.Value = True Then   '''''''esto pregunta p�r fecha de impresion
                                        If rsPAgoDetalle("fecha_impresion_cheque") >= DTPFechaInicio And rsPAgoDetalle("fecha_impresion_cheque") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                                c = c + 1
                                        
                                        '********Buscar la cuenta para acumular
                                        
                                       Select Case rsPAgoDetalle("cta_codigo")
                                           Case "0869"
                                                 '"4.41.1.1.1.402.208.11-2"
                                                 vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                                           Case "0870"
                                                '"4.41.1.1.1.402.208.12-1"
                                                 vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                                           Case "0872"
                                                '"4.41.1.1.1.402.208.14-0"
                                                 vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                                           Case "0873"
                                                '"4.41.1.1.1.402.208.16-8"
                                                 vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                                           Case "2676"
                                                '"4.41.1.1.1.402.208.18-6"
                                                 vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                                           Case "0922"
                                                 '"4.41.1.1.1.402.254.01-7"
                                                 vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                                           Case "0921"
                                                 '"4.41.1.1.1.402.254.02-6"
                                                 vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297792"
                                                 vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297809"
                                                 vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297841"
                                                 vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297867"
                                                 vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297875"
                                                 vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297883"
                                                 vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297891"
                                                 vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297916"
                                                 vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297924"
                                                 vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297932"
                                                 vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297940"
                                                 vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                                           Case "1-297958"
                                                 vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301973"
                                                 vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                                           Case "1-301999"
                                                 vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                                           Case "1-303515"
                                                 vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                                           Case "1-306379"
                                                 vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                                           Case "1-302731"
                                                 vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                                 vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                                        End Select
                                        '**************
                                        End If
                                   End If  '''''''esto pregunta p�r fecha de impresion
                                        rsPAgoDetalle.MoveNext
                              Wend
                              
                              
                                'NoRegistros = c
                                'AdoCuenta.Caption = NoRegistros
                                MsgBox c
                              
                            End If
                            
                            
                            'Filtrando los datos
                             If c = 0 Then
                                 MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
                             Else
                                 Set rsFecha = New ADODB.Recordset
                                 If rsFecha.State = 1 Then rsFecha.Close
                                 rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
                                 If rsFecha.RecordCount > 0 Then
                                     Set DtGPagosDetalle.DataSource = rsFecha
                                     Set AdoCuenta.Recordset = rsFecha
                                     NoRegistros = c
                                     AdoCuenta.Caption = NoRegistros
                                    Exit Sub
                                 End If
                             End If
                             
                     
                            'Asignacion de acumulado
                            TxtMovimientoTodos.Text = Monto_Actual_Bolivianos
                            TxtSaldoTodasCuentas.Text = Val(TxtSaldoInicialTotal) - Val(Monto_Actual_Bolivianos)
            
                            'TxtSACuenta.Text = Monto_Actual_Bolivianos
                            'rsCB("Cta_Acumulado") = Monto_Actual_Bolivianos
                            'rsCB.Update
                            'rsCB.MoveNext
                            'Monto_Actual_Bolivianos = 0
                            'Monto_Actual_Dolares = 0
                    'Wend
                   'resfresca_grid_cuenta_bancaria
            '   End If
            
                '*****  Insertando datos a to_cta_bancaria
                   'Abriendo tabla de cuenta bancaria para las cuentas
                   Set rsCuentaImprimir = New ADODB.Recordset
                   If rsCuentaImprimir.State = 1 Then rsCuentaImprimir.Close
                   rsCuentaImprimir.Open "SELECT * FROM to_cta_bancaria", db, adOpenKeyset, adLockOptimistic
                        Set rsCB = New ADODB.Recordset
                        If rsCB.State = 1 Then rsCB.Close
                        rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
                        While Not rsCB.EOF
                             rsCuentaImprimir.AddNew
                             rsCuentaImprimir("ges_gestion") = rsCB("ges_gestion")
                             rsCuentaImprimir("bco_codigo") = rsCB("bco_codigo")
                             rsCuentaImprimir("cta_codigo") = rsCB("cta_codigo")
                             rsCuentaImprimir("cta_codigo_tgn") = rsCB("cta_codigo_tgn")
                             'rsCuentaImprimir("cta_descripcion_larga") = rsCB("cta_descripcion_larga")
                             rsCuentaImprimir("cta_saldo_inicial") = rsCB("cta_saldo_inicial")
                                       Select Case rsCB("cta_codigo")
                                           Case "0869"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(1)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(1)
                                           Case "0870"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(2)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(2)
                                           Case "0872"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(3)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(3)
                                           Case "0873"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(4)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(4)
                                           Case "2676"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(5)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(5)
                                           Case "0922"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(6)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(6)
                                           Case "0921"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(7)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(7)
                                           Case "1-297792"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(8)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(8)
                                           Case "1-297809"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(9)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(9)
                                           Case "1-297841"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(10)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(10)
                                           Case "1-297867"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(11)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(11)
                                           Case "1-297875"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(12)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(12)
                                           Case "1-297883"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(13)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(13)
                                           Case "1-297891"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(14)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(14)
                                           Case "1-297916"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(15)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(15)
                                           Case "1-297924"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(16)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(16)
                                           Case "1-297932"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(17)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(17)
                                           Case "1-297940"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(18)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(18)
                                           Case "1-297958"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(19)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(19)
                                           Case "1-301973"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(20)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(20)
                                           Case "1-301999"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(21)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(21)
                                           Case "1-302731"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(22)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(22)
                                           Case "1-303515"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(23)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(23)
                                           Case "1-306379"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(24)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(24)
                                           Case "1-302731"
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(25)
                                                 rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(25)
                                        End Select
                             
                             rsCuentaImprimir.Update
                            rsCB.MoveNext
                            
                        Wend
                MsgBox "Termin�"
                
            End If
            AVI.Stop
End If
End Sub

Private Sub CmdPorCuenta_Click()
    Unload Me
End Sub

Private Sub CmdRangoFecha_Click()
Dim c As Long
Dim vectorBs(100) As Double
Dim vectorSus(100) As Double


AVI.Open "c:\AVIS\search.avi"
AVI.Play

'Validaci�n de fecha
If DTPFechaInicio.Value > DTPFechaFin.Value Or DTPFechaFin.Value < DTPFechaInicio.Value Then
     MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
     AVI.Stop
     Exit Sub
End If

'Para fines de impresi�n
swMes = 0
swFecha = 1


'*********
db.Execute "DELETE FROM to_movimiento"

If DtCCuentaOrigen.Text = "" And OptUnaCuenta.Value = True Then
    MsgBox "Lista por cuenta elija cuenta bancaria", vbInformation + vbCritical
    Exit Sub
End If
    If DtCCuentaOrigen <> "" Then
        MsgBox "Calculo por cuenta"
        Proceso_Cuenta "RANGO"
        Exit Sub
    End If
    
    Monto_Actual_Bolivianos = 0
    Monto_Actual_Dolares = 0
    TxtSaldoActualBol.Text = ""
    If CmbAnio.Text = "" Then
        MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
        Exit Sub
    End If
    
If OptUnaCuenta.Value = True Then
    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    
    rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                       " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' ", db, adOpenKeyset, adLockOptimistic
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
    ''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    ''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPAgoDetalle.RecordCount > 0 Then
        While Not rsPAgoDetalle.EOF
            If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                    If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                    'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                    
            End If
            rsPAgoDetalle.MoveNext
            c = c + 1
        Wend
    Else
        MsgBox "No existen registros", vbInformation
        Set DtGPagosDetalle.DataSource = rsNada
    End If
    
    
        'Filtrando los datos
        If c = 0 Then
            MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
        Else
            Set rsFecha = New ADODB.Recordset
            If rsFecha.State = 1 Then rsFecha.Close
            rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
            If rsFecha.RecordCount > 0 Then
                Set DtGPagosDetalle.DataSource = rsFecha
                Set AdoCuenta.Recordset = rsFecha
                NoRegistros = c
                AdoCuenta.Caption = NoRegistros
               Exit Sub
            End If
        End If

    TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
    TxtMovimiento.Text = Monto_Actual_Bolivianos
    TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - TxtMovimiento.Text
    If rsPAgoDetalle.RecordCount > 0 Then
    NoRegistros = rsPAgoDetalle.RecordCount
            AdoCuenta.Caption = NoRegistros
            Set DtGPagosDetalle.DataSource = rsPAgoDetalle
            DtGCuenta.Refresh
    End If
End If

Dim condicion As String
If OptTodasCuentas.Value = True Then
              If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
              rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_impresion_cheque" & _
                                 " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                ''***                   rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.codigo_pago, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_registro " & _
                                ''***                   "FROM (pago_detalle RIGHT JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo ORDER BY pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                        '"Where (((pago_detalle.fecha_pago) = #5/11/2000#))"
                                        '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                                        'rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "' and cta_codigo='" & rsCB("Cta_codigo") & "' ORDER BY fecha_pago", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                    While Not rsPAgoDetalle.EOF
                        If OptFechaPago.Value = True Then 'Fecha de pago
                            If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                    If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                    If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "') "
                                    c = c + 1
                            
                                                    '********Buscar la cuenta para acumular
                            
                           Select Case rsPAgoDetalle("cta_codigo")
                               Case "0869"
                                     '"4.41.1.1.1.402.208.11-2"
                                     vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                               Case "0870"
                                    '"4.41.1.1.1.402.208.12-1"
                                     vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                               Case "0872"
                                    '"4.41.1.1.1.402.208.14-0"
                                     vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                               Case "0873"
                                    '"4.41.1.1.1.402.208.16-8"
                                     vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                               Case "2676"
                                    '"4.41.1.1.1.402.208.18-6"
                                     vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                               Case "0922"
                                     '"4.41.1.1.1.402.254.01-7"
                                     vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                               Case "0921"
                                     '"4.41.1.1.1.402.254.02-6"
                                     vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                               Case "1-297792"
                                     vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                               Case "1-297809"
                                     vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                               Case "1-297841"
                                     vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                               Case "1-297867"
                                     vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                               Case "1-297875"
                                     vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                               Case "1-297883"
                                     vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                               Case "1-297891"
                                     vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                               Case "1-297916"
                                     vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                               Case "1-297924"
                                     vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                               Case "1-297932"
                                     vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                               Case "1-297940"
                                     vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                               Case "1-297958"
                                     vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                               Case "1-301973"
                                     vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                               Case "1-301999"
                                     vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                               Case "1-303515"
                                     vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                               Case "1-306379"
                                     vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                            End Select
                            '**************
                            End If
                       End If 'Fin por fecha de pago
''''''
''''''
                        If OptFechaImpresion.Value = True Then   '''''''esto pregunta p�r fecha de impresion
                            If rsPAgoDetalle("fecha_impresion_cheque") >= DTPFechaInicio And rsPAgoDetalle("fecha_impresion_cheque") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                                    If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                                    If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo,fecha_impresion) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "', '" & rsPAgoDetalle!Fecha_Impresion_Cheque & "' ) "
                                    c = c + 1
                            
                            '********Buscar la cuenta para acumular
                            
                           Select Case rsPAgoDetalle("cta_codigo")
                               Case "0869"
                                     '"4.41.1.1.1.402.208.11-2"
                                     vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                               Case "0870"
                                    '"4.41.1.1.1.402.208.12-1"
                                     vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                               Case "0872"
                                    '"4.41.1.1.1.402.208.14-0"
                                     vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                               Case "0873"
                                    '"4.41.1.1.1.402.208.16-8"
                                     vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                               Case "2676"
                                    '"4.41.1.1.1.402.208.18-6"
                                     vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                               Case "0922"
                                     '"4.41.1.1.1.402.254.01-7"
                                     vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                               Case "0921"
                                     '"4.41.1.1.1.402.254.02-6"
                                     vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                               Case "1-297792"
                                     vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                               Case "1-297809"
                                     vectorBs(9) = vectorBs(9) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(9) = vectorSus(9) + rsPAgoDetalle("monto_dolares")
                               Case "1-297841"
                                     vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                               Case "1-297867"
                                     vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                               Case "1-297875"
                                     vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                               Case "1-297883"
                                     vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                               Case "1-297891"
                                     vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                               Case "1-297916"
                                     vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                               Case "1-297924"
                                     vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                               Case "1-297932"
                                     vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                               Case "1-297940"
                                     vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                               Case "1-297958"
                                     vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                               Case "1-301973"
                                     vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                               Case "1-301999"
                                     vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(22) = vectorBs(22) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(22) = vectorSus(22) + rsPAgoDetalle("monto_dolares")
                               Case "1-303515"
                                     vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                               Case "1-306379"
                                     vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                            End Select
                            '**************
                            End If
                       End If  '''''''esto pregunta p�r fecha de impresion
                            rsPAgoDetalle.MoveNext
                  Wend
                  
                  
                    'NoRegistros = c
                    'AdoCuenta.Caption = NoRegistros
                    MsgBox c
                  
                End If
                
                
                'Filtrando los datos
                 If c = 0 Then
                     MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
                 Else
                     Set rsFecha = New ADODB.Recordset
                     If rsFecha.State = 1 Then rsFecha.Close
                     rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
                     If rsFecha.RecordCount > 0 Then
                         Set DtGPagosDetalle.DataSource = rsFecha
                         Set AdoCuenta.Recordset = rsFecha
                         NoRegistros = c
                         AdoCuenta.Caption = NoRegistros
                        Exit Sub
                     End If
                 End If
                 
         
                'Asignacion de acumulado
                TxtMovimientoTodos.Text = Monto_Actual_Bolivianos
                TxtSaldoTodasCuentas.Text = Val(TxtSaldoInicialTotal) - Val(Monto_Actual_Bolivianos)

                'TxtSACuenta.Text = Monto_Actual_Bolivianos
                'rsCB("Cta_Acumulado") = Monto_Actual_Bolivianos
                'rsCB.Update
                'rsCB.MoveNext
                'Monto_Actual_Bolivianos = 0
                'Monto_Actual_Dolares = 0
        'Wend
       'resfresca_grid_cuenta_bancaria
'   End If

    '*****  Insertando datos a to_cta_bancaria
       'Abriendo tabla de cuenta bancaria para las cuentas
       Set rsCuentaImprimir = New ADODB.Recordset
       If rsCuentaImprimir.State = 1 Then rsCuentaImprimir.Close
       rsCuentaImprimir.Open "SELECT * FROM to_cta_bancaria", db, adOpenKeyset, adLockOptimistic
            Set rsCB = New ADODB.Recordset
            If rsCB.State = 1 Then rsCB.Close
            rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
            While Not rsCB.EOF
                 rsCuentaImprimir.AddNew
                 rsCuentaImprimir("ges_gestion") = rsCB("ges_gestion")
                 rsCuentaImprimir("bco_codigo") = rsCB("bco_codigo")
                 rsCuentaImprimir("cta_codigo") = rsCB("cta_codigo")
                 rsCuentaImprimir("cta_codigo_tgn") = rsCB("cta_codigo_tgn")
                 'rsCuentaImprimir("cta_descripcion_larga") = rsCB("cta_descripcion_larga")
                 rsCuentaImprimir("cta_saldo_inicial") = rsCB("cta_saldo_inicial")
                           Select Case rsCB("cta_codigo")
                               Case "0869"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(1)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(1)
                               Case "0870"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(2)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(2)
                               Case "0872"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(3)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(3)
                               Case "0873"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(4)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(4)
                               Case "2676"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(5)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(5)
                               Case "0922"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(6)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(6)
                               Case "0921"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(7)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(7)
                               Case "1-297792"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(8)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(8)
                               Case "1-297809"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(9)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(9)
                               Case "1-297841"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(10)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(10)
                               Case "1-297867"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(11)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(11)
                               Case "1-297875"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(12)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(12)
                               Case "1-297883"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(13)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(13)
                               Case "1-297891"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(14)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(14)
                               Case "1-297916"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(15)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(15)
                               Case "1-297924"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(16)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(16)
                               Case "1-297932"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(17)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(17)
                               Case "1-297940"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(18)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(18)
                               Case "1-297958"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(19)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(19)
                               Case "1-301973"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(20)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(20)
                               Case "1-301999"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(21)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(21)
                               Case "1-302731"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(22)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(22)
                               Case "1-303515"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(23)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(23)
                               Case "1-306379"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(24)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(24)
                               Case "1-302731"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(25)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(25)
                            End Select
                 
                 rsCuentaImprimir.Update
                rsCB.MoveNext
                
            Wend
    MsgBox "Termin�"
    
End If
AVI.Stop

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTesoreria_Click()
'Realiza la actualizacion de la Cta_Saldo actual de tesoreriaSet rsCuenta = New ADODB.Recordset
Dim suma As Variant
Dim sumartf As Variant

    MsgBox "Esperar mensaje de t�rmino"
    If rsCBria.State = 1 Then rsCBria.Close
    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
    
    While Not rsCBria.EOF
                suma = 0
                If rsPg.State = 1 Then rsPg.Close
                rsPg.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & rsCBria("Cta_Codigo") & "'", db, adOpenKeyset, adLockOptimistic
                While Not rsPg.EOF
                      If Not IsNull(rsPg("monto_bolivianos")) Then
                          suma = suma + rsPg("monto_bolivianos")
                      End If
                      rsPg.MoveNext
                Wend
                rsCBria("Cta_Acumulado") = suma
                rsCBria.Update
                rsCBria.MoveNext
    Wend
MsgBox "T E R M I N �  S I N  T R P"


                    
'Caso traspasos
    If rsCBria.State = 1 Then rsCBria.Close
    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
    While Not rsCBria.EOF
                sumartf = 0
                If rsPg.State = 1 Then rsPg.Close
                rsPg.Open "SELECT pagos.org_codigo AS Expr1, pago_detalle.*, pagos.* " & _
                "FROM pago_detalle INNER JOIN pagos ON pago_detalle.Ges_gestion = pagos.ges_gestion AND " & _
                "pago_detalle.org_codigo = pagos.org_codigo AND pago_detalle.codigo_pago = pagos.codigo_pago WHERE cta_codigo='" & rsCBria("Cta_Codigo") & "' AND cta_codigo<>'0922'", db, adOpenKeyset, adLockOptimistic
                While Not rsPg.EOF
                      If Not IsNull(rsPg("monto_Bolivianos")) And rsPg("tipo_comp") = "TRP" Then
                          sumartf = sumartf + rsPg("monto_bolivianos")
                      End If
                      rsPg.MoveNext
                Wend
                rsCBria("Cta_Acumulado") = rsCBria("Cta_Acumulado") + sumartf
                rsCBria.Update
                rsCBria.MoveNext

    Wend

'DETERMINANDO SALDO ACTUAL

    If rsCBria.State = 1 Then rsCBria.Close
    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
    While Not rsCBria.EOF
                rsCBria("Cta_Saldo_Actual") = rsCBria("Cta_saldo_inicial") - rsCBria("Cta_Acumulado") + rsCBria("Cta_Pco_Debe") - rsCBria("Cta_Pco_Haber") + rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe")
                rsCBria.Update
                rsCBria.MoveNext
    Wend
MsgBox "T E R M I N �  S A L D O   A C T U A L"

End Sub

Private Sub CmdTodasCtas_Click()
    RepCtaBancaria.Show
End Sub

Private Sub CmdTodosRegistros_Click()
Dim c As Long
Dim vectorBs(100) As Double
Dim vectorSus(100) As Double
'*********
db.Execute "DELETE FROM to_movimiento"

If DtCCuentaOrigen.Text = "" And OptUnaCuenta.Value = True Then
    MsgBox "Lista por cuenta elija cuenta bancaria", vbInformation + vbCritical
    Exit Sub
End If
    If DtCCuentaOrigen <> "" Then
        MsgBox "Calculo por cuenta"
        Proceso_Cuenta "RANGO"
        Exit Sub
    End If
    
    Monto_Actual_Bolivianos = 0
    Monto_Actual_Dolares = 0
    TxtSaldoActualBol.Text = ""
    If CmbAnio.Text = "" Then
        MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
        Exit Sub
    End If
    
If OptUnaCuenta.Value = True Then
    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    
    rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                       " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' ", db, adOpenKeyset, adLockOptimistic
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
    ''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    ''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPAgoDetalle.RecordCount > 0 Then
        While Not rsPAgoDetalle.EOF
            If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                    If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                    'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                    
            End If
            rsPAgoDetalle.MoveNext
            c = c + 1
        Wend
    Else
        MsgBox "No existen registros", vbInformation
        Set DtGPagosDetalle.DataSource = rsNada
    End If
    
        'Filtrando los datos
        If c = 0 Then
            MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
        Else
            Set rsFecha = New ADODB.Recordset
            If rsFecha.State = 1 Then rsFecha.Close
            rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
            If rsFecha.RecordCount > 0 Then
                Set DtGPagosDetalle.DataSource = rsFecha
                Set AdoCuenta.Recordset = rsFecha
                NoRegistros = c
                AdoCuenta.Caption = NoRegistros
               Exit Sub
            End If
        End If
    TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
    TxtMovimiento.Text = Monto_Actual_Bolivianos
    TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - TxtMovimiento.Text
    If rsPAgoDetalle.RecordCount > 0 Then
    NoRegistros = rsPAgoDetalle.RecordCount
            AdoCuenta.Caption = NoRegistros
            Set DtGPagosDetalle.DataSource = rsPAgoDetalle
            DtGCuenta.Refresh
    End If
End If


If OptTodasCuentas.Value = True Then

                rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, pago_detalle.monto_bolivianos,pago_detalle.monto_dolares,fc_beneficiario.denominacion_beneficiario, pago_detalle.codigo_pago, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_registro " & _
                "FROM (pago_detalle RIGHT JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo ORDER BY pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                    While Not rsPAgoDetalle.EOF
                            If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin Then
                                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','', '" & rsPAgoDetalle!denominacion_beneficiario & "', 0,'" & rsPAgoDetalle!codigo_pago & "',0,'" & rsPAgoDetalle!tipo_cambio & "','','0','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                    c = c + 1
                            End If
                           If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                           If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                           If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                                    c = c + 1
                           End If
                                    
                      '********Buscar la cuenta para acumular
                      If Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                           Select Case rsPAgoDetalle("cta_codigo")
                               Case "0869"
                                     vectorBs(1) = vectorBs(1) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(1) = vectorSus(1) + rsPAgoDetalle("monto_dolares")
                               Case "0870"
                                     vectorBs(2) = vectorBs(2) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(2) = vectorSus(2) + rsPAgoDetalle("monto_dolares")
                               Case "0872"
                                     vectorBs(3) = vectorBs(3) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(3) = vectorSus(3) + rsPAgoDetalle("monto_dolares")
                               Case "0873"
                                     vectorBs(4) = vectorBs(4) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(4) = vectorSus(4) + rsPAgoDetalle("monto_dolares")
                               Case "2676"
                                     vectorBs(5) = vectorBs(5) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(5) = vectorSus(5) + rsPAgoDetalle("monto_dolares")
                               Case "0922"
                                     vectorBs(6) = vectorBs(6) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(6) = vectorSus(6) + rsPAgoDetalle("monto_dolares")
                               Case "0921"
                                     vectorBs(7) = vectorBs(7) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(7) = vectorSus(7) + rsPAgoDetalle("monto_dolares")
                               Case "1-297792"
                                     vectorBs(8) = vectorBs(8) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(8) = vectorSus(8) + rsPAgoDetalle("monto_dolares")
                               Case "1-297809"
'                                     vectorBs(9) = vectorBs(9) + rsPagoDetalle("monto_bolivianos")
'                                     vectorSus(9) = vectorSus(9) + rsPagoDetalle("monto_dolares")
                               Case "1-297841"
                                     vectorBs(10) = vectorBs(10) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(10) = vectorSus(10) + rsPAgoDetalle("monto_dolares")
                               Case "1-297867"
                                     vectorBs(11) = vectorBs(11) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(11) = vectorSus(11) + rsPAgoDetalle("monto_dolares")
                               Case "1-297875"
                                     vectorBs(12) = vectorBs(12) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(12) = vectorSus(12) + rsPAgoDetalle("monto_dolares")
                               Case "1-297883"
                                     vectorBs(13) = vectorBs(13) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(13) = vectorSus(13) + rsPAgoDetalle("monto_dolares")
                               Case "1-297891"
                                     vectorBs(14) = vectorBs(14) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(14) = vectorSus(14) + rsPAgoDetalle("monto_dolares")
                               Case "1-297916"
                                     vectorBs(15) = vectorBs(15) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(15) = vectorSus(15) + rsPAgoDetalle("monto_dolares")
                               Case "1-297924"
                                     vectorBs(16) = vectorBs(16) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(16) = vectorSus(16) + rsPAgoDetalle("monto_dolares")
                               Case "1-297932"
                                     vectorBs(17) = vectorBs(17) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(17) = vectorSus(17) + rsPAgoDetalle("monto_dolares")
                               Case "1-297940"
                                     vectorBs(18) = vectorBs(18) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(18) = vectorSus(18) + rsPAgoDetalle("monto_dolares")
                               Case "1-297958"
                                     vectorBs(19) = vectorBs(19) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(19) = vectorSus(19) + rsPAgoDetalle("monto_dolares")
                               Case "1-301973"
                                     vectorBs(20) = vectorBs(20) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(20) = vectorSus(20) + rsPAgoDetalle("monto_dolares")
                               Case "1-301999"
                                     vectorBs(21) = vectorBs(21) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(21) = vectorSus(21) + rsPAgoDetalle("monto_dolares")
                               Case "1-302731"
'                                     vectorBs(22) = vectorBs(22) + rsPagoDetalle("monto_bolivianos")
'                                     vectorSus(22) = vectorSus(22) + rsPagoDetalle("monto_dolares")
                               Case "1-303515"
                                     vectorBs(23) = vectorBs(23) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(23) = vectorSus(23) + rsPAgoDetalle("monto_dolares")
                               Case "1-306379"
                                     vectorBs(24) = vectorBs(24) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(24) = vectorSus(24) + rsPAgoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(25) = vectorBs(25) + rsPAgoDetalle("monto_bolivianos")
                                     vectorSus(25) = vectorSus(25) + rsPAgoDetalle("monto_dolares")
                            End Select
                            '**************
                            End If
                    rsPAgoDetalle.MoveNext
                  Wend
                    MsgBox c
                End If
                
                
                'Filtrando los datos
                 If c = 0 Then
                     MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
                 Else
                     Set rsFecha = New ADODB.Recordset
                     If rsFecha.State = 1 Then rsFecha.Close
                     rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
                     If rsFecha.RecordCount > 0 Then
                         Set DtGPagosDetalle.DataSource = rsFecha
                         Set AdoCuenta.Recordset = rsFecha
                         NoRegistros = c
                         AdoCuenta.Caption = NoRegistros
                        Exit Sub
                     End If
                 End If


    '*****  Insertando datos a to_cta_bancaria
       'Abriendo tabla de cuenta bancaria para las cuentas
       Set rsCuentaImprimir = New ADODB.Recordset
       If rsCuentaImprimir.State = 1 Then rsCuentaImprimir.Close
       rsCuentaImprimir.Open "SELECT * FROM to_cta_bancaria", db, adOpenKeyset, adLockOptimistic
            Set rsCB = New ADODB.Recordset
            If rsCB.State = 1 Then rsCB.Close
            rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
            While Not rsCB.EOF
                 rsCuentaImprimir.AddNew
                 rsCuentaImprimir("ges_gestion") = rsCB("ges_gestion")
                 rsCuentaImprimir("bco_codigo") = rsCB("bco_codigo")
                 rsCuentaImprimir("cta_codigo") = rsCB("cta_codigo")
                 rsCuentaImprimir("cta_codigo_tgn") = rsCB("cta_codigo_tgn")
                 'rsCuentaImprimir("cta_descripcion_larga") = rsCB("cta_descripcion_larga")
                 rsCuentaImprimir("cta_saldo_inicial") = rsCB("cta_saldo_inicial")
                           Select Case rsCB("cta_codigo")
                               Case "0869"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(1)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(1)
                               Case "0870"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(2)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(2)
                               Case "0872"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(3)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(3)
                               Case "0873"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(4)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(4)
                               Case "2676"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(5)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(5)
                               Case "0922"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(6)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(6)
                               Case "0921"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(7)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(7)
                               Case "1-297792"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(8)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(8)
                               Case "1-297809"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(9)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(9)
                               Case "1-297841"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(10)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(10)
                               Case "1-297867"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(11)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(11)
                               Case "1-297875"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(12)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(12)
                               Case "1-297883"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(13)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(13)
                               Case "1-297891"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(14)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(14)
                               Case "1-297916"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(15)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(15)
                               Case "1-297924"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(16)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(16)
                               Case "1-297932"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(17)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(17)
                               Case "1-297940"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(18)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(18)
                               Case "1-297958"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(19)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(19)
                               Case "1-301973"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(20)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(20)
                               Case "1-301999"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(21)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(21)
                               Case "1-302731"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(22)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(22)
                               Case "1-303515"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(23)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(23)
                               Case "1-306379"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(24)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(24)
                               Case "1-302731"
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Bs") = vectorBs(25)
                                     rsCuentaImprimir("cta_saldo_acumulado_total_Sus") = vectorSus(25)
                            End Select
                 
                 rsCuentaImprimir.Update
                rsCB.MoveNext
            Wend
    MsgBox "Ya termin�"
End If

End Sub

Private Sub DtCCta_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCCta.BoundText
    DtCTgn.BoundText = DtCCta.BoundText
End Sub

Private Sub CmdUnionTablas_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
    Determina_saldo DtCCuentaOrigen
End Sub

Private Sub DtCCuentaOrigen_Change()
    Determina_saldo DtCCuentaOrigen
    TxtSACuenta.Text = ""
    TxtMovimientoCuenta.Text = ""
End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCCuentaOrigen.BoundText
    DtCTgn.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
   DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
   DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
   Determina_saldo DtCCuentaOrigen
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
   DtCTgn.BoundText = DtCDescripcion.BoundText
   DtCCuentaOrigen.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub DtCTgn_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCTgn.BoundText
    DtCCuentaOrigen.BoundText = DtCTgn.BoundText
End Sub




Private Sub Form_Load()
Dim Acumula_Saldo As Variant
    Set rsCta = New ADODB.Recordset
    If rsCta.State = 1 Then rsCuenta.Close
    'rsCuenta.Open "SELECT cta_codigo,Cta_descripcion_larga,Cta_codigo_tgn  FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    rsCta.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                  "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    If rsCta.RecordCount > 0 Then
        NoRegistros = rsCta.RecordCount
        AdoCuenta.Caption = NoRegistros
        Set DtGCuentaBancaria.DataSource = rsCta
    End If
    
        CmbMes.AddItem "ENERO"
        CmbMes.AddItem "FEBRERO"
        CmbMes.AddItem "MARZO"
        CmbMes.AddItem "ABRIL"
        CmbMes.AddItem "MAYO"
        CmbMes.AddItem "JUNIO"
        CmbMes.AddItem "JULIO"
        CmbMes.AddItem "AGOSTO"
        CmbMes.AddItem "SEPTIEMBRE"
        CmbMes.AddItem "OCTUBRE"
        CmbMes.AddItem "NOVIEMBRE"
        CmbMes.AddItem "DICIEMBRE"
        
        CmbAnio.AddItem "1999"
        CmbAnio.AddItem "2000"
        CmbAnio.AddItem "2001"
        CmbAnio.AddItem "2002"
        CmbAnio.AddItem "2003"
        CmbAnio.AddItem "2004"
        CmbAnio.AddItem "2005"
        CmbAnio.AddItem "2006"
    
    'Determinar fecha actual
        DTPFechaInicio.Value = Date
        DTPFechaFin.Value = Date
        
    'Determinar saldo inicial total
    Set rsCuenta = New ADODB.Recordset
    rsCuenta.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    Set AdoCuenta.Recordset = rsCuenta
    If rsCuenta.RecordCount > 0 Then
      NoRegistros = rsCuenta.RecordCount
      AdoCuenta.Caption = NoRegistros
      While Not rsCuenta.EOF
            If Not IsNull(rsCuenta("cta_saldo_inicial")) Then
              Acumula_Saldo = Acumula_Saldo + rsCuenta("cta_saldo_inicial")
            End If
            rsCuenta.MoveNext
      Wend
        'TxtSaldoInicial.Text = Acumula_Saldo
    End If
    TxtSaldoActualBol.Text = ""
    TxtMovimientoCuenta.Text = ""
    TxtSACuenta.Text = ""
    
    
    DtCDescripcion.BoundText = DtCCuentaOrigen.BoundText
    DtCTgn.BoundText = DtCCuentaOrigen.BoundText

	Call SeguridadSet(Me)
End Sub
Public Sub Proceso_Cuenta(filtro As String)
'Busca por numero de cuenta, rango de fecha, etc.,..
'Determina_saldo DtCCuentaOrigen
Dim c As Long

'db.Execute "DELETE FROM to_movimiento"
If filtro = "MES" Then
    Monto_Actual_Bolivianos = 0
    Monto_Actual_Dolares = 0
    TxtSaldoActualBol.Text = ""

    If CmbMes.Text = "" Then
    MsgBox "No existe dato de mes....", vbCritical + vbInformation, "Validaci�n de datos"
    Exit Sub
    End If
            Select Case CmbMes.Text
                Case "ENERO"
                    mes_numeral = 1
                Case "FEBRERO"
                    mes_numeral = 2
                Case "MARZO"
                    mes_numeral = 3
                Case "ABRIL"
                    mes_numeral = 4
                Case "MAYO"
                    mes_numeral = 5
                Case "JUNIO"
                    mes_numeral = 6
                Case "JULIO"
                    mes_numeral = 7
                Case "AGOSTO"
                    mes_numeral = 8
                Case "SEPTIEMBRE"
                    mes_numeral = 9
                Case "OCTUBRE"
                    mes_numeral = 10
                Case "NOVIEMBRE"
                    mes_numeral = 11
                Case "DICIEMBRE"
                    mes_numeral = 12
            End Select

    Monto_Actual_Bolivianos = 0
    Monto_Actual_Dolares = 0
    If CmbAnio.Text = "" Then
        MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
        Exit Sub
    End If

    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    If ChkDanida.Value = 1 Then
        rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.org_codigo <> '999' and pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and month(fecha_pago)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
    If ChkDanida.Value = 0 Then
        If OptFechaPago.Value = True Then
            rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                           "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and month(fecha_pago)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
        End If
        If OptFechaImpresion.Value = True Then
            rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                           "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and month(fecha_impresion_cheque)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
        End If
    End If

   '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    Set DtGPagosDetalle.DataSource = rsPAgoDetalle
    DtGCuenta.Refresh
    'Esto es por rangoa
        If rsPAgoDetalle.RecordCount > 0 Then
             NoRegistros = rsPAgoDetalle.RecordCount
             AdoCuenta.Caption = NoRegistros
            While Not rsPAgoDetalle.EOF
                If Not IsNull(rsPAgoDetalle("monto_dolares")) And Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then
                    Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                    Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                    If IsNull(rsPAgoDetalle!Fecha_Impresion_Cheque) Then
                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "','" & rsPAgoDetalle!cheque_o_trf & "' ) "
                     Else
                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "','" & rsPAgoDetalle!cheque_o_trf & "','" & rsPAgoDetalle!Fecha_Impresion_Cheque & "' ) "
                    End If
                End If
                rsPAgoDetalle.MoveNext
            Wend
        End If
        'TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
        'TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
        TxtMovimientoCuenta = Monto_Actual_Bolivianos
        TxtSACuenta.Text = Val(TxtSICuenta.Text) - Monto_Actual_Bolivianos
End If

If filtro = "RANGO" Then
    Monto_Actual_Bolivianos = 0
    Monto_Actual_Dolares = 0
    TxtSaldoActualBol.Text = ""
    If CmbAnio.Text = "" Then
        MsgBox "Falta a�o", vbCritical, "Validaci�n de datos"
        Exit Sub
    End If
    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & DtCCuentaOrigen.Text & "' and estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic

    If ChkDanida.Value = 1 Then
        rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                           "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.org_codigo<> '999' and pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If

    If ChkDanida.Value = 0 Then
        rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                        "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPAgoDetalle.RecordCount > 0 Then
    'NoRegistros = rsPagoDetalle.RecordCount
    While Not rsPAgoDetalle.EOF
       If OptFechaPago.Value = True Then
            If rsPAgoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPAgoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                    If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                    c = c + 1
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, impresion_cheque) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "','" & rsPAgoDetalle!cheque_o_trf & "','" & rsPAgoDetalle!Fecha_Impresion_Cheque & "' ) "

            End If
       End If
       If OptFechaImpresion.Value = True Then
            If rsPAgoDetalle("fecha_impresion_cheque") >= DTPFechaInicio And rsPAgoDetalle("fecha_impresion_cheque") <= DTPFechaFin And Not IsNull(rsPAgoDetalle("monto_bolivianos")) And Not IsNull(rsPAgoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                    If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                    c = c + 1
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "', '" & rsPAgoDetalle!cheque_o_trf & "','" & rsPAgoDetalle!Fecha_Impresion_Cheque & "' ) "
            End If
       End If
        rsPAgoDetalle.MoveNext
    Wend
'*********AQUI QUEDO PENDIENTE
    'Filtrando los datos
        If c = 0 Then
            MsgBox "No existen registros con esa fecha", vbInformation + vbcritrical
            Set DtGPagosDetalle.DataSource = rsNada
        Else
            Set rsFecha = New ADODB.Recordset
            If rsFecha.State = 1 Then rsFecha.Close
            rsFecha.Open "SELECT * FROM to_movimiento", db, adOpenKeyset, adLockOptimistic
            If rsFecha.RecordCount > 0 Then
                Set DtGPagosDetalle.DataSource = rsFecha
                Set AdoCuenta.Recordset = rsFecha
                NoRegistros = c
                AdoCuenta.Caption = NoRegistros


                TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
                TxtMovimientoCuenta.Text = Monto_Actual_Bolivianos
                TxtSACuenta.Text = Val(TxtSICuenta.Text) - Monto_Actual_Bolivianos
                If rsPAgoDetalle.RecordCount > 0 Then
                        'NoRegistros = rsPagoDetalle.RecordCount
                        'Set DtGPagosDetalle.DataSource = rsPagoDetalle
                        'DtGCuenta.Refresh
                End If

               Exit Sub
            End If
        End If
    End If
End If

If filtro = "GESTION" Then
    Monto_Actual_Bolivianos = 0
    Monto_Actual_Dolares = 0

    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & DtCCuentaOrigen.Text & "' and estado_aprobacion='A' and ges_gestion='" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic

    If ChkDanida.Value = 1 Then
        rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.org_codigo <> '999' and pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.month(fecha_pago)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
    If ChkDanida.Value = 0 Then
        rsPAgoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                       " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion='" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPAgoDetalle.RecordCount > 0 Then
    NoRegistros = rsPAgoDetalle.RecordCount
    AdoCuenta.Caption = NoRegistros
    While Not rsPAgoDetalle.EOF
                If Not IsNull(rsPAgoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPAgoDetalle("monto_bolivianos")
                If Not IsNull(rsPAgoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPAgoDetalle("monto_dolares")
                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPAgoDetalle!fecha_pago & "','" & rsPAgoDetalle!numero_cheque_trf & "', '" & rsPAgoDetalle!denominacion_beneficiario & "', " & rsPAgoDetalle!monto_bolivianos & ",'" & rsPAgoDetalle!codigo_pago & "'," & rsPAgoDetalle!monto_Dolares & ",'" & rsPAgoDetalle!tipo_cambio & "','" & rsPAgoDetalle!cta_descripcion_larga & "','" & rsPAgoDetalle!cta_codigo & "','" & rsPAgoDetalle!Org_Codigo & "' ) "
                rsPAgoDetalle.MoveNext
    Wend
    End If
    TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
    TxtMovimientoCuenta.Text = Monto_Actual_Bolivianos
    TxtSACuenta.Text = Val(TxtSICuenta.Text) - Monto_Actual_Bolivianos

    If rsPAgoDetalle.RecordCount > 0 Then
            NoRegistros = rsPAgoDetalle.RecordCount
            AdoCuenta.Caption = NoRegistros
            Set DtGPagosDetalle.DataSource = rsPAgoDetalle
            DtGCuenta.Refresh
    End If
End If



End Sub
Public Sub Determina_saldo(cuenta As String)
    'Determinar saldo inicial total
    Set rsCuentaBancaria = New ADODB.Recordset
    If rsCuentaBancaria.State = 1 Then rsCuentaBancaria.Close
    rsCuentaBancaria.Open "select * from fc_cuenta_bancaria where cta_codigo='" & cuenta & "'", db, adOpenKeyset, adLockOptimistic
    If rsCuentaBancaria.RecordCount > 0 Then
        TxtSICuenta.Text = rsCuentaBancaria("cta_saldo_inicial")
    End If
End Sub
Public Sub resfresca_grid_cuenta_bancaria()
 'Determinar saldo inicial total
    Set rsCuentaBancaria = New ADODB.Recordset
    If rsCuentaBancaria.State = 1 Then rsCuentaBancaria.Close
    rsCuentaBancaria.Open "select cta_codigo, cta_codigo_tgn, cta_saldo_inicial, cta_saldo_actual, cta_acumulado, cta_saldo_real from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    If rsCuentaBancaria.RecordCount > 0 Then
        NoRegistros = rsCuentaBancaria.RecordCount
        AdoCuenta.Caption = NoRegistros
        Set DtGCuentaBancaria.DataSource = rsCuentaBancaria
        DtGCuentaBancaria.Refresh
    End If
End Sub

Private Sub OptFecha_Click()
    FraFechas.Visible = True
    FraMes.Visible = False
End Sub

Private Sub OptMes_Click()
    FraFechas.Visible = False
    FraMes.Visible = True
End Sub

Private Sub OptTodasCuentas_Click()
    FraCuenta.Visible = False
'    DtGCuentaBancaria.Visible = True
'    FraTotal.Visible = False
'    FraPorCuenta.Visible = False
'
'    DtCCuentaOrigen.Visible = False
'    DtCTgn.Visible = False
'    DtCDescripcion.Visible = False
'    'lblcuenta.Visible = False
'
    DtCCuentaOrigen.Text = ""
    DtcCtaTGN.Text = ""
    DtCCuentaOrigenDes.Text = ""
    FraTodasCuentas.Visible = True
    FraPorCuenta.Visible = False
'
'
'    'Averiguando el monto total
'    Dim Saldo_Inicial As String
'    Set rsCBria = New ADODB.Recordset
'    rsCBria.Open "SELECT sum(Cta_saldo_inicial) as Saldo_Inicial from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
'    If rsCBria.RecordCount > 0 Then
'           TxtSaldoInicialTotal.Text = rsCBria!Saldo_Inicial
'    End If
    
End Sub

Private Sub OptUnaCuenta_Click()
    FraCuenta.Visible = True
    FraTodasCuentas.Visible = False
    FraPorCuenta.Visible = True
'    DtGCuentaBancaria.Visible = False
'    FraTotal.Visible = True
'    FraPorCuenta.Visible = True
'
'
'    DtCCuentaOrigen.Visible = True
'    DtCTgn.Visible = True
'    DtCDescripcion.Visible = True
'    lblcuenta.Visible = True
'
'    DtCCuentaOrigen.Visible = True
'    DtcCtaTGN.Visible = True
'    DtCCuentaOrigenDes.Visible = True
End Sub



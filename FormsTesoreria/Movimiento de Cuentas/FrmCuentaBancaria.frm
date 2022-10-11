VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCuentaBancaria 
   Caption         =   "Cuenta Bancaria"
   ClientHeight    =   8385
   ClientLeft      =   270
   ClientTop       =   135
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComCtl2.Animation AVI 
      Height          =   1140
      Left            =   7395
      TabIndex        =   67
      Top             =   6300
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2011
      _Version        =   393216
      FullWidth       =   64
      FullHeight      =   76
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
      Height          =   1485
      Left            =   7380
      TabIndex        =   36
      Top             =   4770
      Width           =   4500
      Begin VB.TextBox TxtMovimientoCuenta 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   43
         Top             =   1050
         Width           =   1890
      End
      Begin VB.TextBox TxtSICuenta 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   38
         Top             =   480
         Width           =   1890
      End
      Begin VB.TextBox TxtSACuenta 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2430
         TabIndex        =   37
         Top             =   480
         Width           =   1860
      End
      Begin VB.Label Label14 
         Caption         =   "Movimiento"
         Height          =   210
         Left            =   135
         TabIndex        =   44
         Top             =   825
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Left            =   105
         TabIndex        =   40
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label15 
         Caption         =   "Saldo Actual en Bolivianos"
         Height          =   210
         Left            =   2445
         TabIndex        =   39
         Top             =   270
         Width           =   1890
      End
   End
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
         Caption         =   "MOVIMIENTO DE CUENTA BANCARIA"
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
         Left            =   3435
         TabIndex        =   5
         Top             =   150
         Width           =   5685
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
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   390
      Left            =   1335
      Top             =   7470
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
   Begin VB.Frame FraOpciones 
      Height          =   6810
      Left            =   45
      TabIndex        =   6
      Top             =   1035
      Width           =   1245
      Begin VB.CommandButton CmdTesoreria 
         Caption         =   "Tesoreria Actualiza"
         Height          =   690
         Left            =   180
         TabIndex        =   63
         Top             =   5610
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton CmdTodosRegistros 
         Caption         =   "Todos x Fecha"
         Height          =   705
         Left            =   165
         TabIndex        =   59
         Top             =   4110
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimeGrid 
         Caption         =   "Imprime grid"
         Height          =   735
         Left            =   165
         TabIndex        =   48
         Top             =   2340
         Width           =   930
      End
      Begin VB.CommandButton Cmdanio 
         Caption         =   "Anual"
         Height          =   705
         Left            =   165
         TabIndex        =   11
         Top             =   1635
         Width           =   930
      End
      Begin VB.CommandButton CmdRangoFecha 
         Caption         =   "Movimiento por fecha"
         Height          =   735
         Left            =   165
         TabIndex        =   10
         Top             =   900
         Width           =   930
      End
      Begin VB.CommandButton CmdmovimientoMes 
         Caption         =   "Movimiento por mes"
         Height          =   705
         Left            =   165
         TabIndex        =   9
         Top             =   195
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   165
         Picture         =   "FrmCuentaBancaria.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4815
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Todas las Cuentas  Bancarias"
         Height          =   1035
         Left            =   165
         Picture         =   "FrmCuentaBancaria.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3075
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3780
      Left            =   7350
      TabIndex        =   12
      Top             =   1005
      Width           =   4500
      Begin VB.Frame Frame3 
         Height          =   1260
         Left            =   2310
         TabIndex        =   60
         Top             =   1800
         Width           =   2010
         Begin VB.OptionButton OptFechaImpresion 
            Caption         =   "Fecha de Impresión"
            Height          =   300
            Left            =   180
            TabIndex        =   62
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton OptFechaPago 
            Caption         =   "Fecha de Pago"
            Height          =   300
            Left            =   180
            TabIndex        =   61
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.CheckBox ChkDanida 
         Caption         =   "ORGANISMO 999"
         Height          =   405
         Left            =   2190
         TabIndex        =   49
         Top             =   1335
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.OptionButton OptTodasCuentas 
         Caption         =   "X Todas las Cuentas"
         Height          =   315
         Left            =   2160
         TabIndex        =   46
         Top             =   915
         Width           =   2010
      End
      Begin VB.OptionButton OptUnaCuenta 
         Caption         =   "Por Cuenta"
         Height          =   330
         Left            =   2190
         TabIndex        =   45
         Top             =   465
         Value           =   -1  'True
         Width           =   1830
      End
      Begin VB.ComboBox CmbAnio 
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   90
         TabIndex        =   26
         Text            =   "2000"
         Top             =   2100
         Width           =   1845
      End
      Begin VB.ComboBox CmbMes 
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   330
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   375
         Left            =   90
         TabIndex        =   28
         Top             =   870
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24576001
         CurrentDate     =   36413
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   375
         Left            =   90
         TabIndex        =   29
         Top             =   1500
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24576001
         CurrentDate     =   36413
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen1 
         Bindings        =   "FrmCuentaBancaria.frx":0AAC
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   90
         TabIndex        =   30
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
         Bindings        =   "FrmCuentaBancaria.frx":0AC4
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   90
         TabIndex        =   31
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
         Bindings        =   "FrmCuentaBancaria.frx":0ADC
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   75
         TabIndex        =   32
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
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "FrmCuentaBancaria.frx":0AF4
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   60
         TabIndex        =   64
         Top             =   2670
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
         Bindings        =   "FrmCuentaBancaria.frx":0B0C
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   60
         TabIndex        =   65
         Top             =   3375
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
         Bindings        =   "FrmCuentaBancaria.frx":0B24
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   60
         TabIndex        =   66
         Top             =   3030
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Label LblCuenta 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. "
         Height          =   195
         Left            =   90
         TabIndex        =   33
         Top             =   2445
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   90
         TabIndex        =   27
         Top             =   1890
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin"
         Height          =   240
         Left            =   105
         TabIndex        =   16
         Top             =   1290
         Width           =   1590
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio"
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   675
         Width           =   1590
      End
      Begin VB.Label LblMes 
         Caption         =   "Mes"
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   90
         TabIndex        =   14
         Top             =   135
         Width           =   1425
      End
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
      TabIndex        =   20
      Top             =   4665
      Visible         =   0   'False
      Width           =   4485
      Begin VB.TextBox TxtMovimiento 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   42
         Top             =   1020
         Width           =   1890
      End
      Begin VB.TextBox TxtSaldoActualDol 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   34
         Top             =   1020
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox TxtSaldoActualBol 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   22
         Top             =   480
         Width           =   1860
      End
      Begin VB.TextBox TxtSaldoInicial 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   21
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label Label13 
         Caption         =   "Movimiento"
         Height          =   210
         Left            =   135
         TabIndex        =   41
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Actual en Dolares"
         Height          =   210
         Left            =   2130
         TabIndex        =   35
         Top             =   780
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label Label12 
         Caption         =   "Saldo Actual en Bolivianos"
         Height          =   210
         Left            =   2100
         TabIndex        =   24
         Top             =   270
         Width           =   2070
      End
      Begin VB.Label Label11 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Left            =   105
         TabIndex        =   23
         Top             =   255
         Width           =   1185
      End
   End
   Begin MSDataGridLib.DataGrid DtGPagosDetalle 
      Height          =   6345
      Left            =   1335
      TabIndex        =   25
      Top             =   1095
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   11192
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
      Height          =   6285
      Left            =   1365
      TabIndex        =   17
      Top             =   1125
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   11086
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
   Begin MSDataGridLib.DataGrid DtGCuentaBancaria 
      Height          =   6330
      Left            =   1335
      TabIndex        =   47
      Top             =   1095
      Visible         =   0   'False
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   11165
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
   Begin VB.Frame Frame2 
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
      TabIndex        =   50
      Top             =   4785
      Width           =   4485
      Begin VB.TextBox TxtSaldoInicialTotal 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   54
         Top             =   480
         Width           =   1890
      End
      Begin VB.TextBox TxtSaldoTodasCuentas 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   53
         Top             =   480
         Width           =   1860
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2100
         TabIndex        =   52
         Top             =   1020
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox TxtMovimientoTodos 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         TabIndex        =   51
         Top             =   1020
         Width           =   1890
      End
      Begin VB.Label Label20 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Left            =   105
         TabIndex        =   58
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label19 
         Caption         =   "Saldo Actual en Bolivianos"
         Height          =   210
         Left            =   2100
         TabIndex        =   57
         Top             =   270
         Width           =   2070
      End
      Begin VB.Label Label18 
         Caption         =   "Saldo Actual en Dolares"
         Height          =   210
         Left            =   2130
         TabIndex        =   56
         Top             =   780
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label Label17 
         Caption         =   "Movimiento"
         Height          =   210
         Left            =   135
         TabIndex        =   55
         Top             =   780
         Width           =   1005
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha Inicio"
      Height          =   240
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1590
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha Fin"
      Height          =   240
      Left            =   30
      TabIndex        =   18
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
' Módulo:                   Movimiento de Cuenta Bancaria
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmCuentaBancaria.frm
' Descipción :              Movimientos de cuentas bancaria por mes,
'                           por fecha, por año, por cuenta, por todas
'                           las cuentas, etc.
' Formularios relacionados: Main.frm (Padre)
'                           CryPagos, CryCtaBancaria
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creación         01/Mare/ 2000
' Fecha última modificación 20/May/ 2000
' Versión:                  2.0
'========================================================================================
Dim rsCuenta As New ADODB.Recordset
Dim rsPagoDetalle As New ADODB.Recordset
Dim rsCB As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim Monto_Actual_Bolivianos As Long
Dim Monto_Actual_Dolares As Long
Dim NoRegistros As Long
'Para actualizar saldos
Dim rsCBria As New ADODB.Recordset
Dim rsPg As New ADODB.Recordset

'Para fines de impresión en mdulo se definen
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
    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
''    rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion='" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque" & _
                         " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPagoDetalle.RecordCount > 0 Then
    While Not rsPagoDetalle.EOF
                If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                If Not IsNull(rsPagoDetalle("fecha_impresion_cheque")) Then
                       db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "', '" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                Else
                       db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "' ) "
                End If
                rsPagoDetalle.MoveNext
    Wend
        TxtSaldoInicial.Text = Monto_Actual_Bolivianos
        'TxtSaldoActual.Text = Monto_Actual_Bolivianos
    End If
    
    If rsPagoDetalle.RecordCount > 0 Then
            Set DtGPagosDetalle.DataSource = rsPagoDetalle
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
                If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion='" & CmbAnio.Text & "' and cta_codigo='" & rsCB("Cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                While Not rsPagoDetalle.EOF
                            If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                            If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                            'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                            If Not IsNull(rsPagoDetalle("fecha_impresion_cheque")) Then
                                   db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "', '" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                            Else
                                   db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "' ) "
                            End If
              
                            rsPagoDetalle.MoveNext
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
    'RepCuentaBancaria.Show
    RepCtaBancaria.Show
End Sub
Private Sub CmdmovimientoMes_Click()
Dim dia As Variant
Dim mes As Variant
Dim anio As Variant
Dim mes_numeral As Integer
Dim vectorBs(100) As Double
Dim vectorSus(100) As Double
Dim condicion As String


AVI.Open "c:\AVIS\search.avi"
AVI.Play
   'Para fines de impresión
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
    MsgBox "No existe dato de mes....", vbCritical + vbInformation, "Validación de datos"
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
        MsgBox "Falta año", vbCritical, "Validación de datos"
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
            If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
             rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                                " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE   " & condicion & " and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                    Set DtGPagosDetalle.DataSource = rsPagoDetalle
                    DtGCuenta.Refresh
                    While Not rsPagoDetalle.EOF
                        If Not IsNull(rsPagoDetalle("monto_dolares")) And Not IsNull(rsPagoDetalle("monto_bolivianos")) Then
                            Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                            Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                            MsgBox rsPagoDetalle!fecha_impresion_cheque
                            db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "', '" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                            '********Buscar la cuenta para acumular
                           Select Case rsPagoDetalle("cta_codigo")
                               Case "0869"
                                     '"4.41.1.1.1.402.208.11-2"
                                     vectorBs(1) = vectorBs(1) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(1) = vectorSus(1) + rsPagoDetalle("monto_dolares")
                               Case "0870"
                                    '"4.41.1.1.1.402.208.12-1"
                                     vectorBs(2) = vectorBs(2) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(2) = vectorSus(2) + rsPagoDetalle("monto_dolares")
                               Case "0872"
                                    '"4.41.1.1.1.402.208.14-0"
                                     vectorBs(3) = vectorBs(3) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(3) = vectorSus(3) + rsPagoDetalle("monto_dolares")
                               Case "0873"
                                    '"4.41.1.1.1.402.208.16-8"
                                     vectorBs(4) = vectorBs(4) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(4) = vectorSus(4) + rsPagoDetalle("monto_dolares")
                               Case "2676"
                                    '"4.41.1.1.1.402.208.18-6"
                                     vectorBs(5) = vectorBs(5) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(5) = vectorSus(5) + rsPagoDetalle("monto_dolares")
                               Case "0922"
                                     '"4.41.1.1.1.402.254.01-7"
                                     vectorBs(6) = vectorBs(6) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(6) = vectorSus(6) + rsPagoDetalle("monto_dolares")
                               Case "0921"
                                     '"4.41.1.1.1.402.254.02-6"
                                     vectorBs(7) = vectorBs(7) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(7) = vectorSus(7) + rsPagoDetalle("monto_dolares")
                               Case "1-297792"
                                     vectorBs(8) = vectorBs(8) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(8) = vectorSus(8) + rsPagoDetalle("monto_dolares")
                               Case "1-297809"
                                     vectorBs(9) = vectorBs(9) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(9) = vectorSus(9) + rsPagoDetalle("monto_dolares")
                               Case "1-297841"
                                     vectorBs(10) = vectorBs(10) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(10) = vectorSus(10) + rsPagoDetalle("monto_dolares")
                               Case "1-297867"
                                     vectorBs(11) = vectorBs(11) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(11) = vectorSus(11) + rsPagoDetalle("monto_dolares")
                               Case "1-297875"
                                     vectorBs(12) = vectorBs(12) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(12) = vectorSus(12) + rsPagoDetalle("monto_dolares")
                               Case "1-297883"
                                     vectorBs(13) = vectorBs(13) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(13) = vectorSus(13) + rsPagoDetalle("monto_dolares")
                               Case "1-297891"
                                     vectorBs(14) = vectorBs(14) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(14) = vectorSus(14) + rsPagoDetalle("monto_dolares")
                               Case "1-297916"
                                     vectorBs(15) = vectorBs(15) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(15) = vectorSus(15) + rsPagoDetalle("monto_dolares")
                               Case "1-297924"
                                     vectorBs(16) = vectorBs(16) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(16) = vectorSus(16) + rsPagoDetalle("monto_dolares")
                               Case "1-297932"
                                     vectorBs(17) = vectorBs(17) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(17) = vectorSus(17) + rsPagoDetalle("monto_dolares")
                               Case "1-297940"
                                     vectorBs(18) = vectorBs(18) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(18) = vectorSus(18) + rsPagoDetalle("monto_dolares")
                               Case "1-297958"
                                     vectorBs(19) = vectorBs(19) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(19) = vectorSus(19) + rsPagoDetalle("monto_dolares")
                               Case "1-301973"
                                     vectorBs(20) = vectorBs(20) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(20) = vectorSus(20) + rsPagoDetalle("monto_dolares")
                               Case "1-301999"
                                     vectorBs(21) = vectorBs(21) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(21) = vectorSus(21) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(22) = vectorBs(22) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(22) = vectorSus(22) + rsPagoDetalle("monto_dolares")
                               Case "1-303515"
                                     vectorBs(23) = vectorBs(23) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(23) = vectorSus(23) + rsPagoDetalle("monto_dolares")
                               Case "1-306379"
                                     vectorBs(24) = vectorBs(24) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(24) = vectorSus(24) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(25) = vectorBs(25) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(25) = vectorSus(25) + rsPagoDetalle("monto_dolares")
                            End Select
                            '**************
                        ' rsPagoDetalle.MoveNext
                            'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_Bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!Cta_descripcion_larga & "','" & rsPagoDetalle!Cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                        End If
                        rsPagoDetalle.MoveNext
                    Wend
                Else
                    Set DtGPagosDetalle.DataSource = rsNada
                    MsgBox "No existen registros", vbInformation
                End If
                'Asignacion de acumulado
                TxtMovimientoTodos.Text = Monto_Actual_Bolivianos
                TxtSaldoTodasCuentas.Text = TxtSaldoInicialTotal - Monto_Actual_Bolivianos
                
                
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
    MsgBox "Ya terminó"
End If
AVI.Stop
End Sub
Private Sub CmdRangoFecha_Click()
Dim c As Long
Dim vectorBs(100) As Double
Dim vectorSus(100) As Double


AVI.Open "c:\AVIS\search.avi"
AVI.Play

'Para fines de impresión
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
        MsgBox "Falta año", vbCritical, "Validación de datos"
        Exit Sub
    End If
    
If OptUnaCuenta.Value = True Then
    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    
    rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf,pago_detalle.fecha_impresion_cheque" & _
                       " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' ", db, adOpenKeyset, adLockOptimistic
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
    ''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    ''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPagoDetalle.RecordCount > 0 Then
        While Not rsPagoDetalle.EOF
            If rsPagoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPagoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                    If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                    If Not IsNull(rsPagoDetalle("fecha_impresion_cheque")) Then
                         db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "', '" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                    Else
                         db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "' ) "
                    End If
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                    'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                    
            End If
            rsPagoDetalle.MoveNext
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
    If rsPagoDetalle.RecordCount > 0 Then
    NoRegistros = rsPagoDetalle.RecordCount
            AdoCuenta.Caption = NoRegistros
            Set DtGPagosDetalle.DataSource = rsPagoDetalle
            DtGCuenta.Refresh
    End If
End If

Dim condicion As String
If OptTodasCuentas.Value = True Then
              If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
              rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_impresion_cheque, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque" & _
                                 " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                ''***                   rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.codigo_pago, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_registro " & _
                                ''***                   "FROM (pago_detalle RIGHT JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo ORDER BY pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                                        '"Where (((pago_detalle.fecha_pago) = #5/11/2000#))"
                                        '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
                                        'rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "' and cta_codigo='" & rsCB("Cta_codigo") & "' ORDER BY fecha_pago", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                    While Not rsPagoDetalle.EOF
                        If OptFechaPago.Value = True Then 'Fecha de pago
                            If rsPagoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPagoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                                    If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                                    If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                                    If Not IsNull(rsPagoDetalle("fecha_impresion_cheque")) Then
                                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "', '" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                                    Else
                                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "' ) "
                                    End If
                                    c = c + 1
                            
                                                    '********Buscar la cuenta para acumular
                            
                           Select Case rsPagoDetalle("cta_codigo")
                               Case "0869"
                                     '"4.41.1.1.1.402.208.11-2"
                                     vectorBs(1) = vectorBs(1) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(1) = vectorSus(1) + rsPagoDetalle("monto_dolares")
                               Case "0870"
                                    '"4.41.1.1.1.402.208.12-1"
                                     vectorBs(2) = vectorBs(2) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(2) = vectorSus(2) + rsPagoDetalle("monto_dolares")
                               Case "0872"
                                    '"4.41.1.1.1.402.208.14-0"
                                     vectorBs(3) = vectorBs(3) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(3) = vectorSus(3) + rsPagoDetalle("monto_dolares")
                               Case "0873"
                                    '"4.41.1.1.1.402.208.16-8"
                                     vectorBs(4) = vectorBs(4) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(4) = vectorSus(4) + rsPagoDetalle("monto_dolares")
                               Case "2676"
                                    '"4.41.1.1.1.402.208.18-6"
                                     vectorBs(5) = vectorBs(5) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(5) = vectorSus(5) + rsPagoDetalle("monto_dolares")
                               Case "0922"
                                     '"4.41.1.1.1.402.254.01-7"
                                     vectorBs(6) = vectorBs(6) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(6) = vectorSus(6) + rsPagoDetalle("monto_dolares")
                               Case "0921"
                                     '"4.41.1.1.1.402.254.02-6"
                                     vectorBs(7) = vectorBs(7) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(7) = vectorSus(7) + rsPagoDetalle("monto_dolares")
                               Case "1-297792"
                                     vectorBs(8) = vectorBs(8) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(8) = vectorSus(8) + rsPagoDetalle("monto_dolares")
                               Case "1-297809"
                                     vectorBs(9) = vectorBs(9) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(9) = vectorSus(9) + rsPagoDetalle("monto_dolares")
                               Case "1-297841"
                                     vectorBs(10) = vectorBs(10) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(10) = vectorSus(10) + rsPagoDetalle("monto_dolares")
                               Case "1-297867"
                                     vectorBs(11) = vectorBs(11) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(11) = vectorSus(11) + rsPagoDetalle("monto_dolares")
                               Case "1-297875"
                                     vectorBs(12) = vectorBs(12) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(12) = vectorSus(12) + rsPagoDetalle("monto_dolares")
                               Case "1-297883"
                                     vectorBs(13) = vectorBs(13) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(13) = vectorSus(13) + rsPagoDetalle("monto_dolares")
                               Case "1-297891"
                                     vectorBs(14) = vectorBs(14) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(14) = vectorSus(14) + rsPagoDetalle("monto_dolares")
                               Case "1-297916"
                                     vectorBs(15) = vectorBs(15) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(15) = vectorSus(15) + rsPagoDetalle("monto_dolares")
                               Case "1-297924"
                                     vectorBs(16) = vectorBs(16) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(16) = vectorSus(16) + rsPagoDetalle("monto_dolares")
                               Case "1-297932"
                                     vectorBs(17) = vectorBs(17) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(17) = vectorSus(17) + rsPagoDetalle("monto_dolares")
                               Case "1-297940"
                                     vectorBs(18) = vectorBs(18) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(18) = vectorSus(18) + rsPagoDetalle("monto_dolares")
                               Case "1-297958"
                                     vectorBs(19) = vectorBs(19) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(19) = vectorSus(19) + rsPagoDetalle("monto_dolares")
                               Case "1-301973"
                                     vectorBs(20) = vectorBs(20) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(20) = vectorSus(20) + rsPagoDetalle("monto_dolares")
                               Case "1-301999"
                                     vectorBs(21) = vectorBs(21) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(21) = vectorSus(21) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(22) = vectorBs(22) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(22) = vectorSus(22) + rsPagoDetalle("monto_dolares")
                               Case "1-303515"
                                     vectorBs(23) = vectorBs(23) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(23) = vectorSus(23) + rsPagoDetalle("monto_dolares")
                               Case "1-306379"
                                     vectorBs(24) = vectorBs(24) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(24) = vectorSus(24) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(25) = vectorBs(25) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(25) = vectorSus(25) + rsPagoDetalle("monto_dolares")
                            End Select
                            '**************
                            End If
                       End If 'Fin por fecha de pago
                       
                        If OptFechaImpresion.Value = True Then   '''''''esto pregunta pòr fecha de impresion
                            If rsPagoDetalle("fecha_impresion_cheque") >= DTPFechaInicio And rsPagoDetalle("fecha_impresion_cheque") <= DTPFechaFin And Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                                    If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                                    If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                                    
                                    If Not IsNull(rsPagoDetalle("fecha_impresion_cheque")) Then
                                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "', '" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                                    Else
                                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "' ) "
                                    End If
                              
                                    'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                                    c = c + 1
                            
                            '********Buscar la cuenta para acumular
                            
                           Select Case rsPagoDetalle("cta_codigo")
                               Case "0869"
                                     '"4.41.1.1.1.402.208.11-2"
                                     vectorBs(1) = vectorBs(1) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(1) = vectorSus(1) + rsPagoDetalle("monto_dolares")
                               Case "0870"
                                    '"4.41.1.1.1.402.208.12-1"
                                     vectorBs(2) = vectorBs(2) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(2) = vectorSus(2) + rsPagoDetalle("monto_dolares")
                               Case "0872"
                                    '"4.41.1.1.1.402.208.14-0"
                                     vectorBs(3) = vectorBs(3) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(3) = vectorSus(3) + rsPagoDetalle("monto_dolares")
                               Case "0873"
                                    '"4.41.1.1.1.402.208.16-8"
                                     vectorBs(4) = vectorBs(4) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(4) = vectorSus(4) + rsPagoDetalle("monto_dolares")
                               Case "2676"
                                    '"4.41.1.1.1.402.208.18-6"
                                     vectorBs(5) = vectorBs(5) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(5) = vectorSus(5) + rsPagoDetalle("monto_dolares")
                               Case "0922"
                                     '"4.41.1.1.1.402.254.01-7"
                                     vectorBs(6) = vectorBs(6) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(6) = vectorSus(6) + rsPagoDetalle("monto_dolares")
                               Case "0921"
                                     '"4.41.1.1.1.402.254.02-6"
                                     vectorBs(7) = vectorBs(7) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(7) = vectorSus(7) + rsPagoDetalle("monto_dolares")
                               Case "1-297792"
                                     vectorBs(8) = vectorBs(8) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(8) = vectorSus(8) + rsPagoDetalle("monto_dolares")
                               Case "1-297809"
                                     vectorBs(9) = vectorBs(9) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(9) = vectorSus(9) + rsPagoDetalle("monto_dolares")
                               Case "1-297841"
                                     vectorBs(10) = vectorBs(10) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(10) = vectorSus(10) + rsPagoDetalle("monto_dolares")
                               Case "1-297867"
                                     vectorBs(11) = vectorBs(11) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(11) = vectorSus(11) + rsPagoDetalle("monto_dolares")
                               Case "1-297875"
                                     vectorBs(12) = vectorBs(12) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(12) = vectorSus(12) + rsPagoDetalle("monto_dolares")
                               Case "1-297883"
                                     vectorBs(13) = vectorBs(13) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(13) = vectorSus(13) + rsPagoDetalle("monto_dolares")
                               Case "1-297891"
                                     vectorBs(14) = vectorBs(14) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(14) = vectorSus(14) + rsPagoDetalle("monto_dolares")
                               Case "1-297916"
                                     vectorBs(15) = vectorBs(15) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(15) = vectorSus(15) + rsPagoDetalle("monto_dolares")
                               Case "1-297924"
                                     vectorBs(16) = vectorBs(16) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(16) = vectorSus(16) + rsPagoDetalle("monto_dolares")
                               Case "1-297932"
                                     vectorBs(17) = vectorBs(17) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(17) = vectorSus(17) + rsPagoDetalle("monto_dolares")
                               Case "1-297940"
                                     vectorBs(18) = vectorBs(18) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(18) = vectorSus(18) + rsPagoDetalle("monto_dolares")
                               Case "1-297958"
                                     vectorBs(19) = vectorBs(19) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(19) = vectorSus(19) + rsPagoDetalle("monto_dolares")
                               Case "1-301973"
                                     vectorBs(20) = vectorBs(20) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(20) = vectorSus(20) + rsPagoDetalle("monto_dolares")
                               Case "1-301999"
                                     vectorBs(21) = vectorBs(21) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(21) = vectorSus(21) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(22) = vectorBs(22) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(22) = vectorSus(22) + rsPagoDetalle("monto_dolares")
                               Case "1-303515"
                                     vectorBs(23) = vectorBs(23) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(23) = vectorSus(23) + rsPagoDetalle("monto_dolares")
                               Case "1-306379"
                                     vectorBs(24) = vectorBs(24) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(24) = vectorSus(24) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(25) = vectorBs(25) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(25) = vectorSus(25) + rsPagoDetalle("monto_dolares")
                            End Select
                            '**************
                            End If
                       End If  '''''''esto pregunta pòr fecha de impresion
                            rsPagoDetalle.MoveNext
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
    MsgBox "Terminó"
End If
AVI.Stop
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTesoreria_Click()
'Realiza la actualizacion de la Cta_Saldo actual de tesoreriaSet rsCuenta = New ADODB.Recordset
Dim suma As Double
MsgBox "Esperar mensaje de término"
    If rsCBria.State = 1 Then rsCBria.Close
    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
    suma = 0
    While Not rsCBria.EOF
                If rsPg.State = 1 Then rsPg.Close
                rsPg.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & rsCBria("Cta_Codigo") & "'", db, adOpenKeyset, adLockOptimistic
                While Not rsPg.EOF
                      If Not IsNull(rsPg("monto_bolivianos")) Then
                          suma = suma + rsPg("monto_bolivianos")
                      End If
                      rsPg.MoveNext
                Wend
                rsCBria("Cta_Saldo_Actual") = suma + rsCBria("Cta_debe") - rsCBria("CTA_haber")
                MsgBox rsCBria("Cta_Saldo_Actual")
                rsCBria.Update
                rsCBria.MoveNext
    Wend
MsgBox "T E R M I N Ó"

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
        MsgBox "Falta año", vbCritical, "Validación de datos"
        Exit Sub
    End If
    
If OptUnaCuenta.Value = True Then
    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    
    rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                       " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' ", db, adOpenKeyset, adLockOptimistic
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle where fecha_pago = " & DTPFechaInicio.Value & "", db, adOpenKeyset, adLockOptimistic
    ''    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    ''    rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPagoDetalle.RecordCount > 0 Then
        While Not rsPagoDetalle.EOF
            If rsPagoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPagoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                    If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                    'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                    
            End If
            rsPagoDetalle.MoveNext
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
    If rsPagoDetalle.RecordCount > 0 Then
    NoRegistros = rsPagoDetalle.RecordCount
            AdoCuenta.Caption = NoRegistros
            Set DtGPagosDetalle.DataSource = rsPagoDetalle
            DtGCuenta.Refresh
    End If
End If


If OptTodasCuentas.Value = True Then

''    Set rsCB = New ADODB.Recordset
''    If rsCB.State = 1 Then rsCB.Close
''    rsCB.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
''    If rsCB.RecordCount > 0 Then
''    'NoRegistros = rsPagoDetalle.RecordCount
''    rsCB("Cta_Acumulado") = 50
    
''    While Not rsCB.EOF


'***              If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
'***              rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                                 " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic

                   rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, pago_detalle.monto_bolivianos,pago_detalle.monto_dolares,fc_beneficiario.denominacion_beneficiario, pago_detalle.codigo_pago, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_registro " & _
                   "FROM (pago_detalle RIGHT JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo ORDER BY pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                   '"Where (((pago_detalle.fecha_pago) = #5/11/2000#))"

                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
            
            'rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "' and cta_codigo='" & rsCB("Cta_codigo") & "' ORDER BY fecha_pago", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                    While Not rsPagoDetalle.EOF
                            If rsPagoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPagoDetalle("FECHA_PAGO") <= DTPFechaFin Then
'                                    If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
'                                    If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','', '" & rsPagoDetalle!denominacion_beneficiario & "', 0,'" & rsPagoDetalle!codigo_pago & "',0,'" & rsPagoDetalle!tipo_cambio & "','','0','" & rsPagoDetalle!org_codigo & "' ) "
                                    c = c + 1
                            End If
                           If rsPagoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPagoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                           If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                           If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                                    'db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', '" & rsPagoDetalle!monto_bolivianos & "','" & rsPagoDetalle!codigo_pago & "','" & rsPagoDetalle!monto_Dolares & "','" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                                    c = c + 1
                           End If
                           'End If
                                    
                      '********Buscar la cuenta para acumular
                      If Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                           Select Case rsPagoDetalle("cta_codigo")
                               Case "0869"
                                     '"4.41.1.1.1.402.208.11-2"
                                     vectorBs(1) = vectorBs(1) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(1) = vectorSus(1) + rsPagoDetalle("monto_dolares")
                               Case "0870"
                                    '"4.41.1.1.1.402.208.12-1"
                                     vectorBs(2) = vectorBs(2) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(2) = vectorSus(2) + rsPagoDetalle("monto_dolares")
                               Case "0872"
                                    '"4.41.1.1.1.402.208.14-0"
                                     vectorBs(3) = vectorBs(3) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(3) = vectorSus(3) + rsPagoDetalle("monto_dolares")
                               Case "0873"
                                    '"4.41.1.1.1.402.208.16-8"
                                     vectorBs(4) = vectorBs(4) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(4) = vectorSus(4) + rsPagoDetalle("monto_dolares")
                               Case "2676"
                                    '"4.41.1.1.1.402.208.18-6"
                                     vectorBs(5) = vectorBs(5) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(5) = vectorSus(5) + rsPagoDetalle("monto_dolares")
                               Case "0922"
                                     '"4.41.1.1.1.402.254.01-7"
                                     vectorBs(6) = vectorBs(6) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(6) = vectorSus(6) + rsPagoDetalle("monto_dolares")
                               Case "0921"
                                     '"4.41.1.1.1.402.254.02-6"
                                     vectorBs(7) = vectorBs(7) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(7) = vectorSus(7) + rsPagoDetalle("monto_dolares")
                               Case "1-297792"
                                     vectorBs(8) = vectorBs(8) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(8) = vectorSus(8) + rsPagoDetalle("monto_dolares")
                               Case "1-297809"
'                                     vectorBs(9) = vectorBs(9) + rsPagoDetalle("monto_bolivianos")
'                                     vectorSus(9) = vectorSus(9) + rsPagoDetalle("monto_dolares")
                               Case "1-297841"
                                     vectorBs(10) = vectorBs(10) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(10) = vectorSus(10) + rsPagoDetalle("monto_dolares")
                               Case "1-297867"
                                     vectorBs(11) = vectorBs(11) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(11) = vectorSus(11) + rsPagoDetalle("monto_dolares")
                               Case "1-297875"
                                     vectorBs(12) = vectorBs(12) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(12) = vectorSus(12) + rsPagoDetalle("monto_dolares")
                               Case "1-297883"
                                     vectorBs(13) = vectorBs(13) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(13) = vectorSus(13) + rsPagoDetalle("monto_dolares")
                               Case "1-297891"
                                     vectorBs(14) = vectorBs(14) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(14) = vectorSus(14) + rsPagoDetalle("monto_dolares")
                               Case "1-297916"
                                     vectorBs(15) = vectorBs(15) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(15) = vectorSus(15) + rsPagoDetalle("monto_dolares")
                               Case "1-297924"
                                     vectorBs(16) = vectorBs(16) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(16) = vectorSus(16) + rsPagoDetalle("monto_dolares")
                               Case "1-297932"
                                     vectorBs(17) = vectorBs(17) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(17) = vectorSus(17) + rsPagoDetalle("monto_dolares")
                               Case "1-297940"
                                     vectorBs(18) = vectorBs(18) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(18) = vectorSus(18) + rsPagoDetalle("monto_dolares")
                               Case "1-297958"
                                     vectorBs(19) = vectorBs(19) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(19) = vectorSus(19) + rsPagoDetalle("monto_dolares")
                               Case "1-301973"
                                     vectorBs(20) = vectorBs(20) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(20) = vectorSus(20) + rsPagoDetalle("monto_dolares")
                               Case "1-301999"
                                     vectorBs(21) = vectorBs(21) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(21) = vectorSus(21) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
'                                     vectorBs(22) = vectorBs(22) + rsPagoDetalle("monto_bolivianos")
'                                     vectorSus(22) = vectorSus(22) + rsPagoDetalle("monto_dolares")
                               Case "1-303515"
                                     vectorBs(23) = vectorBs(23) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(23) = vectorSus(23) + rsPagoDetalle("monto_dolares")
                               Case "1-306379"
                                     vectorBs(24) = vectorBs(24) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(24) = vectorSus(24) + rsPagoDetalle("monto_dolares")
                               Case "1-302731"
                                     vectorBs(25) = vectorBs(25) + rsPagoDetalle("monto_bolivianos")
                                     vectorSus(25) = vectorSus(25) + rsPagoDetalle("monto_dolares")
                            End Select
                            '**************
                            End If
                            
'                        End If
                    rsPagoDetalle.MoveNext
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
    MsgBox "Ya termi`nó"
End If

End Sub

Private Sub DtCCta_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCCta.BoundText
    DtCTgn.BoundText = DtCCta.BoundText
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
Dim Acumula_Saldo As Long
    Set rsCuenta = New ADODB.Recordset
    If rsCuenta.State = 1 Then rsCuenta.Close
    'rsCuenta.Open "SELECT cta_codigo,Cta_descripcion_larga,Cta_codigo_tgn  FROM fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    rsCuenta.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                  "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    If rsCuenta.RecordCount > 0 Then
        NoRegistros = rsCuenta.RecordCount
        AdoCuenta.Caption = NoRegistros
        Set DtGCuenta.DataSource = rsCuenta
        Set AdoCuenta.Recordset = rsCuenta
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
        TxtSaldoInicial.Text = Acumula_Saldo
    End If
    TxtSaldoActualBol.Text = ""
    TxtMovimientoCuenta.Text = ""
    TxtSACuenta.Text = ""
    
    
    DtCDescripcion.BoundText = DtCCuentaOrigen.BoundText
    DtCTgn.BoundText = DtCCuentaOrigen.BoundText

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
    MsgBox "No existe dato de mes....", vbCritical + vbInformation, "Validación de datos"
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
        MsgBox "Falta año", vbCritical, "Validación de datos"
        Exit Sub
    End If
    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle where cta_codigo='" & DtCCuentaOrigen.Text & "' and month(fecha_pago)='" & mes_numeral & "' and estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If ChkDanida.Value = 1 Then
        rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.org_codigo <> '999' and pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and month(fecha_pago)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
    If ChkDanida.Value = 0 Then
        If OptFechaPago.Value = True Then
            rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                           "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and month(fecha_pago)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
        End If
        If OptFechaImpresion.Value = True Then
            rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                           "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and month(fecha_impresion_cheque)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago ", db, adOpenKeyset, adLockOptimistic
        End If
    End If
    
   '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    Set DtGPagosDetalle.DataSource = rsPagoDetalle
    DtGCuenta.Refresh
    'Esto es por rangoa
        If rsPagoDetalle.RecordCount > 0 Then
             NoRegistros = rsPagoDetalle.RecordCount
             AdoCuenta.Caption = NoRegistros
            While Not rsPagoDetalle.EOF
                If Not IsNull(rsPagoDetalle("monto_dolares")) And Not IsNull(rsPagoDetalle("monto_bolivianos")) Then
                    Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                    Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")

                    If IsNull(rsPagoDetalle!fecha_impresion_cheque) Then
                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "','" & rsPagoDetalle!cheque_o_trf & "' ) "
                     Else
                        db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "','" & rsPagoDetalle!cheque_o_trf & "','" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                    End If
                End If
                rsPagoDetalle.MoveNext
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
        MsgBox "Falta año", vbCritical, "Validación de datos"
        Exit Sub
    End If
    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & DtCCuentaOrigen.Text & "' and estado_aprobacion='A' and ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    
    If ChkDanida.Value = 1 Then
        rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                           "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.org_codigo<> '999' and pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
    
    If ChkDanida.Value = 0 Then
        rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.cheque_o_trf, pago_detalle.fecha_impresion_cheque " & _
                        "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE  pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPagoDetalle.RecordCount > 0 Then
    'NoRegistros = rsPagoDetalle.RecordCount
    While Not rsPagoDetalle.EOF
       If OptFechaPago.Value = True Then
            If rsPagoDetalle("FECHA_PAGO") >= DTPFechaInicio And rsPagoDetalle("FECHA_PAGO") <= DTPFechaFin And Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                    If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                    c = c + 1
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, impresion_cheque) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "','" & rsPagoDetalle!cheque_o_trf & "','" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
                    
            End If
       End If
       If OptFechaImpresion.Value = True Then
            If rsPagoDetalle("fecha_impresion_cheque") >= DTPFechaInicio And rsPagoDetalle("fecha_impresion_cheque") <= DTPFechaFin And Not IsNull(rsPagoDetalle("monto_bolivianos")) And Not IsNull(rsPagoDetalle("monto_dolares")) Then
                    If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                    If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                    c = c + 1
                    db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo, chq_trf, fecha_impresion) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "', '" & rsPagoDetalle!cheque_o_trf & "','" & rsPagoDetalle!fecha_impresion_cheque & "' ) "
            End If
       End If
        rsPagoDetalle.MoveNext
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
                If rsPagoDetalle.RecordCount > 0 Then
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
    
    If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
    ''rsPagoDetalle.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & DtCCuentaOrigen.Text & "' and estado_aprobacion='A' and ges_gestion='" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    
    If ChkDanida.Value = 1 Then
        rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.org_codigo <> '999' and pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.month(fecha_pago)='" & mes_numeral & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
    If ChkDanida.Value = 0 Then
        rsPagoDetalle.Open "SELECT pago_detalle.fecha_pago,pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
                       " FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.cta_codigo='" & DtCCuentaOrigen.Text & "' and pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion='" & CmbAnio.Text & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
    End If
                         '" FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.estado_aprobacion='A' and pago_detalle.ges_gestion= '" & CmbAnio.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPagoDetalle.RecordCount > 0 Then
    NoRegistros = rsPagoDetalle.RecordCount
    AdoCuenta.Caption = NoRegistros
    While Not rsPagoDetalle.EOF
                If Not IsNull(rsPagoDetalle("monto_bolivianos")) Then Monto_Actual_Bolivianos = Monto_Actual_Bolivianos + rsPagoDetalle("monto_bolivianos")
                If Not IsNull(rsPagoDetalle("monto_dolares")) Then Monto_Actual_Dolares = Monto_Actual_Dolares + rsPagoDetalle("monto_dolares")
                db.Execute "insert into to_movimiento (fecha_pago, numero_cheque_trf, denominacion_beneficiario, monto_bolivianos, codigo_pago, monto_dolares, tipo_cambio, cta_descripcion_larga, cta_codigo, org_codigo) values ('" & rsPagoDetalle!fecha_pago & "','" & rsPagoDetalle!numero_cheque_trf & "', '" & rsPagoDetalle!denominacion_beneficiario & "', " & rsPagoDetalle!monto_bolivianos & ",'" & rsPagoDetalle!codigo_pago & "'," & rsPagoDetalle!monto_Dolares & ",'" & rsPagoDetalle!tipo_cambio & "','" & rsPagoDetalle!cta_descripcion_larga & "','" & rsPagoDetalle!cta_codigo & "','" & rsPagoDetalle!org_codigo & "' ) "
                rsPagoDetalle.MoveNext
    Wend
    End If
    TxtSaldoActualBol.Text = Val(TxtSaldoInicial.Text) - Monto_Actual_Bolivianos
    TxtMovimientoCuenta.Text = Monto_Actual_Bolivianos
    TxtSACuenta.Text = Val(TxtSICuenta.Text) - Monto_Actual_Bolivianos
    
    If rsPagoDetalle.RecordCount > 0 Then
            NoRegistros = rsPagoDetalle.RecordCount
            AdoCuenta.Caption = NoRegistros
            Set DtGPagosDetalle.DataSource = rsPagoDetalle
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

Private Sub OptTodasCuentas_Click()
    DtGCuentaBancaria.Visible = True
    FraTotal.Visible = False
    FraPorCuenta.Visible = False
    
    DtCCuentaOrigen.Visible = False
    DtCTgn.Visible = False
    DtCDescripcion.Visible = False
    LblCuenta.Visible = False
    
    DtCCuentaOrigen.Text = ""
    DtcCtaTGN.Text = ""
    DtCCuentaOrigenDes.Text = ""
    
    
    'Averiguando el monto total
    Dim Saldo_Inicial As String
    Set rsCBria = New ADODB.Recordset
    rsCBria.Open "SELECT sum(Cta_saldo_inicial) as Saldo_Inicial from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    If rsCBria.RecordCount > 0 Then
           TxtSaldoInicialTotal.Text = rsCBria!Saldo_Inicial
    End If
    
End Sub

Private Sub OptUnaCuenta_Click()
    DtGCuentaBancaria.Visible = False
    FraTotal.Visible = True
    FraPorCuenta.Visible = True
    
    
    DtCCuentaOrigen.Visible = True
    DtCTgn.Visible = True
    DtCDescripcion.Visible = True
    LblCuenta.Visible = True
    
    DtCCuentaOrigen.Visible = True
    DtcCtaTGN.Visible = True
    DtCCuentaOrigenDes.Visible = True
End Sub



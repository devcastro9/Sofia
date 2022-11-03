VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmTributosFiscales 
   Caption         =   "Pagos y Tributos Fiscales"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmTributosFiscales.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1740
      Left            =   3135
      TabIndex        =   68
      Top             =   3690
      Width           =   2235
   End
   Begin VB.TextBox TxtJustificacion 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   15
      MultiLine       =   -1  'True
      TabIndex        =   67
      Top             =   8490
      Width           =   11325
   End
   Begin Crystal.CrystalReport CryHistorico 
      Left            =   5625
      Top             =   2775
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryPagosBeneficiario 
      Left            =   1680
      Top             =   5085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CmdImprimirGridSeleccionados 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   9645
      TabIndex        =   27
      Top             =   5115
      Width           =   1665
   End
   Begin VB.CommandButton CmdGuardarHistorico 
      Caption         =   "Vaciar en Histórico"
      Height          =   375
      Left            =   7935
      TabIndex        =   26
      Top             =   5115
      Width           =   1695
   End
   Begin VB.CommandButton CmdLimpiarGrid 
      Caption         =   "Limpiar el grid"
      Height          =   375
      Left            =   6180
      TabIndex        =   25
      Top             =   5115
      Width           =   1740
   End
   Begin VB.CommandButton CmdDevolver 
      Caption         =   "<<"
      Height          =   525
      Left            =   5460
      TabIndex        =   24
      Top             =   1995
      Width           =   660
   End
   Begin VB.CommandButton CmdElegir 
      Caption         =   ">>"
      Height          =   555
      Left            =   5475
      TabIndex        =   23
      Top             =   1305
      Width           =   660
   End
   Begin MSDataGridLib.DataGrid DtGTributos 
      Height          =   3690
      Left            =   6240
      TabIndex        =   21
      Top             =   1320
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   6509
      _Version        =   393216
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
   Begin VB.Frame FraOpciones 
      Height          =   7485
      Left            =   15
      TabIndex        =   6
      Top             =   945
      Width           =   1245
      Begin VB.CommandButton CmdHistoricoTributos 
         Caption         =   "Histórico  Tributos"
         Height          =   795
         Left            =   120
         TabIndex        =   19
         Top             =   2595
         Width           =   945
      End
      Begin VB.CommandButton CmdInforme 
         Caption         =   "Imprimir"
         Height          =   795
         Left            =   120
         Picture         =   "FrmTributosFiscales.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   150
         Picture         =   "FrmTributosFiscales.frx":1534
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6450
         Width           =   960
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   795
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   945
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar"
         Height          =   795
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   915
      Left            =   0
      Picture         =   "FrmTributosFiscales.frx":1976
      ScaleHeight     =   855
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRIBUTOS FISCALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   4530
         TabIndex        =   5
         Top             =   135
         Width           =   3255
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   4
         Top             =   495
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   2
         Top             =   525
         Width           =   2460
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   60
         TabIndex        =   1
         Top             =   495
         Width           =   1110
      End
   End
   Begin MSDataGridLib.DataGrid DtGTributosFiscales 
      Height          =   3720
      Left            =   1305
      TabIndex        =   11
      Top             =   1320
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   6562
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   -135
      TabIndex        =   12
      Top             =   8475
      Visible         =   0   'False
      Width           =   6555
      Begin VB.OptionButton Option3 
         Caption         =   "Bienes"
         Height          =   240
         Left            =   2445
         TabIndex        =   18
         Top             =   855
         Width           =   1620
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sueldos"
         Height          =   240
         Left            =   1380
         TabIndex        =   17
         Top             =   870
         Width           =   1620
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Viático"
         Height          =   240
         Left            =   420
         TabIndex        =   16
         Top             =   855
         Width           =   1620
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   1995
         TabIndex        =   15
         Top             =   495
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   375
         TabIndex        =   14
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label2 
         Caption         =   "Beneficiario"
         Height          =   195
         Left            =   315
         TabIndex        =   13
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Frame FraBusca 
      Height          =   2865
      Left            =   1275
      TabIndex        =   47
      Top             =   5565
      Width           =   10080
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar  Por Nro. Cheque"
         Height          =   420
         Left            =   7500
         TabIndex        =   61
         Top             =   660
         Width           =   2160
      End
      Begin VB.CommandButton CmdBuscaSolicitud 
         Caption         =   "Buscar por Nro. Sol."
         Height          =   405
         Left            =   7515
         TabIndex        =   60
         Top             =   1080
         Width           =   2145
      End
      Begin VB.TextBox TxtCodigoSolicitud 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   75
         TabIndex        =   51
         Top             =   570
         Width           =   1350
      End
      Begin VB.TextBox TxtNroCheque 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   1725
         Width           =   2100
      End
      Begin VB.TextBox TxtBeneficiario 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1455
         TabIndex        =   49
         Top             =   570
         Width           =   3930
      End
      Begin VB.TextBox TxtMonto_Bolivianos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2340
         TabIndex        =   48
         Top             =   1725
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc AdoCuenta 
         Height          =   360
         Left            =   2310
         Top             =   1695
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   635
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
         Caption         =   "AdoCuenta"
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
      Begin MSDataListLib.DataCombo DtCCuentaOrigen1 
         Bindings        =   "FrmTributosFiscales.frx":11660
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   105
         TabIndex        =   52
         Top             =   2385
         Visible         =   0   'False
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes1 
         Bindings        =   "FrmTributosFiscales.frx":11678
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   3855
         TabIndex        =   53
         Top             =   2400
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN1 
         Bindings        =   "FrmTributosFiscales.frx":11690
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   2235
         TabIndex        =   54
         Top             =   2400
         Visible         =   0   'False
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
         Bindings        =   "FrmTributosFiscales.frx":116A8
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   90
         TabIndex        =   63
         Top             =   1140
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "FrmTributosFiscales.frx":116C0
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   3855
         TabIndex        =   64
         Top             =   1140
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "FrmTributosFiscales.frx":116D8
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   2235
         TabIndex        =   65
         Top             =   1140
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. "
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   945
         Width           =   630
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. "
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   2175
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. de Cheque"
         Height          =   180
         Left            =   135
         TabIndex        =   58
         Top             =   1500
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Nro. Solicitud"
         Height          =   195
         Left            =   90
         TabIndex        =   57
         Top             =   330
         Width           =   1230
      End
      Begin VB.Label Label12 
         Caption         =   "Beneficiario"
         Height          =   195
         Left            =   1470
         TabIndex        =   56
         Top             =   360
         Width           =   2250
      End
      Begin VB.Label Label13 
         Caption         =   "Monto Bolivianos"
         Height          =   165
         Left            =   2310
         TabIndex        =   55
         Top             =   1515
         Width           =   1335
      End
   End
   Begin VB.Frame FraTributos 
      Height          =   2880
      Left            =   1275
      TabIndex        =   28
      Top             =   5550
      Width           =   10080
      Begin VB.CommandButton CmdAdicionar 
         Appearance      =   0  'Flat
         Caption         =   "Adicionar"
         Height          =   465
         Left            =   7095
         TabIndex        =   62
         Top             =   420
         Width           =   1380
      End
      Begin VB.TextBox TxtPorcentaje1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3075
         TabIndex        =   44
         Top             =   1395
         Width           =   1815
      End
      Begin VB.TextBox TxtPorcentaje2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3105
         TabIndex        =   43
         Top             =   1905
         Width           =   1815
      End
      Begin VB.TextBox TxtMonto 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   165
         TabIndex        =   40
         Top             =   1380
         Width           =   2445
      End
      Begin VB.TextBox TxtPartida 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   180
         TabIndex        =   39
         Top             =   1890
         Width           =   2445
      End
      Begin VB.TextBox TxtOrganismo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   165
         TabIndex        =   36
         Top             =   2430
         Width           =   2445
      End
      Begin VB.TextBox TxtCT 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3060
         TabIndex        =   35
         Top             =   870
         Width           =   2445
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   195
         TabIndex        =   33
         Top             =   870
         Width           =   2445
      End
      Begin VB.TextBox TxtBenef 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   165
         TabIndex        =   31
         Top             =   360
         Width           =   6375
      End
      Begin VB.CommandButton CmdSalirTributo 
         Caption         =   "Salir"
         Height          =   435
         Left            =   7110
         TabIndex        =   30
         Top             =   1335
         Width           =   1365
      End
      Begin VB.CommandButton CmdGrabarTributo 
         Appearance      =   0  'Flat
         Caption         =   "Grabar"
         Height          =   420
         Left            =   7110
         TabIndex        =   29
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label18 
         Caption         =   "Porcentaje 1"
         Height          =   240
         Left            =   3060
         TabIndex        =   46
         Top             =   1185
         Width           =   1590
      End
      Begin VB.Label Label17 
         Caption         =   "Porcentaje 2"
         Height          =   240
         Left            =   3090
         TabIndex        =   45
         Top             =   1710
         Width           =   825
      End
      Begin VB.Label Label16 
         Caption         =   "Monto"
         Height          =   240
         Left            =   135
         TabIndex        =   42
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label Label15 
         Caption         =   "Partida"
         Height          =   240
         Left            =   165
         TabIndex        =   41
         Top             =   1695
         Width           =   825
      End
      Begin VB.Label Label14 
         Caption         =   "Organismo"
         Height          =   240
         Left            =   150
         TabIndex        =   38
         Top             =   2235
         Width           =   825
      End
      Begin VB.Label Label9 
         Caption         =   "Cheq/Transf."
         Height          =   240
         Left            =   3045
         TabIndex        =   37
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Cmpte."
         Height          =   240
         Left            =   180
         TabIndex        =   34
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Beneficiario"
         Height          =   240
         Left            =   150
         TabIndex        =   32
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Pagos Realizados"
      Height          =   270
      Left            =   1275
      TabIndex        =   22
      Top             =   1050
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "Tributos Fiscales"
      Height          =   210
      Left            =   6240
      TabIndex        =   20
      Top             =   990
      Width           =   2250
   End
End
Attribute VB_Name = "FrmTributosFiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTributo As New ADODB.Recordset
Dim rscuenta As New ADODB.Recordset
Dim rsPAgoDetalle As New ADODB.Recordset
Dim rspago As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset

Public MONTO As Double

Private Sub CmdBuscar_Click()
Dim codigo_solicitud As String
Dim MONTO As Double
MONTO = 0

    If DtCCuentaOrigen.Text = "" Then
        MsgBox "No existe número de cuenta ", vbCritical + vbDefaultButton1
        Exit Sub
    End If

    If TxtNroCheque.Text = "" Then
        MsgBox "No existe número de cheque", vbCritical + vbDefaultButton1
        Exit Sub
    End If

db.Execute "delete from to_PagosTributos"
    Set rsPAgoDetalle = New ADODB.Recordset
    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    rsPAgoDetalle.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & DtCCuentaOrigen.Text & "' and numero_cheque_trf= '" & TxtNroCheque.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPAgoDetalle.RecordCount > 0 Then
                Set rspago = New ADODB.Recordset
                If rspago.State = 1 Then rspago.Close
                rspago.Open "SELECT * FROM pagos WHERE ges_gestion='" & rsPAgoDetalle("ges_gestion") & "' and org_codigo= '" & rsPAgoDetalle("org_codigo") & "' and codigo_Pago= '" & rsPAgoDetalle("codigo_pago") & "' ", db, adOpenKeyset, adLockOptimistic
                If rspago.RecordCount > 0 And Not IsNull(rspago("codigo_solicitud")) Then
                    codigo_solicitud = rspago("codigo_solicitud")
                End If
                
                'Buscando todos con el mismo Nro. de solicitud
                Set DtGTributosFiscales.DataSource = rspago
                'While Not rsPago.EOF
                    Set rsVista = New ADODB.Recordset
                    
                    'rsvista.Open " SELECT  pago_detalle.*, pagos.*, fc_beneficiario.denominacion_beneficiario " &
                    rsVista.Open " SELECT fc_beneficiario.denominacion_beneficiario as BENEFICIARIO, pago_detalle.codigo_pago as CMPTE,Pago_detalle.org_codigo as ORGANISMO,pago_detalle.numero_cheque_trf AS CHEQTRANSF, pago_detalle.codigo_beneficiario as CODIGOBENEF, pago_detalle.monto_bolivianos as MONTO, pago_detalle.par_codigo as PARTIDA,pago_detalle.*, pagos.*  " & _
                                 " FROM (pago_detalle INNER JOIN pagos ON (pago_detalle.Ges_gestion = pagos.ges_gestion) AND (pago_detalle.org_codigo = pagos.org_codigo) AND (pago_detalle.codigo_pago = pagos.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario WHERE pagos.codigo_solicitud='" & codigo_solicitud & "' ", db, adOpenKeyset, adLockOptimistic
                    If rsVista.RecordCount > 0 Then
                            Set DtGTributosFiscales.DataSource = rsVista
                            rsVista.MoveFirst
                            While Not rsVista.EOF
                                    If rsVista("CODIGOBENEF") <> "035_SNII" Then
                                        txtbeneficiario.Text = rsVista("beneficiario")
                                    End If
                                    If rsVista("MONTO") > 0 Then
                                        MONTO = MONTO + rsVista("MONTO")
                                    End If
'                                   lstBeneficiario.AddItem rsVista("beneficiario")
                                   If Not IsNull(rsVista("CheqTransf")) Then
                                   Dim SQLVar As String
                                   SQLVar = "insert into to_PagosTributos (T_beneficiario, codigo_pago, org_codigo, Nro_doc, monto_bolivianos) " & _
                                            "values  ('" & rsVista!beneficiario & "','" & rsVista("cmpte") & "', '" & rsVista("organismo") & "','" & rsVista("CheqTransf") & "'," & CCur(rsVista("monto")) & " ) "
                                   db.Execute SQLVar
                                   End If
                                   rsVista.MoveNext
                            Wend
                    Else
                            MsgBox "No existen registros", vbInformation + vbCritical
                            Set DtGTributosFiscales.DataSource = rsNada
                    End If
                    TxtMonto_bolivianos = MONTO
    Else
        MsgBox "No existen registros", vbCritical + vbDefaultButton1, "VALIDACION DE DATOS"
        Set DtGTributosFiscales.DataSource = rsNada
        
    End If
End Sub

Private Sub CmdBuscarCmpte_Click()
    FraBusca.Visible = True
    FraTributos.Visible = False
End Sub

Private Sub CmdBuscaSolicitud_Click()
Dim codigo_solicitud As String
Dim beneficiario As String
Dim valor As String

MONTO = 0
If TxtCodigoSolicitud.Text = "" Then
    MsgBox "No existe dato de solicitud", vbInformation + vbCritical
    Exit Sub
End If
db.Execute "DELETE from to_PagosTributos "
       Set rsVista = New ADODB.Recordset
       rsVista.Open " SELECT fc_beneficiario.denominacion_beneficiario as BENEFICIARIO, pago_detalle.codigo_pago as CMPTE,Pago_detalle.org_codigo as ORGANISMO,pago_detalle.numero_cheque_trf AS CHEQTRANSF, pago_detalle.codigo_beneficiario as CODIGOBENEF, pago_detalle.monto_bolivianos as MONTO, pago_detalle.par_codigo as PARTIDA, pago_detalle.*, pagos.*  " & _
                 " FROM (pago_detalle INNER JOIN pagos ON (pago_detalle.Ges_gestion = pagos.ges_gestion) AND (pago_detalle.org_codigo = pagos.org_codigo) AND (pago_detalle.codigo_pago = pagos.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario WHERE pagos.codigo_solicitud='" & TxtCodigoSolicitud.Text & "' ", db, adOpenKeyset, adLockOptimistic
        If rsVista.RecordCount > 0 Then
        Set DtGTributosFiscales.DataSource = rsVista
        While Not rsVista.EOF
            If rsVista("CODIGOBENEF") <> "035_SNII" Then
                txtbeneficiario.Text = rsVista("beneficiario")
            End If
            If rsVista("MONTO") > 0 Then
                MONTO = MONTO + rsVista("MONTO")
            End If
            If IsNull(rsVista("CheqTransf")) Then
               rsVista("CheqTransf") = "0000"
            End If
            If IsNull(rsVista("monto")) Then
               rsVista("monto") = 0
            End If
            'LstBeneficiario.AddItem rsVista("Beneficiario")
            db.Execute "insert into to_PagosTributos (T_beneficiario, codigo_pago, org_codigo, Nro_doc, monto_bolivianos) " & _
                                                      "values  ('" & rsVista!beneficiario & "','" & rsVista("cmpte") & "', '" & rsVista("organismo") & "','" & rsVista("CheqTransf") & "'," & rsVista("monto") & " ) "
         rsVista.MoveNext
        Wend
        TxtMonto_bolivianos.Text = MONTO
        Else
          MsgBox "No existen registros", vbCritical + vbDefaultButton1, "VALIDACION DE DATOS"
        End If
                    
    
End Sub

Private Sub CmdDevolver_Click()
If TxtMonto_bolivianos.Text <> "" Then
    db.Execute "delete FROM to_CajaHistoricoPagos where codigo_pago='" & DtGTributos.Columns(1) & "' and org_codigo='" & DtGTributos.Columns(2) & "'"
    
    Set rsHistorico = New ADODB.Recordset
    rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos", db, adOpenKeyset, adLockOptimistic
    If rsHistorico.RecordCount > 0 Then
        Set DtGTributos.DataSource = rsHistorico
    Else
        Set DtGTributos.DataSource = rsNada
    End If
Else
   MsgBox "Elija un registro", vbInformation + vbCritical
End If
End Sub

Private Sub cmdElegir_Click()
'Pasando los datos a histórico y determinando porcentajes
'DtGTributosFiscales.Columns(5)= monto_bolivianos
'DtGTributosFiscales.Columns(6)= codigo de partido
Dim Porcentaje_1 As Double
Dim Porcentaje_2 As Double
If TxtMonto_bolivianos.Text <> "" Then
        MONTO = CDbl(TxtMonto_bolivianos.Text)
        
        Porcentaje_1 = 0
        Porcentaje_2 = 0
        If MONTO > 0 Then
            If DtGTributosFiscales.Columns(6) = "22200" Then
                Porcentaje_1 = MONTO * 0.13
            End If
            If DtGTributosFiscales.Columns(6) = "46200" Then
                Porcentaje_1 = MONTO * 0.125
                Porcentaje_2 = MONTO * 0.03
            End If
            If DtGTributosFiscales.Columns(6) = "43100" Or DtGTributosFiscales.Columns(6) = "43500" Or DtGTributosFiscales.Columns(6) = "43600" Or DtGTributosFiscales.Columns(6) = "43700" Then
                Porcentaje_1 = MONTO * 0.13
                Porcentaje_2 = MONTO * 0.03
            End If
        Else
            MsgBox "No existe monto bolivianos !!! Imposible pasar datos", vbInformation + vbCritical
            Exit Sub
        End If
        db.Execute "insert into to_CajaHistoricoPagos (T_beneficiario, codigo_pago, org_codigo, Nro_doc, monto_bolivianos, par_codigo, Porcentaje_1, Porcentaje_2 )" & _
                   "values  ('" & txtbeneficiario.Text & "','" & DtGTributosFiscales.Columns(1) & "', '" & DtGTributosFiscales.Columns(2) & "','" & DtGTributosFiscales.Columns(3) & "'," & MONTO & ", '" & DtGTributosFiscales.Columns(6) & "', " & Porcentaje_1 & ", " & Porcentaje_2 & "  ) "
                   
        Set rsHistorico = New ADODB.Recordset
        If rsHistorico.State = 1 Then rsHistorico.Close
        rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos", db, adOpenKeyset, adLockOptimistic
        If rsHistorico.RecordCount > 0 Then
           Set DtGTributos.DataSource = rsHistorico
        End If
 Else
    MsgBox "Elija un registro !!!", vbInformation + vbCritical
 End If
End Sub

Private Sub CmdGrabarTributo_Click()

Dim sino As String
sino = MsgBox("Está seguro grabar los datos modificados?", vbYesNo + vbQuestion, "Atenciòn")
If sino = vbYes Then
    If rsH.State = 1 Then rsHistorico.Close
    rsH.Open "SELECT * FROM to_CajaHistoricoPagos WHERE codigo_pago ='" & DtGTributos.Columns(1) & "' and org_codigo= TxtOrganismo.Text = '" & DtGTributos.Columns(2) & " '", db, adOpenKeyset, adLockOptimistic
    If rsH.RecordCount > 0 Then
        rsH.Update
    End If
End If
End Sub

Private Sub CmdGuardarHistorico_Click()
Dim sino As String
Dim rsHistorico As New ADODB.Recordset
If rsHistorico.State = 1 Then rsHistorico.Close
rsHistorico.Open "select * from to_CajaHistoricoPagos", db, adOpenKeyset, adLockOptimistic
If rsHistorico.RecordCount > 0 Then
    sino = MsgBox("Está seguro en el histórico el grid que se ve, se grabarán con la fecha de hoy?", vbYesNo + vbQuestion, "Atenciòn")
    If sino = vbYes Then
        Set rsHistorico = New ADODB.Recordset
        If rsHistorico.State = 1 Then rsHistorico.Close
        rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos", db, adOpenKeyset, adLockOptimistic
        If rsHistorico.RecordCount > 0 Then
            While Not rsHistorico.EOF
                    db.Execute "insert into to_CajaHistoricoPagos_fechas (T_beneficiario, codigo_pago, org_codigo, Nro_doc, monto_bolivianos, porcentaje_1, porcentaje_2, fecha_impresion) " & _
                                                                      "values  ('" & rsHistorico!T_beneficiario & "','" & rsHistorico!codigo_pago & "', '" & rsHistorico!org_codigo & "','" & rsHistorico!Nro_doc & "'," & rsHistorico!monto_Bolivianos & " , " & rsHistorico!Porcentaje_1 & ", " & rsHistorico!Porcentaje_2 & ", '" & Date & "') "
                    rsHistorico.MoveNext
                    
            Wend
        End If
                    'db.Execute "delete from to_CajaHistoricoPagos"
    End If
Else
    MsgBox "No existen regsitros para vaciar al histórico", vbCritical + vbDefaultButton1, "Validación de datos"
End If
End Sub

Private Sub CmdHistoricoTributos_Click()
    FrmHistoricoTributosFiscales.Show
    FraTributos.Visible = True
    
End Sub

Private Sub CmdImprimirGridSeleccionados_Click()
 'RepSoloTributos.Show
 CryHistorico.ReportFileName = App.Path & "\FormsTesoreria\EntregaCheques\Impresiones\Rpt_SoloTributos.rpt"
 iResult = CryHistorico.PrintReport
 If iResult <> 0 Then
    MsgBox CryHistorico.LastErrorNumber & " : " & CryHistorico.LastErrorString, vbCritical + vbOKOnly, "Error..."
 End If
End Sub

Private Sub CmdInforme_Click()
Dim rsTributos As New ADODB.Recordset
     If rsTributos.State = 1 Then rsTributos.Close
     rsTributos.Open "select * from to_PagosTributos", db, adOpenKeyset, adLockOptimistic
     If rsTributos.RecordCount > 0 Then
            CryPagosBeneficiario.ReportFileName = App.Path & "\FormsTesoreria\EntregaCheques\Impresiones\Rpt_PagosBeneficiario.rpt"
            iResult = CryPagosBeneficiario.PrintReport
            If iResult <> 0 Then
                MsgBox CryPagosBeneficiario.LastErrorNumber & " : " & CryPagosBeneficiario.LastErrorString, vbCritical + vbOKOnly, "Error..."
            End If
     Else
          MsgBox "No existen datos para imprimir, selecciónelos", vbCritical + vbDefaultButton1, "Validación de datos"
    End If
End Sub

Private Sub CmdLimpiarGrid_Click()
Dim sino As String
sino = MsgBox("Está seguro limpiar el grid?", vbYesNo + vbQuestion, "Atenciòn")
If sino = vbYes Then
    db.Execute "delete from to_CajaHistoricoPagos"
End If

Set rsHistorico = New ADODB.Recordset
If rsHistorico.State = 1 Then rsHistorico.Close
rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos", db, adOpenKeyset, adLockOptimistic
If rsHistorico.RecordCount > 0 Then
   Set DtGTributos.DataSource = rsHistorico
Else
   Set DtGTributos.DataSource = rsNada
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub CmdSalirBuscar_Click()
    FraBusca.Visible = False
    FraTributos.Visible = True
End Sub

Private Sub CmdSalirTributo_Click()
    FraTributos.Visible = False
End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
   DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
   DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
End Sub

Private Sub DtGTributos_Click()
    TxtBenef.Text = DtGTributos.Columns(0)
    TxtCmpte.Text = DtGTributos.Columns(1)
    txtorganismo.Text = DtGTributos.Columns(2)
    TxtCT.Text = DtGTributos.Columns(3)
    txtmonto.Text = DtGTributos.Columns(4)
    txtpartida.Text = DtGTributos.Columns(5)
    TxtPorcentaje1.Text = DtGTributos.Columns(6)
    TxtPorcentaje2.Text = DtGTributos.Columns(7)
End Sub

Private Sub DtGTributosFiscales_Click()
'If DtGTributosFiscales.Columns(4) = "035_SNII" Then
If DtGTributosFiscales.Columns(88) <> "" Then
    TxtJustificacion.Text = DtGTributosFiscales.Columns(88)
Else
    TxtJustificacion.Text = ""
End If
'End If
End Sub

Private Sub Form_Load()
'Abriendo la tabla de cuentas bancarias
    Set rscuenta = New ADODB.Recordset
    rscuenta.Open "select * from fc_cuenta_bancaria order by cta_codigo", db, adOpenKeyset, adLockOptimistic
    Set AdoCuenta.Recordset = rscuenta
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
'Abriendo histórico
    Set rsHistorico = New ADODB.Recordset
    If rsHistorico.State = 1 Then rsHistorico.Close
    rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos", db, adOpenKeyset, adLockOptimistic
    If rsHistorico.RecordCount > 0 Then
       Set DtGTributos.DataSource = rsHistorico
    Else
       Set DtGTributos.DataSource = rsNada
    End If
'Dejando limpio la tabla temporal    }
db.Execute "DELETE FROM to_PagosTributos"

	Call SeguridadSet(Me)
End Sub

Private Sub TxtCodigoSolicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmAnulacion 
   Caption         =   "Anulaciones y Devoluciones de Comprobantes - Gesti�n 2000"
   ClientHeight    =   8310
   ClientLeft      =   135
   ClientTop       =   285
   ClientWidth     =   11880
   Icon            =   "FrmAnulacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   480
      Left            =   1185
      TabIndex        =   160
      Top             =   1005
      Width           =   3585
      Begin VB.OptionButton OptFilGral2 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   162
         Top             =   150
         Width           =   795
      End
      Begin VB.OptionButton OptFilGral1 
         Caption         =   "Sin Procesar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   161
         Top             =   210
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame FraCopiaRegistro 
      Height          =   5985
      Left            =   4740
      TabIndex        =   105
      Top             =   990
      Visible         =   0   'False
      Width           =   7125
      Begin VB.Frame FraCopiar 
         Height          =   5745
         Left            =   60
         TabIndex        =   124
         Top             =   120
         Width           =   6996
         Begin VB.TextBox TxtFR 
            DataSource      =   "AdoRegularizacion"
            Enabled         =   0   'False
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
            Height          =   315
            Left            =   180
            TabIndex        =   133
            Top             =   2070
            Width           =   1395
         End
         Begin VB.Frame Frame6 
            Height          =   120
            Left            =   30
            TabIndex        =   132
            Top             =   495
            Width           =   6870
         End
         Begin VB.TextBox TxtCO 
            DataSource      =   "AdoRegularizacion"
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
            Height          =   315
            Left            =   5430
            TabIndex        =   131
            Top             =   1440
            Width           =   1395
         End
         Begin VB.TextBox TxtNCA 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataSource      =   "AdoRegularizacion"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   5415
            TabIndex        =   130
            Top             =   180
            Width           =   1395
         End
         Begin VB.Frame Frame7 
            Height          =   120
            Left            =   45
            TabIndex        =   129
            Top             =   4635
            Width           =   6855
         End
         Begin VB.TextBox TxtNS 
            DataSource      =   "AdoRegularizacion"
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
            Height          =   315
            Left            =   5430
            TabIndex        =   128
            Top             =   840
            Width           =   1380
         End
         Begin VB.TextBox TxtNC 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   127
            Top             =   195
            Width           =   1395
         End
         Begin VB.TextBox TxtJ 
            Height          =   720
            Left            =   156
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   126
            Text            =   "FrmAnulacion.frx":038A
            Top             =   4935
            Width           =   6660
         End
         Begin VB.TextBox TxtTR 
            DataSource      =   "AdoRegularizacion"
            Height          =   330
            Left            =   1905
            TabIndex        =   125
            Text            =   "Text1"
            Top             =   780
            Width           =   915
         End
         Begin MSDataListLib.DataCombo DtCDRD 
            Bindings        =   "FrmAnulacion.frx":0390
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   915
            TabIndex        =   134
            Top             =   1455
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Denominacion_documento"
            BoundColumn     =   "codigo_documento"
            Text            =   "DataCombo2"
         End
         Begin MSDataListLib.DataCombo DtCDR 
            Bindings        =   "FrmAnulacion.frx":03AB
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   165
            TabIndex        =   135
            Top             =   1455
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Codigo_Documento"
            BoundColumn     =   "codigo_documento"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DtCOF 
            Bindings        =   "FrmAnulacion.frx":03C6
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   180
            TabIndex        =   136
            Top             =   3705
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Org_Codigo"
            BoundColumn     =   "Org_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcC 
            Bindings        =   "FrmAnulacion.frx":03E1
            DataSource      =   "AdoRegularizacion"
            Height          =   288
            Left            =   180
            TabIndex        =   137
            Top             =   4320
            Width           =   1416
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "codigo_categoria"
            BoundColumn     =   "codigo_categoria"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCFF 
            Bindings        =   "FrmAnulacion.frx":03FC
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   195
            TabIndex        =   138
            Top             =   3180
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Fte_codigo"
            BoundColumn     =   "Fte_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcOFD 
            Bindings        =   "FrmAnulacion.frx":0414
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   1755
            TabIndex        =   139
            Top             =   3705
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Org_descripcion"
            BoundColumn     =   "Org_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcCD 
            Bindings        =   "FrmAnulacion.frx":042F
            DataSource      =   "AdoRegularizacion"
            Height          =   312
            Left            =   1632
            TabIndex        =   140
            Top             =   4308
            Width           =   3852
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "denominacion_categoria"
            BoundColumn     =   "codigo_categoria"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcFFD 
            Bindings        =   "FrmAnulacion.frx":044B
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   1740
            TabIndex        =   141
            Top             =   3195
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Fte_descripcion_larga"
            BoundColumn     =   "Fte_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCUT 
            Bindings        =   "FrmAnulacion.frx":0463
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   195
            TabIndex        =   142
            Top             =   2595
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Uni_codigo"
            BoundColumn     =   "Uni_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCUTD 
            Bindings        =   "FrmAnulacion.frx":047B
            DataSource      =   "AdoRegularizacion"
            Height          =   315
            Left            =   1740
            TabIndex        =   143
            Top             =   2610
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Uni_descripcion_larga"
            BoundColumn     =   "Uni_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcConv2 
            Bindings        =   "FrmAnulacion.frx":0493
            DataSource      =   "AdoRegularizacion"
            Height          =   312
            Left            =   5520
            TabIndex        =   157
            Top             =   4308
            Width           =   1296
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "codigo_convenio"
            BoundColumn     =   "codigo_categoria"
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "..."
            Height          =   270
            Left            =   3180
            TabIndex        =   158
            Top             =   615
            Width           =   555
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Formulario de Solicitud:"
            Height          =   195
            Left            =   3405
            TabIndex        =   155
            Top             =   885
            Width           =   1980
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Justificaci�n:"
            Height          =   195
            Left            =   315
            TabIndex        =   154
            Top             =   4740
            Width           =   915
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Unidad T�cnica:"
            Height          =   195
            Left            =   195
            TabIndex        =   153
            Top             =   2400
            Width           =   1185
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Organismo Financiador:"
            Height          =   195
            Left            =   210
            TabIndex        =   152
            Top             =   3480
            Width           =   1665
         End
         Begin VB.Label LblNroComprobanteAnt_Sig 
            AutoSize        =   -1  'True
            Caption         =   "Nro Comprobante anterior:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3480
            TabIndex        =   151
            Top             =   210
            Width           =   1860
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Documento Respaldo:"
            Height          =   195
            Left            =   165
            TabIndex        =   150
            Top             =   1215
            Width           =   1590
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "No.:"
            Height          =   195
            Left            =   5445
            TabIndex        =   149
            Top             =   1185
            Width           =   300
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Registro:"
            Height          =   195
            Left            =   180
            TabIndex        =   148
            Top             =   1860
            Width           =   1125
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Nro Comprobante:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   135
            TabIndex        =   147
            Top             =   210
            Width           =   1290
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Fuente Financiamiento:"
            Height          =   195
            Left            =   210
            TabIndex        =   146
            Top             =   2985
            Width           =   1650
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Categor�a del Financiador:"
            Height          =   195
            Left            =   195
            TabIndex        =   145
            Top             =   4050
            Width           =   1875
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Registro:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   150
            TabIndex        =   144
            Top             =   855
            Width           =   1215
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoRegularizacion 
      Height          =   375
      Left            =   1260
      Top             =   6360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
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
      Caption         =   "Navegar"
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
   Begin VB.Frame FraDetalle 
      Caption         =   "DETALLE REGULARIZACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1932
      Left            =   1170
      TabIndex        =   19
      Top             =   6765
      Width           =   10770
      Begin MSDataGridLib.DataGrid DtGDetalle 
         Bindings        =   "FrmAnulacion.frx":04AE
         Height          =   1245
         Left            =   180
         TabIndex        =   82
         Top             =   240
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   2196
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "par_codigo"
            Caption         =   "PARTIDA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "pro_programa"
            Caption         =   "PRO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "pro_subprograma"
            Caption         =   "SUB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "pro_proyecto"
            Caption         =   "PRY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "pro_actividad"
            Caption         =   "ACT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "cta_codigo"
            Caption         =   "CTA.ORIGEN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "numero_cheque_trf"
            Caption         =   "Cheque/Trf"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "cta_codigo_destino"
            Caption         =   "CTA.DESTINO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "codigo_beneficiario"
            Caption         =   "BENEFICIARIO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "monto_total"
            Caption         =   "MONTO Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "tipo_cambio"
            Caption         =   "TDC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "monto_dolares"
            Caption         =   "MONTO $US"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               DividerStyle    =   6
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoDetalle 
         Height          =   336
         Left            =   120
         Top             =   1500
         Width           =   10560
         _ExtentX        =   18627
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
         Enabled         =   0
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoDetalle"
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   11820
      TabIndex        =   13
      Top             =   0
      Width           =   11880
      Begin Crystal.CrystalReport Cry 
         Left            =   240
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "."
         ForeColor       =   &H0000C000&
         Height          =   180
         Left            =   4815
         TabIndex        =   103
         Top             =   675
         Width           =   2655
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE COMPROBANTES"
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
         Left            =   3690
         TabIndex        =   14
         Top             =   150
         Width           =   4935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   225
         Left            =   10485
         TabIndex        =   18
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   17
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   16
         Top             =   705
         Width           =   2460
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Left            =   60
         TabIndex        =   15
         Top             =   675
         Width           =   1110
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -15
         Picture         =   "FrmAnulacion.frx":04C7
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11900
      End
   End
   Begin MSDataGridLib.DataGrid DtcRegularizacion 
      Bindings        =   "FrmAnulacion.frx":137B
      Height          =   4830
      Left            =   1200
      TabIndex        =   107
      Top             =   1530
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   8520
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "codigo_pago"
         Caption         =   "Cmbte."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "org_codigo"
         Caption         =   "Org."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "tipo_comp"
         Caption         =   "Tipo"
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
      BeginProperty Column03 
         DataField       =   "Nro_comprobante_anterior"
         Caption         =   "Anterior"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "estado_compromiso"
         Caption         =   "C"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "estado_devengado"
         Caption         =   "D"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "estado_pagado"
         Caption         =   "P"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "estado_reversion_total"
         Caption         =   "R"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "estado_devolucion"
         Caption         =   "V"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "estado_anulado"
         Caption         =   "A"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "fecha_egreso"
         Caption         =   "Fecha Cmbte."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "estado_contabilidad"
         Caption         =   "CB"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   269.858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   299.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   345.26
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraMaestro 
      Enabled         =   0   'False
      Height          =   5745
      Left            =   4800
      TabIndex        =   20
      Top             =   960
      Width           =   7035
      Begin VB.TextBox TxtTipoReg 
         DataField       =   "tipo_formulario"
         DataSource      =   "AdoRegularizacion"
         Height          =   330
         Left            =   1455
         TabIndex        =   159
         Text            =   "Text1"
         Top             =   810
         Width           =   915
      End
      Begin MSDataListLib.DataCombo DtcTipoCod 
         Bindings        =   "FrmAnulacion.frx":139B
         DataField       =   "tipo_formulario"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   2400
         TabIndex        =   109
         Top             =   1155
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_tipo"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox DtpFecha 
         DataField       =   "fecha_egreso"
         DataSource      =   "AdoRegularizacion"
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
         Left            =   180
         TabIndex        =   3
         Top             =   2005
         Width           =   1545
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   75
         TabIndex        =   99
         Top             =   465
         Width           =   6930
      End
      Begin MSDataListLib.DataCombo DtcDcuDes 
         Bindings        =   "FrmAnulacion.frx":13B1
         DataField       =   "codigo_documento"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   900
         TabIndex        =   91
         Top             =   1470
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_documento"
         BoundColumn     =   "codigo_documento"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcDcu 
         Bindings        =   "FrmAnulacion.frx":13CC
         DataField       =   "codigo_documento"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   1470
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Codigo_Documento"
         BoundColumn     =   "codigo_documento"
         Text            =   ""
      End
      Begin VB.TextBox TxtCodigoOrden 
         DataField       =   "codigo_orden"
         DataSource      =   "AdoRegularizacion"
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
         Height          =   315
         Left            =   5490
         TabIndex        =   2
         Top             =   1470
         Width           =   1395
      End
      Begin VB.TextBox TxtComprobanteAnterior 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Nro_Comprobante_Anterior"
         DataSource      =   "AdoRegularizacion"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   5490
         TabIndex        =   39
         Top             =   150
         Width           =   1395
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   30
         TabIndex        =   37
         Top             =   4650
         Width           =   6990
      End
      Begin VB.TextBox txtNroSolicitud 
         DataField       =   "codigo_solicitud"
         DataSource      =   "AdoRegularizacion"
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
         Height          =   315
         Left            =   5730
         TabIndex        =   0
         Top             =   825
         Width           =   1140
      End
      Begin VB.TextBox TxtComprobante 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "codigo_pago"
         DataSource      =   "AdoRegularizacion"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1470
         TabIndex        =   21
         Top             =   150
         Width           =   1395
      End
      Begin VB.TextBox TxtJustificacion 
         DataField       =   "justificacion"
         DataSource      =   "AdoRegularizacion"
         Height          =   720
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "FrmAnulacion.frx":13E7
         Top             =   4950
         Width           =   6750
      End
      Begin MSDataListLib.DataCombo DtcOrg 
         Bindings        =   "FrmAnulacion.frx":13ED
         DataField       =   "Org_codigo"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   3690
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Org_Codigo"
         BoundColumn     =   "Org_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCat 
         Bindings        =   "FrmAnulacion.frx":1408
         DataField       =   "codigo_categoria"
         DataSource      =   "AdoRegularizacion"
         Height          =   312
         Left            =   180
         TabIndex        =   7
         Top             =   4320
         Width           =   1476
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_categoria"
         BoundColumn     =   "codigo_categoria"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DTcFte 
         Bindings        =   "FrmAnulacion.frx":1423
         DataField       =   "fte_codigo"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   3145
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Fte_codigo"
         BoundColumn     =   "Fte_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcDesOrg 
         Bindings        =   "FrmAnulacion.frx":143B
         DataField       =   "Org_codigo"
         DataSource      =   "AdoRegularizacion"
         Height          =   288
         Left            =   1716
         TabIndex        =   11
         Top             =   3708
         Width           =   5172
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Org_descripcion"
         BoundColumn     =   "Org_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCatDes 
         Bindings        =   "FrmAnulacion.frx":1456
         DataField       =   "codigo_categoria"
         DataSource      =   "AdoRegularizacion"
         Height          =   312
         Left            =   1776
         TabIndex        =   12
         Top             =   4320
         Width           =   3672
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_categoria"
         BoundColumn     =   "codigo_categoria"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcFteDes 
         Bindings        =   "FrmAnulacion.frx":1472
         DataField       =   "fte_codigo"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   1710
         TabIndex        =   10
         Top             =   3145
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Fte_descripcion_larga"
         BoundColumn     =   "Fte_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCUnidad 
         Bindings        =   "FrmAnulacion.frx":148A
         DataField       =   "codigo_unidad"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   2580
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Uni_codigo"
         BoundColumn     =   "Uni_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDesUnidad 
         Bindings        =   "FrmAnulacion.frx":14A2
         DataField       =   "codigo_unidad"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   2580
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Uni_descripcion_larga"
         BoundColumn     =   "Uni_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcTipoDes 
         Bindings        =   "FrmAnulacion.frx":14BA
         DataField       =   "tipo_formulario"
         DataSource      =   "AdoRegularizacion"
         Height          =   315
         Left            =   1380
         TabIndex        =   110
         Top             =   795
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_tipo"
         BoundColumn     =   "codigo_tipo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcConv 
         Bindings        =   "FrmAnulacion.frx":14D0
         DataField       =   "codigo_categoria"
         DataSource      =   "AdoRegularizacion"
         Height          =   312
         Left            =   5460
         TabIndex        =   156
         Top             =   4320
         Width           =   1476
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_convenio"
         BoundColumn     =   "codigo_categoria"
         Text            =   ""
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Registro:"
         Height          =   195
         Left            =   165
         TabIndex        =   108
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Formulario Solicitud:"
         Height          =   195
         Left            =   3945
         TabIndex        =   42
         Top             =   885
         Width           =   1755
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Justificaci�n:"
         Height          =   195
         Left            =   315
         TabIndex        =   41
         Top             =   4740
         Width           =   915
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Unidad T�cnica:"
         Height          =   195
         Left            =   165
         TabIndex        =   38
         Top             =   2350
         Width           =   1185
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Organismo Financiador:"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   3480
         Width           =   1665
      End
      Begin VB.Label LblNroCmpte_Ant_Dev 
         AutoSize        =   -1  'True
         Caption         =   "Nro Comprobante anterior:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3555
         TabIndex        =   40
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Documento Respaldo:"
         Height          =   195
         Left            =   165
         TabIndex        =   36
         Top             =   1215
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No.:"
         Height          =   195
         Left            =   5445
         TabIndex        =   27
         Top             =   1185
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Registro:"
         Height          =   195
         Left            =   165
         TabIndex        =   26
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Nro Comprobante:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fuente Financiamiento:"
         Height          =   195
         Left            =   165
         TabIndex        =   23
         Top             =   2935
         Width           =   1650
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Categor�a del Financiador:"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   4050
         Width           =   1875
      End
   End
   Begin VB.Frame FraDetalleG 
      Enabled         =   0   'False
      Height          =   5820
      Left            =   1152
      TabIndex        =   43
      Top             =   915
      Visible         =   0   'False
      Width           =   10710
      Begin VB.Frame FraProyecto 
         Height          =   2355
         Left            =   720
         TabIndex        =   92
         Top             =   840
         Visible         =   0   'False
         Width           =   8925
         Begin VB.CommandButton CmdSalirGrid 
            Caption         =   "Salir sin Elegir Proyecto ..."
            Height          =   330
            Left            =   6135
            TabIndex        =   114
            Top             =   1950
            Width           =   2490
         End
         Begin MSDataGridLib.DataGrid DtGProyecto 
            Bindings        =   "FrmAnulacion.frx":14EB
            Height          =   1620
            Left            =   255
            TabIndex        =   112
            Top             =   240
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   2858
            _Version        =   393216
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "pro_programa"
               Caption         =   "Programa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "pro_subprograma"
               Caption         =   "Subprograma"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "pro_proyecto"
               Caption         =   "Proyecto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Pro_actividad"
               Caption         =   "Actividad"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Pro_descripcion_larga"
               Caption         =   "Descripci�n"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
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
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label9 
            Caption         =   "Doble Click para Elegir Proyecto ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Left            =   255
            TabIndex        =   113
            Top             =   1995
            Width           =   3345
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cta.Bancaria:"
         Height          =   1335
         Left            =   360
         TabIndex        =   115
         Top             =   4380
         Width           =   8145
         Begin VB.TextBox Text2 
            DataField       =   "beneficiario_destino"
            DataSource      =   "AdoDetalle"
            Height          =   315
            Left            =   4440
            TabIndex        =   123
            Text            =   "Text1"
            Top             =   570
            Width           =   3570
         End
         Begin VB.TextBox Text1 
            DataField       =   "cta_codigo_destino"
            DataSource      =   "AdoDetalle"
            Height          =   315
            Left            =   5640
            TabIndex        =   122
            Text            =   "Text1"
            Top             =   195
            Width           =   1410
         End
         Begin VB.TextBox TxtNroCheque 
            DataField       =   "numero_cheque_trf"
            DataSource      =   "AdoDetalle"
            Height          =   315
            Left            =   2805
            TabIndex        =   118
            Text            =   "Text1"
            Top             =   930
            Width           =   1410
         End
         Begin MSDataListLib.DataCombo DtcDesCta 
            Bindings        =   "FrmAnulacion.frx":1505
            DataField       =   "cta_codigo"
            DataSource      =   "AdoDetalle"
            Height          =   315
            Left            =   120
            TabIndex        =   117
            Top             =   550
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "cta_descripcion_larga"
            BoundColumn     =   "cta_codigo"
            Text            =   "DataCombo2"
         End
         Begin MSDataListLib.DataCombo DtcCodCta 
            Bindings        =   "FrmAnulacion.frx":151D
            DataField       =   "cta_codigo"
            DataSource      =   "AdoDetalle"
            Height          =   315
            Left            =   1200
            TabIndex        =   116
            Top             =   195
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "cta_codigo"
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "DESTINO . . . :"
            Height          =   195
            Left            =   4455
            TabIndex        =   121
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "ORIGEN . . . :"
            Height          =   195
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Cheque o Transferencia:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   119
            Top             =   960
            Width           =   2580
         End
      End
      Begin VB.TextBox DtCCuentaDestino 
         DataField       =   "cta_codigo_destino"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   5100
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   2190
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox TxtMontoDolares 
         DataField       =   "monto_Dolares"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   1830
         TabIndex        =   95
         Top             =   4050
         Width           =   1395
      End
      Begin VB.TextBox TxtTipoCambio 
         DataField       =   "tipo_cambio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   4410
         TabIndex        =   94
         Top             =   4110
         Width           =   1410
      End
      Begin VB.CommandButton CmdNuevoBeneficiario 
         Caption         =   "Nuevo Beneficiario"
         Height          =   495
         Left            =   8745
         TabIndex        =   93
         Top             =   2805
         Width           =   975
      End
      Begin VB.CommandButton CmdProyecto 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Elegir Proyecto ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7575
         TabIndex        =   90
         Top             =   1170
         Width           =   1815
      End
      Begin VB.CommandButton CmdCalculo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "Calcular"
         Height          =   855
         Left            =   8745
         Picture         =   "FrmAnulacion.frx":1535
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3525
         Width           =   990
      End
      Begin VB.TextBox TxtMontoFuente 
         DataField       =   "monto_total"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   1845
         TabIndex        =   84
         Top             =   3555
         Width           =   1410
      End
      Begin VB.TextBox TxtSaldo 
         DataField       =   "saldo_bolivianos"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoDetalle"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         TabIndex        =   78
         Top             =   3570
         Width           =   1395
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   180
         TabIndex        =   75
         Top             =   2190
         Width           =   2775
         Begin VB.OptionButton OptTransferenciaOrigen 
            Caption         =   "    Transferencia"
            Height          =   195
            Left            =   1170
            TabIndex        =   77
            Top             =   30
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.OptionButton OptChequeOrigen 
            Caption         =   "    Cheque"
            Height          =   195
            Left            =   90
            TabIndex        =   76
            Top             =   30
            Visible         =   0   'False
            Width           =   1035
         End
      End
      Begin VB.TextBox TxtNoTransferenciaOrigen 
         DataField       =   "numero_cheque_trf"
         DataSource      =   "AdoDetalle"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4650
         TabIndex        =   74
         Top             =   2190
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox TxtDeducciones 
         DataField       =   "deducciones"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   4410
         TabIndex        =   70
         Top             =   3585
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TxtProy 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1665
         TabIndex        =   69
         Top             =   1635
         Width           =   7755
      End
      Begin VB.TextBox TxtProgramad 
         DataField       =   "Pro_programa"
         DataSource      =   "AdoDetalle"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2025
         TabIndex        =   68
         Top             =   1350
         Width           =   675
      End
      Begin VB.TextBox TxtSubprogramad 
         DataField       =   "Pro_subprograma"
         DataSource      =   "AdoDetalle"
         Enabled         =   0   'False
         Height          =   240
         Left            =   3315
         TabIndex        =   67
         Top             =   1350
         Width           =   675
      End
      Begin VB.TextBox TxtProyectod 
         DataField       =   "Pro_proyecto"
         DataSource      =   "AdoDetalle"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4545
         TabIndex        =   66
         Top             =   1350
         Width           =   675
      End
      Begin VB.TextBox TxtActividadd 
         DataField       =   "Pro_actividad"
         DataSource      =   "AdoDetalle"
         Enabled         =   0   'False
         Height          =   240
         Left            =   5700
         TabIndex        =   65
         Top             =   1350
         Width           =   675
      End
      Begin VB.Frame Frame10 
         Height          =   30
         Left            =   600
         TabIndex        =   60
         Top             =   2550
         Width           =   9030
      End
      Begin VB.Frame Frame9 
         Height          =   30
         Left            =   570
         TabIndex        =   58
         Top             =   2055
         Width           =   9030
      End
      Begin VB.TextBox TxtCodigoDetalle 
         Appearance      =   0  'Flat
         DataField       =   "codigo_pago_detalle"
         DataSource      =   "AdoDetalle"
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
         Height          =   315
         Left            =   7560
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.TextBox TxtCodigoOrdend 
         Appearance      =   0  'Flat
         DataField       =   "codigo_pago"
         DataSource      =   "AdoDetalle"
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   3930
         TabIndex        =   51
         Top             =   240
         Visible         =   0   'False
         Width           =   1875
      End
      Begin MSDataListLib.DataCombo dtcNombreRuc 
         Bindings        =   "FrmAnulacion.frx":173F
         DataField       =   "codigo_beneficiario"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   2430
         TabIndex        =   44
         Top             =   2925
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dtcRuc 
         Bindings        =   "FrmAnulacion.frx":1754
         DataField       =   "codigo_beneficiario"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   660
         TabIndex        =   45
         Top             =   2925
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtCPartida 
         Bindings        =   "FrmAnulacion.frx":1769
         DataField       =   "par_codigo"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   1605
         TabIndex        =   54
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "par_codigo"
         BoundColumn     =   "Par_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCPartidaDes 
         Bindings        =   "FrmAnulacion.frx":17A1
         DataField       =   "par_codigo"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   2955
         TabIndex        =   55
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Par_descripcion_larga"
         BoundColumn     =   "par_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "FrmAnulacion.frx":17BA
         DataField       =   "cta_codigo"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   3015
         TabIndex        =   72
         Top             =   2190
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "FrmAnulacion.frx":17D2
         DataField       =   "cta_codigo"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   3945
         TabIndex        =   73
         Top             =   2190
         Visible         =   0   'False
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "FrmAnulacion.frx":17EA
         DataField       =   "cta_codigo"
         DataSource      =   "AdoDetalle"
         Height          =   315
         Left            =   3570
         TabIndex        =   102
         Top             =   2190
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Frame Frame8 
         Height          =   120
         Left            =   15
         TabIndex        =   57
         Top             =   15
         Width           =   9045
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   1140
         TabIndex        =   101
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Monto D�lares:"
         Height          =   195
         Left            =   630
         TabIndex        =   97
         Top             =   4110
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio:"
         Height          =   195
         Left            =   3450
         TabIndex        =   96
         Top             =   4155
         Width           =   930
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Monto en Bs.:"
         Height          =   195
         Left            =   630
         TabIndex        =   85
         Top             =   3615
         Width           =   990
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Categor�a Program�tica:"
         Height          =   390
         Left            =   570
         TabIndex        =   81
         Top             =   1275
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Pago:"
         Height          =   195
         Left            =   2970
         TabIndex        =   80
         Top             =   285
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "L�quido Pagable:"
         Height          =   195
         Left            =   5925
         TabIndex        =   79
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Deducciones:"
         Height          =   195
         Left            =   3345
         TabIndex        =   71
         Top             =   3630
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Monto $us.:"
         Height          =   195
         Left            =   2775
         TabIndex        =   59
         Top             =   5895
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Partida:"
         Height          =   195
         Left            =   1005
         TabIndex        =   56
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Correlativo Detalle:"
         Height          =   195
         Left            =   6180
         TabIndex        =   53
         Top             =   300
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "SubPrograma"
         Height          =   195
         Index           =   3
         Left            =   3165
         TabIndex        =   50
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Proyecto"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   49
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Programa"
         Height          =   195
         Index           =   1
         Left            =   2055
         TabIndex        =   48
         Top             =   1125
         Width           =   675
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         Height          =   195
         Index           =   0
         Left            =   5685
         TabIndex        =   47
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Beneficiario:"
         Height          =   195
         Left            =   660
         TabIndex        =   46
         Top             =   2640
         Width           =   870
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   7752
      Left            =   30
      TabIndex        =   28
      Top             =   990
      Width           =   1140
      Begin VB.CommandButton CmdAprueba 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Aprobar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   120
         Picture         =   "FrmAnulacion.frx":1802
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   4245
         Width           =   945
      End
      Begin VB.CommandButton CmdCopiar 
         Caption         =   "Copiar Reg"
         Enabled         =   0   'False
         Height          =   660
         Left            =   120
         Picture         =   "FrmAnulacion.frx":1A0C
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1590
         Width           =   945
      End
      Begin VB.CommandButton CmdAdicionar 
         Caption         =   "Adicionar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "FrmAnulacion.frx":1C16
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   120
         Picture         =   "FrmAnulacion.frx":1F20
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   930
         Width           =   945
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   660
         Left            =   120
         Picture         =   "FrmAnulacion.frx":212A
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2910
         Width           =   945
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "Anular"
         Enabled         =   0   'False
         Height          =   660
         Left            =   120
         Picture         =   "FrmAnulacion.frx":2334
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2250
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   660
         Left            =   120
         Picture         =   "FrmAnulacion.frx":2A1E
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6330
         Width           =   945
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   660
         Left            =   120
         Picture         =   "FrmAnulacion.frx":2C28
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3570
         Width           =   945
      End
      Begin VB.CommandButton CmdDet 
         Caption         =   "Ir a Detalle"
         Enabled         =   0   'False
         Height          =   660
         Left            =   105
         Picture         =   "FrmAnulacion.frx":3312
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4905
         Width           =   945
      End
      Begin VB.Image Image2 
         Height          =   7425
         Left            =   30
         Picture         =   "FrmAnulacion.frx":3754
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1065
      End
   End
   Begin VB.Frame FraGrabarCancelar 
      Height          =   7728
      Left            =   45
      TabIndex        =   62
      Top             =   1035
      Visible         =   0   'False
      Width           =   1125
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   780
         Left            =   60
         Picture         =   "FrmAnulacion.frx":4608
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2265
         Width           =   1005
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   810
         Left            =   60
         Picture         =   "FrmAnulacion.frx":4812
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Image Image4 
         Height          =   7365
         Left            =   30
         Picture         =   "FrmAnulacion.frx":4A1C
         Stretch         =   -1  'True
         Top             =   150
         Width           =   1065
      End
   End
   Begin VB.Frame FraAdos 
      Height          =   5040
      Left            =   90
      TabIndex        =   83
      Top             =   960
      Visible         =   0   'False
      Width           =   2580
      Begin MSAdodcLib.Adodc AdoCategoria 
         Height          =   330
         Left            =   -120
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
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
         Caption         =   "Adodc1"
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
      Begin MSAdodcLib.Adodc AdoDocumento 
         Height          =   375
         Left            =   -120
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Caption         =   "AdoDocumento"
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
      Begin MSAdodcLib.Adodc AdoRuc 
         Height          =   375
         Left            =   0
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Caption         =   "AdoRuc"
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
      Begin MSAdodcLib.Adodc AdoFuente 
         Height          =   330
         Left            =   -120
         Top             =   1680
         Visible         =   0   'False
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "AdoFuente"
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
      Begin MSAdodcLib.Adodc AdoOrganismo 
         Height          =   330
         Left            =   -120
         Top             =   2040
         Visible         =   0   'False
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "AdoOrganismo"
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
      Begin MSAdodcLib.Adodc AdoProyecto 
         Height          =   330
         Left            =   -105
         Top             =   2760
         Visible         =   0   'False
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "AdoProyecto"
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
      Begin MSAdodcLib.Adodc AdoFormulario 
         Height          =   330
         Left            =   -120
         Top             =   3120
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
         CommandType     =   2
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
         Caption         =   "adoFormulario"
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
      Begin MSAdodcLib.Adodc AdoUnidad 
         Height          =   330
         Left            =   -90
         Top             =   3480
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
         Caption         =   "AdoUnidad"
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
      Begin MSAdodcLib.Adodc AdoTipoMoneda 
         Height          =   330
         Left            =   -90
         Top             =   3720
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
         CommandType     =   2
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
         Caption         =   "AdoTipoMoneda"
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
      Begin MSAdodcLib.Adodc AdoPartida 
         Height          =   330
         Left            =   -120
         Top             =   4080
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
         Caption         =   "AdoPartida"
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
      Begin MSAdodcLib.Adodc AdoCuenta 
         Height          =   330
         Left            =   -120
         Top             =   4425
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
      Begin MSAdodcLib.Adodc AdoTipo 
         Height          =   330
         Left            =   -90
         Top             =   600
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
         Caption         =   "AdoTipo"
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
   Begin VB.Frame FraOpcionesDetalle 
      Height          =   7836
      Left            =   45
      TabIndex        =   61
      Top             =   936
      Visible         =   0   'False
      Width           =   1125
      Begin VB.CommandButton CmdBorrarDetalle 
         Caption         =   "Anular"
         Height          =   735
         Left            =   105
         Picture         =   "FrmAnulacion.frx":58D0
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   2850
         Width           =   945
      End
      Begin VB.CommandButton CmdModificarDetalle 
         Caption         =   "Modificar"
         Height          =   780
         Left            =   105
         Picture         =   "FrmAnulacion.frx":5FBA
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1200
         Width           =   960
      End
      Begin VB.CommandButton CmdGrabaDetalle 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   780
         Left            =   105
         Picture         =   "FrmAnulacion.frx":61C4
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   2040
         Width           =   960
      End
      Begin VB.CommandButton CmdAgregarDetalle 
         Caption         =   "Adicionar"
         Height          =   780
         Left            =   90
         Picture         =   "FrmAnulacion.frx":63CE
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   375
         Width           =   960
      End
      Begin VB.CommandButton CmdSalirDetalle 
         Caption         =   "Salir"
         Height          =   780
         Left            =   105
         Picture         =   "FrmAnulacion.frx":66D8
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3630
         Width           =   945
      End
      Begin VB.Image Image3 
         Height          =   7455
         Left            =   30
         Picture         =   "FrmAnulacion.frx":68E2
         Stretch         =   -1  'True
         Top             =   165
         Width           =   1065
      End
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "mnuAcciones"
      Visible         =   0   'False
      Begin VB.Menu mnuAccion 
         Caption         =   "Devengado"
         Index           =   0
      End
      Begin VB.Menu mnuAccion 
         Caption         =   "Reversi�n"
         Index           =   1
      End
      Begin VB.Menu mnuAccion 
         Caption         =   "Devoluci�n"
         Index           =   2
      End
      Begin VB.Menu mnuAccion 
         Caption         =   "Anulaci�n"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FrmAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsNada As ADODB.Recordset
Dim rsD As ADODB.Recordset
Dim rsM As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsdiario As ADODB.Recordset
Dim rsCorr As ADODB.Recordset
Dim rsdev As ADODB.Recordset
Dim rsCoCoM As ADODB.Recordset
Dim rsPago_dev As ADODB.Recordset
Dim rsPpto As ADODB.Recordset
Dim rsRepCab As ADODB.Recordset
Dim rsRepDet As ADODB.Recordset
Dim rsAnterior As ADODB.Recordset
Dim rsauxiliar As ADODB.Recordset
Dim rsDocumentoRespaldo As ADODB.Recordset
Dim rsUnidad As ADODB.Recordset
Dim rsFuente As ADODB.Recordset
Dim rsOrganismo As ADODB.Recordset
Dim rsCategoria As ADODB.Recordset
Dim rsPartida As ADODB.Recordset
Dim rsProyecto As ADODB.Recordset
Dim rsBeneficiario As ADODB.Recordset
Dim rsCuenta As ADODB.Recordset
Dim rsRegularizacion As ADODB.Recordset
Dim rsTipoComprobante As ADODB.Recordset
Dim rsCorrel_Dev As ADODB.Recordset
Dim rsDet As ADODB.Recordset
Dim rsCtaB As ADODB.Recordset
Dim rsFGasto As ADODB.Recordset
Dim rsPg As ADODB.Recordset
Dim rstipocambio As ADODB.Recordset
Dim rscorrelativo As ADODB.Recordset
Dim swRefresca As Integer

Dim sql_TC As String
Public swModificaDetalle As String
Public swDevolucion As String
Public swGrabaCopia As Integer

Public sw2 As String
Public swA As String
Public Total_MontoBolivianos As Double
Public Total_MontoDolares As Double
Public Total_Deduccion As Double
Public Total_SaldoBolivianos As Double
Public ANTERIOR As String
Dim rsbusca As New ADODB.Recordset
Dim CAMPOS As ADODB.Field
Dim Marca As Integer
'Variables globales para copia de detalles en DEVOLUCION
Public vgFteCodigo As Variant
Public vgCodigoPartida As Variant
Public vgPrograma As Variant
Public vgSubPrograma As Variant
Public vgProyecto As Variant
Public vgActividad As Variant
Public vgCtaOrigen As Variant
Public vgNroChequeOTransferencia As Variant
Public vgCtaDestino As Variant
Public vgCodBeneficiario As Variant
Public vgMontoTotal As Variant
Public vgDeducciones As Variant
'Public vgMontoBolivianos As Double
Public vgMB As Currency
Public vgTipoCambio As Variant
Public vgMontoDolares As Variant
Public vgOrgCodigo As Variant
Public vgGesGestion As Variant
Public vgCodigoPago As Variant
Public vgCodigoPagoDetalle As Variant
Public ComprobanteAnterior As Variant
Public TIPOFORMULARIO As String

Dim sino As String
Dim X As String
Dim Y As String
Dim z As String
Dim swgraba As String
Dim ppto2 As String
Dim Org3 As String
Dim cocmCod_CompDiario As String
Dim cocmTipo_Comp As String
Dim cocmCod_Trans As String
Dim cocmCod_Trans_Detalle As String
Dim cocmOrg_Codigo As String
Dim cocmGes_Gestion As String
Dim cocmNum_Respaldo As String
Dim cocmFecha_A As String
Dim cocmCodigo_Beneficiario As String
Dim cocmCodigo_Documento As String
Dim cocmGlosa As String
Dim cocmStatus As String
Dim cocmUsr_usuario As String
Dim cocmCod_Comp As Variant
Dim cocmCod_Comp_C As Variant
Dim AuxCod_Comp  As String
Dim AuxTipo_Comp As String
Dim AuxCod_Comp_C As Integer
Dim AuxD_Cuenta  As String
Dim AuxD_Nombre  As String
Dim AuxD_SubCta1  As String
Dim AuxD_SubCta2  As String
Dim AuxD_Aux1  As String
Dim AuxD_Aux2  As String
Dim AuxD_Aux3  As String
Dim AuxD_Cta_Larga  As String
Dim AuxD_Des_Larga As String
Dim AuxD_MontoBs As String
Dim AuxD_MontoDL As String
Dim AuxD_Cambio As String
                    
Dim AuxH_Cuenta As String
Dim AuxH_Nombre As String
Dim AuxH_SubCta1 As String
Dim AuxH_SubCta2 As String
Dim AuxH_Aux1 As String
Dim AuxH_Aux2 As String
Dim AuxH_Aux3 As String
Dim AuxH_Cta_Larga As String
Dim AuxH_Des_Larga As String
Dim AuxH_MontoBs As String
Dim AuxH_MontoDL As String
Dim AuxH_Cambio As String
                    
Dim AuxUsr_Usuario As String
Dim AuxFecha_Registro As Variant
Dim AuxHora_Registro As String
Dim AuxCopia As String

Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim EntrarAdo As Boolean 'Para que al aprobar no muestre uno por uno
Dim QueryInicial As String
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String

Private Sub AdoRegularizacion_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Not EntrarAdo Then Exit Sub
If pRecordset.State <> 1 Then Exit Sub
   If Not AdoRegularizacion.Recordset.EOF And Not AdoRegularizacion.Recordset.BOF And AdoRegularizacion.Recordset.RecordCount > 0 Then
      If swA = "2" Then
      '         If Not IsNull(AdoRegularizacion.Recordset("fte_codigo")) Then DTcFte.Text = AdoRegularizacion.Recordset("fte_codigo") Else DTcFte.Text = ""
      '         If Not IsNull(AdoRegularizacion.Recordset("org_codigo")) Then DtcOrg.Text = AdoRegularizacion.Recordset("org_codigo") Else DtcOrg.Text = ""
      '         If Not IsNull(AdoRegularizacion.Recordset("uni_codigo")) Then DtCUnidad.Text = AdoRegularizacion.Recordset("uni_codigo") Else DtCUnidad.Text = ""
      '         If Not IsNull(AdoRegularizacion.Recordset("Codigo_orden")) Then TxtCodigoOrden.Text = AdoRegularizacion.Recordset("Codigo_orden") Else TxtCodigoOrden = ""
      '         If Not IsNull(AdoRegularizacion.Recordset("Codigo_Solicitud")) Then txtNroSolicitud.Text = AdoRegularizacion.Recordset("Codigo_Solicitud") Else txtNroSolicitud = ""
      End If
      ' Detalle
      If Not IsNull(AdoRegularizacion.Recordset("codigo_pago")) And Not IsNull(AdoRegularizacion.Recordset("org_codigo")) Then
            Set rsDetalle = New ADODB.Recordset
              rsDetalle.Open "select * from pago_detalle where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
            Set DtGDetalle.DataSource = rsDetalle
            If rsDetalle.RecordCount > 0 Then
                DtGDetalle.Refresh
            End If
      End If
      ' VERIFICAMOS QUE OPCIONES DEL POPUP ACTIVAMOS
      With AdoRegularizacion
        If IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = True
              mnuAccion(1).Enabled = True
              mnuAccion(2).Enabled = False
              mnuAccion(3).Enabled = False
              AuxCopia = "R"
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = True
              mnuAccion(2).Enabled = False
              mnuAccion(3).Enabled = False
              AuxCopia = "R"
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "S" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = False
              mnuAccion(2).Enabled = True
              mnuAccion(3).Enabled = True
              AuxCopia = "R"
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "S" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = False
              mnuAccion(2).Enabled = False
              mnuAccion(3).Enabled = False
              AuxCopia = "R"
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "S" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = False
              mnuAccion(2).Enabled = False
              mnuAccion(3).Enabled = False
              AuxCopia = "R"
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" And _
                 IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "S" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "S" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = False
              mnuAccion(2).Enabled = False
              mnuAccion(3).Enabled = False
              AuxCopia = "R"
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "S" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "S" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = False
              mnuAccion(2).Enabled = False
              mnuAccion(3).Enabled = False
              AuxCopia = "R"
          ' ADD. por Jorge
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = True
              mnuAccion(2).Enabled = False
              mnuAccion(3).Enabled = False
              AuxCopia = "R"
          ElseIf IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "" And _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "S" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = False
              mnuAccion(2).Enabled = True
              mnuAccion(3).Enabled = True
              AuxCopia = "R"
          ' ADD. por Jorge
          'adiciono gaby
          '***opciones de contabilidad
          'IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "" And _'
          ElseIf IIf(IsNull(.Recordset!org_codigo), "", .Recordset!org_codigo) = "999" And _
             (IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" Or _
             IIf(IsNull(.Recordset!estado_contabilidad), "", .Recordset!estado_contabilidad) = "P") And _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "S" And _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
              mnuAccion(0).Enabled = False
              mnuAccion(1).Enabled = False
              mnuAccion(2).Enabled = True
              mnuAccion(3).Enabled = True
'          ElseIf IIf(IsNull(.Recordset!Org_Codigo), "", .Recordset!Org_Codigo) = "999" And _
'             (IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" Or _
'             IIf(IsNull(.Recordset!estado_contabilidad), "", .Recordset!estado_contabilidad) = "P") And _
'             ((IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "") Or _
'             (IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "N")) And _
'             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "" And _
'             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "" Then
'              mnuAccion(0).Enabled = False
'              mnuAccion(1).Enabled = True
'              mnuAccion(2).Enabled = False
'              mnuAccion(3).Enabled = False
'          'fin de adiciono gaby
        Else
            mnuAccion(0).Enabled = False
            mnuAccion(1).Enabled = False
            mnuAccion(2).Enabled = False
            mnuAccion(3).Enabled = False
        End If
        'greco no molestes aqu� ... que hondas
'        If IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "E" Or _
'           IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "E" Then
'           cmdAprueba.Enabled = False
'           CmdCopiar.Enabled = False
'           CmdModificar.Enabled = False
'           CmdBorrar.Enabled = False
'        Else
'           cmdAprueba.Enabled = True
'           CmdCopiar.Enabled = True
'           CmdModificar.Enabled = True
'           CmdBorrar.Enabled = True
'        End If
        'greco no molestes aqu� ... que hondas
      End With
     Else
            Set DtGDetalle.DataSource = rsNada
            mnuAccion(0).Enabled = False
            mnuAccion(1).Enabled = False
            mnuAccion(2).Enabled = False
            mnuAccion(3).Enabled = False
     End If
    'adiciono gaby
    If AdoRegularizacion.Recordset!org_codigo = "999" Then
      CmdAprueba.Enabled = False
    Else
      CmdAprueba.Enabled = True
    End If
    'fin adiciono gaby
End Sub

Private Sub CmdAceptarDev_Click()
    Devolucion
End Sub

Private Sub Cmdadicionar_Click()
On Error GoTo adiciona
        DtpFecha.Enabled = True
         FraMaestro.Enabled = True
         LblTitulo.Caption = "ADICIONANDO . . . "
         DtcDcu.Refresh
         DtcDcuDes.Refresh
         
         Set rsauxiliar = New ADODB.Recordset
         Set rsauxiliar = rsRegularizacion
        
         AdoRegularizacion.Recordset.AddNew
         TxtCodigoOrden.Text = ""
         TxtComprobante.Text = ""
         TxtComprobanteAnterior.Text = ""
         txtNroSolicitud.Text = ""
         DtCUnidad.Text = ""
         DTcFte.Text = ""
         DtcOrg.Text = ""
         DtcCat.Text = ""
         TxtJustificacion.Text = ""
         TxtDeducciones.Text = ""
         txtNroSolicitud.SetFocus
         FraOpciones.Visible = False
         FraGrabarCancelar.Visible = True
         DtpFecha.Text = CDate(Date)
         DtcTipoDes.Visible = True
         TxtTipoReg.Visible = False
         sw2 = "1"
         swA = "2"
Exit Sub
adiciona:
   MsgBox Err.Number & " " & Err.Description

End Sub

Private Sub CmdAgregarDetalle_Click()
On Error Resume Next
    
    FraDetalleG.Enabled = True
    Label35.Caption = "ADICIONANDO DETALLE . . ."
    
    TxtTipoCambio.Enabled = True
    Set rstipocambio = New ADODB.Recordset
    sql_TC = "select fecha_cambio, Cambio_Oficial  from ac_tipo_cambio  where fecha_cambio = (select max(fecha_cambio) as expr1 from ac_tipo_cambio)"
    rstipocambio.Open sql_TC, db, adOpenKeyset, adLockReadOnly
    GlTipoCambioOficial = rstipocambio!cambio_oficial
    'TFecha = rstipocambio!fecha_cambio
    

    AdoDetalle.Recordset.AddNew
    TxtTipoCambio.Text = GlTipoCambioOficial
    TxtTipoCambio.Enabled = False
    TxtDeducciones.Text = 0
    TxtDeducciones.Enabled = False
    'Set rstipocambio = New ADODB.Recordset
    TxtCodigoDetalle.Text = AdoDetalle.Recordset.RecordCount
    DtCPartida.Text = ""
    DtCCuentaOrigen.Text = ""
    DtCCuentaDestino.Text = ""
    TxtNoTransferenciaOrigen.Text = ""
    CmdGrabaDetalle.Enabled = True
    CmdAgregarDetalle.Enabled = False
    CmdModificarDetalle.Enabled = False
    CmdBorrarDetalle.Enabled = False
    'Command11.Enabled = False
    msgSalir = "1"
Exit Sub
'l:
'   MsgBox "Esta es una prueba", vbCritical
End Sub



'Private Sub Cmd_Pagado(P_codigo_pago As String, P_codigo_pago_detalle As String, P_org_codigo As String, P_ges_gestion As String)
'Dim sw As Boolean
'Dim Sw_Fuente As Boolean
'Dim Cont_Comp As Long
'Dim aux_T As String
'
'On Error GoTo errorPag
'
'db.BeginTrans
'
''        MsgBox AdoPagoDetalle.Recordset("ges_gestion")
''        MsgBox AdoPagoDetalle.Recordset("org_codigo")
''        MsgBox AdoPagoDetalle.Recordset("codigo_pago")
''        MsgBox AdoPagoDetalle.Recordset("codigo_pago_detalle")
''       'Contabiliza_Automatico
'
'
''*******************************************************
''******************** Contabilizar Pagos ***************'
''********************************************************
''************** Para inicializar el contador ******************'
'
''*********** Para obtenerr en el recordset recsetAuxComp losdatos necesarios para almacenar*********"
'
''Set recSetAuxcomp1 = New ADODB.Recordset
''recSetAuxcomp1.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'
''Set recSetAuxcomp1 = New ADODB.Recordset
''recSetAuxcomp1.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
''If recSetAuxcomp1.State = 1 Then recSetAuxcomp.Close
''recSetAuxcomp1.Open "SELECT * from ts_cheque   ", db, adOpenDynamic, adLockOptimistic, adCmdText
''
''If recSetAuxcomp1.RecordCount > 0 Then
''    recSetAuxcomp1.MoveFirst
''End If
'
''While Not (recSetAuxcomp1.EOF)
'
'
'        Set recSetAuxcomp = New ADODB.Recordset
'        recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'
'        If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'        recSetAuxcomp.Open "SELECT distinct pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion,Estado_Pagado,Pago_Detalle.Cta_Codigo,Pago_Detalle.tipo_cambio," & _
'        " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_Bolivianos,estado_Devengado,Pagos.Org_Codigo,Pagos.Codigo_Orden,Pagos.Codigo_Documento," & _
'        " pago_detalle.Monto_Dolares,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and   " & _
'        " pago_Detalle.Org_codigo= '" & P_org_codigo & "' and  pago_detalle.Ges_Gestion='" & P_ges_gestion & "' and pago_detalle.codigo_Pago=" & Val(P_codigo_pago) & " and " & " pago_detalle.Ges_Gestion = pagos.Ges_Gestion and pago_detalle.codigo_pago_detalle='" & P_codigo_pago_detalle & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        'and pago_detalle.codigo_pago_detalle='" & P_codigo_pago_detalle & "'
'        'pagos.Org_codigo='" & rsCheque!cod_org & "' and
'        'pago_detalle.estado_aprobacion ='A' and pago_detalle.Ges_Gestion='" & rsCheque!Ges_Gestion & "' and pago_detalle.codigo_Pago='" & rsCheque!Numero_comprobante & "'
'        'and  pagos.estado_Pagado='S'  AND Pagos.Tipo_comp='PAC'
'        'AND pago_detalle.estado_aprobacion = 'A'
'        If recSetAuxcomp.RecordCount > 0 Then
'        recSetAuxcomp.MoveFirst
'        End If
'While Not (recSetAuxcomp.EOF)
'
'
'        '************Abrimos un record set para adicionar datos*********************'
'        Set recSetAuxActualizar = New ADODB.Recordset
'        If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'        recSetAuxActualizar.Open " select * from CO_Comprobante_M ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
'        Set recSetAuxActualizar1 = New ADODB.Recordset
'        If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar.Close
'        recSetAuxActualizar1.Open " select * from CO_Diario ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        Dim Aux_Cont As String
'
'        aux_T = "select * from Co_comprobante_M"
'
'        'While Not (recSetAuxcomp.EOF)
'
'        If Not Buscar(aux_T, recSetAuxcomp!codigo_pago, recSetAuxcomp!org_codigo, recSetAuxcomp!ges_gestion, "PAC", recSetAuxcomp!codigo_pago_detalle) Then
'
'            Select Case recSetAuxcomp!fte_codigo
'
'            Case "10"
'
'            Set recSetPartida = New ADODB.Recordset
'            recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'            If recSetPartida.State = 1 Then recSetPartida.Close
'            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H, CC_Cuentas_D" & _
'            " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst= 'PAG' and CC_Cuenta_H.Inst= 'PAG' and " & _
'            " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=1 AND " & _
'            " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            Sw_Fuente = True
'
'           Case "70"
'            Set recSetPartida = New ADODB.Recordset
'            recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'            If recSetPartida.State = 1 Then recSetPartida.Close
'            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H, CC_Cuentas_D" & _
'            " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PAG' and CC_Cuenta_H.Inst='PAG' and " & _
'            " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=2 AND " & _
'            " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            Sw_Fuente = True
'
'            Case "80"
'            Set recSetPartida = New ADODB.Recordset
'            recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'            If recSetPartida.State = 1 Then recSetPartida.Close
'            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H, CC_Cuentas_D" & _
'            " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PAG' and CC_Cuenta_H.Inst='PAG' and " & _
'            " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=3 and  " & _
'            " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            Sw_Fuente = True
'
'            Case Else
'            'MsgBox "No esta asociado a ninguna fuente ... partida no relacionada "
'            Sw_Fuente = False
'
'            End Select
'          If Sw_Fuente Then
'
'            recSetAuxActualizar.AddNew
'            recSetAuxActualizar1.AddNew
'            'recSetAuxActualizar!Cod_Comp = Cont_Comp
'            recSetAuxActualizar!Cod_trans = recSetAuxcomp!codigo_pago
'            recSetAuxActualizar!Cod_Trans_Detalle = recSetAuxcomp!codigo_pago_detalle
'            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
'            recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
'            recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
'            recSetAuxActualizar!num_respaldo = recSetAuxcomp!codigo_orden
'            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
'
'            recSetAuxActualizar!fecha_A = recSetAuxcomp!fecha_pago
'            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
'            'recSetAuxActualizar!codigo_solicitud = recSetAuxcomp!codigo_solicitud
'            recSetAuxActualizar!tipo_comp = "PAC"
'
'            recSetAuxActualizar!Status = "S"
'            recSetAuxActualizar1!tipo_comp = "PAC"
'            recSetAuxActualizar1!d_cuenta = recSetPartida!cuenta
'            recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
'            recSetAuxActualizar1!d_subcta1 = recSetPartida!subcta1
'            recSetAuxActualizar1!d_subcta2 = recSetPartida!subcta2
'            recSetAuxActualizar1!d_Aux1 = recSetPartida!aux1
'            recSetAuxActualizar1!d_Aux2 = recSetPartida!aux2
'            recSetAuxActualizar1!d_Aux3 = recSetPartida!aux3
'
'        '************* CONTABILIZA AUXILIAARES DEBITO
'            Select Case recSetPartida!aux1
'            Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
'
'            Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
'
'            Case Else
'            End Select
'        ''****************** finaliza sesion de auxiliares
'
'
'            recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
'            recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
'            recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
'
'        '************* CONTABILIZA AUXILIAARES DEBITO
'
'            Select Case recSetPartida!h_Aux1
'            Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
'
'            Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    'recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga
'
'            Case Else
'            End Select
'        ''****************** finaliza sesion de auxiliares
'
'            recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
'            recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
'            recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
'            recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2
'            recSetAuxActualizar1!d_montobs = recSetAuxcomp!monto_bolivianos
'            recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_Dolares
'            recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_Dolares
'            recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
'
'            recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_bolivianos
'            recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_Dolares
'            recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_Dolares
'            recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
'            ''************ GENERA EL CODIGO DE COMPROBANTE**********
'
'                    Set recSetGenera = New ADODB.Recordset
'                    recSetGenera.CursorLocation = adUseClient
'                    If recSetGenera.State = 1 Then recSetGenera.Close
'                    recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    If recSetGenera.RecordCount > 0 Then
'                     Cont_Comp = Val(recSetGenera!numero_correlativo)
'                     Cont_Comp = Cont_Comp + 1
'                     recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
'
'
'
'        '************TERMINA GENERACION DE COMPROBANTE********
'                     recSetAuxActualizar!Cod_Comp = Cont_Comp
'                     recSetAuxActualizar1!Cod_Comp = Cont_Comp
'                     recSetAuxActualizar1.Update
'                     recSetAuxActualizar.Update
'                     recSetGenera.Update
'
'                    End If
'
'           Else
'                MsgBox "No esta asociado a ninguna fuente ...  "
'
'           End If
'        Else
'            MsgBox "Existe registro....."
'        End If
'            'Cont_Comp = Cont_Comp + 1
'            recSetAuxcomp.MoveNext
'Wend
'''Unload Me
''recSetGenera!Numero_correlativo = Cont_Comp
''recSetGenera.Update
'db.CommitTrans
'MsgBox "Click para continuar la Impresi�n ... "
''Unload Me
'Exit Sub
'errorPag:
'db.RollbackTrans
'MsgBox "Error, No se contabiliz� ... "
'Unload Me
'
'End Sub
'Private Sub Cmd_ContaConf(P_codigo_pago As String, P_org_codigo As String, P_ges_gestion As String)
'Dim sw As Boolean
'Dim Sw_Fuente As Boolean
'Dim Cont_Comp As Long
'
'Dim aux_T As String
'
'On Error GoTo errorComp
'db.BeginTrans
'
'
''********* Para obtener en el recordset recsetAuxComp los datos necesarios para almacenar*********"
'    Set recSetAuxcomp = New ADODB.Recordset
'    recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'    If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'    recSetAuxcomp.Open "SELECT distinct pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion," & _
'    " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_total,Pagos.org_Codigo,pagos.Codigo_orden,Pagos.Codigo_documento," & _
'    " pago_detalle.Monto_Dolares,pago_detalle.Tipo_Cambio,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and TIPO_COMP='DAC' AND " & _
'    " pago_detalle.Ges_Gestion = pagos.Ges_Gestion AND pagos.estado_Devengado= 'S' AND pagos.codigo_Pago= '" & P_codigo_pago & "' and pagos.Org_Codigo='" & P_org_codigo & "' and pago_detalle.Ges_Gestion = '" & P_ges_gestion & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
'   'FrmRegularizacion.AdoRegularizacion.Recordset!Codigo_Pago
'   'FrmRegularizacion.AdoRegularizacion.Recordset!Org_Codigo
'   'FrmRegularizacion.AdoRegularizacion.Recordset!Ges_gestion
'
'   'FrmRegularizacion.AdoRegularizacion.Recordset
'    '*******  Mueve al primer registro
'    If recSetAuxcomp.RecordCount > 0 Then
'    recSetAuxcomp.MoveFirst
'    End If
'
'    '************Abrimos un record set para adicionar datos*********************'
'
'    Set recSetAuxActualizar = New ADODB.Recordset
'    If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'    recSetAuxActualizar.Open " select * from CO_Comprobante_M ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
'    Set recSetAuxActualizar1 = New ADODB.Recordset
'    If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar.Close
'    recSetAuxActualizar1.Open " select * from CO_Diario ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
'    aux_T = "select * from Co_comprobante_M"
'
'    While Not (recSetAuxcomp.EOF)
'    If Not Buscar(aux_T, recSetAuxcomp!codigo_pago, recSetAuxcomp!org_codigo, recSetAuxcomp!ges_gestion, "DAC", recSetAuxcomp!codigo_pago_detalle) Then
'        Set recSetPartida = New ADODB.Recordset
'        If recSetPartida.State = 1 Then recSetPartida.Close
'        recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H,CC_Cuentas_D" & _
'        " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='DEV' and CC_Cuenta_H.Inst='DEV' and" & _
'        " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
''If recSetPartida.RecordCount > 0 Then
'    recSetAuxActualizar.AddNew
'    recSetAuxActualizar1.AddNew
'
'    'recSetAuxActualizar!Cod_Comp = Cont_Comp
'    recSetAuxActualizar!Cod_trans = recSetAuxcomp!codigo_pago
'    recSetAuxActualizar!Cod_Trans_Detalle = recSetAuxcomp!codigo_pago_detalle
'    recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
'
'    recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
'    recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
'    recSetAuxActualizar!fecha_A = recSetAuxcomp!fecha_pago
'    recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
'    recSetAuxActualizar!num_respaldo = recSetAuxcomp!codigo_orden
'    recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
'    recSetAuxActualizar!Status = "S"
'
'   ' recSetAuxActualizar!codigo_solicitud = recSetAuxcomp!codigo_solicitud
'    recSetAuxActualizar!tipo_comp = "DAC"
'
'   ' recSetAuxActualizar1!Cod_Comp = Cont_Comp
'    recSetAuxActualizar1!tipo_comp = "DAC"
'    recSetAuxActualizar1!d_cuenta = recSetPartida!cuenta
'    recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
'    recSetAuxActualizar1!d_subcta1 = recSetPartida!subcta1
'    recSetAuxActualizar1!d_subcta2 = recSetPartida!subcta2
'    recSetAuxActualizar1!d_Aux1 = recSetPartida!aux1
'    recSetAuxActualizar1!d_Aux2 = recSetPartida!aux2
'    recSetAuxActualizar1!d_Aux3 = recSetPartida!aux3
'
'''******* ADICION DE AUXILIARES A DETALLE DEBITO*******
'    Select Case recSetPartida!aux1
'    Case "01"
'            Set recsetAdicion = New ADODB.Recordset
'            If recsetAdicion.State = 1 Then recsetAdicion.Close
'            recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
'            recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
'
'    Case "02"
'            Set recsetAdicion = New ADODB.Recordset
'            If recsetAdicion.State = 1 Then recsetAdicion.Close
'            recsetAdicion.Open " select * from fc_cuenta_Bancaria where cTA_cODIGO='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
'            recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
'
'
'
'    Case Else
'    End Select
'''****************** finaliza sesion de auxiliares DEBITO
'    recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
'    recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
'    recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
'    recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2
'
'    recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
'    recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
'    recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
'''******* ADICION DE AUXILIARES A DETALLE*******
'    Select Case recSetPartida!h_Aux1
'    Case "01"
'            Set recsetAdicion = New ADODB.Recordset
'            If recsetAdicion.State = 1 Then recsetAdicion.Close
'            recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
'            recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
'
'    Case "02"
'            Set recsetAdicion = New ADODB.Recordset
'            If recsetAdicion.State = 1 Then recsetAdicion.Close
'            recsetAdicion.Open " select * from fc_cuenta_Bancaria where CTA_CODIGO='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
'            recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga
'
'    Case Else
'    End Select
'''****************** finaliza sesion de auxiliares
'
'
'    recSetAuxActualizar1!d_montobs = recSetAuxcomp!monto_total
'    recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_Dolares
'    recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
'
'    recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_total
'    recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_Dolares
'    recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
'''************ GENERA EL CODIGO DE COMPROBANTE**********
'
'            Set recSetGenera = New ADODB.Recordset
'            recSetGenera.CursorLocation = adUseClient
'            If recSetGenera.State = 1 Then recSetGenera.Close
'            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'            If recSetGenera.RecordCount > 0 Then
'             Cont_Comp = Val(recSetGenera!numero_correlativo)
'             Cont_Comp = Cont_Comp + 1
'             recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
'
'
'''************TERMINA GENERACION DE COMPROBANTE********
'
'             recSetAuxActualizar!Cod_Comp = Cont_Comp
'             recSetAuxActualizar1!Cod_Comp = Cont_Comp
'
'             recSetAuxActualizar1.Update
'             recSetAuxActualizar.Update
'             recSetGenera.Update
'
'            End If
'
'' Else
' ' MsgBox "No existe partidas"
' 'End If
'
'Else
' MsgBox "Existe registro....."
'End If
''Cont_Comp = Cont_Comp + 1
'recSetAuxcomp.MoveNext
'Wend
'
'db.CommitTrans
'MsgBox "Contabilizo con exito....."
''Unload Frm_Cont_Mat
'
'Exit Sub
'errorComp:
'db.RollbackTrans
'MsgBox "No contabilizo con exito......"
''Unload Frm_Cont_Mat
'
'End Sub
'
'Private Sub CmdAnulacion_Click()
'    Set rsRegularizacion = New ADODB.Recordset
'    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
'    rsRegularizacion.Open "select * from pagos where tipo_comp = 'DAC' and estado_compromiso='S' and estado_devengado='S' and estado_pagado='S' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'    'rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC' or  tipo_comp = 'CYD') and estado_devengado='S' and estado_pagado='S' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'    CmdAprueba.Enabled = True
'    If rsRegularizacion.RecordCount > 0 Then
'        Set DtcRegularizacion.DataSource = AdoRegularizacion
'        Set AdoRegularizacion.Recordset = rsRegularizacion
'    Else
'        MsgBox "No existen datos", vbInformation, "Validaci�n de datos"
'    End If
'    'FraBusqueda.Visible = False
'    FraMaestro.Enabled = True
'    swDevolucion = "A"
'End Sub
'Private Sub Cmd_Pagado(P_codigo_pago As String, P_codigo_pago_detalle As String, P_org_codigo As String, P_ges_gestion As String)
'Dim sw As Boolean
'
'Dim Sw_Fuente As Boolean
'Dim Cont_Comp As Long
'Dim aux_T As String
'
'Dim v_Cuenta As String
'Dim v_SubCta1 As String
'Dim v_SubCta2 As String
'Dim v_NombreCta As String
'Dim v_H_Cuenta As String
'Dim v_H_SubCta1 As String
'Dim v_H_SubCta2 As String
'Dim v_H_NombCta As String
'Dim v_Aux1 As String
'Dim v_Aux2 As String
'Dim v_Aux3 As String
'Dim v_H_Aux1 As String
'Dim v_H_Aux2 As String
'Dim v_H_Aux3 As String
'Dim Aux_Cont As String
'
'On Error GoTo errorPag
'
'db.BeginTrans
'        MsgBox "Contabilizar............", vbOKOnly, "Contabilizaci�n"
'        Set recSetAuxcomp = New ADODB.Recordset
'        recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'
'    If Me.DtCCuentaOrigen.Text = "" Then
'            MsgBox "ERROR, NO SE CONTABILIZO", vbDefaultButton1 + vbOKOnly
'            Exit Sub
'    End If
'        If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'        recSetAuxcomp.Open "SELECT distinct pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion,Estado_Pagado,Pago_Detalle.Cta_Codigo,Pago_Detalle.tipo_cambio," & _
'        " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_Bolivianos,estado_Devengado,Pagos.Org_Codigo,Pagos.Codigo_Orden,Pagos.Codigo_Documento," & _
'        " pago_detalle.Monto_Dolares,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and   " & _
'        " pago_Detalle.Org_codigo= '" & P_org_codigo & "' and  pago_detalle.Ges_Gestion='" & P_ges_gestion & "' and pago_detalle.codigo_Pago=" & Val(P_codigo_pago) & " and " & _
'        " pago_detalle.Ges_Gestion = pagos.Ges_Gestion  and pago_detalle.codigo_pago_detalle='" & P_codigo_pago_detalle & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        If recSetAuxcomp.RecordCount > 0 Then
'            recSetAuxcomp.MoveFirst
'        Else
'            MsgBox "ERROR EN LA CONTABILIZACION", vbCritical + vbDefaultButton1
'        End If
'While Not (recSetAuxcomp.EOF)
''VERIFICA FUENTE
'    Select Case recSetAuxcomp!fte_codigo
'    Case "10", "41"
'        Set recSetPartida = New ADODB.Recordset
'        recSetPartida.CursorLocation = adUseClient
'        If recSetPartida.State = 1 Then recSetPartida.Close
'        recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'        " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst= 'PAG' and CC_Cuenta_H1.Inst= 'PAG' and " & _
'        " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=1 AND " & _
'        " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        Sw_Fuente = True
'    'Asignacion a variables
'
'    Case "70", "43"
'        Set recSetPartida = New ADODB.Recordset
'        recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'        If recSetPartida.State = 1 Then recSetPartida.Close
'        recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'        " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PAG' and CC_Cuenta_H1.Inst='PAG' and " & _
'        " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=2 AND " & _
'        " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        Sw_Fuente = True
'
'    Case "80"
'        Set recSetPartida = New ADODB.Recordset
'        recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'        If recSetPartida.State = 1 Then recSetPartida.Close
'        recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H1, CC_Cuentas_D1" & _
'        " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PAG' and CC_Cuenta_H1.Inst='PAG' and " & _
'        " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=3 and  " & _
'        " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        Sw_Fuente = True
'
'    Case Else
'        Sw_Fuente = False
'        MsgBox "No esta asociado a ninguna fuente ... partida no relacionada "
'    End Select
'
'    If Sw_Fuente Then
''Asignacion a variables
'        v_Cuenta = recSetPartida!cuenta
'        v_SubCta1 = recSetPartida!subcta1
'        v_SubCta2 = recSetPartida!subcta2
'        v_NombreCta = recSetPartida!NombreCta
'        v_H_Cuenta = recSetPartida!h_cuenta
'        v_H_SubCta1 = recSetPartida!h_subcta1
'        v_H_SubCta2 = recSetPartida!h_subcta2
'        v_H_NombCta = recSetPartida!H_NombCta
'
'        v_Aux1 = recSetPartida!aux1
'        v_Aux2 = recSetPartida!aux2
'        v_Aux3 = recSetPartida!aux3
'
'        v_H_Aux1 = recSetPartida!h_Aux1
'        v_H_Aux2 = recSetPartida!h_Aux2
'        v_H_Aux3 = recSetPartida!h_Aux3
'
'        If recSetPartida.State = 1 Then recSetPartida.Close
'
''************Abrimos un record set para adicionar datos*********************'
'        Set recSetAuxActualizar = New ADODB.Recordset
'        If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'        recSetAuxActualizar.Open " select * from CO_Comprobante_M  where Cod_Trans='" & P_codigo_pago & "' and Org_Codigo='" & P_org_codigo & "' " & _
'        " and Ges_Gestion='" & P_ges_gestion & "' and Tipo_comp='PAC' and Cod_Trans_Detalle='" & P_codigo_pago_detalle & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
'        If Not recSetAuxActualizar.BOF Then recSetAuxActualizar.MoveFirst
'        If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
''************* GENERA EL CODIGO DE COMPROBANTE**********
'            Set recSetGenera = New ADODB.Recordset
'            recSetGenera.CursorLocation = adUseClient
'            If recSetGenera.State = 1 Then recSetGenera.Close
'            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'            If recSetGenera.RecordCount > 0 Then
'                Cont_Comp = Val(recSetGenera!numero_correlativo)
'                Cont_Comp = Cont_Comp + 1
'                recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
'                recSetGenera.Update
'            End If
'            If recSetGenera.State = 1 Then recSetGenera.Close
''************TERMINA GENERACION DE COMPROBANTE********
'' Datos Para co_Comprobante
'
'            recSetAuxActualizar.AddNew
'            recSetAuxActualizar!Cod_Comp = Cont_Comp
'            recSetAuxActualizar!Cod_trans = recSetAuxcomp!codigo_pago
'            recSetAuxActualizar!Cod_Trans_Detalle = recSetAuxcomp!codigo_pago_detalle
'            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
'            recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
'            recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
'            recSetAuxActualizar!num_respaldo = recSetAuxcomp!codigo_orden
'            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
'            recSetAuxActualizar!fecha_A = recSetAuxcomp!fecha_pago
'            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
'            recSetAuxActualizar!tipo_comp = "PAC"
'            recSetAuxActualizar!Status = "S"
'            recSetAuxActualizar.Update
'            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'
'' Datos Para co_Diario
'            Set recSetAuxActualizar1 = New ADODB.Recordset
'            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'            recSetAuxActualizar1.Open " select * from CO_Diario where  cod_Comp = " & Cont_Comp & " ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
'                recSetAuxActualizar1.AddNew
'                recSetAuxActualizar1!tipo_comp = "PAC"
'                recSetAuxActualizar1!d_cuenta = v_Cuenta
'                recSetAuxActualizar1!D_Nombre = v_NombreCta
'                recSetAuxActualizar1!d_subcta1 = v_SubCta1
'                recSetAuxActualizar1!d_subcta2 = v_SubCta2
'                recSetAuxActualizar1!d_Aux1 = v_Aux1
'                recSetAuxActualizar1!d_Aux2 = v_Aux2
'                recSetAuxActualizar1!d_Aux3 = v_Aux3
''************* CONTABILIZA AUXILIAARES DEBITO
'                Select Case v_Aux1
'                Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
'
'                Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
'                Case Else
'                End Select
'''****************** finaliza sesion de auxiliares
'                recSetAuxActualizar1!h_Aux1 = v_H_Aux1
'                recSetAuxActualizar1!h_Aux2 = v_H_Aux2
'                recSetAuxActualizar1!h_Aux3 = v_H_Aux3
''************* CONTABILIZA AUXILIAARES CREDITO
'
'                Select Case v_H_Aux1
'                Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
'
'                Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
''recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga
'
'                Case Else
'                End Select
'''****************** finaliza sesion de auxiliares
'
'                recSetAuxActualizar1!h_cuenta = v_H_Cuenta
'                recSetAuxActualizar1!H_Nombre = v_H_NombCta
'                recSetAuxActualizar1!h_subcta1 = v_H_SubCta1
'                recSetAuxActualizar1!h_subcta2 = v_H_SubCta2
'                recSetAuxActualizar1!d_montobs = recSetAuxcomp!monto_bolivianos
'                recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_Dolares
'                recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
'
'                recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_bolivianos
'                recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_Dolares
'                recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
'                recSetAuxActualizar1!Cod_Comp = Cont_Comp
'                recSetAuxActualizar1.Update
'            End If
'        Else
'        MsgBox "Ya fue contabilizado anteriormente...  ", vbOKOnly, "contabilizando...  "
'
'
'' buscar el que ya existe y reemplazar los datos
'
'            If (Not recSetAuxActualizar.BOF) Then recSetAuxActualizar.MoveFirst
''            recSetAuxActualizar!Cod_Comp = Cont_Comp
'            Cont_Comp = recSetAuxActualizar!Cod_Comp
'            recSetAuxActualizar!Cod_trans = recSetAuxcomp!codigo_pago
'            recSetAuxActualizar!Cod_Trans_Detalle = recSetAuxcomp!codigo_pago_detalle
'            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
'            recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
'            recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
'            recSetAuxActualizar!num_respaldo = recSetAuxcomp!codigo_orden
'            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
'            recSetAuxActualizar!fecha_A = recSetAuxcomp!fecha_pago
'            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
''            recSetAuxActualizar!Tipo_Comp = "PAC"
'            recSetAuxActualizar!Status = "S"
'            recSetAuxActualizar.Update
'            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'
'' Datos Para co_Diario
'            Set recSetAuxActualizar1 = New ADODB.Recordset
'            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'            recSetAuxActualizar1.Open " select * from CO_Diario where  cod_Comp = " & Cont_Comp & " ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
'                recSetAuxActualizar1.AddNew
'                recSetAuxActualizar1!tipo_comp = "PAC"
'                recSetAuxActualizar1!Cod_Comp = Cont_Comp
'            Else
'                If (Not recSetAuxActualizar1.BOF) Then recSetAuxActualizar1.MoveFirst
'            End If
'                recSetAuxActualizar1!d_cuenta = v_Cuenta
'                recSetAuxActualizar1!D_Nombre = v_NombreCta
'                recSetAuxActualizar1!d_subcta1 = v_SubCta1
'                recSetAuxActualizar1!d_subcta2 = v_SubCta2
'                recSetAuxActualizar1!d_Aux1 = v_Aux1
'                recSetAuxActualizar1!d_Aux2 = v_Aux2
'                recSetAuxActualizar1!d_Aux3 = v_Aux3
''************* CONTABILIZA AUXILIAARES DEBITO
'                Select Case v_Aux1
'                Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
'
'                Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
'                Case Else
'                End Select
'''****************** finaliza sesion de auxiliares
'                recSetAuxActualizar1!h_Aux1 = v_H_Aux1
'                recSetAuxActualizar1!h_Aux2 = v_H_Aux2
'                recSetAuxActualizar1!h_Aux3 = v_H_Aux3
''************* CONTABILIZA AUXILIAARES CREDITO
'
'                Select Case v_H_Aux1
'                Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
'
'                Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
''recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga
'
'                Case Else
'                End Select
'''****************** finaliza sesion de auxiliares
'
'                recSetAuxActualizar1!h_cuenta = v_H_Cuenta
'                recSetAuxActualizar1!H_Nombre = v_H_NombCta
'                recSetAuxActualizar1!h_subcta1 = v_H_SubCta1
'                recSetAuxActualizar1!h_subcta2 = v_H_SubCta2
'                recSetAuxActualizar1!d_montobs = recSetAuxcomp!monto_bolivianos
'                recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_Dolares
'                recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
'
'                recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_bolivianos
'                recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_Dolares
'                recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
'                recSetAuxActualizar1.Update
'        End If
'    Else
'           MsgBox "No esta asociado a ninguna fuente ...  "
'    End If
'    recSetAuxcomp.MoveNext
'MsgBox "Contabilizacion exitosa...... ", vbOKOnly, "Contabilizacion"
'Wend
'db.CommitTrans
'
'
'    Set recSetAuxcomp = New ADODB.Recordset
'    recSetAuxcomp.CursorLocation = adUseClient
'    If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'
'    Set recSetAuxActualizar = New ADODB.Recordset
'    If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'
'    Set recSetAuxActualizar1 = New ADODB.Recordset
'    If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'
'    Set recSetPartida = New ADODB.Recordset
'    recSetPartida.CursorLocation = adUseClient
'    If recSetPartida.State = 1 Then recSetPartida.Close
'
'
'Exit Sub
'errorPag:
'db.RollbackTrans
'MsgBox "No se contabiliz� ... "
'
'End Sub


Private Sub CmdAprueba_Click()
'On Error GoTo error_grabar:

'DUL
Screen.MousePointer = vbHourglass
Dim rsNada As ADODB.Recordset
Dim auxCodPago As String
Dim auxOrg As String
Dim Encontro As Boolean
auxCodPago = AdoRegularizacion.Recordset!codigo_pago
auxOrg = AdoRegularizacion.Recordset!org_codigo
Set DtcRegularizacion.DataSource = rsNada
Set TxtJustificacion.DataSource = rsNada
EntrarAdo = False
' Actualizar los Datos en el GRID
    Set rsRegularizacion = New ADODB.Recordset
'    If GlUsuario = "FFL001" Or GlUsuario = "F_FLORES" Or GlUsuario = "F_Flores" Or GlUsuario = "J_CAMACHO" Or GlUsuario = "J_Camacho" Or GlUsuario = "M_YA�EZ" Then

'gvi        QueryInicial = "select * from pagos where (tipo_comp = 'DAC') and (estado_compromiso='N' or estado_devengado='N' or estado_pagado='N' or estado_reversion_total='N' or estado_devolucion='N' or estado_anulado='N') "

        'QueryInicial = "select * from pagos where (estado_pagado='S')"
        rsRegularizacion.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
        rsRegularizacion.Sort = "codigo_pago"
        CmdAprueba.Enabled = True
'    Else
'        QueryInicial = "select * from pagos where (tipo_comp = 'DAC') and (usr_usuario = '" & Trim(Label7.Caption) & "')"
'        rsRegularizacion.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'        rsRegularizacion.Sort = "codigo_pago"
'        CmdAprueba.Enabled = False
'        swA = "2"
'    End If
'    Set AdoRegularizacion.Recordset = rsRegularizacion
'    Set DtcRegularizacion.DataSource = AdoRegularizacion.Recordset
'
'    If rsRegularizacion.RecordCount > 0 Then
'        AdoRegularizacion.Recordset.MoveNext
'        AdoRegularizacion.Recordset.MovePrevious
'    End If
''
'rsRegularizacion.Requery
Set AdoRegularizacion.Recordset = rsRegularizacion
Encontro = False
'aqui gvi
AdoRegularizacion.Recordset.Find "codigo_pago = '" & auxCodPago & "'", , adSearchForward, 1
While Not Encontro And Not AdoRegularizacion.Recordset.EOF
  'If (AdoRegularizacion.Recordset!codigo_pago = auxCodPago) And (AdoRegularizacion.Recordset!Org_Codigo = auxOrg) Then
  If (AdoRegularizacion.Recordset!org_codigo = auxOrg) Then
    Encontro = True
  Else
    AdoRegularizacion.Recordset.MoveNext
    AdoRegularizacion.Recordset.Find "codigo_pago = '" & auxCodPago & "'", , adSearchForward
  End If
Wend
MsgBox auxCodPago
Set DtcRegularizacion.DataSource = AdoRegularizacion.Recordset
Set TxtJustificacion.DataSource = AdoRegularizacion.Recordset
EntrarAdo = True
AdoRegularizacion.Recordset.Move 0
Screen.MousePointer = vbDefault
'DUL

'Verifica ppto
Set rsDet = New ADODB.Recordset
If rsDet.State = 1 Then rsDet.Close
rsDet.Open "select * from pago_detalle where codigo_pago= " & AdoRegularizacion.Recordset!codigo_pago & " and org_codigo= '" & AdoRegularizacion.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'  Print rsDet.RecordCount
If rsDet.RecordCount > 0 Then
  ppto2 = "0"
  Set rsPpto = New ADODB.Recordset
  If rsPpto.State = 1 Then rsPpto.Close
  rsPpto.Open "select * from fo_formulacion_gasto where pro_programa='" & rsDet("pro_programa") & "' and pro_subprograma='" & rsDet("pro_subprograma") & "' and pro_proyecto='" & rsDet("pro_proyecto") & "' and pro_actividad='" & rsDet("pro_actividad") & "' and par_codigo='" & rsDet("par_codigo") & "' and org_codigo= '" & rsDet("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
  If rsPpto.RecordCount > 0 Then
     ppto2 = "1"
     If AdoRegularizacion.Recordset("estado_compromiso") = "N" Then
      If (rsPpto("FGS_vigente") - rsPpto("FGS_compromiso") < rsDet("monto_total")) Then
        MsgBox "NO EXISTE PRESUPUESTO PARA COMPROMETER ...", vbOKOnly, "ERROR"
        '----ini se desabilita el control solo por un tiempo SOLICITUD IMA�A
        Exit Sub 'gvi
        rsPpto("fgs_compromiso") = rsPpto("fgs_compromiso") + rsDet("monto_total") 'gvi
        rsPpto.Update 'gvi
        '----fin se desabilita el control solo por un tiempo SOLICITUD IMA�A
      Else
        rsPpto("fgs_compromiso") = rsPpto("fgs_compromiso") + rsDet("monto_total")
        rsPpto.Update
      End If
     End If
     If AdoRegularizacion.Recordset("estado_devengado") = "N" Then
      ' Para Validar lo Devengado
      ' Modificado por Gerardo Rodriguez
      Dim RsDevenga As ADODB.Recordset
      Dim RsCompro As ADODB.Recordset
      Dim GlSqlAux As String
      Set RsDevenga = New ADODB.Recordset
      Set RsCompro = New ADODB.Recordset
      ' Para ACCESS
      'GlSQLAux = "SELECT IIF(ISNULL(SUM(monto_Total)), 0, SUM(monto_Total)) AS TotalDevengado " & _
      '           "FROM pagos, pago_Detalle " & _
      '           "WHERE (pagos.codigo_pago = pago_detalle.codigo_pago) AND (pagos.Tipo_formulario = 'DEV') AND (pagos.estado_devengado = 'S') AND (pagos.Nro_Comprobante_Anterior = '" & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & "')"
      ' Para SQL
'      GlSqlAux = "SELECT ISNULL(SUM(monto_Total), 0) AS TotalDevengado " & _
'                 "FROM pagos, pago_Detalle " & _
'                 "WHERE (pagos.codigo_pago = pago_detalle.codigo_pago) AND (pagos.Tipo_formulario = 'DEV') AND (pagos.estado_devengado = 'S') AND (pagos.Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & ") AND (pagos.org_codigo = '" & AdoRegularizacion.Recordset!Org_Codigo & "')"
      'corregido por jorge
      GlSqlAux = "SELECT ISNULL(SUM(monto_bolivianos), 0) AS TotalDevengado " & _
                 "FROM pagos " & _
                 "WHERE (pagos.Tipo_formulario = 'DEV') AND (pagos.estado_devengado = 'S') AND (pagos.Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!nro_comprobante_anterior & ") AND (pagos.org_codigo = '" & AdoRegularizacion.Recordset!org_codigo & "')"
                 
      RsDevenga.Open GlSqlAux, db, adOpenStatic
      GlSqlAux = "SELECT Sum(Monto_Total) AS MontoTotal FROM pago_detalle " & _
                 "WHERE (pago_detalle.Codigo_Pago = " & AdoRegularizacion.Recordset!nro_comprobante_anterior & ") AND (pago_detalle.org_codigo = '" & AdoRegularizacion.Recordset!org_codigo & "') "
      RsCompro.Open GlSqlAux, db, adOpenStatic
      If (RsCompro!MontoTotal < RsDevenga!TotalDevengado + rsDet("monto_total")) Then
        MsgBox "La Suma de lo DEVENGADO excede el Monto del Compromiso del Comprobante '" & AdoRegularizacion.Recordset!nro_comprobante_anterior & "'.", vbExclamation + vbOKOnly, "ERROR" '"La estructura presupuestaria NO es v�lida o NO EXISTE PRESUPUESTO "
        Exit Sub
      Else
        rsPpto("fgs_devengado") = rsPpto("fgs_devengado") + rsDet("monto_total")
        rsPpto.Update
      End If

    '      If (rsPpto("FGS_compromiso") - rsPpto("FGS_devengado") < rsDet("monto_total")) Then
    '        MsgBox "NO EXISTE PRESUPUESTO PARA DEVENGAR ", vbOKOnly, "ERROR"  '"La estructura presupuestaria NO es v�lida o NO EXISTE PRESUPUESTO "
    '        Exit Sub
    '      Else
    '        rsPpto("fgs_devengado") = rsPpto("fgs_devengado") + rsDet("monto_total")
    '        rsPpto.Update
    '      End If
     End If
     'Verificar por que ...
     If AdoRegularizacion.Recordset("estado_pagado") = "N" Then
        If (rsPpto("FGS_compromiso") - rsPpto("FGS_pagado") < rsDet("monto_total")) Then
           MsgBox "NO EXISTE PRESUPUESTO", vbOKOnly, "ERROR"  '"La estructura presupuestaria NO es v�lida o NO EXISTE PRESUPUESTO "
           Exit Sub
        Else
           rsPpto("fgs_pagado") = rsPpto("fgs_pagado") + rsDet("monto_total")
           rsPpto.Update
        End If
     End If
     'Verificar por que ... hasta aqui ...
   Else
     MsgBox "La estructura presupuestaria NO es v�lida", vbOKOnly, "ERROR"
     Exit Sub
  End If
  If rsPpto.State = 1 Then rsPpto.Close
  '************************

  ' VER NOW

   Dim VARCONTA As String
   VARCONTA = "B"
   If AdoRegularizacion.Recordset("estado_devengado") = "S" Then
      MsgBox "El registro ya est� APROBADO ..."
   Else
      Total_MontoBolivianos = 0
      Total_MontoDolares = 0
      Total_Deduccion = 0
      Total_SaldoBolivianos = 0

      Set rsDet = New ADODB.Recordset
      If rsDet.State = 1 Then rsDet.Close
      rsDet.Open "select * from pago_detalle where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo= '" & AdoRegularizacion.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
      Set AdoDetalle.Recordset = rsDet
      '      ****************** del detalle
      While Not rsDet.EOF
            Total_MontoBolivianos = Total_MontoBolivianos + rsDet("monto_total")
            Total_MontoDolares = Total_MontoDolares + rsDet("monto_dolares")
            Total_Deduccion = Total_Deduccion + rsDet("deducciones")
            Total_SaldoBolivianos = Total_SaldoBolivianos + rsDet("monto_total")
            rsDet.MoveNext
      Wend
      '      ****************** del detalle
      If rsDet.RecordCount > 0 Then
           Set rscorrelativo = New ADODB.Recordset
           '   Contabilizando
           swA = "1"
           If AdoRegularizacion.Recordset("estado_compromiso") = "N" Then 'Compromiso
              AdoRegularizacion.Recordset("estado_compromiso") = "S"
              AdoRegularizacion.Recordset("estado_aprobacion") = "N"
              AdoRegularizacion.Recordset("Deducciones") = 1
              VARCONTA = "C"
           End If
           If AdoRegularizacion.Recordset("estado_devengado") = "N" Then 'Devengado
              AdoRegularizacion.Recordset("estado_devengado") = "S"
              AdoRegularizacion.Recordset("estado_aprobacion") = "N"
              AdoRegularizacion.Recordset("Deducciones") = 1
              VARCONTA = "D"
           End If
           If AdoRegularizacion.Recordset("estado_tesoreria") = "N" Then 'Regularizacion
              AdoRegularizacion.Recordset("estado_tesoreria") = "S"
              AdoRegularizacion.Recordset("estado_aprobacion") = "N"
              AdoRegularizacion.Recordset("Deducciones") = 1
              VARCONTA = "G"
           End If
           If AdoRegularizacion.Recordset("estado_pagado") = "N" Then 'Pagos
              AdoRegularizacion.Recordset("estado_pagado") = "S"
              AdoRegularizacion.Recordset("estado_aprobacion") = "N"
              AdoRegularizacion.Recordset("Deducciones") = 1
              VARCONTA = "P"
           End If
          If AdoRegularizacion.Recordset("estado_devolucion") = "N" Then 'Devolucion
              AdoRegularizacion.Recordset("estado_devolucion") = "S"
              AdoRegularizacion.Recordset("Deducciones") = -1
              VARCONTA = "V"
          End If
          If AdoRegularizacion.Recordset("estado_reversion_total") = "N" Then 'Reversion Total
              AdoRegularizacion.Recordset("estado_reversion_total") = "S"
              AdoRegularizacion.Recordset("Deducciones") = -1
              VARCONTA = "R"
          End If
          If AdoRegularizacion.Recordset("estado_reversion_parcial") = "N" Then 'Reversion Parcial
              AdoRegularizacion.Recordset("estado_reversion_parcial") = "S"
              AdoRegularizacion.Recordset("Deducciones") = -1
              VARCONTA = "L"
          End If
          If AdoRegularizacion.Recordset("estado_anulado") = "N" Then 'Anulado
              AdoRegularizacion.Recordset("estado_anulado") = "S"
              AdoRegularizacion.Recordset("Deducciones") = -1
              VARCONTA = "A"
          End If
          
          'Datos de aprueba
          '      ****************** del detalle
            AdoRegularizacion.Recordset("monto_bolivianos") = Round(Total_MontoBolivianos, 2)
            AdoRegularizacion.Recordset("monto_dolares") = Round(Total_MontoDolares, 2)
'            AdoRegularizacion.Recordset("Deducciones") = Total_Deduccion
            AdoRegularizacion.Recordset("liquido_pagar") = Round(Total_SaldoBolivianos, 2)
          '      ****************** del detalle
            AdoRegularizacion.Recordset("usuario_aprueba") = Label7.Caption
            AdoRegularizacion.Recordset("fecha_aprueba") = Date
            AdoRegularizacion.Recordset("hora_aprueba") = Format(Time, "hh:mm:ss")
            
            AdoRegularizacion.Recordset.Update
            '***GABY*****Regularizacion
            If (AdoRegularizacion.Recordset("estado_compromiso") = "S") And (AdoRegularizacion.Recordset("estado_devengado") = "S") And (AdoRegularizacion.Recordset("estado_pagado") = "S") Then
              VARCONTA = "G"
            End If
            
            '****
               
            swRefresca = 1
            Marca = AdoRegularizacion.Recordset.AbsolutePosition
            ' Variables para actualizar estados
            'Dim GlSqlAux As String
            Dim RsAux As ADODB.Recordset

' GABY revisar desde aqu�

            If VARCONTA = "D" Then 'Devengado
                'Cmd_ContaConf AdoRegularizacion.Recordset!Codigo_Pago, AdoRegularizacion.Recordset!Org_Codigo, AdoRegularizacion.Recordset!Ges_Gestion
                'Frm_Cont_Mat.Show vbModal
            End If

            If VARCONTA = "G" Then 'Regularizacion
'                Cmd_ContaConf AdoRegularizacion.Recordset!codigo_pago, AdoRegularizacion.Recordset!Org_Codigo, AdoRegularizacion.Recordset!ges_gestion
'                Dim rsayuda As ADODB.Recordset
'                Set rsayuda = New ADODB.Recordset
'                If rsayuda.State = 1 Then rsayuda.Close
'                rsayuda.Open "select codigo_pago,codigo_pago_detalle,org_codigo,ges_gestion from pago_detalle where codigo_pago=" & AdoRegularizacion.Recordset!codigo_pago & " and org_codigo='" & AdoRegularizacion.Recordset!Org_Codigo & "' and ges_gestion='" & AdoRegularizacion.Recordset!ges_gestion & "'", db, adOpenKeyset, adLockReadOnly
'
'                Cmd_Pagado rsayuda!codigo_pago, rsayuda!codigo_pago_detalle, rsayuda!Org_Codigo, rsayuda!ges_gestion
            End If
            If VARCONTA = "R" Then 'Reversion Total
'                Reversion_DAC (AdoRegularizacion.Recordset)
'                GlSqlAux = "SELECT * " & _
'                           "FROM Pagos " & _
'                           "WHERE Org_Codigo = '" & AdoRegularizacion.Recordset!Org_Codigo & "' AND Codigo_Pago = '" & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & "'"
'                Set RsAux = New ADODB.Recordset
'                RsAux.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
'                If RsAux.RecordCount > 0 Then
'                  If IIf(IsNull(RsAux!estado_compromiso), "", RsAux!estado_compromiso) = "S" Then
'                    RsAux!estado_compromiso = "R"
'                  End If
'                  If IIf(IsNull(RsAux!estado_devengado), "", RsAux!estado_devengado) = "S" Then
'                    RsAux!estado_devengado = "R"
'                  End If
'                  RsAux.Update
'                End If
            End If
            If VARCONTA = "A" Then 'Anulacion
                Anulacion_DAC (AdoRegularizacion.Recordset)
                GlSqlAux = "SELECT * " & _
                           "FROM Pagos " & _
                           "WHERE Org_Codigo = '" & AdoRegularizacion.Recordset!org_codigo & "' AND Codigo_Pago = '" & AdoRegularizacion.Recordset!nro_comprobante_anterior & "'"
                Set RsAux = New ADODB.Recordset
                RsAux.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
                If RsAux.RecordCount > 0 Then
                  If IIf(IsNull(RsAux!estado_pagado), "", RsAux!estado_pagado) = "S" Then
                    RsAux!estado_pagado = "L"
                  End If
                  RsAux.Update
                End If
            End If
            If VARCONTA = "V" Then 'Devolucion
                Devolucion_PAC_DAC (AdoRegularizacion.Recordset)
                GlSqlAux = "SELECT * " & _
                           "FROM Pagos " & _
                           "WHERE Org_Codigo = '" & AdoRegularizacion.Recordset!org_codigo & "' AND Codigo_Pago = '" & AdoRegularizacion.Recordset!nro_comprobante_anterior & "'"
                Set RsAux = New ADODB.Recordset
                RsAux.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
                If RsAux.RecordCount > 0 Then
                  If IIf(IsNull(RsAux!estado_compromiso), "", RsAux!estado_compromiso) = "S" Then
                    RsAux!estado_compromiso = "V"
                  End If
                  If IIf(IsNull(RsAux!estado_devengado), "", RsAux!estado_devengado) = "S" Then
                    RsAux!estado_devengado = "V"
                  End If
                  If IIf(IsNull(RsAux!estado_pagado), "", RsAux!estado_pagado) = "S" Then
                    RsAux!estado_pagado = "V"
                  End If
                  RsAux.Update
                End If
            End If
'
' GABY revisar hasta aqu�
'
            AdoRegularizacion.Recordset.AbsolutePosition = Marca
'
            swRefresca = 0
            LblTitulo.Caption = ""
          '         FraOpciones.Visible = True
          '         FraGrabarCancelar.Visible = False
          '         DtcRegularizacion.Enabled = True
          ' RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR
            
       Else
             MsgBox "No se puede APROBAR un registro sin detalle ..."
       End If
    End If
  Else
    MsgBox "No se puede APROBAR un registro sin detalle ..."
 End If
MsgBox auxCodPago
Exit Sub

error_grabar:
MsgBox Err.Number & " " & Err.Description

End Sub

Private Sub CmdBorrar_Click()
'On Error GoTo elimina:
 If AdoRegularizacion.Recordset("estado_devengado") = "N" Or AdoRegularizacion.Recordset("estado_compromiso") = "N" Or AdoRegularizacion.Recordset("estado_reversion_total") = "N" Or AdoRegularizacion.Recordset("estado_devolucion") = "N" Or AdoRegularizacion.Recordset("estado_anulado") = "N" Then
    sino = MsgBox("Est� Seguro de Anular este Registro", vbYesNo + vbQuestion, "Atenci�n")
    If sino = vbYes Then
         'El estado de equivocado u errado se lo har� en el momento de aprobado, en otro caso no
         
         MsgBox "Se eliminaran los datos de este registro y el detalle correspondiente", vbYesNo + vbQuestion, "Atenci�n"
         
         Marca = AdoRegularizacion.Recordset.AbsolutePosition
         If sino = vbYes Then
            Set rsD = New ADODB.Recordset
            X = AdoRegularizacion.Recordset("ges_gestion")
            Y = AdoRegularizacion.Recordset("org_codigo")
            z = AdoRegularizacion.Recordset("codigo_pago")
            rsD.Open "select * from pago_detalle where ges_gestion='" & X & "' and org_codigo='" & Y & "' and codigo_pago= '" & z & "' ", db, adOpenKeyset, adLockOptimistic
            'While Not rsD.EOF
            '   rsD.Delete
            '   rsD.MoveNext
            'Wend
            rsD.Close
            
            AdoRegularizacion.Recordset.AbsolutePosition = Marca
            Set rsM = New ADODB.Recordset
            rsM.Open "select * from pagos where ges_gestion='" & X & "' and org_codigo='" & Y & "' and codigo_pago= '" & z & "' ", db, adOpenKeyset, adLockOptimistic
            If rsM("estado_compromiso") = "N" Then rsM("estado_compromiso") = "E"
            If rsM("estado_devengado") = "N" Then rsM("estado_devengado") = "E"
            If rsM("estado_pagado") = "N" Then rsM("estado_pagado") = "E"
            If rsM("estado_reversion_total") = "N" Then rsM("estado_reversion_total") = "E"
            If rsM("estado_devolucion") = "N" Then rsM("estado_devolucion") = "E"
            If rsM("estado_anulado") = "N" Then rsM("estado_anulado") = "E"
            rsM.Update
            rsM.Close
            
            'rsM.Open "select * from pagos where usr_usuario = '" & Trim(Label7.Caption) & "' ", db, adOpenKeyset, adLockOptimistic
            'If GlUsuario = "FFL001" Or GlUsuario = "F_FLORES" Or GlUsuario = "F_Flores" Or GlUsuario = "J_CAMACHO" Or GlUsuario = "J_Camacho" Or GlUsuario = "M_YA�EZ" Then
                rsM.Open "select * from pagos where ((tipo_comp = 'DAC') and (estado_compromiso='N' or estado_devengado='N' or estado_pagado='N' or estado_reversion_total='N' or estado_devolucion='N' or estado_anulado='N')) order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
                'rsM.Open "select * from pagos where tipo_comp = 'DAC' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
                CmdAprueba.Enabled = True
            'Else
            '    rsM.Open "select * from pagos where (tipo_comp = 'DAC') and usr_usuario = '" & Trim(Label7.Caption) & "' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
            '    CmdAprueba.Enabled = True
            '    swA = "2"
            'End If
            Set AdoRegularizacion.Recordset = rsM
            Set DtcRegularizacion.DataSource = AdoRegularizacion
            
            DtcRegularizacion.Refresh
            'Set DtcRegularizacion.DataSource = AdoRegularizacion
         End If
    End If
 Else
    MsgBox "No se puede Anular un registro APROBADO ..."
 End If
    
'Exit Sub
'elimina:
'    'MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdBorrarDetalle_Click()
   If AdoRegularizacion.Recordset("estado_devengado") = "N" Or AdoRegularizacion.Recordset("estado_compromiso") = "N" Then
        
        If AdoDetalle.Recordset.RecordCount > 0 Then 'DtGDetalle.Columns(0) <> "" Then
            sino = MsgBox("Est� seguro de eliminar este registro", vbYesNo + vbQuestion, "Atenci�n")
            If sino = vbYes Then
                AdoDetalle.Recordset.Delete
            End If
        Else
            MsgBox "No existe registro para eliminar", vbCritical + vbInformation, "Validaci�n de Datos"
        End If
    Else
       MsgBox "No se puede modificar un registro APROBADO ..."
   End If
  msgSalir = "1"
End Sub

Private Sub CmdBuscar_Click()
'dul
Dim ClVBusca As CompBusquedas.ClBuscaEnGridPropio 'Componente de busquedas

Dim ClBuscaSec As CompBusquedas.ClBuscaSecuencialEnRS
PosibleApliqueFiltro = False
Dim rsNada As ADODB.Recordset
Dim GrSqlAux As String

  GrFrmOpciones.OptOpciones(1).Caption = "Buscar en Cabecera"
  GrFrmOpciones.OptOpciones(2).Caption = "Buscar en Cabecera y Detalle"
  GrFrmOpciones.Show vbModal
  If GrFrmOpciones.POpcionElegida <> 0 Then
    If GrFrmOpciones.POpcionElegida = 1 Then
      Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexi�n = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = DtcRegularizacion
      ClBuscaGrid.QueryUtilizado = QueryInicial
      Set ClBuscaGrid.RecordsetTrabajo = AdoRegularizacion.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
    Else
      Set ClVBusca = New CompBusquedas.ClBuscaEnGridPropio 'DUL: Instancio Componente de Busqueda
      'Dll con Ventana de b�squeda
      GrSqlAux = "SELECT pg.ges_gestion, pg.codigo_pago, pg.org_codigo, " & _
                 "pg.Nro_Comprobante_Anterior, pg.estado_compromiso, " & _
                 "pg.estado_devengado, pg.estado_pagado, " & _
                 "pg.estado_reversion_total, pg.estado_devolucion, " & _
                 "pg.estado_anulado, Pd.par_codigo, " & _
                 "Pd.Pro_programa, Pd.Pro_subprograma, " & _
                 "Pd.Pro_proyecto, Pd.Pro_actividad, " & _
                 "Pd.cta_codigo, Pd.numero_cheque_trf, " & _
                 "Pd.cta_codigo_destino, " & _
                 "Pd.codigo_beneficiario, Pd.monto_total, " & _
                 "Pd.monto_Dolares " & _
                 "FROM Pagos as Pg INNER JOIN Pago_Detalle as Pd " & _
                 "ON (Pg.Ges_gestion = pd.ges_gestion)and(Pg.codigo_pago = Pd.codigo_pago)and(Pg.org_codigo = Pd.org_codigo) "
                 'NO PUEDE IR : "ORDER BY pg.codigo_pago"
      Set ClVBusca.Conexi�n = db
      ClVBusca.QueryUtilizado = GrSqlAux
      ClVBusca.FiltrosMultiples = True
      ClVBusca.OcultarPrimero = True
      ClVBusca.AliasColumnas = "ges_gestion-Cmbte-Org-Anterior-C-D-P-R-V-A-PARTIDA-" & _
                               "PRO-SUB-PRY-ACT-CTA_ORIGEN-Cheque_Trf-CTA_DESTINO-" & _
                               "BENEFICIARIO-MONTO_Bs-MONTO_US"
      ClVBusca.Tama�oCampos = "1-0.6-0.5-0.4-0.3-0.3-0.3-0.3-0.3-0.3-0.8-" & _
                              "0.4-0.4-0.4-0.4-0.7-0.6-0.7-" & _
                              "0.8-0.6-0.8"
      ClVBusca.Ejecutar
      'Si eligio alguno
      If ClVBusca.ElegidoCol1 <> "" Then
        'Desconecta para que no se disparen los eventos del ado
        If PosibleApliqueFiltro Then
          If rsRegularizacion.State = 1 Then rsRegularizacion.Close
          rsRegularizacion.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
          rsRegularizacion.Sort = "codigo_pago"
          Set AdoRegularizacion.Recordset = rsRegularizacion
          Set DtcRegularizacion.DataSource = AdoRegularizacion.Recordset
        End If
        Set DtcRegularizacion.DataSource = rsNada
        Set TxtJustificacion.DataSource = rsNada
        EntrarAdo = False
        'Buscar Registro
        Set ClBuscaSec = New CompBusquedas.ClBuscaSecuencialEnRS
        Set ClBuscaSec.Recordset = AdoRegularizacion.Recordset
        ClBuscaSec.Campo1 = "ges_Gestion"
        ClBuscaSec.ValorCampo1 = ClVBusca.ElegidoCol1
        ClBuscaSec.Campo2 = "org_codigo"
        ClBuscaSec.ValorCampo2 = ClVBusca.ElegidoCol3
        ClBuscaSec.Campo3 = "codigo_pago"
        ClBuscaSec.ValorCampo3 = ClVBusca.ElegidoCol2
        If Not ClBuscaSec.Ejecutar Then
          MsgBox "Error en b�squeda...", vbInformation + vbOKOnly, "Atenci�n"
        End If
        Set ClBuscaSec = Nothing
        
        Set DtcRegularizacion.DataSource = AdoRegularizacion.Recordset
        Set TxtJustificacion.DataSource = AdoRegularizacion.Recordset
        EntrarAdo = True
        If AdoRegularizacion.Recordset.RecordCount > 0 Then
          AdoRegularizacion.Recordset.Move 0
        End If
      End If
    End If
  End If
End Sub

Private Sub CmdCalculo_Click()
   If TxtTipoCambio.Text <> "" Then
    ' Validamos los Montos
        If Not IsNumeric(TxtMontoFuente.Text) Then
            MsgBox "El Monto en Bolivianos debe ser un Valor Num�rico V�lido.", vbExclamation + vbOKOnly, "Validaci�n de Datos"
            Exit Sub
        End If
        If Not IsNumeric(TxtMontoDolares.Text) Then
            MsgBox "El Monto en D�lares debe ser un Valor Num�rico V�lido.", vbExclamation + vbOKOnly, "Validaci�n de Datos"
            Exit Sub
        End If
      If (Val(TxtMontoDolares.Text) > 0) And (Val(TxtMontoFuente.Text) = 0) Then
         TxtMontoFuente.Text = CDbl(TxtMontoDolares.Text) * CDbl(TxtTipoCambio.Text)
      End If
      If (Val(TxtMontoFuente.Text) > 0) And (Val(TxtMontoDolares.Text) = 0) Then
         TxtMontoDolares.Text = Round(CDbl(TxtMontoFuente.Text) / CDbl(TxtTipoCambio.Text), 2)
      End If
'      If (TxtMontoDolares.Text > 0) And (TxtMontoFuente.Text > 0) Then
'         TxtMontoFuente.Text = CDbl(TxtMontoDolares.Text) * CDbl(TxtTipoCambio.Text)
'         TxtMontoDolares.Text = Round(CDbl(TxtMontoFuente.Text) / CDbl(TxtTipoCambio.Text), 2)
'      End If
   Else
      MsgBox "Introducir tipo de cambio", vbCritical + vbExclamation, "Validaci�n de datos"
      Exit Sub
   End If
   TxtSaldo.Text = Val(TxtMontoFuente.Text) - Val(TxtDeducciones.Text)
   If TxtSaldo.Text = 0 And TxtMontoFuente > 0 Then
        'TxtMontoFuente.Text = CDbl(TxtMontoDolares.Text) * CDbl(TxtTipoCambio.Text)
        TxtMontoDolares.Text = CDbl(TxtMontoFuente.Text) / CDbl(TxtTipoCambio.Text)
        TxtSaldo.Text = Val(TxtMontoFuente.Text) - Val(TxtDeducciones.Text)
   End If
End Sub

Private Sub CmdCancelaCopia_Click()
    FraCopiaRegistro.Visible = False
    FraCopiaRegistro.Enabled = False
End Sub
Private Sub CmdCancelar_Click()
On Error Resume Next
'On Error GoTo error_cancelar:
  LblTitulo.Caption = ""
  FraMaestro.Enabled = False
  
  Set rsNada = New ADODB.Recordset
  Set DtcRegularizacion.DataSource = rsNada

  FraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
 
  FraOpciones.Visible = True
  FraMaestro.Visible = True
  FraMaestro.Enabled = False
  FraGrabarCancelar.Visible = False
  CmdAdicionar.Enabled = True
  CmdBorrar.Enabled = True
  CmdSalir.Enabled = True
  DtcRegularizacion.Enabled = True
  If swGrabaCopia = 1 Then
      FraCopiaRegistro.Visible = False
      FraCopiaRegistro.Enabled = False
      swGrabaCopia = 0
  End If
  swgraba = "0"
    
    AdoRegularizacion.Recordset.CancelUpdate
    AdoRegularizacion.Recordset.Requery
    
'    Set rsRegularizacion = rsauxiliar
'    Set AdoRegularizacion.Recordset = rsauxiliar
    Set DtcRegularizacion.DataSource = AdoRegularizacion
    DtcOrg.Enabled = True
    DtcDesOrg.Enabled = True
    DTcFte.Enabled = True
    DtcFteDes.Enabled = True
    
    DtcTipoDes.Visible = False
    TxtTipoReg.Visible = True

db.RollbackTrans
Exit Sub
error_cancelar:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdCancelarBeneficiario_Click()
    rsBeneficiario.CancelUpdate
    FraBeneficiario.Visible = False
End Sub

Private Sub CmdCmptesDev_Click()
'    LblCabecera.Caption = "COMPROBANTES DE DEVOLUCIONES"
'    FraDev.Visible = True
'    FraCopiar.Visible = False
'    Grid_Devoluciones
End Sub
Private Sub CmdCancelarBusqueda_Click()
    'FraBusqueda.Visible = False
    
    
    'Restaurando el grid
    Set rsRegularizacion = New ADODB.Recordset
    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
        rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and (estado_compromiso='N' or estado_devengado='N' or estado_pagado='N' or estado_reversion_total='N' or estado_devolucion='N' or estado_anulado='N') order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
        'rsRegularizacion.Open "select * from pagos where tipo_comp = 'DAC' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
        If rsRegularizacion.RecordCount > 0 Then
            Set DtcRegularizacion.DataSource = AdoRegularizacion
            Set AdoRegularizacion.Recordset = rsRegularizacion
        End If
End Sub

Private Sub CmdCopiar_Click()
If DtcRegularizacion.Columns(0) <> "" Then
    If rsDetalle.RecordCount <= 0 Then
      MsgBox "No se puede Copiar un Comprobante incompleto.", vbExclamation + vbOKOnly, "Atenci�n"
      Exit Sub
    End If
    CmdAdicionar.Enabled = False
    CmdBorrar.Enabled = False
    CmdSalir.Enabled = False
    CmdGrabar.Visible = True
    FraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraMaestro.Enabled = True
    'LblTitulo.Caption = "Para Confirmar la COPIA, haga CLICK en GRABAR . . . "
    LblTitulo.Caption = ". . . "
    FraMaestro.Enabled = False
    DtcRegularizacion.Enabled = False
    FraCopiaRegistro.Visible = True
    FraCopiaRegistro.Enabled = True
    
    'Datos para el nuevo registro
    'corregir DE aqui ..............
    '    TxtTR.Text = DtcTipoCod
    'corregir aqui ..............
    TxtNS.Text = txtNroSolicitud
    TxtCO.Text = TxtCodigoOrden
    TxtNC.Text = TxtComprobante
    
    '    If DtcTipoDes.Text = "DEVOLUCION" Or DtcTipoDes.Text = "REVERSION TOTAL" Or DtcTipoDes.Text = "ANULACION" Or DtcTipoDes.Text = "DEVENGADO" Then
    '        TxtTR.Text = DtcTipoCod
    '        ComprobanteAnterior = TxtComprobante.Text
    '    Else
    '        TxtTR.Text = DtcTipoCod
    '        TxtNCA.Text = TxtComprobanteAnterior.Text
    '    End If
    
    ' nO EST� BIEN ...........
    
    If TxtTR.Text = "DVL" Or TxtTR.Text = "RVT" Or TxtTR.Text = "ANL" Or TxtTR.Text = "DEV" Then
        TxtNCA.Text = TxtComprobanteAnterior.Text
        ComprobanteAnterior = TxtComprobante.Text
    Else
        'por solo copia
        TxtTR.Text = DtcTipoCod
        ComprobanteAnterior = TxtComprobante.Text
        ANTERIOR = TxtComprobante.Text
    End If
    DtCDR.Text = DtcDcu
    DtCDRD.Text = DtcDcuDes
    DtCUT.Text = DtCUnidad
    DtCUTD.Text = DtCDesUnidad
    DtCFF.Text = DTcFte
    DtcFFD.Text = DtcFteDes
    DtCOF.Text = DtcOrg
    DtcOFD.Text = DtcDesOrg
    DtcC.Text = DtcCat
    DtcCD.Text = DtcCatDes
    DtcConv2.Text = DtcConv
    TxtJ.Text = TxtJustificacion
    TxtFR.Text = DtpFecha
    
    swGrabaCopia = 1
    AuxCopia = "C"
    'Copiar detalle para devolucion declaradas en variables globales
    Dim vector(50) As String
    'Utilizando vector para almacenar los varios registros de detalle
    'vector(1) = DtCFF.Text
    
    vgFteCodigo = DtCFF.Text
    vgCodigoPartida = DtGDetalle.Columns(0).Value
    vgPrograma = DtGDetalle.Columns(1)
    vgSubPrograma = DtGDetalle.Columns(2)
    vgProyecto = DtGDetalle.Columns(3)
    vgActividad = DtGDetalle.Columns(4)
    vgCtaOrigen = DtGDetalle.Columns(5)
    vgNroChequeOTransferencia = DtGDetalle.Columns(6)
    vgCtaDestino = DtGDetalle.Columns(7)
    vgCodBeneficiario = DtGDetalle.Columns(8)
    If DtGDetalle.Columns(9) <> "" Then vgMontoTotal = CCur(DtGDetalle.Columns(9).Value)
    'vgDeducciones = DtGDetalle.Columns(10)
    'vgMontoBolvianos = DtGDetalle.Columns(11)
    'If DtGDetalle.Columns(11) <> "" Then vgMB = CCur(DtGDetalle.Columns(11).Value)
    If DtGDetalle.Columns(10) <> "" Then vgTipoCambio = CCur(DtGDetalle.Columns(10).Value)
    If DtGDetalle.Columns(11) <> "" Then vgMontoDolares = CCur(DtGDetalle.Columns(11).Value)
    vgOrgCodigo = DtcOrg.Text
    vgGesGestion = Year(DtpFecha.Text)
    vgCodigoPago = TxtComprobante.Text
    vgCodigoPagoDetalle = "1"
    
    'FraCopiaRegistro.Enabled = False
    FraCopiaRegistro.Enabled = True
    
 Else
    MsgBox "Falta detalle ", vbCritical + vbExclamation, "Validaci�n de Datos"
    Exit Sub
 End If
End Sub

Private Sub CmdDet_Click()
On Error Resume Next
        FraOpciones.Visible = False
        FraGrabarCancelar.Visible = False
        FraOpcionesDetalle.Visible = True
        FraMaestro.Visible = False
        FraDetalleG.Visible = True
        FraDetalleG.Enabled = False
        Frame5.Visible = False
        
        TxtCodigoOrdend.Text = TxtCodigoOrden.Text
        DtcRegularizacion.Visible = False
        AdoRegularizacion.Visible = False
        AdoDetalle.Enabled = True
        DtcRegularizacion.Enabled = False
        'Detalle
        CmdAgregarDetalle.Enabled = True
        CmdGrabaDetalle.Enabled = False
        CmdModificarDetalle.Enabled = True
        CmdBorrarDetalle.Enabled = True
        If AdoRegularizacion.Recordset("tipo_formulario") = "REG" Then
            Frame1.Visible = True
        Else
            Frame1.Visible = False
        End If
        If rsDetalle.State = 1 Then
            rsDetalle.Close
        End If
         Set rsDetalle = New ADODB.Recordset
         rsDetalle.Open "select * from pago_detalle where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo= '" & AdoRegularizacion.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
         Set DtGDetalle.DataSource = rsDetalle
         If rsDetalle.RecordCount > 0 Then
              Set AdoDetalle.Recordset = rsDetalle
         Else
              Set rsNada = New ADODB.Recordset
              Set AdoDetalle.Recordset = rsDetalle
         End If
' Validamos si ya esta aprobado
      With AdoRegularizacion
        If IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "S" Or _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "S" Or _
             IIf(IsNull(.Recordset!estado_pagado), "", .Recordset!estado_pagado) = "S" Or _
             IIf(IsNull(.Recordset!estado_reversion_total), "", .Recordset!estado_reversion_total) = "S" Or _
             IIf(IsNull(.Recordset!estado_devolucion), "", .Recordset!estado_devolucion) = "S" Or _
             IIf(IsNull(.Recordset!estado_anulado), "", .Recordset!estado_anulado) = "S" Or _
             IIf(IsNull(.Recordset!estado_compromiso), "", .Recordset!estado_compromiso) = "E" Or _
             IIf(IsNull(.Recordset!estado_devengado), "", .Recordset!estado_devengado) = "E" Then
             CmdAgregarDetalle.Enabled = False
             CmdModificarDetalle.Enabled = False
             CmdBorrarDetalle.Enabled = False
        Else
             CmdAgregarDetalle.Enabled = True
             CmdModificarDetalle.Enabled = True
             CmdBorrarDetalle.Enabled = True
        End If
      End With
         
End Sub

Private Sub CmdGrabaCopia_Click()
db.BeginTrans
On Error GoTo error_grabar:
        Set rscorrelativo = New ADODB.Recordset
        AdoRegularizacion.Recordset.AddNew
        If DtCOF.Text <> "" Then
            AdoRegularizacion.Recordset("org_codigo") = DtCOF.Text
        Else
            MsgBox "Introcudir Organismo Financiador", vbCritical + vbExclamation, "Validaci�n de datos"
            Exit Sub
        End If
         
        If DtCOF.Text = "111" Then  'TGN
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org111) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
                  rscorrelativo!Correl_Org111 = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "112" Then 'TGNP
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org112) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
                  rscorrelativo!Correl_Org112 = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "114" Then 'RECON
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org114) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
                  rscorrelativo!Correl_Org114 = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "344" Then 'UNICEF
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org4) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org4) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org4) + 1)
                  rscorrelativo!Correl_Org4 = CDbl(CDbl(rscorrelativo!Correl_Org4) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "381" Then  'FAD
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org5) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org5) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org5) + 1)
                  rscorrelativo!Correl_Org5 = Val(Val(rscorrelativo!Correl_Org5) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "411" Then  'BID
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org411) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
                  rscorrelativo!Correl_Org411 = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "415" Then  'IDA
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org415) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
                  rscorrelativo!Correl_Org415 = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "516" Then  'KFW
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org516) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
                  rscorrelativo!Correl_Org516 = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "541" Then  'ALEM
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org9) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org9) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org9) + 1)
                  rscorrelativo!Correl_Org9 = CDbl(CDbl(rscorrelativo!Correl_Org9) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "551" Then  'DIN
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org10) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org10) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org10) + 1)
                  rscorrelativo!Correl_Org10 = CDbl(CDbl(rscorrelativo!Correl_Org10) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "556" Then  'HOL
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org556) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
                  rscorrelativo!Correl_Org556 = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "565" Then  'SUE
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org565) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
                  rscorrelativo!Correl_Org565 = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "999" Then  'S/N
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org999) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
                  rscorrelativo!Correl_Org999 = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
                  rscorrelativo.Update
            End If
         End If
         
         If DtCOF.Text = "Org14" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org14) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org13) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org14) + 1)
                  rscorrelativo!Correl_Org14 = CDbl(CDbl(rscorrelativo!Correl_Org14) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org14 = 0
                rscorrelativo.Update
            End If
         End If
         
         
         If DtCOF.Text = "Org15" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org15) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
                  rscorrelativo!Correl_Org15 = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org15 = 0
                rscorrelativo.Update
            End If
         End If
         
           If DtCOF.Text = "Org16" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org16) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
                  rscorrelativo!Correl_Org16 = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org16 = 0
                rscorrelativo.Update
            End If
         End If
   
         If DtCOF.Text = "Org17" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org17) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
                  rscorrelativo!Correl_Org17 = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org17 = 0
                rscorrelativo.Update
            End If
         End If
   
         If DtCOF.Text = "Org18" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org18) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
                  rscorrelativo!Correl_Org18 = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org18 = 0
                rscorrelativo.Update
            End If
         End If
 ' Este codigo dependera de organismo financiador y del a�o
          
          If DtCUT.Text <> "" Then
            AdoRegularizacion.Recordset("uni_codigo") = "VEIPS" 'DtCUT.Text
            AdoRegularizacion.Recordset("codigo_unidad") = DtCUT.Text
          Else
            MsgBox "Falta Unidad T�cnica", vbCritical + vbInformation, "Validaci�n de datos"
            Exit Sub
          End If
        
         If TxtCO.Text <> "" Then
            AdoRegularizacion.Recordset("codigo_orden") = TxtCO.Text
         Else
            MsgBox "Introducir n�mero Orden", vbCritical + vbExclamation
            Exit Sub
         End If
         If TxtNS.Text <> "" Then
            AdoRegularizacion.Recordset("codigo_solicitud") = TxtNS.Text
         Else
            MsgBox "Introcudir N�mero de Solicitud", vbCritical + vbExclamation
            Exit Sub
         End If
         If DtCFF.Text <> "" Then
            AdoRegularizacion.Recordset("fte_codigo") = DtCFF.Text
         Else
            MsgBox "Introcudir Fte. de Financiamiento", vbCritical + vbExclamation, "Validaci�n de datos"
            Exit Sub
         End If
         
         AdoRegularizacion.Recordset("codigo_categoria") = DtcC.Text
         MsgBox TxtJ.Text
         If TxtJ.Text <> "" Then
            AdoRegularizacion.Recordset("justificacion") = TxtJ.Text
         Else
            MsgBox "Introducir Justificaci�n", vbCritical + vbExclamation, "Validaci�n de datos"
            Exit Sub
         End If
         AdoRegularizacion.Recordset("tipo_moneda") = "Bs." 'DtCTipoMoneda.Text
         AdoRegularizacion.Recordset("liquido_pagar") = "0" 'Val(TxtLiquido.Text)
         
         'Estados de aprobaci�n
         AdoRegularizacion.Recordset("liquido_pagar") = "0"
         AdoRegularizacion.Recordset("estado_compromiso") = "N"
         AdoRegularizacion.Recordset("estado_devengado") = "N"
         AdoRegularizacion.Recordset("estado_pagado") = "N"
         
         AdoRegularizacion.Recordset("estado_tesoreria") = "N"
         AdoRegularizacion.Recordset("estado_entregado") = "N"
         AdoRegularizacion.Recordset("estado_anulado") = "N"

        'Datos de seguimiento
         AdoRegularizacion.Recordset("ges_gestion") = Year(Now)
         AdoRegularizacion.Recordset("usr_usuario") = Label7.Caption
         AdoRegularizacion.Recordset("fecha_registro") = Date
         AdoRegularizacion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
         
         MsgBox AdoRegularizacion.Recordset("codigo_pago")
         MsgBox AdoRegularizacion.Recordset("org_codigo")
         AdoRegularizacion.Recordset.Update
  'FraCopiaRegistro.Visible = False
  FraCopiaRegistro.Visible = True
db.CommitTrans
Exit Sub
error_grabar:
MsgBox Err.Number & " " & Err.Description
db.RollbackTrans
End Sub
Private Sub CmdDev_Click()
'    DtGDevoluciones.Visible = False
'    DtcRegularizacion.Visible = True
'    FraDev.Visible = False
'    CmdDevolucion_Click
End Sub
Private Sub CmdDevolucion_Click()
  Set rsRegularizacion = New ADODB.Recordset
  If rsRegularizacion.State = 1 Then rsRegularizacion.Close
  rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC' or  tipo_comp = 'CYD') and estado_devengado='S' and estado_pagado='S' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
  CmdAprueba.Enabled = True
        If rsRegularizacion.RecordCount > 0 Then
            Set DtcRegularizacion.DataSource = AdoRegularizacion
            Set AdoRegularizacion.Recordset = rsRegularizacion
        Else
            MsgBox "No se encontraron registros", vbInformation, "Validaci�n de datos"
        End If
'  FraBusqueda.Visible = False
  FraMaestro.Enabled = True
  swDevolucion = "D"
End Sub
Private Sub CmdGrabaDetalle_Click()
On Error GoTo error_grabadetalle

    ' Validaciones
    If Val(DtCPartida.Text) < 10000 Then
            MsgBox "Introduzca C�digo de Partida", vbCritical + vbInformation, "Validaci�n de Datos"
            Exit Sub
    End If
    ' Validamos los Montos
    If CCur(IIf(Trim(TxtSaldo.Text) = "", 0, TxtSaldo.Text)) <= 0 Then
        MsgBox "El L�quido pagable debe ser un Monto Mayor a CERO.", vbExclamation + vbOKOnly, "Validaci�n de Datos"
        Exit Sub
    End If
    If Not IsNumeric(TxtMontoFuente.Text) Then
        MsgBox "El Monto en Bolivianos debe ser un Valor Num�rico V�lido.", vbExclamation + vbOKOnly, "Validaci�n de Datos"
        Exit Sub
    End If
    If Not IsNumeric(TxtMontoDolares.Text) Then
        MsgBox "El Monto en D�lares debe ser un Valor Num�rico V�lido.", vbExclamation + vbOKOnly, "Validaci�n de Datos"
        Exit Sub
    End If
  '-----------------------------
  Set rsPpto = New ADODB.Recordset
  If rsPpto.State = 1 Then rsPpto.Close
  rsPpto.Open "select * from fo_formulacion_gasto where pro_programa='" & TxtProgramad & "' and pro_subprograma='" & TxtSubprogramad & "' and pro_proyecto='" & TxtProyectod & "' and pro_actividad='" & TxtActividadd & "' and par_codigo='" & DtCPartida.Text & "' and org_codigo= '" & DtcOrg.Text & "'", db, adOpenKeyset, adLockOptimistic 'GVI
  If rsPpto.RecordCount > 0 Then
  Else
        MsgBox "La estructura presupuestaria NO es v�lida", vbOKOnly, "ERROR"
        If rsPpto.State = 1 Then rsPpto.Close
        Exit Sub
   End If
  If rsPpto.State = 1 Then rsPpto.Close
  '---------------------------
  db.BeginTrans
        Dim codigo_pago1 As String ' gvi
        Dim ges_gestion1 As String ' gvi
        codigo_pago1 = AdoRegularizacion.Recordset("codigo_pago") ' gvi
        ges_gestion1 = AdoRegularizacion.Recordset("ges_gestion") ' gvi
         
         AdoDetalle.Recordset("codigo_pago") = codigo_pago1
         AdoDetalle.Recordset("ges_gestion") = ges_gestion1
         AdoDetalle.Recordset("org_codigo") = DtcOrg.Text
         AdoDetalle.Recordset("codigo_pago_detalle") = AdoDetalle.Recordset.RecordCount
         AdoDetalle.Recordset("par_codigo") = DtCPartida.Text
         AdoDetalle.Recordset("Pro_programa") = TxtProgramad.Text
         AdoDetalle.Recordset("Pro_subprograma") = TxtSubprogramad.Text
         AdoDetalle.Recordset("Pro_proyecto") = TxtProyectod.Text
         AdoDetalle.Recordset("Pro_actividad") = TxtActividadd.Text
         AdoDetalle.Recordset("codigo_beneficiario") = dtcRuc.Text
         AdoDetalle.Recordset("monto_total") = Val(TxtMontoFuente.Text) 'MontoFuente = MontoBolivianos
         AdoDetalle.Recordset("monto_dolares") = Val(TxtMontoDolares)
         If AdoRegularizacion.Recordset("estado_reversion_total") = "N" Or AdoRegularizacion.Recordset("estado_devolucion") = "N" Or AdoRegularizacion.Recordset("estado_anulado") = "N" Then
            AdoDetalle.Recordset("Deducciones") = -1
          Else
            AdoDetalle.Recordset("Deducciones") = 1
         End If
         'AdoDetalle.Recordset("Deducciones") = Val(TxtDeducciones.Text)
         AdoDetalle.Recordset("saldo_bolivianos") = Val(TxtSaldo.Text)
         AdoDetalle.Recordset("ges_gestion") = Year(Date)
         AdoDetalle.Recordset("estado_aprobacion") = "N"
         AdoDetalle.Recordset("fecha_pago") = Date
         AdoDetalle.Recordset("fecha_registro") = Date
         AdoDetalle.Recordset("usr_usuario") = Label7.Caption
         AdoDetalle.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
         AdoDetalle.Recordset.Update
         AdoDetalle.Recordset.MoveNext
         AdoDetalle.Recordset.MovePrevious
         'DtGDetalle.Refresh
         
         CmdGrabaDetalle.Enabled = False
         CmdAgregarDetalle.Enabled = True
         CmdModificarDetalle.Enabled = True
         Label35.Caption = "."
  '*********
  db.CommitTrans
  msgSalir = "1"
  Exit Sub
error_grabadetalle:
   MsgBox Err.Number & " " & Err.Description
   db.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()
If swGrabaCopia = 1 Then
  Graba_Copia
'  swGrabaCopia = 0
Else
  'On Error GoTo error_grabar
  'validaciones
  
  If txtNroSolicitud.Text = "" Then
    MsgBox "Se requiere Nro. Formulario de Solicitud ...", vbCritical + vbExclamation
    Exit Sub
  End If
    
  If TxtCodigoOrden.Text = "" Then
    MsgBox "Se requiere N�mero de Documento de Respaldo ...", vbCritical + vbExclamation
    Exit Sub
  End If
  
  If DtCUnidad.Text = "" Then
    MsgBox "Se requiere Unidad T�cnica", vbCritical + vbInformation, "Validaci�n de datos"
    Exit Sub
  End If
  
  If DTcFte.Text = "" Then
    MsgBox "Se requiere Fuente de Financiamiento", vbCritical + vbExclamation, "Validaci�n de datos"
    Exit Sub
  End If

  If DtcOrg.Text = "" Then
    MsgBox "Introcudir Organismo Financiador", vbCritical + vbExclamation, "Validaci�n de datos"
    Exit Sub
  End If

  If DtcOrg.Text <> "" Then
    If swgraba = "1" Then
      'AdoRegularizacion.Recordset("org_codigo") = DtcOrg.Text
    Else
      Org3 = DtcOrg.Text
    End If
  Else
    MsgBox "Se requiere el organismo financiador ...", vbCritical + vbExclamation, "Validaci�n de datos"
    Exit Sub
  End If
  
'  If DtCUT.Text <> "" Then
'    'AdoRegularizacion.Recordset("uni_codigo") = TxtCodigoOrden.Text
'  Else
'    MsgBox "Falta Unidad T�cnica", vbCritical + vbInformation, "Validaci�n de datos"
'    Exit Sub
'  End If

  If TxtJustificacion.Text = "" Then
    MsgBox "Introducir Justificaci�n", vbCritical + vbExclamation, "Validaci�n de datos"
    Exit Sub
  End If

'**************  db.BeginTrans
  If sw2 = "1" Then
    Set rscorrelativo = New ADODB.Recordset
    
    If DtcOrg.Text = "111" Then  'TGN
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org111) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
        rscorrelativo!Correl_Org111 = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "112" Then 'TGNP
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org112) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
        rscorrelativo!Correl_Org112 = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
        rscorrelativo.Update
      End If
    End If
    
    If DtcOrg.Text = "114" Then 'RECON
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org114) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
        rscorrelativo!Correl_Org114 = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
        rscorrelativo.Update
      End If
    End If
    
    If DtcOrg.Text = "344" Then 'UNICEF
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org334) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org334) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org334) + 1)
        rscorrelativo!Correl_Org334 = CDbl(CDbl(rscorrelativo!Correl_Org334) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "381" Then  'FAD
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org381) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org381) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org381) + 1)
        rscorrelativo!Correl_Org381 = Val(Val(rscorrelativo!Correl_Org381) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "411" Then  'BID
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org411) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
        rscorrelativo!Correl_Org411 = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "415" Then  'IDA
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org415) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
        rscorrelativo!Correl_Org415 = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "516" Then  'KFW
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org516) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
        rscorrelativo!Correl_Org516 = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "541" Then  'ALEM
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org541) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org541) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org541) + 1)
        rscorrelativo!Correl_Org541 = CDbl(CDbl(rscorrelativo!Correl_Org541) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "551" Then  'DIN
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org551) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org551) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org551) + 1)
        rscorrelativo!Correl_Org551 = CDbl(CDbl(rscorrelativo!Correl_Org551) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "556" Then  'HOL
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org556) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
        rscorrelativo!Correl_Org556 = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "565" Then  'SUE
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org565) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
        rscorrelativo!Correl_Org565 = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "999" Then  'S/N
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org999) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
        rscorrelativo!Correl_Org999 = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "Org14" Then
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org14) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org13) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org14) + 1)
        rscorrelativo!Correl_Org14 = CDbl(CDbl(rscorrelativo!Correl_Org14) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "Org15" Then
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org15) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
        rscorrelativo!Correl_Org15 = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
        rscorrelativo.Update
      End If
    End If
         
    If DtcOrg.Text = "Org16" Then
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org16) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
        rscorrelativo!Correl_Org16 = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
        rscorrelativo.Update
      End If
    End If
    
    If DtcOrg.Text = "Org17" Then
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org17) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
        rscorrelativo!Correl_Org17 = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
        rscorrelativo.Update
      End If
    End If
   
    If DtcOrg.Text = "Org18" Then
      rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
      If Not IsNull(rscorrelativo!Correl_Org18) Then
        AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
        AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
        rscorrelativo!Correl_Org18 = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
        rscorrelativo.Update
      Else
        rscorrelativo!Correl_Org18 = 0
        rscorrelativo.Update
      End If
    End If
         
    'aqui exit transact 1
    If DtcOrg.Text <> "" Then
      If swgraba = "1" Then
        AdoRegularizacion.Recordset("org_codigo") = DtcOrg.Text
      Else
        Org3 = DtcOrg.Text
      End If
    Else
      MsgBox "Se requiere la unidad", vbCritical + vbExclamation, "Validaci�n de datos"
      Exit Sub
    End If
  End If
  ' Este codigo dependera de organismo financiador y del a�o

  'aqui exit transact 2
  If DtCUnidad.Text <> "" Then
    AdoRegularizacion.Recordset("uni_codigo") = DtCUnidad.Text
  Else
    MsgBox "Falta Unidad T�cnica", vbCritical + vbInformation, "Validaci�n de datos"
    Exit Sub
  End If
  
  If TxtCodigoOrden.Text <> "" Then
    AdoRegularizacion.Recordset("codigo_orden") = TxtCodigoOrden.Text
  Else
    MsgBox "Introducir n�mero de documento de respaldo", vbCritical + vbExclamation
    Exit Sub
  End If
  
  If txtNroSolicitud.Text <> "" Then
    AdoRegularizacion.Recordset("codigo_solicitud") = txtNroSolicitud.Text
  Else
    MsgBox "Introcudir dato", vbCritical + vbExclamation
    Exit Sub
  End If
  AdoRegularizacion.Recordset("uni_codigo") = "VEIPS"
  'aqui exit transact 3
  If DTcFte.Text <> "" Then
    AdoRegularizacion.Recordset("fte_codigo") = DTcFte.Text
  Else
    MsgBox "Introcudir Fte. de Financiamiento", vbCritical + vbExclamation, "Validaci�n de datos"
    Exit Sub
  End If
  AdoRegularizacion.Recordset("codigo_categoria") = DtcCat.Text
  AdoRegularizacion.Recordset("codigo_convenio") = DtcConv.Text
  'aqui exit transact 4
  If TxtJustificacion.Text <> "" Then
    AdoRegularizacion.Recordset("justificacion") = TxtJustificacion.Text
  Else
    MsgBox "Introducir Justificaci�n", vbCritical + vbExclamation, "Validaci�n de datos"
    Exit Sub
  End If
  AdoRegularizacion.Recordset("tipo_moneda") = "Bs." 'DtCTipoMoneda.Text
  AdoRegularizacion.Recordset("liquido_pagar") = "0" 'Val(TxtLiquido.Text)
     
  'Estados de aprobaci�n
  If DtcTipoCod.Text <> "" Then
    If DtcTipoCod.Text = "COM" Then
      AdoRegularizacion.Recordset("estado_compromiso") = "N"
      AdoRegularizacion.Recordset("estado_devengado") = ""
      'AdoRegularizacion.Recordset("nro_comprobante_anterior") = AdoRegularizacion.Recordset("codigo_pago")
    End If
    If DtcTipoCod.Text = "DEV" Then
      AdoRegularizacion.Recordset("estado_compromiso") = ""
      AdoRegularizacion.Recordset("estado_devengado") = "N"
      'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = ANTERIOR
    End If
    If DtcTipoCod.Text = "CYD" Then
      AdoRegularizacion.Recordset("estado_compromiso") = "N"
      AdoRegularizacion.Recordset("estado_devengado") = "N"
      'AdoRegularizacion.Recordset("nro_comprobante_anterior") = AdoRegularizacion.Recordset("codigo_pago")
    End If
    If DtcTipoCod.Text = "PAG" Then
      AdoRegularizacion.Recordset("estado_pagado") = "N"
    End If
    If DtcTipoCod.Text = "REG" Then
      AdoRegularizacion.Recordset("estado_compromiso") = "N"
      AdoRegularizacion.Recordset("estado_devengado") = "N"
      AdoRegularizacion.Recordset("estado_pagado") = "N"
    End If
  Else
    MsgBox "Introducir Tipo de Registro", vbCritical + vbExclamation, "Validaci�n de datos"
    Exit Sub
  End If
  AdoRegularizacion.Recordset("tipo_formulario") = DtcTipoCod.Text
  AdoRegularizacion.Recordset("estado_aprobacion") = "X"
  
  'Datos de seguimiento
  AdoRegularizacion.Recordset("tipo_comp") = "DAC"
  AdoRegularizacion.Recordset("ges_gestion") = Year(Date)
  AdoRegularizacion.Recordset("usr_usuario") = Label7.Caption
  AdoRegularizacion.Recordset("fecha_registro") = Date
  AdoRegularizacion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
     
  DtcTipoDes.Visible = False
  TxtTipoReg.Visible = True

  FraOpciones.Visible = True
  FraMaestro.Visible = True
  FraMaestro.Enabled = False
  
  CmdAdicionar.Enabled = True
  CmdBorrar.Enabled = True
  CmdSalir.Enabled = True
  
  AdoRegularizacion.Recordset.Update
  
  'AdoRegularizacion.Recordset.Requery
  LblTitulo.Caption = ""
  FraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  DtcRegularizacion.Enabled = True
  'If AdoRegularizacion.Recordset.RecordCount > 0 Then
  '  AdoRegularizacion.Recordset.MovePrevious
  '  AdoRegularizacion.Recordset.MoveNext
  'End If
  
  '********Apuntando al registro de devolucion
  
'*************  db.CommitTrans
  
End If

'Copia un registro de regularizaci�n
DtcOrg.Enabled = True
DtcDesOrg.Enabled = True
DTcFte.Enabled = True
DtcFteDes.Enabled = True
    
'If DtcTipoDes.Text <> "DEVOLUCION" Or DtcTipoDes.Text = "REVERSION TOTAL" Or DtcTipoDes.Text = "ANULACION" Then
'            If rsRegularizacion.State = 1 Then rsRegularizacion.Close
'            rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and usr_usuario = '" & Trim(Label7.Caption) & "' and codigo_pago='" & TxtComprobante.Text & "' and org_codigo='" & DtcOrg.Text & "'  order by codigo_pago='" & vgOrgCodigo & "' ", db, adOpenKeyset, adLockOptimistic
'                If rsRegularizacion.RecordCount > 0 Then
'                '    AdoRegularizacion.Recordset.MoveNext
'                '    AdoRegularizacion.Recordset.MovePrevious
'                End If
'End If
'Refresca el grid con todos los datos
If DtcTipoDes.Text = "DEVOLUCION" Or DtcTipoDes.Text = "REVERSION TOTAL" Or DtcTipoDes.Text = "ANULACION" Then

  'db.BeginTrans
  
  ' Revisar Jorge
  '            If rsRegularizacion.State = 1 Then rsRegularizacion.Close
  '            rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and usr_usuario = '" & Trim(Label7.Caption) & "' and codigo_pago='" & TxtComprobante.Text & "' and org_codigo='" & DtcOrg.Text & "'  order by codigo_pago", db, adOpenKeyset, adLockOptimistic
  '                If rsRegularizacion.RecordCount > 0 Then
  '                    Set DtcRegularizacion.DataSource = AdoRegularizacion
  '                    Set AdoRegularizacion.Recordset = rsRegularizacion
  '                End If
                
  'Actualizando el campo de comprometido-devengado o pagado a su anterior estado
  Set rsAnterior = New ADODB.Recordset
  If rsAnterior.State = 1 Then rsAnterior.Close
  rsAnterior.Open "select * from pagos where codigo_pago='" & TxtComprobanteAnterior.Text & "' and org_codigo='" & DtcOrg.Text & "'  order by codigo_pago", db, adOpenKeyset, adLockOptimistic
  If rsAnterior.RecordCount > 0 Then
    Select Case TIPOFORMULARIO
      Case "ANULACION"
              rsAnterior("tipo_formulario") = "ANL"
      Case "COMPROMISO"
              rsAnterior("tipo_formulario") = "COM"
      Case "COMPROMISO Y DEVENGADO"
              rsAnterior("tipo_formulario") = "CYD"
      Case "DEVENGADO"
              rsAnterior("tipo_formulario") = "DEV"
      Case "DEVOLUCION"
              rsAnterior("tipo_formulario") = "DVL"
      Case "REGULARIZACION"
              rsAnterior("tipo_formulario") = "REG"
      Case "REVERSION PARCIAL"
              rsAnterior("tipo_formulario") = "RVP"
      Case "REVERSION TOTAL"
              rsAnterior("tipo_formulario") = "RVT"
    End Select
    rsAnterior.Update
  End If
  'If TIPOFORMULARIO = "COMPROMISO Y DEVENGADO" Then
  ' rsAnterior("tipo_formulario") = "CYD"
  ' rsAnterior.Update
  'End If

'  db.CommitTrans
End If
'---------OJO ------------------
'REVISAR CON CELIA

'If DtcTipoDes.Text <> "DEVOLUCION" And DtcTipoDes.Text <> "REVERSION TOTAL" And DtcTipoDes.Text <> "ANULACION" Then
'    Set rsRegularizacion = New ADODB.Recordset
'    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
'    rsRegularizacion.Open "select * from pagos where tipo_comp = 'DAC' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'    CmdAprueba.Enabled = True
'    If rsRegularizacion.RecordCount > 0 Then
'    Set DtcRegularizacion.DataSource = AdoRegularizacion
'    Set AdoRegularizacion.Recordset = rsRegularizacion
'    End If
'End If

    
'db.CommitTrans
''If DtcTipoDes.Text = "DEVOLUCION" Then
''        If rsRegularizacion.State = 1 Then rsRegularizacion.Close
''        rsRegularizacion.Open "select * from pagos where tipo_comp = 'DAC' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
''        CmdAprueba.Enabled = True
''        Set AdoRegularizacion.Recordset = rsRegularizacion
''        Set DtcRegularizacion.DataSource = AdoRegularizacion
''        FraCopiaRegistro.Visible = False
''End If
Exit Sub
error_grabar:
MsgBox Err.Number & " " & Err.Description

'db.RollbackTrans
      
End Sub
Private Sub CmdGrabarBeneficiario_Click()
    If TxtCodigoBeneficiario.Text <> "" Then
        rsBeneficiario!Codigo_beneficiario = TxtCodigoBeneficiario.Text
    Else
        MsgBox "Introducir codigo de beneficiario", vbCritical + vbInformation, "Validadci�n de datos"
    End If
    If TxtDenominacionBeneficiario.Text <> "" Then
        rsBeneficiario!denominacion_beneficiario = TxtDenominacionBeneficiario.Text
    Else
        MsgBox "Introducir nombre del beneficiario", vbCritical + vbInformation, "Validadci�n de datos"
    End If

    If CmbTipoBeneficiario.Text = "Proveedor" And CmbTipoBeneficiario.Text <> "" Then
        rsBeneficiario!Tipo_Beneficiario = "R"
    Else
        rsBeneficiario!Tipo_Beneficiario = "C"
    End If
    'Datos de seguimiento
    rsBeneficiario!Usr_Usuario = Label7.Caption
    rsBeneficiario!fecha_registro = Date
    rsBeneficiario!Hora_Registro = Format(Time, "hh:mm:ss")
    rsBeneficiario.Update
    rsBeneficiario.Close
    
    Set rsBeneficiario = New ADODB.Recordset
    
      rsBeneficiario.Open "select * from fc_beneficiario", db, adOpenKeyset, adLockOptimistic
      Set AdoRuc.Recordset = rsBeneficiario
      rsBeneficiario.MoveFirst
    FraBeneficiario.Visible = False
End Sub

Private Sub CmdImprimir_Click()
'    Dim Report As New CrtComprobante
On Error GoTo error_grabar:
Screen.MousePointer = vbHourglass
        Set rscorrelativo = New ADODB.Recordset
        Set rsRepCab = New ADODB.Recordset

        db.Execute "DELETE from pagos_rep"
        If rsRepCab.State = 1 Then rsRepCab.Close
        rsRepCab.Open "select * from pagos_rep ", db, adOpenKeyset, adLockOptimistic
        rsRepCab.AddNew
        rsRepCab("Maquina") = GlMaquina
        rsRepCab("codigo_pago") = TxtComprobante.Text
        rsRepCab("ges_gestion") = Year(Now)
        rsRepCab("org_codigo") = DtcOrg.Text
        rsRepCab("codigo_unidad") = DtCUnidad.Text
        rsRepCab("nro_comprobante_anterior") = TxtComprobanteAnterior.Text
        rsRepCab("codigo_documento") = DtcDcu.Text
        rsRepCab("codigo_orden") = TxtCodigoOrden.Text
        rsRepCab("codigo_solicitud") = txtNroSolicitud.Text
        rsRepCab("tipo_formulario") = DtcTipoCod.Text
        rsRepCab("fecha_egreso") = CDate(DtpFecha)
        rsRepCab("uni_codigo") = "VEIPS"
        rsRepCab("fte_codigo") = DTcFte.Text
        rsRepCab("codigo_categoria") = DtcCat.Text
        rsRepCab("justificacion") = TxtJustificacion.Text
        rsRepCab("tipo_moneda") = "Bs." 'DtCTipoMoneda.Text
        rsRepCab("monto_bolivianos") = IIf(IsNull(AdoRegularizacion.Recordset("monto_bolivianos")), 0, Round(AdoRegularizacion.Recordset("monto_bolivianos"), 2))
        rsRepCab("monto_dolares") = IIf(IsNull(AdoRegularizacion.Recordset("monto_dolares")), 0, Round(AdoRegularizacion.Recordset("monto_dolares"), 2))
        rsRepCab("Deducciones") = IIf(IsNull(AdoRegularizacion.Recordset("Deducciones")), 1, AdoRegularizacion.Recordset("Deducciones"))
        rsRepCab("liquido_pagar") = AdoRegularizacion.Recordset("liquido_pagar")
        LiteralCry = Str(Round(rsRepCab("monto_bolivianos"), 2))
        
        If AdoRegularizacion.Recordset("tipo_formulario") = "REG" And AdoRegularizacion.Recordset("estado_compromiso") = "S" And AdoRegularizacion.Recordset("estado_devengado") = "S" And AdoRegularizacion.Recordset("estado_pagado") = "S" Then
            rsRepCab("estado_tesoreria") = "A"
           Else
            If AdoRegularizacion.Recordset("estado_compromiso") = "S" Then
               rsRepCab("estado_compromiso") = "A"
            End If
            If AdoRegularizacion.Recordset("estado_devengado") = "S" Then
               rsRepCab("estado_devengado") = "A"
            End If
            If AdoRegularizacion.Recordset("estado_pagado") = "S" Then
               rsRepCab("estado_pagado") = "A"
            End If
            ''
            'Devoluci�n
            If AdoRegularizacion.Recordset("estado_devolucion") = "S" Then
               rsRepCab("estado_devolucion") = "A"
            End If
            'Anulaci�n
            If AdoRegularizacion.Recordset("estado_anulado") = "S" Then
               rsRepCab("estado_anulado") = "A"
            End If
            'Reversi�n Total
            If AdoRegularizacion.Recordset("estado_reversion_total") = "S" Then
               rsRepCab("estado_reversion_total") = "A"
            End If
            'Reversi�n Parcial
            If AdoRegularizacion.Recordset("estado_reversion_parcial") = "S" Then
               rsRepCab("estado_reversion_parcial") = "A"
            End If
            If AdoRegularizacion.Recordset("estado_compromiso") = "N" Then
               rsRepCab("estado_compromiso") = "S"
            End If
            ''
            If AdoRegularizacion.Recordset("estado_devengado") = "N" Then
               rsRepCab("estado_devengado") = "S"
            End If
            If AdoRegularizacion.Recordset("estado_pagado") = "N" Then
               rsRepCab("estado_pagado") = "S"
            End If
            If AdoRegularizacion.Recordset("estado_devolucion") = "N" Then
               rsRepCab("estado_devolucion") = "S"
            End If
            'Anulaci�n
            If AdoRegularizacion.Recordset("estado_anulado") = "N" Then
               rsRepCab("estado_anulado") = "S"
            End If
            'Reversi�n Total
            If AdoRegularizacion.Recordset("estado_reversion_total") = "N" Then
               rsRepCab("estado_reversion_total") = "S"
            End If
            'Reversi�n Parcial
            If AdoRegularizacion.Recordset("estado_reversion_parcial") = "N" Then
               rsRepCab("estado_reversion_parcial") = "S"
            End If
        
        End If
      '            If DtcTipoDes.Text = "DEVOLUCION" Then
      '               RsRepCab("estado_devolucion") = TxtComprobante.Text
      '            End If
       rsRepCab.Update
       LblTitulo.Caption = "."
         
       Set rsRepDet = New ADODB.Recordset
       db.Execute "DELETE from pago_detalle_rep"
       
       If rsRepDet.State = 1 Then rsRepDet.Close
       rsRepDet.Open "select * from pago_detalle_rep ", db, adOpenKeyset, adLockOptimistic
       'rsRepDet.Open "select * from pago_detalle_rep ", db, adOpenKeyset, adLockOptimistic
'       While Not rsRepDet.EOF And rsRepDet.RecordCount > 0
'                     rsRepDet.Delete
'                     rsRepDet.MoveNext
'       Wend
        
       If Not IsNull(AdoRegularizacion.Recordset("codigo_pago")) And Not IsNull(AdoRegularizacion.Recordset("org_codigo")) Then
           Set rsDetalle = New ADODB.Recordset
           rsDetalle.Open "select * from pago_detalle where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
      'revvvvvvvvvv ---------
        If AdoRegularizacion.Recordset("estado_devengado") = "S" Then
      '           AdoRegularizacion.Recordset("estado_aprobacion") = "N"
      '           AdoRegularizacion.Recordset.Update
              End If
        '''RepComprobante.Show vbModal
        
        If rsDetalle("monto_total") <> 0 Then
      '-----------
           'Set DtGDetalle.DataSource = rsDetalle
          If rsDetalle.RecordCount > 0 Then
          While Not rsDetalle.EOF
            rsRepDet.AddNew
            rsRepDet("Maquina") = GlMaquina
            rsRepDet("codigo_pago") = rsDetalle("codigo_pago")
            rsRepDet("ges_gestion") = rsDetalle("ges_gestion")
            rsRepDet("org_codigo") = rsDetalle("org_codigo")
            rsRepDet("codigo_pago_detalle") = rsDetalle("codigo_pago_detalle")
            rsRepDet("par_codigo") = rsDetalle("par_codigo")
            rsRepDet("pro_programa") = rsDetalle("pro_programa")
            rsRepDet("Pro_subprograma") = rsDetalle("Pro_subprograma")
            rsRepDet("Pro_proyecto") = rsDetalle("Pro_proyecto")
            rsRepDet("Pro_actividad") = rsDetalle("Pro_actividad")
            If rsDetalle("codigo_beneficiario") <> "" Then
                rsRepDet("codigo_beneficiario") = rsDetalle("codigo_beneficiario")
              Else
                rsRepDet("codigo_beneficiario") = "-"
            End If
            rsRepDet("monto_total") = Round(rsDetalle("monto_total"), 2)
            If AdoRegularizacion.Recordset("monto_bolivianos") = 0 Then
              LiteralCry = Str(Round(rsDetalle("monto_total"), 2))
            End If
            rsRepDet("monto_dolares") = Round(rsDetalle("monto_dolares"), 2)
            rsRepDet("tipo_cambio") = rsDetalle("tipo_cambio")
            rsRepDet("Deducciones") = rsDetalle("Deducciones")
            rsRepDet("saldo_bolivianos") = rsDetalle("saldo_bolivianos")
            rsRepDet("literal") = Literal(LiteralCry) + "  Bolivianos"
            rsRepDet.Update
            rsDetalle.MoveNext
            
          Wend
          End If
        End If
        
        'Report.lblFecha.SetText (Literal(LiteralCry) + "  Bolivianos")
'        RepComprobante.CRViewer1.ReportSource = Report
'        RepComprobante.CRViewer1.ViewReport
        
' JORGE
'    MsgBox "EL USUARIO NO TIENE ACCESO AL SERVIDOR SQL ..."
'    Exit Sub
        'RepComprobante.Show vbModal
        Dim iResult As Integer
'        Cry.Reset
        Cry.ReportFileName = "c:\Saf-2000\FormsPresupuesto\Dise�adores\CrtComprobantePpto.rpt"
        'Cry.SelectionFormula = {fv_comprobante2.Maquina} = '"& glmaQUINA &"'"
        iResult = Cry.PrintReport
        If iResult <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
        End If
' JORGE
       Else
        MsgBox "No se registr� el detalle del comprobante ..."
       End If
  Screen.MousePointer = vbDefault
Exit Sub
error_grabar:
  Screen.MousePointer = vbDefault
MsgBox Err.Number & " " & Err.Description

End Sub

Private Sub CmdModificar_Click()
    If AdoRegularizacion.Recordset("estado_devengado") = "N" Or AdoRegularizacion.Recordset("estado_compromiso") = "N" Or AdoRegularizacion.Recordset("estado_reversion_total") = "N" Or AdoRegularizacion.Recordset("estado_devolucion") = "N" Or AdoRegularizacion.Recordset("estado_anulado") = "N" Then
        DtpFecha.Enabled = False
        CmdAdicionar.Enabled = False
        CmdBorrar.Enabled = False
        CmdSalir.Enabled = False
        CmdGrabar.Visible = True
        FraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        FraMaestro.Enabled = True
        LblTitulo.Caption = "MODIFICANDO . . . "
        DtcRegularizacion.Enabled = False
        sw2 = "2"
        swA = "2"
        DtcOrg.Enabled = False
        DtcDesOrg.Enabled = False
        DTcFte.Enabled = False
        DtcFteDes.Enabled = False
        TxtTipoReg.Enabled = False
    Else
        MsgBox "No se puede modificar un registro APROBADO ..."
    End If
        
End Sub

Private Sub CmdModificarDetalle_Click()
    If AdoRegularizacion.Recordset("estado_devengado") = "N" Or AdoRegularizacion.Recordset("estado_compromiso") = "N" Or AdoRegularizacion.Recordset("estado_reversion_total") = "N" Or AdoRegularizacion.Recordset("estado_devolucion") = "N" Then
       FraDetalleG.Enabled = True
       Label35.Caption = "MODIFICANDO DETALLE . . ."
       CmdGrabaDetalle.Enabled = True
       CmdAgregarDetalle.Enabled = False
       CmdModificarDetalle.Enabled = False
       CmdBorrarDetalle.Enabled = False
       swModificaDetalle = 2 'Editando detalle
       TxtDeducciones.Text = "0"
       TxtDeducciones.Enabled = False
    Else
       MsgBox "No se puede modificar un registro APROBADO ..."
    End If
  msgSalir = "1"
End Sub

Private Sub CmdNuevoBeneficiario_Click()
   FraBeneficiario.Visible = True
   Set rsBeneficiario = New ADODB.Recordset
   rsBeneficiario.Open "select * from fc_beneficiario", db, adOpenKeyset, adLockOptimistic
   
   TxtCodigoBeneficiario.Text = ""
   TxtDenominacionBeneficiario.Text = ""
   CmbTipoBeneficiario.Text = ""
   rsBeneficiario.AddNew
End Sub

Private Sub CmdOrdenar_Click()
'Buscar . . .
        '''    If ValidaCriterio(CmbCampo.Text, CmbOperador.Text, TxtValor.Text) = 2 Then
        '''        If (Not rsRegularizacion.BOF) Then
        '''            rsRegularizacion.MoveFirst
        '''            rsRegularizacion.Find CmbCampo.Text & " " & CmbOperador.Text & " '" & TxtValor.Text & "'", , adSearchForward
        '''            CmdOrdenar.Enabled = True
        '''        End If
        '''    Else
        '''        MsgBox ErrCriterio, vbExclamation, "Error ..."
        '''    End If
Dim cadena_busqueda As String
    cadena_busqueda = ""
'    If CmbCampo = "ges_gestion" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "codigo_pago" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "org_codigo" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "tipo_comp" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "Nro_Comprobante_Anterior" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "fecha_egreso" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + " = " + "#" + TxtValor + "#"
'    End If
'    If CmbCampo = "estado_devolucion" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "estado_anulado" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "estado_comprometido" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "estado_reversion_total" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    If CmbCampo = "estado_reversion_parcial" Then
'        cadena_busqueda = "pagos." + CmbCampo.Text + CmbOperador + "'" + TxtValor + "'"
'    End If
'    'Realizar la busqueda dado un criterio
'    Set rsRegularizacion = New ADODB.Recordset
'    If cadena_busqueda <> "" Then
'        rsRegularizacion.Open "select * from pagos where " & cadena_busqueda & " ", db, adOpenKeyset, adLockOptimistic
'        If rsRegularizacion.RecordCount > 0 Then
'            Set DtcRegularizacion.DataSource = rsRegularizacion
'            Set AdoRegularizacion.Recordset = rsRegularizacion
'        Else
'            Set DtcRegularizacion.DataSource = rsNada
'        End If
'    Else
'        MsgBox "Coloque datos"
'    End If
'    FraBusqueda.Visible = False
'
'' Filtrar . . .
''    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
''    'esta bien
''    If CmbCampo.Text <> "" And CmbOperador.Text <> "" And "'" & TxtValor.Text & "'" <> "" Then
''        If GlUsuario = "FFL001" Or GlUsuario = "jgc001" Then
''            rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & " order by codigo_pago", db, adOpenStatic, adLockReadOnly
'''            AdoRegularizacion.Recordset.Open "select * from pagos where (tipo_comp = 'DAC') and " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & " order by codigo_pago", db, adOpenKeyset, adLockOptimistic
''            CmdAprueba.Enabled = True
''        Else
''            rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and usr_usuario = '" & Trim(Label7.Caption) & "' AND " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & "order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'''            AdoRegularizacion.Recordset.Open "select * from pagos where (tipo_comp = 'DAC') and usr_usuario = '" & Trim(Label7.Caption) & "' AND " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & "order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
''            CmdAprueba.Enabled = False
''            swA = "2"
''        End If
''        Set DtcRegularizacion.DataSource = AdoRegularizacion
''        Set AdoRegularizacion.Recordset = rsRegularizacion
'''        AdoRegularizacion.Recordset.Requery
''        rsRegularizacion.Requery
''        If rsRegularizacion.RecordCount = 0 Then
''            MsgBox "La Seleci�n NO tiene registros ... "
''            If rsRegularizacion.State = 1 Then rsRegularizacion.Close
''            If GlUsuario = "FFL001" Or GlUsuario = "jgc001" Then
''                rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & " order by codigo_pago", db, adOpenStatic, adLockReadOnly
''                CmdAprueba.Enabled = True
''            Else
''                rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and usr_usuario = '" & Trim(Label7.Caption) & "' AND " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & "order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
''                CmdAprueba.Enabled = False
''                swA = "2"
''            End If
''            'rsRegularizacion.Open "select * from pagos where estado_compromiso = 'S' or estado_compromiso = 'N' or estado_compromiso='E' or estado_tesoreria='N' order by codigo_pago ", db, adOpenStatic, adLockReadOnly
''            Set AdoRegularizacion.Recordset = rsRegularizacion
''            Set DtcRegularizacion.DataSource = rsRegularizacion
''            rsRegularizacion.Requery
''        End If
''    Else
''        MsgBox ErrCriterio, vbExclamation, "ERROR"
''        If rsRegularizacion.State = 1 Then rsRegularizacion.Close
''        If GlUsuario = "FFL001" Or GlUsuario = "jgc001" Then
''            rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & " order by codigo_pago", db, adOpenStatic, adLockReadOnly
''            CmdAprueba.Enabled = True
''        Else
''            rsRegularizacion.Open "select * from pagos where (tipo_comp = 'DAC') and usr_usuario = '" & Trim(Label7.Caption) & "' AND " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'" & "order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
''            CmdAprueba.Enabled = False
''            swA = "2"
''        End If
''        Set AdoRegularizacion.Recordset = rsRegularizacion
''        Set DtcRegularizacion.DataSource = rsRegularizacion
''        rsRegularizacion.Requery
''    End If
'    FraBusqueda.Visible = False
End Sub

Private Sub CmdProyecto_Click()
   FraProyecto.Visible = True
   
      'Set rsProyecto = New ADODB.Recordset
      'rsProyecto.Open "select pro_programa as Programa, pro_subprograma as Subprograma, pro_proyecto as Proyecto,pro_Actividad as Actividad,pro_descripcion_larga as Nombre_del_Proyecto  from fc_estructura_programatica ", db, adOpenKeyset, adLockOptimistic
      'rsProyecto.Open "select * from fc_estructura_programatica ", db, adOpenKeyset, adLockOptimistic
      'Set AdoProyecto.Recordset = rsProyecto
      'If AdoProyecto.Recordset.RecordCount > 0 Then
      '      Set DtGProyecto.DataSource = rsProyecto
      'End If
      
End Sub

Private Sub CmdReversion_Click()
    Set rsRegularizacion = New ADODB.Recordset
    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
    'rsRegularizacion.Open "select * from pagos where tipo_comp = 'DAC' and estado_compromiso='S' and estado_devengado='S' and estado_pagado='S' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
    rsRegularizacion.Open "select * from pagos where (tipo_formulario = 'COM' or  tipo_formulario = 'CYD' or  tipo_formulario = 'DEV') and (estado_devengado='S' OR estado_pagado='S' OR estado_compromiso='S') order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
    CmdAprueba.Enabled = True
    If rsRegularizacion.RecordCount > 0 Then
        Set DtcRegularizacion.DataSource = AdoRegularizacion
        Set AdoRegularizacion.Recordset = rsRegularizacion
    Else
        MsgBox "No existen datos", vbInformation, "Validaci�n de datos"
    End If
'    FraBusqueda.Visible = False
    FraMaestro.Enabled = True
    swDevolucion = "R"
End Sub

Private Sub CmdSalir_Click()

   'If AdoRegularizacion.Recordset.State = 1 Then AdoRegularizacion.Recordset.Close
   'If AdoDetalle.Recordset.State = 1 Then AdoDetalle.Recordset.Close
   If AdoCategoria.Recordset.State = 1 Then AdoCategoria.Recordset.Close
   If AdoCuenta.Recordset.State = 1 Then AdoCuenta.Recordset.Close
   If AdoDocumento.Recordset.State = 1 Then AdoDocumento.Recordset.Close
   If AdoFuente.Recordset.State = 1 Then AdoFuente.Recordset.Close
   If AdoOrganismo.Recordset.State = 1 Then AdoOrganismo.Recordset.Close
   If AdoPartida.Recordset.State = 1 Then AdoPartida.Recordset.Close
   If AdoProyecto.Recordset.State = 1 Then AdoProyecto.Recordset.Close
   If AdoRuc.Recordset.State = 1 Then AdoRuc.Recordset.Close
   If AdoUnidad.Recordset.State = 1 Then AdoUnidad.Recordset.Close
   'If rsRegularizacion.State = 1 Then rsRegularizacion.Close
   'If rsDetalle.State = 1 Then rsDetalle.Close
   Unload Me
   
End Sub

Private Sub CmdSalirBeneficiario_Click()
   FraBeneficiario.Visible = False
End Sub

Private Sub CmdSalirDetalle_Click()
  If msgSalir = "1" Then
    sino = MsgBox("Est� seguro de Salir . . .", vbYesNo + vbQuestion, "Atenci�n")
    If sino = vbYes Then
        FraOpcionesDetalle.Visible = False
        FraGrabarCancelar.Visible = False
        AdoRegularizacion.Visible = True
        DtcRegularizacion.Visible = True
        DtcRegularizacion.Enabled = True
        AdoDetalle.Enabled = False
        FraOpciones.Visible = True
        FraMaestro.Visible = True
        FraDetalleG.Visible = False
        Frame5.Visible = True
    Else
    '     MsgBox "No existe registro para eliminar", vbCritical + vbInformation, "Validaci�n de Datos"
    End If
  Else
    FraOpcionesDetalle.Visible = False
    FraGrabarCancelar.Visible = False
    AdoRegularizacion.Visible = True
    DtcRegularizacion.Visible = True
    DtcRegularizacion.Enabled = True
    AdoDetalle.Enabled = False
    FraOpciones.Visible = True
    FraMaestro.Visible = True
    FraDetalleG.Visible = False
    Frame5.Visible = True
    msgSalir = "0"
  End If
End Sub

Private Sub CmdSalirDev_Click()
    FraOpciones.Visible = True
    FraOpcionesDetalle.Visible = False
    FraGrabarCancelar.Visible = False
'rev Celia
  '  FraDevolucion.Visible = False
    LblCodigo.Caption = ""
    
    'Restaurando el grid
     Set rsRegularizacion = New ADODB.Recordset
    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
    rsRegularizacion.Open "select * from pagos where tipo_comp = 'DAC' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
    If rsRegularizacion.RecordCount > 0 Then
    Set DtcRegularizacion.DataSource = AdoRegularizacion
    Set AdoRegularizacion.Recordset = rsRegularizacion
    End If
'rev Celia
'DtGDevoluciones.Visible = False
    LblCodigo.Caption = "Nro Comprobante:"
    LblCabecera.Caption = "REGISTRO DE COMPROBANTES"
'rev Celia
   ' FraDev.Visible = False
End Sub

Private Sub CmdSalirGrid_Click()
   FraProyecto.Visible = False
End Sub

Private Sub ContableDevolucion_Click()
'ESTO COLOCAR CUANDO SE GRABA
'Devolucion_PAC_DAC
'evolucion_DAC
'Reversion_DAC

'gaby  Anulacion_DAC

'''
''''Devoluci�n contablemente
'''    'recogiendo los datos de devolucion Nro de comprobante al que pertenece la devoluci�n
'''    Set rsdev = New ADODB.Recordset
'''    If rsdev.State = 1 Then rsdev.Close
'''    rsdev.Open "select * from pagos where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "' and ges_gestion='" & AdoRegularizacion.Recordset("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
'''    If rsdev.RecordCount > 0 Then
'''            Set rsCoCoM = New ADODB.Recordset
'''            If rsCoCoM.State = 1 Then rsCoCoM.Close
'''            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC'", db, adOpenKeyset, adLockOptimistic
'''            If rsCoCoM.RecordCount > 0 Then
'''                Set rsDiario = New ADODB.Recordset
'''                If rsDiario.State = 1 Then rsDiario.Close
'''                rsDiario.Open "select * from co_Diario where Cod_Comp=" & rsCoCoM("Cod_Comp") & "", db, adOpenKeyset, adLockOptimistic
'''                If rsDiario.RecordCount > 0 Then
'''                    'Recuperando datos
'''                    Set rsCorr = New ADODB.Recordset
'''                    If rsCorr.State = 1 Then rsCorr.Close
'''                    rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'''                    If rsCorr.RecordCount > 0 Then
'''                        AuxCod_Comp = rsCorr("numero_correlativo") + 1
'''                        rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'''                        rsCorr.Update
'''                    End If
'''                    AuxTipo_Comp = rsDiario("Tipo_Comp")
'''                    AuxCod_Comp_C = rsDiario("Cod_Comp_C")
'''                    AuxD_Cuenta = rsDiario("D_Cuenta")
'''                    AuxD_Nombre = rsDiario("D_Nombre")
'''                    AuxD_SubCta1 = rsDiario("D_SubCta1")
'''                    AuxD_SubCta2 = rsDiario("D_SubCta2")
'''                    AuxD_Aux1 = rsDiario("D_Aux1")
'''                    AuxD_Aux2 = rsDiario("D_Aux2")
'''                    AuxD_Aux3 = rsDiario("D_Aux3")
'''                    AuxD_Cta_Larga = rsDiario("D_Cta_Larga")
'''                    AuxD_Des_Larga = rsDiario("D_Des_Larga")
'''                    AuxD_MontoBs = rsDiario("D_MontoBs")
''''                    AuxD_MontoDL = rsDiario("D_MontoDL")
'''                    AuxD_Cambio = rsDiario("D_Cambio")
'''
'''                    AuxH_Cuenta = rsDiario("H_Cuenta")
'''                    AuxH_Nombre = rsDiario("H_Nombre")
'''                    AuxH_SubCta1 = rsDiario("H_SubCta1")
'''                    AuxH_SubCta2 = rsDiario("H_SubCta2")
'''                    AuxH_Aux1 = rsDiario("H_Aux1")
'''                    AuxH_Aux2 = rsDiario("H_Aux2")
'''                    AuxH_Aux3 = rsDiario("H_Aux3")
'''                    AuxH_Cta_Larga = rsDiario("H_Cta_Larga")
'''                    AuxH_Des_Larga = rsDiario("H_Des_Larga")
'''                    AuxH_MontoBs = rsDiario("H_MontoBs")
''''                    AuxH_MontoDL = rsDiario("H_MontoDL")
'''                    AuxH_Cambio = rsDiario("H_Cambio")
'''
'''                    AuxUsr_Usuario = rsDiario("Usr_Usuario")
'''                    AuxFecha_Registro = rsDiario("Fecha_Registro")
'''                    AuxHora_Registro = rsDiario("Hora_Registro")
'''
'''                    'Adicionando una copia del registro
'''                    rsDiario.AddNew
'''                    rsDiario("Cod_Comp") = AuxCod_Comp
'''                    rsDiario("Tipo_Comp") = AuxTipo_Comp
'''                    rsDiario("Cod_Comp_C") = AuxCod_Comp_C
'''
'''                    rsDiario("D_Cuenta") = AuxH_Cuenta
'''                    rsDiario("D_Nombre") = AuxH_Nombre
'''                    rsDiario("D_SubCta1") = AuxH_SubCta1
'''                    rsDiario("D_SubCta2") = AuxH_SubCta2
'''                    rsDiario("D_Aux1") = AuxH_Aux1
'''                    rsDiario("D_Aux2") = AuxH_Aux2
'''                    rsDiario("D_Aux3") = AuxH_Aux3
'''                    rsDiario("D_Cta_Larga") = AuxH_Cta_Larga
'''                    rsDiario("D_Cta_Larga") = AuxH_Des_Larga
'''                    rsDiario("D_MontoBs") = AuxH_MontoBs
'''                    'rsDiario("D_MontoDL") = AuxH_MontoDL
'''                    rsDiario("D_Cambio") = AuxH_Cambio
'''
'''                    rsDiario("H_Cuenta") = AuxD_Cuenta
'''                    rsDiario("H_Nombre") = AuxD_Nombre
'''                    rsDiario("H_SubCta1") = AuxD_SubCta1
'''                    rsDiario("H_SubCta2") = AuxD_SubCta2
'''                    rsDiario("H_Aux1") = AuxD_Aux1
'''                    rsDiario("H_Aux2") = AuxD_Aux2
'''                    rsDiario("H_Aux3") = AuxD_Aux3
'''                    rsDiario("H_Cta_Larga") = AuxD_Cta_Larga
'''                    rsDiario("H_Cta_Larga") = AuxD_Des_Larga
'''                    rsDiario("H_MontoBs") = AuxD_MontoBs
'''                    'rsDiario("H_MontoDL") = AuxD_MontoDL
'''                    rsDiario("H_Cambio") = AuxD_Cambio
'''
'''                    rsDiario("Usr_Usuario") = AuxUsr_Usuario
'''                    rsDiario("Fecha_Registro") = AuxFecha_Registro
'''                    rsDiario("Hora_Registro") = AuxHora_Registro
'''                    rsDiario.Update
'''
'''                End If
'''          Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'''    End If
'''       Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'''End If

  End Sub


Private Sub DtcC_Click(Area As Integer)
    DtcCD.BoundText = DtcC.BoundText
    DtcConv2.BoundText = DtcC.BoundText
End Sub

Private Sub DtcCat_Click(Area As Integer)
   DtcCatDes.BoundText = DtcCat.BoundText
   DtcConv.BoundText = DtcCat.BoundText
End Sub

Private Sub DtcCatDes_Click(Area As Integer)
   DtcCat.BoundText = DtcCatDes.BoundText
   DtcConv.BoundText = DtcCatDes.BoundText
End Sub

Private Sub DtcCD_Click(Area As Integer)
   DtcC.BoundText = DtcCD.BoundText
   DtcConv2.BoundText = DtcCD.BoundText
End Sub

Private Sub DtcCodCta_Click(Area As Integer)
    DtcDesCta.BoundText = DtcCodCta.BoundText
End Sub

Private Sub DtcConv_Click(Area As Integer)
  DtcCat.BoundText = DtcConv.BoundText
  DtcCatDes.BoundText = DtcConv.BoundText
End Sub

Private Sub DtcConv2_Click(Area As Integer)
   DtcC.BoundText = DtcConv2.BoundText
   DtcCD.BoundText = DtcConv2.BoundText
End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
   DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
   DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
End Sub
Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
    DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
End Sub

Private Sub DtcDcu_Click(Area As Integer)
   DtcDcuDes.BoundText = DtcDcu.BoundText
End Sub

Private Sub DtcDcuDes_Click(Area As Integer)
   DtcDcu.BoundText = DtcDcuDes.BoundText
End Sub

Private Sub DtcDesCta_Click(Area As Integer)
    DtcCodCta.BoundText = DtcDesCta.BoundText
End Sub

Private Sub DtcDesOrg_Click(Area As Integer)
    DtcOrg.BoundText = DtcDesOrg.BoundText
    Call pCat(DtcDesOrg.BoundText)
End Sub

'Private Sub DtCDesTipoMoneda_Click(Area As Integer)
'    DtCTipoMoneda.BoundText = DtCDesTipoMoneda.BoundText
'End Sub

Private Sub DtCDesUnidad_Click(Area As Integer)
   DtCUnidad.BoundText = DtCDesUnidad.BoundText
End Sub

Private Sub DtCDR_Click(Area As Integer)
    DtCDRD.BoundText = DtCDR.BoundText
End Sub

Private Sub DtCDRD_Click(Area As Integer)
    DtCDR.BoundText = DtCDRD.BoundText
End Sub

Private Sub DtCFF_Click(Area As Integer)
    DtcFFD.BoundText = DtCFF.BoundText
End Sub

Private Sub DtcFFD_Click(Area As Integer)
    DtCFF.BoundText = DtcFFD.BoundText
End Sub

Private Sub DtCfte_Click(Area As Integer)
   DtcFteDes.BoundText = DTcFte.BoundText
   Call pOrganismo(DtcFteDes.BoundText)
End Sub

Private Sub DtcFteDes_Click(Area As Integer)
    DTcFte.BoundText = DtcFteDes.BoundText
    Call pOrganismo(DtcFteDes.BoundText)
End Sub

Private Sub dtcNombreRuc_Click(Area As Integer)
   dtcRuc.BoundText = dtcNombreRuc.BoundText
End Sub

Private Sub DtCOF_Click(Area As Integer)
    DtcOFD.BoundText = DtCOF.BoundText
    
End Sub

Private Sub DtcOFD_Click(Area As Integer)
    DtCOF.BoundText = DtcOFD.BoundText
End Sub
Private Sub DtcOrg_Click(Area As Integer)
      DtcDesOrg.BoundText = DtcOrg.BoundText
      Call pCat(DtcOrg.BoundText)
End Sub
Private Sub DtCPartida_Click(Area As Integer)
   DtCPartidaDes.BoundText = DtCPartida.BoundText
End Sub
Private Sub DtCPartidaDes_Click(Area As Integer)
   DtCPartida.BoundText = DtCPartidaDes.BoundText
End Sub
Private Sub DtcRegularizacion_Click()
    TIPOFORMULARIO = DtcTipoDes.Text
End Sub
Private Sub DtcRegularizacion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then Me.PopupMenu mnuAcciones
End Sub
Private Sub DtcRegularizacion_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'If Button = vbRightButton Then Me.PopupMenu mnuAcciones
End Sub
Private Sub dtcRuc_Click(Area As Integer)
   dtcNombreRuc.BoundText = dtcRuc.BoundText
'   Set rsBeneficiario = New ADODB.Recordset
'   If FraBeneficiario.Visible = False Then
'      rsBeneficiario.Open "select * from fc_beneficiario", db, adOpenKeyset, adLockOptimistic
'      rsBeneficiario.MoveFirst
'   End If
End Sub
'Private Sub DtcTipo_Click(Area As Integer)
'   DtcTipoDes.BoundText = DtcTipo.BoundText
'End Sub
Private Sub DtcTipoCod_Click(Area As Integer)
    DtcTipoDes.BoundText = DtcTipoCod.BoundText
End Sub
Private Sub dtctipoDes_Click(Area As Integer)
'  Dim sw As Integer
'   sw = 0
'   DtcTipoCod.BoundText = DtcTipoDes.BoundText
'   If DtcTipoDes.Text = "DEVOLUCION" Then
'        TxtTR.Text = "DEVOLUCION"
'        CmdCopiar_Click
'    End If
'    If DtcTipoDes.Text = "REVERSION TOTAL" Then
'        TxtTR.Text = "REVERSION TOTAL"
'        CmdCopiar_Click
'    End If
'    If DtcTipoDes.Text = "ANULACION" Then
'        TxtTR.Text = "ANULACION"
'        CmdCopiar_Click
'    End If
'        Set rsPg = New ADODB.Recordset
'        If rsPg.State = 1 Then rsPg.Close
'        rsPg.Open "select * from pagos where Nro_Comprobante_Anterior='" & TxtComprobante & "' and (estado_devolucion='S' or estado_anulado='S' or estado_reversion_total='S' or estado_reversion_parcial='S' )order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'        'rsPg.Open "select * from pagos where Nro_Comprobante_Anterior='" & TxtComprobante & "' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'        If rsPg.RecordCount > 0 Then
'                MsgBox "Ya existe comprobante de anulaci�n o de devoluci�n o reversi�n", vbInformation
'    '            MsgBox rsPg!estado_devolucion
'    '            MsgBox rsPg!estado_anulado
'    '            MsgBox rsPg!estado_reversion_total
'    '            MsgBox rsPg!estado_reversion_parcial
'          FraMaestro.Enabled = False
'          Exit Sub
'        End If
'----OJO----
' Arreglar con CELIA

'    CmdCopiar_Click
    
        'CmdAdicionar_Click
End Sub

'Private Sub DtCTipoMoneda_Click(Area As Integer)
'    DtCDesTipoMoneda.BoundText = DtCTipoMoneda.BoundText
'End Sub

'Private Sub DtcTipoDes_LostFocus()
''    If DtcTipoCod.Text = "DEV" Then
''       TxtComprobanteAnterior.Enabled = True
''       TxtComprobanteAnterior.SetFocus
''    End If
'End Sub

Private Sub DtCUnidad_Click(Area As Integer)
   DtCDesUnidad.BoundText = DtCUnidad.BoundText
End Sub

Private Sub DtCUT_Click(Area As Integer)
    DtCUTD.BoundText = DtCUT.BoundText
End Sub

Private Sub DtCUTD_Click(Area As Integer)
    DtCUT.BoundText = DtCUTD.BoundText
End Sub

Private Sub DtGProyecto_DblClick()
   TxtProgramad.Text = DtGProyecto.Columns(0)
   TxtSubprogramad.Text = DtGProyecto.Columns(1)
   TxtProyectod.Text = DtGProyecto.Columns(2)
   TxtActividadd.Text = DtGProyecto.Columns(3)
   txtProy.Text = DtGProyecto.Columns(4)
   FraProyecto.Visible = False
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    EntrarAdo = True  'Para problema de aprobar
   'Ojo por utilizaci�n del ado da el error de irowset.
    BotonesHabilitar Me, GlTipoAcceso
    Label7.Caption = GlUsuario
    Set rsRegularizacion = New ADODB.Recordset
    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
        QueryInicial = "select * from pagos where (estado_pagado='S') or (estado_contabilidad='P')"
        rsRegularizacion.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
        rsRegularizacion.Sort = "codigo_pago"
        CmdAprueba.Enabled = True
    Set AdoRegularizacion.Recordset = rsRegularizacion
    Set DtcRegularizacion.DataSource = AdoRegularizacion.Recordset
    
    If rsRegularizacion.RecordCount > 0 Then
        AdoRegularizacion.Recordset.MoveNext
        AdoRegularizacion.Recordset.MovePrevious
    End If
    
    'Obteniendo datos de clasificadores
    Set rsDocumentoRespaldo = New ADODB.Recordset
    rsDocumentoRespaldo.Open "select * from ac_documento_respaldo", db, adOpenKeyset, adLockOptimistic
    Set AdoDocumento.Recordset = rsDocumentoRespaldo
    DtcDcuDes.BoundText = DtcDcu.BoundText
     
    Set rsUnidad = New ADODB.Recordset
    rsUnidad.Open "select * from fc_unidad_ejecutora", db, adOpenKeyset, adLockOptimistic
    Set AdoUnidad.Recordset = rsUnidad
    DtCDesUnidad.BoundText = DtCUnidad.BoundText
    
    Set rsFuente = New ADODB.Recordset
    rsFuente.Open "select * from fc_fuente_financiamiento", db, adOpenKeyset, adLockOptimistic
    Set AdoFuente.Recordset = rsFuente
    DtcFteDes.BoundText = DTcFte.BoundText
     
    Set rsOrganismo = New ADODB.Recordset
    rsOrganismo.Open "select * from fc_organismo_financiamiento", db, adOpenKeyset, adLockOptimistic
    Set AdoOrganismo.Recordset = rsOrganismo
    DtcDesOrg.BoundText = DtcOrg.BoundText
    
    Set rsCategoria = New ADODB.Recordset
    rsCategoria.Open "select * from fc_categoria_financiador", db, adOpenKeyset, adLockOptimistic
    Set AdoCategoria.Recordset = rsCategoria
    DtcCatDes.BoundText = DtcCat.BoundText
    
    Set rsPartida = New ADODB.Recordset
    rsPartida.Open "select * from fc_partida_gasto", db, adOpenKeyset, adLockOptimistic
    Set AdoPartida.Recordset = rsPartida
    DtCPartidaDes.BoundText = DtCPartida.BoundText
 
    Set rsProyecto = New ADODB.Recordset
    rsProyecto.Open "select * from fc_estructura_programatica", db, adOpenKeyset, adLockOptimistic
    Set AdoProyecto.Recordset = rsProyecto
    Set DtGProyecto.DataSource = AdoProyecto
    
    Set rsBeneficiario = New ADODB.Recordset
    rsBeneficiario.Open "select * from fc_beneficiario", db, adOpenKeyset, adLockOptimistic
    Set AdoRuc.Recordset = rsBeneficiario
    dtcNombreRuc.BoundText = dtcRuc.BoundText
    
    Set rsCuenta = New ADODB.Recordset
    rsCuenta.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    Set AdoCuenta.Recordset = rsCuenta
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
    
    Set rsTipoComprobante = New ADODB.Recordset
    rsTipoComprobante.Open "select * from Tipo_Comprobante where ingresos='P'", db, adOpenKeyset, adLockOptimistic
    Set AdoTipo.Recordset = rsTipoComprobante
    DtcTipoDes.BoundText = DtcTipoCod.BoundText
    
    DtcTipoDes.Visible = False
    TxtTipoReg.Visible = True
    
    Screen.MousePointer = vbDefault
	Call SeguridadSet(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set rsRegularizacion = New ADODB.Recordset
    rsRegularizacion.CursorLocation = adUseClient

    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
    
    If rsDocumentoRespaldo.State = 1 Then rsDocumentoRespaldo.Close
    If rsUnidad.State = 1 Then rsUnidad.Close
    If rsFuente.State = 1 Then rsFuente.Close
    If rsOrganismo.State = 1 Then rsOrganismo.Close
    If rsCategoria.State = 1 Then rsCategoria.Close
    If rsPartida.State = 1 Then rsPartida.Close
    If rsProyecto.State = 1 Then rsProyecto.Close
    If rsBeneficiario.State = 1 Then rsBeneficiario.Close
    If rsCuenta.State = 1 Then rsCuenta.Close
    
  Set ClBuscaGrid = Nothing
'  Set ClVBusca = Nothing  'DUL:Libero el componente de Busqueda
End Sub
Private Sub mnuAccion_Click(Index As Integer)
Dim GlSqlAux As String
Dim RsAux As ADODB.Recordset
  Select Case Index
    Case 0 ' Devengado
      'Valida si puede
      Set RsAux = New ADODB.Recordset
      GlSqlAux = "SELECT sum(monto_Bolivianos) as SumaMonto FROM Pagos " & _
                 "WHERE (Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!codigo_pago & ")and" & _
                        "(org_codigo= '" & AdoRegularizacion.Recordset!org_codigo & "')and" & _
                        "(estado_devengado='S')"
      RsAux.Open GlSqlAux, db, adOpenStatic

        If RsAux!SumaMonto >= AdoRegularizacion.Recordset!monto_bolivianos Then
          MsgBox "No puede devengar ya que ......" & vbCrLf & _
               "Suma de Devengados = Bs " & RsAux!SumaMonto & vbCrLf & _
               "Monto Comprometido = Bs " & AdoRegularizacion.Recordset!monto_bolivianos, vbExclamation + vbOKOnly, "Atenci�n"
          RsAux.Close
          Exit Sub
        End If
        RsAux.Close
        'Realiza
        MsgBox "Realizando el Devengado.", vbInformation + vbOKOnly, "Atenci�n"
        TxtTR.Text = "DEV"
        swDevolucion = "E"
        CmdCopiar_Click
    Case 1 ' Reversi�n
      'Valida si puede
      ' If AdoRegularizacion.Recordset!Nro_Comprobante_Anterior <> AdoRegularizacion.Recordset!codigo_pago Then
      
      If AdoRegularizacion.Recordset!estado_compromiso = "S" And IsNull(AdoRegularizacion.Recordset!estado_devengado) Then
        Set RsAux = New ADODB.Recordset
        GlSqlAux = "SELECT sum(monto_bolivianos) as SumaMonto FROM Pagos " & _
                 "WHERE (Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!codigo_pago & ")and" & _
                        "(org_codigo= '" & AdoRegularizacion.Recordset!org_codigo & "')and" & _
                        "(estado_devengado='S')"
        RsAux.Open GlSqlAux, db, adOpenStatic
        If RsAux!SumaMonto > 0 Then
          MsgBox "No puede revertir ya que existe un compromiso realizado ..." & vbCrLf & _
               "Suma de Devengados = Bs " & RsAux!SumaMonto, vbExclamation + vbOKOnly, "Atenci�n"
          RsAux.Close
          Exit Sub
        Else
          MsgBox "Error En registro anterior, Verifique los datos ... "
          RsAux.Close
          Exit Sub
        End If
        RsAux.Close
      Else
      'Realiza
        MsgBox "Realizando la Reversi�n.", vbInformation + vbOKOnly, "Atenci�n"
        swDevolucion = "R"
        TxtTR.Text = "RVT"
        CmdCopiar_Click
      End If
    Case 2 ' Devoluci�n
      If AdoRegularizacion.Recordset!org_codigo = "999" Then
          Dim opt As Integer
          opt = MsgBox("Est� seguro de la Devoluci�n del comprobante  " & AdoRegularizacion.Recordset!codigo_pago & "  " & AdoRegularizacion.Recordset!org_codigo & "  ??", vbQuestion + vbYesNo, "ANULACION")
          If opt = vbYes Then
            Anulacion999 AdoRegularizacion.Recordset!codigo_pago, AdoRegularizacion.Recordset!org_codigo
            Reversion999 AdoRegularizacion.Recordset!codigo_pago, AdoRegularizacion.Recordset!org_codigo
            AdoRegularizacion.Recordset.Requery
            Exit Sub
          Else
            Exit Sub
          End If
      
      Else
        'Valida si puede
        Set RsAux = New ADODB.Recordset
        GlSqlAux = "SELECT codigo_pago, org_codigo, estado_devolucion FROM Pagos " & _
                   "WHERE (Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!codigo_pago & ")and" & _
                          "(org_codigo= '" & AdoRegularizacion.Recordset!org_codigo & "')and" & _
                          "(estado_devolucion='S')"
        RsAux.Open GlSqlAux, db, adOpenStatic
        If RsAux.RecordCount > 0 Then
          MsgBox "No puede devolver ya que ya se encuentra devuelto:" & vbCrLf & _
                 "Cmbte: " & RsAux!codigo_pago & ";  Org: " & RsAux!org_codigo, vbExclamation + vbOKOnly, "Atenci�n"
          RsAux.Close
          Exit Sub
        End If
        RsAux.Close
        'Realiza
        MsgBox "Realizando la Devoluci�n.", vbInformation + vbOKOnly, "Atenci�n"
        swDevolucion = "D"
        TxtTR.Text = "DVL"
        CmdCopiar_Click
      End If
    Case 3 ' Anulaci�n
      '****adiciono gaby
      If AdoRegularizacion.Recordset!org_codigo = "999" Then
          Dim opt1 As Integer
          opt1 = MsgBox("Est� seguro de anular este comprobante??", vbQuestion + vbYesNo, "ANULACION")
          If opt1 = vbYes Then
            Anulacion999 AdoRegularizacion.Recordset!codigo_pago, AdoRegularizacion.Recordset!org_codigo
            AdoRegularizacion.Recordset.Requery
            Exit Sub
          Else
            Exit Sub
          End If
      Else 'adiciono gaby
                
        'Valida si puede
        Set RsAux = New ADODB.Recordset
        GlSqlAux = "SELECT codigo_pago, org_codigo, estado_anulado FROM Pagos " & _
                 "WHERE (Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!codigo_pago & ")and" & _
                        "(org_codigo= '" & AdoRegularizacion.Recordset!org_codigo & "')and" & _
                        "(estado_anulado='S')"
        RsAux.Open GlSqlAux, db, adOpenStatic
        If RsAux.RecordCount > 0 Then
          MsgBox "No puede anular ya que ya se encuentra anulado:" & vbCrLf & _
               "Cmbte: " & RsAux!codigo_pago & ";  Org: " & RsAux!org_codigo, vbExclamation + vbOKOnly, "Atenci�n"
          RsAux.Close
          Exit Sub
        End If
        RsAux.Close
      'Realiza
        MsgBox "Realizando la Anulaci�n.", vbInformation + vbOKOnly, "Atenci�n"
        swDevolucion = "A"
        TxtTR.Text = "ANL"
        CmdCopiar_Click
      End If   ' adiciono gaby
  End Select
  'Celia Ctrl.Reversion, Devoluci�n, Anulaci�n
'  Dim sw As Integer
'  sw = 0
'  DtcTipoCod.BoundText = DtcTipoDes.BoundText
'  If DtcTipoDes.Text = "DEVOLUCION" Then
'        TxtTR.Text = "DEVOLUCION"
'        CmdCopiar_Click
'  End If
'  If DtcTipoDes.Text = "REVERSION TOTAL" Then
'        TxtTR.Text = "REVERSION TOTAL"
'        CmdCopiar_Click
'  End If
'  If DtcTipoDes.Text = "ANULACION" Then
'        TxtTR.Text = "ANULACION"
'        CmdCopiar_Click
'  End If
'
'        Set rsPg = New ADODB.Recordset
'        If rsPg.State = 1 Then rsPg.Close
'        rsPg.Open "select * from pagos where Nro_Comprobante_Anterior='" & TxtComprobante & "' and (estado_devolucion='S' or estado_anulado='S' or estado_reversion_total='S' or estado_reversion_parcial='S' )order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'        'rsPg.Open "select * from pagos where Nro_Comprobante_Anterior='" & TxtComprobante & "' order by codigo_pago ", db, adOpenKeyset, adLockOptimistic
'        If rsPg.RecordCount > 0 Then
'                MsgBox "Ya existe comprobante de anulaci�n o de devoluci�n o reversi�n", vbInformation
'    '            MsgBox rsPg!estado_devolucion
'    '            MsgBox rsPg!estado_anulado
'    '            MsgBox rsPg!estado_reversion_total
'    '            MsgBox rsPg!estado_reversion_parcial
'          FraMaestro.Enabled = False
'          Exit Sub
'        End If
        
End Sub

Private Sub OptChequeOrigen_Click()
'   LblNumeroOrigen.Caption = "No. Cheque: "
'   TxtNoTransferenciaOrigen.Enabled = True
'   DtCCuentaDestino.Visible = False
'   Label40.Visible = False
End Sub

Private Sub OptTransferenciaDestino_Click()
'   LblNumeroDestino.Caption = "transferencia: "
'   TxtNoTransferenciaDestino.Enabled = True
End Sub

Private Sub OptFilGral1_Click()
    Screen.MousePointer = vbHourglass
    QueryInicial = "select * from pagos where (estado_pagado='S') "
    If rsRegularizacion.State = 1 Then rsRegularizacion.CancelUpdate
    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
    rsRegularizacion.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
    rsRegularizacion.Sort = "codigo_pago"
    CmdAprueba.Enabled = True
    rsRegularizacion.Requery
    
    Set AdoRegularizacion.Recordset = rsRegularizacion
    Set DtcRegularizacion.DataSource = AdoRegularizacion.Recordset
    Screen.MousePointer = vbDefault
End Sub

Private Sub OptFilGral2_Click()
    Screen.MousePointer = vbHourglass
    QueryInicial = "select * from pagos where (estado_pagado='S' or estado_pagado='L' or estado_pagado='D')"
    If rsRegularizacion.State = 1 Then rsRegularizacion.CancelUpdate
    If rsRegularizacion.State = 1 Then rsRegularizacion.Close
    rsRegularizacion.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
    rsRegularizacion.Sort = "codigo_pago"
    CmdAprueba.Enabled = True
    rsRegularizacion.Requery
    
    Set AdoRegularizacion.Recordset = rsRegularizacion
    Set DtcRegularizacion.DataSource = AdoRegularizacion.Recordset
    Screen.MousePointer = vbDefault
End Sub

Private Sub OptTransferenciaOrigen_Click()
'   LblNumeroOrigen.Caption = "No. Transferencia: "
'   TxtNoTransferenciaOrigen.Enabled = True
'   DtCCuentaDestino.Visible = True
'   Label40.Visible = True
End Sub
Public Sub Graba_Copia()
'On Error GoTo error_grabar:
db.BeginTrans
        Set rscorrelativo = New ADODB.Recordset
        ComprobanteAnterior = AdoRegularizacion.Recordset("codigo_pago")     'gvi
        AdoRegularizacion.Recordset.AddNew
        If DtCOF.Text <> "" Then
            AdoRegularizacion.Recordset("org_codigo") = DtCOF.Text
        Else
            MsgBox "Introcudir Organismo Financiador", vbCritical + vbExclamation, "Validaci�n de datos"
            Exit Sub
        End If

        If DtCOF.Text = "111" Then  'TGN
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org111) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
                  rscorrelativo!Correl_Org111 = CDbl(CDbl(rscorrelativo!Correl_Org111) + 1)
                  rscorrelativo.Update
            End If
        End If

        If DtCOF.Text = "112" Then 'TGNP
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org112) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
                  rscorrelativo!Correl_Org112 = CDbl(CDbl(rscorrelativo!Correl_Org112) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "114" Then 'RECON
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org114) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
                  rscorrelativo!Correl_Org114 = CDbl(CDbl(rscorrelativo!Correl_Org114) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "344" Then 'UNICEF
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org344) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org344) + 1)
                  rscorrelativo!Correl_Org344 = CDbl(CDbl(rscorrelativo!Correl_Org344) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "381" Then  'FAD
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org381) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org381) + 1)
                  rscorrelativo!Correl_Org381 = Val(Val(rscorrelativo!Correl_Org381) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "411" Then  'BID
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org411) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
                  rscorrelativo!Correl_Org411 = CDbl(CDbl(rscorrelativo!Correl_Org411) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "415" Then  'IDA
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org415) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
                  rscorrelativo!Correl_Org415 = CDbl(CDbl(rscorrelativo!Correl_Org415) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "516" Then  'KFW
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org516) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
                  rscorrelativo!Correl_Org516 = CDbl(CDbl(rscorrelativo!Correl_Org516) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "541" Then  'ALEM
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org541) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org541) + 1)
                  rscorrelativo!Correl_Org541 = CDbl(CDbl(rscorrelativo!Correl_Org541) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "551" Then  'DIN
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org551) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org551) + 1)
                  rscorrelativo!Correl_Org551 = CDbl(CDbl(rscorrelativo!Correl_Org551) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "556" Then  'HOL
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org556) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
                  rscorrelativo!Correl_Org556 = CDbl(CDbl(rscorrelativo!Correl_Org556) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "565" Then  'SUE
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org565) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
                  rscorrelativo!Correl_Org565 = CDbl(CDbl(rscorrelativo!Correl_Org565) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "999" Then  'S/N
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org999) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org13) + 1)
                  rscorrelativo!Correl_Org999 = CDbl(CDbl(rscorrelativo!Correl_Org999) + 1)
                  rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "Org14" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org14) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org13) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org14) + 1)
                  rscorrelativo!Correl_Org14 = CDbl(CDbl(rscorrelativo!Correl_Org14) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org14 = 0
                rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "Org15" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org15) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
                  rscorrelativo!Correl_Org15 = CDbl(CDbl(rscorrelativo!Correl_Org15) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org15 = 0
                rscorrelativo.Update
            End If
        End If
         
        If DtCOF.Text = "Org16" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org16) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
                  rscorrelativo!Correl_Org16 = CDbl(CDbl(rscorrelativo!Correl_Org16) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org16 = 0
                rscorrelativo.Update
            End If
        End If
   
        If DtCOF.Text = "Org17" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org17) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
                  'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
                  rscorrelativo!Correl_Org17 = CDbl(CDbl(rscorrelativo!Correl_Org17) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org17 = 0
                rscorrelativo.Update
            End If
        End If
   
        If DtCOF.Text = "Org18" Then
            rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            If Not IsNull(rscorrelativo!Correl_Org18) Then
                  AdoRegularizacion.Recordset("codigo_pago") = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
                    'AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
                  rscorrelativo!Correl_Org18 = CDbl(CDbl(rscorrelativo!Correl_Org18) + 1)
                  rscorrelativo.Update
            Else
                rscorrelativo!Correl_Org18 = 0
                rscorrelativo.Update
            End If
        End If
 
   ' Greco para Copia y los de abajo
   'JORGE 21 JULIO 2000
   'If TxtTR.Text = "DVL" Or TxtTR.Text = "RVT" Or TxtTR.Text = "ANL" Or TxtTR.Text = "DEV" Or swGrabaCopia = 1 Then
   'If AuxCopia = "R" Then
        'AdoRegularizacion.Recordset("codigo_solicitud") = TxtNS.Text
        AdoRegularizacion.Recordset("codigo_documento") = DtCDR.Text
        AdoRegularizacion.Recordset("fecha_egreso") = Date
        AdoRegularizacion.Recordset("tipo_formulario") = TxtTR.Text
        AdoRegularizacion.Recordset("tipo_comp") = "DAC"
       
        If swDevolucion = "E" Or swDevolucion = "R" Or swDevolucion = "D" Or swDevolucion = "A" Then
              AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = ComprobanteAnterior
           Else
              AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = AdoRegularizacion.Recordset("codigo_pago")
        End If
        '        If TxtTR.Text = "DEV" Then
        '          AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = ComprobanteAnterior
        '        End If
        '        If swGrabaCopia = 1 Then
        '          AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = AdoRegularizacion.Recordset("codigo_pago")
        '        End If
        'Validaciones
        If DtCUT.Text = "" Then
            MsgBox "Falta Unidad T�cnica, El proceso ser� interrumpido ! ...", vbCritical + vbInformation, "Validaci�n de datos"
            Exit Sub
        End If
        If TxtCO.Text = "" Then
            MsgBox "Falta n�mero Orden de Pago, El proceso ser� interrumpido ! ...", vbCritical + vbExclamation
            Exit Sub
        End If
        If TxtNS.Text = "" Then
            MsgBox "Falta N�mero de Solicitud, El proceso ser� interrumpido ! ...", vbCritical + vbExclamation
            Exit Sub
        End If
        If DtCFF.Text = "" Then
            MsgBox "Falta Fte. de Financiamiento, El proceso ser� interrumpido ! ...", vbCritical + vbExclamation, "Validaci�n de datos"
            Exit Sub
        End If
        If TxtJ.Text = "" Then
            MsgBox "Falta Justificaci�n, El proceso ser� interrumpido ! ...", vbCritical + vbExclamation, "Validaci�n de datos"
            Exit Sub
        End If

        AdoRegularizacion.Recordset("uni_codigo") = "VEIPS" 'DtCUT.Text
        AdoRegularizacion.Recordset("codigo_unidad") = DtCUT.Text
        AdoRegularizacion.Recordset("codigo_orden") = TxtCO.Text
        AdoRegularizacion.Recordset("codigo_solicitud") = TxtNS.Text
        AdoRegularizacion.Recordset("fte_codigo") = DtCFF.Text
        AdoRegularizacion.Recordset("justificacion") = TxtJ.Text
        
        ' no quiere asignar ....
        AdoRegularizacion.Recordset("codigo_categoria") = DtcC.Text
        AdoRegularizacion.Recordset("codigo_convenio") = DtcConv2.Text
        AdoRegularizacion.Recordset("tipo_moneda") = "Bs." 'DtCTipoMoneda.Text
        AdoRegularizacion.Recordset("liquido_pagar") = "0" 'Val(TxtLiquido.Text)
        If Not IsNull(vgMontoTotal) Then AdoRegularizacion.Recordset("liquido_pagar") = CDbl(vgMontoTotal)
        
        'Estados de aprobaci�n
        If TxtTR.Text = "COM" Then
            AdoRegularizacion.Recordset("estado_compromiso") = "N"
        End If
        If TxtTR.Text = "DEV" Then
            AdoRegularizacion.Recordset("estado_devengado") = "N"
        End If
        If TxtTR.Text = "CYD" Then
            AdoRegularizacion.Recordset("estado_compromiso") = "N"
            AdoRegularizacion.Recordset("estado_devengado") = "N"
        End If
        If TxtTR.Text = "REG" Then
            AdoRegularizacion.Recordset("estado_compromiso") = "N"
            AdoRegularizacion.Recordset("estado_devengado") = "N"
            AdoRegularizacion.Recordset("estado_pagado") = "N"
        End If

        'Datos de seguimiento
        AdoRegularizacion.Recordset("ges_gestion") = Year(Now)
        AdoRegularizacion.Recordset("usr_usuario") = Label7.Caption
        AdoRegularizacion.Recordset("fecha_registro") = Date
        AdoRegularizacion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
         
        ' Para Borrar . . .
        '        If DtcTipoDes.Text = "DEVOLUCION" Or DtcTipoDes.Text = "REVERSION TOTAL" Or DtcTipoDes.Text = "ANULACION" Or DtcTipoDes.Text = "DEVENGADO" Then
        '            AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") = ComprobanteAnterior
        '        End If
        
        If TxtTR.Text = "DVL" Then
            AdoRegularizacion.Recordset("estado_devolucion") = "N"
        End If
        If TxtTR.Text = "RVT" Then
            AdoRegularizacion.Recordset("estado_reversion_total") = "N"
        End If
        If TxtTR.Text = "RVP" Then
            AdoRegularizacion.Recordset("estado_reversion_parcial") = "N"
        End If
        If TxtTR.Text = "ANL" Then
            AdoRegularizacion.Recordset("estado_anulado") = "N"
'            AdoRegularizacion.Recordset("estado_devengado") = "N"
        End If
        
        AdoRegularizacion.Recordset.Update

       FraCopiaRegistro.Visible = False
       FraCopiaRegistro.Enabled = False
      
       CmdAdicionar.Enabled = True
       CmdBorrar.Enabled = True
       CmdSalir.Enabled = True
       LblTitulo.Caption = ""
       FraOpciones.Visible = True
       FraGrabarCancelar.Visible = False
       DtcRegularizacion.Enabled = True
   
   'End If
 
    'Determinamos si se trata o no de devolucion, si es asi copiamos el detalle
    '     If TxtTR.Text = "DVL" Then
    Set rsDet = New ADODB.Recordset
    If rsDet.State = 1 Then rsDet.Close
    'gaby revisar no traer todo el recordset
    rsDet.Open "select * from pago_detalle", db, adOpenKeyset, adLockOptimistic
    rsDet.AddNew
    rsDet("par_codigo") = vgCodigoPartida
    rsDet("pro_programa") = vgPrograma
    rsDet("pro_subprograma") = vgSubPrograma
    rsDet("pro_proyecto") = vgProyecto
    rsDet("pro_actividad") = vgActividad
    rsDet("cta_codigo") = vgCtaOrigen
    
    rsDet("numero_cheque_trf") = vgNroChequeOTransferencia
    rsDet("cta_codigo_destino") = IIf(vgCtaDestino = "", "", vgCtaDestino)
    
    rsDet("codigo_beneficiario") = vgCodBeneficiario
    If Not IsNull(vgMontoTotal) Then rsDet("monto_total") = CDbl(vgMontoTotal) 'CCur(vgMontoTotal.Value)
    
    rsDet("deducciones") = vgDeducciones
    rsDet("monto_bolivianos") = CDbl(vgMB)
    
    TxtTipoCambio.Enabled = True
    Set rstipocambio = New ADODB.Recordset
    sql_TC = "select fecha_cambio, Cambio_Oficial  from ac_tipo_cambio  where fecha_cambio = (select max(fecha_cambio) as expr1 from ac_tipo_cambio)"
    rstipocambio.Open sql_TC, db, adOpenKeyset, adLockReadOnly
    GlTipoCambioOficial = rstipocambio!cambio_oficial

    rsDet("tipo_cambio") = GlTipoCambioOficial
    'rsDet("tipo_cambio") = CDbl(vgTipoCambio)
    rsDet("monto_dolares") = CDbl(vgMontoDolares)
    rsDet("org_codigo") = AdoRegularizacion.Recordset("org_codigo")
    vgOrgCodigo = AdoRegularizacion.Recordset("org_codigo")
    rsDet("ges_gestion") = vgGesGestion
 '   MsgBox "NUMERO DE COMPROBANTE NUEVO . . . " & AdoRegularizacion.Recordset("codigo_pago")
'    MsgBox "NUMERO DE ORGANISMO NUEVO . . . " & AdoRegularizacion.Recordset("org_codigo")
    rsDet("codigo_pago") = AdoRegularizacion.Recordset("codigo_pago")
    rsDet("codigo_pago_detalle") = vgCodigoPagoDetalle
    ' campos no existen en tabla
    '             rsDet("usuario_aprueba") = Label7.Caption
     
    '             rsDet("hora_aprueba") = Format(Time, "hh:mm:ss")
    rsDet("fecha_pago") = Date
    rsDet("usr_usuario") = Label7.Caption
    rsDet("fecha_registro") = Date
    rsDet("hora_registro") = Format(Time, "hh:mm:ss")
    
    rsDet.Update
        
    'If TxtTR.Text = "DVL" Then ' p/verificar
    If swDevolucion = "D" Then
        'Acumulando datos en el campo de cta_acum_dev de fc_cuenta_bancaria
        Set rsCtaB = New ADODB.Recordset
        If rsCtaB.State = 1 Then rsCtaB.Close
        rsCtaB.Open "select * FROM fc_cuenta_bancaria WHERE Cta_codigo='" & vgCtaOrigen & "'", db, adOpenKeyset, adLockOptimistic
        If rsCtaB.RecordCount > 0 Then
            rsCtaB("cta_acum_dev") = rsCtaB("cta_acum_dev") + CDbl(vgMontoTotal)
            rsCtaB.Update
        End If
        
        'Acumulando datos en el campo de fgs_acum_dev de fc_cuenta_bancaria
        Set rsFGasto = New ADODB.Recordset
        If rsFGasto.State = 1 Then rsFGasto.Close
            rsFGasto.Open "select * FROM fo_formulacion_gasto WHERE fte_codigo='" & vgFteCodigo & "' and org_codigo='" & vgOrgCodigo & "' and pro_programa='" & vgPrograma & "' and pro_Subprograma='" & vgSubPrograma & "' and pro_Proyecto='" & vgProyecto & "' and pro_Actividad='" & vgActividad & "' and par_codigo= '" & vgCodigoPartida & "' and ges_gestion='2000'", db, adOpenKeyset, adLockOptimistic
            If rsFGasto.RecordCount > 0 Then
                    rsFGasto("fgs_acum_dev") = rsFGasto("fgs_acum_dev") + CDbl(vgMontoTotal)
                    rsFGasto.Update
            End If
        'GROVER .....Devolucion_PAC_DAC
        'Devolucion_PAC_DAC
     End If
   
     'If DtcTipoDes.Text = "REVERSION TOTAL" Then
     If swDevolucion = "R" Then
        'Acumulando datos en el campo de fgs_acum_dev de fc_cuenta_bancaria
        Set rsFGasto = New ADODB.Recordset
        If rsFGasto.State = 1 Then rsFGasto.Close
            rsFGasto.Open "select * FROM fo_formulacion_gasto WHERE fte_codigo='" & vgFteCodigo & "' and org_codigo='" & vgOrgCodigo & "' and pro_programa='" & vgPrograma & "' and pro_Subprograma='" & vgSubPrograma & "' and pro_Proyecto='" & vgProyecto & "' and pro_Actividad='" & vgActividad & "' and par_codigo= '" & vgCodigoPartida & "' and ges_gestion='2000'", db, adOpenKeyset, adLockOptimistic
            If rsFGasto.RecordCount > 0 Then
                    rsFGasto("fgs_acum_rev") = rsFGasto("fgs_acum_rev") + vgMontoTotal
                    rsFGasto.Update
            End If
        'GROVER .....Reversion_DAC
        'Reversion_DAC
     End If
     
     'If DtcTipoDes.Text = "ANULACION" Then
     If swDevolucion = "A" Then
        'Acumulando datos en el campo de cta_acum_dev de fc_cuenta_bancaria
        'Jorge 07/08/2000
        Set rsCtaB = New ADODB.Recordset
        If rsCtaB.State = 1 Then rsCtaB.Close
        rsCtaB.Open "select * FROM fc_cuenta_bancaria WHERE Cta_codigo='" & vgCtaOrigen & "'", db, adOpenKeyset, adLockOptimistic
        If rsCtaB.RecordCount > 0 Then
            rsCtaB("cta_acum_anl") = rsCtaB("cta_acum_anl") + vgMontoTotal
            rsCtaB.Update
        End If
        'GROVER ..... Anulacion_DAC
        'Anulacion_DAC (AdoRegularizacion.Recordset)
     End If

    If rsDetalle.State = adStateOpen Then rsDetalle.Close
    rsDetalle.Open "select * from pago_detalle where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
    Set DtGDetalle.DataSource = rsDetalle
    If rsDetalle.RecordCount > 0 Then
        DtGDetalle.Refresh
    End If

swGrabaCopia = 0
db.CommitTrans
Exit Sub
error_grabar:

MsgBox Err.Number & " " & Err.Description
db.RollbackTrans
End Sub


Private Sub pCat(CodOrganismo As String)
   Dim strConsulta As String
   
   strConsulta = "select * from fc_categoria_financiador where org_codigo='" & CodOrganismo & "'"
   
   Set DtcCat.RowSource = Nothing
   Set DtcCat.RowSource = db.Execute(strConsulta, , adCmdText)
   DtcCat.ReFill
   DtcCat.BoundText = Empty
   
   Set DtcCatDes.RowSource = Nothing
   Set DtcCatDes.RowSource = db.Execute(strConsulta, , adCmdText)
   DtcCatDes.ReFill
   DtcCatDes.BoundText = Empty

End Sub

Private Sub TxtComprobanteAnterior_LostFocus()
'    ANTERIOR = TxtComprobanteAnterior.Text
End Sub

Private Sub TxtMontoDolares_Click()
    TxtMontoFuente.Text = 0
End Sub

Private Sub TxtMontoDolares_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtMontoFuente_Click()
    TxtMontoDolares.Text = 0
End Sub
Public Sub Devolucion()
Dim sino As Variant
    
    sino = MsgBox("Est� seguro de realizar la devoluci�n ?", vbYesNo + vbQuestion, "Atenci�n")
    If sino = vbYes Then
            DtcTipoDes.Text = "DEVOLUCION"
            'Abriendo la base para colocar numero de devolucion de devoluci�n
            Set rsCorrel_Dev = New ADODB.Recordset
            If rsCorrel_Dev.State = 1 Then rsCorrel_Dev.Close
            rsCorrel_Dev.Open "select * from fc_correl where tipo_tramite='Devolucion'", db, adOpenKeyset, adLockOptimistic
            If rsCorrel_Dev.RecordCount > 0 Then
                    TxtComprobanteAnterior.Text = AdoRegularizacion.Recordset("codigo_pago")
                    LblCodigo.Caption = "Nro. Devolucion"
                    TxtComprobante.Text = rsCorrel_Dev("numero_correlativo") + 1
                    rsCorrel_Dev("numero_correlativo") = rsCorrel_Dev("numero_correlativo") + 1
                    rsCorrel_Dev.Update
            Else
                    MsgBox "No existe correlativo"
            End If
            
            Set rsP = New ADODB.Recordset
            If rsP.State = 1 Then rsdev.Close
            rsP.Open "select * from pagos where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "' and ges_gestion='" & AdoRegularizacion.Recordset("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
            If rsP.RecordCount > 0 Then
                    rsP("estado_devolucion") = "S"
                    rsP("nro_devolucion") = Val(TxtComprobante.Text)
                    rsP.Update
            End If
            
            'Estado Devolucion en pagos es Si
'            AdoRegularizacion.Recordset("estado_devolucion") = "S"
'            AdoRegularizacion.Recordset("nro_devolucion") = Val(TxtComprobante.Text)
'            AdoRegularizacion.Recordset.Update

            'Abriendo la base para a�adir un registro en devolucion
            Set rsdev = New ADODB.Recordset
            If rsdev.State = 1 Then rsdev.Close
            rsdev.Open "select * from fc_devolucion", db, adOpenKeyset, adLockOptimistic
                rsdev.AddNew
                rsdev("Nro_Dev") = Val(TxtComprobante.Text)
                'If Not IsNull(txtobs_dev.Text) Then rsDev("Obs_Dev") = Val(TxtComprobante.Text)
                    rsdev("Nro_Dev") = Val(TxtComprobante.Text)
                    rsdev("usr_usuario") = Label7.Caption
                    rsdev("fecha_registro") = Date
                    rsdev("hora_registro") = Format(Time, "hh:mm:ss")
                    rsdev.Update
            'Mostrando el grid con datos de devoluci�n
    End If
    
    Grid_Devoluciones
    'FraDev.Visible = True
End Sub

Public Sub Grid_Devoluciones()
'Colocando el  nuevo grid con datos de pago y devolucion
'    Set rsPago_dev = New ADODB.Recordset
'    If rsPago_dev.State = 1 Then rsPago_dev.Close
'    rsPago_dev.Open "SELECT Fc_Devolucion.Nro_Dev,PAGOS.codigo_pago, PAGOS.codigo_orden, PAGOS.org_codigo, PAGOS.tipo_comp, PAGOS.estado_compromiso, PAGOS.estado_devengado, PAGOS.estado_pagado, PAGOS.estado_devolucion FROM PAGOS INNER JOIN Fc_Devolucion ON PAGOS.Nro_devolucion = Fc_Devolucion.Nro_Dev", db, adOpenKeyset, adLockOptimistic
'    If rsPago_dev.RecordCount > 0 Then
'        Set DtGDevoluciones.DataSource = rsPago_dev
'        Set AdoDevolucion.Recordset = rsPago_dev
'    End If
'    DtGDevoluciones.Visible = True
'    AdoDevolucion.Visible = True
'    CmdImprimirDev.Enabled = True
'    LblCabecera.Caption = "COMPROBANTES DE DEVOLUCIONES"
End Sub
'Public Sub Devolucion_PAC_DAC()
'    'Devoluci�n contablemente
'    'recogiendo los datos de devolucion Nro de comprobante al que pertenece la devoluci�n
'    Set rsdev = New ADODB.Recordset
'    If rsdev.State = 1 Then rsdev.Close
'    rsdev.Open "select * from pagos where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "' and ges_gestion='" & AdoRegularizacion.Recordset("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
'    If rsdev.RecordCount > 0 Then
'            Set rsCoCoM = New ADODB.Recordset
'            If rsCoCoM.State = 1 Then rsCoCoM.Close
'            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and (Tipo_Comp='DAC' or Tipo_Comp='CAD') ", db, adOpenKeyset, adLockOptimistic
'            If rsCoCoM.RecordCount > 0 Then
'                'Creaci�n de la cabecera o registros maestro en CO_COMPROBANTE_M
'                'Recuperando datos de co_comprobante_m
'                cocmCod_CompDiario = rsCoCoM("Cod_Comp")
'                cocmTipo_Comp = rsCoCoM("Tipo_Comp")
'                cocmCod_Trans = TxtComprobante.Text 'AdoRegularizacion.Recordset("codigo_pago") 'TxtComprobante.text TxtNC.Text 'rsCoCoM("Cod_Trans")
'                cocmCod_Trans_Detalle = rsCoCoM("Cod_Trans_Detalle")
'                cocmOrg_Codigo = rsCoCoM("Org_Codigo")
'                cocmGes_Gestion = rsCoCoM("Ges_Gestion")
'                cocmNum_Respaldo = rsCoCoM("Num_Respaldo")
'                cocmFecha_A = rsCoCoM("Fecha_A")
'                cocmCodigo_Beneficiario = rsCoCoM("Codigo_Beneficiario")
'                cocmCodigo_Documento = rsCoCoM("Codigo_Documento")
'                cocmGlosa = rsCoCoM("Glosa")
'                cocmStatus = rsCoCoM("Status")
'                cocmUsr_usuario = rsCoCoM("Usr_Usuario")
'                'Adicionando un nuevo registro
'                'Generando nuevo c�digo
'                        Set rsCorr = New ADODB.Recordset
'                        If rsCorr.State = 1 Then rsCorr.Close
'                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'                        If rsCorr.RecordCount > 0 Then
'                            cocmCod_Comp = rsCorr("numero_correlativo") + 1
'                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'                            rsCorr.Update
'                        End If
'                        MsgBox "NUMERO DE 1era. CUENTA DAC" & cocmCod_Comp
'                        rsCorr.Close
'                rsCoCoM.AddNew
'                    rsCoCoM("Cod_Comp") = cocmCod_Comp
'                    rsCoCoM("Tipo_Comp") = cocmTipo_Comp
'                    rsCoCoM("Cod_Trans") = TxtComprobante.Text 'AdoRegularizacion.Recordset("codigo_pago") 'TxtNC.Text 'cocmCod_Trans
'                    rsCoCoM("Cod_Trans_Detalle") = cocmCod_Trans_Detalle
'                    rsCoCoM("org_codigo") = cocmOrg_Codigo
'                    rsCoCoM("Ges_Gestion") = cocmGes_Gestion
'                    rsCoCoM("Num_Respaldo") = cocmNum_Respaldo
'                    rsCoCoM("Fecha_A") = cocmFecha_A
'                    rsCoCoM("Codigo_Beneficiario") = cocmCodigo_Beneficiario
'                    rsCoCoM("Codigo_Documento") = cocmCodigo_Documento
'                    rsCoCoM("Glosa") = cocmGlosa
'                    rsCoCoM("Status") = cocmStatus
'                    rsCoCoM("usr_usuario") = Label7.Caption
'                    rsCoCoM("fecha_registro") = Date
'                    rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
'                rsCoCoM.Update
'
'                Set rsdiario = New ADODB.Recordset
'                If rsdiario.State = 1 Then rsdiario.Close
'                rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
'                'rsDiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_Comp & "", db, adOpenKeyset, adLockOptimistic
'                If rsdiario.RecordCount > 0 Then
'                    AuxCod_Comp = cocmCod_Comp
'                    AuxTipo_Comp = rsdiario("Tipo_Comp")
'                    AuxCod_Comp_C = IIf(IsNull(rsdiario("Cod_Comp_C")), 0, rsdiario("Cod_Comp_C"))
'                    AuxD_Cuenta = rsdiario("D_Cuenta")
'                    AuxD_Nombre = rsdiario("D_Nombre")
'                    AuxD_SubCta1 = rsdiario("D_SubCta1")
'                    AuxD_SubCta2 = rsdiario("D_SubCta2")
'                    AuxD_Aux1 = rsdiario("D_Aux1")
'                    AuxD_Aux2 = rsdiario("D_Aux2")
'                    AuxD_Aux3 = rsdiario("D_Aux3")
'                    AuxD_Cta_Larga = IIf(IsNull(rsdiario("D_Cta_Larga")), "-", rsdiario("D_Cta_Larga"))
'                    AuxD_Des_Larga = IIf(IsNull(rsdiario("D_Des_Larga")), "-", rsdiario("D_Des_Larga"))
'                    AuxD_MontoBs = rsdiario("D_MontoBs")
'                    AuxD_MontoDL = rsdiario("D_MontoDL")
'                    AuxD_Cambio = rsdiario("D_Cambio")
'
'                    AuxH_Cuenta = rsdiario("H_Cuenta")
'                    AuxH_Nombre = rsdiario("H_Nombre")
'                    AuxH_SubCta1 = rsdiario("H_SubCta1")
'                    AuxH_SubCta2 = rsdiario("H_SubCta2")
'                    AuxH_Aux1 = rsdiario("H_Aux1")
'                    AuxH_Aux2 = rsdiario("H_Aux2")
'                    AuxH_Aux3 = rsdiario("H_Aux3")
'                    AuxH_Cta_Larga = rsdiario("H_Cta_Larga")
'                    AuxH_Des_Larga = rsdiario("H_Des_Larga")
'                    AuxH_MontoBs = rsdiario("H_MontoBs")
'                    AuxH_MontoDL = rsdiario("H_MontoDL")
'                    AuxH_Cambio = rsdiario("H_Cambio")
'
'                    AuxUsr_Usuario = rsdiario("Usr_Usuario")
'                    AuxFecha_Registro = Date
'                    AuxHora_Registro = Format(Time, "hh:mm:ss")
'
'                    'Adicionando una copia del registro
'                    rsdiario.AddNew
'                    rsdiario("Cod_Comp") = AuxCod_Comp 'AuxCod_Comp_C
'                    rsdiario("Tipo_Comp") = AuxTipo_Comp
'                    rsdiario("Cod_Comp_C") = AuxCod_Comp_C
'
'                    rsdiario("D_Cuenta") = AuxH_Cuenta
'                    rsdiario("D_Nombre") = AuxH_Nombre
'                    rsdiario("D_SubCta1") = AuxH_SubCta1
'                    rsdiario("D_SubCta2") = AuxH_SubCta2
'                    rsdiario("D_Aux1") = AuxH_Aux1
'                    rsdiario("D_Aux2") = AuxH_Aux2
'                    rsdiario("D_Aux3") = AuxH_Aux3
'                    rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
'                    rsdiario("D_Cta_Larga") = AuxH_Des_Larga
'                    rsdiario("D_MontoBs") = AuxH_MontoBs
'                    rsdiario("D_MontoDL") = AuxH_MontoDL
'                    rsdiario("D_Cambio") = AuxH_Cambio
'
'                    rsdiario("H_Cuenta") = AuxD_Cuenta
'                    rsdiario("H_Nombre") = AuxD_Nombre
'                    rsdiario("H_SubCta1") = AuxD_SubCta1
'                    rsdiario("H_SubCta2") = AuxD_SubCta2
'                    rsdiario("H_Aux1") = AuxD_Aux1
'                    rsdiario("H_Aux2") = AuxD_Aux2
'                    rsdiario("H_Aux3") = AuxD_Aux3
'                    rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
'                    rsdiario("H_Cta_Larga") = AuxD_Des_Larga
'                    rsdiario("H_MontoBs") = AuxD_MontoBs
'                    rsdiario("H_MontoDL") = AuxD_MontoDL
'                    rsdiario("H_Cambio") = AuxD_Cambio
'
'                    rsdiario("Usr_Usuario") = AuxUsr_Usuario
'                    rsdiario("Fecha_Registro") = AuxFecha_Registro
'                    rsdiario("Hora_Registro") = AuxHora_Registro
'                    rsdiario.Update
'
'                End If
'
'
'                'Comprobantes PAC
'                If rsCoCoM.State = 1 Then rsCoCoM.Close
'                rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='PAC' or Tipo_Comp='CAP'", db, adOpenKeyset, adLockOptimistic
'                If rsCoCoM.RecordCount > 0 Then
'
''                Set rsCoCoM = New ADODB.Recordset
''                If rsCoCoM.State = 1 Then rsCoCoM.Close
''                rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC'", db, adOpenKeyset, adLockOptimistic
'            If rsCoCoM.RecordCount > 0 Then
'                'Creaci�n de la cabecera o registros maestro en CO_COMPROBANTE_M
'                'Recuperando datos de co_comprobante_m
'                cocmCod_CompDiario = rsCoCoM("Cod_Comp")
'                cocmTipo_Comp = rsCoCoM("Tipo_Comp")
'                cocmCod_Trans = TxtComprobante.Text 'AdoRegularizacion.Recordset("codigo_pago") 'TxtNC.Text 'rsCoCoM("Cod_Trans")
'                cocmCod_Trans_Detalle = rsCoCoM("Cod_Trans_Detalle")
'                cocmOrg_Codigo = rsCoCoM("Org_Codigo")
'                cocmGes_Gestion = rsCoCoM("Ges_Gestion")
'                cocmNum_Respaldo = rsCoCoM("Num_Respaldo")
'                cocmFecha_A = rsCoCoM("Fecha_A")
'                cocmCodigo_Beneficiario = rsCoCoM("Codigo_Beneficiario")
'                cocmCodigo_Documento = rsCoCoM("Codigo_Documento")
'                cocmGlosa = rsCoCoM("Glosa")
'                cocmStatus = rsCoCoM("Status")
'                cocmUsr_usuario = IIf(IsNull(rsCoCoM("Usr_Usuario")), "", rsCoCoM("Usr_Usuario"))
'                'Adicionando un nuevo registro
'                'Generando nuevo c�digo
'                'Segunda genera*********
'                        Set rsCorr = New ADODB.Recordset
'                        If rsCorr.State = 1 Then rsCorr.Close
'                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'                        If rsCorr.RecordCount > 0 Then
'                            cocmCod_Comp = rsCorr("numero_correlativo") + 1
'                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'                            rsCorr.Update
'                        End If
'                        MsgBox "NUMERO DE 2da. CUENTA PAC " & cocmCod_Comp
'                        rsCorr.Close
'                rsCoCoM.AddNew
'
'                    rsCoCoM("Cod_Comp") = cocmCod_Comp
'                    rsCoCoM("Tipo_Comp") = cocmTipo_Comp
'                    rsCoCoM("Cod_Trans") = TxtComprobante.Text 'AdoRegularizacion.Recordset("codigo_pago") 'TxtNC.Text 'cocmCod_Trans
'                    rsCoCoM("Cod_Trans_Detalle") = cocmCod_Trans_Detalle
'                    rsCoCoM("org_codigo") = cocmOrg_Codigo
'                    rsCoCoM("Ges_Gestion") = cocmGes_Gestion
'                    rsCoCoM("Num_Respaldo") = cocmNum_Respaldo
'                    rsCoCoM("Fecha_A") = cocmFecha_A
'                    rsCoCoM("Codigo_Beneficiario") = cocmCodigo_Beneficiario
'                    rsCoCoM("Codigo_Documento") = cocmCodigo_Documento
'                    rsCoCoM("Glosa") = cocmGlosa
'                    rsCoCoM("Status") = cocmStatus
'                    rsCoCoM("usr_usuario") = Label7.Caption
'                    rsCoCoM("fecha_registro") = Date
'                    rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
'                rsCoCoM.Update
'                    Set rsdiario = New ADODB.Recordset
'                    If rsdiario.State = 1 Then rsdiario.Close
'                    'rsDiario.Open "select * from co_Diario where Cod_Comp=" & rsCoCoM("Cod_Comp") & "", db, adOpenKeyset, adLockOptimistic
'                    rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
'                    If rsdiario.RecordCount > 0 Then
''                        'Recuperando datos
''                        Set rsCorr = New ADODB.Recordset
''                        If rsCorr.State = 1 Then rsCorr.Close
''                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
''                        If rsCorr.RecordCount > 0 Then
''                            AuxCod_Comp = rsCorr("numero_correlativo") + 1
''                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
''                            rsCorr.Update
''                        End If
'                        'AuxCod_Comp_C = rsDiario("Cod_Comp_C")
'                        AuxCod_Comp = cocmCod_Comp
'                        AuxTipo_Comp = rsdiario("Tipo_Comp")
'                        AuxCod_Comp_C = cocmCod_Comp_C
'                        AuxD_Cuenta = rsdiario("D_Cuenta")
'                        AuxD_Nombre = rsdiario("D_Nombre")
'                        AuxD_SubCta1 = rsdiario("D_SubCta1")
'                        AuxD_SubCta2 = rsdiario("D_SubCta2")
'                        AuxD_Aux1 = rsdiario("D_Aux1")
'                        AuxD_Aux2 = rsdiario("D_Aux2")
'                        AuxD_Aux3 = rsdiario("D_Aux3")
'                        AuxD_Cta_Larga = rsdiario("D_Cta_Larga")
'                        AuxD_Des_Larga = rsdiario("D_Des_Larga")
'                        AuxD_MontoBs = rsdiario("D_MontoBs")
'                        AuxD_MontoDL = rsdiario("D_MontoDL")
'                        AuxD_Cambio = rsdiario("D_Cambio")
'
'                        AuxH_Cuenta = rsdiario("H_Cuenta")
'                        AuxH_Nombre = rsdiario("H_Nombre")
'                        AuxH_SubCta1 = rsdiario("H_SubCta1")
'                        AuxH_SubCta2 = rsdiario("H_SubCta2")
'                        AuxH_Aux1 = rsdiario("H_Aux1")
'                        AuxH_Aux2 = rsdiario("H_Aux2")
'                        AuxH_Aux3 = rsdiario("H_Aux3")
'                        AuxH_Cta_Larga = rsdiario("H_Cta_Larga")
'                        AuxH_Des_Larga = rsdiario("H_Des_Larga")
'                        AuxH_MontoBs = rsdiario("H_MontoBs")
'                        AuxH_MontoDL = rsdiario("H_MontoDL")
'                        AuxH_Cambio = rsdiario("H_Cambio")
'
'                        AuxUsr_Usuario = IIf(IsNull(rsdiario("Usr_Usuario")), "", rsdiario("Usr_Usuario"))
'                        AuxFecha_Registro = rsdiario("Fecha_Registro")
'                        AuxHora_Registro = IIf(IsNull(rsdiario("Hora_Registro")), Time, rsdiario("Hora_Registro"))
'
'                        'Adicionando una copia del registro
'                        rsdiario.AddNew
'                        rsdiario("Cod_Comp") = AuxCod_Comp
'                        rsdiario("Tipo_Comp") = AuxTipo_Comp
'                        rsdiario("Cod_Comp_C") = AuxCod_Comp_C
'
'                        rsdiario("D_Cuenta") = AuxH_Cuenta
'                        rsdiario("D_Nombre") = AuxH_Nombre
'                        rsdiario("D_SubCta1") = AuxH_SubCta1
'                        rsdiario("D_SubCta2") = AuxH_SubCta2
'                        rsdiario("D_Aux1") = AuxH_Aux1
'                        rsdiario("D_Aux2") = AuxH_Aux2
'                        rsdiario("D_Aux3") = AuxH_Aux3
'                        rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
'                        rsdiario("D_Des_Larga") = AuxH_Des_Larga
'                        rsdiario("D_MontoBs") = AuxH_MontoBs
'                        rsdiario("D_MontoDL") = AuxH_MontoDL
'                        rsdiario("D_Cambio") = AuxH_Cambio
'
'                        rsdiario("H_Cuenta") = AuxD_Cuenta
'                        rsdiario("H_Nombre") = AuxD_Nombre
'                        rsdiario("H_SubCta1") = AuxD_SubCta1
'                        rsdiario("H_SubCta2") = AuxD_SubCta2
'                        rsdiario("H_Aux1") = AuxD_Aux1
'                        rsdiario("H_Aux2") = AuxD_Aux2
'                        rsdiario("H_Aux3") = AuxD_Aux3
'                        rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
'                        rsdiario("H_Cta_Larga") = AuxD_Des_Larga
'                        rsdiario("H_MontoBs") = AuxD_MontoBs
'                        rsdiario("H_MontoDL") = AuxD_MontoDL
'                        rsdiario("H_Cambio") = AuxD_Cambio
'
'                        rsdiario("Usr_Usuario") = AuxUsr_Usuario
'                        rsdiario("Fecha_Registro") = AuxFecha_Registro
'                        rsdiario("Hora_Registro") = Format(AuxHora_Registro, "hh:mm:ss")
'                        rsdiario.Update
'                End If
'                  Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'              End If
'          Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'    End If
'       Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'End If
'End If
'End Sub

'Public Sub Reversion_DAC()
'    'Devoluci�n contablemente
'    'recogiendo los datos de devolucion Nro de comprobante al que pertenece la devoluci�n
'    Set rsdev = New ADODB.Recordset
'    If rsdev.State = 1 Then rsdev.Close
'    rsdev.Open "select * from pagos where codigo_pago='" & AdoRegularizacion.Recordset("codigo_pago") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "' and ges_gestion='" & AdoRegularizacion.Recordset("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
'    If rsdev.RecordCount > 0 Then
'            Set rsCoCoM = New ADODB.Recordset
'            If rsCoCoM.State = 1 Then rsCoCoM.Close
'            'Verificar en PAC-DAC
'            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC' ", db, adOpenKeyset, adLockOptimistic
'            If rsCoCoM.RecordCount > 0 Then
'                'Creaci�n de la cabecera o registros maestro en CO_COMPROBANTE_M
'                'Recuperando datos de co_comprobante_m
'                cocmCod_CompDiario = rsCoCoM("Cod_Comp")
'                cocmTipo_Comp = rsCoCoM("Tipo_Comp")
'                cocmCod_Trans = rsCoCoM("Cod_Trans")
'                cocmCod_Trans_Detalle = rsCoCoM("Cod_Trans_Detalle")
'                cocmOrg_Codigo = rsCoCoM("Org_Codigo")
'                cocmGes_Gestion = rsCoCoM("Ges_Gestion")
'                cocmNum_Respaldo = rsCoCoM("Num_Respaldo")
'                cocmFecha_A = rsCoCoM("Fecha_A")
'                cocmCodigo_Beneficiario = rsCoCoM("Codigo_Beneficiario")
'                cocmCodigo_Documento = rsCoCoM("Codigo_Documento")
'                cocmGlosa = rsCoCoM("Glosa")
'                cocmStatus = rsCoCoM("Status")
'                cocmUsr_usuario = rsCoCoM("Usr_Usuario")
'                'Adicionando un nuevo registro
'                'Generando nuevo c�digo
'                        Set rsCorr = New ADODB.Recordset
'                        If rsCorr.State = 1 Then rsCorr.Close
'                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'                        If rsCorr.RecordCount > 0 Then
'                            cocmCod_Comp = rsCorr("numero_correlativo") + 1
'                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'                            rsCorr.Update
'                        End If
'                        rsCorr.Close
'                        MsgBox "NUMERO DE 1era. CUENTA DAC" & cocmCod_Comp
'                rsCoCoM.AddNew
'                    rsCoCoM("Cod_Comp") = cocmCod_Comp
'                    rsCoCoM("Tipo_Comp") = cocmTipo_Comp
'                    rsCoCoM("Cod_Trans") = cocmCod_Trans
'                    rsCoCoM("Cod_Trans_Detalle") = cocmCod_Trans_Detalle
'                    rsCoCoM("org_codigo") = cocmOrg_Codigo
'                    rsCoCoM("Ges_Gestion") = cocmGes_Gestion
'                    rsCoCoM("Num_Respaldo") = cocmNum_Respaldo
'                    rsCoCoM("Fecha_A") = cocmFecha_A
'                    rsCoCoM("Codigo_Beneficiario") = cocmCodigo_Beneficiario
'                    rsCoCoM("Codigo_Documento") = cocmCodigo_Documento
'                    rsCoCoM("Glosa") = cocmGlosa
'                    rsCoCoM("Status") = cocmStatus
'                    rsCoCoM("usr_usuario") = Label7.Caption
'                    rsCoCoM("fecha_registro") = Date
'                    rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
'                rsCoCoM.Update
'
'                Set rsdiario = New ADODB.Recordset
'                If rsdiario.State = 1 Then rsdiario.Close
'                rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
'                'rsDiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_Comp & "", db, adOpenKeyset, adLockOptimistic
'                If rsdiario.RecordCount > 0 Then
'                    AuxCod_Comp = cocmCod_Comp
'                    AuxTipo_Comp = rsdiario("Tipo_Comp")
'                    AuxCod_Comp_C = IIf(IsNull(rsdiario("Cod_Comp_C")), 0, rsdiario("Cod_Comp_C"))
'                    AuxD_Cuenta = rsdiario("D_Cuenta")
'                    AuxD_Nombre = rsdiario("D_Nombre")
'                    AuxD_SubCta1 = rsdiario("D_SubCta1")
'                    AuxD_SubCta2 = rsdiario("D_SubCta2")
'                    AuxD_Aux1 = rsdiario("D_Aux1")
'                    AuxD_Aux2 = rsdiario("D_Aux2")
'                    AuxD_Aux3 = rsdiario("D_Aux3")
'                    AuxD_Cta_Larga = IIf(IsNull(rsdiario("D_Cta_Larga")), "-", rsdiario("D_Cta_Larga"))
'                    AuxD_Des_Larga = IIf(IsNull(rsdiario("D_Des_Larga")), "-", rsdiario("D_Des_Larga"))
'                    AuxD_MontoBs = rsdiario("D_MontoBs")
'                    AuxD_MontoDL = rsdiario("D_MontoDL")
'                    AuxD_Cambio = rsdiario("D_Cambio")
'
'                    AuxH_Cuenta = rsdiario("H_Cuenta")
'                    AuxH_Nombre = rsdiario("H_Nombre")
'                    AuxH_SubCta1 = rsdiario("H_SubCta1")
'                    AuxH_SubCta2 = rsdiario("H_SubCta2")
'                    AuxH_Aux1 = rsdiario("H_Aux1")
'                    AuxH_Aux2 = rsdiario("H_Aux2")
'                    AuxH_Aux3 = rsdiario("H_Aux3")
'                    AuxH_Cta_Larga = rsdiario("H_Cta_Larga")
'                    AuxH_Des_Larga = rsdiario("H_Des_Larga")
'                    AuxH_MontoBs = rsdiario("H_MontoBs")
'                    AuxH_MontoDL = rsdiario("H_MontoDL")
'                    AuxH_Cambio = rsdiario("H_Cambio")
'
'                    AuxUsr_Usuario = rsdiario("Usr_Usuario")
'                    AuxFecha_Registro = rsdiario("Fecha_Registro")
'                    AuxHora_Registro = Format(Time, "hh:mm:ss")
'
'                    'Adicionando una copia del registro
'                    rsdiario.AddNew
'                    rsdiario("Cod_Comp") = AuxCod_Comp 'AuxCod_Comp_C
'                    rsdiario("Tipo_Comp") = AuxTipo_Comp
'                    rsdiario("Cod_Comp_C") = AuxCod_Comp_C
'
'                    rsdiario("D_Cuenta") = AuxH_Cuenta
'                    rsdiario("D_Nombre") = AuxH_Nombre
'                    rsdiario("D_SubCta1") = AuxH_SubCta1
'                    rsdiario("D_SubCta2") = AuxH_SubCta2
'                    rsdiario("D_Aux1") = AuxH_Aux1
'                    rsdiario("D_Aux2") = AuxH_Aux2
'                    rsdiario("D_Aux3") = AuxH_Aux3
'                    rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
'                    rsdiario("D_Cta_Larga") = AuxH_Des_Larga
'                    rsdiario("D_MontoBs") = AuxH_MontoBs
'                    rsdiario("D_MontoDL") = AuxH_MontoDL
'                    rsdiario("D_Cambio") = AuxH_Cambio
'
'                    rsdiario("H_Cuenta") = AuxD_Cuenta
'                    rsdiario("H_Nombre") = AuxD_Nombre
'                    rsdiario("H_SubCta1") = AuxD_SubCta1
'                    rsdiario("H_SubCta2") = AuxD_SubCta2
'                    rsdiario("H_Aux1") = AuxD_Aux1
'                    rsdiario("H_Aux2") = AuxD_Aux2
'                    rsdiario("H_Aux3") = AuxD_Aux3
'                    rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
'                    rsdiario("H_Cta_Larga") = AuxD_Des_Larga
'                    rsdiario("H_MontoBs") = AuxD_MontoBs
'                    rsdiario("H_MontoDL") = AuxD_MontoDL
'                    rsdiario("H_Cambio") = AuxD_Cambio
'
'                    rsdiario("Usr_Usuario") = AuxUsr_Usuario
'                    rsdiario("Fecha_Registro") = AuxFecha_Registro
'                    rsdiario("Hora_Registro") = AuxHora_Registro
'                    rsdiario.Update
'
'                End If
'          Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'    End If
'       Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'End If
'
'End Sub
'Public Sub Anulacion_DAC()
'                'Comprobantes PAC
'                Set rsCoCoM = New ADODB.Recordset
'                If rsCoCoM.State = 1 Then rsCoCoM.Close
'                rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "' and Tipo_Comp='PAC'", db, adOpenKeyset, adLockOptimistic
'                If rsCoCoM.RecordCount > 0 Then
'                    '             Set rsCoCoM = New ADODB.Recordset
'                    '            If rsCoCoM.State = 1 Then rsCoCoM.Close
'                    '            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC'", db, adOpenKeyset, adLockOptimistic
'                    '            If rsCoCoM.RecordCount > 0 Then
''               'Creaci�n de la cabecera o registros maestro en CO_COMPROBANTE_M
'                'Recuperando datos de co_comprobante_m
'                cocmCod_CompDiario = rsCoCoM("Cod_Comp")
'                cocmTipo_Comp = rsCoCoM("Tipo_Comp")
'                cocmCod_Trans = rsCoCoM("Cod_Trans")
'                cocmCod_Trans_Detalle = rsCoCoM("Cod_Trans_Detalle")
'                cocmOrg_Codigo = rsCoCoM("Org_Codigo")
'                cocmGes_Gestion = rsCoCoM("Ges_Gestion")
'                cocmNum_Respaldo = rsCoCoM("Num_Respaldo")
'                cocmFecha_A = rsCoCoM("Fecha_A")
'                cocmCodigo_Beneficiario = rsCoCoM("Codigo_Beneficiario")
'                cocmCodigo_Documento = rsCoCoM("Codigo_Documento")
'                cocmGlosa = rsCoCoM("Glosa")
'                cocmStatus = rsCoCoM("Status")
'                cocmUsr_usuario = IIf(IsNull(rsCoCoM("Usr_Usuario")), "", rsCoCoM("Usr_Usuario"))
'                'Adicionando un nuevo registro
'                'Generando nuevo c�digo
'                'Segunda genera*********
'                        Set rsCorr = New ADODB.Recordset
'                        If rsCorr.State = 1 Then rsCorr.Close
'                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'                        If rsCorr.RecordCount > 0 Then
'                            cocmCod_Comp = rsCorr("numero_correlativo") + 1
'                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'                            rsCorr.Update
'                        End If
'                        rsCorr.Close
'                        MsgBox "NUMERO DE 1era. CUENTA PAC" & cocmCod_Comp
'                rsCoCoM.AddNew
'
'                    rsCoCoM("Cod_Comp") = cocmCod_Comp
'                    rsCoCoM("Tipo_Comp") = cocmTipo_Comp
'                    rsCoCoM("Cod_Trans") = cocmCod_Trans
'                    rsCoCoM("Cod_Trans_Detalle") = cocmCod_Trans_Detalle
'                    rsCoCoM("org_codigo") = cocmOrg_Codigo
'                    rsCoCoM("Ges_Gestion") = cocmGes_Gestion
'                    rsCoCoM("Num_Respaldo") = cocmNum_Respaldo
'                    rsCoCoM("Fecha_A") = cocmFecha_A
'                    rsCoCoM("Codigo_Beneficiario") = cocmCodigo_Beneficiario
'                    rsCoCoM("Codigo_Documento") = cocmCodigo_Documento
'                    rsCoCoM("Glosa") = cocmGlosa
'                    rsCoCoM("Status") = cocmStatus
'                    rsCoCoM("usr_usuario") = Label7.Caption
'                    rsCoCoM("fecha_registro") = Date
'                    rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
'                rsCoCoM.Update
'                    Set rsdiario = New ADODB.Recordset
'                    If rsdiario.State = 1 Then rsdiario.Close
'                    'rsDiario.Open "select * from co_Diario where Cod_Comp=" & rsCoCoM("Cod_Comp") & "", db, adOpenKeyset, adLockOptimistic
'                    rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
'                    If rsdiario.RecordCount > 0 Then
''                        'Recuperando datos
''                        Set rsCorr = New ADODB.Recordset
''                        If rsCorr.State = 1 Then rsCorr.Close
''                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
''                        If rsCorr.RecordCount > 0 Then
''                            AuxCod_Comp = rsCorr("numero_correlativo") + 1
''                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
''                            rsCorr.Update
''                        End If
'                        'AuxCod_Comp_C = rsDiario("Cod_Comp_C")
'                        AuxCod_Comp = cocmCod_Comp
'                        AuxTipo_Comp = rsdiario("Tipo_Comp")
'                        AuxCod_Comp_C = cocmCod_Comp_C
'                        AuxD_Cuenta = rsdiario("D_Cuenta")
'                        AuxD_Nombre = rsdiario("D_Nombre")
'                        AuxD_SubCta1 = rsdiario("D_SubCta1")
'                        AuxD_SubCta2 = rsdiario("D_SubCta2")
'                        AuxD_Aux1 = rsdiario("D_Aux1")
'                        AuxD_Aux2 = rsdiario("D_Aux2")
'                        AuxD_Aux3 = rsdiario("D_Aux3")
'                        AuxD_Cta_Larga = rsdiario("D_Cta_Larga")
'                        AuxD_Des_Larga = rsdiario("D_Des_Larga")
'                        AuxD_MontoBs = rsdiario("D_MontoBs")
'    '                    AuxD_MontoDL = rsDiario("D_MontoDL")
'                        AuxD_Cambio = rsdiario("D_Cambio")
'
'                        AuxH_Cuenta = rsdiario("H_Cuenta")
'                        AuxH_Nombre = rsdiario("H_Nombre")
'                        AuxH_SubCta1 = rsdiario("H_SubCta1")
'                        AuxH_SubCta2 = rsdiario("H_SubCta2")
'                        AuxH_Aux1 = rsdiario("H_Aux1")
'                        AuxH_Aux2 = rsdiario("H_Aux2")
'                        AuxH_Aux3 = rsdiario("H_Aux3")
'                        AuxH_Cta_Larga = rsdiario("H_Cta_Larga")
'                        AuxH_Des_Larga = rsdiario("H_Des_Larga")
'                        AuxH_MontoBs = rsdiario("H_MontoBs")
'    '                    AuxH_MontoDL = rsDiario("H_MontoDL")
'                        AuxH_Cambio = rsdiario("H_Cambio")
'
'                        AuxUsr_Usuario = IIf(IsNull(rsdiario("Usr_Usuario")), "", rsdiario("Usr_Usuario"))
'                        AuxFecha_Registro = rsdiario("Fecha_Registro")
'                        AuxHora_Registro = IIf(IsNull(rsdiario("Hora_Registro")), Time, rsdiario("Hora_Registro"))
'
'                        'Adicionando una copia del registro
'                        rsdiario.AddNew
'                        rsdiario("Cod_Comp") = AuxCod_Comp
'                        rsdiario("Tipo_Comp") = AuxTipo_Comp
'                        rsdiario("Cod_Comp_C") = AuxCod_Comp_C
'
'                        rsdiario("D_Cuenta") = AuxH_Cuenta
'                        rsdiario("D_Nombre") = AuxH_Nombre
'                        rsdiario("D_SubCta1") = AuxH_SubCta1
'                        rsdiario("D_SubCta2") = AuxH_SubCta2
'                        rsdiario("D_Aux1") = AuxH_Aux1
'                        rsdiario("D_Aux2") = AuxH_Aux2
'                        rsdiario("D_Aux3") = AuxH_Aux3
'                        rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
'                        rsdiario("D_Des_Larga") = AuxH_Des_Larga
'                        rsdiario("D_MontoBs") = AuxH_MontoBs
'                        'rsDiario("D_MontoDL") = AuxH_MontoDL
'                        rsdiario("D_Cambio") = AuxH_Cambio
'
'                        rsdiario("H_Cuenta") = AuxD_Cuenta
'                        rsdiario("H_Nombre") = AuxD_Nombre
'                        rsdiario("H_SubCta1") = AuxD_SubCta1
'                        rsdiario("H_SubCta2") = AuxD_SubCta2
'                        rsdiario("H_Aux1") = AuxD_Aux1
'                        rsdiario("H_Aux2") = AuxD_Aux2
'                        rsdiario("H_Aux3") = AuxD_Aux3
'                        rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
'                        rsdiario("H_Des_Larga") = AuxD_Des_Larga
'                        rsdiario("H_MontoBs") = AuxD_MontoBs
'                        'rsDiario("H_MontoDL") = AuxD_MontoDL
'                        rsdiario("H_Cambio") = AuxD_Cambio
'
'                        rsdiario("Usr_Usuario") = AuxUsr_Usuario
'                        rsdiario("Fecha_Registro") = AuxFecha_Registro
'                        rsdiario("Hora_Registro") = Format(AuxHora_Registro, "hh:mm:ss")
'                        rsdiario.Update
'                End If
'                  Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'              End If
'End Sub

'Public Sub Anulacion_DAC()
'    'Comprobantes PAC
'  db.BeginTrans
'    Set rsCoCoM = New ADODB.Recordset
'    If rsCoCoM.State = 1 Then rsCoCoM.Close
'    rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & AdoRegularizacion.Recordset("Nro_Comprobante_Anterior") & "' and org_codigo='" & AdoRegularizacion.Recordset("org_codigo") & "' and Tipo_Comp='PAC'", db, adOpenKeyset, adLockOptimistic
'    If rsCoCoM.RecordCount > 0 Then
'        '             Set rsCoCoM = New ADODB.Recordset
'        '            If rsCoCoM.State = 1 Then rsCoCoM.Close
'        '            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC'", db, adOpenKeyset, adLockOptimistic
'        '            If rsCoCoM.RecordCount > 0 Then
''               'Creaci�n de la cabecera o registros maestro en CO_COMPROBANTE_M
'    'Recuperando datos de co_comprobante_m
'    cocmCod_CompDiario = IIf(IsNull(rsCoCoM("Cod_Comp")), " ", rsCoCoM("Cod_Comp"))
'    cocmTipo_Comp = IIf(IsNull(rsCoCoM("Tipo_Comp")), " ", rsCoCoM("Tipo_Comp"))
'    cocmCod_Trans = IIf(IsNull(rsCoCoM("Cod_Trans")), " ", rsCoCoM("cod_trans"))
'    cocmCod_Trans_Detalle = IIf(IsNull(rsCoCoM("Cod_Trans_Detalle")), "", (rsCoCoM("Cod_Trans_Detalle")))
'    cocmOrg_Codigo = IIf(IsNull(rsCoCoM("Org_Codigo")), "", rsCoCoM("Org_Codigo"))
'    cocmGes_Gestion = IIf(IsNull(rsCoCoM("Ges_Gestion")), "", rsCoCoM("Ges_Gestion"))
'    cocmNum_Respaldo = IIf(IsNull(rsCoCoM("Num_Respaldo")), "", rsCoCoM("Num_Respaldo"))
'    cocmFecha_A = CDate(rsCoCoM("Fecha_A"))
'    cocmCodigo_Beneficiario = IIf(IsNull(rsCoCoM("Codigo_Beneficiario")), "", rsCoCoM("Codigo_Beneficiario"))
'    cocmCodigo_Documento = IIf(IsNull(rsCoCoM("Codigo_Documento")), "", rsCoCoM("Codigo_Documento"))
'    cocmGlosa = IIf(IsNull(rsCoCoM("Glosa")), "", rsCoCoM("Glosa"))
'    cocmStatus = IIf(IsNull(rsCoCoM("Status")), "", rsCoCoM("Status"))
'    cocmUsr_usuario = IIf(IsNull(rsCoCoM("Usr_Usuario")), "", rsCoCoM("Usr_Usuario"))
'    'Adicionando un nuevo registro
'    'Generando nuevo c�digo
'    'Segunda genera*********
'            Set rsCorr = New ADODB.Recordset
'            If rsCorr.State = 1 Then rsCorr.Close
'            rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'            If rsCorr.RecordCount > 0 Then
'                cocmCod_Comp = rsCorr("numero_correlativo") + 1
'                rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'                rsCorr.Update
'            End If
'            rsCorr.Close
'            MsgBox "NUMERO DE 1era. CUENTA PAC" & cocmCod_Comp
'    rsCoCoM.AddNew
'
'        rsCoCoM("Cod_Comp") = cocmCod_Comp
'        rsCoCoM("Tipo_Comp") = Trim(cocmTipo_Comp)
'        rsCoCoM("Cod_Trans") = Trim(cocmCod_Trans)
'        rsCoCoM("Cod_Trans_Detalle") = Trim(cocmCod_Trans_Detalle)
'        rsCoCoM("org_codigo") = Trim(cocmOrg_Codigo)
'        rsCoCoM("Ges_Gestion") = Trim(cocmGes_Gestion)
'        rsCoCoM("Num_Respaldo") = Trim(cocmNum_Respaldo)
'        rsCoCoM("Fecha_A") = CDate(cocmFecha_A)
'        rsCoCoM("Codigo_Beneficiario") = Trim(cocmCodigo_Beneficiario)
'        rsCoCoM("Codigo_Documento") = Trim(cocmCodigo_Documento)
'        rsCoCoM("Glosa") = Trim(cocmGlosa)
'        rsCoCoM("Status") = Trim(cocmStatus)
'        rsCoCoM("usr_usuario") = Label7.Caption
'        rsCoCoM("fecha_registro") = CDate(Format(Date, "dd/mm/yyyy"))
'        rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
'    rsCoCoM.Update
'        Set rsdiario = New ADODB.Recordset
'        If rsdiario.State = 1 Then rsdiario.Close
'        'rsDiario.Open "select * from co_Diario where Cod_Comp=" & rsCoCoM("Cod_Comp") & "", db, adOpenKeyset, adLockOptimistic
'        rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
'        If rsdiario.RecordCount > 0 Then
''                        'Recuperando datos
''                        Set rsCorr = New ADODB.Recordset
''                        If rsCorr.State = 1 Then rsCorr.Close
''                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
''                        If rsCorr.RecordCount > 0 Then
''                            AuxCod_Comp = rsCorr("numero_correlativo") + 1
''                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
''                            rsCorr.Update
''                        End If
'            'AuxCod_Comp_C = rsDiario("Cod_Comp_C")
'            AuxCod_Comp = cocmCod_Comp
'            AuxTipo_Comp = IIf(IsNull(rsdiario("Tipo_Comp")), "", rsdiario("Tipo_Comp"))
'            AuxCod_Comp_C = IIf(IsNull(cocmCod_Comp_C), 0, cocmCod_Comp_C)
'            AuxD_Cuenta = rsdiario("D_Cuenta")
'            AuxD_Nombre = IIf(IsNull(rsdiario("D_Nombre")), "", rsdiario("D_Nombre"))
'            AuxD_SubCta1 = rsdiario("D_SubCta1")
'            AuxD_SubCta2 = rsdiario("D_SubCta2")
'            AuxD_Aux1 = rsdiario("D_Aux1")
'            AuxD_Aux2 = rsdiario("D_Aux2")
'            AuxD_Aux3 = rsdiario("D_Aux3")
'            AuxD_Cta_Larga = IIf(IsNull(rsdiario("D_Cta_Larga")), "", rsdiario("D_Cta_Larga"))
'            AuxD_Des_Larga = IIf(IsNull(rsdiario("D_Des_Larga")), "", rsdiario("D_Des_Larga"))
'            AuxD_MontoBs = rsdiario("D_MontoBs")
'            AuxD_MontoDL = rsdiario("D_MontoDL")
'            AuxD_Cambio = rsdiario("D_Cambio")
'
'            AuxH_Cuenta = rsdiario("H_Cuenta")
'            AuxH_Nombre = IIf(IsNull(rsdiario("H_Nombre")), "", rsdiario("H_Nombre"))
'            AuxH_SubCta1 = rsdiario("H_SubCta1")
'            AuxH_SubCta2 = rsdiario("H_SubCta2")
'            AuxH_Aux1 = rsdiario("H_Aux1")
'            AuxH_Aux2 = rsdiario("H_Aux2")
'            AuxH_Aux3 = rsdiario("H_Aux3")
'            AuxH_Cta_Larga = IIf(IsNull(rsdiario("H_Cta_Larga")), "", rsdiario("H_Cta_Larga"))
'            AuxH_Des_Larga = IIf(IsNull(rsdiario("H_Des_Larga")), "", rsdiario("H_Des_Larga"))
'            AuxH_MontoBs = rsdiario("H_MontoBs")
'            AuxH_MontoDL = rsdiario("H_MontoDL")
'            AuxH_Cambio = rsdiario("H_Cambio")
'
'            AuxUsr_Usuario = IIf(IsNull(rsdiario("Usr_Usuario")), "", rsdiario("Usr_Usuario"))
'            AuxFecha_Registro = CDate(rsdiario("Fecha_Registro"))
'            AuxHora_Registro = IIf(IsNull(rsdiario("Hora_Registro")), Time, rsdiario("Hora_Registro"))
'
'            'Adicionando una copia del registro
'            rsdiario.AddNew
'            rsdiario("Cod_Comp") = AuxCod_Comp
'            rsdiario("Tipo_Comp") = Trim(AuxTipo_Comp)
'            rsdiario("Cod_Comp_C") = AuxCod_Comp_C
'
'            rsdiario("D_Cuenta") = AuxH_Cuenta
'            rsdiario("D_Nombre") = IIf(IsNull(AuxH_Nombre), "", AuxH_Nombre)
'            rsdiario("D_SubCta1") = AuxH_SubCta1
'            rsdiario("D_SubCta2") = AuxH_SubCta2
'            rsdiario("D_Aux1") = AuxH_Aux1
'            rsdiario("D_Aux2") = AuxH_Aux2
'            rsdiario("D_Aux3") = AuxH_Aux3
'            rsdiario("D_Cta_Larga") = IIf(IsNull(AuxH_Cta_Larga), "", AuxH_Cta_Larga)
'            rsdiario("D_Des_Larga") = IIf(IsNull(AuxH_Des_Larga), "", AuxH_Des_Larga)
'            rsdiario("D_MontoBs") = AuxH_MontoBs
'            rsdiario("D_MontoDL") = AuxH_MontoDL
'            rsdiario("D_Cambio") = AuxH_Cambio
'
'            rsdiario("H_Cuenta") = AuxD_Cuenta
'            rsdiario("H_Nombre") = IIf(IsNull(AuxD_Nombre), "", AuxD_Nombre)
'            rsdiario("H_SubCta1") = AuxD_SubCta1
'            rsdiario("H_SubCta2") = AuxD_SubCta2
'            rsdiario("H_Aux1") = AuxD_Aux1
'            rsdiario("H_Aux2") = AuxD_Aux2
'            rsdiario("H_Aux3") = AuxD_Aux3
'            rsdiario("H_Cta_Larga") = IIf(IsNull(AuxD_Cta_Larga), "", AuxD_Cta_Larga)
'            rsdiario("H_Des_Larga") = IIf(IsNull(AuxD_Des_Larga), "", AuxD_Des_Larga)
'            rsdiario("H_MontoBs") = AuxD_MontoBs
'            rsdiario("H_MontoDL") = AuxD_MontoDL
'            rsdiario("H_Cambio") = AuxD_Cambio
'
'            rsdiario("Usr_Usuario") = AuxUsr_Usuario
'            rsdiario("Fecha_Registro") = CDate(AuxFecha_Registro)
'            rsdiario("Hora_Registro") = Format(AuxHora_Registro, "hh:mm:ss")
'            rsdiario.Update
'    End If
'      Else: MsgBox "No se contabiliz�", vbCritical + vbInformation, "CONTABILIZACION"
'  End If
' db.CommitTrans
'End Sub
Private Sub pOrganismo(CodFuente As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from fc_organismo_financiamiento where fte_codigo='" & CodFuente & "'"
   
   Set DtcOrg.RowSource = Nothing
   Set DtcOrg.RowSource = db.Execute(strConsultaF, , adCmdText)
   DtcOrg.ReFill
   DtcOrg.BoundText = Empty
   
   Set DtcDesOrg.RowSource = Nothing
   Set DtcDesOrg.RowSource = db.Execute(strConsultaF, , adCmdText)
   DtcDesOrg.ReFill
   DtcDesOrg.BoundText = Empty

End Sub

Private Sub TxtMontoFuente_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub


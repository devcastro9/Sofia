VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fw_conciliacion_bancaria 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Financiero - Tesorer�a - Conciliaci�n Bancaria"
   ClientHeight    =   8790
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14835
   Icon            =   "fw_conciliacion_bancaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.Frame FraBusca4 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   2175
      Left            =   7080
      TabIndex        =   91
      Top             =   3480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtBusca 
         Height          =   285
         Left            =   1080
         TabIndex        =   97
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3000
         TabIndex        =   96
         Text            =   "0"
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5520
         TabIndex        =   92
         Top             =   240
         Width           =   5520
         Begin VB.PictureBox BtnBuscar4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   1200
            Picture         =   "fw_conciliacion_bancaria.frx":0A02
            ScaleHeight     =   585
            ScaleWidth      =   1275
            TabIndex        =   94
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   3000
            Picture         =   "fw_conciliacion_bancaria.frx":12C1
            ScaleHeight     =   585
            ScaleWidth      =   1395
            TabIndex        =   93
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   14175
            TabIndex        =   95
            Top             =   195
            Width           =   1005
         End
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   3000
         TabIndex        =   98
         Top             =   1920
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   42880
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ESCRIBA DATOS PARA BUSCAR EN TODAS  COLUMANAS"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   101
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO MENOR DE BUSQUEDA"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   100
         Top             =   2160
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIO BUSQUEDA"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   99
         Top             =   1920
         Visible         =   0   'False
         Width           =   1995
      End
   End
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FFFF00&
      Height          =   2895
      Left            =   7080
      TabIndex        =   76
      Top             =   2640
      Visible         =   0   'False
      Width           =   5775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5520
         TabIndex        =   79
         Top             =   240
         Width           =   5520
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3000
            Picture         =   "fw_conciliacion_bancaria.frx":1C80
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   81
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1200
            Picture         =   "fw_conciliacion_bancaria.frx":256C
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   80
            Top             =   0
            Width           =   1280
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   14175
            TabIndex        =   82
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   3000
         TabIndex        =   78
         Text            =   "0"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   240
         TabIndex        =   77
         Text            =   "0"
         Top             =   2400
         Visible         =   0   'False
         Width           =   5295
      End
      Begin MSComCtl2.DTPicker DTP_Finicio 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   3000
         TabIndex        =   83
         Top             =   1200
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   44457
      End
      Begin MSComCtl2.DTPicker DTP_Ffin 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   3000
         TabIndex        =   84
         Top             =   1200
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   42880
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA TRANSACCION"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   87
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE DEPOSITO"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   86
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCIAS DEL DEPOSITANTE"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   85
         Top             =   2160
         Visible         =   0   'False
         Width           =   2685
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   19245
      _ExtentX        =   33946
      _ExtentY        =   17171
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "IMPORTAR EXTRACTO BANCARIO"
      TabPicture(0)   =   "fw_conciliacion_bancaria.frx":2D42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOpciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_ABM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraNavega"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnImportarDato"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnCargarArchivo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "CONCILIACION BANCARIA"
      TabPicture(1)   =   "fw_conciliacion_bancaria.frx":2D5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BtnImprimir2"
      Tab(1).Control(1)=   "BtnBuscar2"
      Tab(1).Control(2)=   "BtnBuscar3"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "FrmDetalle"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "FrmDetalle2"
      Tab(1).Control(7)=   "FrmABMDet"
      Tab(1).ControlCount=   8
      Begin VB.PictureBox BtnImprimir2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   -74940
         Picture         =   "fw_conciliacion_bancaria.frx":2D7A
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   102
         ToolTipText     =   "Busca Registros CONCILIADOS"
         Top             =   7800
         Width           =   1390
      End
      Begin VB.PictureBox BtnBuscar2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   -74880
         Picture         =   "fw_conciliacion_bancaria.frx":3A55
         ScaleHeight     =   585
         ScaleWidth      =   1335
         TabIndex        =   90
         ToolTipText     =   "Busca Registros CONCILIADOS"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.PictureBox BtnBuscar3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -74880
         Picture         =   "fw_conciliacion_bancaria.frx":4314
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   89
         ToolTipText     =   "Busca Registros CONCILIADOS (TODAS LAS COLUMNAS)"
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "REGISTROS CONCILIADOS - EXTRACTOS BANCARIOS BMSC"
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
         Height          =   2360
         Left            =   -73440
         TabIndex        =   65
         Top             =   5400
         Width           =   17655
         Begin MSDataGridLib.DataGrid dg_datos3 
            Bindings        =   "fw_conciliacion_bancaria.frx":4AC9
            Height          =   1695
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   17400
            _ExtentX        =   30692
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483624
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
            ColumnCount     =   21
            BeginProperty Column00 
               DataField       =   "fecha_transaccion"
               Caption         =   "Fecha.Transaccion"
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
               DataField       =   "cod_bancarizacion"
               Caption         =   "Cod.Bancarizacion"
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
            BeginProperty Column02 
               DataField       =   "nro_cheque"
               Caption         =   "Nro.Cheque"
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
            BeginProperty Column03 
               DataField       =   "plantilla"
               Caption         =   "Plantilla"
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
            BeginProperty Column04 
               DataField       =   "cod_cliente"
               Caption         =   "Cod.Cliente"
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
            BeginProperty Column05 
               DataField       =   "id_depositante"
               Caption         =   "Id.Depositante"
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
            BeginProperty Column06 
               DataField       =   "nombre_depositante"
               Caption         =   "Nombre.Depositante"
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
            BeginProperty Column07 
               DataField       =   "tipo_transaccion"
               Caption         =   "Tipo.Transaccion"
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
            BeginProperty Column08 
               DataField       =   "descripcion"
               Caption         =   "Descripcion"
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
            BeginProperty Column09 
               DataField       =   "oficina"
               Caption         =   "Oficina"
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
            BeginProperty Column10 
               DataField       =   "banco"
               Caption         =   "Banco"
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
            BeginProperty Column11 
               DataField       =   "tipo_deposito"
               Caption         =   "Tipo.Deposito"
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
            BeginProperty Column12 
               DataField       =   "nombre_destinatario"
               Caption         =   "Nombre.Destinatario"
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
            BeginProperty Column13 
               DataField       =   "glosa"
               Caption         =   "Glosa"
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
            BeginProperty Column14 
               DataField       =   "originador"
               Caption         =   "Originador"
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
            BeginProperty Column15 
               DataField       =   "originador_ACH"
               Caption         =   "Originador_ACH"
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
            BeginProperty Column16 
               DataField       =   "ciudad_origen"
               Caption         =   "Ciudad_origen"
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
            BeginProperty Column17 
               DataField       =   "debito"
               Caption         =   "Debito"
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
            BeginProperty Column18 
               DataField       =   "credito"
               Caption         =   "Credito"
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
            BeginProperty Column19 
               DataField       =   "Cuenta"
               Caption         =   "Cuenta"
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
            BeginProperty Column20 
               DataField       =   "estado_conciliado"
               Caption         =   "Estado_conciliado"
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
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
               EndProperty
               BeginProperty Column07 
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
               EndProperty
               BeginProperty Column13 
               EndProperty
               BeginProperty Column14 
               EndProperty
               BeginProperty Column15 
               EndProperty
               BeginProperty Column16 
               EndProperty
               BeginProperty Column17 
               EndProperty
               BeginProperty Column18 
               EndProperty
               BeginProperty Column19 
               EndProperty
               BeginProperty Column20 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_datos3 
            Height          =   330
            Left            =   120
            Top             =   1920
            Width           =   17415
            _ExtentX        =   30718
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
            Caption         =   "REGISTROS CONCILIADOS - EXTRACTOS BANCARIOS BMSC"
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
      Begin VB.Frame FrmDetalle 
         BackColor       =   &H00C0C0C0&
         Caption         =   "REGISTROS CONCILIADOS - SOFIA"
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
         Height          =   1905
         Left            =   -73440
         TabIndex        =   61
         Top             =   7760
         Width           =   17655
         Begin MSDataGridLib.DataGrid DtGLista 
            Bindings        =   "fw_conciliacion_bancaria.frx":4AE2
            Height          =   1500
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   17460
            _ExtentX        =   30798
            _ExtentY        =   2646
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   13
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   19
            BeginProperty Column00 
               DataField       =   "Correl_doc"
               Caption         =   "Rbo.Tes."
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
               DataField       =   "cobranza_fecha"
               Caption         =   "Fecha.Recibo"
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
               DataField       =   "doc_numero"
               Caption         =   "Recibo.Cobr."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "trans_descripcion"
               Caption         =   "Tipo.Transac."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "tipo_moneda"
               Caption         =   "Moneda"
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
               DataField       =   "cmpbte_deposito"
               Caption         =   "#Cheque/Transf."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4105
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "cmpbte_deposito_bco"
               Caption         =   "#Cmpbte.Deposito"
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
               DataField       =   "fecha_registro_bco"
               Caption         =   "Fecha.Deposito"
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
               DataField       =   "cta_codigo"
               Caption         =   "Cuenta.Bancaria"
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
               DataField       =   "cobranza_bs"
               Caption         =   "Cobrado Bs."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "cobranza_dol"
               Caption         =   "Cobrado Dol."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "cobranza_observaciones"
               Caption         =   "Concepto"
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
            BeginProperty Column12 
               DataField       =   "estado_codigo_bco"
               Caption         =   "Cobrado"
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
            BeginProperty Column13 
               DataField       =   "edif_codigo_corto"
               Caption         =   "Edificio"
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
            BeginProperty Column14 
               DataField       =   "edif_descripcion"
               Caption         =   "Nombre.Edificio"
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
            BeginProperty Column15 
               DataField       =   "correl_doc_trp"
               Caption         =   "Nro.Traspaso"
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
            BeginProperty Column16 
               DataField       =   "IdRecibo"
               Caption         =   "Id.Tes."
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
            BeginProperty Column17 
               DataField       =   "estado_codigo_rbo"
               Caption         =   "Aceptado"
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
            BeginProperty Column18 
               DataField       =   "usr_codigo"
               Caption         =   "Usuario"
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
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1260.284
               EndProperty
               BeginProperty Column06 
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1319.811
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column10 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   3674.835
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   4305.26
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   1065.26
               EndProperty
               BeginProperty Column16 
                  Alignment       =   2
                  ColumnWidth     =   659.906
               EndProperty
               BeginProperty Column17 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   780.095
               EndProperty
               BeginProperty Column18 
                  Alignment       =   2
                  Object.Visible         =   0   'False
                  ColumnWidth     =   645.165
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "REGISTROS SIN CONCILIAR - EXTRACTOS BANCARIOS BMSC"
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
         Height          =   2295
         Left            =   -73440
         TabIndex        =   60
         Top             =   3100
         Width           =   17655
         Begin MSDataGridLib.DataGrid dg_datos2 
            Bindings        =   "fw_conciliacion_bancaria.frx":4AFC
            Height          =   1935
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   17400
            _ExtentX        =   30692
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
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
      End
      Begin VB.Frame FrmDetalle2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "REGISTROS SIN CONCILIAR - SOFIA (TESORERIA)"
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
         Height          =   2625
         Left            =   -73440
         TabIndex        =   58
         Top             =   480
         Width           =   17655
         Begin VB.OptionButton OptFilGral07 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "7. Santa Cruz"
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
            Left            =   12120
            TabIndex        =   75
            Top             =   2280
            Width           =   1425
         End
         Begin VB.OptionButton OptFilGral08 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "8. Beni . . . . "
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
            Left            =   14040
            TabIndex        =   74
            Top             =   2280
            Width           =   1305
         End
         Begin VB.OptionButton OptFilGral09 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "9. Pando . . . "
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
            Left            =   15960
            TabIndex        =   73
            Top             =   2280
            Width           =   1305
         End
         Begin VB.OptionButton OptFilGral06 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "6. Tarija . . . "
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
            Left            =   10200
            TabIndex        =   72
            Top             =   2280
            Width           =   1305
         End
         Begin VB.OptionButton OptFilGral05 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "5. Potosi . . . ."
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
            Left            =   8280
            TabIndex        =   71
            Top             =   2280
            Width           =   1425
         End
         Begin VB.OptionButton OptFilGral01 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "1. Chuquisaca ."
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
            Left            =   240
            TabIndex        =   70
            Top             =   2280
            Width           =   1545
         End
         Begin VB.OptionButton OptFilGral02 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "2. La Paz . . . "
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
            Left            =   2280
            TabIndex        =   69
            Top             =   2280
            Width           =   1305
         End
         Begin VB.OptionButton OptFilGral04 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "4. Oruro . . . "
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
            Left            =   6360
            TabIndex        =   68
            Top             =   2280
            Width           =   1305
         End
         Begin VB.OptionButton OptFilGral03 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "3. Cochabamba "
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
            Left            =   4200
            TabIndex        =   67
            Top             =   2280
            Width           =   1550
         End
         Begin MSDataGridLib.DataGrid DtGLista11 
            Bindings        =   "fw_conciliacion_bancaria.frx":4B15
            Height          =   1980
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   17460
            _ExtentX        =   30798
            _ExtentY        =   3493
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   13
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   21
            BeginProperty Column00 
               DataField       =   "Correl_Doc"
               Caption         =   "Rbo.Tes."
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
               DataField       =   "cobranza_fecha"
               Caption         =   "Fecha.Recibo"
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
               DataField       =   "doc_numero"
               Caption         =   "doc_numero"
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
            BeginProperty Column03 
               DataField       =   "doc_numero"
               Caption         =   "Recibo.Cobr."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "trans_descripcion"
               Caption         =   "Tipo.Transac."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "tipo_moneda"
               Caption         =   "Moneda"
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
               DataField       =   "cmpbte_deposito"
               Caption         =   "#Cheque/Transf."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4105
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "cmpbte_deposito_bco"
               Caption         =   "#Cmpbte.Deposito"
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
               DataField       =   "fecha_registro_bco"
               Caption         =   "Fecha.Deposito"
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
               DataField       =   "cta_codigo"
               Caption         =   "Cuenta.Bancaria"
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
               DataField       =   "cobranza_bs"
               Caption         =   "Cobrado Bs."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "cobranza_dol"
               Caption         =   "Cobrado Dol."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column12 
               DataField       =   "cobranza_observaciones"
               Caption         =   "Concepto"
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
            BeginProperty Column13 
               DataField       =   "estado_codigo_bco"
               Caption         =   "Cobrado"
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
            BeginProperty Column14 
               DataField       =   "edif_codigo_corto"
               Caption         =   "Edificio"
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
            BeginProperty Column15 
               DataField       =   "edif_descripcion"
               Caption         =   "Nombre.Edificio"
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
            BeginProperty Column16 
               DataField       =   "IdRecibo"
               Caption         =   "Id.Tes."
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
            BeginProperty Column17 
               DataField       =   "estado_codigo_rbo"
               Caption         =   "Aceptado"
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
            BeginProperty Column18 
               DataField       =   "usr_codigo"
               Caption         =   "Usuario"
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
            BeginProperty Column19 
               DataField       =   "depto_codigo"
               Caption         =   "Departamento"
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
            BeginProperty Column20 
               DataField       =   "observaciones"
               Caption         =   "Referencias.del.Depositante"
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
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   689.953
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1260.284
               EndProperty
               BeginProperty Column07 
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1319.811
               EndProperty
               BeginProperty Column10 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column11 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   2594.835
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   2355.024
               EndProperty
               BeginProperty Column16 
                  Alignment       =   2
                  ColumnWidth     =   689.953
               EndProperty
               BeginProperty Column17 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   780.095
               EndProperty
               BeginProperty Column18 
                  Alignment       =   2
                  Object.Visible         =   0   'False
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column19 
               EndProperty
               BeginProperty Column20 
                  ColumnWidth     =   3825.071
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton btnCargarArchivo 
         BackColor       =   &H80000010&
         Caption         =   "Elegir Archivo Excel (Extracto Bancario)"
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
         Left            =   15360
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton btnImportarDato 
         BackColor       =   &H80000010&
         Caption         =   "Importar Datos a SOFIA (para Conciliar)"
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
         Left            =   15360
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   5760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox FrmABMDet 
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         Height          =   4905
         Left            =   -74880
         Negotiate       =   -1  'True
         ScaleHeight     =   20.188
         ScaleMode       =   4  'Character
         ScaleWidth      =   11.625
         TabIndex        =   50
         Top             =   480
         Width           =   1455
         Begin VB.PictureBox BtnBuscar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":4B2F
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   88
            ToolTipText     =   "Busca NO Conciliados en Extracto BMSC"
            Top             =   3000
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":52E4
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   51
            ToolTipText     =   "Imprime Reporte de Conciliaci�n"
            Top             =   2160
            Width           =   1400
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":5BB1
            ScaleHeight     =   735
            ScaleWidth      =   1365
            TabIndex        =   63
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   4200
            Width           =   1365
         End
         Begin VB.CommandButton BtnAnlDetalle 
            BackColor       =   &H80000015&
            Height          =   525
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":6373
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Deshabilitar el Registro activo (volver al paso anterior)"
            Top             =   4275
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton BtnAddDetalle 
            BackColor       =   &H80000015&
            Height          =   525
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":6ABF
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Adiciona Detalle"
            Top             =   2160
            Width           =   1365
         End
         Begin VB.PictureBox BtnBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":72AD
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   53
            ToolTipText     =   "Busca NO Conciliados en SOFIA"
            Top             =   120
            Width           =   1215
         End
         Begin VB.PictureBox BtnModDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":7A62
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   52
            ToolTipText     =   "Modifica Datos en SOFIA"
            Top             =   720
            Visible         =   0   'False
            Width           =   1430
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Borrar registro de Facturaci�n Electr�nica"
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
         Height          =   2355
         Left            =   360
         TabIndex        =   34
         Top             =   4680
         Visible         =   0   'False
         Width           =   9540
         Begin VB.ComboBox cb_aguinaldo 
            Height          =   315
            Left            =   5280
            TabIndex        =   43
            Top             =   840
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "TODAS INTERIOR"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6480
            TabIndex        =   42
            Top             =   2280
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "TODAS LAS PLANILLAS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6480
            TabIndex        =   41
            Top             =   1920
            Width           =   2115
         End
         Begin VB.ComboBox cmb_gestion 
            Height          =   315
            Left            =   1920
            TabIndex        =   40
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cbo_mes_rep 
            Height          =   315
            Left            =   5280
            TabIndex        =   39
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txt_mes 
            BackColor       =   &H00C0C0C0&
            DataField       =   "mes_grupo"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.PictureBox FraGrabarCancelar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            FillColor       =   &H00404040&
            FillStyle       =   2  'Horizontal Line
            ForeColor       =   &H80000008&
            Height          =   676
            Left            =   0
            ScaleHeight     =   675
            ScaleWidth      =   9720
            TabIndex        =   35
            Top             =   1680
            Visible         =   0   'False
            Width           =   9720
            Begin VB.PictureBox BtnCancelar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   4275
               Picture         =   "fw_conciliacion_bancaria.frx":8377
               ScaleHeight     =   615
               ScaleWidth      =   1395
               TabIndex        =   37
               Top             =   0
               Width           =   1400
            End
            Begin VB.PictureBox BtnGrabar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   2880
               Picture         =   "fw_conciliacion_bancaria.frx":8C63
               ScaleHeight     =   615
               ScaleWidth      =   1305
               TabIndex        =   36
               Top             =   0
               Width           =   1300
            End
         End
         Begin MSDataListLib.DataCombo dtc_rep_det 
            DataField       =   "planilla_codigo"
            Height          =   315
            Left            =   2880
            TabIndex        =   44
            Top             =   1920
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "planilla_descripcion"
            BoundColumn     =   "planilla_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_rep_cod 
            DataField       =   "planilla_codigo"
            Height          =   315
            Left            =   1920
            TabIndex        =   45
            Top             =   1920
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "planilla_codigo"
            BoundColumn     =   "planilla_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_depto 
            DataField       =   "planilla_codigo"
            Height          =   315
            Left            =   1920
            TabIndex        =   46
            Top             =   2160
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "depto_codigo"
            BoundColumn     =   "planilla_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_datos_rep 
            Height          =   330
            Left            =   120
            Top             =   240
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
            Caption         =   "Ado_cuenta"
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
         Begin VB.Label Label34 
            BackColor       =   &H00C0C0C0&
            Caption         =   "PLANILLA"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   49
            Top             =   1935
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0C0C0&
            Caption         =   "MES"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4800
            TabIndex        =   48
            Top             =   855
            Width           =   735
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0C0C0&
            Caption         =   "GESTI�N"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   47
            Top             =   855
            Width           =   735
         End
      End
      Begin VB.Frame FraNavega 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registros del EXTRACTO BANCARIO"
         ForeColor       =   &H00FF0000&
         Height          =   7320
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   12975
         Begin MSAdodcLib.Adodc Ado_datos 
            Height          =   330
            Left            =   120
            Top             =   6840
            Width           =   12705
            _ExtentX        =   22410
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
            BackColor       =   16777215
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
            Caption         =   " <-- Inicio                                                  Asistencia                                              Fin -->"
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
         Begin MSDataGridLib.DataGrid dg_datos 
            Bindings        =   "fw_conciliacion_bancaria.frx":9451
            Height          =   6495
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   12720
            _ExtentX        =   22437
            _ExtentY        =   11456
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
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
      End
      Begin VB.Frame Fra_ABM 
         BackColor       =   &H00C0C0C0&
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
         Height          =   7320
         Left            =   13200
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   5805
         Begin VB.OptionButton rbtDia 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Por d�a"
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
            Index           =   0
            Left            =   2880
            TabIndex        =   25
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton rbtMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Por mes"
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
            Left            =   1320
            TabIndex        =   24
            Top             =   720
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.ComboBox cmb_departamento 
            Height          =   315
            ItemData        =   "fw_conciliacion_bancaria.frx":9469
            Left            =   360
            List            =   "fw_conciliacion_bancaria.frx":9488
            TabIndex        =   23
            Top             =   2400
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cmb_equipo 
            Height          =   315
            ItemData        =   "fw_conciliacion_bancaria.frx":94DC
            Left            =   3240
            List            =   "fw_conciliacion_bancaria.frx":9501
            TabIndex        =   22
            Top             =   2400
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox LblMensaje 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   1320
            TabIndex        =   21
            Text            =   "IMPORTANDO DATOS ..."
            Top             =   6360
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.ComboBox cmb_mes_ini 
            DataField       =   "mes_inicio_crono"
            DataSource      =   "Ado_datos"
            Height          =   315
            ItemData        =   "fw_conciliacion_bancaria.frx":9547
            Left            =   360
            List            =   "fw_conciliacion_bancaria.frx":956F
            TabIndex        =   20
            Top             =   1800
            Width           =   2340
         End
         Begin VB.ComboBox cmb_gestion_rep 
            Height          =   315
            Left            =   2760
            TabIndex        =   19
            Top             =   1800
            Width           =   1095
         End
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   6840
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   661
            _Version        =   327682
            Appearance      =   1
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   285
            Left            =   360
            TabIndex        =   26
            Top             =   1800
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   42570
         End
         Begin VB.Label lbl_inicial 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Elija: Procesar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   31
            Top             =   1440
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label lbl_inicialr 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Elija ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   30
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lbl_inicialw 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar (Departamento)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   29
            Top             =   2160
            Visible         =   0   'False
            Width           =   2280
         End
         Begin VB.Label lbl_inicialq 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Equipo Biom�trico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   3360
            TabIndex        =   28
            Top             =   2160
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label LblTime 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Por Mes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   27
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   1440
            Picture         =   "fw_conciliacion_bancaria.frx":95D8
            Top             =   3360
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   1440
            Picture         =   "fw_conciliacion_bancaria.frx":98E2
            Top             =   4800
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   18960
         TabIndex        =   8
         Top             =   360
         Width           =   18960
         Begin VB.CommandButton BtnDesAprobar 
            BackColor       =   &H00808080&
            Height          =   600
            Left            =   8400
            Picture         =   "fw_conciliacion_bancaria.frx":9BEC
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.PictureBox BtnA�adir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_conciliacion_bancaria.frx":9DF6
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   14
            Top             =   0
            Width           =   1200
         End
         Begin VB.PictureBox BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1305
            Picture         =   "fw_conciliacion_bancaria.frx":A5B5
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnEliminar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            Picture         =   "fw_conciliacion_bancaria.frx":AECA
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   12
            ToolTipText     =   "Borrar el Registro de Facturas Migradas"
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox BtnAprobar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4080
            Picture         =   "fw_conciliacion_bancaria.frx":B616
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnImprimir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5520
            Picture         =   "fw_conciliacion_bancaria.frx":BE49
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.PictureBox BtnSalir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   17520
            Picture         =   "fw_conciliacion_bancaria.frx":C716
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   9
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.Label lbl_titulo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IMPORTAR EXTRACTO BANCARIO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   9825
            TabIndex        =   16
            Top             =   195
            Width           =   4035
         End
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
      ScaleWidth      =   14835
      TabIndex        =   0
      Top             =   8790
      Width           =   14835
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
         Left            =   0
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
   Begin Crystal.CrystalReport Cr01 
      Left            =   2400
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   2880
      Top             =   9480
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   5280
      Top             =   9480
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "ado_datos14"
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
      Left            =   0
      Top             =   10200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Crystal.CrystalReport Cr02 
      Left            =   2880
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "fw_conciliacion_bancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombreArchivo As String
Dim SiEstaImportado As Boolean
Dim Mensaje As String
Dim Fecha As Date

Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset

Dim rs_aux7 As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset

Dim queryinicial1 As String
Dim queryinicial2 As String

Dim Nro As String
Dim varbusca As String
Dim ac_no As Integer

Private Sub Ado_datos3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Ado_datos3.Recordset.RecordCount > 0 Then
    'DESTINO - CONCILIADOS SOFIA
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
            DtGLista.Visible = False
            rs_datos14.Open "select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'APR' and cmpbte_deposito_bco = '" & Ado_datos3.Recordset!cod_bancarizacion & "' AND fecha_registro_bco = '" & CDate(Ado_datos3.Recordset!fecha_transaccion) & "') ", db, adOpenKeyset, adLockOptimistic
            'queryinicial2 = " select * from fv_ventas_cobranza_det_VS_extracto_BMSC_2021 where (estado_conciliado_tes = 'APR') "
            'queryinicial2 = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'APR' and cmpbte_deposito_bco = '" & Ado_datos3.Recordset!cod_bancarizacion & "' AND fecha_registro_bco = '" & Ado_datos3.Recordset!fecha_transaccion & "') "
            'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
            'rs_datos14.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
        Set ado_datos14.Recordset = rs_datos14.DataSource
        'ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            DtGLista.Visible = True
            Set DtGLista.DataSource = ado_datos14.Recordset
            'Set dg_datos3.DataSource = ado_datos14.Recordset
        Else
            deta2 = 0
            DtGLista.Visible = False
        End If
    End If
End Sub


Private Sub BtnBuscar_Click()
  If Ado_datos11.Recordset.RecordCount > 0 Then
    buscados = 1
    'OptFilGral2.Visible = False
    'OptFilGral1.Visible = False
'    Call OptFilGral2_Click
'    Call ABRIR_DETALLE
    PosibleApliqueFiltro = False
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexi�n = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = DtGLista11
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = Ado_datos11.Recordset
    ClBuscaGrid.CamposVisibles = "110"
    ClBuscaGrid.Ejecutar
    PosibleApliqueFiltro = True

  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atenci�n!"
    'OptFilGral1.Visible = True
    'OptFilGral2.Visible = True
  End If
End Sub

Private Sub BtnBuscar1_Click()
  If Ado_datos2.Recordset.RecordCount > 0 Then
    buscados = 1
    'OptFilGral2.Visible = False
    'OptFilGral1.Visible = False
'    Call OptFilGral2_Click
'    Call ABRIR_DETALLE
    PosibleApliqueFiltro = False
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexi�n = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos2
    ClBuscaGrid.QueryUtilizado = queryinicial1
    Set ClBuscaGrid.RecordsetTrabajo = Ado_datos2.Recordset
    ClBuscaGrid.CamposVisibles = "110"
    ClBuscaGrid.Ejecutar
    PosibleApliqueFiltro = True

  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atenci�n!"
    'OptFilGral1.Visible = True
    'OptFilGral2.Visible = True
  End If

End Sub

Private Sub BtnBuscar2_Click()
  If Ado_datos3.Recordset.RecordCount > 0 Then
    buscados = 1
    'OptFilGral2.Visible = False
    'OptFilGral1.Visible = False
'    Call OptFilGral2_Click
'    Call ABRIR_DETALLE
    PosibleApliqueFiltro = False
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexi�n = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos3
    ClBuscaGrid.QueryUtilizado = queryinicial2
    Set ClBuscaGrid.RecordsetTrabajo = Ado_datos3.Recordset
    ClBuscaGrid.CamposVisibles = "110"
    ClBuscaGrid.Ejecutar
    PosibleApliqueFiltro = True

  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atenci�n!"
    'OptFilGral1.Visible = True
    'OptFilGral2.Visible = True
  End If

End Sub

Private Sub BtnBuscar3_Click()
    Call ABRIR_LAS4
    FraBusca4.Visible = True
End Sub

Private Sub BtnBuscar4_Click()
    'DESTINO - CONCILIADOS SOFIA
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
        DtGLista.Visible = False
        varbusca = "%" + txtBusca.Text + "%"
        queryinicial2 = " select * from FO_extracto_BMSC_202101 where ((estado_conciliado = 'APR') AND " & _
            " (id_depositante like '" & varbusca & "' OR nombre_depositante like '" & varbusca & "' OR descripcion like '" & varbusca & "' OR oficina like '" & varbusca & "' OR banco like '" & varbusca & "' OR nombre_destinatario like '" & varbusca & "' OR  glosa like '" & varbusca & "' OR " & _
            " originador  like '" & varbusca & "'  OR  originador_ACH like '" & varbusca & "' OR ciudad_origen like '" & varbusca & "' )) "
        
        'fecha_transaccion like '', hora_transaccion, cod_bancarizacion, nro_cheque, plantilla, cod_cliente,
        ', debito, credito, saldo, cuenta, estado_conciliado, correlativo
        rs_datos3.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        'rs_datos14.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos3.Recordset = rs_datos3.DataSource
    'ado_datos14.Recordset.Requery
    If Ado_datos3.Recordset.RecordCount > 0 Then
        deta2 = 1
        dg_datos3.Visible = True
        'Set DtGLista.DataSource = ado_datos14.Recordset
        Set dg_datos3.DataSource = Ado_datos3.Recordset
    Else
        deta2 = 0
        dg_datos3.Visible = False
    End If
End Sub

Private Sub BtnCancelar_Click()
    Frame1.Visible = False
End Sub

Private Sub BtnCancelar3_Click()
    Fra_reporte.Visible = False
End Sub

Private Sub BtnEliminar_Click()
    Frame1.Visible = True
End Sub

Private Sub BtnGrabar_Click()
    sino = MsgBox("�Est� Seguro de Eliminar el Registro de Facturas Electr�nicas ?", vbYesNo + vbQuestion, "Atenci�n")
    If sino = vbYes Then
        db.Execute " UPDATE ao_ventas_cobranza SET factura_impresa = 'N', cobranza_nro_factura = '0', estado_codigo_fac1 = 'REG', cta_codigo2 = 'NN', trans_codigo = 'O' WHERE (factura_impresa = 'S') AND (cobranza_nro_factura > '0') AND (ges_gestion = '2020') AND (estado_codigo_fac1 = 'APR') AND (cmpbte_deposito2 <> '0') AND (cta_codigo2 = '" & NombreArchivo & "') and (trans_codigo = 'L') "
        MsgBox "Se anularon los registros de Facturas Electr�nicas ..."
        Frame1.Visible = False
    End If
End Sub

Private Sub BtnImprimir1_Click()
    CR01.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_tesoreria_regional_detalle_concilia.rpt"
    titulo2 = "MODULO TESORERIA"
    subtitulo2 = "CONCILIACION P/CUENTA BANCARIA"
    CR01.Formulas(2) = "Titulo = '" & titulo2 & "'"
    CR01.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
    
    iResult = CR01.PrintReport
    If iResult <> 0 Then
        MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
End Sub

Private Sub BtnImprimir2_Click()
    CR02.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_tesoreria_regional_NO_conciliados.rpt"
    titulo2 = "EXTRACTOS BANCARIOS"
    subtitulo2 = "NO CONCILIADOS"
    CR02.Formulas(2) = "Titulo = '" & titulo2 & "'"
    CR02.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
    
    iResult = CR02.PrintReport
    If iResult <> 0 Then
        MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
End Sub

Private Sub BtnModDetalle_Click()
        If Ado_datos11.Recordset.RecordCount > 0 Then         '<> "" Then
'            'GRABA RECIBO DETALLE
            Text11.Text = IIf(IsNull(Ado_datos11.Recordset!cmpbte_deposito_bco), 0, Ado_datos11.Recordset!cmpbte_deposito_bco)
            DTP_Finicio.Value = IIf(IsNull(Ado_datos11.Recordset!fecha_registro_bco), Date, Ado_datos11.Recordset!fecha_registro_bco)
            
'            Label6.Caption = Ado_datos11.Recordset!trans_descripcion
            Fra_reporte.Visible = True
            'DtGLista.Enabled = True
        Else
            MsgBox "Debe elegir un registro cobrado para modificar, verifique y vuelva a intentar ...", , "Atenci�n"
        End If
End Sub

Private Sub cbo_mes_rep_Change()
    txt_mes.Text = cbo_mes_rep.ListIndex
    txt_mes.Text = Val(txt_mes.Text) + 1
End Sub

Private Sub cmb_mes_ini_Click()
    txt_mes.Text = cmb_mes_ini.ListIndex
    txt_mes.Text = Val(txt_mes.Text) + 1
End Sub

Private Sub CmdElim2_Click()
    Frame1.Visible = False
End Sub

Private Sub dtc_rep_cod_Click(Area As Integer)
    dtc_rep_det.BoundText = dtc_rep_cod.BoundText
    dtc_rep_det.BoundText = dtc_depto.BoundText
    Option1.Value = False
End Sub

Private Sub dtc_rep_det_Click(Area As Integer)
    dtc_rep_cod.BoundText = dtc_rep_det.BoundText
    dtc_depto.BoundText = dtc_rep_det.BoundText
    Option1.Value = False
End Sub

Private Sub Form_Load()
    Call CargarControles
    NombreArchivo = ""
    SiEstaImportado = False
    Call limpiar
    'Call sstab1_Click
    If SSTab1.Tab = 1 Then
        'ACTUALIZA LOS CONCILIADOS
        db.Execute "UPDATE fo_extracto_BMSC_202101 SET Cuenta = '4010620792' WHERE  (Cuenta IS NULL) "
        db.Execute "UPDATE fo_recibos_detalle SET estado_conciliado = 'REG' WHERE  (estado_conciliado IS NULL) "
        db.Execute "UPDATE fo_extracto_BMSC_202101 SET estado_conciliado = 'REG' WHERE  (estado_conciliado IS NULL) "
        
        db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.edif_codigo_corto = av_venta_cobranza_APR.edif_codigo_corto FROM fo_recibos_detalle inner JOIN av_venta_cobranza_APR ON fo_recibos_detalle.correl_cobro = av_venta_cobranza_APR.correl_cobro WHERE fo_recibos_detalle.edif_codigo_corto IS NULL"
    
        db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.estado_conciliado = 'APR' FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion "
        db.Execute "UPDATE fo_extracto_BMSC_202101 SET fo_extracto_BMSC_202101.estado_conciliado = 'APR' FROM fo_extracto_BMSC_202101 INNER JOIN fo_recibos_detalle ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion "
        
        '-- 1. TODOS = (#BANCARIZACION, CUENTA, FECHA, MONTO, CLIENTE)
        db.Execute "UPDATE fo_recibos_detalle SET  fo_recibos_detalle.nivel_conciliado = 1 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  = fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  = fo_extracto_BMSC_202101.cod_cliente WHERE fo_recibos_detalle.estado_conciliado = 'APR' "
        
        '-- 2. = (#BANCARIZACION, CUENTA, FECHA, MONTO ) Y <> (CLIENTE)
        db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.nivel_conciliado = 2 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  = fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  <> fo_extracto_BMSC_202101.cod_cliente WHERE fo_recibos_detalle.estado_conciliado = 'APR' AND  fo_recibos_detalle.nivel_conciliado <> 1 "
    
        '-- 3. = (#BANCARIZACION, CUENTA, FECHA) Y <> (CLIENTE, MONTO)
        db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.nivel_conciliado = 3 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  <> fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  <> fo_extracto_BMSC_202101.cod_cliente WHERE fo_recibos_detalle.estado_conciliado = 'APR' AND  fo_recibos_detalle.nivel_conciliado <> 1 AND  fo_recibos_detalle.nivel_conciliado <> 2 "
    
        '-- 4. = (#BANCARIZACION, CUENTA) Y <> (CLIENTE, MONTO, FECHA)
        'db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.estado_conciliado = 'APR', nivel_conciliado = 4 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  <> fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  <> fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  <> fo_extracto_BMSC_202101.cod_cliente "
        db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.nivel_conciliado = 4 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion WHERE fo_recibos_detalle.estado_conciliado = 'APR' AND  fo_recibos_detalle.nivel_conciliado <> 1 AND  fo_recibos_detalle.nivel_conciliado <> 2 AND  fo_recibos_detalle.nivel_conciliado <> 3 "
    
        Call ABRIR_LAS4
    End If
'    Set rs_aux7 = New ADODB.Recordset
'    If rs_aux7.State = 1 Then rs_aux7.Close
'    rs_aux7.Open "SELECT * FROM rc_planilla_grupo", db, adOpenStatic
'    Set Ado_datos_rep.Recordset = rs_aux7
'    dtc_rep_det.BoundText = dtc_rep_cod.BoundText
End Sub

Private Sub BtnA�adir_Click()
'    Call limpiar
'    LblMensaje.Visible = False
'    Fra_ABM.Enabled = True
'    BtnA�adir.Visible = True
'    btnCargarArchivo.Visible = True
'    Image1.Visible = True
'    btnImportarDato.Visible = False
'    Image2.Visible = False
'    cmb_gestion_rep.Text = Year(Date)

    If glusuario = "VPAREDES" Or glusuario = "ADMIN" Or glusuario = "MWILDE" Or glusuario = "RCUELA" Or glusuario = "MVALDIVIA" Or glusuario = "CSALINAS" Then         'Or glusuario = "MVALDIVIA"
        Dim e As Long
        e = Shell(App.Path & "\Extractos\SofiaNetCore.exe", 1)
    Else
        MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
    End If
    ' esto es...
    '-- ACTUALIZA INGRESOS fo_extracto_ingreso_GRAL
    'db.Execute "DELETE fo_extracto_ingreso_GRAL "
    'db.Execute "INSERT INTO fo_extracto_ingreso_GRAL (correlativo, cuenta, fecha_transaccion, monto, cod_bancarizacion, agencia, descripcion, glosa, nro_cheque, plantilla, estado_conciliado, usuario, nombre_archivo, cod_cliente, id_depositante, nombre_depositante, banco) SELECT * FROM fv_extracto_ingreso_GRAL "

    '-- ACTUALIZA BOLIVIANOS
    'db.Execute "UPDATE fo_extracto_ingreso_GRAL SET monto_bs = monto, monto_dol = MONTO/6.96 WHERE cuenta ='2015046557-03-054' OR cuenta ='4010439742' OR cuenta ='4010620792' OR cuenta ='4010644195' OR cuenta ='4010772049' OR cuenta ='4011005599' OR cuenta ='4011048967' OR cuenta ='4011048981' OR cuenta ='4069626219' OR cuenta ='4069626233' OR cuenta ='10000019133060' "

    '-- ACTUALIZA DOLARES
    'db.Execute "UPDATE fo_extracto_ingreso_GRAL SET monto_dol = monto, monto_bs = MONTO*6.96 WHERE cuenta ='201-5041743-2-18' OR cuenta ='096359-201-9' OR cuenta ='4010038393' OR cuenta ='4010620785' OR cuenta ='4010780124' OR cuenta ='4011005601' OR cuenta ='4011048974' OR cuenta ='4069626242' OR cuenta ='4069626265' "

    '-- ACTUALIZA EGREOS fo_extracto_egreso_GRAL        'select * from fo_extracto_egreso_GRAL
    'db.Execute "Delete fo_extracto_egreso_GRAL "
    'db.Execute "INSERT INTO fo_extracto_egreso_GRAL (correlativo, cuenta, fecha_transaccion, monto, cod_bancarizacion, agencia, descripcion, glosa, nro_cheque, plantilla, estado_conciliado, usuario, nombre_archivo, cod_cliente, id_depositante, nombre_depositante, banco) SELECT * FROM fv_extracto_egreso_GRAL "

    '-- ACTUALIZA BOLIVIANOS
    'db.Execute "UPDATE fo_extracto_egreso_GRAL SET monto_bs = monto, monto_dol = MONTO/6.96 WHERE cuenta ='2015046557-03-054' OR cuenta ='4010439742' OR cuenta ='4010620792' OR cuenta ='4010644195' OR cuenta ='4010772049' OR cuenta ='4011005599' OR cuenta ='4011048967' OR cuenta ='4011048981' OR cuenta ='4069626219' OR cuenta ='4069626233' OR cuenta ='10000019133060' "

    '-- ACTUALIZA DOLARES
    'db.Execute "UPDATE fo_extracto_egreso_GRAL SET monto_dol = monto, monto_bs = MONTO*6.96 WHERE cuenta ='201-5041743-2-18' OR cuenta ='096359-201-9' OR cuenta ='4010038393' OR cuenta ='4010620785' OR cuenta ='4010780124' OR cuenta ='4011005601' OR cuenta ='4011048974' OR cuenta ='4069626242' OR cuenta ='4069626265' "

End Sub

Private Sub limpiar()
    Mensaje = ""
    SiEstaImportado = False
    
    btnCargarArchivo.Enabled = True
    btnImportarDato.Enabled = True
    cmb_departamento = ""
    cmb_equipo = ""
    DtpFecha.Value = Date
    
End Sub

Private Sub btnCargarArchivo_Click()
  Dim rutaArchivo As String
  rutaArchivo = App.Path & "\EXTRACTOS\"
  LblMensaje.Visible = False
  Dim existeRuta As Boolean
  Dim oDir As New Scripting.FileSystemObject
  existeRuta = oDir.FolderExists(rutaArchivo)
   
  ' Valida si existe ruta destino.
  If existeRuta = Falso Then
     ' Consulta no existe ruta.
     sino = MsgBox("No existe ruta destino 'EXTRACTOS' � Desea crearla ? ", vbYesNo + vbQuestion, "Atenci�n")
     If sino = vbYes Then
           Dim f As FileSystemObject
           Set f = New FileSystemObject
           f.CreateFolder (rutaArchivo)
           existeRuta = True
     End If
  End If
   
  If existeRuta Then
     ' Carga archivo.
     Dim rsCantExistente As New ADODB.Recordset
     Dim esValido As Boolean
     esValio = True
     Call valida_campos(esValio)
    
     If esValio Then
        If rbtMes.Value = True Then
            'sino = MsgBox("�Esta seguro de subir la asistencia del MES con los siguientes datos?" & vbCrLf & "Gestion: " & cmb_gestion_rep.Text & vbCrLf & "Mes:" & cmb_mes_ini.Text & vbCrLf & "Equipo Biom�trico: " & cmb_equipo.Text & vbCrLf & "Departamento: " & cmb_departamento.Text, vbYesNo + vbQuestion, "Atenci�n")
            sino = MsgBox("�Esta seguro de subir el Extracto con los siguientes datos?" & vbCrLf & "Gestion: " & cmb_gestion_rep.Text & vbCrLf & "Mes:" & cmb_mes_ini.Text, vbYesNo + vbQuestion, "Atenci�n")
        End If
        If rbtDia(0).Value = True Then
            sino = MsgBox("�Esta seguro de subir el Extracto con los siguientes datos?" & vbCrLf & "Fecha:" & DtpFecha.Value & vbCrLf & "Equipo Biom�trico: " & cmb_equipo.Text & vbCrLf & "Departamento: " & cmb_departamento.Text, vbYesNo + vbQuestion, "Atenci�n")
        End If
        If sino = vbYes Then
            GLCarpeta = ""
            BtnA�adir.Visible = False
            Fra_ABM.Enabled = False
            Dim dia As String, mes As String
        
            Fecha = DtpFecha.Value
            Call ObtenerDiaMes(DatePart("m", Fecha), mes)
            ' Tipo de exportaci�n por mes o dia.
            If rbtMes.Value = True Then
                NombreArchivo = UCase(Trim$("EB" & "_" & cmb_gestion_rep.Text & cmb_mes_ini.Text))
'                If cmb_mes_ini.Text = "NO ASIGNADO" Then
'                    NombreArchivo = UCase(Trim$("EB" & "_" & cmb_gestion_rep.Text & txt_mes.Text))
'                Else
'                    NombreArchivo = UCase(Trim$("EB" & "_" & Replace(cmb_departamento, " ", "")) & "_" & cmb_gestion_rep.Text & txt_mes.Text)
'                End If
            Else
                Call ObtenerDiaMes(DatePart("d", Fecha), dia)
                NombreArchivo = UCase(Trim$("EB" & "_" & DatePart("yyyy", Fecha) & mes & dia))
                'NombreArchivo = UCase(Trim$(Replace(cmb_departamento, " ", "")) & "_" & Trim$(cmb_equipo) & "_" & DatePart("yyyy", Fecha) & mes & dia)
            End If
            ' Asigna nombre archivo a variable global
            GLCarpeta2 = NombreArchivo
            rutaArchivo = App.Path & "\EXTRACTOS\"
            GlArch = "EXBCO"
            Frmexporta.DirDestino.Path = rutaArchivo
            Frmexporta.DirDestino2.Path = rutaArchivo
            Frmexporta.Show vbModal
            ' Verifica si nombre de hoja es diferente a vacio.
            If GLCarpeta2 <> "" Then
                MsgBox "El archivo " & NombreArchivo & " se copio correctamente."
                btnImportarDato.Enabled = True
            End If
        
            ' Consulta verifica si los datos del archivo con NombreArchivo se registraron.
'            'rsCantExistente.Open "SELECT COUNT(*) AS 'cuantos' FROM auxiliar_asistencia AS ax INNER JOIN ro_controlasistencia AS ctr ON ax.Id_AuxAsis =ctr.Id_AuxAsis WHERE ax.Nombre_Archivo = '" & NombreArchivo & "' ", db, adOpenStatic
'            rsCantExistente.Open "SELECT COUNT(*) AS cuantos FROM fo_auxiliar_facturacion AS ax INNER JOIN ao_ventas_cobranza AS ctr ON ax.cobranza_codigo = ctr.cobranza_codigo WHERE ax.Nombre_Archivo = '" & NombreArchivo & "' ", db, adOpenStatic
'            rsCantExistente.MoveFirst
'
'            If rsCantExistente![Cuantos] > 0 Then SiEstaImportado = True Else SiEstaImportado = False
'            rsCantExistente.Close
'
'            db.Execute "delete fo_auxiliar_facturacion "
            Set dg_datos.DataSource = rsNada
            btnCargarArchivo.Visible = False
            Image1.Visible = False
            btnImportarDato.Visible = True
            Image2.Visible = True
        End If
     End If
  End If
End Sub

Private Sub CargarControles()
    Dim rsDepartamento As New ADODB.Recordset
    Dim rsEquipo As New ADODB.Recordset
    rsDepartamento.Open "SELECT DISTINCT * FROM gc_departamento ", db, adOpenStatic
    rsDepartamento.MoveFirst
    With Me.cmb_departamento
        .Clear
        Do
            .AddItem rsDepartamento![depto_descripcion]
            rsDepartamento.MoveNext
        Loop Until rsDepartamento.EOF
    End With
    ' Equipo
    rsEquipo.Open "SELECT * FROM rc_equipo_asistencia ", db, adOpenStatic
    rsEquipo.MoveFirst
    With Me.cmb_equipo
        .Clear
        Do
            .AddItem rsEquipo![descripcion_asist]
            rsEquipo.MoveNext
        Loop Until rsEquipo.EOF
    End With
    
'UserForm_Initialize_Exit:
    On Error Resume Next
    rsDepartamento.Close
    rsEquipo.Close
End Sub


Private Sub valida_campos(esValio)
  Dim inicial As Integer
  If rbtDia(0).Value = True Then
    If DtpFecha.Value = "" Then
      MsgBox " El campo Fecha es requerido."
      esValio = False
    End If
  End If
  
  If rbtMes.Value = True Then
     If txt_mes.Text = "0" Or txt_mes.Text = "" Then
        MsgBox " El campo Mes requerido."
        esValio = False
     End If
  End If
  
'  If cmb_departamento = "" Then
'    MsgBox " Seleccione un departamento."
'    esValio = False
'  End If
  
'  If cmb_equipo = "" Then
'    MsgBox " Seleccione un equipo."
'    esValio = False
'  End If

End Sub


Private Sub btnImportarDato_Click()
        btnCargarArchivo.Visible = False
        Image1.Visible = False
        
        If SiEstaImportado Then
            sino = MsgBox("�Existen datos para '" & NombreArchivo & "',desea reemplazarlos?", vbQuestion + vbYesNo, "Confirmando ... ")
            If sino = vbYes Then
                Call EliminarDatoAnterior
                MsgBox "Los datos anteriores se anularon ..."
                'Call ImportarDato
                btnImportarDato.Enabled = False
'                db.Execute "UPDATE fo_auxiliar_facturacion SET Notas = '0' WHERE NOTAS IS NULL "
'                db.Execute "UPDATE fo_auxiliar_facturacion SET nro_factura = substring(IdFactura,1,CHARINDEX('-', idfactura,1)-1) "
'                db.Execute "UPDATE fo_auxiliar_facturacion SET cobranza_codigo = cast(Notas as integer)"
'                db.Execute "UPDATE fo_auxiliar_facturacion SET edif_codigo_corto = CampoProducto"
'
'                db.Execute "UPDATE ao_ventas_cobranza SET ao_ventas_cobranza.cobranza_nro_factura = fo_auxiliar_facturacion.nro_factura, " & _
'                " ao_ventas_cobranza.factura_impresa = 'S', ao_ventas_cobranza.estado_codigo_fac1 = 'APR', ao_ventas_cobranza.estado_codigo_fac = 'APR',  " & _
'                " ao_ventas_cobranza.cta_codigo2 = fo_auxiliar_facturacion.Nombre_Archivo , " & _
'                " ao_ventas_cobranza.trans_codigo = 'L', ao_ventas_cobranza.cmpbte_deposito2 = fo_auxiliar_facturacion.Factura, " & _
'                " ao_ventas_cobranza.cobranza_fecha_fac = fo_auxiliar_facturacion.Fecha_emision " & _
'                " FROM ao_ventas_cobranza INNER JOIN fo_auxiliar_facturacion " & _
'                " ON (ao_ventas_cobranza.cobranza_codigo = fo_auxiliar_facturacion.cobranza_codigo ) "
            End If
        Else
           'Call ImportarDato           ' HABILITAR
           'Call ABRIR_TABLA             ' DESHABILITAR
           'db.Execute "UPDATE ro_controlasistencia SET ges_gestion = year(Fecha_control), Mes_control = month(Fecha_control), Dia_control= day(Fecha_control)"
           btnImportarDato.Enabled = False
           ' Call ABRIR_TABLA
            
        End If
        Fra_ABM.Enabled = False
        ProgressBar1.Visible = False
        Image2.Visible = False
End Sub

' Eliminar datos de anterior importacion
Private Sub EliminarDatoAnterior()
     ' FALTA ---- estado_codigo_fac1 = "FIN"  ---- trans_codigo = "L"  (Tipo de Transaccion L=Fact.Electr.)
     db.Execute " UPDATE ao_ventas_cobranza SET factura_impresa = 'N', cobranza_nro_factura = '0', estado_codigo_fac1 = 'REG', cta_codigo2 = 'NN', trans_codigo = 'O' WHERE (factura_impresa = 'S') AND (cobranza_nro_factura > '0') AND (ges_gestion = '" & cmb_gestion_rep.Text & "') AND (estado_codigo_fac1 = 'APR') AND (cmpbte_deposito2 <> '0') AND (cta_codigo2 = '" & NombreArchivo & "') and (trans_codigo = 'L') "
     db.Execute " DELETE FROM fo_auxiliar_facturacion WHERE Nombre_Archivo = '" & NombreArchivo & "' "
End Sub

' Importar excel
Private Sub ImportarDato()
  On Error GoTo ErrorHandler
            
        LblMensaje.Visible = True
        MsgBox " Se inicia el proceso de importaci�n de datos..."
                
        Dim conExcel As New ADODB.Connection
        Dim rsExcel As New ADODB.Recordset
        
        Dim rsTablaAuxiliar As ADODB.Recordset
        
        Dim sqlDatosAux As String
        Dim indice As Integer
        
        If conExcel.State = adStateOpen Then conExcel.Close
        If rsExcel.State = adStateOpen Then rsExcel.Close
        
        Dim origenExcel As String
        Dim ruta As String
        origenExcel = NombreArchivo '
        ' ruta = App.Path & "\EXTRACTOS\" & NombreArchivo & ".xls"
        ruta = App.Path & "\EXTRACTOS\" & NombreArchivo & "." & GlExtension
        
        '--------------------------------- Obtiene nombre de hoja
'        Dim ObjExcel As Excel.Application
'        Dim ObjExcelLibro As Excel.Workbook
'        Set ObjExcel = New Excel.Application
'        Set ObjExcelLibro = ObjExcel.Workbooks.Open(ruta)
'
'        If ObjExcelLibro.Sheets.Count > 0 Then
'            ' Asigna nombre de primera hoja.
'            GLCarpeta = ObjExcelLibro.Sheets(1).Name
'        End If
'        ObjExcelLibro.Close
'        ObjExcel.Quit
'        Set ObjExcelLibro = Nothing
'        Set ObjExcel = Nothing
       
        '---------------------------------
        
        ' Coneccion a excel
'        conExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'            "Data Source= " & ruta & ";" & _
'                "Extended Properties=""Excel 8.0;"";"
                
         conExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source= " & ruta & "; Extended Properties=""Excel 12.0 Xml;"";"
        
        ' Consulta obtiene datos de excel.
        ' GLCarpeta contiene nombre de hoja desde frmexport
        If GLCarpeta2 <> "Worksheet" And GLCarpeta2 <> "Hoja1" Then
            rsExcel.Open "SELECT * FROM [" & GLCarpeta2 & "$]", conExcel, 3, 1
        Else
            rsExcel.Open "SELECT * FROM [Hoja1$] ", conExcel, 3, 1
        End If
        ' INSERTA REGISTROS A TABLA AUXILIAR
         indice = 0
        ' Variables de registros auxiliar
        Dim nroxl As Integer, cantRegistro As Integer
        Dim sql As String
        Dim sqlValue As String
        cantRegistro = 1
        'JQ
        CANTOT = rsExcel.RecordCount
        ProgressBar1.Visible = True
        With ProgressBar1
            .Max = CANTOT     'rs_datos6.RecordCount
            .Min = 0
            .Value = 0
        End With
        While Not rsExcel.EOF
                
            If rsExcel.Fields(0) <> "" Or rsExcel.Fields(0) <> Nulo Then
                For indice = 0 To rsExcel.Fields.Count - 1
                    sqlDatosAux = sqlDatosAux & "'" & rsExcel.Fields(indice).Value & "',"
                Next
            End If
            
            If sqlValue = "" And Trim$(sqlDatosAux) <> "" Then
                 sqlValue = " (" & Mid(sqlDatosAux, 1, Len(sqlDatosAux) - 4) & " ,'" & GLCarpeta2 & "' )"
            Else
                 If Trim$(sqlDatosAux) <> "" Then
                       sqlValue = sqlValue & ", (" & Mid(sqlDatosAux, 1, Len(sqlDatosAux) - 4) & " ,'" & GLCarpeta2 & "' )"
                 End If
            End If
            ' Sql server solo permite registrar 1000 registros por insert.
            'If cantRegistro = 1000 Then
            'If cantRegistro = 100 Then   ' Factura, Fecha_emision, IdFactura, NombreCliente, GlosaSevicio, TipoFactura, EstadoPago, FechaVencimiento, Moneda, TipoCambio, Cantidadtotal, Notas, CampoProducto, CampoCliente, nro_factura, edif_codigo_corto, FechaFactura, cobranza_codigo, Nombre_Archivo
                 sql = sql & " INSERT INTO fo_auxiliar_facturacion (Factura, Fecha_emision, IdFactura, NombreCliente, GlosaSevicio, TipoFactura, EstadoPago, FechaVencimiento, Moneda, TipoCambio, Cantidadtotal, Notas, CampoProducto, CampoCliente, nro_factura, edif_codigo_corto, FechaFactura, cobranza_codigo, Nombre_Archivo) VALUES  " & sqlValue & " ;"
                 'fo_auxiliar_facturacion  Factura, Fecha_emision, IdFactura, NombreCliente, GlosaSevicio, TipoFactura, EstadoPago, FechaVencimiento, Moneda, TipoCambio, Cantidadtotal, Notas, CampoProducto , CampoCliente, Nombre_Archivo
                 'db.Execute " UPDATE fo_auxiliar_facturacion SET cobranza_codigo = Notas WHERE (Notas = " & ac_no & ")  "
                 cantRegistro = 0
                 sqlValue = ""
                  ' Inserta registros.
                 'db.Execute sql
                 'sql = ""
            'End If
                
            sqlDatosAux = ""
            rsExcel.MoveNext
            cantRegistro = cantRegistro + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
        Wend
        
        If sqlValue <> "" Then
            'sql = sql & " INSERT INTO fo_auxiliar_facturacion (Factura, Fecha_emision, IdFactura, NombreCliente, GlosaSevicio, TipoFactura, EstadoPago, FechaVencimiento, Moneda, TipoCambio, Cantidadtotal, Notas, CampoProducto, CampoCliente, nro_factura, edif_codigo_corto, FechaFactura, cobranza_codigo, Nombre_Archivo) VALUES  " & sqlValue & " ;"
            sql = sql & " INSERT INTO fo_auxiliar_facturacion (Factura, Fecha_emision, IdFactura, NombreCliente, GlosaSevicio, TipoFactura, EstadoPago, FechaVencimiento, Moneda, TipoCambio, Cantidadtotal, Notas,CampoProducto , CampoCliente, nro_factura, edif_codigo_corto, FechaFactura, cobranza_codigo, Nombre_Archivo) VALUES  " & sqlValue & " ;"
        End If
        If sql <> "" Then
             ' Inserta registros.
            db.Execute sql
            sql = ""
            ' Actualiza Datos fo_auxiliar_facturacion
            db.Execute "UPDATE fo_auxiliar_facturacion SET Notas = '0' WHERE NOTAS IS NULL "
            'db.Execute "UPDATE fo_auxiliar_facturacion SET nro_factura = substring(IdFactura,1,CHARINDEX('-', idfactura,1)-1) "
            db.Execute "UPDATE fo_auxiliar_facturacion SET cobranza_codigo = cast(Notas as integer)"
            'db.Execute "UPDATE fo_auxiliar_facturacion SET edif_codigo_corto = CampoProducto"
            
            db.Execute "UPDATE ao_ventas_cobranza SET ao_ventas_cobranza.cobranza_nro_factura = fo_auxiliar_facturacion.nro_factura, " & _
            " ao_ventas_cobranza.factura_impresa = 'S', ao_ventas_cobranza.estado_codigo_fac1 = 'APR', ao_ventas_cobranza.estado_codigo_fac = 'APR',  " & _
            " ao_ventas_cobranza.cta_codigo2 = fo_auxiliar_facturacion.Nombre_Archivo , " & _
            " ao_ventas_cobranza.trans_codigo = 'L', ao_ventas_cobranza.cmpbte_deposito2 = fo_auxiliar_facturacion.Factura, " & _
            " ao_ventas_cobranza.cobranza_fecha_fac = fo_auxiliar_facturacion.FechaFactura " & _
            " FROM ao_ventas_cobranza INNER JOIN fo_auxiliar_facturacion " & _
            " ON (ao_ventas_cobranza.cobranza_codigo = fo_auxiliar_facturacion.cobranza_codigo ) "
        End If
        
        ' INSERTA REGISTROS A TABLA OFICIAL
        Set rsTablaAuxiliar = New ADODB.Recordset
         
        If rsTablaAuxiliar.State = 1 Then rsTablaAuxiliar.Close
        Dim sqlSelect As String
        ' Tipo de exportaci�n por mes o dia.
        If rbtMes.Value = True Then
            ' Consulta por mes
            'sqlSelect = "SELECT * FROM auxiliar_asistencia WHERE MONTH(Fecha) = '" & txt_mes.Text & "' AND YEAR(Fecha) = '" & cmb_gestion_rep.Text & "' AND Nombre_Archivo = '" & NombreArchivo & "' "
            sqlSelect = "SELECT * FROM fo_auxiliar_facturacion WHERE MONTH(Fecha_emision) = '" & txt_mes.Text & "' AND YEAR(Fecha_emision) = '" & cmb_gestion_rep.Text & "' AND Nombre_Archivo = '" & NombreArchivo & "' "
            Else
               ' Consulta por dia
                sqlSelect = "SELECT * FROM fo_auxiliar_facturacion WHERE DAY(Fecha_emision) = DAY('" & Fecha_emision & "') AND MONTH(Fecha_emision) = MONTH('" & Fecha_emision & "') AND YEAR(Fecha_emision) = YEAR('" & Fecha_emision & "') AND Nombre_Archivo = '" & NombreArchivo & "' "
            End If
 
            rsTablaAuxiliar.Open sqlSelect, db, 3, 1
          
           sqlValue = ""
           cantRegistro = 1
           sql = ""
           ' Recorre registros de auxiliar asistencia
           Dim strValorInser As String
           Dim esdebein As String, esfalta As String, esdebesal As String
           
           Dim tardanzaval As String
           Dim normal As String, tiemporeal As String, nday As String, ndiasot As String, tardanza As String
           Dim minutoTardanza As Integer
           Dim Formato As String
           Formato = "#,##0"
           
           If rsTablaAuxiliar.RecordCount > 0 Then
              rsTablaAuxiliar.MoveFirst
              While Not rsTablaAuxiliar.EOF
                'Factura, Fecha_emision, IdFactura, NombreCliente, GlosaSevicio, TipoFactura, EstadoPago, FechaVencimiento, Moneda, TipoCambio, Cantidadtotal, Notas, CampoProducto , CampoCliente, Nombre_Archivo
'                Call ObtenerValorNumero(rsTablaAuxiliar!TipoCambio, TipoCambio)
'                Call ObtenerValorNumero(rsTablaAuxiliar!Cantidadtotal, Cantidadtotal)
'                Call ObtenerValorNumero(rsTablaAuxiliar!Notas, Notas)
'                   Call ObtenerValorNumero(rsTablaAuxiliar!TiemReal, tiemporeal)
'                   Call ObtenerValorBool(rsTablaAuxiliar!Falta, esfalta)
'                   Call ObtenerValorBool(rsTablaAuxiliar!Debe_C_In, esdebein)
'                   Call ObtenerValorBool(rsTablaAuxiliar!Debe_C_Sal, esdebesal)
'                   Call ObtenerValorNumero(rsTablaAuxiliar!NDays, nday)
'                   Call ObtenerValorNumero(rsTablaAuxiliar!ndiasot, ndiasot)
                   
                   'tardanzaval = rsTablaAuxiliar!tardanza
                   
'                If rsTablaAuxiliar!tardanza = "NULL" Then
'                    tardanzaval = "00:00"
'                End If
'                If Trim(rsTablaAuxiliar!tardanza) = "" Then
'                    tardanzaval = "00:00"
'                End If
'
'                minutoTardanza = Format(DateDiff("n", "00:00", tardanzaval), Formato)
'
'                   Dim tardanzaCadena As String
'                   tardanzaCadena = rsTablaAuxiliar!tardanza
'                   If tardanzaCadena = "" Then
'                    tardanzaCadena = "0000"
'                   Else
'                    tardanzaCadena = Replace(rsTablaAuxiliar!tardanza, ":", "")
'                   End If
                
                'JQA    AQUI ---------------------------------------------------------
                ' Cadena de datos para insert.
'                strValorInser = " " & Nro & ", " & ac_no & ", '" & rsTablaAuxiliar!Cedula_No & "', " & _
'                                " '" & rsTablaAuxiliar!Nombre & "', '" & CStr(rsTablaAuxiliar!Auto_asigna) & "', '" & CStr(rsTablaAuxiliar!Fecha) & "', " & _
'                                " '" & CStr(rsTablaAuxiliar!Horario) & "', '" & Replace(rsTablaAuxiliar!HoraEnt, ":", "") & "', '" & CStr(rsTablaAuxiliar!HoraEnt) & "', " & _
'                                " '" & Replace(rsTablaAuxiliar!horaSal, ":", "") & "', '" & CStr(rsTablaAuxiliar!horaSal) & "', '" & Replace(rsTablaAuxiliar!Marc_Ent, ":", "") & "', " & _
'                                " '" & CStr(rsTablaAuxiliar!Marc_Ent) & "', '" & Replace(rsTablaAuxiliar!Marc_Sal, ":", "") & "', '" & CStr(rsTablaAuxiliar!Marc_Sal) & "', " & _
'                                 " " & Replace(normal, ",", ".") & ", " & Replace(tiemporeal, ",", ".") & ", '" & tardanzaval & "', " & _
'                                 " '" & CStr(rsTablaAuxiliar!SalioTempr) & "', " & esfalta & ", '" & Trim$(Replace(Replace(CStr(rsTablaAuxiliar!HoraExtra), "a.m.", ""), "p.m.", "")) & "', " & _
'                                 " '" & CStr(rsTablaAuxiliar!WorkTime) & "', '" & CStr(rsTablaAuxiliar!Excepcion) & "', " & esdebein & ", " & _
'                                 "  " & esdebesal & ", '" & CStr(rsTablaAuxiliar!Depto) & "', " & Replace(nday, ",", ".") & ", " & _
'                                 " '" & CStr(rsTablaAuxiliar!FinSemana) & "', '" & CStr(rsTablaAuxiliar!Feriado) & "', '" & CStr(rsTablaAuxiliar!TiemAsist) & "', " & _
'                                 "  " & Replace(ndiasot, ",", ".") & ", '" & CStr(rsTablaAuxiliar!FinSemanaOT) & "', '" & CStr(rsTablaAuxiliar!FeriadoOT) & "', " & rsTablaAuxiliar!Id_AuxAsis & " , " & _
'                                 " '" & tardanzaCadena & "', '" & Replace(rsTablaAuxiliar!TiemAsist, ":", "") & "' " & " , " & _
'                                 " " & minutoTardanza & " "
                
'                If Nro <> "NULL" Then
'                    If sqlValue = "" Then
'                         sqlValue = " (" & strValorInser & ")"
'                    Else
'                         sqlValue = sqlValue & ", (" & strValorInser & ") "
'                    End If
'                End If
                 ' Sql server solo permite registrar 1000 registros por insert.
'                If cantRegistro = 100 Then
'                    'UPDATE AO_VENTAS_COBRANZA
'                     sql = sql & " INSERT INTO ro_controlasistencia (Correl,Correl_ac,beneficiario_codigo,Nombre,Autoasigna,Fecha_control,TipoHorario,Hora1, HoraUno,Hora2,HoraDos,Hora3,HoraTres,Hora4,HoraCuatro,Normal,TiemReal,Tardanza,SalioTempr,EsFalta,HoraExtra,WorkTime,Excepcion,Debe_C_In,Debe_C_Sal,Depto,NDays,FinSemana,Feriado,TiemAsist,NDiasOT,FinSemanaOT,FeriadoOT, Id_AuxAsis,TardanzaCadena,TiempoTrabajoCadena, AtrasoMin1) VALUES  " & sqlValue & " ;"
'                     cantRegistro = 0
'                     sqlValue = ""
'                End If
                Dim posicion As Integer
                
                cantRegistro = 0
                'db.Execute " UPDATE fo_auxiliar_facturacion SET notas = '0' WHERE (notas is null) "
                posicion = InStr(rsTablaAuxiliar!IdFactura, "-")
                Nro = Left(rsTablaAuxiliar!IdFactura, posicion - 1)
                ac_no = CDbl(rsTablaAuxiliar!Notas)
                If ac_no <> 0 Then
                    'db.Execute " UPDATE ao_ventas_cobranza SET factura_impresa = 'S', cobranza_nro_factura = '" & Nro & "', estado_codigo_fac1 = 'APR', cta_codigo2 = '" & NombreArchivo & "', trans_codigo = 'L', cmpbte_deposito2 = '" & rsTablaAuxiliar!Factura & "', cobranza_observiones = '" & rsTablaAuxiliar!GlosaServicio & "' WHERE (cobranza_codigo = " & ac_no & ") AND (estado_codigo_fac1 = 'REG') "
                    
                End If
                rsTablaAuxiliar.MoveNext
                cantRegistro = cantRegistro + 1
              Wend
             
'              If sqlValue <> "" Then
'                    sql = sql & " INSERT INTO ro_controlasistencia (Correl,Correl_ac,beneficiario_codigo,Nombre,Autoasigna,Fecha_control,TipoHorario,Hora1, HoraUno,Hora2,HoraDos,Hora3,HoraTres,Hora4,HoraCuatro,Normal,TiemReal,Tardanza,SalioTempr,EsFalta,HoraExtra,WorkTime,Excepcion,Debe_C_In,Debe_C_Sal,Depto,NDays,FinSemana,Feriado,TiemAsist,NDiasOT,FinSemanaOT,FeriadoOT, Id_AuxAsis,TardanzaCadena,TiempoTrabajoCadena, AtrasoMin1) VALUES " & sqlValue & " ;"
'              End If
'
'              If sql <> "" Then
'                     ' Inserta registros.
'                    db.Execute sql
'              End If
              LblMensaje.Visible = False
              MsgBox "Los datos de las Facturas se registraron correctamente."
             
           Else
              'MsgBox " No existen datos coincidentes."
           End If
           
           Call ABRIR_TABLA
           ' VERIFICA REGISTROS
ErrorHandler:
    If Trim(Err.Description) <> "" Then
       LblMensaje.Visible = False
       MsgBox Err.Description, , "Error"
       Fra_ABM.Enabled = True
       BtnA�adir.Visible = True
    End If
End Sub

' Validacion cabecera.
Private Function ValidarCabecera(registros As ADODB.Recordset) As String ' Notice the As String
            Dim Mensaje As String
            Mensaje = ""
            
            Dim nombreEnc As String
            Dim nomCabecera As String
            For i = 0 To rsExcel.Fields.Count - 1
              nomCabecera = rsExcel.Fields(i).Name
              If i = 0 Then
                 If LTrim$(nomCabecera) <> "No#" Then
                    Mensaje = Mensaje & " La columna " & (i + 1) & " debe nombrarse 'No.'"
                 End If
              End If
              
            Next
            
            'return mensaje
End Function

' Retorna valor por defecto campo decimal o entero vacio
Private Function ObtenerValorNumero(dato As String, rvalor As String) As String
    If LTrim$(dato) = "" Then
       rvalor = "0"
    Else
       rvalor = dato
    End If
End Function

' Retorna valor por defecto campo bool
Private Function ObtenerValorBool(dato As String, rvalor As String) As String
                   If LTrim$(dato) = "True" Then
                        rvalor = "1"
                   ElseIf LTrim$(dato) = "False" Then
                        rvalor = "0"
                   Else
                        rvalor = "NULL"
                   End If
End Function

Private Function ObtenerDiaMes(dato As String, rvalor As String) As String
                   rvalor = Trim$(dato)
                   If Len(dato) = 1 Then
                        rvalor = "0" & Trim$(dato)
                   End If
End Function


Private Sub BtnSalir_Click()
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = " SELECT * FROM fo_extracto_BMSC_202101  "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral01_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - CHUQUISACA
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '1') "
        'queryinicial = " select * from fv_ventas_cobranza_det_traspasos where estado_conciliado = 'REG'  "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral02_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - LA PAZ
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '2') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral03_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - COCHABAMBA
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '3') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral04_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - ORURO
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '4') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral05_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - POTOSI
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '5') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral06_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - TARIJA
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '6') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral07_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - SANTA CRUZ
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '7') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral08_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - BENI
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '8') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub OptFilGral09_Click()
    'ORIGEN - NO CONCILIADOS SOFIA  - PANDO
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG' AND depto_codigo = '9') "
        rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub Picture2_Click()
    'db.Execute "update fo_recibos_detalle set CMPBTE_DEPOSITO_BCO = '" & Text11.Text & "', fecha_registro_bco= '" & DTP_Finicio & "', fecha_destino = '" & Date & "'  where correl_cobro = " & Ado_datos11.Recordset!correl_cobro & " "
    db.Execute "update fo_recibos_detalle set CMPBTE_DEPOSITO_BCO = '" & Text11.Text & "', fecha_registro_bco= '" & DTP_Finicio & "', fecha_destino = '" & Date & "', observaciones = '" & Text12.Text & "'  where correl_cobro = " & Ado_datos11.Recordset!correl_cobro & " "
    db.Execute "update fo_recibos_detalle set estado_conciliado = 'APR', fecha_concilia='" & Date & "', usr_concilia='" & glusuario & "'  where correl_cobro = " & Ado_datos11.Recordset!correl_cobro & " "
    
    db.Execute "update fo_extracto_BMSC_202101 set estado_conciliado = 'APR'  where cod_bancarizacion = '" & Ado_datos2.Recordset!cod_bancarizacion & "' "
    Fra_reporte.Visible = False
    Call ABRIR_LAS4
End Sub

Private Sub Picture3_Click()
    Unload Me
End Sub

Private Sub Picture5_Click()
    'Call ABRIR_LAS4
    FraBusca4.Visible = False
End Sub

'Private Sub Option1_Click()
'If Option1.Value = True Then
'dtc_rep_cod.Text = "%"
'dtc_rep_det.Text = "TODAS LAS PLANILLAS"
'dtc_depto.Text = "%"
'Else
'dtc_rep_cod.Text = ""
'dtc_rep_det.Text = ""
'End If
'End Sub

Private Sub rbtDia_Click(Index As Integer)
    If rbtDia(0).Value = True Then
        LblTime.Caption = rbtDia(0).Caption
        'lbl_inicial(0).Visible = True
        DtpFecha.Visible = True
        'lbl_inicial(1).Visible = False
        cmb_mes_ini.Visible = False
        cmb_gestion_rep.Visible = False
    End If
End Sub

Private Sub rbtMes_Click()

    If rbtMes.Value = True Then
        LblTime.Caption = rbtMes.Caption
        'lbl_inicial(0).Visible = False
        DtpFecha.Visible = False
        'lbl_inicial(1).Visible = True
        cmb_mes_ini.Visible = True
        cmb_gestion_rep.Visible = True
    End If

End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
  If SSTab1.Tab = 1 Then
    'ACTUALIZA LOS CONCILIADOS
    db.Execute "UPDATE fo_extracto_BMSC_202101 SET Cuenta = '4010620792' WHERE  (Cuenta IS NULL) "
    db.Execute "UPDATE fo_recibos_detalle SET estado_conciliado = 'REG' WHERE  (estado_conciliado IS NULL) "
    db.Execute "UPDATE fo_extracto_BMSC_202101 SET estado_conciliado = 'REG' WHERE  (estado_conciliado IS NULL) "
    
    db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.edif_codigo_corto = av_venta_cobranza_APR.edif_codigo_corto FROM fo_recibos_detalle inner JOIN av_venta_cobranza_APR ON fo_recibos_detalle.correl_cobro = av_venta_cobranza_APR.correl_cobro WHERE fo_recibos_detalle.edif_codigo_corto IS NULL"

    db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.estado_conciliado = 'APR' FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion "
    db.Execute "UPDATE fo_extracto_BMSC_202101 SET fo_extracto_BMSC_202101.estado_conciliado = 'APR' FROM fo_extracto_BMSC_202101 INNER JOIN fo_recibos_detalle ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion "
    
    '-- 1. TODOS = (#BANCARIZACION, CUENTA, FECHA, MONTO, CLIENTE)
    db.Execute "UPDATE fo_recibos_detalle SET  fo_recibos_detalle.nivel_conciliado = 1 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  = fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  = fo_extracto_BMSC_202101.cod_cliente WHERE fo_recibos_detalle.estado_conciliado = 'APR' "
    
    '-- 2. = (#BANCARIZACION, CUENTA, FECHA, MONTO ) Y <> (CLIENTE)
    db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.nivel_conciliado = 2 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  = fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  <> fo_extracto_BMSC_202101.cod_cliente WHERE fo_recibos_detalle.estado_conciliado = 'APR' AND  fo_recibos_detalle.nivel_conciliado <> 1 "

    '-- 3. = (#BANCARIZACION, CUENTA, FECHA) Y <> (CLIENTE, MONTO)
    db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.nivel_conciliado = 3 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  <> fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  <> fo_extracto_BMSC_202101.cod_cliente WHERE fo_recibos_detalle.estado_conciliado = 'APR' AND  fo_recibos_detalle.nivel_conciliado <> 1 AND  fo_recibos_detalle.nivel_conciliado <> 2 "

    '-- 4. = (#BANCARIZACION, CUENTA) Y <> (CLIENTE, MONTO, FECHA)
    'db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.estado_conciliado = 'APR', nivel_conciliado = 4 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  <> fo_extracto_BMSC_202101.fecha_transaccion AND fo_recibos_detalle.cobranza_bs  <> fo_extracto_BMSC_202101.credito AND fo_recibos_detalle.edif_codigo_corto  <> fo_extracto_BMSC_202101.cod_cliente "
    db.Execute "UPDATE fo_recibos_detalle SET fo_recibos_detalle.nivel_conciliado = 4 FROM fo_recibos_detalle INNER JOIN fo_extracto_BMSC_202101 ON fo_recibos_detalle.cmpbte_deposito_bco = fo_extracto_BMSC_202101.cod_bancarizacion AND fo_recibos_detalle.cta_codigo_destino = fo_extracto_BMSC_202101.cuenta AND fo_recibos_detalle.fecha_registro_bco  = fo_extracto_BMSC_202101.fecha_transaccion WHERE fo_recibos_detalle.estado_conciliado = 'APR' AND  fo_recibos_detalle.nivel_conciliado <> 1 AND  fo_recibos_detalle.nivel_conciliado <> 2 AND  fo_recibos_detalle.nivel_conciliado <> 3 "

    Call ABRIR_LAS4
  End If
End Sub

Private Sub ABRIR_LAS4()
    'ORIGEN - NO CONCILIADOS SOFIA
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
    BtnModDetalle.Visible = True
    queryinicial = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'REG')  "
    rs_datos11.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos11.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
    
    'EXTRACTO BANCARIO - NO CONCILIADOS
    Set rs_datos2 = New Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    queryinicial1 = " select * from fo_extracto_BMSC_202101 WHERE ((credito IS NOT NULL or credito <> '0') AND estado_conciliado = 'REG') "
    rs_datos2.Open queryinicial1, db, adOpenKeyset, adLockOptimistic
    'rs_datos2.Open " SELECT * FROM fo_extracto_BMSC_202101 where estado_conciliado = 'REG' ", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos2.Recordset = rs_datos2.DataSource
    Set dg_datos2.DataSource = Ado_datos2.Recordset
    
    
    'EXTRACTO BANCARIO - CONCILIADOS
    Set rs_datos3 = New Recordset
    If rs_datos3.State = 1 Then rs_datos2.Close
    'rs_datos3.Open " SELECT * FROM fo_extracto_BMSC_202101 where estado_conciliado = 'APR' ", db, adOpenKeyset, adLockOptimistic
    queryinicial2 = " SELECT * FROM fo_extracto_BMSC_202101 where (estado_conciliado = 'APR') "
    rs_datos3.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos3.Recordset = rs_datos3.DataSource
    Set dg_datos3.DataSource = Ado_datos3.Recordset
    If Ado_datos3.Recordset.RecordCount > 0 Then
    
'        'DESTINO - CONCILIADOS SOFIA
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'            DtGLista.Visible = False
'            'rs_datos14.Open "select * from fv_ventas_cobranza_det_traspasos where estado_conciliado = 'APR' ", db, adOpenKeyset, adLockOptimistic
'            'queryinicial2 = " select * from fv_ventas_cobranza_det_VS_extracto_BMSC_2021 where (estado_conciliado_tes = 'APR') "
'            queryinicial2 = " select * from fv_ventas_cobranza_det_traspasos where (estado_conciliado = 'APR' and cmpbte_deposito_bco = '" & Ado_datos3.Recordset!cod_bancarizacion & "' AND fecha_registro_bco = '" & Ado_datos3.Recordset!fecha_transaccion & "') "
'            rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'            rs_datos14.Sort = "fecha_registro_bco, cmpbte_deposito_bco "
'        Set ado_datos14.Recordset = rs_datos14.DataSource
'        'ado_datos14.Recordset.Requery
'        If ado_datos14.Recordset.RecordCount > 0 Then
'            deta2 = 1
'            DtGLista.Visible = True
'            Set DtGLista.DataSource = ado_datos14.Recordset
'            'Set dg_datos3.DataSource = ado_datos14.Recordset
'        Else
'            deta2 = 0
'            DtGLista.Visible = False
'        End If
    End If

End Sub

Private Sub SSTab1_DblClick()
'    If SSTab1.Tab = 0 Then
'        'SSTab1.TabEnabled(0) = True
'        'SSTab1.TabEnabled(1) = False
'    Else
''        SSTab1.Tab = 0
''        SSTab1.TabEnabled(0) = True
''        SSTab1.TabEnabled(1) = False
''        SSTab1.TabEnabled(2) = False
''           FrmEditaDet.Visible = False
''           DtGLista.Visible = False
''           adoao_solicitud_lista.Visible = False
'    End If

End Sub

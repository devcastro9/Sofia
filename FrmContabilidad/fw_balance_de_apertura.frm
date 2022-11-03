VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_balance_de_apertura 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Contabilidad - Balance de Apertura"
   ClientHeight    =   9165
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "fw_balance_de_apertura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   16
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "fw_balance_de_apertura.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   25
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         Picture         =   "fw_balance_de_apertura.frx":11C4
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   24
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "fw_balance_de_apertura.frx":1A91
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   23
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "fw_balance_de_apertura.frx":2246
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   22
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "fw_balance_de_apertura.frx":2A79
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "fw_balance_de_apertura.frx":31C5
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   20
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "fw_balance_de_apertura.frx":3ADA
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   19
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "fw_balance_de_apertura.frx":4299
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "fw_balance_de_apertura.frx":46DB
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE APERTURA"
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
         Left            =   12450
         TabIndex        =   26
         Top             =   195
         Width           =   2625
      End
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
      ScaleWidth      =   20280
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "fw_balance_de_apertura.frx":48E5
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "fw_balance_de_apertura.frx":51D1
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   13
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE APERTURA"
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
         Left            =   12405
         TabIndex        =   15
         Top             =   195
         Width           =   2625
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00800000&
      Height          =   2880
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   17535
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Aprobados"
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
         Left            =   10440
         TabIndex        =   10
         Top             =   2490
         Width           =   1815
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todas"
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
         TabIndex        =   9
         Top             =   2490
         Width           =   915
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "fw_balance_de_apertura.frx":59A7
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   17280
         _ExtentX        =   30480
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "cuenta"
            Caption         =   "Cuenta"
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
            DataField       =   "subcta1"
            Caption         =   "SubCta1"
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
            DataField       =   "subcta2"
            Caption         =   "SubCta2"
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
            DataField       =   "aux1"
            Caption         =   "Aux1"
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
            DataField       =   "aux2"
            Caption         =   "Aux2"
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
            DataField       =   "aux3"
            Caption         =   "Aux3"
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
            DataField       =   "denominacion_aux1"
            Caption         =   "Denominacion_Aux1"
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
            DataField       =   "denominacion_aux2"
            Caption         =   "Denominacion_Aux2"
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
            DataField       =   "denominacion_aux3"
            Caption         =   "Denomincacion_Aux3"
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
            DataField       =   "NombreCta"
            Caption         =   "NombreCta"
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
            DataField       =   "DebeSaldoIBs"
            Caption         =   "SaldoDebeBs"
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
            DataField       =   "HaberSaldoIBs"
            Caption         =   "SaldoHaberBs"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
            DataField       =   "fecha_registro"
            Caption         =   "Fecha_Reg."
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
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
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
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2400
         Width           =   17265
         _ExtentX        =   30454
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
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Registro"
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
      Height          =   5385
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   17535
      Begin VB.Frame Fra_Aux 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipos de Auxiliares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1695
         Left            =   240
         TabIndex        =   54
         Top             =   3600
         Width           =   17175
         Begin VB.PictureBox Buscar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15480
            Picture         =   "fw_balance_de_apertura.frx":59BF
            ScaleHeight     =   495
            ScaleWidth      =   1215
            TabIndex        =   72
            Top             =   840
            Width           =   1215
         End
         Begin VB.PictureBox Buscar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   9480
            Picture         =   "fw_balance_de_apertura.frx":6174
            ScaleHeight     =   495
            ScaleWidth      =   1215
            TabIndex        =   71
            Top             =   840
            Width           =   1215
         End
         Begin VB.PictureBox Buscar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000015&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3360
            Picture         =   "fw_balance_de_apertura.frx":6929
            ScaleHeight     =   495
            ScaleWidth      =   1215
            TabIndex        =   70
            Top             =   840
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "fw_balance_de_apertura.frx":70DE
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "Aux1"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "fw_balance_de_apertura.frx":70F7
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6240
            TabIndex        =   56
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "Aux2"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "fw_balance_de_apertura.frx":7110
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   12240
            TabIndex        =   57
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "Aux3"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "fw_balance_de_apertura.frx":7129
            DataField       =   "Cod_Aux1"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desc1"
            BoundColumn     =   "codigo1"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc9 
            Bindings        =   "fw_balance_de_apertura.frx":7142
            DataField       =   "Cod_Aux2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6240
            TabIndex        =   45
            Top             =   1320
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desc2"
            BoundColumn     =   "codigo2"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc10 
            Bindings        =   "fw_balance_de_apertura.frx":715B
            DataField       =   "Cod_Aux3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   12240
            TabIndex        =   47
            Top             =   1320
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desc3"
            BoundColumn     =   "codigo3"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "fw_balance_de_apertura.frx":7175
            DataField       =   "Cod_Aux1"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo1"
            BoundColumn     =   "codigo1"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo9 
            Bindings        =   "fw_balance_de_apertura.frx":718E
            DataField       =   "Cod_Aux2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6240
            TabIndex        =   67
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo2"
            BoundColumn     =   "codigo2"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo10 
            Bindings        =   "fw_balance_de_apertura.frx":71A7
            DataField       =   "Cod_Aux3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   12240
            TabIndex        =   68
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo3"
            BoundColumn     =   "codigo3"
            Text            =   "Todos"
         End
         Begin VB.Label Txt_campo5 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            DataField       =   "Cod_Aux1"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Txt_campo10 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            DataField       =   "Denominacion_Aux3"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   12240
            TabIndex        =   79
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Label Txt_campo9 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            DataField       =   "Cod_Aux3"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   12240
            TabIndex        =   78
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Txt_campo8 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            DataField       =   "Denominacion_Aux2"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   6240
            TabIndex        =   77
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Label Txt_campo7 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            DataField       =   "Cod_Aux2"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   6240
            TabIndex        =   76
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Txt_campo6 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            DataField       =   "Denominacion_Aux1"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Label lblAux1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 1"
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
            Index           =   1
            Left            =   165
            TabIndex        =   60
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblAux2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 2"
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
            Left            =   6240
            TabIndex        =   59
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblAux3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 3"
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
            Left            =   12240
            TabIndex        =   58
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Fra_ABM3 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   240
         TabIndex        =   39
         Top             =   2400
         Width           =   17175
         Begin VB.TextBox TxtDebe 
            DataField       =   "DebeSaldoIBs"
            DataSource      =   "Ado_datos"
            Height          =   405
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   41
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox TxtCodigo 
            DataField       =   "Cod_Anterior"
            DataSource      =   "Ado_datos"
            Height          =   405
            Left            =   14280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox TxtHaber 
            DataField       =   "HaberSaldoIBs"
            DataSource      =   "Ado_datos"
            Height          =   405
            Left            =   8280
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   2475
         End
         Begin VB.Label lblCodigo 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Codigo Anterior CGI"
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
            Left            =   12240
            TabIndex        =   65
            Top             =   720
            Width           =   1770
         End
         Begin VB.Label lblSHaber 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Saldo Haber Bs."
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
            Left            =   6480
            TabIndex        =   64
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Txt_campo4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00"
            DataField       =   "SubCta2"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   14280
            TabIndex        =   53
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Txt_campo3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "SubCta1"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8280
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sub Cuenta - 1"
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
            Left            =   6600
            TabIndex        =   51
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sub Cuenta - 2"
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
            Left            =   12600
            TabIndex        =   50
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Codigo de la Cuenta"
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
            TabIndex        =   49
            Top             =   240
            Width           =   1830
         End
         Begin VB.Label Txt_campo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "cuenta"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   48
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblSDebe 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Saldo Debe Bs."
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
            TabIndex        =   46
            Top             =   720
            Width           =   1425
         End
      End
      Begin VB.Frame Fra_ABM2 
         BackColor       =   &H00C0C0C0&
         Height          =   1335
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   17175
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "fw_balance_de_apertura.frx":71C1
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3600
            TabIndex        =   40
            Top             =   600
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NombreCta"
            BoundColumn     =   "correl"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "fw_balance_de_apertura.frx":71DA
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   15240
            TabIndex        =   34
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "subcta2"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "fw_balance_de_apertura.frx":71F3
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   12000
            TabIndex        =   35
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "subcta1"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "fw_balance_de_apertura.frx":720C
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1920
            TabIndex        =   38
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "Cuenta"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_Aux1 
            Bindings        =   "fw_balance_de_apertura.frx":7225
            DataField       =   "correl"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   360
            TabIndex        =   37
            Top             =   600
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "correl"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin VB.Label lblNomCta 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nombre de la Cuenta"
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
            Left            =   3600
            TabIndex        =   73
            Top             =   240
            Width           =   1905
         End
         Begin VB.Label lblcorelativo 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Correlativo"
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
            TabIndex        =   69
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sub Cuenta 2"
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
            Left            =   13920
            TabIndex        =   63
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sub Cuenta 1"
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
            TabIndex        =   62
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label lblcuenta 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cuenta"
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
            Left            =   1920
            TabIndex        =   61
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.Frame Fra_ABM1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   17175
         Begin VB.ComboBox Cmbgestion 
            DataField       =   "ges_gestion"
            DataSource      =   "Ado_datos"
            Height          =   315
            ItemData        =   "fw_balance_de_apertura.frx":723E
            Left            =   10560
            List            =   "fw_balance_de_apertura.frx":7260
            TabIndex        =   36
            Text            =   "0"
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label lblGestion 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   9720
            TabIndex        =   81
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Txt_cuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "5"
            DataField       =   "Mov"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   2040
            TabIndex        =   80
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nivel de la Cuenta"
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
            TabIndex        =   32
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estado de Registro"
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
            Left            =   13200
            TabIndex        =   31
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Titulo/Subtitulo/Detalle"
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
            Left            =   5040
            TabIndex        =   30
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label Txt_campo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "D"
            DataField       =   "Mov"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   7200
            TabIndex        =   29
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Txt_estado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "estado_codigo"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   15120
            TabIndex        =   28
            Top             =   360
            Width           =   855
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
      ScaleWidth      =   15120
      TabIndex        =   0
      Top             =   9165
      Width           =   15120
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
   Begin Crystal.CrystalReport cr01 
      Left            =   16200
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2400
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4680
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6960
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   9240
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   11640
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   13920
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   120
      Top             =   9480
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   2400
      Top             =   9480
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
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      DataField       =   "estado_codigo"
      DataSource      =   "Ado_datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "fw_balance_de_apertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset

Dim rs_det0 As New ADODB.Recordset
Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset
Dim rs_det3 As New ADODB.Recordset
Dim rs_det4 As New ADODB.Recordset
Dim rs_det5 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String


Dim var_cod, VAR_COD2, VAR_COD1, VAR_COD3 As String
Dim VAR_VAL As String
Dim VAR_SUB1 As String
Dim VAR_SW As String
Dim VAR_CTA, VAR_SUB2 As String

Dim VAR_TABLA, VAR_CODIGO, VAR_DES As String


Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If Ado_datos.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datos.Recordset!estado_codigo = "APR"
         Ado_datos.Recordset!fecha_registro = Date
        ' Ado_datos.Recordset!usr_codigo = glusuario
         Ado_datos.Recordset.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   'error
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
  On Error GoTo UpdateErr
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelBatch
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        Call OptFilGral2_Click
            
        dtc_codigo8.Visible = False
        dtc_desc8.Visible = False
        dtc_codigo9.Visible = False
        dtc_desc9.Visible = False
        dtc_codigo10.Visible = False
        dtc_desc10.Visible = False
            
        Buscar1.Visible = False
        Buscar2.Visible = False
        Buscar3.Visible = False
            
        rs_datos.MoveFirst
'        rs_datos.Enabled = False
        mbDataChanged = False
        Fra_Aux.Enabled = False
        FraGlobal.Enabled = False
        FraNavega.Enabled = True
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        txt_codigo.Enabled = True
        dtc_desc1.Enabled = True
        VAR_SW = ""
    End If
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(rs_datos!cuenta, rs_datos!subcta1, rs_datos!subcta2) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If Ado_datos.Recordset!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datos.Recordset!estado_codigo = "ANL"
         Ado_datos.Recordset!fecha_registro = Date
         'Ado_datos.Recordset!usr_codigo = glusuario
         Ado_datos.Recordset.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
        Fra_Aux.Enabled = False
        dtc_codigo8.Visible = False
        dtc_desc8.Visible = False
        dtc_codigo9.Visible = False
        dtc_desc9.Visible = False
        dtc_codigo10.Visible = False
        dtc_desc10.Visible = False
        Buscar1.Visible = False
        Buscar2.Visible = False
        Buscar3.Visible = False
        
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
    
        db.Execute " INSERT INTO fo_balance_apertura (Cuenta,                             SubCta1,                     SubCta2,                  aux1,                        aux2,                   aux3,                 denominacion_aux1,        denominacion_aux2,       denominacion_aux3,        NombreCta,              DebeSaldoIBs,               DebeSaldoISus,                                     HaberSaldoIBs,           HaberSaldoISus,                               Usr_Usuario,    Fecha_Registro,estado_codigo,Usr_aprobacion,fecha_aprobacion,Verificado,correl,        Cod_Anterior,                Cod_Aux1,               Cod_Aux2,                     Cod_Aux3,          cod_cta,    ges_gestion) " & _
                                            " VALUES ('" & dtc_codigo1.Text & "','" & dtc_codigo2.Text & "','" & dtc_codigo3.Text & "','" & dtc_codigo5.Text & "','" & dtc_codigo6.Text & "','" & dtc_codigo7.Text & "','" & dtc_desc8.Text & "','" & dtc_desc9.Text & "','" & dtc_desc10.Text & "','" & dtc_desc1.Text & "'," & CDbl(TxtDebe.Text) & "," & CDbl(TxtDebe.Text) / GlTipoCambioOficial & "," & CDbl(TxtHaber.Text) & "," & CDbl(TxtHaber.Text) / GlTipoCambioOficial & ",'" & glusuario & "','" & Date & "','REG','" & glusuario & "','" & Date & "','N','" & dtc_aux1 & "','" & TxtCodigo.Text & "','" & dtc_codigo8.Text & "','" & dtc_codigo9.Text & "','" & dtc_codigo10.Text & "','""','" & Cmbgestion.Text & "')"

       ' Set rs_det0 = New ADODB.Recordset
'            If rs_det0.State = 1 Then rs_det0.Close
'            rs_det0.Open "Select correl  as correl2 from CC_plan_cuentas where Cuenta= '" & VAR_CTA & "' and SubCta1= '00' and SubCta2='00' and nivel = '" & dtc_cuenta.Text & "'", db, adOpenStatic
'            If rs_det0.RecordCount > 0 Then
'                VAR_COD3 = rs_det0!correl2
'            End If

'
'        If dtc_cuenta.Text = 5 Then
'
'             db.Execute " INSERT INTO CC_plan_cuentas (Cuenta,SubCta1,SubCta2,NombreCta,Aux1,Aux2,Aux3,Mov,Niv_Public,Naturaleza,nivel,t_comision,estado_codigo,Usr_codigo,Fecha_registro, correl) " & _
'            " VALUES ('" & VAR_CTA & "','00','00','" & TxtConcepto.Text & "','00','00','00','T','N','D','" & dtc_cuenta.Text & "','N','REG','" & glusuario & "','" & Date & "', " & VAR_COD3 & ")"
'        End If
    End If

    If VAR_SW = "MOD" Then
             db.Execute " UPDATE fo_balance_apertura set DebeSaldoIBs='" & TxtDebe.Text & "',HaberSaldoIBs='" & TxtHaber.Text & "',Cod_Anterior='" & TxtCodigo.Text & "',Cod_Aux1='" & dtc_codigo8.Text & "',Cod_Aux2='" & dtc_codigo9.Text & "',Cod_Aux3='" & dtc_codigo10.Text & "',Denominacion_Aux1='" & dtc_desc8 & "',Denominacion_Aux2='" & dtc_desc9 & "',Denominacion_Aux3='" & dtc_desc10 & "',ges_gestion='" & Cmbgestion.Text & "' WHERE fo_balance_apertura.correl_ba='" & Ado_datos.Recordset!correl_ba & "' "
'
     End If
     

     Call OptFilGral2_Click
     rs_datos.Update
     rs_datos.MoveLast
     mbDataChanged = False
      
      Fra_ABM.Enabled = True
      FraNavega.Enabled = True
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      Fra_ABM2.Visible = False
      dg_datos.Enabled = True
      'txt_codigo.Enabled = True
     'dtc_cuenta.Enabled = True
    ' dtc_desc1.Enabled = True
      VAR_SW = ""
End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'habilitar codigo cuando se transcribe
If Cmbgestion.Text = "" Then
    MsgBox "Debe registrar la " + lblGestion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_aux1.Text = "" Then
    MsgBox "Debe registrar el " + lblcorelativo.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar la " + Lblcuenta.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_desc1.Text = "" Then
    MsgBox "Debe registrar: " + lblNomCta.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If TxtDebe.Text = "" Then
    MsgBox "Debe registrar: " + lblSDebe.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If TxtHaber.Text = "" Then
    MsgBox "Debe registrar: " + lblSHaber.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If TxtCodigo.Text = "" Then
    MsgBox "Debe registrar: " + lblCodigo.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If dtc_desc8.Text = "" Then
'    MsgBox "Debe registrar: " + lblAux1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'   If dtc_desc9.Text = "" Then
'    MsgBox "Debe registrar: " + lblAux2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'   If dtc_desc10.Text = "" Then
'    MsgBox "Debe registrar: " + lblAux3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub BtnImprimir_Click()
  On Error GoTo UpdateErr
  Dim iResult As Integer
  CR01.WindowShowPrintSetupBtn = True
  CR01.WindowShowRefreshBtn = True
  CR01.ReportFileName = App.Path & "\REPORTES\contabilidad\fr_balance_apertura.rpt"
   CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  CR01.WindowState = crptMaximized
     Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
   If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_ABM.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        FraNavega.Enabled = False
        dg_datos.Enabled = False
        VAR_SW = "MOD"
        Fra_ABM1.Enabled = True
    
        Fra_ABM3.Enabled = True
        Fra_Aux.Enabled = True
         
        'TxtConcepto.Enabled = True
        Fra_Aux.Visible = True
        Fra_ABM2.Visible = False
        Fra_ABM3.Visible = True
        Buscar1.Visible = True
        Buscar2.Visible = True
        Buscar3.Visible = True
        
       'dtc_cuenta.Text = rs_datos!nivel
        Txt_campo2.Caption = cuenta
        txt_campo3.Caption = subcta1
        txt_campo4.Caption = subcta2
        
        
           
   Else

'             TxtDebe.Visible = True
'             TxtHaber.Visible = True
'             TxtCodigo.Visible = True
'             dtc_codigo8.Visible = True
'             dtc_desc8.Visible = True
'             dtc_codigo9.Visible = True
'             dtc_desc9.Visible = True
'             dtc_codigo10.Visible = True
'             dtc_desc10.Visible = True
             
'            Call ABRIR_AUX1
'            Call ABRIR_AUX2
'            Call ABRIR_AUX3
         MsgBox "No se puede MODIFICAR un registro APROBADO o Errado ...", vbExclamation, "Validación de Registro"
    End If
'  lblStatus.Caption = "Modificar registro"
        'TxtConcepto.SetFocus
  Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
'  If glPersOtro = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_datos!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_datos!ocup_descripcion
'  End If
'  glPersOtro = "N"
  Unload Me
End Sub

Private Sub Buscar1_Click()
 Call ABRIR_AUX1

    If VAR_TABLA = "NN" And dtc_codigo5 = "00" Then
        dtc_codigo8.Text = "0"
        dtc_desc8.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
  dtc_codigo8.Visible = True
 dtc_desc8.Visible = True
        Set rs_datos8 = New ADODB.Recordset
        If rs_datos8.State = 1 Then rs_datos8.Close
            rs_datos8.Open "Select " + VAR_CODIGO + " as codigo1 , " + VAR_DES + " as desc1 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos8.Recordset = rs_datos8
            dtc_desc8.BoundText = dtc_codigo8.BoundText
    End If
End Sub

Private Sub Buscar2_Click()
 Call ABRIR_AUX2
 
    If VAR_TABLA = "NN" And dtc_codigo6 = "00" Then
        dtc_codigo9.Text = "0"
        dtc_desc9.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
    dtc_codigo9.Visible = True
    dtc_desc9.Visible = True
        Set rs_datos9 = New ADODB.Recordset
        If rs_datos9.State = 1 Then rs_datos9.Close
            rs_datos9.Open "Select " + VAR_CODIGO + " as codigo2 , " + VAR_DES + " as desc2 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos9.Recordset = rs_datos9
            dtc_desc9.BoundText = dtc_codigo9.BoundText
    End If
End Sub

Private Sub Buscar3_Click()
 Call ABRIR_AUX3
    If VAR_TABLA = "NN" And dtc_codigo7 = "00" Then
        dtc_codigo10.Text = ""
        dtc_desc10.Text = ""
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else

        dtc_codigo10.Visible = True
        dtc_desc10.Visible = True
        Set rs_datos10 = New ADODB.Recordset
        If rs_datos10.State = 1 Then rs_datos10.Close
            rs_datos10.Open "Select " + VAR_CODIGO + " as codigo3 , " + VAR_DES + " as desc3 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos10.Recordset = rs_datos10
            dtc_desc10.BoundText = dtc_codigo10.BoundText
    End If
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
  dtc_desc1.BoundText = dtc_aux1.BoundText
  dtc_codigo1.BoundText = dtc_aux1.BoundText
  dtc_codigo2.BoundText = dtc_aux1.BoundText
  dtc_codigo3.BoundText = dtc_aux1.BoundText
  dtc_codigo5.BoundText = dtc_aux1.BoundText
  dtc_codigo6.BoundText = dtc_aux1.BoundText
  dtc_codigo7.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_desc2.BoundText = Dtc_aux2.BoundText
    dtc_codigo2.BoundText = Dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_aux3.BoundText
    dtc_codigo3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
'    dtc_desc4.BoundText = dtc_Aux4.BoundText
'    dtc_codigo4.BoundText = dtc_Aux4.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
  dtc_desc1.BoundText = dtc_codigo1.BoundText
  dtc_aux1.BoundText = dtc_codigo1.BoundText
  dtc_codigo2.BoundText = dtc_codigo1.BoundText
  dtc_codigo3.BoundText = dtc_codigo1.BoundText
  dtc_codigo5.BoundText = dtc_codigo1.BoundText
  dtc_codigo6.BoundText = dtc_codigo1.BoundText
  dtc_codigo7.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
 dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
  dtc_codigo1.BoundText = dtc_codigo2.BoundText
  dtc_aux1.BoundText = dtc_codigo2.BoundText
  dtc_desc1.BoundText = dtc_codigo2.BoundText
  dtc_codigo3.BoundText = dtc_codigo2.BoundText
  dtc_codigo5.BoundText = dtc_codigo2.BoundText
  dtc_codigo6.BoundText = dtc_codigo2.BoundText
  dtc_codigo7.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
  dtc_codigo1.BoundText = dtc_codigo3.BoundText
  dtc_aux1.BoundText = dtc_codigo3.BoundText
  dtc_desc1.BoundText = dtc_codigo3.BoundText
  dtc_codigo2.BoundText = dtc_codigo3.BoundText
  dtc_codigo5.BoundText = dtc_codigo3.BoundText
  dtc_codigo6.BoundText = dtc_codigo3.BoundText
  dtc_codigo7.BoundText = dtc_codigo3.BoundText
  
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
    
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
  dtc_desc1.BoundText = dtc_codigo5.BoundText
  dtc_codigo1.BoundText = dtc_codigo5.BoundText
  dtc_codigo2.BoundText = dtc_codigo5.BoundText
  dtc_codigo3.BoundText = dtc_codigo5.BoundText
  dtc_aux1.BoundText = dtc_codigo5.BoundText
  dtc_codigo6.BoundText = dtc_codigo5.BoundText
  dtc_codigo7.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
  dtc_desc1.BoundText = dtc_codigo6.BoundText
  dtc_codigo1.BoundText = dtc_codigo6.BoundText
  dtc_codigo2.BoundText = dtc_codigo6.BoundText
  dtc_codigo3.BoundText = dtc_codigo6.BoundText
  dtc_aux1.BoundText = dtc_codigo6.BoundText
  dtc_codigo5.BoundText = dtc_codigo6.BoundText
  dtc_codigo7.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
  dtc_desc1.BoundText = dtc_codigo7.BoundText
  dtc_codigo1.BoundText = dtc_codigo7.BoundText
  dtc_codigo2.BoundText = dtc_codigo7.BoundText
  dtc_codigo3.BoundText = dtc_codigo7.BoundText
  dtc_aux1.BoundText = dtc_codigo7.BoundText
  dtc_codigo6.BoundText = dtc_codigo7.BoundText
  dtc_codigo5.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
  dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
  dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub
'Private Sub dtc_cuenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii >= 0 Then
'    KeyAscii = 0
'    Else
'    Exit Sub
'    End If
'End Sub
Private Sub dtc_cuenta_LostFocus()
Call Abrir_Aux
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub
'
'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_codigo7.BoundText
'    dtc_codigo1.BoundText = dtc_codigo7.BoundText
'    dtc_codigo10.BoundText = dtc_codigo7.BoundText
'End Sub
'
'Private Sub dtc_codigo8_Click(Area As Integer)
'    dtc_desc2.BoundText = dtc_codigo8.BoundText
'    dtc_codigo2.BoundText = dtc_codigo8.BoundText
'    dtc_codigo11.BoundText = dtc_codigo8.BoundText
'End Sub
'
'Private Sub dtc_codigo9_Click(Area As Integer)
'    dtc_desc3.BoundText = dtc_codigo9.BoundText
'    dtc_codigo3.BoundText = dtc_codigo9.BoundText
'    dtc_codigo12.BoundText = dtc_codigo9.BoundText
'End Sub

Private Sub dtc_desc1_Click(Area As Integer)
  dtc_codigo1.BoundText = dtc_desc1.BoundText
  dtc_aux1.BoundText = dtc_desc1.BoundText
  dtc_codigo2.BoundText = dtc_desc1.BoundText
  dtc_codigo3.BoundText = dtc_desc1.BoundText
  dtc_codigo5.BoundText = dtc_desc1.BoundText
  dtc_codigo6.BoundText = dtc_desc1.BoundText
  dtc_codigo7.BoundText = dtc_desc1.BoundText
    
'    If dtc_cuenta.Text > 2 Then
'        Call pnivel2(dtc_codigo1.Text)
'        dtc_desc2.Enabled = True
'    End If

End Sub

Private Sub pnivel2(codigo1 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_plan_nivel2 where left(Cuenta,1)= '" & Left(codigo1, 1) & "'"
   Set dtc_codigo2.RowSource = Nothing
   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo2.ReFill
   dtc_codigo2.BoundText = Empty
   
   Set dtc_desc2.RowSource = Nothing
   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc2.ReFill
   dtc_desc2.BoundText = Empty

End Sub

Private Sub dtc_desc2_Click(Area As Integer)
  dtc_codigo2.BoundText = dtc_desc2.BoundText
   ' dtc_Aux2.BoundText = dtc_desc2.BoundText
    
End Sub

Private Sub pnivel3(codigo2 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_plan_nivel3 where left(Cuenta,2)= '" & Left(codigo2, 2) & "'"
   Set dtc_codigo3.RowSource = Nothing
   Set dtc_codigo3.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo3.ReFill
   dtc_codigo3.BoundText = Empty
   
   Set dtc_desc3.RowSource = Nothing
   Set dtc_desc3.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc3.ReFill
   dtc_desc3.BoundText = Empty

End Sub

Private Sub dtc_desc3_Click(Area As Integer)
 dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
'    If dtc_cuenta.Text > 4 Then
'        Call pnivel4(dtc_codigo3.Text)
'        dtc_desc4.Enabled = True
'    End If
End Sub
Private Sub pnivel4(codigo3 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_plan_nivel4 where left(Cuenta,3)= '" & Left(codigo3, 3) & "'"
   Set dtc_codigo4.RowSource = Nothing
   Set dtc_codigo4.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo4.ReFill
   dtc_codigo4.BoundText = Empty
   
   Set dtc_desc4.RowSource = Nothing
   Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc4.ReFill
   dtc_desc4.BoundText = Empty

End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    'dtc_Aux4.BoundText = dtc_desc4.BoundText
    
End Sub

Private Sub pnivel5(codigo4 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_tipo_auxiliar where left(Cuenta,4)= '" & Left(codigo4, 4) & "'"
   Set dtc_codigo5.RowSource = Nothing
   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo5.ReFill
   dtc_codigo5.BoundText = Empty
   
   Set dtc_desc5.RowSource = Nothing
   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc5.ReFill
   dtc_desc5.BoundText = Empty

End Sub

Private Sub dtc_desc5_Click(Area As Integer)
  dtc_codigo5.BoundText = dtc_desc5.BoundText

'   If dtc_cuenta.Text > 6 Then
'        Call pnivel6(dtc_codigo5.Text)
'        dtc_desc5.Enabled = True
'       dtc_desc5.Enabled = True
'       Fra_Aux.Enabled = True
'    End If
End Sub

Private Sub pnivel6(codigo5 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_tipo_auxiliar where left(aux,5)= '" & Left(codigo5, 5) & "'"
   Set dtc_codigo6.RowSource = Nothing
   Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo6.ReFill
   dtc_codigo6.BoundText = Empty
   
   Set dtc_desc6.RowSource = Nothing
   Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc6.ReFill
   dtc_desc6.BoundText = Empty

End Sub

Private Sub dtc_desc6_Click(Area As Integer)
dtc_codigo6.BoundText = dtc_desc6.BoundText
  'dtc_codigo5.BoundText = dtc_desc5.BoundText

End Sub

Private Sub dtc_desc7_Click(Area As Integer)
dtc_codigo7.BoundText = dtc_desc7.BoundText
  'dtc_codigo5.BoundText = dtc_desc5.BoundText
'   If dtc_cuenta.Text > 6 Then
'        Call pnivel7(dtc_codigo7.Text)
'        dtc_desc7.Enabled = True
'       dtc_desc7.Enabled = True
'       Fra_Aux.Enabled = True
'    End If
End Sub

Private Sub dtc_desc10_Click(Area As Integer)
  dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
  dtc_codigo8.BoundText = dtc_desc8.BoundText
  
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
dtc_codigo9.BoundText = dtc_desc9.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click
'   Call ABRIR_TABLA
'   txt_codigo.Enabled = True
    VAR_SW = ""
    mbDataChanged = False
    Fra_ABM.Enabled = False
    Buscar1.Enabled = True
    Buscar2.Enabled = True
    Buscar3.Enabled = True
    Buscar1.Visible = False
    Buscar2.Visible = False
    Buscar3.Visible = False
    Fra_Aux.Enabled = False
    dg_datos.Enabled = True
 'Call Buscar2_Click
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
  '  txt_Tcta.Visible = False
'    txt_Tscta1.Visible = False
'    txt_Tscta2.Visible = False
   ' txt_desc1.Visible = False
    
'    txt_Tcta2.Visible = False
'    txt_Tscta12.Visible = False
'    txt_Tscta22.Visible = False
'    txt_desc2.Visible = False
	Call SeguridadSet(Me)
End Sub

Private Sub OptFilGral2_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from fo_balance_apertura "      ' todos
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub OptFilGral1_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from fo_balance_apertura where estado_codigo= 'REG' "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub OptFilGral3_Click()
'  Set rs_datos = New Recordset
'  If rs_datos.State = 1 Then rs_datos.Close
'  queryinicial = "select  * from CC_Plan_Cuentas  where estado_codigo= 'APR' "
'  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'  Set Ado_datos.Recordset = rs_datos.DataSource
'  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub OptFilGral4_Click()
'  Set rs_datos = New Recordset
'  If rs_datos.State = 1 Then rs_datos.Close
'  queryinicial = "select  * from CC_Plan_Cuentas  where mov= 'D' "
'  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'  Set Ado_datos = rs_datos.DataSource
'  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE nivel = '5' order by Cuenta ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText

'    Set rs_datos2 = New ADODB.Recordset
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    rs_datos2.Open "Select * from cc_tipo_auxiliar order  by correl ", db, adOpenStatic
'    Set Ado_datos2.Recordset = rs_datos
'    dtc_desc2.BoundText = dtc_codigo2.BoundText

'    Set rs_datos3 = New ADODB.Recordset
'    If rs_datos3.State = 1 Then rs_datos3.Close
'    'rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 <> '00' order by Cuenta ", db, adOpenStatic
'    rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'D' order by Cuenta ", db, adOpenStatic
'    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText

End Sub

Private Sub ABRIR_AUX1()
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from cc_tipo_auxiliar where aux = '" & dtc_codigo5 & "' order by aux ", db, adOpenStatic
    If rs_datos5.RecordCount > 0 Then
        VAR_TABLA = rs_datos5!NombreTabla
        VAR_CODIGO = rs_datos5!nombre_codigo
        VAR_DES = rs_datos5!nombre_descripcion
    Else
        VAR_TABLA = "NN"
        VAR_CODIGO = "NN"
        VAR_DES = "NN"
    End If
'    Set Ado_datos5.Recordset = rs_datos5
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub ABRIR_AUX2()
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from cc_tipo_auxiliar where aux = '" & dtc_codigo6 & "' order by aux ", db, adOpenStatic
    If rs_datos6.RecordCount > 0 Then
        VAR_TABLA = rs_datos6!NombreTabla
        VAR_CODIGO = rs_datos6!nombre_codigo
        VAR_DES = rs_datos6!nombre_descripcion
    Else
        VAR_TABLA = "NN"
        VAR_CODIGO = "NN"
        VAR_DES = "NN"
    End If
    
'    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub ABRIR_AUX3()
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from cc_tipo_auxiliar where aux = '" & dtc_codigo7 & "' order by aux ", db, adOpenStatic
    If rs_datos7.RecordCount > 0 Then
        VAR_TABLA = rs_datos7!NombreTabla
        VAR_CODIGO = rs_datos7!nombre_codigo
        VAR_DES = rs_datos7!nombre_descripcion
    Else
        VAR_TABLA = "NN"
        VAR_CODIGO = "NN"
        VAR_DES = "NN"
    End If
'    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

'Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub ABRIR_NIVEL1()
    Set rs_datos1 = New ADODB.Recordset
        If rs_datos1.State = 1 Then rs_datos1.Close
        'WHERE correl = '" & rs_datos!CORREL & "'
        rs_datos1.Open "Select * from cc_plan_nivel1   order by Cuenta ", db, adOpenStatic
        Set Ado_datos1.Recordset = rs_datos1
        dtc_desc1.BoundText = dtc_aux1.BoundText
        dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

Private Sub ABRIR_NIVEL2()
Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
   
'   WHERE correl = '" & rs_datos!CORREL & "'
     rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    'dtc_desc2.BoundText = dtc_Aux2.BoundText
'    dtc_codigo2.BoundText = dtc_Aux2.BoundText

End Sub

Private Sub ABRIR_NIVEL3()
Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
     rs_datos3.Open "Select * from cc_plan_nivel3 order by Cuenta ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'     dtc_desc3.BoundText = dtc_Aux3.BoundText
    'dtc_codigo3.BoundText = dtc_Aux3.BoundText
End Sub

Private Sub ABRIR_NIVEL4()
Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
     rs_datos4.Open "Select * from cc_plan_nivel4  order by Cuenta ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
'     dtc_desc4.BoundText = dtc_Aux4.BoundText
   ' dtc_codigo4.BoundText = dtc_Aux4.BoundText
    
End Sub

Private Sub Abrir_Aux()
    Fra_ABM2.Visible = True
    If dtc_cuenta.Text = 1 Then
        
        dtc_codigo1.Visible = False
        dtc_desc1.Visible = False
        dtc_codigo2.Visible = False
        dtc_desc2.Visible = False
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False
        
        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
    End If
    
    If dtc_cuenta.Text = 2 Then
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = False
        dtc_desc2.Visible = False
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False
        
        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
    End If

      If dtc_cuenta.Text = 3 Then
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False

        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
    End If
   
    If dtc_cuenta.Text = 4 Then
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = True
        dtc_desc3.Visible = True
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False
        
        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
        Call ABRIR_NIVEL4
    End If
    
    If dtc_cuenta.Text = 5 Then
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
'        dtc_desc2.Visible = True
        dtc_codigo3.Visible = True
'        dtc_desc3.Visible = True
        'dtc_codigo4.Visible = True
'        dtc_desc4.Visible = True
        
       
        Fra_Aux.Visible = True
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
        Call ABRIR_NIVEL4
        Call ABRIR_AUX1
        Call ABRIR_AUX2
        Call ABRIR_AUX3
        
    End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
 ' dtc_codigo8.Text = Ado_datos.Recordset!Cod_Aux1
     ' Ado_datos.Caption = rs_datos.AbsolutePosition & " / " & rs_datos.RecordCount
     If VAR_SW = "" Then
        Fra_ABM2.Visible = False
     Else
        Fra_ABM2.Visible = True
        
     End If
  End If
  Fra_ABM3.Enabled = True
  Fra_Aux.Enabled = True
  Fra_ABM1.Enabled = True
  
''  Call Abrir_Aux
'
''     If Ado_datos.Recordset!mov = "T" Then
'       ' lbl_observacion.Caption = "TITULO"
'        Fra_Aux.Visible = True
''
'
''        If Ado_datos.Recordset!mov = "S" Then
''           Fra_Aux.Visible = False
'           'lbl_observacion.Caption = "SUB TITULO"
''
'
'        var_cod = Ado_datos.Recordset!cuenta
'
'       If rs_aux2.State = 1 Then rs_aux2.Close
'              rs_aux2.Open "select  * from fo_balance_apertura where cuenta = '" & var_cod & "'  ", db, adOpenKeyset, adLockOptimistic
''              If rs_aux2.RecordCount > 0 Then
'''                txt_Tcta.Text = rs_aux2("Cuenta")
'''                txt_Tscta1.Text = rs_aux2("SubCta1")
'''                txt_Tscta2.Text = rs_aux2("SubCta2")
'''                txt_desc1.Text = rs_aux2("NombreCta")
''              End If
''
'        Else
'           Fra_Aux.Visible = True
'
'           'lbl_observacion.Caption = "DETALLE"
'            var_cod = Ado_datos.Recordset!cuenta
'            VAR_COD1 = Ado_datos.Recordset!subcta1
'            VAR_COD2 = Ado_datos.Recordset!subcta2
'
''
'     If rs_aux1.State = 1 Then rs_aux1.Close
'             rs_aux1.Open "select  * from fo_balance_apertura where  cuenta = '" & var_cod & "' and subcta1 = '" & VAR_COD1 & "'  ", db, adOpenKeyset, adLockOptimistic
'              If rs_aux1.RecordCount > 0 Then
''                txt_Tcta2.Text = rs_aux1("Cuenta")
''                txt_Tscta12.Text = rs_aux1("SubCta1")
''                txt_Tscta22.Text = rs_aux1("SubCta2")
''                txt_desc2.Text = rs_aux1("NombreCta")
'              End If
'
'              If rs_aux2.State = 1 Then rs_aux2.Close
'                   rs_aux2.Open "select  * from fo_balance_apertura where cuenta = '" & var_cod & "'  and subcta2 = '" & VAR_COD2 & "' ", db, adOpenKeyset, adLockOptimistic
'                   If rs_aux2.RecordCount > 0 Then
'    '                  txt_Tcta.Text = rs_aux2("Cuenta")
'    '                  txt_Tscta1.Text = rs_aux2("SubCta1")
'    '                  txt_Tscta2.Text = rs_aux2("SubCta2")
'    '                  txt_desc1.Text = rs_aux2("NombreCta")
'                   End If
'                   If Ado_datos.Recordset!aux1 = "00" Then
'    '                   Chkaux1.Value = 0
'                        dtc_codigo5.Visible = False
'                        dtc_desc5.Visible = False
''                       dtc_desc4.Visible = False
'                        Else
'            '               Chkaux1.Value = 1
'                            Call ABRIR_AUX1
'                            dtc_codigo5.Visible = True
'                            dtc_desc5.Visible = True
'                   End If
'                   If Ado_datos.Recordset!AUX2 = "00" Then
'                '           Chkaux2.Value = 0
'                            dtc_codigo6.Visible = False
'                            dtc_desc6.Visible = False
'                   Else
'            '               Chkaux2.Value = 1
'                            Call ABRIR_AUX2
'                            dtc_codigo6.Visible = True
'                            dtc_desc6.Visible = True
'                   End If
'                   If Ado_datos.Recordset!aux3 = "00" Then
'            '                Chkaux3.Value = 0
'                             dtc_codigo7.Visible = False
'                             dtc_desc7.Visible = False
'                   Else
'            '               Chkaux3.Value = 1
'                            Call ABRIR_AUX3
'                            dtc_codigo7.Visible = True
'                            dtc_desc7.Visible = True
'                   End If
'            End If

End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    Call OptFilGral2_Click
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    Fra_ABM2.Visible = True
    FraNavega.Enabled = False
    
    Fra_Aux.Enabled = True
   
    Buscar1.Visible = True
    Buscar2.Visible = True
    Buscar3.Visible = True
    Buscar1.Enabled = True
    Buscar2.Enabled = True
    Buscar3.Enabled = True
'   dg_datos.Enabled = False
    VAR_SW = "ADD"
    'dtc_cuenta.Text = "5"
'   txt_codigo.Enabled = False

   ' dtc_cuenta.SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(cuenta2 As String, scuenta1 As String, scuenta2 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM co_diario WHERE D_Cuenta = '" & cuenta2 & "' and D_Subcta1= '" & scuenta1 & "' and D_SubCta2= '" & scuenta2 & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_gestion_Click()

End Sub

Private Sub TxtDebe_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub TxtHaber_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

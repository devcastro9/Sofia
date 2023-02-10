VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_seguimiento_ventas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Ventas - Seguimiento de Contratos"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   17115
   Icon            =   "fw_seguimiento_ventas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   2.47341e7
   ScaleMode       =   0  'User
   ScaleWidth      =   1.34771e9
   WindowState     =   2  'Maximized
   Begin VB.Frame FraImprime 
      BackColor       =   &H80000018&
      Caption         =   "Reportes de a cuerdo a criterio seleccionado ..."
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
      Height          =   7455
      Left            =   8280
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton btnSalirPanel 
         Caption         =   "Salir"
         Height          =   495
         Left            =   2160
         TabIndex        =   39
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton btnPrintOption 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   480
         TabIndex        =   38
         Top             =   6600
         Width           =   1335
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H80000018&
         Caption         =   "9. KARDEX por COBRADOR seleccionado (Todos los Clientes) (Dólares)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   37
         Top             =   6000
         Width           =   7335
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H80000018&
         Caption         =   "8. KARDEX por SERVICIO seleccionado (Todos los Clientes) (Dólares)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   36
         Top             =   5520
         Width           =   7335
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H80000018&
         Caption         =   "7. KARDEX completo por Cliente seleccionado (TODOS los Servicios, con Tesorería) (Dólares)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   35
         Top             =   5040
         Width           =   8895
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000018&
         Caption         =   "6. KARDEX por COBRADOR seleccionado (Todos los Clientes)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   33
         Top             =   3960
         Width           =   6495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000018&
         Caption         =   "3. KARDEX del Cliente (Sólo del SERVICIO Seleccionado)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   32
         Top             =   1920
         Width           =   6495
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000018&
         Caption         =   "5. KARDEX por SERVICIO seleccionado (Todos los Clientes)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   29
         Top             =   3480
         Width           =   6495
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000018&
         Caption         =   "4. KARDEX completo por Cliente seleccionado (TODOS los Servicios, con Tesorería)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   3000
         Width           =   8055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000018&
         Caption         =   "2. KARDEX del Cliente (Sólo del CONTRATO Seleccionado)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Top             =   1440
         Width           =   6375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000018&
         Caption         =   "1. KARDEX completo por Cliente seleccionado (TODOS los Servicios)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   26
         Top             =   960
         Width           =   6735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         Caption         =   "Kardex para uso interno de CGI (Dolares)"
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
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   4560
         Width           =   7215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000018&
         Caption         =   "Kardex para uso interno de CGI"
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
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   2520
         Width           =   7215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Kardex para el Cliente"
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
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   600
         Width           =   6615
      End
   End
   Begin Crystal.CrystalReport CryQ01 
      Left            =   2280
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryF02 
      Left            =   1680
      Top             =   10320
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
   Begin Crystal.CrystalReport CryR01 
      Left            =   1080
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryF01 
      Left            =   480
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE COBRANZAS REGISTRADAS"
      ForeColor       =   &H00FF0000&
      Height          =   1635
      Left            =   120
      TabIndex        =   11
      Top             =   7800
      Width           =   18825
      Begin MSDataGridLib.DataGrid dg_datos06 
         Height          =   1260
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   18540
         _ExtentX        =   32703
         _ExtentY        =   2223
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "cobranza_detalle"
            Caption         =   "Nro."
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
            DataField       =   "cobranza_fecha"
            Caption         =   "Fecha.Cobro"
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
         BeginProperty Column03 
            DataField       =   "cmpbte_deposito"
            Caption         =   "Cmpbte.Depósito"
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
         BeginProperty Column04 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Recibo"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Cobrador"
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
         BeginProperty Column10 
            DataField       =   "estado_codigo"
            Caption         =   "Contabilizado"
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
            DataField       =   "trans_codigo"
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
         BeginProperty Column12 
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
               Alignment       =   2
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1769.953
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   7035.024
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   645.165
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos06 
         Height          =   330
         Left            =   75
         Top             =   1200
         Width           =   11460
         _ExtentX        =   20214
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
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   6
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnImprimir2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "fw_seguimiento_ventas.frx":058A
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   13
         ToolTipText     =   "Kardex por Cliente (de TODOS sus Contratos)"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3000
         Picture         =   "fw_seguimiento_ventas.frx":0EF3
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   10
         ToolTipText     =   "Kardex por Cliente (de TODOS sus Contratos)"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "fw_seguimiento_ventas.frx":197D
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   8
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7680
         Picture         =   "fw_seguimiento_ventas.frx":213F
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRONOGRAMA"
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
         Left            =   12855
         TabIndex        =   9
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.Frame FraNavega2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CUOTAS PARA COBRANZAS (PENDIENTES Y FACTURADOS)"
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
      Height          =   3480
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   18825
      Begin MSDataGridLib.DataGrid dg_datos2 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   18540
         _ExtentX        =   32703
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   33
         BeginProperty Column00 
            DataField       =   "cobranza_prog_codigo"
            Caption         =   "Nro.Cuota"
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
            DataField       =   "prog_fecha"
            Caption         =   "Fecha.Cuota"
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
            DataField       =   "prog_estado"
            Caption         =   "Estado.Cuota"
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
            DataField       =   "prog_bs"
            Caption         =   "Cuota.en.Bs"
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
            DataField       =   "prog_dol"
            Caption         =   "Cuota.en.Dol"
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
            DataField       =   "prog_observaciones"
            Caption         =   "Conepto.de.Cuota"
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
            DataField       =   "doc_numero2"
            Caption         =   "Certificado"
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
            DataField       =   "cobranza_fecha_conformidad"
            Caption         =   "Fecha.Conformidad"
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
            DataField       =   "cobranza_codigo"
            Caption         =   "Cod.Cobranza"
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
            DataField       =   "concepto_factura"
            Caption         =   "Concepto Factura"
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
            DataField       =   "tipo_documento"
            Caption         =   "Doc.ISO"
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
            DataField       =   "cobranza_fecha_fac"
            Caption         =   "F.Facturacion"
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
            DataField       =   "numero_recibo"
            Caption         =   "Nro.Recibo"
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
            DataField       =   "numero_factura"
            Caption         =   "Nro.Factura"
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
            DataField       =   "doc_numero"
            Caption         =   "Nro.O.C."
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
            DataField       =   "cobranza_nro_factura"
            Caption         =   "Nro. Factura"
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
            DataField       =   "monto_facturado"
            Caption         =   "Monto Facturado"
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
            DataField       =   "beneficiario_nit"
            Caption         =   "NIT Facturado"
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
            DataField       =   "beneficiario_denominacion_fac"
            Caption         =   "Factura a Nombre de:"
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
            DataField       =   "estado_facturacion"
            Caption         =   "Estado Fac."
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
            DataField       =   "cobro1_cuenta"
            Caption         =   "1ra.Cta.Bancaria"
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
         BeginProperty Column21 
            DataField       =   "cobro1_comprobante"
            Caption         =   "1er.Cmpbte.Deposito"
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
         BeginProperty Column22 
            DataField       =   "cobranza_deuda_bs"
            Caption         =   "Tot.Cobrado.Bs."
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
         BeginProperty Column23 
            DataField       =   "cobranza_deuda_dol"
            Caption         =   "Tot.Cobrado.Dol."
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
         BeginProperty Column24 
            DataField       =   "cobro1_fecha"
            Caption         =   "1er.Fecha.Cobro"
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
         BeginProperty Column25 
            DataField       =   "cobro2_cuenta"
            Caption         =   "2da.Cta.Bancaria"
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
         BeginProperty Column26 
            DataField       =   "saldo_bs"
            Caption         =   "Saldo.X.Cobrar.Bs"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column27 
            DataField       =   "saldo_dol"
            Caption         =   "Saldo.X.Cobrar.Dol"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column28 
            DataField       =   "cobro2_fecha"
            Caption         =   "2da.Fecha.Cobro"
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
         BeginProperty Column29 
            DataField       =   "cobrador"
            Caption         =   "Cobrador"
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
         BeginProperty Column30 
            DataField       =   "estado_cobranza"
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
         BeginProperty Column31 
            DataField       =   "beneficiario_codigo"
            Caption         =   "NIT/CI del Cliente"
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
         BeginProperty Column32 
            DataField       =   "estado_codigo"
            Caption         =   "Contabilizado"
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
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1590.236
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               DividerStyle    =   1
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1769.953
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column15 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   1874.835
            EndProperty
            BeginProperty Column22 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1590.236
            EndProperty
            BeginProperty Column23 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1649.764
            EndProperty
            BeginProperty Column24 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column25 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column26 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1649.764
            EndProperty
            BeginProperty Column27 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1755.213
            EndProperty
            BeginProperty Column28 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column29 
               Alignment       =   2
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column30 
               Alignment       =   2
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column31 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column32 
               Alignment       =   2
               ColumnWidth     =   1244.976
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos02 
         Height          =   330
         Left            =   75
         Top             =   3195
         Visible         =   0   'False
         Width           =   18540
         _ExtentX        =   32703
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
   Begin Crystal.CrystalReport CryQ02 
      Left            =   2880
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DATOS DE CONTRATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   18855
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   2880
         Width           =   17370
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
            Left            =   5160
            TabIndex        =   24
            Top             =   120
            Width           =   1550
         End
         Begin VB.OptionButton OptFilGral04 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "4. Oruro . . ."
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
            Left            =   7200
            TabIndex        =   23
            Top             =   120
            Width           =   1305
         End
         Begin VB.OptionButton OptFilGral00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Todos . . . ."
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
            TabIndex        =   22
            Top             =   120
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton OptFilGral02 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "2. La Paz . . ."
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
            Left            =   3600
            TabIndex        =   21
            Top             =   120
            Width           =   1305
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
            Left            =   1680
            TabIndex        =   20
            Top             =   120
            Width           =   1545
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
            Left            =   8760
            TabIndex        =   19
            Top             =   120
            Width           =   1425
         End
         Begin VB.OptionButton OptFilGral06 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "6. Tarija . . ."
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
            Left            =   10560
            TabIndex        =   18
            Top             =   120
            Width           =   1185
         End
         Begin VB.OptionButton OptFilGral09 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "9. Pando . . ."
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
            Left            =   15600
            TabIndex        =   17
            Top             =   120
            Width           =   1305
         End
         Begin VB.OptionButton OptFilGral08 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "8. Beni . . ."
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
            TabIndex        =   16
            Top             =   120
            Width           =   1185
         End
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
            TabIndex        =   15
            Top             =   120
            Width           =   1425
         End
      End
      Begin MSDataGridLib.DataGrid dg_datos16 
         Height          =   2610
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   18600
         _ExtentX        =   32808
         _ExtentY        =   4604
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
         ColumnCount     =   26
         BeginProperty Column00 
            DataField       =   "depto_codigo_vta"
            Caption         =   "Depto."
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
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad/Oficina"
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
            DataField       =   "estado_cancelado"
            Caption         =   "Vigente(S)-NoVigente(N)"
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
            DataField       =   "edif_codigo_corto"
            Caption         =   "Cod.Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Denominacion del Edificio"
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
            DataField       =   "venta_fecha_inicio"
            Caption         =   "Fecha.Inicio"
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
            DataField       =   "venta_fecha_fin"
            Caption         =   "Fecha.Fin"
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
            DataField       =   "venta_monto_total_bs"
            Caption         =   "Total.Contrato.Bs"
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
            DataField       =   "venta_monto_facturado_bs"
            Caption         =   "Tot.Facturado.Bs"
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
            DataField       =   "venta_saldo_p_facturar_bs"
            Caption         =   "Saldo.P/Facturar"
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
            DataField       =   "venta_monto_cobrado_bs"
            Caption         =   "Tot.Cobrado.Bs"
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
            DataField       =   "venta_saldo_p_cobrar_bs"
            Caption         =   "Saldo.P/Cobar"
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
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "estado_descripcion"
            Caption         =   "Tramite"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cod.Adm./Contrato"
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
            DataField       =   "subproceso_descripcion"
            Caption         =   "Proceso"
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
            DataField       =   "unimed_codigo"
            Caption         =   "Periodicidad"
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
            DataField       =   "venta_cantidad_total"
            Caption         =   "Cantidad/Cuotas"
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
            DataField       =   "zona_denominacion"
            Caption         =   "Zona"
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
         BeginProperty Column19 
            DataField       =   "calle_tipo"
            Caption         =   "Via.Acceso"
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
         BeginProperty Column20 
            DataField       =   "calle_denominacion"
            Caption         =   "Nombre de Calle, Av u otro"
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
         BeginProperty Column21 
            DataField       =   "edif_nro"
            Caption         =   "Nro."
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
         BeginProperty Column22 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
         BeginProperty Column23 
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Cliente/Representante.Legal"
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
         BeginProperty Column24 
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad.E."
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
         BeginProperty Column25 
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Tramite"
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
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   -1  'True
               ColumnWidth     =   4529.764
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column14 
               Alignment       =   2
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   2055.118
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   2385.071
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   464.882
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column23 
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos16 
         Height          =   330
         Left            =   120
         Top             =   2925
         Width           =   18585
         _ExtentX        =   32782
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
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00808080&
      Caption         =   "DETALLE DE BIENES / SERVICIOS VENDIDOS"
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
      Height          =   1485
      Left            =   120
      TabIndex        =   0
      Top             =   9525
      Visible         =   0   'False
      Width           =   18855
      Begin MSDataGridLib.DataGrid DtGLista 
         Height          =   1140
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   18615
         _ExtentX        =   32835
         _ExtentY        =   2011
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Bien"
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
            DataField       =   "concepto_venta"
            Caption         =   "Descripcion y Características del Bien"
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
            DataField       =   "venta_det_cantidad"
            Caption         =   "Cantidad"
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
         BeginProperty Column04 
            DataField       =   "venta_precio_unitario_bs"
            Caption         =   "Prec.Unitario"
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
         BeginProperty Column05 
            DataField       =   "venta_descuento_bs"
            Caption         =   "Descuento"
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
         BeginProperty Column06 
            DataField       =   "venta_precio_total_bs"
            Caption         =   "Precio Total"
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
         BeginProperty Column07 
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Vendido"
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
            DataField       =   "almacen_codigo"
            Caption         =   "Almacen"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   5009.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   720
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6840
      Top             =   10080
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
      Top             =   10080
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
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   0
      Top             =   10800
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
   Begin MSAdodcLib.Adodc ado_datos17 
      Height          =   330
      Left            =   9120
      Top             =   10440
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
      Caption         =   "ado_datos17"
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
      Left            =   -120
      Top             =   10440
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
   Begin MSAdodcLib.Adodc ado_datos15 
      Height          =   330
      Left            =   6840
      Top             =   10440
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
      Caption         =   "ado_datos15"
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
   Begin MSAdodcLib.Adodc AdoDsctos 
      Height          =   330
      Left            =   11400
      Top             =   10080
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
      Caption         =   "AdoDsctos"
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2280
      Top             =   10440
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
      Caption         =   "Ado_Datos12"
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
   Begin MSAdodcLib.Adodc Ado_datos13 
      Height          =   330
      Left            =   4560
      Top             =   10440
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
      Caption         =   "Ado_datos13"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   13680
      Top             =   10080
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
      Caption         =   "AdoAux"
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
      Left            =   4560
      Top             =   10080
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   10080
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
   Begin MSAdodcLib.Adodc ado_datos4A 
      Height          =   330
      Left            =   9120
      Top             =   10080
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
      Caption         =   "ado_datos4A"
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
   Begin MSAdodcLib.Adodc Ado_datos20 
      Height          =   330
      Left            =   4560
      Top             =   10800
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
      Caption         =   "Ado_datos20"
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
      Left            =   6840
      Top             =   10800
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   9120
      Top             =   10800
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
      Left            =   11400
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   13680
      Top             =   10440
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
End
Attribute VB_Name = "fw_seguimiento_ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ventas
Dim rs_datos As New ADODB.Recordset     'FACTURACION
Dim rs_datos01 As New ADODB.Recordset     'INICIO COBRANZAS
Dim rs_datos02 As New ADODB.Recordset     'REG. COBRANZAS
Dim rs_datos06 As New ADODB.Recordset     'DETALLE COBRANZAS
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos4A As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset
Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset
Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
Dim rs_datos20 As New ADODB.Recordset   'Cta Bancaria

Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset

'CLASIFICADORES
Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset
'IMAGENES
Dim m_stream    As ADODB.Stream
'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
'Dim queryinicial As String
Dim queryinicial1 As String
Dim queryinicial2 As String

'Dim descri_bien As String
'VARIABLES
Dim iResult As Variant  ', i%, y%
Dim marca1 As Variant

Dim VAR_CANT As Integer         'Cant_Alm,
Dim correlativo1 As Integer
Dim swgrabar, swnuevo, deta2 As Integer
Dim nroventa, correlv, NRO_COBR As Long
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SW, VAR_TSOL As Integer
Dim VAR_SOL As Integer
Dim i As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, COBR_BS As Double
Dim VAR_CONTAB As Double
Dim Monto_Bs As Double

Dim gestion0, var_literal, VAR_PROY2, VAR_CTA, VAR_PROY3 As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_MONEDA As String
Dim VAR_COD1, VAR_COD2, VAR_COD3 As String
Dim VAR_ANIO, VAR_MES, VAR_DIA, VAR_FECHA As String
Dim VAR_COD4, VAR_TIPOV, VAR_CITE  As String
Dim DESAUX, VARAUX, VARCODIG As String
Dim Numero As String
Dim Autorizacion As String
Dim NroFactura As String
Dim NitCi As String
Dim Fecha As String
Dim Monto As String
Dim Llave As String
Dim CodigoContro As String
Dim montoLiteral As String
        
'Dim Exel As New Excel.Application
Dim fs As FileSystemObject      'Variable de tipo file System Object
    
Private Sub CmdDetalle_Click()
    FrmCobranza.Visible = True
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  ' Verifica si es factura
  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
  If Not IsNull(Ado_datos.Recordset("doc_codigo_fac")) Then
    If Ado_datos.Recordset("doc_codigo_fac") = "R-101" Then
'      BtnImprimir3.Visible = True
    Else
'      BtnImprimir3.Visible = False
    End If
  End If
  End If
    
'  Call cambiarEtiquetaFactura
'  Dim descri_bien As String
'  Dim Cant_Alm As Integer
  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
            BtnModificar.Visible = True
         
            If Ado_datos.Recordset!doc_codigo_fac <> "R-101" Then
                'BtnImprimir3.Caption = "Recibo"
                lbl_factura.Caption = "Nro.de Recibo"
            Else
                'BtnImprimir3.Caption = "Facturar"
                lbl_factura.Caption = "Nro.de Factura"
            End If
            If (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 16) Then
                TxtDsctoTot.backColor = &HFF&             'ROJO
                DTPFechaProg.backColor = &HFF&             'ROJO
            Else
                If (Ado_datos.Recordset("cobranza_fecha_sol") > Date - 16) And (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 1) Then
                    TxtDsctoTot.backColor = &H80FF&           'NARANJA
                    DTPFechaProg.backColor = &H80FF&           'NARANJA
                Else
                    TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
                    DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
                End If
            End If
        Else
            BtnModificar.Visible = False
'            BtnEliminar.Visible = False
'            BtnAprobar.Visible = False
'            BtnVer.Visible = True
'            FrmABMDet.Visible = False
'            FrmABMDet2.Visible = True
'            FrmCobranza.Visible = True
            TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
'            BtnImprimir3.Visible = False
        End If

        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and correl_venta = " & Ado_datos.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
        Else
            deta2 = 0
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
'    BtnImprimir3.Visible = False
'                BtnDesAprobar.Visible = True
    BtnModificar.Visible = False
'    BtnEliminar.Visible = False
'    BtnVer.Visible = False
    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF
End Sub

Private Sub Ado_datos01_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not Ado_datos01.Recordset.BOF) And (Not Ado_datos01.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos01.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos01.Recordset("estado_codigo_sol") = "REG") Then          'REG
            If (Ado_datos01.Recordset("cobranza_fecha_prog") <= Date - 16) Then
                TxtDsctoTot1.backColor = &HFF&             'ROJO
                DTPFechaProg1.backColor = &HFF&             'ROJO
            Else
                If (Ado_datos01.Recordset("cobranza_fecha_prog") > Date - 16) And (Ado_datos01.Recordset("cobranza_fecha_prog") <= Date - 1) Then
                    TxtDsctoTot1.backColor = &H80FF&           'NARANJA
                    DTPFechaProg1.backColor = &H80FF&           'NARANJA
                Else
                    TxtDsctoTot1.backColor = &H404040        '&H80000013      'Fondo Oscuro
                    DTPFechaProg1.backColor = &H404040       '&H80000013      'Fondo Oscuro
                End If
            End If
            BtnModificar1.Visible = True
            BtnAprobar1.Visible = True
            If Ado_datos01.Recordset!doc_codigo_fac = "R-103" Then
                cmd_fac = "RECIBO"
            Else
                cmd_fac = "FACTURA"
            End If
        Else
            BtnModificar1.Visible = False
            BtnAprobar1.Visible = False
            TxtDsctoTot1.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg1.backColor = &H404040       '&H80000013      'Fondo Oscuro
        End If
'        If Ado_datos01.Recordset("beneficiario_codigo") <> "" Then
'            Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos01.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'            'RS_BENEF.Recordset.Requery
'            If RS_BENEF.RecordCount > 0 Then
'                If RS_BENEF!beneficiario_deudor = "SI" Then
'                    Dtc_deudor2.BackColor = &HFF&
'                Else
'                    Dtc_deudor2.BackColor = &H80000010
'                End If
'            End If
'
'        End If
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos01.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos01.Recordset!venta_codigo & " and correl_venta = " & Ado_datos01.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            'TxtMontoBs.Text = Ado_datos01.Recordset!monto_total_bS
            'TxtMontoUs.Text = Ado_datos01.Recordset!deuda_cobrada
            'Text2.Text = Ado_datos01.Recordset!saldo_p_cobrar
            'Call AbreAlmacen
        Else
            deta2 = 0
'            'TxtMontoBs.Text = 0
'            'TxtMontoUs.Text = 0
'            'Text2.Text = 0
'            FrmABMDet2.Visible = False
'            FrmCobranza.Visible = False
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos01.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos01.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos01.Recordset("venta_codigo")))
        
'        TxtCobrador1 = Trim(dtc_desc4A.Text)
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos01.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
    BtnAprobar1.Visible = False
    BtnModificar1.Visible = False
    'BtnEliminar1.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF
End Sub


Private Sub AbreAlmacen()
'    Set rs_datos13 = New ADODB.Recordset
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    'rs_datos13.Open "select * from Av_DestinoDet where coddetalle= '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    rs_datos13.Open "select * from Av_almacen_detalle where bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_datos13.Recordset = rs_datos13
'    Ado_datos13.Refresh

End Sub

Private Sub Ado_datos02_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  ' Verifica si no es factura
  
      
  If (Not Ado_datos02.Recordset.BOF) And (Not Ado_datos02.Recordset.EOF) Then   'EOF
  
     If Not IsNull(Ado_datos02.Recordset("venta_codigo")) Then            'venta_codigo
        NRO_COBR = IIf(IsNull(Ado_datos02.Recordset!cobranza_codigo), 0, Ado_datos02.Recordset!cobranza_codigo)
        If (Ado_datos02.Recordset("estado_codigo_bco") = "REG") Then          'REG
          
        Else
'           ' TxtDsctoTot2.backColor = &H404040        '&H80000013      'Fondo Oscuro
'            'DTPFechaProg2.backColor = &H404040       '&H80000013      'Fondo Oscuro
'            If Ado_datos02.Recordset!estado_codigo = "APR" Then
'
'                OptFilGral05.Visible = False
'            Else
'                If (glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS") Then
'
'                    OptFilGral05.Visible = True
'                Else
'
'                    OptFilGral05.Visible = False
'                End If
'            End If
        End If
        Call OptFilGral10_Click         'Cobranza Det

'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos02.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos02.Recordset!venta_codigo & " and correl_venta = " & Ado_datos02.Recordset!correl_venta & " "
'        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'        Set ado_datos14.Recordset = rs_datos14
'        ado_datos14.Recordset.Requery
'        If ado_datos14.Recordset.RecordCount > 0 Then
'            deta2 = 1
'            'TxtMontoBs.Text = Ado_datos02.Recordset!monto_total_bS
'            'TxtMontoUs.Text = Ado_datos02.Recordset!deuda_cobrada
'            'Text2.Text = Ado_datos02.Recordset!saldo_p_cobrar
'            'Call AbreAlmacen
'        Else
'            deta2 = 0
''            'TxtMontoBs.Text = 0
''            'TxtMontoUs.Text = 0
''            'Text2.Text = 0
''            FrmABMDet2.Visible = False
''            FrmCobranza.Visible = False
'        End If
        
'        Set rs_datos16 = New ADODB.Recordset
'        If rs_datos16.State = 1 Then rs_datos16.Close
'        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos02.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_datos16.Recordset = rs_datos16
'        Ado_datos16.Recordset.Requery
'        If Ado_datos16.Recordset.RecordCount > 0 Then
'            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
'            FrmCobranza.Visible = True
'            'BtnImprimir2.Visible = True
'            'BtnImprimir3.Visible = True
'        Else
'            FrmCobranza.Visible = False
'            'BtnImprimir2.Visible = False
'            'BtnImprimir3.Visible = False
'        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
'        Set rs_datos5 = New ADODB.Recordset
'        If rs_datos5.State = 1 Then rs_datos5.Close
'        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
'        Set Ado_datos5.Recordset = rs_datos5
       ' dtc_desc5.BoundText = dtc_codigo5.BoundText
        'dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        'FrmDetalle.Caption = "VENTA NRO. " + str((Ado_datos02.Recordset("venta_codigo")))
        
        'FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + str((Ado_datos02.Recordset("venta_codigo")))
        
'        TxtCobrador1 = Trim(dtc_desc4A.Text)
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos02.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     'FrmDetalle.Enabled = True
     'FrmCobranza.Visible = True
  Else
    'BtnAprobar2.Visible = False
    'BtnModificar2.Visible = False
    'BtnEliminar2.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    'FrmABMDet.Visible = False
    'FrmABMDet2.Visible = False
  End If                            'EOF

End Sub

Private Sub BtnAñadir_Click()
marca1 = Ado_datos.Recordset.Bookmark
  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Then
    If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
        swnuevo = 1
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
'        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
'        TxtCobrador.Visible = False
        Ado_datos16.Recordset.AddNew
        dtc_codigo2A.Text = dtc_codigo2.Text
        dtc_desc2A.Text = dtc_desc2.Text
        TxtMonto.SetFocus
        DTPFechaProg.Visible = True
        DTPFechaCobro.Visible = True
        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
        Txt_parche.Visible = True
        'Ado_datos.Recordset.Move marca1 - 1
    Else
        MsgBox "Ya se cobró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "La Venta (al Contado o Donación) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atención!"
  End If
End Sub


'Private Sub Ado_datos16_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then   'EOF
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos16.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
        Else
            deta2 = 0
        End If

        Set rs_datos02 = New Recordset
        If rs_datos02.State = 1 Then rs_datos02.Close
        queryinicial2 = " SELECT * FROM av_seguimiento_prog_cobranza where venta_codigo = '" & Ado_datos16.Recordset!venta_codigo & "' "
        rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        rs_datos02.Sort = "cobranza_prog_codigo"
        Set Ado_datos02.Recordset = rs_datos02.DataSource
        Set dg_datos2.DataSource = Ado_datos02.Recordset


'        Set rs_datos5 = New ADODB.Recordset
'        If rs_datos5.State = 1 Then rs_datos5.Close
'        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
'        Set Ado_datos5.Recordset = rs_datos5
Else

End If

End Sub

Private Sub BtnBuscar_Click()
'JQA
 If Ado_datos16.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos16
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos16.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If



' If Ado_datos.Recordset.RecordCount > 0 Then
'    'JQA
'    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
'      PosibleApliqueFiltro = False
'      Dim rsNada As ADODB.Recordset
'      Dim GrSqlAux As String
'      Set ClBuscaGrid = New ClBuscaEnGridExterno
'      Set ClBuscaGrid.Conexión = db
'      ClBuscaGrid.EsTdbGrid = False
'      Set ClBuscaGrid.GridTrabajo = dg_datos
'      ClBuscaGrid.QueryUtilizado = queryinicial1
'      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
'      ClBuscaGrid.CamposVisibles = "110"
'      ClBuscaGrid.Ejecutar
'      PosibleApliqueFiltro = True
'  Else
'    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
End Sub

'Private Sub BtnBuscar1_Click()
''JQA
' If Ado_datos01.Recordset.RecordCount > 0 Then
'    'JQA
'    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
'      PosibleApliqueFiltro = False
'      Dim rsNada As ADODB.Recordset
'      Dim GrSqlAux As String
'      Set ClBuscaGrid = New ClBuscaEnGridExterno
'      Set ClBuscaGrid.Conexión = db
'      ClBuscaGrid.EsTdbGrid = False
'      Set ClBuscaGrid.GridTrabajo = dg_datos1
'      ClBuscaGrid.QueryUtilizado = queryinicial
'      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos01.Recordset
'      ClBuscaGrid.CamposVisibles = "110"
'      ClBuscaGrid.Ejecutar
'      PosibleApliqueFiltro = True
'  Else
'    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'
'End Sub

'Private Sub BtnBuscar2_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'    'JQA
'    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
'      PosibleApliqueFiltro = False
'      Dim rsNada As ADODB.Recordset
'      Dim GrSqlAux As String
'      Set ClBuscaGrid = New ClBuscaEnGridExterno
'      Set ClBuscaGrid.Conexión = db
'      ClBuscaGrid.EsTdbGrid = False
'      Set ClBuscaGrid.GridTrabajo = dg_datos2
'      ClBuscaGrid.QueryUtilizado = queryinicial2
'      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos02.Recordset
'      ClBuscaGrid.CamposVisibles = "110"
'      ClBuscaGrid.Ejecutar
'      PosibleApliqueFiltro = True
'  Else
'    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'
'End Sub

Private Sub btnEliminar_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    If Ado_datos.Recordset!estado_codigo_fac = "APR" And Ado_datos.Recordset!estado_codigo_bco = "REG" Then      'Ado_datos.Recordset("estado_codigo_anl") = "REG"
'      sino = MsgBox("Esta seguro de ANULAR la facturación registrada ?", vbYesNo, "Confirmando")
'      If sino = vbYes Then
'        sino = MsgBox("Volverá a emitir otra FACTURA con este mismo registro ? (Si elige NO, se cierra el registro)", vbYesNo, "Confirmando")
'        If sino = vbYes Then
'          db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'REG' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set factura_impresa = 'N' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'        Else
'          db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'ANL' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'        End If
'          db.Execute "update ao_ventas_cobranza set cobranza_nro_factura_anl = '" & Ado_datos.Recordset!cobranza_nro_factura & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_fecha_anl = '" & Format(Date, "dd/mm/yyyy") & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set usr_codigo_anl = '" & glusuario & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set estado_codigo_anl = 'APR' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_fecha_ant = cobranza_fecha_fac Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_codigo_control_anl = cobranza_codigo_control Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set correl_contab_anl = correl_contab Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion_anl = cobranza_nro_autorizacion Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'
'          Set rs_datos12 = New ADODB.Recordset
'          If rs_datos12.State = 1 Then rs_datos12.Close
'          rs_datos12.Open "Select * from ao_ventas_cobro_anl where cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " and cobranza_nro_factura_anl = " & Ado_datos.Recordset!cobranza_nro_factura & " ", db, adOpenKeyset, adLockOptimistic
'          If rs_datos12.RecordCount > 0 Then
'            MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
'          Else
'            'wwwwwwwwwwwwwwwwwwwww
'              ' hora_registro
'            rs_datos12.AddNew
'            rs_datos12!ges_gestion = glGestion
'            rs_datos12!cobranza_codigo = Ado_datos.Recordset!cobranza_codigo
'            rs_datos12!venta_codigo = Ado_datos.Recordset!venta_codigo
'
'            rs_datos12!cobranza_nro_factura_anl = Ado_datos.Recordset!cobranza_nro_factura
'            rs_datos12!cobranza_prog_codigo = Ado_datos.Recordset!cobranza_prog_codigo
'            rs_datos12!beneficiario_codigo_fac = Ado_datos.Recordset!beneficiario_codigo_fac
'            rs_datos12!cobranza_anuladal_bs = Ado_datos.Recordset!cobranza_total_bs
'            rs_datos12!cobranza_anulada_dol = Ado_datos.Recordset!cobranza_total_dol
'
'            rs_datos12!cobranza_fecha_anl = Ado_datos.Recordset!cobranza_fecha_fac      'Format(Date, "dd/mm/yyyy")
'            rs_datos12!cobranza_fecha_fac2 = Ado_datos.Recordset!cobranza_fecha_fac2
'            rs_datos12!cobranza_observaciones = Ado_datos.Recordset!cobranza_observaciones
'            rs_datos12!cobranza_codigo_control_anl = Ado_datos.Recordset!cobranza_codigo_control
'            rs_datos12!Literal = Ado_datos.Recordset!Literal
'
'            rs_datos12!cobranza_nro_autorizacion_anl = Ado_datos.Recordset!cobranza_nro_autorizacion
'            rs_datos12!correl_contab_anl = Ado_datos.Recordset!correl_contab
'            rs_datos12!estado_codigo_anl = "APR"            'Ado_datos.Recordset!estado_codigo_anl
'            rs_datos12!usr_codigo_anl = glusuario           'Ado_datos.Recordset!usr_codigo_anl
'            rs_datos12!fecha_registro = Ado_datos.Recordset!fecha_registro
'
'            rs_datos12!trans_codigo = Ado_datos.Recordset!trans_codigo
'            rs_datos12!cmpbte_deposito = Ado_datos.Recordset!cmpbte_deposito
'            rs_datos12!cta_codigo = Ado_datos.Recordset!cta_codigo
'            rs_datos12.Update
'          End If
'      End If
'        '  rs_datos12!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
'          'wwwwwwwwwwwwwwwwwwwww
'          'marca1 = Ado_datos.Recordset.Bookmark
'          'Call OptFilGral2_Click
'          'Ado_datos.Recordset.Move marca1 - 1
'    Else
'      MsgBox "NO se puede ANULAR, porque el registro NO fue Facturado o ya fue Cobrado...", , "Atencion"
'    End If
'  Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
'  End If
End Sub


Private Sub BtnGrabar_Click()
  Call cambiarEtiquetaFactura
  If dtc_codigo4A.Text = "" Then
    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If dtc_codigo5.Text = "" Then
    MsgBox "Debe Elejir <<Factura a Nombre de:>> !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If TxtMonto = "" Or TxtMonto = "0" Or TxtMonto = "0.00" Then
    MsgBox "Debe Registrar el " + lbl_monto.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If TxtObs = "" Then
    MsgBox "Debe Registrar el " + lbl_obs.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  'If swnuevo = 2 Then
  'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'  If DTPFechaProg.Visible = False Then
'    If TxtCmpbte = "" Or TxtCmpbte = "0" Then
'       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'      Exit Sub
'    End If
'  End If
  'fin PARA COBRANZA WWWWWWWWWWWWWWWWWWW
  'valida = 1
  'If valida = 1 And dtc_codigo4A <> "" Then
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
    db.BeginTrans
    If swnuevo = 1 Then
'      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
'      Set Ado_datos16.Recordset = rstdestino
'      Ado_datos16.Recordset.AddNew
      Ado_datos.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
      Ado_datos.Recordset!ges_gestion = glGestion       'Ado_datos.Recordset("ges_gestion")
      'Ado_datos.Recordset!cobranza_fecha_prog = DTPFechaProg                                'Fecha Programada a Cobrar
    End If
'      If Ado_datos.Recordset!beneficiario_codigo = "0" Then
'        Ado_datos.Recordset!beneficiario_codigo = dtc_codigo5.Text        'lbl_nit.Caption                                  'Codigo Beneficiario (Cliente)
'      End If
      Ado_datos.Recordset!beneficiario_codigo_fac = IIf(dtc_codigo5.Text = "", "0", dtc_codigo5.Text)       ' dtc_codigo5.Text  'dtc_codigo2A.Text                            'Beneficiario (Factura a nombre de ...)
      Ado_datos.Recordset!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
      Ado_datos.Recordset!trans_codigo = IIf(dtc_codigo6.Text = "", "O", dtc_codigo6.Text) 'tipo de Transaccion
      'Ado_datos.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
      Ado_datos.Recordset!cmpbte_deposito = IIf(Txt_deposito.Text = "", "0", Txt_deposito.Text)
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      Ado_datos.Recordset!cta_codigo = IIf(dtc_cta.Text = "", "NN", dtc_cta.Text)
      Ado_datos.Recordset!cta_codigo2 = IIf(dtc_codigo7.Text = "", "NN", dtc_codigo7.Text)
      If TxtMonto.Text = "" Then
        Ado_datos.Recordset!cobranza_deuda_bs = "0"                                  'Monto Cobrado Bs.
        Ado_datos.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
      Else
        Ado_datos.Recordset!cobranza_tdc = IIf(IsNull(Txt_tdc = ""), 6.96, CDbl(Txt_tdc.Text))                               'Monto Cobrado Bs.
'        Ado_datos.Recordset!cobranza_deuda_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
        Ado_datos.Recordset!cobranza_total_bs = CDbl(TxtMonto.Text)                                  'Monto Cobrado Bs.
        Ado_datos.Recordset!cobranza_total_dol = CDbl(TxtMontoDol)        'CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
      End If
      'VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) + " - Nro.: " + Trim(VAR_CITE)
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos.Recordset!cobranza_total_bs <> 0 Then
            Ado_datos.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
      End If
      'Ado_datos.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value                                'Fecha de Cobranza
      'Call acumulaMont(Ado_datos.Recordset!ges_gestion, Ado_datos.Recordset!correl_venta, Ado_datos.Recordset!venta_codigo)
      Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
      '        '===== ini GENERA NRO. AUTORIZACION DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FAC_AUTORIZA'", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'          VAR_COD2 = CDbl(rs_aux1!numero_correlativo)
'          'rs_aux1!numero_correlativo = Trim(Str(VAR_COD2))
'          'rs_aux1.Update
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION NRO. AUTORIZACION DE FACTURA =====
'        'GENERA CORREL NOTA DEBITO POR DEPTO INI
'        Set rs_aux5 = New ADODB.Recordset
'        If rs_aux5.State = 1 Then rs_aux5.Close
'        'rs_aux5.Open "Select correl_contab as Codigo from gc_departamento where depto_codigo = '" & Left(VAR_PROY3, 1) & "'    ", db, adOpenStatic
'        rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
'        If Not rs_aux5.EOF Then
'            VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
'        End If
'        'rs_aux5!Codigo = VAR_CONTAB
'        'rs_aux5.Update
'        db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'        db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
'        'Ado_datos.Recordset!correl_contab = VAR_CONTAB
'        'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'        If VAR_CONTAB < 10 Then
'            Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'        End If
'        If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
'           Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'        End If
'        If VAR_CONTAB > 99 And VAR_CONTAB < 1000 Then
'           Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'        End If
      'Ado_datos.Recordset!proceso_codigo = "FIN"
      'Ado_datos.Recordset!subproceso_codigo = "FIN-01"
      Ado_datos.Recordset!etapa_codigo = "FIN-01-02"
      'Ado_datos.Recordset!clasif_codigo = "ADM"
      'Ado_datos.Recordset!doc_codigo = IIf(lbl_doc1 = "", "R-105", lbl_doc1)
      'Ado_datos.Recordset!doc_numero = IIf(lbl_docnro = "", "0", lbl_docnro)
      Ado_datos.Recordset!cmpbte_deposito = IIf(Txt_deposito = "", "0", Txt_deposito)
      If lbl_fac = "R-103" Then
        Ado_datos01.Recordset!doc_codigo_fac = "R-103"
      Else
        Ado_datos01.Recordset!doc_codigo_fac = "R-101"
      End If
      If Ado_datos.Recordset!factura_impresa = "N" And Ado_datos01.Recordset!doc_codigo_fac = "R-101" Then
         TxtCmpbte.Text = "0"
         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
      Else
         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
      End If
      Ado_datos.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion = "", "0", Trim(TxtAutorizacion))
      Ado_datos.Recordset!poa_codigo = "3.1.2"
      Ado_datos.Recordset!cobranza_fecha_fac = DTPFechaCobro.Value         'Fecha de Facturacion
        'VAR_ANIO = CStr(glGestion)
        'VAR_MES = CStr(Month(Date))
        'VAR_DIA = CStr(Day(Date))
      Ado_datos.Recordset!cobranza_fecha_fac2 = ""        'VAR_ANIO & VAR_MES & VAR_DIA          'Fecha de Facturacion Texto
      Ado_datos.Recordset!estado_codigo_fac = "REG"
      Ado_datos.Recordset!usr_codigo = glusuario
      Ado_datos.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      Ado_datos.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos.Recordset.Update
    db.CommitTrans
    
      MsgBox "El registro se guardo correctamente"
    'Ado_datos.Recordset!doc_numero = Ado_datos.Recordset!cobranza_codigo       'Txt_cod_cobro.Text     ' "0"
  If swnuevo = 1 Then
    'Call abre_solicitud_lista
    'rc_Cobranza.Requery
    'Ado_datos.Refresh
    'Ado_datos.Recordset.MoveLast
  End If
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    FraNavega.Enabled = True
'    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FrmDetalle.Enabled = True
    FrmCobranza.Visible = True
    FrmCobros.Enabled = False
'    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
'    BtnImprimir2.Visible = True
'    BtnImprimir3.Visible = True
    
    swnuevo = 0
    
  'Else
  '  MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
  'End If

End Sub

Private Sub BtnImprimir_Click()
    FraImprime.Visible = True
    fraOpciones.Visible = False
    FrmDetalle.Enabled = False
    FraNavega2.Enabled = False
    Frame1.Enabled = False
    
'  If Ado_datos16.Recordset.RecordCount > 0 Then
'  '            Dim iResult As Variant  ', i%, y%
'      CryF01.ReportFileName = App.Path & "\reportes\ventas\fr_contrato_kardex_cliente.rpt"
'      CryF01.WindowShowRefreshBtn = True
'        If Ado_datos16.Recordset!edif_codigo = "" Then
'              CryF01.StoredProcParam(0) = "%"
'        Else
'              'CryF01.StoredProcParam(0) = cmb_codigoedificio.Text ' para REPORTE POR OPCIONES (PENDIENTE)
'              CryF01.StoredProcParam(0) = Ado_datos16.Recordset!edif_codigo
'        End If
''      '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'      iResult = CryF01.PrintReport
'      If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"
'    Else
'      MsgBox "No se puede IMPRIMIR debe elegir un registro, verifique y vuelva a intentar ...", , "Atención"
'    End If
End Sub

Private Sub generar(Autorizacion As String, Numero As String, NitCi As String, Fecha As String, Monto As String, Llave As String)
' paso 1
'    Dim suma As String
'    Dim digitos As String
'    Dim digitossum(4) As Integer
'    Dim cadenas(4) As String
'    Dim inicio As Integer
'    Dim x As Integer
'
'    Dim arc4 As String
'    Dim suma_total As Long
'    Dim sumas(4) As Long
'    Dim strlen_arc4 As Integer
'    Dim i As Integer
'    Dim total As Long
'
'    Dim mensaje As String
'    Dim last As String
'
'        numero = verhoeff_add_recursive(numero, 2)
'        nitci = verhoeff_add_recursive(nitci, 2)
'        fecha = verhoeff_add_recursive(fecha, 2)
'        monto = verhoeff_add_recursive(monto, 2)
''            Dim suma As String = CType((Long.Parse(numero) _
''                        + (Long.Parse(nitci) _
''                        + (Long.Parse(fecha) + Long.Parse(monto)))),Long).ToString
'        suma = (CStr(numero) + (CStr(nitci) + (Trim(fecha) + CStr(monto))))
'        suma = verhoeff_add_recursive(suma, 5)
'' paso2
''            Dim digitos As String = ("" + suma.Substring((suma.Length - 5), 5))
''            Dim digitossum() As Integer = New Integer() {0, 0, 0, 0, 0}
''            Dim cadenas() As String = New String() {"", "", "", "", ""}
''            Dim inicio As Integer = 0
''            Dim x As Integer = 0
'    digitos = ("" + suma.Substring((suma.Length - 5), 5))
'    digitossum(0) = 0
'    digitossum(1) = 0
'    digitossum(2) = 0
'    digitossum(3) = 0
'    digitossum(4) = 0
'    cadenas(0) = ""
'    cadenas(1) = ""
'    cadenas(2) = ""
'    cadenas(3) = ""
'    cadenas(4) = ""
'    inicio = 0
'    x = 0
''    For Each d As Char In digitos.ToCharArray
''                digitossum(x) = (Integer.Parse(d.ToString) + 1)
''                cadenas(x) = llave.Substring(inicio, (Integer.Parse(d.ToString) + 1))
''                inicio = (inicio _
''                            + (Integer.Parse(d.ToString) + 1))
''                x = (x + 1)
''    Next
'    For x = 0 To Len(digitos)
'        digitossum(x) = (CInt(digitos) + 1)
'        cadenas(x) = llave.Substring(inicio, (CInt(digitos) + 1))
'        inicio = (inicio + (CInt(digitos) + 1))
'        x = (x + 1)
'    Next x
'            autorizacion = (autorizacion + cadenas(0))
'            numero = (numero + cadenas(1))
'            nitci = (nitci + cadenas(2))
'            fecha = (fecha + cadenas(3))
'            monto = (monto + cadenas(4))
'' paso3
'    arc4 = allegedrc4((autorizacion + (numero + (nitci + (fecha + monto)))), (llave + digitos))
'' paso4
'    suma_total = 0
'    sumas(0) = 0
'    sumas(1) = 0
'    sumas(2) = 0
'    sumas(3) = 0
'    sumas(4) = 0
'    strlen_arc4 = Len(arc4)
'    i = 0
'    Do While (i < strlen_arc4)
'                x = CInt(arc4(i))
'                sumas((i Mod 5)) = (sumas((i Mod 5)) + x)
'                suma_total = (suma_total + x)
'                i = (i + 1)
'    Loop
'' paso5
'    total = 0
'    i = 0
'    Do While (i < Len(sumas))
'                total = (total + (suma_total * (sumas(i) / digitossum(i))))
'                i = (i + 1)
'    Loop
'    mensaje = big_base_convert(total, 64)
'    last = allegedrc4(mensaje, (llave + digitos)).Insert(2, "-").Insert(5, "-").Insert(8, "-")
'            If (last.Length > 11) Then
'                last = last.Insert(11, "-")
'            End If
'    'Return last

End Sub

Private Sub big_base_convert(ByVal Numero As Long, ByVal baseconv As Long)
'    Dim dic(63) As Char
'    Dim cociente As Long
'    Dim resto As Long
'    Dim palabra As String
'
'    dic(0) = Microsoft.VisualBasic.ChrW(48)
'    dic(1) = Microsoft.VisualBasic.ChrW(49)
'    dic(2) = Microsoft.VisualBasic.ChrW(50)
'    dic(3) = Microsoft.VisualBasic.ChrW(51)
'    dic(4) = Microsoft.VisualBasic.ChrW(52)
'    dic(5) = Microsoft.VisualBasic.ChrW(53)
'    dic(6) = Microsoft.VisualBasic.ChrW(54)
'    dic(7) = Microsoft.VisualBasic.ChrW(55)
'    dic(8) = Microsoft.VisualBasic.ChrW(56)
'    dic(9) = Microsoft.VisualBasic.ChrW(57)
'    dic(10) = Microsoft.VisualBasic.ChrW(65)
'    dic(11) = Microsoft.VisualBasic.ChrW(66)
'    dic(12) = Microsoft.VisualBasic.ChrW(67)
'    dic(13) = Microsoft.VisualBasic.ChrW(68)
'    dic(14) = Microsoft.VisualBasic.ChrW(69)
'    dic(15) = Microsoft.VisualBasic.ChrW(70)
'    dic(16) = Microsoft.VisualBasic.ChrW(71)
'    dic(17) = Microsoft.VisualBasic.ChrW(72)
'    dic(18) = Microsoft.VisualBasic.ChrW(73)
'    dic(19) = Microsoft.VisualBasic.ChrW(74)
'    dic(20) = Microsoft.VisualBasic.ChrW(75)
'    dic(21) = Microsoft.VisualBasic.ChrW(76)
'    dic(22) = Microsoft.VisualBasic.ChrW(77)
'    dic(23) = Microsoft.VisualBasic.ChrW(78)
'    dic(24) = Microsoft.VisualBasic.ChrW(79)
'    dic(25) = Microsoft.VisualBasic.ChrW(80)
'    dic(26) = Microsoft.VisualBasic.ChrW(81)
'    dic(27) = Microsoft.VisualBasic.ChrW(82)
'    dic(28) = Microsoft.VisualBasic.ChrW(83)
'    dic(29) = Microsoft.VisualBasic.ChrW(84)
'    dic(30) = Microsoft.VisualBasic.ChrW(85)
'    dic(31) = Microsoft.VisualBasic.ChrW(86)
'    dic(32) = Microsoft.VisualBasic.ChrW(87)
'    dic(33) = Microsoft.VisualBasic.ChrW(88)
'    dic(34) = Microsoft.VisualBasic.ChrW(89)
'    dic(35) = Microsoft.VisualBasic.ChrW(90)
'    dic(36) = Microsoft.VisualBasic.ChrW(97)
'    dic(37) = Microsoft.VisualBasic.ChrW(98)
'    dic(38) = Microsoft.VisualBasic.ChrW(99)
'    dic(39) = Microsoft.VisualBasic.ChrW(100)
'    dic(40) = Microsoft.VisualBasic.ChrW(101)
'    dic(41) = Microsoft.VisualBasic.ChrW(102)
'    dic(42) = Microsoft.VisualBasic.ChrW(103)
'    dic(43) = Microsoft.VisualBasic.ChrW(104)
'    dic(44) = Microsoft.VisualBasic.ChrW(105)
'    dic(45) = Microsoft.VisualBasic.ChrW(106)
'    dic(46) = Microsoft.VisualBasic.ChrW(107)
'    dic(47) = Microsoft.VisualBasic.ChrW(108)
'    dic(48) = Microsoft.VisualBasic.ChrW(109)
'    dic(49) = Microsoft.VisualBasic.ChrW(110)
'    dic(50) = Microsoft.VisualBasic.ChrW(111)
'    dic(51) = Microsoft.VisualBasic.ChrW(112)
'    dic(52) = Microsoft.VisualBasic.ChrW(113)
'    dic(53) = Microsoft.VisualBasic.ChrW(114)
'    dic(54) = Microsoft.VisualBasic.ChrW(115)
'    dic(55) = Microsoft.VisualBasic.ChrW(116)
'    dic(56) = Microsoft.VisualBasic.ChrW(117)
'    dic(57) = Microsoft.VisualBasic.ChrW(118)
'    dic(58) = Microsoft.VisualBasic.ChrW(119)
'    dic(59) = Microsoft.VisualBasic.ChrW(120)
'    dic(60) = Microsoft.VisualBasic.ChrW(121)
'    dic(61) = Microsoft.VisualBasic.ChrW(122)
'    dic(62) = Microsoft.VisualBasic.ChrW(43)
'    dic(63) = Microsoft.VisualBasic.ChrW(47)
'
'    cociente = 1
'    resto = 0
'    palabra = ""
'    While (cociente > 0)
'                cociente = (numero / baseconv)
'                resto = (numero Mod baseconv)
'                palabra = (dic(resto) + palabra)
'                numero = cociente
'
'    End
'    '        Return palabra
End Sub
        
Private Sub SWAP(ByRef num1 As Integer, ByRef num2 As Integer)
    Dim temp As Integer
    temp = num2
    num2 = num1
    num1 = temp
End Sub
        
'Private Sub allegedrc4(mensaje As String, llaverc4 As String)
'            Dim state() As Integer = New Integer((256) - 1) {}
'            Dim x As Integer = 0
'            Dim y As Integer = 0
'            Dim index1 As Integer = 0
'            Dim index2 As Integer = 0
'            Dim nmen As Integer = 0
'            Dim i As Integer = 0
'            Dim cifrado As String = ""
'            i = 0
'            Do While (i < 256)
'                state(i) = i
'                i = (i + 1)
'            Loop
'            Dim strlen_llave As Integer = llaverc4.Length
'            Dim strlen_mensaje As Integer = mensaje.Length
'            i = 0
'            Do While (i < 256)
'                index2 = ((CType(llaverc4(index1),Integer) _
'                            + (state(i) + index2)) _
'                            Mod 256)
'                swap(state(index2), state(i))
'                index1 = ((index1 + 1) _
'                            Mod strlen_llave)
'                i = (i + 1)
'            Loop
'            Dim cadtemp As String = ""
'            i = 0
'            Do While (i < strlen_mensaje)
'                x = ((x + 1) _
'                            Mod 256)
'                y = ((state(x) + y) _
'                            Mod 256)
'                swap(state(y), state(x))
'                ' ^ = XOR function
'                nmen = (CType(mensaje(i),Integer) Or state(((state(x) + state(y)) _
'                            Mod 256)))
'                'The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
'                cadtemp = ("0" + big_base_convert(nmen, 16))
'                cifrado = (cifrado + cadtemp.Substring((cadtemp.Length - 2), 2))
'                i = (i + 1)
'            Loop
'            Return cifrado
'End Sub
'
'Private Shared Function calcsum(ByVal number As String) As Integer
'            Dim c As Integer = 0
'            Dim n As String = reverse(number)
'            Dim len As Integer = n.Length
'            Dim nchar() As Char = n.ToCharArray
'            Dim i As Integer = 0
'            Do While (i < len)
'                c = table_d(c, table_p(((i + 1) _
'                            Mod 8), Integer.Parse(nchar(i).ToString)))
'                i = (i + 1)
'            Loop
'            Return table_inv(c)
'End Sub
'
'Private Shared Function verhoeff_add_recursive(ByVal number As String, ByVal digits As Integer) As String
'            Dim temp As String = number
'
'            While (digits > 0)
'                temp = (temp + calcsum(temp))
'                digits = (digits - 1)
'
'            End While
'            Return temp
'End Sub
'
'Private Shared Function reverse(ByVal cadena As String) As String
'            Dim str() As Char = cadena.ToCharArray
'            Array.Reverse(str)
'            Return New String(str)
'End Sub

Private Sub IMPRIME_FACTURA()
'        'IMPRIMIR FACTURA
'    Dim iResult As Variant  ', i%, y%
'    sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
'    If sino = vbYes Then
'        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior_rep.rpt"
'    Else
'        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior.rpt"
'    End If
'        CryF01.WindowShowRefreshBtn = True
'        CryF01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
'        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
'        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
'        'var_literal = "-"   'Ado_datos.Recordset!Literal
'        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
'        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
'        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'        iResult = CryF01.PrintReport
'        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"

End Sub

Private Sub BtnImprimir4_Click()
'    Select Case SSTab1.Tab
'        Case 0
'            If Ado_datos16.Recordset.RecordCount > 0 Then
'              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
'              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
'              CryV01.WindowShowRefreshBtn = True
'              CryV01.StoredProcParam(0) = Me.Ado_datos01.Recordset!ges_gestion            'glGestion
'              CryV01.StoredProcParam(1) = Me.Ado_datos01.Recordset!venta_codigo           'nroventa        '
'              CryV01.StoredProcParam(2) = Me.Ado_datos01.Recordset!cobranza_prog_codigo   'NRO_COBR        '
'              'Literal por el Total de la Compra
'              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
'              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_prog_codigo & "' "
'              iResult = CryV01.PrintReport
'              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'            Else
'              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'            End If
'        Case 1
'            If Ado_datos16.Recordset.RecordCount > 0 Then
'              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
'              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
'              CryV01.WindowShowRefreshBtn = True
'              CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion            'glGestion
'              CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo           'nroventa        '
'              CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_prog_codigo   'NRO_COBR        '
'              'Literal por el Total de la Compra
'              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
'              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
'              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_prog_codigo & "' "
'              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'              iResult = CryV01.PrintReport
'              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'            Else
'              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'            End If
'        Case 2  'Ado_datos02
'            If Ado_datos16.Recordset.RecordCount > 0 Then
'              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
'              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
'              CryV01.WindowShowRefreshBtn = True
'              CryV01.StoredProcParam(0) = Me.Ado_datos02.Recordset!ges_gestion            'glGestion
'              CryV01.StoredProcParam(1) = Me.Ado_datos02.Recordset!venta_codigo           'nroventa        '
'              CryV01.StoredProcParam(2) = Me.Ado_datos02.Recordset!cobranza_prog_codigo   'NRO_COBR        '
'              'Literal por el Total de la Compra
'              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
'              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
'              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_prog_codigo & "' "
'              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'              iResult = CryV01.PrintReport
'              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'            Else
'              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'            End If
'    End Select
  
End Sub


Private Sub cambiarEtiquetaFactura()
    If lbl_fac.Caption <> "R-101" Then
       TxtCmpbte.Locked = False
       TxtCmpbte.backColor = &H80000005
       TxtCmpbte.ForeColor = &H80000008
       lbl_factura.Caption = "Nro.de Recibo"
    Else
       TxtCmpbte.Locked = True
       TxtCmpbte.backColor = &H404040
       TxtCmpbte.ForeColor = &HFFFFFF
       lbl_factura.Caption = "Nro.de Factura"
    End If
End Sub

' Modificar factura.
Private Sub BtnModificar_Click()
   Dim codigo_doc As String
   codigo_doc = lbl_fac.Caption
   ' Verifica si existen registros.
  If Ado_datos.Recordset.RecordCount > 0 Then
       
        If codigo_doc <> "R-101" Then
             
             Call cambiarEtiquetaFactura
             Dim Cmd1 As ADODB.Command
             Dim rs  As ADODB.Recordset
             Set Cmd1 = New ADODB.Command
             Set rs = New ADODB.Recordset
            
             Cmd1.ActiveConnection = db 'sqlServer
             Cmd1.CommandType = adCmdStoredProc
             Cmd1.CommandText = "ap_genera_codigoregistro"
             Set Parm1 = Cmd1.CreateParameter("@codigo_doc", adVarChar, adParamInput, 200, codigo_doc)
             Cmd1.Parameters.Append Parm1
             rs.Open Cmd1
             rs.MoveFirst
             TxtCmpbte.Text = rs!Codigo
             rs.Close
        Else
            Call cambiarEtiquetaFactura
        End If
  Else
      Call cambiarEtiquetaFactura
  End If

  If Ado_datos.Recordset.RecordCount > 0 Then
    If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") And (Ado_datos16.Recordset!venta_tipo = "E" Or Ado_datos16.Recordset!venta_tipo = "V" Or Ado_datos16.Recordset!venta_tipo = "C" Or Ado_datos16.Recordset!venta_tipo = "L") Then
      FraNavega.Enabled = False
'      fraOpciones.Visible = False
      FraGrabarCancelar.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      'swgrabar = 0
      swnuevo = 2
      Txt_tdc.Text = GlTipoCambioMercado    'GlTipoCambioOficial
      SSTab1.Tab = 1
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(2) = False
      FrmCobros.Visible = True
      FrmCobros.Enabled = True
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      
'      BtnImprimir2.Visible = False
'      BtnImprimir3.Visible = False
      CmdFoto.Visible = False
      If Ado_datos.Recordset!factura_impresa = "N" And codigo_doc = "R-101" Then
      '    sino = MsgBox("Registrará la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
      '    If sino = vbYes Then
              'DTPFechaProg.Visible = True
              DTPFechaCobro.Visible = True
              DTPFechaCobro.Value = Date
'              Lbl_nombre_fac.Caption = "Factura a Nombre de:"
'              lbl_fechas.Caption = "Fecha de Cobranza"
              TxtCmpbte.Text = "0"
      '        Txt_parche.Visible = False      '&H80000013&
      '        'dtc_desc2A.BackColor = &H80000013
      '    Else
      '        DTPFechaProg.Visible = True
      '        DTPFechaCobro.Visible = False
      '        Lbl_nombre_fac.Caption = "Cliente :"
      '        lbl_fechas.Caption = "Fecha Programada de Cobranza"
      '        Txt_parche.Visible = True       '&H80000005&
      '        'dtc_desc2A.BackColor = &H80000005
      '    End If
      Else
      '    DTPFechaProg.Visible = True
      '    DTPFechaCobro.Visible = False
      '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
      End If
      'TxtMonto.Text = CDbl(TxtDsctoTot)
      TxtMonto.SetFocus
    Else
      MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModificar1_Click()
  If Ado_datos01.Recordset.RecordCount > 0 Then
    If Ado_datos01.Recordset!estado_codigo_sol = "REG" Then
      SSTab1.Tab = 0
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
      'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
      FraNavega1.Enabled = False
'      fraOpciones1.Visible = False
      FraGrabarCancelar1.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      FrmCobros1.Visible = True
      FrmCobros1.Enabled = True
      swnuevo = 2
      DTPfechasol.Value = Date
      Txt_deposito.Text = "0"
      TxtMonto1.SetFocus
    Else
      MsgBox "No se puede editar, porque el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If

End Sub

Private Sub BtnModificar2_Click()
     
  If Ado_datos02.Recordset.RecordCount > 0 Then
    If Ado_datos02.Recordset!estado_codigo_fac = "APR" And Ado_datos02.Recordset!estado_codigo_bco = "REG" Then
      SSTab1.Tab = 2
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = True
      'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
      FraNavega2.Enabled = False
'      fraOpciones1.Visible = False
      FraGrabarCancelar2.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      FrmCobros2.Visible = True
      FrmCobros2.Enabled = True
      swnuevo = 2
      'DTPFechaCobro2.Value = Date
      'DTPFechaCobro02.Value = Date
      'Txt_deposito.Text = "0"
      TxtMonto02.SetFocus
    Else
      If Ado_datos02.Recordset!estado_codigo_bco = "APR" And Ado_datos02.Recordset!estado_codigo = "REG" And usr_codigo = "ASANTIVAÑEZ" Then
            SSTab1.Tab = 2
          SSTab1.TabEnabled(0) = False
          SSTab1.TabEnabled(1) = False
          SSTab1.TabEnabled(2) = True
          'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
          FraNavega2.Enabled = False
    '      fraOpciones1.Visible = False
          FraGrabarCancelar2.Visible = True
          FrmDetalle.Enabled = False
          FrmCobranza.Enabled = False
          FrmABMDet.Visible = False
          FrmABMDet2.Visible = False
          FrmCobros2.Visible = True
          FrmCobros2.Enabled = True
          swnuevo = 2
          'DTPFechaCobro2.Value = Date
          'DTPFechaCobro02.Value = Date
          'Txt_deposito.Text = "0"
          TxtMonto02.SetFocus
      Else
            MsgBox "No se puede editar, porque el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
      End If
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If

End Sub

'Private Sub BtnImprimir1_Click()
'    If Ado_datos16.Recordset.RecordCount > 0 Then
'        If Ado_datos02.Recordset.RecordCount > 0 Then       'And Ado_datos02.Recordset!cobranza_codigo <> Null
'            Monto_Bs = IIf(IsNull(Ado_datos02.Recordset!cobranza_total_bs), 0, Ado_datos02.Recordset!cobranza_total_bs)
'            montoLiteral = Literal(CStr(Monto_Bs)) + " Bolivianos"
'    '            Dim iResult As Variant  ', i%, y%
'          'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'          CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
'          CryR01.WindowShowRefreshBtn = True
'          CryR01.StoredProcParam(0) = Ado_datos16.Recordset!venta_codigo
'          CryR01.StoredProcParam(1) = IIf(IsNull(Ado_datos02.Recordset!cobranza_codigo), 0, Ado_datos02.Recordset!cobranza_codigo)
'          CryR01.Formulas(1) = "literalcobro = '" & montoLiteral & "' "
'          CryR01.Formulas(2) = "correlcobro = '" & IIf(IsNull(Ado_datos02.Recordset!cobranza_codigo), 0, Ado_datos02.Recordset!cobranza_codigo) & "' "
'          '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'          iResult = CryR01.PrintReport
'          If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'        Else
'            MsgBox "No se puede IMPRIMIR, no existen datos de Facturación o Cobranzas Solicitadas y vuelva a intentar ...", , "Atención"
'        End If
'    Else
'      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'    End If
'End Sub

Private Sub BtnImprimir2_Click()
    Fw_ReportesVentas.lbl_titulo = "REPORTES DE VENTAS"
    Fw_ReportesVentas.Show
End Sub

Private Sub btnPrintOption_Click()
    If Option1.Value = True Then
        If Ado_datos16.Recordset.RecordCount > 0 Then
            CryF01.ReportFileName = App.Path & "\reportes\ventas\fr_contrato_kardex_cliente.rpt"
            CryF01.WindowShowRefreshBtn = True
            If Ado_datos16.Recordset!edif_codigo = "" Then
                CryF01.StoredProcParam(0) = "%"
            Else
                'CryF01.StoredProcParam(0) = cmb_codigoedificio.Text ' para REPORTE POR OPCIONES (PENDIENTE)
                CryF01.StoredProcParam(0) = Ado_datos16.Recordset!edif_codigo
            End If
            iResult = CryF01.PrintReport
            If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"
            Else
                MsgBox "No se puede IMPRIMIR, debe elegir un registro, verifique y vuelva a intentar ...", , "Atención"
        End If
    ElseIf Option2.Value = True Then
        If Ado_datos16.Recordset.RecordCount > 0 Then
            CryR01.ReportFileName = App.Path & "\reportes\ventas\fr_contrato_kardex.rpt"
            CryR01.WindowShowRefreshBtn = True
            If Ado_datos16.Recordset!venta_codigo = "" Then
                CryR01.StoredProcParam(0) = "%"
                CryR01.StoredProcParam(1) = CStr(Ado_datos16.Recordset!venta_codigo)
            Else
                CryR01.StoredProcParam(0) = Ado_datos16.Recordset!edif_codigo
                CryR01.StoredProcParam(1) = CStr(Ado_datos16.Recordset!venta_codigo)
            End If
            iResult = CryR01.PrintReport
            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
            Else
                MsgBox "No se puede IMPRIMIR, el registro, verifique los datos y vuelva a intentar ...", , "Atención"
            End If
    ElseIf Option3.Value = True Then
        If Ado_datos16.Recordset.RecordCount > 0 Then
            CryF02.ReportFileName = App.Path & "\reportes\ventas\fr_contrato_kardex_servicio.rpt"
            CryF02.WindowShowRefreshBtn = True
            CryF02.StoredProcParam(0) = Ado_datos16.Recordset!edif_codigo
            CryF02.StoredProcParam(1) = Ado_datos16.Recordset!subproceso_codigo
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
            Else
                MsgBox "No se puede IMPRIMIR, debe elegir un registro, verifique y vuelva a intentar ...", , "Atención"
            End If
    ElseIf Option4.Value = True Then
        If Ado_datos16.Recordset.RecordCount > 0 Then
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\fr_contrato_kardex_cliente_tes.rpt"
            CryQ01.WindowShowRefreshBtn = True
            If Ado_datos16.Recordset!edif_codigo = "" Then
                CryQ01.StoredProcParam(0) = "%"
            Else
                CryQ01.StoredProcParam(0) = Ado_datos16.Recordset!edif_codigo
            End If
            iResult = CryQ01.PrintReport
            If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"
            Else
                MsgBox "No se puede IMPRIMIR, debe elegir un registro, verifique y vuelva a intentar ...", , "Atención"
            End If
    ElseIf Option5.Value = True Then
        MsgBox "Reporte aun en desarrollo ...", , "Atención"
    ElseIf Option6.Value = True Then
        If Ado_datos16.Recordset.RecordCount > 0 Then
            CryQ02.ReportFileName = App.Path & "\reportes\ventas\fr_contrato_kardex_cobrador_tes.rpt"
            CryQ02.WindowShowRefreshBtn = True
            If Ado_datos16.Recordset!edif_codigo = "" Then
                CryQ02.StoredProcParam(0) = "%"
            Else
                CryQ02.StoredProcParam(0) = Ado_datos16.Recordset!beneficiario_codigo_cobr
            End If
            iResult = CryQ02.PrintReport
            If iResult <> 0 Then MsgBox CryQ02.LastErrorNumber & " : " & CryQ02.LastErrorString, vbCritical, "Error de impresión"
            Else
                MsgBox "No se puede IMPRIMIR, debe elegir un registro, verifique y vuelva a intentar ...", , "Atención"
            End If
    ElseIf Option7.Value = True Then
        If Ado_datos16.Recordset.RecordCount > 0 Then
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\fr_contrato_kardex_cliente_tes_dol.rpt"
            CryQ01.WindowShowRefreshBtn = True
            If Ado_datos16.Recordset!edif_codigo = "" Then
                CryQ01.StoredProcParam(0) = "%"
            Else
                CryQ01.StoredProcParam(0) = Ado_datos16.Recordset!edif_codigo
            End If
            iResult = CryQ01.PrintReport
            If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"
            Else
                MsgBox "No se puede IMPRIMIR, debe elegir un registro, verifique y vuelva a intentar ...", , "Atención"
            End If
    ElseIf Option8.Value = True Then
        MsgBox "Reporte aun en desarrollo ...", , "Atención"
    ElseIf Option9.Value = True Then
        MsgBox "Reporte aun en desarrollo ...", , "Atención"
    End If
End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

'Private Sub Cmd_Cliente_Click()
'    glPersNew = "P"
'    frmBeneficiario.Show 'vbModal
'End Sub

Private Sub CmdCancelaCobro_Click()
End Sub

Private Sub BtnModDetalle2_Click()
'  If ado_datos14.Recordset.RecordCount > 0 Then
'    SSTab1.Tab = 2
'    SSTab1.TabEnabled(2) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(1) = False
'
'    FrmEdita.Visible = True
''    BtnImprimir2.Visible = False
''    BtnImprimir3.Visible = False
'  Else
'    MsgBox "No Existen Bienes Registrados, Verifique por favor !! ", vbExclamation, "Atención!"
'  End If

    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'FACTURAS EMITIDAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If

End Sub

Private Sub BtnDesAprobar_Click()
'  sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
'  If sino = vbYes Then
'    Dim rstdestino As New ADODB.Recordset
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " ", db, adOpenDynamic, adLockOptimistic
'    If Not rstdestino.BOF Then rstdestino.MoveFirst
'    If Not rstdestino.BOF And Not rstdestino.EOF Then
'      rstdestino("estado_codigo") = "REG"
'      rstdestino.Update
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
'    marca1 = Ado_datos.Recordset.Bookmark
'    Call OptFilGral1_Click
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo_fac = "REG" And Ado_datos.Recordset!factura_impresa = "N" Then
        Ado_datos.Recordset!estado_codigo_sol = "REG"
        Ado_datos.Recordset!estado_codigo_fac = "REG"
        Ado_datos.Recordset.Update
          'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
    Else
        MsgBox "No se puede DEVOLVER, el registro ya fue FACTURADO, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
 End If
End Sub

'Private Sub CmdDetallePoa_Click()
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'   MsgBox "No Existen Registros ", vbInformation, "Formulario 11"
'  Else
'    marca1 = Ado_datos.Recordset.BookMark
'    FrmPoasCapturaALB.Lblformulario = "F11"
'    FrmPoasCapturaALB.lblges_gestion = Ado_datos.Recordset!ges_gestion
'    FrmPoasCapturaALB.lblcodigo_unidad = Ado_datos.Recordset!codigo_unidad
'    FrmPoasCapturaALB.lblcodigo_solicitud = Ado_datos.Recordset!codigo_solicitud
'    FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'Ado_datos.Recordset!tipoben_codigo
'    FrmPoasCapturaALB.Show vbModal
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'    '
'  Else
'    Ado_datos.Refresh
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  End If
'End Sub

'Private Sub cmdElige_Click()
'  With ALFrmMateriales
'        .ALPrincipal
'        If .QResp Then
'            txtCodigo.Text = .QCodigo
'            txtDesc.Text = .QItem
'        End If
'    End With
'    Txtcant_alm = 0
'    Cant_Alm = 0
'    DE.dbo_albSacaDetalleMaterial Mid(txtCodigo, 3, 12), descri_bien, Cant_Alm
'    Txtcant_alm = Cant_Alm
'    If Cant_Alm >= TxtCantPedi Then
'        optSi = True
'    Else
'        optNo = True
'    End If
'End Sub

Private Sub Contabiliza_venta()
'    Call graba_proyecto
    If VAR_SW = 1 Then
        Call graba_ingreso
    End If
    'If VAR_SW = 1 Then
        Set rstdestino = New ADODB.Recordset
        If VAR_TIPOV = "V" Or VAR_TIPOV = "C" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
        Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
        End If
        If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo
            If VAR_SW = 1 Then
                VAR_CODTIPO = "REF"
            Else
                VAR_CODTIPO = "REC"
            End If
            'Modificar con CASE WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW MAY-2015
            If VAR_COD4 = "DVTA" Then
               VAR_TSOL = "3"
               VAR_PARTIDA = "11200"
            Else
               VAR_TSOL = "10"
               VAR_PARTIDA = "11300"
            End If
            If VAR_COD4 = "DNMAN" Then
               VAR_TSOL = "10"
               VAR_PARTIDA = "11320"
            End If
            If VAR_COD4 = "DNREP" Then
               VAR_TSOL = "7"
               VAR_PARTIDA = "11330"
            End If
            If VAR_COD4 = "DNMOD" Then
               VAR_TSOL = "9"
               VAR_PARTIDA = "11340"
            End If
        End If
    'End If
  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  'If sino = vbYes Then
    ' INI CORRECCION 18-JUN-2014
    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

'               gestion0 = Ado_datos.Recordset("ges_gestion")
'               correlv = Ado_datos.Recordset("venta_codigo")
'               nroventa = Ado_datos.Recordset("venta_codigo")
'
'               VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'               VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
'               VAR_DOL2 = Round(Ado_datos.Recordset("cobranza_deuda_dol"), 2)
'               VAR_BS2 = Round(Ado_datos.Recordset("cobranza_deuda_bs"), 2)
'
'               VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'               VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'               VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'               VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'               VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'                VAR_CODANT = rstdestino!ingreso_codigo
'            VAR_ORG = rstdestino!org_codigo
'            VAR_FTE = rstdestino!org_codigo
'            VAR_CODTIPO = "REC"
'            VAR_PARTIDA = "11200"

    fte_codigo1 = VAR_FTE
    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DEI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
              'cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
              'Subcta_deb11 = rstdestino!Subcta_cred1
              'Subcta_deb21 = rstdestino!Subcta_cred2
    
              'cta_credito1 = rstdestino2!cta_deb
              'Subcta_cred11 = rstdestino2!Subcta_deb1
              'Subcta_cred21 = rstdestino2!Subcta_deb2
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
        Case "REC"
            If VAR_MONEDA = "BOB" Then
                rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and SubCta_Deb1 = '01' ", db, adOpenKeyset, adLockReadOnly
            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            Else
                rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "  and SubCta_Deb1 = '02' ", db, adOpenKeyset, adLockReadOnly
            End If
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
                        
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
            'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                'JQA FEB-2016
                'Exit Sub
              End If
            End If
            If rs_aux1.State = 1 Then rs_aux1.Close
        Case "REF"
            If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REF' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REF' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
                        
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
            'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                Exit Sub
              End If
            End If
            If rs_aux1.State = 1 Then rs_aux1.Close
            
        Case "DYR"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
        Case "DES"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "ANI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DVI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
            '' 02/07/2014 VERIFICAR
            'If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
            'If rstdestino2.State = 1 Then rstdestino2.Close
            'rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            'If rstdestino.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
            '  MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
            '  Exit Sub
            'End If
        Case Else
            MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
            If rstdestino.State = 1 Then rstdestino.Close
            Exit Sub
    End Select
    'If rstdestino.State = 1 Then rstdestino.Close
    '**** FIN VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************

    Dim cta_deb1 As String
    Dim Subcta_deb11 As String
    Dim Subcta_deb21 As String

    Dim cta_credito1 As String
    Dim Subcta_cred11 As String
    Dim Subcta_cred21 As String

    Dim cod_ant As Integer
    Dim org_ant As String

    'If DtCCta_codigo.Text <> "01" Then
    '  If rstdestino.State = 1 Then rstdestino.Close
    '  rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
    '  If Not rstFc_cuenta_bancaria.EOF Then
    '    fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
    '  Else
    '  End If
    'Else
    '    fte_codigo1 = Me.DtCFte_codigo.Text
    'End If
    'If VAR_CODTIPO = "DEI" Or VAR_CODTIPO = "DES" Then
    '  fte_codigo1 = Me.DtCFte_codigo.Text
    'End If
    
'    fte_codigo1 = VAR_FTE
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim v_Tipo_Comp(1, 2)
'
'    v_Tipo_Comp(1, 1) = VAR_CODTIPO
    
'    If VAR_CODTIPO = "DYR" Then
'      'j = 2
'      'v_Tipo_Comp(1, 1) = "CAD"
'      'v_Tipo_Comp(1, 2) = "CAR"
'      j = 2
'      v_Tipo_Comp(1, 1) = "DYR"
'    Else
'      j = 1
'      v_Tipo_Comp(1, 1) = IIf(VAR_CODTIPO = "DEI", "DEI", IIf(VAR_CODTIPO = "REC", "REC", IIf(VAR_CODTIPO = "DES", "DES", IIf(VAR_CODTIPO = "ANI", "ANI", ""))))
'    End If
'
'    If VAR_CODTIPO = "DVI" Then
'      j = 1
'      v_Tipo_Comp(1, 1) = "DVI"
'    End If

'    For i = 1 To j
'      If rstdestino.State = 1 Then rstdestino.Close
'      If v_Tipo_Comp(1, i) = "DEI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "" Then
'        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
'        Exit Sub
'      End If

    ' INI CORRECCION 18-JUN-2014
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' 02/07/2014 VERIFICAR
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rs_aux2.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
'          MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'          Exit Sub
'        End If
'      End If
'
'      If rs_aux2.RecordCount < 1 Then
'        MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'        Exit Sub
'      End If
'    Next

    'If rstdestino.State = 1 Then rstdestino.Close

    v_Tipo_Comp(1, 1) = VAR_CODTIPO
    
    db.BeginTrans
'    Frmmensaje.Visible = True
'    LblMensaje.Caption = "Este proceso tomará solo unos segundos, gracias"
    '========================================
    '==== verifica si ya fue contabilizado
      yacontabilizo = 0
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '" & VAR_CODANT & "' and org_codigo = '" & VAR_ORG & "' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      'rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '2' and org_codigo = '111' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      If rs_aux2.RecordCount > 0 Then
        ' revisar para validar mejor si YA contabilizo !!
        'yacontabilizo = 1
        yacontabilizo = 0
      Else
        yacontabilizo = 0
      End If
      If yacontabilizo = 1 Then
        'MsgBox "aqui recontabilizar" & rstdestino!Cod_trans & " -- " & rstdestino!org_codigo & " / " & rstdestino!Cod_Comp
        Var_Comp = rs_aux2!Cod_Comp
      Else
        '===== ini GENERA EL CODIGO DE COMPROBANTE ====
        Set rstCodComp = New ADODB.Recordset
        rstCodComp.CursorLocation = adUseClient
        If rstCodComp.State = 1 Then rstCodComp.Close
        rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
        If rstCodComp.RecordCount > 0 Then
          Var_Comp = CDbl(rstCodComp!numero_correlativo)
          Var_Comp = Var_Comp + 1
          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
          rstCodComp.Update
        End If
        If rstCodComp.State = 1 Then rstCodComp.Close
        '===== fin TERMINA GENERACION DE COMPROBANTE =====

      '==== ini registro co_comprobante_m

        rs_aux2.AddNew
        rs_aux2("cod_comp") = Var_Comp
      End If
    '========================================
    'anterior
    '      If rstdestino.State = 1 Then rstdestino.Close
    '      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
    '      If rstdestino.RecordCount > 0 Then
    '      End If
    '      rstdestino.AddNew
    
    '      rstdestino("cod_comp") = Var_Comp
    'anterior
      rs_aux2("Tipo_Comp") = VAR_CODTIPO        'v_Tipo_Comp(1, i)
      rs_aux2("cod_trans") = VAR_CODANT
      rs_aux2("org_codigo") = VAR_ORG
      rs_aux2("ges_gestion") = glGestion        'Year(Date)
      'rstdestino("Num_Respaldo") = Ado_datos.Recordset("numero_documento")
      If yacontabilizo = 0 Then
        rs_aux2("Fecha_transacion") = Date
      End If
      rs_aux2("beneficiario_codigo") = VAR_BENEF
      rs_aux2("glosa") = "CONTABILIZA: " + VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4           'Ado_datos16.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = VAR_SOL         'Ado_datos16.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE
      
      rs_aux2("proceso_codigo") = "FIN"
      rs_aux2("subproceso_codigo") = "FIN-02"
      Select Case VAR_CODTIPO
        Case "DEI"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "REC"
            rs_aux2("etapa_codigo") = "FIN-02-03"
        Case "DYR"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "DES"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "ANI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "REF"
            rs_aux2("etapa_codigo") = "FIN-02-02"
      End Select
      
      rs_aux2("clasif_codigo") = "ADM"
      rs_aux2("doc_codigo") = "R-110"
      rs_aux2("doc_numero") = Var_Comp
      rs_aux2("pro_codigo_det") = VAR_PROY2
    
      rs_aux2("estado_codigo") = "APR"

      If yacontabilizo = 0 Then
        rs_aux2("usr_codigo") = glusuario
        rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
        rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If
      rs_aux2.Update
      '==== fin registro co_comprobantre_m

    Dim d_cta_nombre_1 As String
    Dim d_aux1_1 As String
    Dim d_aux2_1 As String
    Dim d_aux3_1 As String
    Dim h_cta_nombre_1 As String
    Dim h_aux1_1 As String
    Dim h_aux2_1 As String
    Dim h_aux3_1 As String
    'If rstdestino.State = 1 Then rstdestino.Close
    
    For i = 1 To j
'    ' nuevo ini
'      If v_Tipo_Comp(1, i) = "DEI" Then     'Devengado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then     'Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then     'Devengado y Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then     'Desafectado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then     'Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then     'Desafectado y Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If

'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' VERIFICAR SI SE ESTA CONTROLANDA con el DYR
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rstdestino.RecordCount > 0 And rstdestino2.RecordCount > 0 Then
'          cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
'          Subcta_deb11 = rstdestino!Subcta_cred1
'          Subcta_deb21 = rstdestino!Subcta_cred2
'
'          cta_credito1 = rstdestino2!cta_deb
'          Subcta_cred11 = rstdestino2!Subcta_deb1
'          Subcta_cred21 = rstdestino2!Subcta_deb2
'        Else
'          MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
''          Exit Sub
'        End If
'      End If
'
'      If rstdestino.RecordCount > 0 And v_Tipo_Comp(1, i) <> "DVI" Then
'        cta_deb1 = rstdestino("cta_deb")
'        Subcta_deb11 = rstdestino("Subcta_deb1")
'        Subcta_deb21 = rstdestino("Subcta_deb2")
'        cta_credito1 = rstdestino("cta_cred")
'        Subcta_cred11 = rstdestino("Subcta_cred1")
'        Subcta_cred21 = rstdestino("Subcta_cred2")
'      Else
'        'MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
'        'Exit Sub
'
'      End If
      
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")
        
        cta_credito1 = rstdestino("cta_cred")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")
      Else
        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
        Subcta_deb11 = rstdestino!Subcta_cred1
        Subcta_deb21 = rstdestino!Subcta_cred2
    
        cta_credito1 = rstdestino!cta_deb
        Subcta_cred11 = rstdestino!Subcta_deb1
        Subcta_cred21 = rstdestino!Subcta_deb2
      End If
      
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        d_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        h_aux1_1 = rs_aux1("aux1")
        h_aux2_1 = rs_aux1("aux2")
        h_aux3_1 = rs_aux1("aux3")
      End If
    ' nuevo fin
    
      '===== ini registra CO_diaRIO =========
      Set rstdestino2 = New ADODB.Recordset
      If rstdestino2.State = 1 Then rstdestino2.Close
      rstdestino2.Open "select * from co_diario where Cod_Comp = " & Var_Comp, db, adOpenKeyset, adLockOptimistic
      'If rstdestino2.RecordCount > 0 Then
      '  MsgBox "Ya Existe el asiento, se reemplazará con los nuevos datos..."
      'Else
        rstdestino2.AddNew
        rstdestino2("Cod_Comp") = Var_Comp
      'End If
        rstdestino2("Cod_Comp_Detalle") = rstdestino2.RecordCount
      'rstdestino2("Tipo_Comp") = "DEI"   'v_Tipo_Comp(1, i)
      'rstdestino2("Cod_Comp_C") = Var_Comp
      'If v_Tipo_Comp(1, i) = "DEI" Or v_Tipo_Comp(1, i) = "REC" Then
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_deb11
        rstdestino2("D_SubCta2") = Subcta_deb21
        rstdestino2("D_Aux1") = d_aux1_1
        rstdestino2("D_Aux2") = d_aux2_1
        rstdestino2("D_Aux3") = d_aux3_1
        ' para Aux1
'        Select Case d_aux1_1
'                Case "01"
'                    VAR_COD1 = VAR_BENEF
'                Case "02"
'                    VAR_COD1 = VAR_CTA
'                Case "03"
'                    VAR_COD1 = VAR_PROY2
'                Case "04"
'                    VAR_COD1 = Ado_datos.Recordset("unidad_codigo")
'                Case "05"
'                    VAR_COD1 = ""
'                Case "06"
'                    VAR_COD1 = ""
'                Case "07"
'                    VAR_COD1 = ""
'                Case "08"
'                    VAR_COD1 = ""
'                Case "09"
'                    VAR_COD1 = VAR_ORG
'                Case "10"
'                    VAR_COD1 = ""
'                Case "11"
'                    VAR_COD1 = ""
'                Case "12"
'                    VAR_COD1 = ""
'        End Select
        ' ini PARA EL FUTURO ******** REVISAR
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        SQL_FOR = "select * from cc_tipo_auxiliar where aux = '" & d_aux1_1 & "' "
'        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux4.RecordCount > 0 Then
'            Set rs_aux1 = New ADODB.Recordset
'            If rs_aux1.State = 1 Then rs_aux1.Close
'            SQL_FOR = "select * from " + rs_aux4!NombreTabla + " where " + rs_aux4!nombre_codigo + " = " + VAR_COD1
'            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'            If rs_aux1.RecordCount > 0 Then
'        Else
'        End If
        ' fin PARA EL FUTURO ******** REVISAR
        Select Case d_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(d_aux1_1, CStr(VAR_BENEF))    'DESAUX =
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                Call DESCAUX(d_aux1_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(d_aux1_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux1_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux1_1, rstdestino2!D_Cta_Aux1)
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                Call DESCAUX(d_aux1_1, CStr(VAR_ORG))
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux1 = DESAUX
        Select Case d_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(d_aux2_1, CStr(VAR_BENEF))
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                Call DESCAUX(d_aux2_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(d_aux2_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux2_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux2_1, rstdestino2!D_Cta_Aux2)
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                Call DESCAUX(d_aux2_1, CStr(VAR_ORG))
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux2 = DESAUX
        Select Case d_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(d_aux3_1, CStr(VAR_BENEF))
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                Call DESCAUX(d_aux3_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(d_aux3_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux3_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux3_1, rstdestino2!D_Cta_Aux3)
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                Call DESCAUX(d_aux3_1, CStr(VAR_ORG))
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux3 = DESAUX
        
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        If cta_deb1 = "6112" Then
            rstdestino2("D_MontoBs") = VAR_BS2 * 0.03
            rstdestino2("D_MontoDl") = VAR_DOL2 * 0.03
        Else
            If cta_credito1 = "2112" Then
                rstdestino2("D_MontoBs") = VAR_BS2 * 0.13
                rstdestino2("D_MontoDl") = VAR_DOL2 * 0.13
            Else
                If cta_deb1 = "1111" Then
                    rstdestino2("D_MontoBs") = VAR_BS2
                    rstdestino2("D_MontoDl") = VAR_DOL2
                Else
                    rstdestino2("D_MontoBs") = VAR_BS2 * 0.87
                    rstdestino2("D_MontoDl") = VAR_DOL2 * 0.87
                    'rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
                    'rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
                End If
            End If
        End If
        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        'AQUI MONEDA 02/07/01
        'rstdestino2("D_Cambio") = GlTipoCambioMercado
        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
        rstdestino2("H_Cuenta") = cta_credito1
        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_cred11
        rstdestino2("H_SubCta2") = Subcta_cred21
        rstdestino2("H_Aux1") = h_aux1_1
        rstdestino2("H_Aux2") = h_aux2_1
        rstdestino2("H_Aux3") = h_aux3_1
        'rstdestino2("H_Cta_Aux1") = ""
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(h_aux1_1, CStr(VAR_BENEF))
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                Call DESCAUX(h_aux1_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(h_aux1_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux1_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux1_1, rstdestino2!H_Cta_Aux1)
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                Call DESCAUX(h_aux1_1, CStr(VAR_ORG))
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux1 = DESAUX
        
        Select Case h_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(h_aux2_1, CStr(VAR_BENEF))
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                Call DESCAUX(h_aux2_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(h_aux2_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux2_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux2_1, rstdestino2!H_Cta_Aux2)
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                Call DESCAUX(h_aux2_1, CStr(VAR_ORG))
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux2 = DESAUX
        Select Case h_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(h_aux3_1, CStr(VAR_BENEF))
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                Call DESCAUX(h_aux3_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(h_aux3_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux3_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux3_1, rstdestino2!H_Cta_Aux3)
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                Call DESCAUX(h_aux3_1, CStr(VAR_ORG))
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux3 = DESAUX
        
        
'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        If cta_deb1 = "6112" Then
            rstdestino2("H_MontoBs") = VAR_BS2 * 0.03
            rstdestino2("H_MontoDl") = VAR_DOL2 * 0.03
        Else
            If cta_credito1 = "2112" Then
                rstdestino2("H_MontoBs") = VAR_BS2 * 0.13
                rstdestino2("H_MontoDl") = VAR_DOL2 * 0.13
            Else
                If cta_deb1 = "1111" Then
                    rstdestino2("H_MontoBs") = VAR_BS2
                    rstdestino2("H_MontoDl") = VAR_DOL2
                Else
                    rstdestino2("H_MontoBs") = VAR_BS2 * 0.87
                    rstdestino2("H_MontoDl") = VAR_DOL2 * 0.87
                    'rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
                    'rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
                End If
            End If
            'rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
            'rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        End If
        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
      End If

      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
      If (VAR_CODTIPO = "DES") Or (VAR_CODTIPO = "ANI") Or (VAR_CODTIPO = "DVI") Then
        'desafecta un devengado
        rstdestino2("D_Cuenta") = cta_credito1
        rstdestino2("D_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_cred11
        rstdestino2("D_SubCta2") = Subcta_cred21
        rstdestino2("D_Aux1") = h_aux1_1
        rstdestino2("D_Aux2") = h_aux2_1
        rstdestino2("D_Aux3") = h_aux3_1
'        rstdestino2("D_Cta_Aux1") = "VESCT"
        Select Case h_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = ""
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
        End Select
        
        Select Case h_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = ""
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
        End Select
        
        Select Case h_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = ""
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
        End Select
'        If h_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("D_Cambio") = GlTipoCambioMercado

        rstdestino2("H_Cuenta") = cta_deb1
        rstdestino2("H_Nombre") = d_cta_nombre_1  ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_deb11
        rstdestino2("H_SubCta2") = Subcta_deb21
        rstdestino2("H_Aux1") = d_aux1_1
        rstdestino2("H_Aux2") = d_aux2_1
        rstdestino2("H_Aux3") = d_aux3_1
'        rstdestino2("H_Cta_Aux1") = "VESCT"
        Select Case d_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = ""
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
        End Select
        
        Select Case d_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = ""
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
        End Select
        
        Select Case d_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = ""
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("H_Cambio") = GlTipoCambioMercado
      End If

'      '==== INI DVI ====
'      If (VAR_CODTIPO = "DVI") Then
'        rstdestino2("D_Cuenta") = cta_deb1
''        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_deb11
'        rstdestino2("D_SubCta2") = Subcta_deb21
'        rstdestino2("D_Aux1") = d_aux1_1
'        rstdestino2("D_Aux2") = d_aux2_1
'        rstdestino2("D_Aux3") = d_aux3_1
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("D_Cambio") = GlTipoCambioMercado
'        rstdestino2("H_Cuenta") = cta_credito1
''        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("H_SubCta1") = Subcta_cred11
'        rstdestino2("H_SubCta2") = Subcta_cred21
'        rstdestino2("H_Aux1") = h_aux1_1
'        rstdestino2("H_Aux2") = h_aux2_1
'        rstdestino2("H_Aux3") = h_aux3_1
'        'rstdestino2("H_Cta_Aux1") = "VESCT"
'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("H_Cambio") = GlTipoCambioMercado
'      End If
'      '==== FIN DVI ====

      If yacontabilizo = 0 Then
        rstdestino2("Usr_codigo") = glusuario
        rstdestino2("Fecha_registro") = Date
        rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If
      
      rstdestino2.Update
      If rstdestino2.State = 1 Then rstdestino2.Close
      '======= fin registra co_diario ==========
      rstdestino.MoveNext
    Next i
    '======= inI Actualiza campos de estatus de ingresos ==========
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '" & correlativo1 & "' and org_codigo = '" & VAR_ORG & "' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' ", db, adOpenDynamic, adLockOptimistic
'    rstdestino.MoveFirst
'    If Not (rstdestino.EOF) Then
'      rstdestino("estado_aprobacion") = "S"
'        If VAR_CODTIPO = "DEI" Then
'          rstdestino("estado_devengado") = "S"
'        End If
'        If VAR_CODTIPO = "REC" Then
'          rstdestino("estado_recaudado") = "S"
'        End If
'        If VAR_CODTIPO = "DYR" Then
'          rstdestino("estado_devengado") = "S"
'          rstdestino("estado_recaudado") = "S"
'        End If
'
'        If VAR_CODTIPO = "DES" Then
'          rstdestino("estado_desafectado") = "S"
'        End If
'        If VAR_CODTIPO = "ANI" Then
'          rstdestino("estado_anulado") = "S"
'        End If
'        If VAR_CODTIPO = "DVI" Then
'          rstdestino!estado_desafectado = "S"
'          rstdestino!estado_anulado = "S"
'        End If
'       rstdestino.Update
'       If rstdestino.State = 1 Then rstdestino.Close
'    End If
    '======= fin Actualiza campos de estatus de ingresos ==========
    ' AAAAAAAAAQQQQQQQQQQQUUUUUUUUUUUIIIIIIIIIII
    cod_ant = 0
    org_ant = ""
    '======= ini Actualiza el monto recaudado  ==========
    If (VAR_CODTIPO = "REC") Then
      '      If rstdestino.State = 1 Then rstdestino.Close
      '      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      '      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
      '        cod_ant = rstdestino("ingreso_codigo_anterior")
      '        org_ant = rstdestino("org_codigo")
      '      End If
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") + VAR_DOL2
          rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") + VAR_BS2
          rstdestino.Update
      End If
      If rstdestino.State = 1 Then rstdestino.Close
    End If

    If (VAR_CODTIPO = "DES") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      Print VAR_CODANT
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
'        org_ant = rstdestino("org_codigo")
'      End If

      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "DEI" Then 'And VAR_CODTIPO = "DES"
'          rstdestino!estado_desafectado = "S" 02/07/01
          rstdestino!estado_codigo = "DES"
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        Else
          rstdestino("estado_codigo") = "DES"
'          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
          cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
          org_ant = rstdestino("org_codigo")
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
          'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & cod_ant & " and org_codigo = '" & org_ant & "' ", db, adOpenKeyset, adLockOptimistic
          rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
          If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
            rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
            rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") - VAR_BS2
          End If
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        End If
      End If
    End If

    If (VAR_CODTIPO = "ANI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "REC" Then
'          rstdestino("estado_desafectado") = ""
          rstdestino("estado_codigo") = "ANI"
'          rstdestino("estado_devengado") = "S" 02/07/01
'          rstdestino("estado_anulado") = ""
'          rstdestino("codigo_tipo") = "DEI" 02/07/01
          rstdestino("monto_recaudado_dolares") = 0
        End If
      End If
      rstdestino.Update
'      Print rstdestino!ingreso_codigo_anterior
'      Print rstdestino!monto_recaudado
      cod_ant = 0
      org_ant = ""
      
      'Call f_actual_rec(rstdestino!org_codigo, rstdestino!ingreso_codigo_anterior)
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    If (VAR_CODTIPO = "DVI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        rstdestino!estado_codigo = "DVI"
      End If
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    '======= fin Actualiza el monto recaudado  ==========

    '======= ini Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    If VAR_CODTIPO = "REC" Or VAR_CODTIPO = "DYR" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    If VAR_CODTIPO = "ANI" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    '======= fin Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    'LblMensaje.Caption = "El proceso concluyó exitosamente, gracias"
    'Frmmensaje.Visible = False
    db.CommitTrans
  'End If
'  'marca1 = Ado_datos.Recordset.Bookmark
'  rs_datos.Update
'  rs_datos.Requery
'  Set Ado_datos.Recordset = rs_datos
'  If rs_datos.RecordCount > 0 Then
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"

End Sub

Private Function DESCAUX(VARAUX As String, VARCODIG As String)
    Set rsAuxDetalle = New ADODB.Recordset
    If rsAuxDetalle.State = 1 Then rsAuxDetalle.Close
    Select Case VARAUX
        Case "01"
            rsAuxDetalle.Open "SELECT beneficiario_denominacion AS DESAUX2 FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT beneficiario_denominacion AS DESAUX FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' "
        Case "02"
            rsAuxDetalle.Open "SELECT cta_descripcion AS DESAUX2 FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT cta_descripcion AS DESAUX FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "' "
        Case "03"
            rsAuxDetalle.Open "SELECT pro_codigo_det_descripcion AS DESAUX2 FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT pro_codigo_det_descripcion AS DESAUX FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "' "
        Case "04"
            rsAuxDetalle.Open "SELECT unidad_descripcion AS DESAUX2 FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "05"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "06"
            rsAuxDetalle.Open "SELECT depto_descripcion AS DESAUX2 FROM gc_departamento where depto_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT depto_descripcion AS DESAUX FROM gc_departamento where depto_codigo = '" & VARCODIG & "' "
        Case "07"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "08"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "09"
            rsAuxDetalle.Open "SELECT Org_descripcion AS DESAUX2 FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT Org_descripcion AS DESAUX FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' "
        Case "10"
            'db.Execute "SELECT impuesto_descripcion AS DESAUX FROM fc_impuestos where impuesto_codigo = '" & VARCODIG & "' "
        Case "11"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "12"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "00"
            DESAUX = ""
    End Select
    If rsAuxDetalle.RecordCount > 0 Then
      DESAUX = RTrim(rsAuxDetalle!DESAUX2)
    Else
      DESAUX = ""
    End If
End Function

'Private Sub f_actual_rec(org, codant)
'  Dim acumDl As Double
'  Dim rsrecalc As New ADODB.Recordset
'  Set rsrecalc = New ADODB.Recordset
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select sum(monto_dolares) as acumDl from fo_ingresos_cabecera where org_codigo = '" & org & "' and  correlativo_anterior = '" & codant & "' and codigo_tipo = 'REC' and estado_recaudado= 'S'", db, adOpenKeyset, adLockReadOnly
'  If rsrecalc.RecordCount > 0 Then
'    acumDl = IIf(IsNull(rsrecalc!acumDl), 0, rsrecalc!acumDl)
'  Else
'    acumDl = 0
'  End If
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select * from fo_ingresos_cabecera where org_codigo = '" & org & "' and correlativo_ingreso = '" & codant & "' ", db, adOpenKeyset, adLockOptimistic
'  If rsrecalc.RecordCount > 0 Then
'    rsrecalc!monto_recaudado_dolares = acumDl
'  End If
'  rsrecalc.Update
'  If rsrecalc.State = 1 Then rsrecalc.Close
'
'End Sub

Private Sub graba_proyecto()
'    Select Case Ado_datos.Recordset!unidad_codigo
'        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
'            VAR_PROY = 12
'        Case "UCOM"
'            VAR_PROY = 17
'        Case "DVTA"
'            VAR_PROY = 18
'
'    End Select
'
'    Set rs_aux1 = New ADODB.Recordset
'    If rs_aux1.State = 1 Then rs_aux1.Close
'    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    Else
'        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
'           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & Ado_datos.Recordset!ges_gestion & ", 'APR', '" & GlUsuario & "', '" & Date & "')"
'    End If
End Sub

Private Sub graba_ingreso()
    '======= Ini grabado de datos
   'swgraba = 0
   'Call valida
    
'   If swgraba = 1 Then
'      FraOpciones2.Visible = False
'      fraOpciones.Visible = True
'      FraIngresosNav.Enabled = True
'      FraIngresosDat.Enabled = False
      
      'If v_añadir = 1 Then
        'EFECTIVO o a CREDITO
         'db.BeginTrans
         'Call add_correl
         Set rstdestino = New ADODB.Recordset
         If VAR_TIPOV = "V" Or VAR_TIPOV = "C" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
         Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
         End If
         
         If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo
            'Call add_correl
         Else
            Call add_correl
            'EXEPCION PARA GRABAR CONTRATO EN INGRESOS
             rstdestino.AddNew
             rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
             rstdestino("ingreso_codigo") = correlativo1
             rstdestino("org_codigo") = VAR_ORG
             rstdestino("ingreso_codigo_anterior") = VAR_ORG
             'rstdestino("Codigo_tipo") = "DEI"
             rstdestino("proceso_codigo") = "FIN"
             rstdestino("subproceso_codigo") = "FIN-01"
             rstdestino("etapa_codigo") = "FIN-01-02"
             rstdestino("clasif_codigo") = "ADM"
             rstdestino("doc_codigo") = "R-110"
             rstdestino("doc_numero") = correlativo1
             rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos16.Recordset("unidad_codigo")
             rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos16.Recordset("solicitud_codigo")
             If VAR_COD4 = "DVTA" Then
                rstdestino("solicitud_tipo") = "3"
                VAR_PARTIDA = "11200"
                rstdestino("tipo_comp") = "DEY"
                rstdestino("Codigo_tipo") = "DEY"
             Else
                rstdestino("solicitud_tipo") = "10"
                VAR_PARTIDA = "11300"
                rstdestino("tipo_comp") = "DEI"
                rstdestino("Codigo_tipo") = "DEI"
             End If
             If VAR_COD4 = "DNMAN" Then
                rstdestino("solicitud_tipo") = "10"
                VAR_PARTIDA = "11320"
             End If
             If VAR_COD4 = "DNREP" Then
                rstdestino("solicitud_tipo") = "7"
                VAR_PARTIDA = "11330"
             End If
             If VAR_COD4 = "DNMOD" Then
                rstdestino("solicitud_tipo") = "9"
                VAR_PARTIDA = "11340"
             End If
             'OJO JQA
             rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
             rstdestino("fecha_ingreso") = Date
             rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
             rstdestino("tipo_moneda") = VAR_MONEDA
             'VAR_MONEDA = "BOB"
             rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA       'Ado_datos.Recordset("cobranza_observaciones")
             'CAMBIAR FTE
             rstdestino("fte_codigo") = VAR_FTE
             'CAMBIAR RUBROS
             rstdestino("rubro_codigo") = VAR_PARTIDA
             'CAMBIAR RUBROS
             rstdestino("cheque_o_trf") = "T"
             'CAMBIAR CTA
             rstdestino("cta_codigo") = VAR_CTA
             If VAR_CTA = "NN" Then
                rstdestino("Bco_codigo") = "BCP"
             Else
                rstdestino("Bco_codigo") = "BMS"
             End If
             'CAMBIAR CTA
             rstdestino("numero_documento") = VAR_COD1
             rstdestino("unidad_codigo_ant") = VAR_CITE
             rstdestino("monto_dolares") = VAR_DOL2 * 12
             rstdestino("monto_bolivianos") = VAR_BS2 * 12
             rstdestino("monto_recaudado_dolares") = VAR_DOL2 * 12 'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
             rstdestino("monto_recaudado_bolivianos") = VAR_BS2 * 12   'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
             rstdestino("convenio_codigo") = "NN"
             rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
             rstdestino("estado_CODIGO") = "APR"
             'rstdestino("estado_codigo_dr") = "DEI"
    
             rstdestino("usr_CODIGO") = glusuario
             rstdestino("fecha_registro") = Date
             rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
             
             rstdestino.Update
             VAR_CODANT = rstdestino!ingreso_codigo
             VAR_ORG = rstdestino!org_codigo
             VAR_FTE = rstdestino!fte_codigo
             If rstdestino.State = 1 Then rstdestino.Close
             If VAR_TIPOV = "V" Or VAR_TIPOV = "C" Then
                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
             Else
                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
             End If
         End If
         Call add_correl
         ' OJO CAMBIAR FINANCIADOR WWWWWWWWWWWWWWWWWWWWW
         rstdestino.AddNew
         rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
         rstdestino("ingreso_codigo") = correlativo1
         'VAR_CODANT = correlativo1
         'CAMBIAR org_codigo
         'rstdestino("org_codigo") = "111"
         'VAR_ORG = "111"
         rstdestino("org_codigo") = VAR_ORG
         'CAMBIAR org_codigo
         'CAMBIAR COD ingreso_codigo_anterior
         rstdestino("ingreso_codigo_anterior") = VAR_CODANT
         'CAMBIAR COD ingreso_codigo_anterior
         'CAMBIAR DEI O REC
         rstdestino("Codigo_tipo") = "REC"
         VAR_CODTIPO = "REC"
         'CAMBIAR DEI O REC
         rstdestino("proceso_codigo") = "FIN"
         rstdestino("subproceso_codigo") = "FIN-01"
         rstdestino("etapa_codigo") = "FIN-01-02"
         rstdestino("clasif_codigo") = "ADM"
         rstdestino("doc_codigo") = "R-110"
         rstdestino("doc_numero") = correlativo1
         rstdestino("unidad_codigo") = VAR_COD4
         rstdestino("solicitud_codigo") = VAR_SOL
         If VAR_COD4 = "DVTA" Then
            rstdestino("solicitud_tipo") = "3"
            VAR_PARTIDA = "11200"
         Else
            rstdestino("solicitud_tipo") = "10"
            VAR_PARTIDA = "11300"
         End If
         If VAR_COD4 = "DNMAN" Then
            rstdestino("solicitud_tipo") = "10"
            VAR_PARTIDA = "11320"
         End If
         If VAR_COD4 = "DNREP" Then
            rstdestino("solicitud_tipo") = "7"
            VAR_PARTIDA = "11330"
         End If
         If VAR_COD4 = "DNMOD" Then
            rstdestino("solicitud_tipo") = "9"
            VAR_PARTIDA = "11340"
         End If
         'OJO JQA
         rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
         rstdestino("fecha_ingreso") = Date
         rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
         rstdestino("tipo_moneda") = VAR_MONEDA
         'VAR_MONEDA = "BOB"
         rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA       'Ado_datos.Recordset("cobranza_observaciones")
         'VAR_GLOSA = "INGRESO POR: " + Ado_datos.Recordset("cobranza_observaciones")
         If Ado_datos16.Recordset("venta_tipo") = "E" Then
            rstdestino("tipo_comp") = "DYR"
         Else
            rstdestino("tipo_comp") = "REC"
         End If
         'CAMBIAR FTE
         rstdestino("fte_codigo") = VAR_FTE
         'CAMBIAR FTE OJO JQAW
         'CAMBIAR RUBROS
         rstdestino("rubro_codigo") = VAR_PARTIDA
         'CAMBIAR RUBROS
         rstdestino("cheque_o_trf") = "T"
         'CAMBIAR CTA
         rstdestino("cta_codigo") = VAR_CTA
         If VAR_CTA = "2015046557-03-054" Then
            rstdestino("Bco_codigo") = "BCP"
         Else
            rstdestino("Bco_codigo") = "BMS"
         End If
         'CAMBIAR CTA
         NroFactura = Trim(Str(VAR_COD1))
         rstdestino("numero_documento") = NroFactura        'Ado_datos.Recordset!cobranza_nro_factura
         rstdestino("unidad_codigo_ant") = VAR_CITE
         rstdestino("monto_dolares") = VAR_DOL2
         rstdestino("monto_bolivianos") = VAR_BS2
         rstdestino("monto_recaudado_dolares") = VAR_DOL2   'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
         rstdestino("monto_recaudado_bolivianos") = VAR_BS2     'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
         rstdestino("convenio_codigo") = "NN"
         rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
         rstdestino("estado_CODIGO") = "APR"
         'rstdestino("estado_codigo_dr") = "DEI"

         rstdestino("usr_CODIGO") = glusuario
         rstdestino("fecha_registro") = Date
         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
         
         rstdestino.Update
         If rstdestino.State = 1 Then rstdestino.Close
        'db.CommitTrans
          
'          If rstIngresos.State = 1 Then rstIngresos.Close
'          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'          rstIngresos.Sort = "ingreso_codigo"
'          rstIngresos.Requery
          
'          rstIngresos.Requery
'          Set AdoIngresos.Recordset = rstIngresos
'          AdoIngresos.Refresh
'          AdoIngresos.Recordset.Find "ultimo = 'S'"
'          If Not (AdoIngresos.Recordset.EOF) Then
'            marca1 = AdoIngresos.Recordset.Bookmark
'            AdoIngresos.Recordset("ultimo") = "N"
'            AdoIngresos.Recordset.Update
'          End If

'          AdoIngresos.Recordset.Move marca1 - 1

'          marca1 = 0
      'End If
'   Else*
'      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
''      FraOpciones2.Visible = False
''      FraOpciones.Visible = True
''      FraIngresosNav.Enabled = True
''      FraIngresosDat.Enabled = False
''      AdoIngresos.Refresh
'   End If
'   LblAccion = ""
'AAQQQQQUIIIIIIIIII    JQA

End Sub

Private Sub add_correl()
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' ", db, adOpenDynamic, adLockOptimistic
  If rstcorrel_ing.RecordCount = 0 Then
     VAR_ORG = "112"
     VAR_FTE = "10"
     If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
     rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "'  ", db, adOpenDynamic, adLockOptimistic
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = VAR_ORG   'Trim(DtCorg_codigo.Text)
'     rstcorrel_ing("ges_gestion") = Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
'     rstcorrel_ing("fte_codigo") = "10"
'     'rstcorrel_ing("correlativo") = 1
'     rstcorrel_ing("correlativo_ingreso") = 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
  Else
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
  End If
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close

End Sub

'Private Sub CmdGrabaCobranza()
'    If swnuevo = 1 Then
''      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
''      Set Ado_datos16.Recordset = rstdestino
''      Ado_datos16.Recordset.AddNew
'      Ado_datos16.Recordset!correl_venta = Val(lblcorrelVenta.Caption)
'      Ado_datos16.Recordset!venta_codigo = Val(TxtNroVenta.Text)
'      Ado_datos16.Recordset!ges_gestion = Year(Date)    'Trim(LblGestion.Caption)
'    End If
'      Ado_datos16.Recordset!beneficiario_codigo = dtc_codigo2A.Text                                 'Codigo Beneficiario/Cliente
'      Ado_datos16.Recordset!ci = dtc_codigo4A.Text                                                     'Codigo Cobrador
'      Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text + " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
'      Ado_datos16.Recordset!deuda_cobrada = Val(TxtMonto.Text)                                  'Monto Cobrado
'      Ado_datos16.Recordset!deuda_cobrada_dol = Val(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
'      Ado_datos16.Recordset!fecha_cobranza = DTPFechaCobro.Value                                'Fecha de Cobranza
'      'Call acumulaMont(Ado_datos16.Recordset!ges_gestion, Ado_datos16.Recordset!correl_venta, Ado_datos16.Recordset!venta_codigo)
'      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))
'
'      Ado_datos16.Recordset!obs_cobranza = TxtObs
'      Ado_datos16.Recordset!nro_cmpbte = Trim(TxtCmpbte)
'      Ado_datos16.Recordset!usr_usuario = GlUsuario
'      Ado_datos16.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'      Ado_datos16.Recordset.Update
'End Sub

'Private Sub CmdModDetalle_Click()
'  FraDetalle.Visible = True
'  FraDetalle.Enabled = True
'  txtnosolicitud1.Enabled = False
'  txtcorrdet.Enabled = False
'  dtccodpar.SetFocus
'  CmdGraDetalle.Enabled = True
'  CmdAddDetalle.Enabled = False
'  CmdModDetalle.Enabled = False
'  CmdSalDetalle.Enabled = False
'  CmdCanDetalle.Enabled = True
'  swgrabar = 2
'End Sub

'Private Sub CmdGraDetalle_Click()
'    If swgrabar = 1 Then
'        Dim rstdestino As New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle_correl where formulario = '" & "F11" & "' and correl_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("correl_solicitud_detalle") = rstdestino("correl_solicitud_detalle") + 1
'        Else
'            rstdestino.AddNew
'            rstdestino("formulario") = "F11"
'            rstdestino("correl_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correl_solicitud_detalle") = 1
'        End If
'        correldetalle = rstdestino("correl_solicitud_detalle")
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correlativo_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        rstdestino.AddNew
'        rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'        rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'        rstdestino("correlativo_detalle") = correldetalle
'        rstdestino("Par_codigo") = dtccodpar.Text
'        rstdestino("Importe_nacional") = txtsolpeso.Text
'        rstdestino("formulario") = "F11"
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    If swgrabar = 2 Then
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoDetalleSolicitud.Recordset("ges_gestion") & "' and correlativo_solicitud = " & adoDetalleSolicitud.Recordset("correlativo_solicitud") & " and correlativo_detalle =" & adoDetalleSolicitud.Recordset("correlativo_detalle"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'            rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correlativo_detalle") = correldetalle
'            rstdestino("Par_codigo") = dtccodpar.Text
'            rstdestino("Importe_nacional") = txtsolpeso.Text
'            rstdestino("formulario") = "F11"
'            rstdestino.Update
'        End If
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    CmdGraDetalle.Enabled = False
'    CmdAddDetalle.Enabled = True
'    CmdModDetalle.Enabled = True
'    CmdSalDetalle.Enabled = True
'    CmdCanDetalle.Enabled = False
'    FraDetalle.Enabled = False
'    swgrabar = 0
'End Sub

Private Sub CmdNOunidad_Click()
    swunidad = 0
    Frmunidad.Visible = False
End Sub

Private Sub CmdOKunidad_Click()
    swunidad = 1
        If swunidad = 1 Then
            Dim rstpagos As New ADODB.Recordset
            Set rstpagos = New ADODB.Recordset
            If rstpagos.State = 1 Then rstpagos.Close
            rstpagos.Open "select * from pagos where GES_gestion = '5000'", db, adOpenKeyset, adLockOptimistic
            rstpagos.AddNew
                rstpagos("ges_gestion") = Ado_datos.Recordset("ges_gestion")
                rstpagos("org_codigo") = DataCombo1.Text   'Ado_datos.Recordset("formulario")
                rstpagos("codigo_pago") = "" 'genera jorge
                rstpagos("codigo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
                rstpagos("formulario") = Ado_datos.Recordset("formulario")
                rstpagos("codigo_unidad") = Ado_datos.Recordset("codigo_unidad")
                rstpagos("monto_bolivianos") = Ado_datos.Recordset("monto_bolivianos")
                rstpagos("estado_compromiso") = "N"
                rstpagos("justificacion") = Ado_datos.Recordset("justificacion_solicitud")
             rstpagos.Update
        End If
End Sub

Private Sub CmdGrabaCobro_Click()
End Sub

'Private Sub CmdGrabaDet_Click()
''If dtc_desc12 = "" Then
''    MsgBox "Debe Elejir un Descuento X Tipo de Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
''    Exit Sub
''  End If
'  If dtc_codigo15 = "" Then
'     MsgBox "Debe Elejir un Producto para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
''  If dtc_desc13 = "" Then
''    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atención"
''    Exit Sub
''  End If
'    'If Val(dtc_stocktotal15.Text) >= Val(TxtCantidad.Text) Then
'    '    VAR_PARTIDA = "OK"
'    If Val(Dtc_Stock13.Text) >= Val(TxtCantidad.Text) Or Dtc_partida15.Text = "43340" Then
'          'fraOpciones.Visible = True
'          'FraGrabarCancelar.Visible = False
'          'TxtNroVenta.Enabled = True
'          FrmEdita.Enabled = False
'        '  DtGListaN.Enabled = True
'          'cmdElige.Enabled = False
'        '  dtc_codigo15.Visible = False
'        '  dtc_desc15.Visible = False
'          'txt_descripcion_venta.Enabled = False
'        If swnuevo = 1 Then
'          'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
'          ado_datos14.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
'          ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset("ges_gestion")
'        End If
'          'ado_datos14.Recordset!nro_licitacion = dtc_partida15.Text                       'Compra ??
'          'ado_datos14.Recordset!nro_adjudica = 0 'Trim(DtcNroAdjudica.Text)                 'Codigo de Adjudicacion
'          ado_datos14.Recordset!bien_codigo = Trim(dtc_codigo15.Text)                       'Codigo Bien (Equipo, Producto, etc)
'          ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
'          ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
'          ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
'          ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
'          ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
'          ado_datos14.Recordset!almacen_codigo = IIf(dtc_codigo13.Text = "", "0", dtc_codigo13.Text)
'          If TxtCantidad.Text = "" Then
'            TxtCantidad.Text = "1"
'          End If
'          ado_datos14.Recordset!venta_det_cantidad = Val(IIf(TxtCantidad = "", 1, TxtCantidad)) 'Cantidad Vendida
'          'ado_datos14.Recordset!codigo_solicitud = 0                                     'Nro.Solicitud de compra
'          ado_datos14.Recordset!venta_precio_unitario_bs = CDbl(TxtPrecioU.Text)             'Precio Unitario de Venta
'          If CDbl(TxtDescuento) > 0 Then
'            ado_datos14.Recordset!venta_descuento_bs = CDbl(TxtDescuento.Text)      'Dcto por producto CON DESCUENTO
'            ado_datos14.Recordset!venta_descuento_dol = Val(TxtDescuento) / GlTipoCambioMercado
'          Else
'            'ado_datos14.Recordset!descuento_venta = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) * (CDbl(Dtc_aux12)) 'Dcto por producto DE LA TABLA
'            TxtDescuento.Text = "0"
'            ado_datos14.Recordset!venta_descuento_bs = 0
'            ado_datos14.Recordset!venta_descuento_dol = 0
'          End If
'          ado_datos14.Recordset!venta_precio_total_bs = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) - (CDbl(TxtDescuento)) 'Precio Total Producto
'          'If Val(lbltipo_Cambio) = 0 Then lbltipo_Cambio = 1
'          ado_datos14.Recordset!venta_precio_unitario_dol = CDbl(TxtPrecioU.Text) / GlTipoCambioMercado                'Precio Unitario Dolares
'          ado_datos14.Recordset!venta_precio_total_dol = (ado_datos14.Recordset!venta_precio_total_bs) / GlTipoCambioMercado
'          'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
'          ado_datos14.Recordset!modelo_codigo = Txt_modelo.Text
'          ado_datos14.Recordset!modelo_codigo1 = Txt_modelo1.Text
'          ado_datos14.Recordset!modelo_codigo_h = Txt_modelo2.Text
'          ado_datos14.Recordset!modelo_codigo_x = Txt_modelo3.Text
'          If OpMod1.Value = True Then
'            ado_datos14.Recordset!modelo_elegido = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido = "N"
'          End If
'          If OpMod2.Value = True Then
'            ado_datos14.Recordset!modelo_elegido_h = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido_h = "N"
'          End If
'          If OpMod2.Value = True Then
''            ado_datos14.Recordset!modelo_elegido_x = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido_x = "N"
'          End If
'          ado_datos14.Recordset!estado_codigo = "REG"
'          ado_datos14.Recordset!usr_codigo = GlUsuario
'          ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'          ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'          ado_datos14.Recordset.Update
'        'db.CommitTrans
'
'        'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
'        Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = False
'        FraNavega.Enabled = True
'        FrmDetalle.Enabled = True
'        'FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
'        FrmABMDet.Visible = True
'        FrmABMDet2.Visible = True
'        Call OptFilGral1_Click
'        If swnuevo = 1 Then
'          'Call abre_ventas_det
'          'rs_datos14.Requery
'          'ado_datos14.Refresh
'          'ado_datos14.Recordset.MoveLast
'
'        End If
'        swnuevo = 0
'    Else
'        MsgBox "Saldo Insuficiente en Almacen Origen, debe realizar Transferencia de otro Almacen, Luego Intente nuevamente !..."
'    End If
'  'Else
'  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
'  'End If
'End Sub

'Private Sub BtnModDetalle_Click()
'  If Ado_datos16.Recordset.RecordCount > 0 Then
'
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmCabecera.Visible = True
''    BtnImprimir2.Visible = False
''    BtnImprimir3.Visible = False
'  Else
'    MsgBox "No existen datos de la Venta, Verifique por favor !! ", vbExclamation, "Atención!"
'  End If
'End Sub

Private Sub BtnSalir2_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmCabecera.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
End Sub

Private Sub BtnSalir3_Click()
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmEdita.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
End Sub

Private Sub BtnSalir1_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If

End Sub

Private Sub cmd_benef_Click()
    Set rs_datos8 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_beneficiario where tipoben_codigo <> '0' and tipoben_codigo <> '1' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    If Ado_datos8.Recordset.RecordCount > 0 Then
        dtc_desc8.BoundText = dtc_codigo8.BoundText
        FraGrabarCancelar.Enabled = False
        frm_benef.Visible = True
    End If
End Sub

Private Sub cmd_moneda1_LostFocus()
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria where tipo_moneda = '" & cmd_moneda1.Text & "' ", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
    dtc_ctaDes.BoundText = dtc_cta.BoundText
End Sub

Private Sub cmd_moneda2_LostFocus()
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_cuenta_bancaria where tipo_moneda = '" & cmd_moneda2.Text & "' ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub CmdFoto_Click()
'    Frm_Imprime_Factura.Show

    On Error GoTo QError
    Set fs = New FileSystemObject   'Creamos la Nueva referencia Fso
    
    Set rs_aux6 = New ADODB.Recordset     'Iniciales del Cliente - gc_beneficiario
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        db.Execute "update ao_ventas_cobranza set beneficiario_iniciales = '" & rs_aux6!beneficiario_iniciales & "'   Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
    End If
    'If Ado_datos.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
    If Ado_datos.Recordset!archivo_foto_cargado = "N" Or IsNull(Ado_datos.Recordset!archivo_foto_cargado) Then
      NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      DirOrigen = App.Path & "\CLIENTES\"
      DirDestino = App.Path & "\CLIENTES\"
      'DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG"       'Ado_datos.Recordset!cobranza_nro_factura        'ARCHIVO_Foto
      Ado_datos.Recordset!ARCHIVO_Foto = Trim(Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG")
      Ado_datos.Recordset!archivo_foto_cargado = "S"
      
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "Q_R"
''      If GlServidor = "SERVIDOR2" Then
''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
''      Else
'         e = NombreCarpeta
''      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          DirOrigen = App.Path & "\CLIENTES\"
          DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos.Recordset!ARCHIVO_Foto
          frmBeneficiario_Admin.Adolista.Recordset!archivo_foto_cargado = "S"
          
    '      Frmexporta.DirDestino.Path = NombreCarpeta
    '      GlArch = "Q_R"
    ''      If GlServidor = "SERVIDOR2" Then
    ''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
    ''      Else
    '         e = NombreCarpeta
    ''      End If
    '      Frmexporta.DirDestino2.Path = e
    '      Frmexporta.Show vbModal      End If
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SERVIDOR2" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" + Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        'ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(rs_aux6!beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
        ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
'    End If
    'ARCH_FOTO = App.Path + "\" + "CLIENTES" + "\" + Ado_datos.Recordset!beneficiario_codigo + "\" + Ado_datos.Recordset("beneficiario_codigo") + "-FOTO.JPG"
    CodBenef = Ado_datos.Recordset!cobranza_codigo
    'If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where beneficiario_codigo= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
    If Guardar_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
        MsgBox "Se cargo la Imagen Correctamente !!"
        Exit Sub
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    db.RollbackTrans
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub BntImprimir3_Click()
    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas_dol.rpt"
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If
End Sub

'Private Sub Command111_Click()
'   Call generarRepRecibo
'End Sub

'' Genera reporte de recibo
'Private Sub generarRepRecibo()
''    ' Verifica si codigo y numero son validos para recibo.
''    If Label38.Caption <> "" And TxtCmpbte2.Text <> "" And Label38.Caption <> "R-101" Then
''        Dim iResult As Integer
''        Dim montoLiteral As String
''        Dim Monto As Double
''        Monto = 0
''        If TxtCmpbte2.Text <> "" Then Monto = TxtCmpbte2.Text
''        If TxtDscto2.Text <> "" Then Monto = Monto + CInt(TxtDscto2.Text)
''        montoLiteral = Monto
''        montoLiteral = Literal(montoLiteral)
''        crRecibo.WindowShowPrintSetupBtn = True
''        crRecibo.WindowShowRefreshBtn = True
''        crRecibo.ReportFileName = App.Path & "\Reportes\Ventas\rr_recibo_oficial.rpt"
''        crRecibo.StoredProcParam(0) = Label38.Caption ' codigo
''        crRecibo.StoredProcParam(1) = TxtCmpbte2.Text ' numero
''        crRecibo.StoredProcParam(2) = "Bs. " + TxtMonto02.Text + "(" + montoLiteral + " BOLIVIANOS)" ' monto
''        crRecibo.WindowState = crptMaximized
''        iResult = crRecibo.PrintReport
''        If iResult <> 0 Then
''              MsgBox crRecibo.LastErrorNumber & " : " & crRecibo.LastErrorString, vbExclamation + vbOKOnly, "Error"
''        End If
''   Else
''       MsgBox "No se puede generar reporte por falta de codigo y numero de recibo."
''   End If
'End Sub

Private Sub DataCombo8_Change()
   If Trim(DataCombo8) = "CHEQUE" Then
      Label47.Caption = "Nro. Cheque"
   Else
      Label47.Caption = "Nro. Recibo"
   End If
End Sub

Private Sub dtc_aux8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_aux8.BoundText
    dtc_codigo8.BoundText = dtc_aux8.BoundText
End Sub

Private Sub dtc_codigo4A1_Click(Area As Integer)
    dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText
End Sub



Private Sub Command1_Click()
   fw_seg_cobranza_parametro.Show vbModal
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    dtc_aux8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_cta_Click(Area As Integer)
    dtc_ctaDes.BoundText = dtc_cta.BoundText
End Sub

Private Sub dtc_ctades_Click(Area As Integer)
    dtc_cta.BoundText = dtc_ctaDes.BoundText
End Sub

Private Sub dtc_desc4A1_Click(Area As Integer)
    dtc_codigo4A1.BoundText = dtc_desc4A1.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    dtc_aux8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    DataCombo2.Text = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DataCombo1.Text = DataCombo2.BoundText
End Sub

Private Sub dtccodmanejo_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodmanejo.BoundText
    DtCDescripcion.BoundText = dtccodmanejo.BoundText
    dtcunidadmedida.BoundText = dtccodmanejo.BoundText
    dtccodpeso.BoundText = dtccodmanejo.BoundText
End Sub

Private Sub dtccodpeso_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodpeso.BoundText
    DtCDescripcion.BoundText = dtccodpeso.BoundText
    dtcunidadmedida.BoundText = dtccodpeso.BoundText
    dtccodmanejo.BoundText = dtccodpeso.BoundText
End Sub


Private Sub dtccodpar_Click(Area As Integer)
    dtcdescripar.Text = dtccodpar.BoundText
End Sub

Private Sub dtccodpoa_Click(Area As Integer)
    dtcdespoa.Text = dtccodpoa.BoundText
End Sub

Private Sub dtccodpuesto_Click(Area As Integer)
    dtcdenopuesto.Text = dtccodpuesto.BoundText
End Sub

Private Sub dtccodtipoid_Click(Area As Integer)
    dtcdescrtipoid.BoundText = dtccodtipoid.BoundText
End Sub

Private Sub dtccoduni_Click(Area As Integer)
    dtcdescripuni.Text = dtccoduni.BoundText
End Sub

Private Sub dtccorrcompromiso_Click(Area As Integer)
    dtcfechacompromiso.BoundText = dtccorrcompromiso.BoundText
End Sub

Private Sub dtccorrsol_Click(Area As Integer)
 dtcfechasol.BoundText = dtccorrsol.BoundText
End Sub

Private Sub dtcdenominacionruc_Click(Area As Integer)
    dtcnroruc.BoundText = dtcdenominacionruc.BoundText
End Sub

Private Sub dtcdenopuesto_Click(Area As Integer)
    dtccodpuesto.Text = dtcdenopuesto.BoundText
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
    DtCCodigo.BoundText = DtCDescripcion.BoundText
    dtcunidadmedida.BoundText = DtCDescripcion.BoundText
    dtccodmanejo.BoundText = DtCDescripcion.BoundText
    dtccodpeso.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub dtcdescripar_Click(Area As Integer)
    dtccodpar.Text = dtcdescripar.BoundText
End Sub

Private Sub dtcdescripuni_Click(Area As Integer)
    dtccoduni.Text = dtcdescripuni.BoundText
End Sub

Private Sub dtcdescrtipoid_Click(Area As Integer)
    dtccodtipoid.BoundText = dtcdescrtipoid.BoundText
End Sub

Private Sub dtcfechacompromiso_Click(Area As Integer)
    dtccorrcompromiso.BoundText = dtcfechacompromiso.BoundText
End Sub

Private Sub dtcfechasol_Click(Area As Integer)
    dtccorrsol.BoundText = dtcfechasol.BoundText
End Sub

Private Sub dtcnroruc_Click(Area As Integer)
    dtcdenominacionruc.Text = dtcnroruc.BoundText
End Sub


Private Sub dtc_desc4A_Click(Area As Integer)
    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
End Sub

Private Sub dtctipodoc_Click(Area As Integer)
    dtcdenodoc.Text = dtctipodoc.BoundText
End Sub

Private Sub dtcunidadmedida_Click(Area As Integer)
    DtCCodigo.BoundText = dtcunidadmedida.BoundText
    DtCDescripcion.BoundText = dtcunidadmedida.BoundText
    dtccodmanejo.BoundText = dtcunidadmedida.BoundText
    dtccodpeso.BoundText = dtcunidadmedida.BoundText
End Sub

Private Sub dtcdespoa_Click(Area As Integer)
    dtccodpoa.Text = dtcdespoa.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub DTPFechaCobro02_LostFocus()
    If (CDate(DTPFechaCobro2.Value) > CDate(DTPFechaCobro02.Value)) Then
        MsgBox "La <<Fecha Cobranza2>> No puede ser MENOR a la <<Fecha Cobranza1>>, Vuelva a Intentar !! ", vbExclamation, "Atención!"
        DTPFechaCobro02.SetFocus
    End If
End Sub
'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

Private Sub btnSalirPanel_Click()
    FraImprime.Visible = False
    fraOpciones.Visible = True
    FrmDetalle.Enabled = True
    FraNavega2.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = 0
    parametro = Aux
     db.Execute "UPDATE  ao_ventas_cabecera SET ao_ventas_cabecera.edif_codigo_corto = gc_edificaciones.edif_codigo_corto FROM ao_ventas_cabecera INNER JOIN gc_edificaciones ON ao_ventas_cabecera.edif_codigo = gc_edificaciones.edif_codigo where (ao_ventas_cabecera.edif_codigo_corto Is Null)"
     
     db.Execute "UPDATE ao_ventas_cobranza_prog SET ao_ventas_cobranza_prog.cobranza_fecha_conformidad  = to_cronograma_diario_final.fecha_conformidad FROM ao_ventas_cobranza_prog INNER JOIN to_cronograma_diario_final ON ao_ventas_cobranza_prog.fmes_plan = to_cronograma_diario_final.fmes_plan AND ao_ventas_cobranza_prog.venta_codigo = to_cronograma_diario_final.venta_codigo where ao_ventas_cobranza_prog.cobranza_fecha_conformidad Is Null And to_cronograma_diario_final.fecha_conformidad Is Not Null "

    db.Execute "UPDATE ao_ventas_cobranza_prog SET ao_ventas_cobranza_prog.doc_numero = to_cronograma_diario_final.doc_numero FROM ao_ventas_cobranza_prog INNER JOIN to_cronograma_diario_final ON ao_ventas_cobranza_prog.fmes_plan = to_cronograma_diario_final.fmes_plan AND ao_ventas_cobranza_prog.venta_codigo = to_cronograma_diario_final.venta_codigo WHERE ao_ventas_cobranza_prog.doc_numero ='0' AND  to_cronograma_diario_final.doc_numero > '0' "
    
    mbDataChanged = False
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
    BtnBuscar.Visible = False
    BtnImprimir.Visible = False
'    BtnImprimir1.Visible = False
'    BtnImprimir3.Visible = False
    BtnImprimir2.Visible = True
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
    'db.Execute "fp_saldos"
    If glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "APALACIOS" Or glusuario = "JCASTRO" Or glusuario = "RCUELA" Or glusuario = "CSALINAS" Or glusuario = "PLOPEZ" Or glusuario = "VBELLIDO" Then
        OptFilGral00.Visible = True
        'SSTab1.Tab = 0
        'SSTab1.TabEnabled(0) = True
        'SSTab1.TabEnabled(1) = True
        'SSTab1.TabEnabled(2) = True
    Else
        OptFilGral00.Visible = False
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = True
    End If
''    FrmEdita.Enabled = False
''    Cmd_Cliente.Visible = False
    swnuevo = 0
    'FraNavega.Caption = lbl_titulo.Caption
    'lbl_titulo2.Caption = lbl_titulo.Caption
    'lbl_titulo1.Caption = lbl_titulo.Caption
        Call SeguridadSet(Me)
End Sub

Private Sub OptFilGral10_Click()
   '===== Proceso para filtrado de datos(registros Pendientes para Cobrar)
    Set rs_datos06 = New Recordset
    If rs_datos06.State = 1 Then rs_datos06.Close
    rs_datos06.Open "select * From ao_ventas_cobranza_det where cobranza_codigo = " & NRO_COBR & "  ", db, adOpenKeyset, adLockOptimistic
    'queryinicial2 = "select * From ao_ventas_cobranza_det where cobranza_codigo = " & NRO_COBR & "  "
    'rs_datos06.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    'rs_datos02.Sort = "cobranza_fecha_fac"
    Set Ado_datos06.Recordset = rs_datos06.DataSource
    Set dg_datos06.DataSource = Ado_datos06.Recordset
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador en Fac.
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    'rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
   ' dtc_desc4A.BoundText = dtc_codigo4A.BoundText
    
    Set rs_datos4A = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador
    If rs_datos4A.State = 1 Then rs_datos4A.Close
    'rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario ", db, adOpenStatic  '4333735
    'rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set ado_datos4A.Recordset = rs_datos4A
    'dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText

    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_tipo_transaccion order by trans_descripcion", db, adOpenStatic
    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    'dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "ac_tipo_compra_venta", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText

    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh
    
    'Solo para Equipos (*)
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    'rs_datos15.Open "select * from av_lista_productos where saldo_actual >= 0 order by DescDetalle ", db, adOpenKeyset, adLockReadOnly  'JQA 06/2008
    rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
    Set ado_datos15.Recordset = rs_datos15
    ado_datos15.Refresh
    
   'wwwwwwwwwwwwwwwwwwww
    'db.Execute "DELETE ao_ventas_cabecera where venta_codigo = 0 "
    'Call ABREVENTAS
  
'    Set rs_Dsctos = New ADODB.Recordset
'    If rs_Dsctos.State = 1 Then rs_Dsctos.Close
'    rs_Dsctos.Open "select * from ac_ventas_descuentos ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
'    Set AdoDsctos.Recordset = rs_Dsctos
'    AdoDsctos.Refresh

    Set rs_datos17 = New ADODB.Recordset
    If rs_datos17.State = 1 Then rs_datos17.Close
    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
    Set ado_datos17.Recordset = rs_datos17
    ado_datos17.Refresh
       
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
    'dtc_ctades.BoundText = dtc_cta.BoundText
    
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    'dtc_desc7.BoundText = dtc_codigo7.BoundText

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
'  End If
'  If glPersNew = "L" Then
'    frmmo_formulario_M1.Dtc_doc_id_lab = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell_lab = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2apell_lab = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre_lab = rs_Personal!pers_nombres
'  End If
'  If glPersNew = "PL" Then
'    frmeo_Larvas_mosquitos.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_Larvas_mosquitos.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_Larvas_mosquitos.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_Larvas_mosquitos.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
'  If glPersNew = "PMA" Then
'    frmeo_mosquito_adulto.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_mosquito_adulto.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_mosquito_adulto.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_mosquito_adulto.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
'  glPersNew = "N"

End Sub


Private Sub OptFilGral00_Click()
'===== Proceso para filtrado de TODOS
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
    BtnImprimir.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A'   AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset
End Sub

Private Sub OptFilGral01_Click()
  '===== Proceso para filtrado de Chuquisaca
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '1'  AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset
    
End Sub

Private Sub OptFilGral02_Click()
'===== Proceso para filtrado de La Paz
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
''    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '2' AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset

End Sub

Private Sub OptFilGral03_Click()
    '===== Proceso para filtrado de Cochabamba
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '3' AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset
    
End Sub

Private Sub OptFilGral04_Click()
    '===== Proceso para filtrado de ORURO
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '4'  AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset

End Sub

Private Sub OptFilGral05_Click()
    '===== Proceso para filtrado de POTOSI
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '5'  AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset

End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
     rs_datos9.Open " Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    If glusuario = "RCUELA" Then
        queryinicial1 = " select cobranza_fecha_sol,edif_codigo, cobranza_codigo, beneficiario_codigo_resp, cobranza_fecha_fac, cobranza_total_bs, cobranza_total_dol, doc_numero, cobranza_nro_factura, estado_codigo_fac, beneficiario_codigo From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' and doc_codigo_fac <> 'R-103' "     'ORDER BY cobranza_fecha_prog
        'queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' and doc_codigo_fac <> 'R-103' "     'ORDER BY cobranza_fecha_prog
    Else
        If glusuario = "HBUSTILLOS" Then
                queryinicial1 = " select cobranza_fecha_sol,edif_codigo, cobranza_codigo, beneficiario_codigo_resp, cobranza_fecha_fac, cobranza_total_bs, cobranza_total_dol, doc_numero, cobranza_nro_factura, estado_codigo_fac, beneficiario_codigo From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR'  AND estado_codigo_bco = 'REG' and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
                'queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR'  AND estado_codigo_bco = 'REG' and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
            Else
                If glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
                    queryinicial1 = " select cobranza_fecha_sol,edif_codigo, cobranza_codigo, beneficiario_codigo_resp, cobranza_fecha_fac, cobranza_total_bs, cobranza_total_dol, doc_numero, cobranza_nro_factura, estado_codigo_fac, beneficiario_codigo From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' "      'ORDER BY cobranza_fecha_prog
                    'queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' "      'ORDER BY cobranza_fecha_prog
                End If
            End If
    '    queryinicial = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
    End If
    'queryinicial = "select * From ao_ventas_cobranza  ORDER BY cobranza_fecha_prog "
    rs_datos.Open queryinicial1, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "cobranza_fecha_sol"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

'Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
' If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
'    'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
'       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
'    Else
'       Txtmonto_dolares_contra.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
''solo numeros y , .
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'End Sub
'
'Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
'    'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
'       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
'    Else
'       Txtmonto_dolares.Text = 0
'    End If
'  End If
'
'End Sub
'
'Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
'      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
'    Else
'      TxtMonto_bolivianos_contra.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
'      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
'    Else
'      TxtMonto_bolivianos.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
'    'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub txtsolpeso_KeyPress(KeyAscii As Integer)
''solo numeros y , .
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'End Sub
'
'Private Sub txtterref_KeyPress(KeyAscii As Integer)
'    If KeyAscii < 58 And KeyAscii > 47 Then
'        KeyAscii = Asc(UCase(Chr(0)))
'    Else
'        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "N" Or KeyAscii = 8 Then
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Else
'            KeyAscii = Asc(UCase(Chr(0)))
'            MsgBox "Debe escribir solo 'N' o 'S'", vbOKOnly, "Error..."
'        End If
'    End If
'End Sub

Private Sub cerea()
  txt_venta = " "
  dtc_codigo4.Text = " "
  Dtcpaternosol.Text = " "  'dtc_codigo4.BoundText
'  dtcmaternosol.Text = " "
'  dtcnombresol.Text = " "
  txtCantTotal = "0"
  TxtMontoBs = "0"
  TxtMontoUs = "0"
  TxtConcepto = ""
  dtc_codigo2 = ""
  dtc_desc2 = ""
  TxtTDC.Text = GlTipoCambioMercado ' GlTipoCambioOficial
  
'  DtCDenominacion_moneda = ""
'  TxtMonto_bolivianos = 0
'  Txtmonto_dolares = 0
'  TxtMonto_bolivianos_contra = 0
'  Txtmonto_dolares_contra = 0
'  DtCOrg_descripcion = ""
'  txtjustifica = ""
'  txt_venta = ""
'  txtterref = ""
End Sub
'Private Sub fbuscaunidad()
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  If rstFc_unidad_ejecutora.RecordCount > 0 Then
'    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
'  Else
'    LblUni_descripcion_larga.Caption = ""
'  End If
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'End Sub

Sub creaVista()
'db.Execute "drop view vwF04"
'
'db.Execute "create view vwF04 as " & _
'            "select  ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.tipoben_codigo, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, ao_solicitud_lista.telefono, ao_solicitud_lista.razon_s, ao_solicitud.codigo_solicitud, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_numero, ao_solicitud.por_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.caracteristicas, ao_solicitud.duracion_estimada_tiempo, " & _
'            "ao_solicitud.tr_adjuntos AS docAdjunta, " & _
'            "ao_solicitud.codigo_bien, ac_bienes.bie_descripcion , ao_solicitud.observaciones, fc_unidad_ejecutora.uni_descripcion_larga, ao_solicitud.fecha_solicitud, " & _
'            "(rc_personal.paterno) + ' ' + (rc_personal.materno) + ' ' +(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
'            "from ao_solicitud_lista  ,     " & _
'                 "ao_solicitud       ,     " & _
'                 "fc_unidad_ejecutora,     " & _
'                 "rc_personal,             " & _
'                 "ac_bienes                " & _
'            "where  ao_solicitud_lista.ges_Gestion       = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
'                    "ao_solicitud_lista.codigo_unidad    = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
'                    "ao_solicitud_lista.codigo_solicitud =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
'                    "ao_solicitud_lista.ges_Gestion      = ao_solicitud.ges_gestion            and " & _
'                    "ao_solicitud_lista.codigo_unidad    = ao_solicitud.codigo_unidad          and " & _
'                    "ao_solicitud_lista.codigo_solicitud = ao_solicitud.codigo_solicitud       and " & _
'                    "ao_solicitud.codigo_unidad          = fc_unidad_ejecutora.codigo_unidad   and " & _
'                    "ao_solicitud.codigo_bien            = ac_bienes.codigo_bien               and " & _
'                    "ao_solicitud.ci                     = rc_personal.ci                      " & _
'            "GROUP BY ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.tipoben_codigo, " & _
'            "ao_solicitud.codigo_solicitud, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.razon_s, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, " & _
'            "ao_solicitud_lista.telefono, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.nacional_extranjero, ao_solicitud.por_tiempo, ao_solicitud.codigo_bien, ac_bienes.bie_descripcion, ao_solicitud.duracion_estimada_numero, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.esparaRH, ao_solicitud.tr_adjuntos, ao_solicitud.observaciones, ao_solicitud.caracteristicas, fc_unidad_ejecutora.Uni_descripcion_larga, ao_solicitud.fecha_solicitud, (rc_personal.paterno)+' '+(rc_personal.materno)+' '+(rc_personal.nombres)+' ['+ao_solicitud.ci+']', ao_solicitud_lista.id_beneficiario "
'
''            "trim$(rc_personal.paterno) + ' ' + trim$(rc_personal.materno) + ' ' +trim$(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
'
''''db.Execute "create view vwF05 as " & _
''''            "select  ao_solicitud_lista.* " & _
''''            "from ao_solicitud_lista"
End Sub

Sub CREAVISTAF11()
'db.Execute "drop view VWF11"
'db.Execute "create view VWF11 as " & _
'    "SELECT ao_Solicitud.Ges_Gestion, ao_Solicitud.codigo_unidad, " & _
'    "ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, " & _
'    "ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, " & _
'    "ao_Solicitud.fecha_solicitud, ao_Solicitud.codigo_bien, " & _
'    "ac_bienes_grupo.DescGrupo, RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres, " & _
'    "ao_Solicitud.observaciones, ao_Solicitud.caracteristicas, " & _
'    "ao_Solicitud.tr_adjuntos, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, " & _
'    "ao_Solicitud.duracion_estimada_numero, ao_Solicitud.duracion_estimada_tiempo, " & _
'    "ao_solicitud_lista.codDetalle AS ci_material,  ao_solicitud_lista.profesion, ao_solicitud_lista.Aplanilla, " & _
'    "ao_solicitud_lista.razon_s, ao_solicitud_lista.Nro_pagos, ao_solicitud_lista.Monto_solicitud_dl, ao_solicitud_lista.AUnidad " & _
'"FROM ao_Solicitud, ao_Solicitud_detalle, ac_bienes_grupo, RC_Personal, ao_solicitud_lista " & _
'"WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
'    "(ao_Solicitud.codigo_unidad) = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
'    "(ao_Solicitud.codigo_solicitud) =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
'    "ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion AND " & _
'    "ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad AND " & _
'    "ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND " & _
'    "ao_Solicitud.codigo_unidad = ao_Solicitud_lista.codigo_unidad AND " & _
'    "ao_Solicitud.codigo_solicitud = ao_Solicitud_lista.codigo_solicitud AND " & _
'    "ao_Solicitud.CodGrupo = ac_bienes_grupo.CodGrupo AND " & _
'    "ao_Solicitud.ci = RC_Personal.ci"
End Sub

Private Sub acumulaMont(ges, Nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  Set rs_datos19 = New ADODB.Recordset
  If rs_datos19.State = 1 Then rs_datos19.Close
'  LblGestion
'  lblcorrelVenta
'  lblNroVenta
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!CANTOT
  End If
  
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If
  
  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  Dim sqlAux As String
  sqlAux = "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & Nro & " "
  db.Execute sqlAux  '"update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & Nro & " "
  
'  TxtMontoBs.Text = VAR_AUX
'  TxtCobrado.Text = Cobrobs
'  TxtBstotal.Text = VAR_Bs
  
  If rstacumdet.State = 1 Then rstacumdet.Close
  
End Sub

'Private Sub sstab1_Click(PreviousTab As Integer)
'    Select Case SSTab1.Tab
'        Case 0
'            lbl_titulo1.Caption = SSTab1.Caption
''            lbl_titulo3.Caption = SSTab1.Caption
'            FraNavega1.Caption = SSTab1.Caption
'            FraGrabarCancelar1.Visible = False
'            OptFilGral01.Value = True
'            Call OptFilGral01_Click
'        Case 1
'            If glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Then
'                lbl_titulo = SSTab1.Caption
''                lbl_titulo2 = SSTab1.Caption
'                FraNavega.Caption = SSTab1.Caption
'                FraGrabarCancelar.Visible = False
'                Call ABRIR_TABLAS_AUX
'                OptFilGral1.Value = True
'                Call OptFilGral1_Click
'                'FACTURA O RECIBO
'            Else
'                SSTab1.Tab = 0
'            End If
'            Picture1.Visible = True
'        Case 2
'            lbl_titulo2 = SSTab1.Caption
''            lbl_titulo5 = SSTab1.Caption
'            FraNavega2.Caption = SSTab1.Caption
'            FraGrabarCancelar2.Visible = False
'            OptFilGral03.Value = True
'            Call OptFilGral03_Click
'    End Select
'End Sub


'adelante
Private Function CodigoControl(NAuto As String, NFactura As String, Nit As String, Fecha As String, Monto As String, Key As String) As String
Dim Suma As Currency
Dim CodControl As String, Cadena As String, NroVer As String
Dim Pos As Integer, i As Integer, Nro As Integer, j As Integer
Dim SumTot As Long, SumPar(1 To 5) As Currency

  Suma = 0
  Cadena = NFactura
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  NFactura = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox NFactura
  'Para el Nit o CI del Cliente.
  Cadena = Nit
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Nit = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox Nit
  'Para la Fecha de transaccion.
  Cadena = Fecha
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Fecha = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox Fecha
  'Para el monto de transaccion.
  Cadena = Monto
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Monto = Cadena
  'MsgBox Monto
  Suma = Suma + CDbl(Cadena)
  'MsgBox Suma
  
  'Para Obtener los 5 numeros Verhoeff.
  Cadena = Str(Suma)
  For i = 1 To 5
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  NroVer = Right(Cadena, 5)
  'MsgBox NroVer
  
  'Para obtener las nuevas cadenas.
  Cadena = ""
  Pos = 1
  For i = 1 To 5
    Nro = (Val(Mid(NroVer, i, 1)) + 1)
    Select Case i
      Case 1: Cadena = Cadena & NAuto & Mid(Key, Pos, Nro)
      Case 2: Cadena = Cadena & NFactura & Mid(Key, Pos, Nro)
      Case 3: Cadena = Cadena & Nit & Mid(Key, Pos, Nro)
      Case 4: Cadena = Cadena & Fecha & Mid(Key, Pos, Nro)
      Case 5: Cadena = Cadena & Monto & Mid(Key, Pos, Nro)
    End Select
    Pos = Pos + Nro
  Next i

  Cadena = AllegedRC4(Cadena, (Key & NroVer))

  
  SumTot = 0
  i = 0
  Do While i < Len(Cadena)
    i = i + 1
    SumTot = SumTot + Asc(Mid(Trim(Cadena), i, 1))
  Loop
 
  
  For i = 1 To 5
    SumPar(i) = 0
    j = i
    Do While j <= Len(Cadena)
      SumPar(i) = SumPar(i) + Asc(Mid(Cadena, j, 1))
      j = j + 5
    Loop
  
  Next i
  
  Suma = 0
  For i = 1 To 5
    SumPar(i) = Int((SumTot * SumPar(i)) / (Val(Mid(NroVer, i, 1)) + 1))
    Suma = Suma + SumPar(i)
  Next i
  Cadena = Base64(Str(Suma))
  
  Cadena = AllegedRC4(Cadena, (Key & NroVer))
  

  CodigoControl = ""
  i = 0
  j = 1
  
  Do While i < Len(Cadena)
    i = i + 1
    If i Mod 2 = 0 Then
      CodigoControl = CodigoControl & Mid(Cadena, j, 2) & "-"
      j = i + 1
    End If
  Loop
  
  CodigoControl = Mid(CodigoControl, 1, (Len(CodigoControl) - 1))
End Function

'Public Function Redondear(dNumero As Double, iDecimales As Integer) As Double
'    Dim lMultiplicador As Long
'    Dim dRetorno As Double
'
'    If iDecimales > 9 Then iDecimales = 9
'    lMultiplicador = 10 ^ iDecimales
'    dRetorno = CDbl(CLng(dNumero * lMultiplicador)) / lMultiplicador
'
'    Redondear = dRetorno
'End Function
'Private Function Redondeo(ByVal Numero, ByVal Decimales)
'      Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
'End Function
'
'Private Sub TxtMonto02_LostFocus()
'    TxtMonto02D.Text = Round(CDbl(TxtMonto02.Text) / Ado_datos02.Recordset!cobranza_tdc, 2)
'End Sub
'
'Private Sub TxtMonto02D_LostFocus()
'    TxtMonto02.Text = Round(CDbl(TxtMonto02D.Text) * Ado_datos02.Recordset!cobranza_tdc, 2)
'End Sub
'
'Private Sub TxtMontoDol_Change()
'    'TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(txt_tdc.Text)
'End Sub
'
'Private Sub TxtMontoDol_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
'End Sub
'
'Private Sub TxtMontoDol_LostFocus()
'    TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(txt_tdc.Text)
'End Sub

Private Sub OptFilGral06_Click()
    '===== Proceso para filtrado de TARIJA
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '6'  AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset

End Sub

Private Sub OptFilGral07_Click()
    '===== Proceso para filtrado de SANTA CRUZ
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '7'  AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset

End Sub

Private Sub OptFilGral08_Click()
    '===== Proceso para filtrado de BENI
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '8'  AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset

End Sub

Private Sub OptFilGral09_Click()
    '===== Proceso para filtrado de PANDO
    BtnBuscar.Visible = True
    BtnImprimir.Visible = True
'    BtnImprimir1.Visible = True
'    BtnImprimir3.Visible = True
    
    Set rs_datos16 = New ADODB.Recordset
    If rs_datos16.State = 1 Then rs_datos16.Close
    queryinicial = " SELECT * FROM av_ventas_cabecera WHERE  (venta_tipo <> 'A' and left(edif_codigo,1) = '9'  AND estado_codigo = 'APR') "
    rs_datos16.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos16.Sort = "depto_codigo_vta, unidad_codigo, edif_codigo, venta_fecha"
    Set Ado_datos16.Recordset = rs_datos16.DataSource
    Set dg_datos16.DataSource = Ado_datos16.Recordset

End Sub

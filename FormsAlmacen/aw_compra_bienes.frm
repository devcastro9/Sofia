VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_compra_bienes 
   BackColor       =   &H00000000&
   Caption         =   "Procesos Administrativos - COMEX - Compra Bienes"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   12720
      Picture         =   "aw_compra_bienes.frx":0000
      ScaleHeight     =   1425
      ScaleWidth      =   2715
      TabIndex        =   84
      Top             =   8160
      Width           =   2775
      Begin VB.CommandButton BtnAprobar2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   60
         Picture         =   "aw_compra_bienes.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Anula Detalle de Registro"
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1285
         Picture         =   "aw_compra_bienes.frx":6C77E
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Imprime Detalle de Registro"
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1285
         Picture         =   "aw_compra_bienes.frx":6D04B
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Modifica Detalle de Registro"
         Top             =   180
         Width           =   1370
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   60
         Picture         =   "aw_compra_bienes.frx":6D960
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Adiciona Detalle de Registro"
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.PictureBox FrmABMDet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   10200
      ScaleHeight     =   3915
      ScaleWidth      =   1125
      TabIndex        =   74
      Top             =   840
      Width           =   1160
      Begin VB.CommandButton BtnEliminar4 
         BackColor       =   &H80000016&
         Caption         =   "Retor.Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   0
         Picture         =   "aw_compra_bienes.frx":6E11F
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Retorna todos los registros del Destino al Origen "
         Top             =   3120
         Width           =   1125
      End
      Begin VB.CommandButton BtnAprobar4 
         BackColor       =   &H80000016&
         Caption         =   "Envía Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         MaskColor       =   &H00FFFF80&
         Picture         =   "aw_compra_bienes.frx":6E561
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Envia todos los registros del Origen al Destino"
         Top             =   2400
         Width           =   1125
      End
      Begin VB.CommandButton BtnEliminar3 
         Caption         =   "Retornar 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   0
         Picture         =   "aw_compra_bienes.frx":6E9A3
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Retorna un equipo al ORIGEN"
         Top             =   840
         Width           =   1125
      End
      Begin VB.CommandButton BtnAprobar3 
         Appearance      =   0  'Flat
         Caption         =   "Envia Uno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         MaskColor       =   &H80000014&
         Picture         =   "aw_compra_bienes.frx":6EDE5
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Envia un solo item al DESTINO"
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   120
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   63
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "aw_compra_bienes.frx":6F227
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   72
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
         Picture         =   "aw_compra_bienes.frx":6F9E9
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   71
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "aw_compra_bienes.frx":702B6
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   70
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
         Picture         =   "aw_compra_bienes.frx":70A6B
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   69
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
         Picture         =   "aw_compra_bienes.frx":7129E
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   68
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
         Picture         =   "aw_compra_bienes.frx":719EA
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   67
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
         Picture         =   "aw_compra_bienes.frx":722FF
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   66
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "aw_compra_bienes.frx":72ABE
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "aw_compra_bienes.frx":72CC8
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   73
         Top             =   195
         Width           =   1815
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
      TabIndex        =   59
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
         Picture         =   "aw_compra_bienes.frx":7310A
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   61
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
         Picture         =   "aw_compra_bienes.frx":739F6
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   60
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl_titulo2 
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
         Left            =   13215
         TabIndex        =   62
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega1 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFFF&
      Height          =   5280
      Left            =   11280
      TabIndex        =   55
      Top             =   720
      Width           =   8055
      Begin MSDataGridLib.DataGrid dg_datos1 
         Height          =   4530
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7990
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         ForeColor       =   0
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "compra_fecha"
            Caption         =   "Fecha.Reg."
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
            DataField       =   "adjudica_codigo"
            Caption         =   "Compra Nro"
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
         BeginProperty Column03 
            DataField       =   "adjudica_descripcion"
            Caption         =   "Descripcion Bien o Servicio"
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
            DataField       =   "adjudica_cantidad_total"
            Caption         =   "Cantidad"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral4 
         Caption         =   "Todos"
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
         Left            =   4680
         TabIndex        =   58
         Top             =   4920
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral3 
         Caption         =   "Pendientes"
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
         Left            =   1440
         TabIndex        =   57
         Top             =   4920
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos13 
         Height          =   330
         Left            =   120
         Top             =   4800
         Width           =   7785
         _ExtentX        =   13732
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
         BackColor       =   -2147483633
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
   Begin VB.Frame FraDet2 
      BackColor       =   &H00000000&
      Caption         =   "CRONOGRAMA - ORDEN DE PAGO"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   240
      TabIndex        =   12
      Top             =   7920
      Width           =   15375
      Begin MSDataGridLib.DataGrid dg_det3 
         Bindings        =   "aw_compra_bienes.frx":741CC
         Height          =   1320
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   2328
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "pago_codigo"
            Caption         =   "Nro.O.P."
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
            DataField       =   "pago_descripcion"
            Caption         =   "Concepto.de.la. Orden de Pago"
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
            DataField       =   "pago_fecha_prog"
            Caption         =   "Fecha.O.P."
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
            DataField       =   "pago_total_bs"
            Caption         =   "Monto.a.Pagar.Bs"
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
            DataField       =   "pago_total_dol"
            Caption         =   "Monto.a.Pagar.USD"
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
         BeginProperty Column06 
            DataField       =   "adjudica_codigo"
            Caption         =   "Nro.Adjudica"
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
            DataField       =   "pago_emite_factura"
            Caption         =   "Emite.Factura"
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
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   5564.977
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "ADJUDICACION"
      ForeColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   17895
      Begin VB.PictureBox FrmABMDet2 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   1490
         Left            =   14880
         Picture         =   "aw_compra_bienes.frx":741E7
         ScaleHeight     =   1425
         ScaleWidth      =   2715
         TabIndex        =   79
         Top             =   120
         Width           =   2775
         Begin VB.CommandButton BtnAddDetalle2 
            BackColor       =   &H00FFC0C0&
            Height          =   525
            Left            =   60
            Picture         =   "aw_compra_bienes.frx":E0219
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Adiciona Detalle de Registro"
            Top             =   180
            Width           =   1245
         End
         Begin VB.CommandButton BtnModDetalle2 
            BackColor       =   &H00FFC0C0&
            Height          =   525
            Left            =   1285
            Picture         =   "aw_compra_bienes.frx":E09D8
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Modifica Detalle de Registro"
            Top             =   180
            Width           =   1370
         End
         Begin VB.CommandButton Command22 
            BackColor       =   &H00FFC0C0&
            Height          =   525
            Left            =   60
            Picture         =   "aw_compra_bienes.frx":E12ED
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Anula Detalle de Registro"
            Top             =   720
            Width           =   1245
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
            Height          =   525
            Left            =   1285
            Picture         =   "aw_compra_bienes.frx":E1A39
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Imprime Detalle de Registro"
            Top             =   720
            Width           =   1365
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "aw_compra_bienes.frx":E2306
         Height          =   1440
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   2540
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
         Caption         =   " "
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cod.Proveedor"
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
            DataField       =   "concepto_proveedor"
            Caption         =   "Denominación.Proveedor"
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
            DataField       =   "adjudica_monto_dol"
            Caption         =   "Precio.Total.USD"
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
         BeginProperty Column03 
            DataField       =   "adjudica_monto_bs"
            Caption         =   "Precio.Total_BOB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "adjudica_cantidad_total"
            Caption         =   "Cantidad"
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
            DataField       =   "fecha_inicio_contrato"
            Caption         =   "Fecha.Inicio"
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
            DataField       =   "fecha_fin_contrato"
            Caption         =   "Fecha.Finalizacion"
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
            DataField       =   "fecha_recibe_almacen"
            Caption         =   "Fecha.Ingreso Almacen"
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
            DataField       =   "nro_nota_remision"
            Caption         =   "Nro Proforma"
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
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3885.166
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFFF&
      Height          =   5280
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   9975
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   4530
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7990
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         ForeColor       =   0
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "Trámite"
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
            Caption         =   "U.Ejecutora"
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
            DataField       =   "edif_codigo"
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
         BeginProperty Column03 
            DataField       =   "compra_fecha"
            Caption         =   "Fecha.Reg."
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
            Caption         =   "Nombre Edificio"
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
         BeginProperty Column06 
            DataField       =   "compra_descripcion"
            Caption         =   "Descripcion Bien o Servicio"
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
            DataField       =   "compra_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "estado_codigo_eqp"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         Caption         =   "Pendientes"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   4920
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptFilGral2 
         Caption         =   "Todos"
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
         Left            =   7080
         TabIndex        =   11
         Top             =   4905
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   240
         Top             =   4800
         Width           =   9615
         _ExtentX        =   16960
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
         BackColor       =   -2147483633
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
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11280
      TabIndex        =   0
      Top             =   10260
      Width           =   11280
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9960
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
   Begin Crystal.CrystalReport CR01 
      Left            =   7200
      Top             =   11040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
      Top             =   9960
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
      Left            =   4440
      Top             =   9960
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
      Left            =   6720
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   9000
      Top             =   9960
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
      Left            =   11280
      Top             =   9960
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
      Left            =   13560
      Top             =   9960
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
      Left            =   120
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   2280
      Top             =   10320
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
      Left            =   4440
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   120
      Top             =   10680
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
      Caption         =   "Ado_detalle1"
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2400
      Top             =   10680
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
      Caption         =   "Ado_detalle2"
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
      Left            =   6720
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Ado_datos12 
      Height          =   330
      Left            =   9000
      Top             =   10320
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
      Caption         =   "Ado_datos12"
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
   Begin MSAdodcLib.Adodc Ado_detalle3 
      Height          =   330
      Left            =   4800
      Top             =   10680
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
      Caption         =   "Ado_detalle3"
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
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   4080
      Left            =   5760
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "compra_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   555
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2400
         Width           =   8025
      End
      Begin VB.TextBox g 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   8955
         TabIndex        =   18
         Top             =   1175
         Width           =   270
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7680
         TabIndex        =   17
         Top             =   525
         Width           =   290
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4515
         TabIndex        =   16
         Top             =   1170
         Width           =   270
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   15
         Top             =   2000
         Visible         =   0   'False
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3360
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         DataField       =   "unidad_codigo_adm"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_sigla"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         DataField       =   "venta_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "venta_tipo"
         BoundColumn     =   "venta_tipo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "compra_fecha"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   7785
         TabIndex        =   23
         Top             =   1980
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   98893825
         CurrentDate     =   41678
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3180
         TabIndex        =   24
         Top             =   3135
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo5"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   1155
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4785
         TabIndex        =   30
         Top             =   1155
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         DataField       =   "unidad_codigo_adm"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6480
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         DataField       =   "venta_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4395
         TabIndex        =   32
         Top             =   1980
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "venta_tipo_descripcion"
         BoundColumn     =   "venta_tipo"
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
      Begin MSDataListLib.DataCombo d 
         DataField       =   "unidad_codigo_adm"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3255
         TabIndex        =   33
         Top             =   510
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   34
         Top             =   3140
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   35
         Top             =   1980
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "ESTADO "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   2
         Left            =   7200
         TabIndex        =   36
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cod.Trámite"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   54
         Top             =   225
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Registro"
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
         Height          =   195
         Index           =   12
         Left            =   7785
         TabIndex        =   53
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Doc. Respaldo"
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
         Height          =   195
         Index           =   13
         Left            =   3240
         TabIndex        =   52
         Top             =   3675
         Width           =   1695
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "doc_numero"
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
         Height          =   300
         Left            =   5100
         TabIndex        =   51
         Top             =   3645
         Width           =   1365
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
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
         Height          =   300
         Left            =   180
         TabIndex        =   50
         Top             =   510
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9360
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9360
         Y1              =   3555
         Y2              =   3555
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad Ejecutora"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3285
         TabIndex        =   49
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Representante Legal / Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4740
         TabIndex        =   48
         Top             =   885
         Width           =   2130
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Responsable del Proceso:"
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
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   1710
         Width           =   2235
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código Registro"
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
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   3675
         Width           =   1350
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Actividad del POA"
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
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   3165
         Width           =   1560
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Concepto:"
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
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   2490
         Width           =   885
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Edificio"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   885
         Width           =   510
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "unidad_codigo_ant"
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
         Height          =   300
         Left            =   1595
         TabIndex        =   42
         Top             =   510
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite del Trámite"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   41
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Compra"
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
         Height          =   195
         Index           =   1
         Left            =   4395
         TabIndex        =   40
         Top             =   1710
         Width           =   1350
      End
      Begin VB.Label dtc_codigo9 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "doc_codigo"
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
         Height          =   300
         Left            =   1680
         TabIndex        =   39
         Top             =   3645
         Width           =   1245
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   8040
         TabIndex        =   38
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "compra_codigo"
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
         Height          =   300
         Left            =   8040
         TabIndex        =   37
         Top             =   510
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Ado_detalle4 
      Height          =   330
      Left            =   7080
      Top             =   10680
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
      Caption         =   "Ado_detalle4"
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
Attribute VB_Name = "aw_compra_bienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
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
Dim rs_datos11 As New ADODB.Recordset
Dim query As String
Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset
Dim rs_det3 As New ADODB.Recordset
Dim rs_det4 As New ADODB.Recordset
Dim rs_det5 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset
Dim rs_aux11 As New ADODB.Recordset


Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim VAR_PAIS As String
Dim VAR_GES As String
Dim Var_Comp As Integer
Dim VAR_BIEN As String
Dim VAR_AP As String

Dim VAR_CMPBTE As Integer
Dim VAR_ADJUDICA As Integer

Dim VAR_AUX, VAR_CONT2 As Double
Dim VAR_FOBSEG, VAR_FOBSEG2 As Double
Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAddDetalle_Click()
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Visible = False
    FrmABMDet2.Visible = False
    FraDet3.Visible = False
'    FrmABMDet3.Visible = False
    Fra_datos.Enabled = False
'    Select Case dtc_codigo2.Text
'        Case "1"    'SOLO COMPRAS BB y SS
'        Case "2"    'SOLO VENTA DE BIENES
'        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'        Case "L"    'IMPORTACION DIRECTA CLIENTE - COMEX
'        marca1 = Ado_datos.Recordset.Bookmark
'            Call ABRIR_TABLA_DET
            Ado_detalle3.Recordset.AddNew
            
    aw_bienes_adjudica.txtCodigo1 = Me.Ado_datos.Recordset!solicitud_codigo  'cod_cabecera
    aw_bienes_adjudica.txt_campo1.Text = Me.Ado_datos.Recordset!unidad_codigo  'Unidad
'    aw_bienes_adjudica.Txt_descripcion = Me.dtc_desc1.Text
    aw_bienes_adjudica.txtCodigo1.Caption = Me.Ado_datos.Recordset!compra_codigo
    aw_bienes_adjudica.lbl_adjudica.Caption = Me.Ado_detalle3.Recordset!adjudica_codigo
'            frm_ao_comex_pagos.txtSW.Text = Me.Ado_datos.Recordset!venta_tipo
'            frm_ao_comex_pagos.txt_pais.Text = VAR_PAIS
'            frm_ao_comex_pagos.Txtestado.Text = "REG"

    frm_ao_comex_pagos.Show vbModal
'        Case "V"    'FACTURACION LOCAL - COMEX
'    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Visible = True
    FrmABMDet2.Visible = True
    FraDet3.Visible = True
'    FrmABMDet3.Visible = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAddDetalle1_Click()
  marca1 = Ado_detalle4.Recordset.Bookmark
  If rs_det4!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
'    FraDet3.Enabled = False
'    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    Call ABRIR_TABLA_DET
    Select Case Glaux
        Case "PROVI"    'PROVISION DE EQUIPOS
            'NO HAY
        Case "TRANS"    'TRANSPORTE
            Ado_detalle2.Recordset.AddNew
            aw_bienes_adjudica.txtCodigo1.Caption = Me.txt_codigo.Caption
            aw_bienes_adjudica.txt_campo1 = Me.dtc_codigo1.Text
'            aw_bienes_adjudica.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            aw_bienes_adjudica.lbl_edif.Caption = dtc_codigo3.Text
'            aw_bienes_adjudica.lbl_det.Caption = Glaux
'            aw_bienes_adjudica.Txt_estado.Caption = "REG"
            aw_bienes_adjudica.Show vbModal
            
        Case "UALMI"    'INSUMOS
            Ado_detalle4.Recordset.AddNew
'            aw_bienes_adjudica.ges_gestion.Caption = Me.txt_codigo.Caption
'
'            aw_bienes_adjudica.adjudica_codigo.Caption = Me.dtc_codigo1.Text
'            aw_bienes_adjudica.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            aw_bienes_adjudica.lbl_edif.Caption = dtc_codigo3.Text
'            aw_bienes_adjudica.lbl_det.Caption = Glaux
'            aw_bienes_adjudica.Txt_estado.Caption = "REG"
            aw_bienes_adjudica.Show vbModal
        Case "ADUAN"    'DESADUANIZACION
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes2.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes2.txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes2.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes2.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes2.lbl_det.Caption = Glaux
            frm_solicitud_bienes2.Txt_estado.Caption = "REG"
            frm_solicitud_bienes2.Show vbModal
        Case "DESCA"    'DESCARGO Y OTROS
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes2.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes2.txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes2.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes2.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes2.lbl_det.Caption = Glaux
            frm_solicitud_bienes2.Txt_estado.Caption = "REG"
            frm_solicitud_bienes2.Show vbModal
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
'    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnAddDetalle2_Click()
  If rs_datos!estado_codigo = "REG" Then
  
    'VAR_PAIS = "BRA"
'    If VAR_PAIS = "NN" Then
'        MsgBox "ERROR, No ha sido registrada la industria. consulte con Gerencia Comercial y vuelva a intentar !! ", vbExclamation
'    Else
''        'FOB + SEG de la Cotizacion
'        Set rs_datos5 = New ADODB.Recordset
'        If rs_datos5.State = 1 Then rs_datos5.Close
'        rs_datos5.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and pais_codigo = '" & VAR_PAIS & "' ", db, adOpenStatic
'        'Set Ado_datos5.Recordset = rs_datos5
'        If rs_datos5.RecordCount > 0 Then
'            If IsNull(rs_datos5!cotiza_fob_seg_dol) Then
'                MsgBox "ERROR, No ha sido registrado el precio FOB. Consulte con Gerencia Comercial para corregirlo. !! ", vbExclamation
'                'Exit Sub
'            Else
'                VAR_FOBSEG = rs_datos5!cotiza_fob_seg_dol
'                VAR_FOBSEG2 = IIf(IsNull(rs_datos5!cotiza_fob_seg_bs), VAR_FOBSEG * GlTipoCambioOficial, rs_datos5!cotiza_fob_seg_bs)
'            End If
'        Else
'            MsgBox "ERROR, No ha sido identificado el registro. Consulte con Gerencia Comercial y vuelva a intentar !! ", vbExclamation
'            Exit Sub
'        End If
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
        swnuevo = 1
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FraDet2.Visible = False
        FrmABMDet2.Visible = False
'        FraDet3.Visible = False
''        FrmABMDet3.Visible = False
        Fra_datos.Enabled = False
        
                Call proveedor
                Ado_detalle4.Recordset.AddNew
'                aw_bienes_adjudica.txt_codigo.Caption = Me.Ado_datos13.Recordset!solicitud_codigo  'cod_cabecera
                aw_bienes_adjudica.txt_campo1.Text = Me.Ado_datos13.Recordset!unidad_codigo  'Unidad
'                aw_bienes_adjudica.Txt_descripcion.Caption = Me.dtc_desc1.Text
                aw_bienes_adjudica.txtCodigo1.Caption = Me.Ado_datos13.Recordset!ges_gestion
                aw_bienes_adjudica.lbl_adjudica.Caption = Me.Ado_detalle4.Recordset!adjudica_codigo
                aw_bienes_adjudica.txtSW.Text = Me.Ado_datos13.Recordset!venta_tipo
                aw_bienes_adjudica.txt_total_dol = VAR_FOBSEG
                aw_bienes_adjudica.txt_total_bs = VAR_FOBSEG2
                aw_bienes_adjudica.txt_pais.Text = VAR_PAIS
'                aw_bienes_adjudica.txtEstado.Text = "REG"
                aw_bienes_adjudica.Show vbModal
    '        Case "V"    'FACTURACION LOCAL - COMEX
    '    End Select
        swnuevo = 0
        fraOpciones.Enabled = True
        FraNavega.Enabled = True
        FraDet2.Visible = True
        FrmABMDet2.Visible = True
        FraDet3.Visible = True
'        FrmABMDet3.Visible = True
    '    Fra_datos.Enabled = True
'    End If
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
  If Ado_detalle1.Recordset.RecordCount > 0 Then
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado ...", vbExclamation, "Validación de Registro"
   End If
 Else
     MsgBox "No se puede ANULAR, el registro no fue identificado correctamente ...", vbExclamation, "Validación de Registro"
 End If
End Sub

Private Sub BtnAnlDetalle2_Click()
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
End Sub



Private Sub BtnAprobar_Click()
      On Error GoTo UpdateErr
  If Ado_datos13.Recordset.RecordCount > 0 Then
   VAR_AP = "APR"
   
   db.Execute " INSERT INTO ao_compra_apertura_sobres( ges_gestion, adjudica_codigo,  apertura_fecha, apertura_sobre_a, apertura_sobre_b, apertura_sobre_c, apertura_sobre_todos, " & _
                    "  beneficiario_codigo, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, apertura_califica_sobre_a, " & _
                    "  apertura_califica_sobre_b, apertura_califica_sobre_c, apertura_califica_total, nro_nota_remision, apertura_fecha_estimada_adjudica, " & _
                    " apertura_tipo_recomendacion, poa_codigo, estado_codigo, usr_codigo, fecha_registro, hora_registro, adjudica_cantidad_total, adjudica_monto_bs, " & _
                    "  adjudica_monto_dol , fecha_inicio_contrato, fecha_fin_contrato, fecha_recibe_almacen, concepto_proveedor  )" & _
                    " (SELECT ges_gestion, adjudica_codigo, '" & Date & "', 'N', 'N', 'N', 'N', " & _
                    "  beneficiario_codigo, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, '0', " & _
                    "  '0', '0', '0', '', '" & Date & "', " & _
                    " '0', poa_codigo, estado_codigo, usr_codigo, fecha_registro, hora_registro, adjudica_cantidad_total, adjudica_monto_bs, " & _
                    "  adjudica_monto_dol , fecha_inicio_contrato, fecha_fin_contrato, fecha_recibe_almacen, '' " & _
                    " FROM av_compra_adjudica_bienes " & _
                    " WHERE ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND adjudica_codigo = '" & Ado_datos13.Recordset!adjudica_codigo & "') "


   db.Execute "UPDATE ao_compra_adjudica SET estado_codigo = 'APR' WHERE ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND adjudica_codigo = '" & Ado_datos13.Recordset!adjudica_codigo & "' "
     db.Execute "UPDATE ao_compra_adjudica_bienes SET estado_codigo = 'APR' WHERE ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND adjudica_codigo = '" & Ado_datos13.Recordset!adjudica_codigo & "' AND bien_codigo = '" & Ado_datos13.Recordset!bien_codigo & "' "
   
    
    db.Execute "UPDATE ao_compra_apertura_sobres SET ao_compra_apertura_sobres.adjudica_cantidad_total = (SELECT SUM(adjudica_cantidad)AS TOTAL FROM ao_compra_adjudica_bienes  WHERE adjudica_codigo = " & Ado_datos13.Recordset!adjudica_codigo & " ) WHERE ao_compra_apertura_sobres.adjudica_codigo = " & Ado_datos13.Recordset!adjudica_codigo & " "
'   Ado_datos13.Recordset.Requery
  Else
  
      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
 
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar2_Click()
    'PAGOS
    'WWWWWWWWWWWWWW
    If Ado_detalle3.Recordset!estado_codigo = "REG" Then
        VAR_COD4 = parametro    'UNIDAD
        VAR_SOL = Ado_datos.Recordset!solicitud_codigo  '
        'tipo_formulario TERCER PARAMETRO
        'org_codigo CUARTO PARAMETRO
        ' ini generación de correlativo
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from fc_organismo_financiamiento where org_codigo = '111'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
           rs_aux2!correlativo_gasto = rs_aux2!correlativo_gasto + 1
           VAR_CMPBTE = rs_aux2!correlativo_gasto
           rs_aux2.Update
        End If
        'WWWWWWWWWWWWWWW
        'correlv = Ado_datos.Recordset!venta_codigo
        'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
        Set rs_aux3 = New ADODB.Recordset
        If rs_aux3.State = 1 Then rs_aux3.Close
        rs_aux3.Open "select * from  fo_gastos_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
        If rs_aux3.RecordCount = 0 Then
            rs_aux3.AddNew
            rs_aux3!ges_gestion = glGestion     'Year(Date)
            rs_aux3!org_codigo = "111"
            rs_aux3!gasto_codigo = VAR_CMPBTE
            rs_aux3!tipo_comp = "DEV"
            rs_aux3!gasto_codigo_anterior = VAR_CMPBTE
            rs_aux3!unidad_codigo = VAR_COD4
            rs_aux3!solicitud_codigo = VAR_SOL
            rs_aux3!tipo_formulario = "F04"     'GC_TIPO_SOLICITUD
            'rs_aux3!pago_codigo = rs_aux3.RecordCount + 1
            
            rs_aux3!proceso_codigo = "FIN"
            rs_aux3!subproceso_codigo = "FIN-03"
            rs_aux3!etapa_codigo = "FIN-03-03"
            rs_aux3!clasif_codigo = "ADM"
            rs_aux3!doc_codigo = "R-111"
            rs_aux3!doc_numero = 0
            rs_aux3!poa_codigo = "4.2.3"
            
            rs_aux3!fecha_egreso = Date
            rs_aux3!tipo_moneda = "BOB"
            rs_aux3!da_codigo = "1.1"
            
            rs_aux3!fte_codigo = "10"   'DEVISAR DE LA TABLA fc_organismo_financiamiento
            rs_aux3!monto_Bolivianos = 0
            rs_aux3!monto_dolares = 0
            rs_aux3!liquido_pagar = 0
            rs_aux3!monto_Bolivianos_pag = 0
            rs_aux3!monto_dolares_pag = 0
            rs_aux3!Deducciones = 0
            rs_aux3!fecha_autorizacion = Date
            rs_aux3!justificacion = Txt_descripcion.Text
            rs_aux3!es_base = "S"
            
            rs_aux3!CODIGO_GRUPO = VAR_CMPBTE   'rs_aux3!pago_codigo
            rs_aux3!NUMERO_PAGO = 1
            rs_aux3!observaciones = txt_obs.Text
            
            'rs_aux3!edif_codigo = VAR_PROY2
            'rs_aux3!beneficiario_codigo = VAR_BENEF
            'rs_aux3!solicitud_tipo = "10"
            'rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
            
            rs_aux3!estado_devengado = "APR"
            rs_aux3!estado_pagado = "REG"
            rs_aux3!estado_contabilidad = "REG"
            
            rs_aux3!estado_codigo = "REG"
            rs_aux3!usr_codigo = glusuario
            rs_aux3!Fecha_Registro = Date
            rs_aux3!usr_codigo_aprueba = glusuario
            rs_aux3!fecha_aprueba = Date
            rs_aux3.Update
            
            'DETALLE Carga fo_gastos_detalle
            Set rstdestino = New ADODB.Recordset
            If rstdestino.State = 1 Then rstdestino.Close
            rstdestino.Open "select * from fo_gastos_detalle where org_codigo = '111' AND gasto_codigo= " & VAR_CMPBTE & "  ", db, adOpenKeyset, adLockBatchOptimistic
            If rstdestino.RecordCount > 0 Then
            End If
            Set rs_aux4 = New ADODB.Recordset
            If rs_aux4.State = 1 Then rs_aux4.Close
            rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux4.RecordCount > 0 Then
               VAR_REG = 1
               rs_aux4.MoveFirst
               While Not rs_aux4.EOF
               '     db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
               '     "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
               '     rs_aux4.MoveNext
                    db.Execute "INSERT INTO fo_gastos_detalle (ges_gestion, org_codigo, gasto_codigo, gasto_codigo_detalle, par_codigo, pro_codigo, codigo_beneficiario, concepto_pago, monto_total, monto_dolares_dev, tipo_cambio_dev, monto_Bolivianos, monto_Dolares, saldo_bolivianos, tipo_cambio, Porcentaje, deducciones, fecha_pago, depto_codigo, estado_aprobacion, fecha_autorizacion, Observacion, codigo_dev, usr_usuario, fecha_registro, hora_registro,  estado_conciliacion, codigo_poa " & _
                    "VALUES ('" & glGestion & "', '111', " & VAR_CMPBTE & ", '" & VAR_REG & "', " & rs_aux4!par_codigo & ", '8', '" & VAR_BENEF & "', '" & Txt_descripcion.Text & "', " & rs_aux4!bien_total_compra & ", " & rs_aux4!bien_total_compra * GlTipoCambioOficial & ", " & GlTipoCambioOficial & ", " & rs_aux4!bien_total_compra & ", " & rs_aux4!bien_total_compra * GlTipoCambioOficial & " , '0', " & GlTipoCambioOficial & ", '100', '0', '2', 'REG', '" & Date & "', '" & txt_obs.Text & "', " & VAR_CMPBTE & ", '" & glusuario & "', '" & Date & "', 'REG', '4.2.3')"
                   VAR_REG = VAR_REG + 1
               '     'cta_codigo, cheque_o_trf, numero_cheque_trf, cta_codigo_destino, cheque_o_trf_destino, numero_cheque_trf_destino,
               '     'Fecha_Aprobacion_tesoreria, fecha_impresion_cheque, banco_destino,
               Wend
            End If
            If rstdestino.State = 1 Then rstdestino.Close
        End If
        db.Execute "update ao_compra_planilla_pagos set estado_codigo = 'APR' where ges_gestion = " & Ado_detalle2.Recordset!ges_gestion & " and adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & " and pago_codigo=" & Ado_detalle3.Recordset!pago_codigo & "   "
        Call ABRIR_TABLA_DET
    Else
        MsgBox "NO se puede APROBAR un registro Anulado o previamente Aprobado. ", vbExclamation, "Atención!"
    End If
        'WWWWWWWWWW
End Sub



Private Sub BtnAprobar3_Click()
VAR_AP = "ENV"
Set rs_aux7 = New ADODB.Recordset
If rs_aux7.State = 1 Then rs_aux7.Close
rs_aux7.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
If rs_aux7.RecordCount > 0 Then


    Set rs_aux6 = New ADODB.Recordset
    SQL_FOR = "select * from fc_correl where tipo_tramite = 'Adjudica' "
    rs_aux6.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
    If rs_aux6.RecordCount > 0 Then
        rs_aux6!numero_correlativo = rs_aux6!numero_correlativo + 1
        VAR_ADJUDICA = rs_aux6!numero_correlativo
        rs_aux6.Update
    End If
    


'db.Execute " INSERT INTO ao_compra_adjudica(ges_gestion,compra_codigo, adjudica_codigo,unidad_codigo, solicitud_codigo ,adjudica_fecha ,proceso_codigo ,subproceso_codigo, etapa_codigo , " & _
'" clasif_codigo, doc_codigo,doc_numero, beneficiario_codigo,adjudica_descripcion, adjudica_cantidad_total,adjudica_monto_bs,tipo_moneda, adjudica_monto_dol, " & _
' " nro_nota_remision, fecha_inicio_contrato, fecha_fin_contrato, fecha_envio_proveedor,fecha_recibe_almacen,almacen_codigo,poa_codigo, mes_inicio_crono, " & _
'" cantidad_cuotas_pag,unimed_codigo_pag,correl_pagos_prog,estado_codigo,usr_codigo,fecha_registro,hora_registro,usr_codigo_aprueba, fecha_aprueba)" & _
'
'" SELECT (ges_gestion,compra_codigo," & VAR_ADJUDICA & "', unidad_codigo, solicitud_codigo,'" & Date & "',proceso_codigo, subproceso_codigo, etapa_codigo, " & _
'" clasif_codigo,doc_codigo,doc_numero, beneficiario_codigo, compra_descripcion,'0', '0',tipo_moneda, '0', " & _
'" nro_nota_remision,'" & Date & "', '" & Date & "', '" & Date & "','" & Date & "','0',poa_codigo, 'Enero', " & _
' " '0','Mes','0',estado_codigo,usr_codigo,fecha_registro,hora_registro, usr_codigo_aprueba,  fecha_aprueba) " & _
'
'                     " From av_compra_cabecera_detalle " & _
'" WHERE ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' and compra_codigo = '" & Ado_datos.Recordset!compra_codigo & "' and adjudica_codigo = '" & Ado_datos.Recordset!adjudica_codigo & "' "
'Ado_datos.Recordset.RecordCount


   sino = MsgBox("Está Seguro de enviar ya tiene una compra pendiente? ", vbYesNo + vbQuestion, "Atención")
 If sino = vbYes Then
db.Execute " INSERT INTO ao_compra_adjudica (ges_gestion, compra_codigo, adjudica_codigo, unidad_codigo, solicitud_codigo, adjudica_fecha, proceso_codigo, subproceso_codigo, etapa_codigo, " & _
                     " clasif_codigo, doc_codigo, doc_numero, beneficiario_codigo, adjudica_descripcion, adjudica_cantidad_total, adjudica_monto_bs, tipo_moneda, adjudica_monto_dol, " & _
                     " nro_nota_remision, fecha_inicio_contrato, fecha_fin_contrato, fecha_envio_proveedor, fecha_recibe_almacen, almacen_codigo, poa_codigo, mes_inicio_crono, " & _
                     " cantidad_cuotas_pag, unimed_codigo_pag, correl_pagos_prog, estado_codigo, usr_codigo, fecha_registro, hora_registro, usr_codigo_aprueba, fecha_aprueba) " & _
" (SELECT DISTINCT ges_gestion, compra_codigo, '" & VAR_ADJUDICA & "', unidad_codigo, solicitud_codigo, '" & Date & "', proceso_codigo, subproceso_codigo, etapa_codigo, " & _
                     " clasif_codigo, doc_codigo, doc_numero, beneficiario_codigo, compra_descripcion, compra_cantidad_total, '0', tipo_moneda, '0', " & _
                     " nro_nota_remision, '" & Date & "', '" & Date & "', '" & Date & "', '" & Date & "', '0', poa_codigo, 'ENERO', " & _
                     " '0', 'Mes', '0', estado_codigo, usr_codigo, fecha_registro, hora_registro, usr_codigo_aprueba, fecha_aprueba " & _
                     " From av_compra_cabecera_detalle " & _
" WHERE ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos.Recordset!compra_codigo & "') "


db.Execute " INSERT INTO ao_compra_adjudica_bienes (ges_gestion, compra_codigo, adjudica_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, adjudica_cantidad, " & _
                    "  bien_precio_adjudica_bs, bien_total_adjudica_bs, tipo_moneda, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo,  " & _
                    "  modelo_codigo, bien_nro_lote, bien_fecha_vencimiento, estado_codigo, usr_codigo, fecha_registro)  " & _
" (SELECT DISTINCT ges_gestion, compra_codigo, '" & VAR_ADJUDICA & "', bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, compra_cantidad,  " & _
                     " compra_precio_unitario_bs, compra_precio_total_bs, tipo_moneda, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo,  " & _
                     " modelo_codigo, '0',  '" & Date & "', 'REG', '" & glusuario & "',  '" & Date & "'  " & _
" From av_compra_cabecera_detalle  " & _
" WHERE ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos.Recordset!compra_codigo & "' AND bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "')  "
End If
    db.Execute "UPDATE ao_compra_cabecera SET estado_codigo_eqp = 'APR' WHERE ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos.Recordset!compra_codigo & "' "
    
    Call OptFilGral3_Click
Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
End If
'and ao_compra_adjudica.adjudica_codigo = '" & VAR_ADJUDICA & "' "
End Sub

Private Sub BtnAprobar4_Click()
VAR_AP = "ENV"

SQL_FOR = "select * from fc_correl where tipo_tramite = 'Adjudica' "
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
    If rs_aux2.RecordCount > 0 Then
        rs_aux2!numero_correlativo = rs_aux2!numero_correlativo + 1
        VAR_ADJUDICA = rs_aux2!numero_correlativo
        rs_aux2.Update
 End If
    



'query = "(" + queryinicial + ")"
'query = "(" + GlSqlAux + ")"
'db.Execute " INSERT INTO ao_compra_adjudica (ges_gestion, compra_codigo, adjudica_codigo, unidad_codigo, solicitud_codigo, adjudica_fecha, proceso_codigo, subproceso_codigo, etapa_codigo, " & _
'                     " clasif_codigo, doc_codigo, doc_numero, beneficiario_codigo, adjudica_descripcion, adjudica_cantidad_total, adjudica_monto_bs, tipo_moneda, adjudica_monto_dol, " & _
'                     " nro_nota_remision, fecha_inicio_contrato, fecha_fin_contrato, fecha_envio_proveedor, fecha_recibe_almacen, almacen_codigo, poa_codigo, mes_inicio_crono, " & _
'                     " cantidad_cuotas_pag, unimed_codigo_pag, correl_pagos_prog, estado_codigo, usr_codigo, fecha_registro, hora_registro, usr_codigo_aprueba, fecha_aprueba) " & _
'" '" & query & "' "
'
'
'db.Execute " INSERT INTO ao_compra_adjudica_bienes (ges_gestion, compra_codigo, adjudica_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, adjudica_cantidad, " & _
'                    "  bien_precio_adjudica_bs, bien_total_adjudica_bs, tipo_moneda, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo,  " & _
'                    "  modelo_codigo, bien_nro_lote, bien_fecha_vencimiento, estado_codigo, usr_codigo, fecha_registro)  " & _
'" '" & query & "' "
VAR_GES = 0
Var_Comp = 0
Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_compra_adjudica    ", db, adOpenKeyset, adLockOptimistic
        'If rs_aux2.RecordCount > 0 Then
             
             VAR_COD0 = 0
'             VAR_GES = ges_gestion
'             VAR_COMP = compra_codigo
             Set rs_aux3 = New ADODB.Recordset
             If rs_aux3.State = 1 Then rs_aux3.Close
             rs_aux3.Open GlSqlAux, db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux3.RecordCount > 0 Then
                 rs_aux3.MoveFirst
                 While Not rs_aux3.EOF
                    'If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then
                    If VAR_GES <> rs_aux3!ges_gestion And Var_Comp <> rs_aux3!compra_codigo Then
                    
                       rs_aux2.AddNew
                       rs_aux2!ges_gestion = rs_aux3!ges_gestion
                       rs_aux2!compra_codigo = rs_aux3!compra_codigo
                       VAR_GES = rs_aux2!ges_gestion
                       Var_Comp = rs_aux2!compra_codigo
                       rs_aux2!adjudica_codigo = VAR_ADJUDICA
                       rs_aux2!unidad_codigo = rs_aux3!unidad_codigo
                       rs_aux2!solicitud_codigo = rs_aux3!solicitud_codigo
                       rs_aux2!adjudica_fecha = Date
                       rs_aux2!proceso_codigo = rs_aux3!proceso_codigo
                       rs_aux2!subproceso_codigo = rs_aux3!subproceso_codigo
                       rs_aux2!etapa_codigo = rs_aux3!etapa_codigo
                       rs_aux2!clasif_codigo = rs_aux3!clasif_codigo
                       rs_aux2!doc_codigo = rs_aux3!doc_codigo
                       rs_aux2!doc_numero = rs_aux3!doc_numero
                       rs_aux2!beneficiario_codigo = rs_aux3!beneficiario_codigo
                       rs_aux2!adjudica_descripcion = rs_aux3!compra_descripcion
                       rs_aux2!adjudica_cantidad_total = rs_aux3!compra_cantidad_total
                       rs_aux2!adjudica_monto_bs = 0
                       rs_aux2!tipo_moneda = rs_aux3!tipo_moneda
                       rs_aux2!adjudica_monto_dol = 0
                       rs_aux2!nro_nota_remision = rs_aux3!nro_nota_remision
                       rs_aux2!fecha_inicio_contrato = Date
                       rs_aux2!fecha_fin_contrato = Date
                       rs_aux2!fecha_envio_proveedor = Date
                       rs_aux2!fecha_recibe_almacen = Date
                       rs_aux2!almacen_codigo = 0
                       rs_aux2!poa_codigo = rs_aux3!poa_codigo
                       rs_aux2!mes_inicio_crono = "ENERO"
                       rs_aux2!cantidad_cuotas_pag = 0
                       rs_aux2!unimed_codigo_pag = "MES"
                       rs_aux2!correl_pagos_prog = 0
                       rs_aux2!estado_codigo = "REG"
                       rs_aux2!usr_codigo = glusuario
                       rs_aux2!Fecha_Registro = Date
                       'rs_aux2!hora_registro = rs_aux3!hora_registro
                       rs_aux2!usr_codigo_aprueba = glusuario
                       rs_aux2!fecha_aprueba = Date
                       rs_aux2!edif_codigo = rs_aux3!edif_codigo
                       rs_aux2.Update
                        
                        VAR_COD0 = VAR_COD0 + 1
                       
                    End If
                    rs_aux3.MoveNext
                   
                 Wend
             End If
                        
'             --------ADJUDICA_BIENES-----------

SQL_FOR = "select * from fc_correl where tipo_tramite = 'Adjudica' "
    
   VAR_BIEN = 0

Set rs_aux4 = New ADODB.Recordset
        If rs_aux4.State = 1 Then rs_aux4.Close
        rs_aux4.Open "select * from ao_compra_adjudica_bienes    ", db, adOpenKeyset, adLockOptimistic
        'If rs_aux4.RecordCount > 0 Then
             
             VAR_COD0 = 0
'             VAR_GES = ges_gestion
'             VAR_COMP = compra_codigo
             Set rs_aux5 = New ADODB.Recordset
             If rs_aux5.State = 1 Then rs_aux5.Close
             rs_aux5.Open GlSqlAux, db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux5.RecordCount > 0 Then
                 rs_aux5.MoveFirst
                 While Not rs_aux5.EOF
                    'If VAR_COD0 < 1 And rs_aux5!estado_activo = "REG" Then
                    If VAR_BIEN <> rs_aux5!bien_codigo Then
                    
                       rs_aux4.AddNew
                       rs_aux4!ges_gestion = rs_aux5!ges_gestion
                       rs_aux4!compra_codigo = rs_aux5!compra_codigo
                       VAR_GES = rs_aux4!ges_gestion
                       Var_Comp = rs_aux4!compra_codigo
                       rs_aux4!adjudica_codigo = VAR_ADJUDICA
                       rs_aux4!bien_codigo = rs_aux5!bien_codigo
                       VAR_BIEN = rs_aux4!bien_codigo
                       rs_aux4!grupo_codigo = rs_aux5!grupo_codigo
                       rs_aux4!subgrupo_codigo = rs_aux5!subgrupo_codigo
                       rs_aux4!par_codigo = rs_aux5!par_codigo
                       rs_aux4!adjudica_cantidad = rs_aux5!compra_cantidad
                       rs_aux4!bien_precio_adjudica_bs = rs_aux5!compra_precio_unitario_bs
                       rs_aux4!bien_total_adjudica_bs = rs_aux5!compra_precio_total_bs
                       rs_aux4!tipo_moneda = rs_aux5!tipo_moneda
                       rs_aux4!unimed_codigo = rs_aux5!unimed_codigo
                       rs_aux4!unimed_codigo_empaque = rs_aux5!unimed_codigo_empaque
                       rs_aux4!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
                       rs_aux4!marca_codigo = rs_aux5!marca_codigo
                       rs_aux4!modelo_codigo = rs_aux5!modelo_codigo
                       rs_aux4!bien_nro_lote = 0
                       rs_aux4!bien_fecha_vencimiento = Date
                       rs_aux4!estado_codigo = "REG"
                       rs_aux4!usr_codigo = glusuario
                       rs_aux4!Fecha_Registro = Date
                      
                            
                       rs_aux4.Update
                        
                        VAR_COD0 = VAR_COD0 + 1
                       
                    End If
                    rs_aux5.MoveNext
                   
                 Wend
             End If
             db.Execute "UPDATE ao_compra_cabecera SET estado_codigo_eqp = 'APR' WHERE ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos.Recordset!compra_codigo & "' "
        Call OptFilGral3_Click
'End If
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        If Ado_datos.Recordset!estado_codigo = "REG" Then
            Call OptFilGral1_Click
        Else
            Call OptFilGral2_Click
        End If
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        'txt_codigo.Enabled = True
        FraDet3.Visible = True
        FraDet2.Visible = True
        FraDet1.Visible = True
'        FrmABMDet3.Visible = True
        FrmABMDet2.Visible = True
        FrmABMDet.Visible = True

        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
    If ExisteReg(Ado_datos.Recordset!edif_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
    If rs_datos!estado_codigo = "APR" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
          rs_datos!estado_codigo = "ERR"
          rs_datos!Fecha_Registro = Date
          rs_datos!usr_codigo = glusuario
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub

UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!Fecha_Registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnEliminar3_Click()
VAR_AP = "ENV"

'  sino = MsgBox("Está Seguro de eliminar el registro? ", vbYesNo + vbQuestion, "Atención")
' If sino = vbYes Then
 If Ado_datos13.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      
   db.Execute "UPDATE ao_compra_cabecera SET estado_codigo_eqp = 'REG' WHERE ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos13.Recordset!compra_codigo & "' "
    db.Execute " DELETE ao_compra_adjudica_bienes  WHERE ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos13.Recordset!compra_codigo & "' AND adjudica_codigo = '" & Ado_datos13.Recordset!adjudica_codigo & "' AND bien_codigo = '" & Ado_datos13.Recordset!bien_codigo & "' "
    Set rs_aux7 = New ADODB.Recordset
       If rs_aux7.State = 1 Then rs_aux7.Close
        SQL_FOR = "select * from ao_compra_adjudica_bienes WHERE ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos13.Recordset!compra_codigo & "' AND adjudica_codigo = '" & Ado_datos13.Recordset!adjudica_codigo & "' "
        rs_aux7.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux7.RecordCount > 0 Then
            
            'MsgBox "El código ya existe, consulte con el administrador del Sistema..."
            
        Else
            db.Execute " DELETE ao_compra_adjudica  WHERE ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos13.Recordset!compra_codigo & "' AND adjudica_codigo = '" & Ado_datos13.Recordset!adjudica_codigo & "' "
        End If
 
 Call OptFilGral3_Click
' End If
 
  Else
      MsgBox "NO se puede ANULAR !!. El registro esta APROBADO. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnEliminar4_Click()
VAR_GES = 0
Var_Comp = 0
Set rs_aux8 = New ADODB.Recordset
        If rs_aux8.State = 1 Then rs_aux8.Close
        rs_aux8.Open "select * from ao_compra_adjudica    ", db, adOpenKeyset, adLockOptimistic
        'If rs_aux2.RecordCount > 0 Then
             
             VAR_COD0 = 0
'             VAR_GES = ges_gestion
'             VAR_COMP = compra_codigo
                db.Execute "UPDATE ao_compra_cabecera SET estado_codigo_eqp = 'REG' WHERE ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND compra_codigo = '" & Ado_datos.Recordset!compra_codigo & "' "
             Set rs_aux9 = New ADODB.Recordset
             If rs_aux9.State = 1 Then rs_aux9.Close
                
             rs_aux9.Open GlSqlAux, db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux9.RecordCount > 0 Then
                 rs_aux9.MoveFirst
                 While Not rs_aux9.EOF
                    db.Execute " DELETE ao_compra_adjudica_bienes  WHERE ges_gestion = '" & rs_aux9!ges_gestion & "' AND compra_codigo = '" & rs_aux9!compra_codigo & "' AND bien_codigo = '" & rs_aux9!bien_codigo & "' AND estado_codigo = 'REG'  "
                    Set rs_aux7 = New ADODB.Recordset
                   If rs_aux7.State = 1 Then rs_aux7.Close
                    SQL_FOR = "select * from ao_compra_adjudica_bienes WHERE ges_gestion = '" & rs_aux9!ges_gestion & "' AND compra_codigo = '" & rs_aux9!compra_codigo & "' AND bien_codigo = '" & rs_aux9!bien_codigo & "'   "
                    rs_aux7.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                    If rs_aux7.RecordCount > 0 Then
                        
                        'MsgBox "El código ya existe, consulte con el administrador del Sistema..."
                        
                    Else
                        db.Execute " DELETE ao_compra_adjudica  WHERE ges_gestion = '" & rs_aux9!ges_gestion & "' AND compra_codigo = '" & rs_aux9!compra_codigo & "' AND estado_codigo = 'REG'"
                    End If
                    rs_aux9.MoveNext
                   
                 Wend
                  Call OptFilGral4_Click
             End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
        VAR_UNI = dtc_codigo1.Text
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        'SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & var_cod & "  "
        SQL_FOR = "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            var_cod = rs_aux1.RecordCount + 1
            'MsgBox "El código ya existe, consulte con el administrador del Sistema..."
            'var_cod = 0
            'Exit Sub
        Else
            'var_cod = rs_datos.RecordCount '+ 1
            var_cod = 1
        End If
        'var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        txt_codigo.Caption = var_cod
        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = glGestion    ' Year(Date)   'no cambia
        rs_datos!unidad_codigo = VAR_UNI
        'Actualiza correaltivo ...
        db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!doc_numero = "0"    'txt_campo1.Caption
        'rs_datos!correl_edificacion = 0
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!correl_bitacora = 0
     End If
     rs_datos!compra_fecha = dtpFecha1.Value
     rs_datos!solicitud_tipo = "15"
     rs_datos!venta_tipo = dtc_codigo2.Text
     rs_datos!edif_codigo = dtc_codigo3.Text
     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
        rs_datos!beneficiario_codigo = dtc_aux3.Text
     Else
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     End If
     Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text + " Cite: " + Txt_campo2.Caption
     rs_datos!compra_descripcion = Txt_descripcion.Text

     Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
            rs_datos!proceso_codigo = "CMX"
            rs_datos!subproceso_codigo = "CMX-01"
            rs_datos!etapa_codigo = "CMX-01-01"
            rs_datos!clasif_codigo = "CMX"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-207"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL

            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            If VAR_UNI = "DNINS" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNAJS" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNMAN" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNIREP" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNEME" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNMOD" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
        Case "5"    ' SERVICIO MODERNIZACION
     End Select
     rs_datos!poa_codigo = dtc_codigo10.Text
     If txt_obs.Text = "" Then
        rs_datos!compra_observaciones = dtc_desc3.Text   'txt_obs.Text
     Else
        rs_datos!compra_observaciones = txt_obs.Text
     End If
     'rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text

'     rs_datos!ges_gestion_ant = glGestion       'Year(Date)
'     If var_cod < 10 Then
'        rs_datos!unidad_codigo_ant = VAR_UNI + "-00000" + Trim(txt_codigo)
'     End If
'     If var_cod > 9 And var_cod < 100 Then
'        rs_datos!unidad_codigo_ant = VAR_UNI + "-0000" + Trim(txt_codigo)
'     End If
'     If var_cod > 99 And var_cod < 1000 Then
'        rs_datos!unidad_codigo_ant = VAR_UNI + "-000" + Trim(txt_codigo)
'     End If
'     If var_cod > 999 And var_cod < 10000 Then
'        rs_datos!unidad_codigo_ant = VAR_UNI + "-00" + Trim(txt_codigo)
'     End If
'     If var_cod > 9999 And var_cod < 100000 Then
'        rs_datos!unidad_codigo_ant = VAR_UNI + "-0" + Trim(txt_codigo)
'     End If
'     If var_cod > 99999 Then
'        rs_datos!unidad_codigo_ant = VAR_UNI + "-" + Trim(txt_codigo)
'     End If
'     rs_datos!solicitud_codigo_ant = 0
'     rs_datos!usr_codigo_aprueba = ""
'     rs_datos!fecha_aprueba = Date
     'rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     rs_datos!Fecha_Registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     If Ado_datos.Recordset!estado_codigo = "REG" Then
        Call OptFilGral1_Click
     Else
        Call OptFilGral2_Click
     End If
     rs_datos.MoveLast
     mbDataChanged = False

     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     
        FraDet3.Visible = True
        FraDet2.Visible = True
        FraDet1.Visible = True
'        FrmABMDet3.Visible = True
        FrmABMDet2.Visible = True
        FrmABMDet.Visible = True
        
'     dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
'     dtc_codigo9.Enabled = True
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo8.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo9.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo10.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
 If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        cr01.ReportFileName = App.Path & "\Reportes\Comex\rr_proceso_contratacion.rpt"
        cr01.WindowShowPrintSetupBtn = True
        cr01.WindowShowRefreshBtn = True
        cr01.StoredProcParam(0) = Me.Ado_datos.Recordset!solicitud_codigo
        cr01.StoredProcParam(1) = Glaux
        iResult = cr01.PrintReport
        If iResult <> 0 Then MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbCritical, "Error de impresión"
        cr01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir1_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle3.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        cr01.ReportFileName = App.Path & "\Reportes\comex\ar_orden_pago.rpt"
        cr01.WindowShowPrintSetupBtn = True
        cr01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
'          CR01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
'          CR01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        cr01.StoredProcParam(0) = Me.Ado_detalle3.Recordset!adjudica_codigo
        cr01.StoredProcParam(1) = Me.Ado_detalle3.Recordset!pago_codigo
        
        iResult = cr01.PrintReport
        If iResult <> 0 Then MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbCritical, "Error de impresión"
        cr01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnModDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
'  If rs_det5.RecordCount > 0 And rs_det5!estado_codigo = "REG" And Ado_detalle1.Recordset.RecordCount > 0 Then
If rs_det5.RecordCount > 0 And rs_det5!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
'    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False

'            aw_orden_pago.txt_codigo.Caption = Me.Ado_datos.Recordset!solicitud_codigo  'cod_cabecera
'            aw_orden_pago.txt_campo1.Text = Me.Ado_datos.Recordset!unidad_codigo  'Unidad
'            aw_orden_pago.Txt_descripcion = Me.dtc_desc1.Text
            aw_orden_pago.txt_codigo.Caption = Me.Ado_datos.Recordset!ges_gestion
            aw_orden_pago.lbl_adjudica.Caption = Me.Ado_detalle3.Recordset!adjudica_codigo
            aw_orden_pago.txtCodigo1.Caption = Me.Ado_detalle3.Recordset!pago_codigo
            aw_orden_pago.txt_campo1.Text = Me.Ado_detalle3.Recordset!beneficiario_codigo
            aw_orden_pago.Txt_descripcion.BoundText = aw_orden_pago.txt_campo1.BoundText
            
            aw_orden_pago.DTPFechaProg.Value = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_fecha_prog), Date, Me.Ado_detalle3.Recordset!pago_fecha_prog)
            aw_orden_pago.DTPFechaPago.Value = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_fecha_efectiva), Date, Me.Ado_detalle3.Recordset!pago_fecha_efectiva)
            aw_orden_pago.TxtMontoBs.Text = Me.Ado_detalle3.Recordset!pago_total_bs
            aw_orden_pago.TxtMontoDol.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_total_dol), 0, Me.Ado_detalle3.Recordset!pago_total_dol)
            aw_orden_pago.txt_factura.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_nro_cmpbte_factura), 0, Me.Ado_detalle3.Recordset!pago_nro_cmpbte_factura)
            aw_orden_pago.txtDoc.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_nro_autorizacion), 0, Me.Ado_detalle3.Recordset!pago_nro_autorizacion)
            
            aw_orden_pago.TxtConcepto.Text = Me.Ado_detalle3.Recordset!pago_descripcion
            aw_orden_pago.txt_respaldos.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_respaldos), "FACTURA", Me.Ado_detalle3.Recordset!pago_respaldos)
'            aw_orden_pago.dtc_codigo4A.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!beneficiario_codigo_resp), 0, Me.Ado_detalle3.Recordset!beneficiario_codigo_resp)
            aw_orden_pago.TxtCobrador = IIf(IsNull(Me.Ado_detalle3.Recordset!beneficiario_codigo_resp), "", Me.Ado_detalle3.Recordset!beneficiario_codigo_resp)
            aw_orden_pago.Show vbModal


    Call ABRIR_TABLA_DET

    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
'    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar el registro, verifique si está Aprobado o fue correctamente identificado !! ", vbExclamation
  End If
End Sub

Private Sub BtnModDetalle1_Click()
    If Ado_detalle1.Recordset.RecordCount > 0 Then
      If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
        marca1 = Ado_detalle1.Recordset.Bookmark
        swnuevo = 2
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FraDet2.Enabled = False
        FrmABMDet2.Enabled = False
        FraDet3.Enabled = False
'        FrmABMDet3.Enabled = False
        Fra_datos.Enabled = False
    
        Select Case dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
            Case "L"    'IMPORTACION DIRECTA CLIENTE
                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                
                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bien_codigo")
                
                frm_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
                
                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
                frm_solicitud_bienes.Txt_campo19.Text = Me.Ado_detalle1.Recordset("bien_cantidad_por_empaque")
                
                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
                frm_solicitud_bienes.Txt_campo15.Text = "10" 'Me.Ado_detalle1.Recordset("fosa_dimension_frente")
                
                frm_solicitud_bienes.lbl_det.Caption = "43340"
                frm_solicitud_bienes.Show vbModal
            Case "V"    'FACTURACION LOCAL
                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                
                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bien_codigo")
                    
                frm_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
                
                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
                
                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
    '            frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
    '            frm_solicitud_bienes.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
                frm_solicitud_bienes.lbl_det.Caption = "43340"
                frm_solicitud_bienes.Show vbModal
            
        End Select
        swnuevo = 0
        fraOpciones.Enabled = True
        FraNavega.Enabled = True
        FraDet2.Enabled = True
        FrmABMDet2.Enabled = True
        FraDet3.Enabled = True
'        FrmABMDet3.Enabled = True
    '    Fra_datos.Enabled = True
        Call ABRIR_TABLA_DET
        Ado_detalle1.Recordset.Move marca1 - 1
      Else
        MsgBox "No se puede MODIFICAR, porque ya está APROBADO o ANULADO, Verifique por favor!! ", vbExclamation
      End If
  Else
     MsgBox "No se puede MODIFICAR, el registro No fue identificado o No Existe, Verifique por favor ...", vbExclamation, "Validación de Registro"
  End If

End Sub

Private Sub BtnModDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_det4.RecordCount > 0 And rs_det4!estado_codigo = "REG" Then
     If Ado_detalle4.Recordset.RecordCount > 0 Then
        swnuevo = 2
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FraDet2.Enabled = False
        FrmABMDet2.Enabled = False
'        FraDet3.Enabled = False
'        FrmABMDet3.Enabled = False
        Fra_datos.Enabled = False

        '    'Call ABRIR_TABLA_DET
        'ges_gestion,     adjudica_fecha, proceso_codigo, subproceso_codigo, etapa_codigo,
        'clasif_codigo, doc_codigo, doc_numero,  adjudica_descripcion, adjudica_cantidad_total,  tipo_moneda,
        '    fecha_recibe_almacen, almacen_codigo, poa_codigo, estado_codigo,
         'usr_codigo , fecha_registro, hora_registro, usr_codigo_aprueba, fecha_aprueba
'
'            aw_bienes_adjudica.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
'            aw_bienes_adjudica.txt_campo1.Text = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
           ' aw_bienes_adjudica.Txt_descripcion.Caption = Me.dtc_desc1.Text
            aw_bienes_adjudica.txtCodigo1.Caption = Me.Ado_detalle4.Recordset("ges_gestion")
            'frm_ao_comex_adjudica.Txt_estado.Caption = "REG"
            aw_bienes_adjudica.lbl_adjudica.Caption = Me.Ado_detalle4.Recordset("adjudica_codigo")
            aw_bienes_adjudica.dtc_codigo5.Text = Me.Ado_detalle4.Recordset("beneficiario_codigo")
            aw_bienes_adjudica.dtc_desc5.BoundText = aw_bienes_adjudica.dtc_codigo5.BoundText
            aw_bienes_adjudica.dtc_aux4.BoundText = aw_bienes_adjudica.dtc_codigo5.BoundText
            aw_bienes_adjudica.dtc_aux5.BoundText = aw_bienes_adjudica.dtc_codigo5.BoundText

            aw_bienes_adjudica.txt_Nota.Text = IIf(IsNull(Me.Ado_detalle4.Recordset("nro_nota_remision")), "", Me.Ado_detalle4.Recordset("nro_nota_remision"))
            aw_bienes_adjudica.txt_total_bs.Text = IIf(IsNull(Me.Ado_detalle4.Recordset("adjudica_monto_bs")), 0, Me.Ado_detalle4.Recordset("adjudica_monto_bs"))
            aw_bienes_adjudica.txt_total_dol.Text = IIf(IsNull(Me.Ado_detalle4.Recordset!adjudica_monto_dol), 0, Me.Ado_detalle4.Recordset!adjudica_monto_dol)
            aw_bienes_adjudica.txtFecha.Value = IIf(IsNull(Me.Ado_detalle4.Recordset("fecha_inicio_contrato")), Date, Me.Ado_detalle4.Recordset("fecha_inicio_contrato"))
            aw_bienes_adjudica.txtFecha2.Value = IIf(IsNull(Me.Ado_detalle4.Recordset("fecha_fin_contrato")), Date, Me.Ado_detalle4.Recordset("fecha_fin_contrato"))
            aw_bienes_adjudica.txtFecha3.Value = IIf(IsNull(Me.Ado_detalle4.Recordset("fecha_recibe_almacen")), Date, Me.Ado_detalle4.Recordset("fecha_recibe_almacen"))
            aw_bienes_adjudica.txt_adjudica.Text = IIf(IsNull(Me.Ado_detalle4.Recordset("adjudica_cantidad_total")), Date, Me.Ado_detalle4.Recordset("adjudica_cantidad_total"))
            
            aw_bienes_adjudica.cmb_mes_ini = IIf(IsNull(Me.Ado_detalle4.Recordset!mes_inicio_crono), "ENERO", Me.Ado_detalle4.Recordset!mes_inicio_crono)
            aw_bienes_adjudica.txtCantCuota.Text = IIf(IsNull(Me.Ado_detalle4.Recordset!cantidad_cuotas_pag), "1", Me.Ado_detalle4.Recordset!cantidad_cuotas_pag)
            aw_bienes_adjudica.cmd_unimed2 = IIf(IsNull(Me.Ado_detalle4.Recordset!unimed_codigo_pag), "MES", Me.Ado_detalle4.Recordset!unimed_codigo_pag)
            
            aw_bienes_adjudica.txt_tipocambio.Text = GlTipoCambioOficial
'            aw_bienes_adjudica.txtSW.Text = Me.Ado_datos.Recordset!venta_tipo
            aw_bienes_adjudica.txt_pais.Text = VAR_PAIS
            
            aw_bienes_adjudica.Show vbModal
        swnuevo = 0
        fraOpciones.Enabled = True
        FraNavega.Enabled = True
        FraDet2.Enabled = True
        FrmABMDet2.Enabled = True
        
        
'        FraDet3.Enabled = True
'        FrmABMDet3.Enabled = True
    '    Fra_datos.Enabled = True
     Else
        MsgBox "No se puede Modificar un registro inexistente, vuelva a intentar!! ", vbExclamation
     End If
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        
        FraDet3.Visible = False
        FraDet2.Visible = False
        FraDet1.Visible = False
'        FrmABMDet3.Visible = False
        FrmABMDet2.Visible = False
        FrmABMDet.Visible = False
        
        VAR_SW = "MOD"
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        dtc_desc11.SetFocus
    '    BtnVer.Visible = True
'        dtc_codigo9.Enabled = False
        FraGrabarCancelar.Visible = True
        BtnCancelar.Visible = True
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
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

Private Sub BtnVer_Click()
  On Error GoTo QError
  If rs_datos!estado_codigo = "APR" Then
    Dim ARCH_FOTO As String
    Dim SW0 As String
    Select Case Left(Trim(Ado_datos.Recordset("edif_codigo")), 1)
        Case "1"    'CHQ
            VAR_DPTO = "CHQ"
        Case "2"    'LPZ
            VAR_DPTO = "LPZ"
        Case "3"    'CBB
            VAR_DPTO = "CBB"
        Case "4"    'SCZ
            VAR_DPTO = "SCZ"
        Case "5"    'PTS
            VAR_DPTO = "PTS"
        Case "6"    'ORU
            VAR_DPTO = "ORU"
        Case "7"    'TJA
            VAR_DPTO = "TJA"
        Case "8"    'BEN
            VAR_DPTO = "BEN"
        Case "9"    'PDO
            VAR_DPTO = "PDO"
    End Select
    If Ado_datos.Recordset!archivo_respaldo_cargado = "N" Then
      'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "DED2"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
      SW0 = 1
    Else
      'MsgBox ""
      'negocia_codigo, unidad_codigo, negocia_fecha_inicio as fecha1, negocia_descripcion, estado_codigo, fecha_registro, usr_codigo, solicitud_tipo as codigo2, edif_codigo as codigo3, beneficiario_codigo as codigo4, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, hora_registro, ges_gestion, archivo_respaldo, archivo_respaldo_cargado
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "DED2"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
          SW0 = 1
      Else
        SW0 = 0
        'e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
        e = ShellExecute(0, vbNullString, App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\" & Trim(Ado_datos.Recordset("archivo_respaldo")), vbNullString, vbNullString, vbNormalFocus)
      End If
    End If
    '    If SW0 = 1 Then
    '    '    If GlServidor = "SRVPRO" Then
    '    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
    '    '    Else
    '            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo)
    '            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo) + ".JPG"
    '    '    End If
    '        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    '        CodBien = Ado_datos.Recordset!edif_codigo
    '        If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
    '            MsgBox "Se cargo la Imagen Correctamente !!"
    '        Else
    '            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    '        End If
    '    Else
    '        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    '        Image2 = Img_Foto
    '    End If
  Else
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validación de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If
End Sub







Private Sub Command5_Click()

End Sub

Private Sub Command3_Click()

End Sub





Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub

'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'End Sub

'Private Sub dtc_codigo8_Click(Area As Integer)
'    dtc_desc8.BoundText = dtc_codigo8.BoundText
'End Sub

'Private Sub dtc_codigo9_Click(Area As Integer)
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
'End Sub

'Private Sub dtc_codigo9_LostFocus()
''  If VAR_SW = "ADD" Then
''    Set rs_aux2 = New ADODB.Recordset
''    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
''    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''    If rs_aux2.RecordCount > 0 Then
''        rs_aux2!correl_doc = rs_aux2!correl_doc + 1
''        txt_campo1.Caption = rs_aux2!correl_doc
''        rs_aux2.Update
''    End If
''  End If
'  txt_aux9.Text = dtc_desc9.Text
'End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
'    dtc_codigo5.BoundText = dtc_desc5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub

'Private Sub pnivel5(codigo5 As String)
'   'Dim strConsultaF As String
'   'strConsultaF = "select * from gc_proceso_nivel2 where proceso_codigo = '" & codigo5 & "'"
'
'   Set dtc_codigo6.RowSource = Nothing
'   'Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_codigo6.ReFill
'   dtc_codigo6.BoundText = Empty
'
'   Set dtc_desc6.RowSource = Nothing
'   'Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_desc6.ReFill
'   dtc_desc6.BoundText = Empty
'End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
 pnivel1 (dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub pnivel1(codigo1 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"

   Set dtc_codigo10.RowSource = Nothing
'   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty

   Set dtc_desc10.RowSource = Nothing
   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub

'Private Sub pnivel11(codigo1 As String)
'   Dim strConsultaF As String
'   'strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'   strConsultaF = "Select * from gv_personal_contratado where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"
'
'   Set dtc_codigo11.RowSource = Nothing
'   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo11.ReFill
'   dtc_codigo11.BoundText = Empty
'
'   Set dtc_desc11.RowSource = Nothing
'   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc11.ReFill
'   dtc_desc11.BoundText = Empty
'End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc11_LostFocus()
    Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text + " Cite: " + Txt_campo2.Caption
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc3_LostFocus()
'    dtc_codigo4.Text = dtc_aux3.Text
'    Txt_descripcion.Text = lbl_titulo + " - " + dtc_desc3.Text
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'    dtc_codigo6.BoundText = dtc_desc6.BoundText
''    Call pnivel6(dtc_codigo6.BoundText)
''    dtc_desc7.Enabled = True
'End Sub

'Private Sub pnivel6(codigo6 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from gc_proceso_nivel3 where subproceso_codigo = '" & codigo6 & "'"
'
'   Set dtc_codigo7.RowSource = Nothing
'   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute("EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_codigo7.ReFill
'   dtc_codigo7.BoundText = Empty
'
'   Set dtc_desc7.RowSource = Nothing
'   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_desc7.ReFill
'   dtc_desc7.BoundText = Empty
'End Sub

'Private Sub dtc_desc7_Click(Area As Integer)
'    dtc_codigo7.BoundText = dtc_desc7.BoundText
'End Sub

'Private Sub dtc_desc8_Click(Area As Integer)
'    dtc_codigo8.BoundText = dtc_desc8.BoundText
'    Call pnivel8(dtc_codigo8.BoundText)
'    'dtc_desc9.Enabled = True
'    dtc_codigo9.Enabled = True
'End Sub

'Private Sub pnivel8(codigo8 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_documentos_respaldo where clasif_codigo = '" & codigo8 & "'"
'
'   Set dtc_codigo9.RowSource = Nothing
'   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo9.ReFill
'   dtc_codigo9.BoundText = Empty
'
'   Set dtc_desc9.RowSource = Nothing
'   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc9.ReFill
'   dtc_desc9.BoundText = Empty
'End Sub

'Private Sub dtc_desc9_Click(Area As Integer)
'    dtc_codigo9.BoundText = dtc_codigo9.BoundText
'End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    parametro = Aux
    Glaux = Aux
    '    Aux = "COMEX"
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    GlSqlAux = queryinicial
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'JQA 2014-JUL-14
    'db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
    FraNavega.Caption = lbl_titulo.Caption
    'lbl_titulo2.Caption = lbl_titulo.Caption
    If Glaux = "PROVI" Then
        FraDet2.Caption = "EQUIPOS A COMPRAR"
    Else
        FraDet2.Caption = "EQUIPOS A IMPORTAR"
    End If
'    If Glaux = "PROVI" Then
'    End If
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText

'    'gc_tipo_solicitud
'    Set rs_datos11 = New ADODB.Recordset
'    If rs_datos11.State = 1 Then rs_datos11.Close
'    rs_datos11.Open "Select * from gc_tipo_solicitud where solicitud_tipo = '3' order by solicitud_tipo", db, adOpenStatic
'    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
    
    'ac_tipo_compra_venta
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "select * from ac_tipo_compra_venta where venta_tipo = 'L' or venta_tipo = 'V' ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from fo_proyectos_ejecucion order by pro_codigo_det_descripcion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText

    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText

'
'    Select Case Glaux
'        Case "PROVI"    'PROVISION DE EQUIPOS
'            rs_datos2.Open "Select * from gc_tipo_solicitud where solicitud_tipo = '3' order by solicitud_tipo", db, adOpenStatic
'        Case "TRANS"    'TRANSPORTE
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'TRANSP' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'        Case "ADUAN"    'DESADUANIZACION
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'ADUANA'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'        Case "DESCA"    'DESCARGUIO Y OTROS
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'DESCAR'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    End Select

    'pc_poa_actividad
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText

    'gc_beneficiario (Personal CGI)
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

'Private Sub ABRIR_TABLA()
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "select solicitud_codigo, unidad_codigo, solicitud_justificacion, solicitud_observaciones, estado_codigo, fecha_registro, usr_codigo, hora_registro, ges_gestion, solicitud_fecha_solicitud as fecha1,  solicitud_fecha_recepción as fecha2, solicitud_tipo as codigo2, beneficiario_codigo as codigo4, beneficiario_codigo_resp as codigo11, edif_codigo as codigo3, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, archivo_respaldo, archivo_respaldo_cargado, ges_gestion_ant, unidad_codigo_ant, solicitud_codigo_ant, usr_codigo_aprueba, fecha_aprueba, hora_aprueba From ao_solicitud WHERE estado_codigo = 'REG' "
'    queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub

Private Sub ABRIR_TABLA_DET()
    'BIENES (Equipos a Comprar)
'    Select Case Glaux
'        Case "PROVI"    'PROVISION DE EQUIPOS
'            Set rs_det1 = New ADODB.Recordset
'            If rs_det1.State = 1 Then rs_det1.Close
'            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            Set Ado_detalle1.Recordset = rs_det1
'            If Ado_detalle1.Recordset.RecordCount > 0 Then
'                VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
'                dg_det1.Visible = True
'                Set dg_det1.DataSource = Ado_detalle1.Recordset
'            Else
'                dg_det1.Visible = False
'                Set dg_det1.DataSource = rsNada
'            End If
'        Case "TRANS"    'TRANSPORTE
'            Set rs_det1 = New ADODB.Recordset
'            If rs_det1.State = 1 Then rs_det1.Close
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'TRANSP' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            Set Ado_detalle1.Recordset = rs_det1
'            Set dg_det1.DataSource = Ado_detalle1.Recordset
'            If Ado_detalle1.Recordset.RecordCount > 0 Then
'                VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
'                dg_det1.Visible = True
'                Set dg_det1.DataSource = Ado_detalle1.Recordset
'            Else
'                dg_det1.Visible = False
'                Set dg_det1.DataSource = rsNada
'            End If
'        Case "ADUAN"    'DESADUANIZACION
'            Set rs_det1 = New ADODB.Recordset
'            If rs_det1.State = 1 Then rs_det1.Close
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'ADUANA'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            Set Ado_detalle1.Recordset = rs_det1
'            Set dg_det1.DataSource = Ado_detalle1.Recordset
'            If Ado_detalle1.Recordset.RecordCount > 0 Then
'                VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
'                dg_det1.Visible = True
'                Set dg_det1.DataSource = Ado_detalle1.Recordset
'            Else
'                dg_det1.Visible = False
'                Set dg_det1.DataSource = rsNada
'            End If
'        Case "DESCA"    'DESCARGUIO Y OTROS
'            Set rs_det1 = New ADODB.Recordset
'            If rs_det1.State = 1 Then rs_det1.Close
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'DESCAR'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            Set Ado_detalle1.Recordset = rs_det1
'            Set dg_det1.DataSource = Ado_detalle1.Recordset
'            If Ado_detalle1.Recordset.RecordCount > 0 Then
'                VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
'                dg_det1.Visible = True
'                Set dg_det1.DataSource = Ado_detalle1.Recordset
'            Else
'                dg_det1.Visible = False
'                Set dg_det1.DataSource = rsNada
'            End If
'
'             Case "UALMI"    'DESCARGUIO Y OTROS
'            Set rs_det1 = New ADODB.Recordset
'            If rs_det1.State = 1 Then rs_det1.Close
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            Set Ado_detalle1.Recordset = rs_det1
''            Set dg_det1.DataSource = Ado_detalle1.Recordset
'            If Ado_detalle1.Recordset.RecordCount > 0 Then
''                VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
''                dg_det1.Visible = True
''                Set dg_det1.DataSource = Ado_detalle1.Recordset
'            Else
'                dg_det1.Visible = False
'                Set dg_det1.DataSource = rsNada
'            End If
''        Case Else
''            Set rs_det1 = New ADODB.Recordset
''            If rs_det1.State = 1 Then rs_det1.Close
''            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '43340'   ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    End Select



'    rs_det1.Open "select * from ao_compra_detalle where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText

'    Set rs_det2 = New ADODB.Recordset
'    If rs_det2.State = 1 Then rs_det2.Close
'    rs_det2.Open "select * from ao_compra_apertura_sobres where ges_gestion = '" & Ado_detalle4.Recordset!ges_gestion & "' and adjudica_codigo = " & Ado_detalle4.Recordset!adjudica_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Ado_detalle4.Recordset = rs_det2
'    If Ado_detalle4.Recordset.RecordCount > 0 Then
'        Set rs_det3 = New ADODB.Recordset
'        If rs_det3.State = 1 Then rs_det3.Close
'        rs_det3.Open "select * from ao_compra_planilla_pagos where  adjudica_codigo = " & rs_det2!adjudica_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'        Set Ado_detalle3.Recordset = rs_det3
'         Set dg_det3.DataSource = Ado_detalle3.Recordset
'        If Ado_detalle3.Recordset.RecordCount > 0 Then
'              dg_det3.Visible = True
'                Set dg_det3.DataSource = Ado_detalle3.Recordset
'            Else
''                dg_det3.Visible = False
'                Set dg_det3.DataSource = rsNada
'        End If
'        dg_det2.Visible = True
'        Set dg_det2.DataSource = Ado_detalle4.Recordset
'    Else
'        dg_det3.Visible = True
'        Set dg_det3.DataSource = rsNada
'        dg_det2.Visible = True
'        Set dg_det2.DataSource = rsNada
'    End If
 

'         Set rs_det3 = New ADODB.Recordset
'        If rs_det3.State = 1 Then rs_det3.Close
'        rs_det3.Open "select * from ao_compra_planilla_pagos where ges_gestion = " & rs_det4!ges_gestion & " and adjudica_codigo = " & rs_det4!adjudica_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'        Set Ado_detalle3.Recordset = rs_det3
'         Set dg_det3.DataSource = Ado_detalle3.Recordset
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub


Private Sub compra_cant()
             
Set rs_aux8 = New ADODB.Recordset
'        If rs_aux8.State = 1 Then rs_aux8.Close
'        rs_aux8.Open "select * from ao_compra_adjudica    ", db, adOpenKeyset, adLockOptimistic
        'If rs_aux2.RecordCount > 0 Then
             
             VAR_COD0 = 0
             
             Set rs_aux9 = New ADODB.Recordset
            If rs_aux9.State = 1 Then rs_aux9.Close
             rs_aux9.Open GlSqlAux, db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux9.RecordCount > 0 Then
                 rs_aux9.MoveFirst
                 While Not rs_aux9.EOF
                    
                    VAR_COD0 = VAR_COD0 + rs_aux9!compra_cantidad
                    rs_aux9.MoveNext
                   
                 Wend
                  
           
             End If
           Ado_datos.Caption = "                                                                Cantidad total --> " + Str(VAR_COD0)
End Sub


Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
    If VAR_SW <> "ADD" Then
        VAR_COD4 = parametro
        VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'        Call proveedor
''        Call ABRIR_TABLA_AUX2
'        Call compra_cant

    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
'    Call proveedor
    'FraDet1.Caption = "BITÁCORA DE: " + dtc_desc1.Text
'    txt_aux9.Text = dtc_desc9.Text
    If Ado_datos.Recordset!estado_codigo_eqp = "APR" Then
            FrmABMDet2.Visible = False
    Else
            FrmABMDet2.Visible = True
    End If
  Else
'    Set dg_det1.DataSource = rsNada

    Set dg_det2.DataSource = rsNada
    Set dg_det3.DataSource = rsNada
  End If
End Sub

Private Sub Ado_datos13_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
If Ado_datos13.Recordset.RecordCount > 0 Then
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
    If VAR_AP <> "APR" Or VAR_AP <> "ENV" Then
     Call proveedor
    End If
    
    
        Call proveedor
'        Call ABRIR_TABLA_AUX2
        Call compra_cant

End If
End Sub
Private Sub Ado_detalle4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
If Ado_detalle4.Recordset.RecordCount > 0 Then
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    Set rs_det5 = New ADODB.Recordset
        If rs_det5.State = 1 Then rs_det5.Close
        rs_det5.Open "select * from ao_compra_planilla_pagos where ges_gestion = " & Ado_detalle4.Recordset!ges_gestion & " AND adjudica_codigo = " & Ado_detalle4.Recordset!adjudica_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle3.Recordset = rs_det5
        If Ado_detalle3.Recordset.RecordCount > 0 Then
                dg_det3.Visible = True
                Set dg_det3.DataSource = Ado_detalle3.Recordset
'                Call ABRIR_TABLA_DET
            Else
            Set Ado_detalle3.Recordset = rs_det5
'         Set dg_det3.DataSource = Ado_detalle3.Recordset
'                dg_det3.Visible = False
                Set dg_det3.DataSource = rsNada
        End If
        Set dg_det3.DataSource = Ado_detalle3.Recordset ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
         dg_det3.Visible = True

End If
End Sub

Private Sub proveedor()
    If Ado_datos13.Recordset.RecordCount > 0 Then
        Set rs_det4 = New ADODB.Recordset
        If rs_det4.State = 1 Then rs_det4.Close
               rs_det4.Open "select * from ao_compra_apertura_sobres where ges_gestion = '" & Ado_datos13.Recordset!ges_gestion & "' AND adjudica_codigo = '" & Ado_datos13.Recordset!adjudica_codigo & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle4.Recordset = rs_det4
        Set dg_det2.DataSource = rs_det4.DataSource
End If

 End Sub

'Private Sub Ado_detalle3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    VAR_COD4 = parametro
'    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'End Sub

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
    VAR_SW = "ADD"
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    'txt_codigo.Enabled = False
'    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
'    rs_datos.AddNew
    Ado_datos.Recordset.AddNew
    dtc_desc11.SetFocus
    'dtc_desc1.BackColor = &H80000005
    dtc_codigo1.Text = parametro
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_desc2.Locked = True
    Select Case parametro
        Case "DVTA"        'INI COMERCIAL
            dtc_codigo2.Text = 3
        Case "COMEX"        'INI COMEX
            dtc_codigo2.Text = 3
        Case "DNINS"                        'INI GRABA INSTALACIONES
            '
            dtc_codigo2.Text = 4
        Case "DNAJS"
        
             dtc_codigo2.Text = 4
        Case "UALMI"
            '
            dtc_codigo2.Text = 4
        Case "DNMAN"
            '
            dtc_codigo2.Text = 4
        Case Else
            dtc_codigo2.Text = 5
    End Select
    dtc_desc2.BoundText = dtc_codigo2.BoundText
'    dtc_codigo5.Text = "COM"
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    dtc_codigo6.Text = "COM-01"
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'    dtc_codigo7.Text = "COM-01-02"
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'    BtnVer.Visible = False
'    dtc_codigo9.Enabled = False
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

Private Function ExisteReg(Unidad As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function


Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from av_compra_cabecera_detalle where estado_codigo_eqp = 'REG' AND unidad_codigo_adm = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from av_compra_cabecera_detalle where unidad_codigo_adm = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral3_Click()
Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from av_compra_adjudica_bienes where estado_codigo = 'REG' AND adjudica_codigo = " & VAR_ADJUDICA & " "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
    Set Ado_datos13.Recordset = rs_datos.DataSource
    Set dg_datos1.DataSource = Ado_datos13.Recordset
End Sub

Private Sub OptFilGral4_Click()
Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from av_compra_adjudica_bienes  ORDER BY adjudica_codigo ASC "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
    Set Ado_datos13.Recordset = rs_datos.DataSource
    Set dg_datos1.DataSource = Ado_datos13.Recordset
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


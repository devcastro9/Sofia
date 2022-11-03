VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form tw_tecnico_cronograma 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Instalaciones - Cronograma por Contrato"
   ClientHeight    =   10005
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   15735
   ForeColor       =   &H00000000&
   Icon            =   "tw_tecnico_cronograma.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   1.52901e8
   ScaleMode       =   0  'User
   ScaleWidth      =   2.03549e6
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H80000010&
      Height          =   975
      Left            =   4080
      TabIndex        =   90
      Top             =   8160
      Visible         =   0   'False
      Width           =   10335
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   270
         Left            =   120
         TabIndex        =   99
         Top             =   600
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESPERE UN MOMENTO POR FAVOR !!!"
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
         Left            =   240
         TabIndex        =   91
         Top             =   240
         Width           =   9900
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20520
      TabIndex        =   51
      Top             =   0
      Width           =   20520
      Begin VB.PictureBox BtnImprimir3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7200
         Picture         =   "tw_tecnico_cronograma.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   92
         ToolTipText     =   "Organización vs. Cronograma por Contrato"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   18000
         Picture         =   "tw_tecnico_cronograma.frx":12CF
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   89
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton BtnImprimir0 
         BackColor       =   &H00808000&
         Caption         =   "R-224"
         Height          =   600
         Left            =   10440
         Picture         =   "tw_tecnico_cronograma.frx":1A91
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Imprime Formulario"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "tw_tecnico_cronograma.frx":204E
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10920
         Picture         =   "tw_tecnico_cronograma.frx":2258
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8400
         Picture         =   "tw_tecnico_cronograma.frx":269A
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   105
         Picture         =   "tw_tecnico_cronograma.frx":2E59
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   1
         ToolTipText     =   "Editar Datos de ""Cabecera Cronograma"""
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "tw_tecnico_cronograma.frx":376E
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   2
         ToolTipText     =   "Anula el Registro Activo"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "tw_tecnico_cronograma.frx":3EBA
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   3
         ToolTipText     =   "Aprueba el Cronograma"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "tw_tecnico_cronograma.frx":46ED
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   4
         ToolTipText     =   "Busca un Registro"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5640
         Picture         =   "tw_tecnico_cronograma.frx":4EA2
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   5
         ToolTipText     =   "Imprime Lista de Cronogramas"
         Top             =   0
         Width           =   1400
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13095
         TabIndex        =   55
         Top             =   180
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
      ScaleWidth      =   20400
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   20400
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "tw_tecnico_cronograma.frx":576F
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   6
         Top             =   0
         Width           =   1280
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5475
         Picture         =   "tw_tecnico_cronograma.frx":5F45
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   7
         Top             =   0
         Width           =   1420
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13455
         TabIndex        =   50
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   1935
      TabIndex        =   43
      Top             =   7350
      Width           =   1935
      Begin VB.PictureBox BtnAnlDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         Picture         =   "tw_tecnico_cronograma.frx":6831
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   87
         ToolTipText     =   "Borra el ""Detalle del Cronograma por Contrato"""
         Top             =   960
         Width           =   1215
      End
      Begin VB.PictureBox BtnAddDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         Picture         =   "tw_tecnico_cronograma.frx":6F7D
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   8
         ToolTipText     =   "Genera Nuevo ""Detalle del Cronograma por Contrato"""
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cronogr."
         Height          =   640
         Left            =   600
         Picture         =   "tw_tecnico_cronograma.frx":773C
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   1110
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAprobar2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Aprobar"
         Height          =   640
         Left            =   960
         Picture         =   "tw_tecnico_cronograma.frx":8EBE
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Aprueba Registro Identificado"
         Top             =   180
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle2AA 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Anular-->"
         Height          =   640
         Left            =   960
         Picture         =   "tw_tecnico_cronograma.frx":90C8
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   945
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   120
         Picture         =   "tw_tecnico_cronograma.frx":9D92
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Modifica la Cobranza Identifiacada"
         Top             =   990
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DEL CRONOGRAMA POR CONTRATO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2085
      Left            =   2160
      TabIndex        =   42
      Top             =   7260
      Width           =   15975
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "tw_tecnico_cronograma.frx":AA5C
         Height          =   1785
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   3149
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "fmes_plan"
            Caption         =   "Mes"
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
            DataField       =   "dia_correl"
            Caption         =   "#.Dia"
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
            DataField       =   "dia_fecha"
            Caption         =   "Fecha"
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
            DataField       =   "dia_nombre"
            Caption         =   "Nombre.Dia"
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
            DataField       =   "horario_codigo"
            Caption         =   "Horario"
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
            DataField       =   "hora_ingreso"
            Caption         =   "Hora.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "hora_salida"
            Caption         =   "Hora.Fin"
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
         BeginProperty Column07 
            DataField       =   "nro_total_horas"
            Caption         =   "Nro.Horas"
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
         BeginProperty Column08 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Equipo"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Tecnico.1"
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
            DataField       =   "beneficiario_codigo_resp2"
            Caption         =   "Tecnico.2"
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
         BeginProperty Column11 
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
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
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               ColumnWidth     =   1470.047
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   4860.284
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox FrmABMDet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   5.688
      ScaleMode       =   4  'Character
      ScaleWidth      =   16.125
      TabIndex        =   32
      Top             =   5835
      Width           =   1935
      Begin VB.PictureBox BtnModDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "tw_tecnico_cronograma.frx":AA77
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   80
         ToolTipText     =   "Editar Datos del Equipo seleccionado"
         Top             =   120
         Width           =   1430
      End
      Begin VB.CommandButton BtnImprimir2 
         BackColor       =   &H80000018&
         Caption         =   "Cronogr."
         Height          =   640
         Left            =   585
         Picture         =   "tw_tecnico_cronograma.frx":B38C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   465
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   135
         Picture         =   "tw_tecnico_cronograma.frx":CB0E
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Adiciona Detalle"
         Top             =   96
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Anular-->"
         Height          =   640
         Left            =   120
         Picture         =   "tw_tecnico_cronograma.frx":CF50
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   828
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   7440
      TabIndex        =   13
      Top             =   765
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8837
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CABECERA CRONOGRAMA"
      TabPicture(0)   =   "tw_tecnico_cronograma.frx":DC1A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CRONOGRAMA POR EQUIPO"
      TabPicture(1)   =   "tw_tecnico_cronograma.frx":DC36
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEdita"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrmCabecera 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4590
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   10575
         Begin VB.TextBox txt_venta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "tec_plan_codigo"
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
            Height          =   315
            Left            =   9300
            Locked          =   -1  'True
            TabIndex        =   141
            Top             =   300
            Width           =   885
         End
         Begin VB.Frame Fra_Total 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Periodos        Periodicidad          Fecha Inicio Crono.        Fecha Final Crono.           Mes Inicio Cronograma"
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
            Height          =   675
            Left            =   60
            TabIndex        =   120
            Top             =   3720
            Width           =   10455
            Begin VB.ComboBox cmb_mes_ini 
               DataField       =   "mes_inicio_crono"
               DataSource      =   "Ado_datos"
               Height          =   315
               ItemData        =   "tw_tecnico_cronograma.frx":DC52
               Left            =   7920
               List            =   "tw_tecnico_cronograma.frx":DC7A
               TabIndex        =   128
               Text            =   "SEPTIEMBRE"
               Top             =   240
               Width           =   1900
            End
            Begin VB.TextBox txt_cant 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "tec_cantidad_unidades"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   360
               TabIndex        =   127
               Text            =   "0"
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox cmd_unimed2 
               BackColor       =   &H00FFFFFF&
               DataField       =   "unimed_codigo"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "tw_tecnico_cronograma.frx":DCE3
               Left            =   1800
               List            =   "tw_tecnico_cronograma.frx":DCF9
               TabIndex        =   126
               Text            =   "ANUAL"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   9550
               TabIndex        =   124
               Top             =   975
               Width           =   270
            End
            Begin VB.ComboBox cmb_dia 
               DataField       =   "dia_nombre"
               DataSource      =   "Ado_datos"
               Height          =   315
               ItemData        =   "tw_tecnico_cronograma.frx":DD21
               Left            =   240
               List            =   "tw_tecnico_cronograma.frx":DD3D
               TabIndex        =   123
               Text            =   "AUTOMATICO"
               Top             =   480
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.ComboBox cmb_lunes 
               BackColor       =   &H00FFFFFF&
               DataField       =   "lunes_cambia"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "tw_tecnico_cronograma.frx":DD89
               Left            =   3840
               List            =   "tw_tecnico_cronograma.frx":DD93
               TabIndex        =   122
               Text            =   "SI"
               Top             =   1440
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.ComboBox cmb_primero 
               BackColor       =   &H00FFFFFF&
               DataField       =   "primero_mes"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "tw_tecnico_cronograma.frx":DD9F
               Left            =   9120
               List            =   "tw_tecnico_cronograma.frx":DDA9
               TabIndex        =   121
               Text            =   "SI"
               Top             =   1440
               Visible         =   0   'False
               Width           =   735
            End
            Begin MSDataListLib.DataCombo dtc_desc5 
               Bindings        =   "tw_tecnico_cronograma.frx":DDB5
               DataField       =   "horario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   5760
               TabIndex        =   125
               Top             =   960
               Width           =   4080
               _ExtentX        =   7197
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "horario_descripcion"
               BoundColumn     =   "horario_codigo"
               Text            =   "0"
            End
            Begin MSComCtl2.DTPicker lbl_fecha_ini 
               DataField       =   "fecha_inicio_tec"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd-MMM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   3480
               TabIndex        =   129
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   -2147483646
               CheckBox        =   -1  'True
               Format          =   114491393
               CurrentDate     =   44197
               MinDate         =   36526
            End
            Begin MSComCtl2.DTPicker lbl_fecha_fin 
               DataField       =   "fecha_fin_tec"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd-MMM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   5640
               TabIndex        =   130
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   114491393
               CurrentDate     =   44561
               MinDate         =   36526
            End
            Begin MSDataListLib.DataCombo dtc_horaD5 
               Bindings        =   "tw_tecnico_cronograma.frx":DDCE
               DataField       =   "horario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   5040
               TabIndex        =   131
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "hora_salida2"
               BoundColumn     =   "horario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_horaC5 
               Bindings        =   "tw_tecnico_cronograma.frx":DDE7
               DataField       =   "horario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4320
               TabIndex        =   132
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "hora_ingreso2"
               BoundColumn     =   "horario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_horaB5 
               Bindings        =   "tw_tecnico_cronograma.frx":DE00
               DataField       =   "horario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3600
               TabIndex        =   133
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "hora_salida"
               BoundColumn     =   "horario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_horaA5 
               Bindings        =   "tw_tecnico_cronograma.frx":DE19
               DataField       =   "horario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2880
               TabIndex        =   134
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "hora_ingreso"
               BoundColumn     =   "horario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_codigo5 
               Bindings        =   "tw_tecnico_cronograma.frx":DE32
               DataField       =   "horario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2160
               TabIndex        =   135
               Top             =   960
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "horario_codigo"
               BoundColumn     =   "horario_codigo"
               Text            =   "0"
            End
            Begin VB.Label lbl_unimed 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               DataField       =   "unimed_codigo"
               DataSource      =   "Ado_datos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   285
               Left            =   1920
               TabIndex        =   140
               Top             =   300
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Programar Hora Fija"
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
               Height          =   270
               Left            =   2160
               TabIndex        =   139
               Top             =   720
               Width           =   1965
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Programar Día Fijo"
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
               Height          =   195
               Left            =   240
               TabIndex        =   138
               Top             =   720
               Width           =   1740
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Asignar el Lunes al primer equipo 4 Hrs.?"
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
               Height          =   195
               Left            =   240
               TabIndex        =   137
               Top             =   1485
               Visible         =   0   'False
               Width           =   3540
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Asignar el 1ro. del Mes al primer equipo 4 Hrs.?"
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
               Height          =   195
               Left            =   5040
               TabIndex        =   136
               Top             =   1485
               Visible         =   0   'False
               Width           =   4110
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Fra_datos 
            BackColor       =   &H00C0C0C0&
            Caption         =   $"tw_tecnico_cronograma.frx":DE4B
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
            Height          =   660
            Left            =   60
            TabIndex        =   113
            Top             =   2625
            Width           =   10425
            Begin MSDataListLib.DataCombo dtc_desc7 
               Bindings        =   "tw_tecnico_cronograma.frx":DEDF
               DataField       =   "zpiloto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   240
               TabIndex        =   114
               Top             =   240
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "zpiloto_descripcion"
               BoundColumn     =   "zpiloto_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc4 
               Bindings        =   "tw_tecnico_cronograma.frx":DEF8
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4320
               TabIndex        =   115
               Top             =   255
               Width           =   4335
               _ExtentX        =   7646
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo4 
               Bindings        =   "tw_tecnico_cronograma.frx":DF11
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   8160
               TabIndex        =   116
               Top             =   480
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_aux4 
               Bindings        =   "tw_tecnico_cronograma.frx":DF2A
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   7200
               TabIndex        =   117
               Top             =   480
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "tipoben_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "-"
            End
            Begin MSDataListLib.DataCombo dtc_codigo7 
               Bindings        =   "tw_tecnico_cronograma.frx":DF43
               DataField       =   "zpiloto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3000
               TabIndex        =   118
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "zpiloto_codigo"
               BoundColumn     =   "zpiloto_codigo"
               Text            =   "0"
            End
            Begin MSComCtl2.DTPicker DTPfechasol 
               DataField       =   "tec_plan_fecha"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd-MMM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   8760
               TabIndex        =   119
               Top             =   240
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   114491393
               CurrentDate     =   44197
               MinDate         =   36526
            End
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   5770
            TabIndex        =   111
            Top             =   900
            Width           =   270
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   10155
            TabIndex        =   110
            Top             =   900
            Width           =   260
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6240
            TabIndex        =   109
            Top             =   310
            Width           =   345
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Caption         =   $"tw_tecnico_cronograma.frx":DF5C
            ForeColor       =   &H00400000&
            Height          =   1155
            Left            =   120
            TabIndex        =   101
            Top             =   1320
            Width           =   10335
            Begin VB.TextBox TxtConcepto 
               BackColor       =   &H00C0C0C0&
               DataField       =   "tec_plan_concepto"
               DataSource      =   "Ado_datos"
               Height          =   405
               Left            =   990
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   102
               Top             =   600
               Width           =   9180
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               DataField       =   "cantidad_unidades_vta"
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
               Height          =   300
               Left            =   2160
               TabIndex        =   108
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               DataField       =   "unimed_codigo_vta"
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
               Height          =   300
               Left            =   4320
               TabIndex        =   107
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               DataField       =   "venta_fecha_inicio"
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
               Height          =   300
               Left            =   6360
               TabIndex        =   106
               Top             =   240
               Width           =   1485
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               DataField       =   "venta_fecha_fin"
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
               Height          =   300
               Left            =   8520
               TabIndex        =   105
               Top             =   240
               Width           =   1485
            End
            Begin VB.Label Txt_campo2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
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
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   120
               TabIndex        =   104
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Concepto"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   103
               Top             =   615
               Width           =   900
               WordWrap        =   -1  'True
            End
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "tw_tecnico_cronograma.frx":DFFF
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8385
            TabIndex        =   112
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "tw_tecnico_cronograma.frx":E018
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6180
            TabIndex        =   142
            Top             =   880
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "tw_tecnico_cronograma.frx":E031
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4260
            TabIndex        =   143
            Top             =   0
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
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "tw_tecnico_cronograma.frx":E04A
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1635
            TabIndex        =   144
            Top             =   300
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo Dtc_aux2 
            Bindings        =   "tw_tecnico_cronograma.frx":E063
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7920
            TabIndex        =   145
            Top             =   600
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   -2147483624
            ListField       =   "codigo2"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "tw_tecnico_cronograma.frx":E07C
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8200
            TabIndex        =   146
            Top             =   1140
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo2"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "tw_tecnico_cronograma.frx":E095
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4720
            TabIndex        =   147
            Top             =   880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "tw_tecnico_cronograma.frx":E0AE
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   180
            TabIndex        =   148
            Top             =   880
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   128
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6180
            TabIndex        =   163
            Top             =   675
            Width           =   525
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Correl. Cronograma"
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   9060
            TabIndex        =   162
            Top             =   75
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Venta"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   6915
            TabIndex        =   161
            Top             =   75
            Width           =   960
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Ejecutora Origen"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1665
            TabIndex        =   160
            Top             =   75
            Width           =   1740
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Trámite"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   159
            Top             =   75
            Width           =   885
         End
         Begin VB.Label txt_codigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   180
            TabIndex        =   158
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Edificio:"
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
            Height          =   195
            Left            =   180
            TabIndex        =   157
            Top             =   670
            Width           =   705
         End
         Begin VB.Label txt_codigo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3375
            TabIndex        =   156
            Top             =   2145
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Documento"
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
            Height          =   195
            Index           =   13
            Left            =   4635
            TabIndex        =   155
            Top             =   2160
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   1980
            TabIndex        =   154
            Top             =   2160
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label txt_campo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6000
            TabIndex        =   153
            Top             =   2145
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   8160
            TabIndex        =   152
            Top             =   75
            Width           =   660
         End
         Begin VB.Label lbl_Gestion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2016"
            DataField       =   "ges_gestion"
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
            Height          =   300
            Left            =   8160
            TabIndex        =   151
            Top             =   300
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "PARAMETROS PARA EL CRONOGRAMA DE INSTALACIONES"
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
            Height          =   195
            Index           =   3
            Left            =   2160
            TabIndex        =   150
            Top             =   3465
            Width           =   5355
         End
         Begin VB.Label lbl_venta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "36NO"
            DataField       =   "venta_codigo"
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
            Height          =   300
            Left            =   6840
            TabIndex        =   149
            Top             =   300
            Width           =   1125
         End
      End
      Begin VB.Frame FrmEdita 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4590
         Left            =   -74960
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   10695
         Begin VB.ComboBox txt_carta 
            BackColor       =   &H00FFFFFF&
            DataField       =   "carta"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "tw_tecnico_cronograma.frx":E0C7
            Left            =   9960
            List            =   "tw_tecnico_cronograma.frx":E0D1
            TabIndex        =   97
            Text            =   "NO"
            Top             =   1480
            Width           =   615
         End
         Begin VB.TextBox txt_certif 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cite_certificado"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5880
            TabIndex        =   96
            Text            =   "0"
            Top             =   1520
            Width           =   1575
         End
         Begin VB.TextBox txt_dias 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "bien_cantidad_por_empaque"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1080
            TabIndex        =   88
            Text            =   "0"
            Top             =   1520
            Width           =   735
         End
         Begin VB.ComboBox txt_carta2 
            BackColor       =   &H00FFFFFF&
            DataField       =   "estado_certificado"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "tw_tecnico_cronograma.frx":E0DD
            Left            =   8280
            List            =   "tw_tecnico_cronograma.frx":E0E7
            TabIndex        =   86
            Text            =   "REG"
            Top             =   1480
            Width           =   855
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "__________ Descripción del Insumo _______________________________ Código Insumo ______ Cant.X.Periodo _____"
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
            Height          =   2655
            Left            =   0
            TabIndex        =   58
            Top             =   1920
            Width           =   10575
            Begin VB.TextBox Txt_cant5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "cantidad5"
               DataSource      =   "ado_datos14"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8880
               TabIndex        =   79
               Text            =   "0"
               Top             =   2200
               Width           =   855
            End
            Begin VB.TextBox Txt_cant4 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "cantidad4"
               DataSource      =   "ado_datos14"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8880
               TabIndex        =   78
               Text            =   "0"
               Top             =   1720
               Width           =   855
            End
            Begin VB.TextBox Txt_cant3 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "cantidad3"
               DataSource      =   "ado_datos14"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8880
               TabIndex        =   77
               Text            =   "0"
               Top             =   1240
               Width           =   855
            End
            Begin VB.TextBox Txt_cant2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "cantidad2"
               DataSource      =   "ado_datos14"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8880
               TabIndex        =   76
               Text            =   "0"
               Top             =   760
               Width           =   855
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   8280
               TabIndex        =   75
               Top             =   2215
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   8280
               TabIndex        =   74
               Top             =   1735
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   8280
               TabIndex        =   73
               Top             =   1255
               Width           =   255
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   8280
               TabIndex        =   72
               Top             =   775
               Width           =   255
            End
            Begin VB.TextBox Txt_cant1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "cantidad1"
               DataSource      =   "ado_datos14"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8880
               TabIndex        =   67
               Text            =   "0"
               Top             =   280
               Width           =   855
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   8280
               TabIndex        =   64
               Top             =   295
               Width           =   255
            End
            Begin MSDataListLib.DataCombo dtc_codigo6 
               Bindings        =   "tw_tecnico_cronograma.frx":E0F5
               DataField       =   "bien_codigo1"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   6600
               TabIndex        =   66
               Top             =   280
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "bien_codigo"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo6A 
               Bindings        =   "tw_tecnico_cronograma.frx":E10E
               DataField       =   "bien_codigo2"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   6600
               TabIndex        =   68
               Top             =   760
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "bien_codigo"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo6B 
               Bindings        =   "tw_tecnico_cronograma.frx":E127
               DataField       =   "bien_codigo3"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   6600
               TabIndex        =   69
               Top             =   1240
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "bien_codigo"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo6C 
               Bindings        =   "tw_tecnico_cronograma.frx":E140
               DataField       =   "bien_codigo4"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   6600
               TabIndex        =   70
               Top             =   1720
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "bien_codigo"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo6D 
               Bindings        =   "tw_tecnico_cronograma.frx":E159
               DataField       =   "bien_codigo5"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   6600
               TabIndex        =   71
               Top             =   2200
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "bien_codigo"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc6 
               Bindings        =   "tw_tecnico_cronograma.frx":E172
               DataField       =   "bien_codigo1"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   1200
               TabIndex        =   65
               Top             =   285
               Width           =   5760
               _ExtentX        =   10160
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "bien_descripcion"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc6A 
               Bindings        =   "tw_tecnico_cronograma.frx":E18B
               DataField       =   "bien_codigo2"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   1200
               TabIndex        =   81
               Top             =   765
               Width           =   5760
               _ExtentX        =   10160
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "bien_descripcion"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc6B 
               Bindings        =   "tw_tecnico_cronograma.frx":E1A4
               DataField       =   "bien_codigo3"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   1200
               TabIndex        =   82
               Top             =   1245
               Width           =   5760
               _ExtentX        =   10160
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "bien_descripcion"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc6C 
               Bindings        =   "tw_tecnico_cronograma.frx":E1BD
               DataField       =   "bien_codigo4"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   1200
               TabIndex        =   83
               Top             =   1725
               Width           =   5760
               _ExtentX        =   10160
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "bien_descripcion"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc6D 
               Bindings        =   "tw_tecnico_cronograma.frx":E1D6
               DataField       =   "bien_codigo5"
               DataSource      =   "ado_datos14"
               Height          =   315
               Left            =   1200
               TabIndex        =   84
               Top             =   2205
               Width           =   5760
               _ExtentX        =   10160
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "bien_descripcion"
               BoundColumn     =   "bien_codigo"
               Text            =   "Todos"
            End
            Begin VB.Label lbl_insumo3 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Insumo 3"
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
               Height          =   195
               Left            =   360
               TabIndex        =   63
               Top             =   1260
               Width           =   780
            End
            Begin VB.Label lbl_insumo1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Insumo 1"
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
               Height          =   195
               Left            =   360
               TabIndex        =   62
               Top             =   300
               Width           =   780
            End
            Begin VB.Label lbl_insumo4 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Insumo 4"
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
               Height          =   195
               Left            =   360
               TabIndex        =   61
               Top             =   1740
               Width           =   780
            End
            Begin VB.Label lbl_insumo2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Insumo 2"
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
               Height          =   195
               Left            =   360
               TabIndex        =   60
               Top             =   780
               Width           =   780
            End
            Begin VB.Label lbl_insumo5 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Insumo 5"
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
               Height          =   195
               Left            =   360
               TabIndex        =   59
               Top             =   2220
               Width           =   780
            End
         End
         Begin VB.PictureBox FraGrabarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   0
            ScaleHeight     =   630
            ScaleWidth      =   10710
            TabIndex        =   36
            Top             =   0
            Width           =   10740
            Begin VB.PictureBox CmdCancelaDet 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   5160
               Picture         =   "tw_tecnico_cronograma.frx":E1EF
               ScaleHeight     =   615
               ScaleWidth      =   1395
               TabIndex        =   57
               Top             =   0
               Width           =   1400
            End
            Begin VB.PictureBox CmdGrabaDet 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   3720
               Picture         =   "tw_tecnico_cronograma.frx":EADB
               ScaleHeight     =   615
               ScaleWidth      =   1275
               TabIndex        =   56
               Top             =   0
               Width           =   1280
            End
            Begin VB.CommandButton cmdElige 
               BackColor       =   &H80000018&
               Caption         =   "New Prod"
               Height          =   525
               Left            =   7440
               MaskColor       =   &H00000000&
               Picture         =   "tw_tecnico_cronograma.frx":F2B1
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   120
               Visible         =   0   'False
               Width           =   825
            End
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   6915
            TabIndex        =   28
            Top             =   1055
            Width           =   270
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   8780
            TabIndex        =   26
            Top             =   1055
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   10080
            TabIndex        =   25
            Top             =   1055
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_subgrupo15 
            Bindings        =   "tw_tecnico_cronograma.frx":F6F3
            CausesValidation=   0   'False
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   5040
            TabIndex        =   22
            Top             =   720
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "subgrupo_codigo"
            BoundColumn     =   "bien_codigo"
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
         Begin MSDataListLib.DataCombo dtc_grupo15 
            Bindings        =   "tw_tecnico_cronograma.frx":F70D
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   4080
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "grupo_codigo"
            BoundColumn     =   "bien_codigo"
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
         Begin VB.TextBox TxtNroVenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "tec_plan_codigo"
            DataSource      =   "ado_datos14"
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
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1040
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dtc_desc15 
            Bindings        =   "tw_tecnico_cronograma.frx":F727
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   1035
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "concepto_venta"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_unimed15 
            Height          =   315
            Left            =   9060
            TabIndex        =   23
            Top             =   1035
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   ""
            BoundColumn     =   ""
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
         Begin MSDataListLib.DataCombo Dtc_partida15 
            Bindings        =   "tw_tecnico_cronograma.frx":F741
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6120
            TabIndex        =   27
            Top             =   720
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "par_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DTPFechaInicio 
            DataField       =   "fecha_inicio"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "ado_datos14"
            Height          =   285
            Left            =   4080
            TabIndex        =   39
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   114491393
            CurrentDate     =   41791
            MinDate         =   36526
         End
         Begin MSComCtl2.DTPicker DTPFechaFin 
            DataField       =   "fecha_fin"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "ado_datos14"
            Height          =   285
            Left            =   7800
            TabIndex        =   40
            Top             =   1320
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   114491393
            CurrentDate     =   41791
            MinDate         =   36526
         End
         Begin MSDataListLib.DataCombo dtc_codigo15 
            Bindings        =   "tw_tecnico_cronograma.frx":F75B
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7220
            TabIndex        =   41
            Top             =   1035
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTP_FechaDetif 
            DataField       =   "fecha_certificado"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "ado_datos14"
            Height          =   285
            Left            =   3200
            TabIndex        =   93
            Top             =   1515
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   114491393
            CurrentDate     =   43101
            MinDate         =   36526
         End
         Begin VB.Label Label18 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Envía Carta?"
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
            Height          =   435
            Left            =   9360
            TabIndex        =   98
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cite Entrega Definitiva"
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
            Height          =   435
            Left            =   4800
            TabIndex        =   95
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label Label13 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Entrega del Equipo"
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
            Height          =   435
            Left            =   1920
            TabIndex        =   94
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aprueba Entrega"
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
            Height          =   435
            Left            =   7560
            TabIndex        =   85
            Top             =   1425
            Width           =   840
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Horas X Día"
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
            Height          =   435
            Left            =   240
            TabIndex        =   38
            Top             =   1425
            Width           =   915
         End
         Begin VB.Label Label16 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Medida"
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
            Height          =   195
            Left            =   9075
            TabIndex        =   24
            Top             =   810
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Correlativo"
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
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   810
            Width           =   930
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Equipo"
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
            Height          =   195
            Left            =   1320
            TabIndex        =   19
            Top             =   810
            Width           =   1965
         End
         Begin VB.Label lbl_bien 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Código del Equipo"
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
            Height          =   195
            Left            =   7200
            TabIndex        =   18
            Top             =   810
            Width           =   1545
         End
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00C00000&
      Height          =   5077
      Left            =   0
      TabIndex        =   29
      Top             =   690
      Width           =   7425
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4920
         TabIndex        =   31
         Top             =   4680
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   30
         Top             =   4680
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4605
         Width           =   7200
         _ExtentX        =   12700
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
         BackColor       =   12632256
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "tw_tecnico_cronograma.frx":F775
         Height          =   4305
         Left            =   120
         TabIndex        =   164
         Top             =   240
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   7594
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "unidad_codigo"
            Caption         =   "U.E.Origen"
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
            DataField       =   "tec_plan_codigo"
            Caption         =   "Cronograma"
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
            DataField       =   "venta_codigo"
            Caption         =   "Contrato"
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
            DataField       =   "zpiloto_codigo"
            Caption         =   "Grupo.Piloto"
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
            DataField       =   "zona_edif_orden"
            Caption         =   "Orden"
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
            DataField       =   "fecha_inicio_tec"
            Caption         =   "Fech.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "fecha_fin_tec"
            Caption         =   "Fecha.Fin"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre_Edificio"
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
         BeginProperty Column10 
            DataField       =   "estado_detalle"
            Caption         =   "Detalle"
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
            DataField       =   "solicitud_codigo"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   675.213
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EQUIPOS QUE INTERVIENEN EN EL CRONOGRAMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   2160
      TabIndex        =   14
      Top             =   5760
      Width           =   15975
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "tw_tecnico_cronograma.frx":F78D
         Height          =   1140
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   2011
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "tec_plan_codigo"
            Caption         =   "Cronograma"
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
            Caption         =   "Cod.Equipo"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion y Características del Equipo"
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
            DataField       =   "bien_cantidad_por_empaque"
            Caption         =   "Hrs.X Día"
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
            DataField       =   "bien_codigo1"
            Caption         =   "Insumo1"
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
         BeginProperty Column05 
            DataField       =   "cantidad1"
            Caption         =   "Cantidad1"
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
            DataField       =   "bien_codigo2"
            Caption         =   "Insumo2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "cantidad2"
            Caption         =   "Cantidad2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "bien_codigo3"
            Caption         =   "Insumo3"
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
            DataField       =   "cantidad3"
            Caption         =   "Cantidad3"
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
            DataField       =   "bien_codigo4"
            Caption         =   "Insumo4"
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
            DataField       =   "cantidad4"
            Caption         =   "Cantidad4"
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
            DataField       =   "bien_codigo5"
            Caption         =   "Insumo5"
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
            DataField       =   "cantidad5"
            Caption         =   "Cantidad5"
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
            DataField       =   "estado_certificado"
            Caption         =   "Entregado"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column13 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column14 
               Alignment       =   2
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column15 
               Alignment       =   2
               ColumnWidth     =   615.118
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   2760
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6840
      Top             =   9480
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
      Top             =   9480
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
      Left            =   13920
      Top             =   9840
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
      Top             =   9840
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
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   13920
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
   Begin MSAdodcLib.Adodc ado_datos15 
      Height          =   330
      Left            =   6840
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2280
      Top             =   9840
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
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4560
      Top             =   9480
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
      Top             =   9480
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
   Begin MSAdodcLib.Adodc ado_datos6 
      Height          =   330
      Left            =   11520
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
      Caption         =   "ado_datos6"
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
   Begin Crystal.CrystalReport CryR01 
      Left            =   3360
      Top             =   10200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   11520
      Top             =   9840
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
      ConnectStringType=   3
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   9120
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
   Begin Crystal.CrystalReport Cry02 
      Left            =   0
      Top             =   10200
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
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lblUni_codigo 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "tw_tecnico_cronograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CRONOGRAMA
Dim rs_datos As New ADODB.Recordset     'CRONOGRAMA
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset   'Zonas Piloto
Dim rs_datos11 As New ADODB.Recordset
'Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Cronograma_detalle
Dim rs_datos15 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2, rs_aux3, rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset

Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset
'CLASIFICADORES
Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
'Dim queryinicial As String
Dim queryinicial2 As String
'Almacenes
Dim descri_bien, var_titulo As String
Dim Cant_Alm, VAR_CANT As Integer
Dim correlativo1 As Integer
'VARIABLES
Dim marca1 As Variant

Dim swgrabar, swnuevo, deta2 As Integer
Dim nroventa, VAR_MES, correldet2 As Integer
Dim UMED_NRO, VAR_TECCOD, VAR_AUX2, correldetalle As Integer
Dim VAR_COD0, VAR_ZONA, VAR_CONT, var_cod5, UMED_NRO2 As Integer
Dim DIA_MES, DIA_ORDEN As Integer
Dim VAR_DIAS As Integer
Dim VAR_DOCNUM, VAR_SOLTIPO As Integer

Dim Cobrobs, VAR_COBR As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double

Dim gestion0, var_literal, VAR_MED, VAR_CITE As String
Dim VAR_CODTIPO, VAR_BENEF, VAR_GLOSA, VAR_MONEDA As String
Dim VAR_UNITEC, VAR_COD2, VAR_COD3, VAR_COD4 As String
Dim VAR_FEC2, MControl, VAR_MESINI2, VAR_EMES As String
Dim VAR_EDIF, VAR_VALD, VAR_DA, VAR_DPTOC As String
Dim VAR_LUN, VAR_PRIM, VAR_UORIGEN As String

Dim VAR_PROC, VAR_SUB, VAR_ETAPA  As String
Dim VAR_CLASIF, VAR_DOC, VAR_POA  As String

Dim FInicio, FControl, FFin As Date
    
Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim descri_bien As String
Dim Cant_Alm As Integer
If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
   If Not IsNull(Ado_datos.Recordset("tec_plan_codigo")) Then
        If (Ado_datos.Recordset("estado_codigo") = "REG") Then
'            BtnAprobar.Visible = True
'            BtnDesAprobar.Visible = False
            BtnModificar.Visible = True
            BtnEliminar.Visible = True
'            BtnVer.Visible = False
'            FrmABMDet.Visible = True
        Else
'            BtnAprobar.Visible = False
'            BtnDesAprobar.Visible = True
            BtnModificar.Visible = False
            BtnEliminar.Visible = False
'            BtnVer.Visible = True
'            FrmABMDet.Visible = False
        End If
        If glusuario = "VPAREDES" Then
                BtnModificar.Visible = False
                BtnEliminar.Visible = False
                BtnAprobar.Visible = False
                BtnModDetalle.Visible = False
                BtnAddDetalle2.Visible = False
                BtnAnlDetalle2.Visible = False
                CmdGrabaDet.Visible = False
        End If
        'If Ado_datos.Recordset("beneficiario_codigo") <> "" And Ado_datos.Recordset("beneficiario_codigo") <> "VD" Then
        If Ado_datos.Recordset("beneficiario_codigo") <> "" Then
            Set RS_BENEF = New ADODB.Recordset
            If RS_BENEF.State = 1 Then RS_BENEF.Close
            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
            'RS_BENEF.Recordset.Requery
            If RS_BENEF.RecordCount > 0 Then
'                If RS_BENEF!beneficiario_deudor = "SI" Then
'                    Dtc_deudor2.backColor = &HFF&
'                Else
'                    Dtc_deudor2.backColor = &H80000010
'                End If
            End If
            
        End If
        Call ABRIR_DETALLE
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from tv_tecnico_cronograma_detalle_inst where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "'  AND tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & " ", db, adOpenKeyset, adLockOptimistic
        'rs_datos14.Open "select * from to_cronograma_detalle_inst where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "'  AND tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & " ", db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        'ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            FrmDetalle.Visible = True
            FrmCobranza.Visible = True
            Call ABRIR_TABLA_DET
        Else
            deta2 = 0
            'TxtMontoUs.Text = 0
            'Text2.Text = 0
'            FrmABMDet2.Visible = False
            FrmDetalle.Visible = False
            FrmCobranza.Visible = False
        End If

        FrmDetalle.Caption = "EQUIPOS del Cronograma Nro. " + Str((Ado_datos.Recordset("tec_plan_codigo")))
        
        End If
        FrmDetalle.Visible = True
        FrmCobranza.Visible = True
    Else
'        BtnAprobar.Visible = False
'            BtnDesAprobar.Visible = True
        BtnModificar.Visible = False
        BtnEliminar.Visible = False
'        BtnVer.Visible = True
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
'        FrmABMDet.Visible = False
    End If
End Sub

Private Sub ABRIR_DETALLE()
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from tv_tecnico_cronograma_detalle_inst where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "'  AND tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & " ", db, adOpenKeyset, adLockOptimistic
        'rs_datos14.Open "select * from to_cronograma_detalle_inst where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "'  AND tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & " ", db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        'ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            FrmDetalle.Visible = True
            FrmCobranza.Visible = True
            FrmCobranza.Caption = "Detalle del Cronograma Nro. " + Str((Ado_datos.Recordset("tec_plan_codigo")))
            Call ABRIR_TABLA_DET
        Else
            deta2 = 0
            'TxtMontoUs.Text = 0
            'Text2.Text = 0
'            FrmABMDet2.Visible = False
            FrmDetalle.Visible = False
            FrmCobranza.Visible = False
        End If
End Sub

Private Sub BtnAddDetalle_Click()
'  marca1 = Ado_datos.Recordset.Bookmark
''    Ado_datos.Recordset.Move marca1 - 1
'    swnuevo = 1
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'    FrmEdita.Visible = True
'    FrmEdita.Enabled = True
'    FraNavega.Enabled = False
'    FrmDetalle.Enabled = False
'    FrmDetalle.Visible = False
''    FrmABMDet.Visible = False
''    FrmABMDet2.Visible = False
End Sub

Private Sub BtnAddDetalle2_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
  VAR_VALD = "OK"
  Call valida_campos
  If VAR_VALD = "ERR" Then
      Exit Sub
  Else
        'WWWWW GENERA CRONOGRAMA DIARIO WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        FrmABMDet2.Enabled = False
        FrmABMDet.Enabled = False
        fraOpciones.Enabled = False
        'Screen.MousePointer = vbHourglass
                
        FInicio = Format(Ado_datos.Recordset!fecha_inicio_tec, "dd/mm/yyyy")
        FFin = Format(Ado_datos.Recordset!fecha_fin_tec, "dd/mm/yyyy")
        CANTOT = IIf(IsNull(Ado_datos.Recordset!tec_cantidad_unidades), 12, Ado_datos.Recordset!tec_cantidad_unidades)
        VAR_MED = IIf(IsNull(Ado_datos.Recordset!unimed_codigo), "MES", Ado_datos.Recordset!unimed_codigo)
        VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
        VAR_UNITEC = Ado_datos.Recordset!unidad_codigo_tec
        VAR_TECCOD = Ado_datos.Recordset!tec_plan_codigo
        VAR_EDIF = RTrim(dtc_desc3.Text)
        VAR_LUN = Ado_datos.Recordset!lunes_cambia
        VAR_PRIM = Ado_datos.Recordset!primero_mes
        
        VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
        dtc_codigo5.Text = "0"
        
      Set rs_datos6 = New ADODB.Recordset
      If rs_datos6.State = 1 Then rs_datos6.Close
      rs_datos6.Open "Select * from to_cronograma_inst WHERE estado_detalle = 'APR' AND unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenStatic
      If rs_datos6.RecordCount > 0 Then
           MsgBox "El Cronograma ya existe, verifique y vuelva a intentar ...", vbExclamation, "Validación de Registro"
           Frame2.Visible = False
           ProgressBar1.Visible = False
           Exit Sub
      Else
        ' estado_activo = 'ANL'
        DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
        MControl = Ado_datos.Recordset!mes_inicio_crono
        VAR_MESINI2 = Ado_datos.Recordset("mes_inicio_crono_nro")
        FControl = FInicio
        CONT4 = 0
        UMED_NRO = Ado_datos.Recordset("unimed_codigo_nro")     ' Fijo MES=1, BMES=2, TMES=3
        VAR_CONT = 1
        VAR_MES = Month(FControl)
        UMED_NRO2 = VAR_MESINI2      'UMED_NRO
        Frame2.Visible = True
        ProgressBar1.Visible = True
        With ProgressBar1
            .Max = CANTOT     'rs_datos6.RecordCount
            .Min = 0
            .Value = 0
        End With
        gestion0 = Year(FControl)
        While CANTOT >= VAR_CONT And FFin >= FControl   'UNIMED veces (12, 24, etc.)
            If UMED_NRO2 = 13 And gestion0 <> Year(FControl) Then
                UMED_NRO2 = 1
                'gestion0 = Year(FControl)
             End If
          gestion0 = Year(FControl)
          
          CONT3 = 0
          If VAR_MES = UMED_NRO2 Then
             Set rs_aux1 = New ADODB.Recordset
             rs_aux1.Open "select * from to_cronograma_detalle_inst where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux1.RecordCount > 0 Then
                 ' De acuerdo a la cantidad de equipos
                 'var_cod5 = IIf(IsNull(rs_aux1!bien_cantidad_por_empaque), 2, rs_aux1!bien_cantidad_por_empaque) / 2
                 var_cod5 = IIf(IsNull(rs_aux1!bien_cantidad_por_empaque), 2, rs_aux1!bien_cantidad_por_empaque)
                 rs_aux1.MoveFirst
                 While Not rs_aux1.EOF
                     
                     Set rs_aux2 = New ADODB.Recordset
                     If rs_aux2.State = 1 Then rs_aux2.Close
                     'rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "  and unidad_codigo_tec = '" & VAR_UNITEC & "'   ", db, adOpenKeyset, adLockOptimistic
                     rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "    ", db, adOpenKeyset, adLockOptimistic
                     If rs_aux2.RecordCount > 0 Then
                         VAR_AUX2 = rs_aux2!fmes_plan
                         VAR_COD0 = 0
                         'UMED_NRO2 = 0
                         Set rs_aux3 = New ADODB.Recordset
                         If rs_aux3.State = 1 Then rs_aux3.Close
                         rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
                         If rs_aux3.RecordCount > 0 Then
                             rs_aux3.MoveFirst
                             While Not rs_aux3.EOF
                                If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
                                   Set rs_aux4 = New ADODB.Recordset
                                   If rs_aux4.State = 1 Then rs_aux4.Close
                                   rs_aux4.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & "  AND estado_activo = 'REG'  ", db, adOpenKeyset, adLockBatchOptimistic
                                   If rs_aux4.RecordCount > 0 Then
                                    If VAR_COD0 < var_cod5 And rs_aux3!estado_activo = "REG" Then
                                       db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                       db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
                                       db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                       db.Execute "update to_cronograma_diario set nro_total_horas = " & var_cod5 & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                       'VAR_COD0 = VAR_COD0 + 1
                                       VAR_COD0 = VAR_COD0 + var_cod5
                                       CONT3 = 1
                                       'If VAR_MES Then
                                       VAR_EMES = "NADA"
                                       'End If
'                                       If VAR_LUN = "SI" Or VAR_PRIM = "SI" Then
'                                          'TODOS LOS LUNES O EL 1RO. DE CADA MES
'                                          If (rs_aux3!dia_nombre = "LUNES" Or rs_aux3!dia_correl = "1") And rs_aux3!hora_ingreso = "08:00" Then
'                                             rs_aux3.MoveNext
'                                             db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                             db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                             db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                             VAR_COD0 = VAR_COD0 + 1
'                                             CONT3 = 1
'                                          End If
'                                       End If
                                       db.Execute "Update to_cronograma_inst Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
                                    End If
                                   Else
                                        MsgBox "Ya no existen horarios laborales LIBRES, para la gestion: " & gestion0 & ", el Mes: " & VAR_MES & " y la Zona: " & VAR_ZONA, vbInformation, "Información"
                                        rs_aux3.MoveLast
                                   End If
                                Else
                                   If cmb_dia.Text = rs_aux3!dia_nombre And dtc_codigo5.Text = "0" Then
    '                                         If rs_aux3!dia_nombre = "SÁBADO" Or rs_aux3!dia_nombre = "DOMINGO" Or rs_aux3!estado_activo = "ANL" Then
    '                                            db.Execute "update to_cronograma_diario set observaciones = 'DIA NO LABORABLE' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
    '                                            db.Execute "update to_cronograma_diario set estado_activo = 'ANL' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
    '                                         Else
                                     If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
                                         db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                         db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
                                         db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                         VAR_COD0 = VAR_COD0 + 1
                                         CONT3 = 1
                                         db.Execute "Update to_cronograma_inst Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
                                     End If
                                   End If
                                   If dtc_codigo5.Text = rs_aux3!horario_codigo Then
                                     If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
                                         db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                         db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
                                         db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                         VAR_COD0 = VAR_COD0 + 1
                                         CONT3 = 1
                                         db.Execute "Update to_cronograma_inst Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
                                     End If
                                   End If
                                End If
                                rs_aux3.MoveNext
                                'db.Execute "Update to_cronograma_inst Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
                             Wend
                         End If
                     End If
                     rs_aux1.MoveNext
                 Wend
             End If
             VAR_CONT = VAR_CONT + 1
             UMED_NRO2 = UMED_NRO2 + UMED_NRO
             ProgressBar1.Value = ProgressBar1.Value + 1
          Else
            'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique u vuelva a intentar..."
          End If
             Select Case VAR_MES
                 Case 2
                     If gestion0 = "2016" Or gestion0 = "2020" Or gestion0 = "2024" Or gestion0 = "2028" Then
                         Dias_Mes = 29
                     Else
                         Dias_Mes = 28
                     End If
                 Case 1, 3, 5, 7, 8, 10, 12
                     Dias_Mes = 31
                 Case 4, 6, 9, 11
                     Dias_Mes = 30
             End Select
             'rs_aux2!cobranza_fecha_prog = FControl
             'rs_aux2!cobranza_fecha_cobro = FControl + 10
             FControl = CDate(FControl) + Dias_Mes
             VAR_MES = Month(FControl)
             Select Case VAR_MED
                Case "MES"    'MENSUAL
'                    UMED_NRO2 = VAR_CONT
'                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
'                        'UMED_NRO2 = (VAR_MES * UMED_NRO) - 1
'                        UMED_NRO2 = VAR_CONT
'                    Else
'                        UMED_NRO2 = VAR_MES * UMED_NRO
'                    End If
                Case "BMES"    'BIMESTRAL
'                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
'                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
'                    Else
'                        UMED_NRO2 = VAR_CONT * UMED_NRO
'                    End If
                Case "TMES"    'TRIMESTRAL
                    'UMED_NRO2 = (UMED_NRO2 + UMED_NRO)
'                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
'                        UMED_NRO2 = (VAR_CONT * UMED_NRO) '- 2
'                        'UMED_NRO2 = (UMED_NRO2 + VAR_MESINI2)
'                    Else
'                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
'                    End If
'                    'UMED_NRO2 = 3
'                    If VAR_MES = UMED_NRO2 Then
'                        UMED_NRO2 = UMED_NRO2 + VAR_MESINI2
'                    End If
'                    UMED_NRO2 = VAR_CONT * UMED_NRO
                Case "CMES"    'CUATRIMESTRAL
                Case "QMES"    'CADA 5 MESES
                Case "SMES"    'SEMESTRAL
                    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
                Case "ANUAL"    'ANUAL
             End Select
'             If VAR_MED = "TMES" And CONT3 = 1 Then
'                UMED_NRO2 = (VAR_CONT * UMED_NRO) - 2
'                VAR_CONT = VAR_CONT + 1
'             End If
'                If CONT3 = 1 And VAR_MED = "MES" And (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) And (UMED_NRO = 2) Then
'                    UMED_NRO2 = (VAR_MES * UMED_NRO) - 1
'                Else
'                    UMED_NRO2 = VAR_MES * UMED_NRO
'                End If
'                'If CONT3 = 1 And VAR_MED = "BMES" Then
'                If VAR_MED = "BMES" Then
'                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) And (UMED_NRO = 2) Then
'                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
'                        VAR_CONT = VAR_CONT + 1
'                    Else
'                        UMED_NRO2 = VAR_CONT * UMED_NRO
'                        VAR_CONT = VAR_CONT + 1
'                    End If
'                End If
'             'End If
        Wend
        
        FrmABMDet2.Enabled = True
        FrmABMDet.Enabled = True
        fraOpciones.Enabled = True
'        Screen.MousePointer = vbDefault
        If VAR_EMES = "NADA" Then
            MsgBox "El Cronograma fue creado Satisfactoriamente ...", vbInformation, "Información"
            ProgressBar1.Visible = False
            Frame2.Visible = False
        Else
            MsgBox VAR_EMES, vbInformation, "Información"
        End If
        Call ABRIR_DETALLE
      End If
      ProgressBar1.Visible = False
      Frame2.Visible = False
      'WWWWW GENERA CRONOGRAMA DIARIO (FIN)
  End If
 Else
        MsgBox "NO se puede generar un NUEVO CRONOGRAMA, en un Registro APROBADO o ANULADO !! ", vbExclamation, "Atención!"
 End If

End Sub

Private Sub BtnAnlDetalle2_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
  Set rs_datos6 = New ADODB.Recordset
  If rs_datos6.State = 1 Then rs_datos6.Close
  rs_datos6.Open "Select * from to_cronograma_inst WHERE estado_detalle = 'APR' AND unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "   ", db, adOpenStatic
  If rs_datos6.RecordCount > 0 Then
     ProgressBar1.Visible = True
     With ProgressBar1
        .Max = rs_datos6.RecordCount
        .Min = 0
        .Value = 0
     End With
     ProgressBar1.Value = ProgressBar1.Value + 1
       db.Execute "Update to_cronograma_inst Set estado_detalle = 'REG' Where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "   "
       db.Execute "update to_cronograma_diario set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = '', observaciones = 'HORARIO LABORABLE', edif_descripcion = '', estado_activo = 'REG', estado_codigo = 'REG' Where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "   "     'WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & "  "
       MsgBox "El Cronograma fue deshabilitado exitosamente ...", vbExclamation, "Validación de Registro"
       Call ABRIR_DETALLE
       ProgressBar1.Visible = False
  Else
        ProgressBar1.Visible = False
        MsgBox "El Cronograma ya fue deshabilitado, verifique y vuelva a generarlo (Nuevo) ...", vbExclamation, "Validación de Registro"
  End If
 Else
        MsgBox "NO se puede ANULAR EL CRONOGRAMA, en un Registro APROBADO o ANULADO !! ", vbExclamation, "Atención!"
 End If
End Sub

Private Sub BtnAñadir_Click()
'    FrmCabecera.Enabled = True
'    FrmDetalle.Visible = False
'    FrmCobranza.Visible = False
'    fraOpciones.Visible = False
'    FraNavega.Enabled = False
'    FraGrabarCancelar.Visible = True
'    Fra_datos.Enabled = True
'    Fra_Total.Visible = False
'    swgrabar = 1
'    Cmd_Cliente.Visible = True
'    'FrmABMDet.Enabled = False
'    'FrmABMDet2.Enabled = False
'    Call cerea
'    Dim rstdestino As New ADODB.Recordset
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_ventas_cabecera where estado_codigo <> 'E'", db, adOpenDynamic, adLockOptimistic
'    Set Ado_datos.Recordset = rstdestino
'    Ado_datos.Recordset.AddNew
'    DTPfechasol.Value = Date
'    DTPfechasol.CheckBox = True
'    dtc_codigo11.Text = "E"
'    dtc_desc11.Text = "EFECTIVO (Contado)"
'
'    Set rs_personalUsr = New ADODB.Recordset
'    If rs_personalUsr.State = 1 Then rs_personalUsr.Close
'    'rs_personalUsr.Open "select * from fv_beneficiarioUsr WHERE usuario='" & GlUsuario & "' OR usuario='ADMIN' ORDER BY denominacion_beneficiario ", db, adOpenKeyset, adLockReadOnly
'    rs_personalUsr.Open "select * from fv_beneficiarioUsr WHERE tipoben_codigo=1 or tipoben_codigo=6 or tipoben_codigo=7 ORDER BY denominacion_beneficiario ", db, adOpenKeyset, adLockReadOnly
'    'WHERE tipoben_codigo=1 or tipoben_codigo=6 or tipoben_codigo=7
'    'Set adopuestosol.Recordset = rs_personalUsr
'    'adopuestosol.Refresh
'    If rs_personalUsr.RecordCount > 0 Then
'        Dtcpaternosol.Text = rs_personalUsr("denominacion_beneficiario")
''    dtcmaternosol.Text = rstrc_personalSoli!segundo_apellido
''    dtcnombresol.Text = rstrc_personalSoli!NombreS
'        dtc_codigo4.Text = rs_personalUsr("beneficiario_codigo")
'    Else
'        Dtcpaternosol.Text = "ADMIN"
'        dtc_codigo4.Text = "-"
'    End If
'    dtc_codigo2.Text = "VD"
'    dtc_desc2.Text = "VENTA DIRECTA"
'    TxtConcepto.Text = "VENTA DIRECTA AL CLIENTE"
'    TxtPlazo.Visible = False
'    Label7.Visible = True
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False

End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  VAR_UNITEC = Ado_datos.Recordset!unidad_codigo_tec
  VAR_TECCOD = Ado_datos.Recordset!tec_plan_codigo
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
   If rs_datos!estado_codigo = "REG" And rs_datos!estado_detalle = "APR" Then       'And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        db.Execute "Update to_cronograma_inst Set estado_codigo = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
        db.Execute "Update to_cronograma_detalle_inst Set estado_codigo = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "  "
        
'        Select Case Ado_datos.Recordset!subproceso_codigo  'dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "TEC-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
'
'            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Case "TEC-05"    '5. SERVICIO MODERNIZACION
'            Case "TEC-02"    '10. SERVICIO MANTENIMIENTO
'               db.Execute "Update to_cronograma_inst Set estado_codigo = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'               db.Execute "Update to_cronograma_detalle_inst Set estado_codigo = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "  "
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'        End Select
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
   End If
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

'Private Sub BtnAprobar2_Click()
' If IsNull(Ado_datos16.Recordset("cobranza_observaciones")) Or (Ado_datos16.Recordset("cobranza_programada_bs") = 0) Then
'    MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'    Exit Sub
' Else
'    If Ado_datos16.Recordset("estado_codigo") = "REG" Then
'       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'       If sino = vbYes Then
'            db.Execute "update gc_documentos_respaldo set gc_documentos_respaldo.correl_doc = " & Ado_datos.Recordset!venta_codigo & " Where gc_documentos_respaldo.doc_codigo = '" & Ado_datos16.Recordset!doc_codigo & "' "
'
'            db.Execute "INSERT INTO ao_ventas_cobranza (ges_gestion, cobranza_prog_codigo, venta_codigo, beneficiario_codigo, beneficiario_codigo_resp, cobranza_programada_bs, cobranza_programada_dol, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, Literal, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, poa_codigo, estado_codigo, usr_codigo, fecha_registro) " & _
'            "VALUES ('" & Ado_datos16.Recordset!ges_gestion & "', " & Ado_datos16.Recordset!cobranza_prog_codigo & ", " & Ado_datos16.Recordset!venta_codigo & ", '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo_resp & "', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '0', '0', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '" & Ado_datos16.Recordset!Literal & "', '" & Ado_datos16.Recordset!cobranza_fecha_prog & "', '" & Ado_datos16.Recordset!cobranza_fecha_prog & "', '" & Ado_datos16.Recordset!cobranza_observaciones & "', 'COM', 'COM-02', 'COM-02-04', 'ADM', 'R-105', '0', '', '0', '0', '3.1.2', 'REG', '" & GlUsuario & "', '" & Date & "')"
'
'            ' APRUEBA ao_ventas_cobranza_prog
'            db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' And venta_codigo = " & Ado_datos.Recordset!venta_codigo & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
'            Ado_datos16.Refresh
'       End If
'    End If
' End If
'End Sub

Private Sub BtnBuscar_Click()
'JQA
  If (Ado_datos.Recordset.RecordCount > 0) Then
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnCancelar_Click()
  'Ado_datos.Refresh
    FrmDetalle.Enabled = True
    FrmCobranza.Enabled = True
    FraNavega.Enabled = True
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FrmCabecera.Enabled = False
'    Fra_datos.Enabled = True
'    Fra_Total.Enabled = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
        
    marca1 = Ado_datos.Recordset.Bookmark
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
      Call OptFilGral1_Click
    Else
      Call OptFilGral2_Click
    End If
  
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = False
'  SSTab1.TabEnabled(2) = True
  'Ado_datos.Recordset.Move marca1 - 1
End Sub

Private Sub BtnEliminar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_detalle = "REG" Then
      sino = MsgBox("Esta seguro de ANULAR el Registro ?", vbYesNo, "Confirmando")
      If sino = vbYes Then
          db.Execute "update to_cronograma_inst set estado_codigo = 'ANL' Where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' And tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "  "
          marca1 = Ado_datos.Recordset.Bookmark
          Call OptFilGral1_Click
          Ado_datos.Recordset.Move marca1 - 1
      End If
    Else
      MsgBox "NO se puede ANULAR el registro que ya tiene Cronograma generado ...", , "Atencion"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnGrabar_Click()
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    correlativo1 = Ado_datos.Recordset!correlativo
    FrmCabecera.Enabled = False
    Call grabar
    If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
     'VAR_SW = ""
        rs_datos.Find "correlativo = " & correlativo1 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
     'VAR_SW = ""
        rs_datos.MoveLast
     End If
    'dg_datos.Visible = True
    FrmDetalle.Enabled = True
    FrmCobranza.Enabled = True
    FraNavega.Enabled = True
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FrmCabecera.Enabled = False
'    Fra_datos.Enabled = True
'    Fra_Total.Enabled = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
        
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = True
  End If

End Sub

Private Sub valida_campos()
  If DTPfechasol = "" Then
    MsgBox "Debe Elejir: " + lbl_fecha.Caption + " !! , Vuelva a Intentar por favor ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  If dtc_codigo4.Text = "" Then
    'MsgBox "Debe Elejir: " + lbl_resp.Caption + "!! , Vuelva a Intentar ...", vbExclamation, "Atención"
    MsgBox "Debe Elejir: Responsable CGI !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  '
  If dtc_codigo7.Text = "" Then
    MsgBox "Debe Elejir: Zona Piloto !! , Vuelva a Intentar por favor ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    VAR_VALD = "ERR"
    Exit Sub
  End If
'  If TxtConcepto = "" Then
'    MsgBox "Debe Elejir: " + lbl_concepto.Caption + "!! , Vuelva a Intentar ...", vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    VAR_VALD = "ERR"
'    Exit Sub
'  End If
  '
  If txt_cant = "" Then
    MsgBox "Debe Registrar: Total Periodos !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  If cmd_unimed2 = "" Then
    MsgBox "Debe Elejir: Periodicidad !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  'If Val(Format(lbl_fecha_fin.Caption, "dd/mm/yyyy")) <= Val(Format(lbl_fecha_ini.Caption, "dd/mm/yyyy")) Then
  If CDate(Format(lbl_fecha_fin, "dd/mm/yyyy")) <= CDate(Format(lbl_fecha_ini, "dd/mm/yyyy")) Then
    MsgBox "La Fecha de Inicio debe ser MENOR a la Fecha de Fin del Cronograma!! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  If cmb_mes_ini = "" Then
    MsgBox "Debe Elejir: Mes Inicio Cronograma !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  Select Case RTrim(cmb_mes_ini)
        Case "ENERO"
            VAR_MESINI2 = 1
        Case "FEBRERO"
            VAR_MESINI2 = 2
        Case "MARZO"
            VAR_MESINI2 = 3
        Case "ABRIL"
            VAR_MESINI2 = 4
        Case "MAYO"
            VAR_MESINI2 = 5
        Case "JUNIO"
            VAR_MESINI2 = 6
        Case "JULIO"
            VAR_MESINI2 = 7
        Case "AGOSTO"
            VAR_MESINI2 = 8
        Case "SEPTIEMBRE"
            VAR_MESINI2 = 9
        Case "OCTUBRE"
            VAR_MESINI2 = 10
        Case "NOVIEMBRE"
            VAR_MESINI2 = 11
        Case "DICIEMBRE"
            VAR_MESINI2 = 12
  End Select
  If Month(CDate(Format(lbl_fecha_ini, "dd/mm/yyyy"))) <> 12 And VAR_MESINI2 <> 1 Then
    If Val(VAR_MESINI2) < Month(CDate(Format(lbl_fecha_ini, "dd/mm/yyyy"))) Then
        MsgBox "El MES de Inicio del Crono. NO puede ser MENOR al de la Fecha de Inicio del Cronograma!! , Vuelva a Intentar ...", vbExclamation, "Atención"
        VAR_VAL = "ERR"
        VAR_VALD = "ERR"
        Exit Sub
    End If
  End If
'  If dtc_codigo11.Text = "C" And dtc_codigo2 = "VD" Then
'        MsgBox "NO se puede realizar la Venta a Credito, Debe cambiar de Cliente ..."
'  Else

End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If (ado_datos14.Recordset.RecordCount > 0) Then
      Dim iResult As Variant  ', i%, y%
      CryV01.ReportFileName = App.Path & "\reportes\tecnico\tr_R-362_instalaciones_seguimiento.rpt"
      CryV01.WindowShowRefreshBtn = True
      'CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion           '"%"
      CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo_tec          '"%"
      'CryR01.StoredProcParam(2) = "0"         'Me.Ado_datos.Recordset!tec_plan_codigo
      'Literal por el Total de la Compra
      var_literal = "TODOS"    'Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
      '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
      CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
      'CryR01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
      CryV01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!tec_plan_codigo & "' "
      Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      CryV01.Formulas(3) = "titulo = '" & var_titulo & "' "
      CryV01.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "
      iResult = CryV01.PrintReport
      If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
    Else
      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub bTNiMPRIMIR0_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If (ado_datos14.Recordset.RecordCount > 0) Then
      Dim iResult As Variant  ', i%, y%
      CryV01.ReportFileName = App.Path & "\reportes\tecnico\tr_R-362_instalaciones_seguimiento_PILOTO.rpt"
      CryV01.WindowShowRefreshBtn = True
      'CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion           '"%"
      CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo_tec          '"%"
      'CryR01.StoredProcParam(2) = "0"         'Me.Ado_datos.Recordset!tec_plan_codigo
      'Literal por el Total de la Compra
      var_literal = "TODOS"    'Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
      '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
      CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
      'CryR01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
      CryV01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!tec_plan_codigo & "' "
      Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      CryV01.Formulas(3) = "titulo = '" & var_titulo & "' "
      CryV01.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "
      iResult = CryV01.PrintReport
      If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
    Else
      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If

End Sub

Private Sub BtnImprimir3_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
      Dim iResult As Variant  ', i%, y%
      Cry02.ReportFileName = App.Path & "\reportes\tecnico\tr_zonas_vs_edificios.rpt"
      Cry02.WindowShowRefreshBtn = True
      Cry02.StoredProcParam(0) = "%"       'Me.Ado_datos.Recordset!unidad_codigo_tec
      'Literal por el Total de la Compra
      var_literal = "TODOS"    'Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
      '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
      Cry02.Formulas(1) = "literalcobro = '" & var_literal & "' "
      Cry02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!zpiloto_codigo & "' "
      Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      Cry02.Formulas(3) = "titulo = '" & var_titulo & "' "
      Cry02.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "
      iResult = Cry02.PrintReport
      If iResult <> 0 Then MsgBox Cry02.LastErrorNumber & " : " & Cry02.LastErrorString, vbCritical, "Error de impresión"
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnModificar_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    If Ado_datos.Recordset.RecordCount > 0 Then
        FrmDetalle.Enabled = False
        FrmCobranza.Enabled = False
        FraNavega.Enabled = False
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        FrmCabecera.Enabled = True
'        Fra_datos.Enabled = True
'        Fra_Total.Enabled = True
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
        DTPfechasol.SetFocus
        swgrabar = 0
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        'update to_cronograma_mensual set to_cronograma_mensual.beneficiario_codigo_resp = tc_zonas_piloto_inst.beneficiario_codigo
'from to_cronograma_mensual inner join tc_zonas_piloto_inst on to_cronograma_mensual.zpiloto_codigo  = tc_zonas_piloto_inst.zpiloto_codigo
'        cmb_lunes
    Else
        MsgBox "NO se puede MODIFICAR, verifique si existen registros y vuelva a intentar !! ", vbExclamation, "Atención!"
    End If
  Else
        MsgBox "NO se puede MODIFICAR, un Registro APROBADO o ANULADO !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

Private Sub CmdCancelaDet_Click()
  'TxtNroVenta.Enabled = True
  FrmEdita.Enabled = False
  swgrabar = 0
  'Call cerea
  swnuevo = 0
  'cmdElige.Enabled = False
  marca1 = Ado_datos.Recordset.Bookmark
  
'  If Ado_datos.Recordset("estado_codigo") = "REG" Then
'    Call OptFilGral1_Click
'  Else
'    Call OptFilGral2_Click
'  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False

    FraNavega.Enabled = True
    FrmDetalle.Enabled = True
    FrmDetalle.Visible = True
    FrmEdita.Enabled = False
    fraOpciones.Enabled = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
    Call ABRIR_DETALLE
  'Ado_datos.Recordset.Move marca1 - 1
End Sub

'Private Sub BtnAnlDetalle2_Click()
' If Ado_datos.Recordset!estado_codigo = "REG" Then
'   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
'   If sino = vbYes Then
'      db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.estado_codigo = 'ANL' Where ao_ventas_cobranza_prog.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & Ado_datos16.Recordset("cobranza_codigo") & " "
'      'db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_deuda_bs = '0', ao_ventas_cobranza_prog.cobranza_deuda_dol = '0'  Where ao_ventas_cobranza_prog.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & ado_datos16.Recordset("cobranza_codigo") & " "
'
'     'ado_ventas_COBRANZAS.Recordset.Delete
'     'ado_ventas_COBRANZAS.Recordset.Update
'     'ado_ventas_COBRANZAS.Requery
'     'ado_ventas_COBRANZAS.Refresh
'     ''cerea
'     'ado_ventas_COBRANZAS.Refresh
'   End If
'  Else
'    MsgBox "Los productos del registro sin Aprobar, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
'  End If
'End Sub

'Private Sub BtnModDetalle2_Click()
'  'If Ado_datos.Recordset!venta_tipo <> "E" And Ado_datos16.Recordset!estado_codigo = "REG" Then
'  If Ado_datos16.Recordset!estado_codigo = "REG" And (Ado_datos.Recordset!venta_tipo = "E" Or Ado_datos.Recordset!venta_tipo = "V" Or Ado_datos.Recordset!venta_tipo = "C") Then
'    FraNavega.Enabled = False
'    fraOpciones.Enabled = False
'    FrmDetalle.Visible = False
'    FrmCobranza.Visible = False
'    'swgrabar = 0
'    swnuevo = 2
'    'TxtMonto.SetFocus
'    'TxtNroVenta.Enabled = False
'    'marca1 = ado_datos14.Recordset.BookMark
'    'txt_descripcion_venta.Enabled = True
'    'TxtNroVenta.Text = txt_venta.Text
'    'lbltipoVenta.Caption = dtc_desc11.Text
'    'lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
'    SSTab1.Tab = 2
'    SSTab1.TabEnabled(2) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(1) = False
'    FrmCobros.Visible = True
'    FrmCobros.Enabled = True
'    FrmABMDet.Visible = False
'    FrmABMDet2.Visible = False
'    'If Ado_datos.Recordset!estado_codigo = "APR" Then
'        'sino = MsgBox("Registrará la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
'        'If sino = vbYes Then
'        '    DTPFechaProg.Visible = False
'        '    DTPFechaCobro.Visible = True
'        '    Lbl_nombre_fac.Caption = "Factura a Nombre de:"
'        '    lbl_fechas.Caption = "Fecha Efectiva de Cobranza"
'        '    Txt_parche.Visible = False      '&H80000013&
'        '    'dtc_desc2A.BackColor = &H80000013
'        'Else
'        '    DTPFechaProg.Visible = True
'        '    DTPFechaCobro.Visible = False
'        '    Lbl_nombre_fac.Caption = "Cliente :"
'        '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
'        '    Txt_parche.Visible = True       '&H80000005&
'        '    'dtc_desc2A.BackColor = &H80000005
'        'End If
'    'Else
'        DTPFechaProg.Visible = True
'        DTPFechaCobro.Visible = False
'        Lbl_nombre_fac.Caption = "Cliente :"
'        lbl_fechas.Caption = "Fecha Programada de Cobranza"
'        Txt_parche.Visible = True       '&H80000005&
'        'dtc_desc2A.BackColor = &H80000005
'    'End If
'    VAR_MBS2 = Ado_datos16.Recordset!cobranza_programada_bs
'    TxtMonto.SetFocus
'  Else
'    MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
'  End If
'End Sub

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

Private Sub cmdElige_Click()
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
End Sub

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

Private Sub CmdGrabaDet_Click()
  If dtc_codigo6 = "" Then
        MsgBox "Debe Elejir: " + lbl_insumo1.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
  End If
  If dtc_codigo6A = "" Then
     MsgBox "Debe Elejir : " + lbl_insumo2.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If dtc_codigo6B = "" Then
    MsgBox "Debe Registrar: " + lbl_insumo3.Caption + ", !! en el Proyecto de Edificación, Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  VAR_FECHA1 = Format(Ado_datos.Recordset!fecha_inicio_tec, "dd/mm/yyyy")
  VAR_FECHA2 = Format(Ado_datos.Recordset!fecha_fin_tec, "dd/mm/yyyy")
'  If (DTPFechaInicio.Value < CDate(VAR_FECHA1)) Or (DTPFechaInicio.Value > CDate(VAR_FECHA2)) Then
'    MsgBox "La -" + lbl_fechai.Caption + "- debe ser mayor o igual a la -" + lbl_fechaini.Caption + "- o menor o igual a la -" + lbl_fechafin.Caption + "- de los Datos Acumulados del Cronograma, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
'  If (DTPFechaFin.Value < DTPFechaInicio.Value) Or (DTPFechaFin.Value > CDate(VAR_FECHA2)) Then
'    MsgBox "La -" + lbl_fechaf.Caption + "- debe ser mayor o igual a la -" + lbl_fechai.Caption + "- o menor o igual a la -" + lbl_fechafin.Caption + "- , !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If

        If swnuevo = 1 Then
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            rs_aux2.Open "select * from to_cronograma_detalle_inst where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' AND tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & " AND bien_codigo = '" & VAR_COD4 & "' AND ges_gestion = '" & lbl_Gestion.Caption & "' ", db, adOpenKeyset, adLockOptimistic
            If rs_aux2.RecordCount > 0 Then
               MsgBox "El equipo ya fue asignado..., Intente nuevamente !!"
               Exit Sub
            End If
            ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset("ges_gestion")
            ado_datos14.Recordset!unidad_codigo_tec = Ado_datos.Recordset("unidad_codigo_tec")    'Trim(parametro)
            ado_datos14.Recordset!bien_codigo = IIf(dtc_codigo15.Text = "", VAR_COD4, dtc_codigo15.Text)                  'Codigo Bien (Equipo, Producto, etc)
            ado_datos14.Recordset!tec_plan_codigo = Ado_datos.Recordset("tec_plan_codigo")      'correldetalle
        End If
            VAR_DIAS = Val(VAR_FECHA2) - Val(VAR_FECHA1)
            
            db.Execute "update to_cronograma_detalle_inst set bien_cantidad_por_empaque= " & IIf(txt_dias.Text = "", 2, txt_dias.Text) & ", bien_tiempo_dias = " & VAR_DIAS & ", carta= '" & txt_carta.Text & "', " & _
            " bien_codigo1 = '" & dtc_codigo6.Text & "', bien_codigo2 = '" & dtc_codigo6A.Text & "', bien_codigo3 = '" & dtc_codigo6B.Text & "', bien_codigo4 = '" & dtc_codigo6C.Text & "', bien_codigo5 = '" & dtc_codigo6D.Text & "', fecha_inicio = '" & VAR_FECHA1 & "', fecha_fin = '" & VAR_FECHA2 & "',  " & _
            " cantidad1 = " & IIf(Txt_cant1.Text = "", 0, Txt_cant1.Text) & ", cantidad2 = " & IIf(Txt_cant2.Text = "", 0, Txt_cant2.Text) & ", cantidad3 = " & IIf(Txt_cant3.Text = "", 0, Txt_cant3.Text) & " , cantidad4 = " & IIf(Txt_cant4.Text = "", 0, Txt_cant4.Text) & " , cantidad5 = " & IIf(Txt_cant5.Text = "", 0, Txt_cant5.Text) & "  " & _
            " where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & " aND bien_codigo = '" & ado_datos14.Recordset!bien_codigo & "' "
            
            'fecha_certificado = '" & IIf(DTP_FechaDet.Value = "", Date, DTP_FechaDet.Value) & "', cite_certificado = '" & IIf(txt_certif = "", "", txt_certif) & "', estado_certificado = '" & IIf(txt_carta2.Text = "", "REG", txt_carta2.Text) & "'
            
'            'beneficiario_codigo,munic_codigo,
'            '          fecha_inicio, fecha_fin, hora_inicio, hora_fin, bien_tiempo_dias,
'            '          , cantidad2, cantidad3, cantidad4, cantidad5, carta, estado_codigo, usr_codigo, fecha_registro, hora_registro
'
'          'ado_datos14.Recordset!beneficiario_codigo = Trim(dtc_codigo4A.Text)               'Técnico Asignado
'          ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
'          ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
'          ado_datos14.Recordset!par_codigo = Trim(Dtc_partida15)                            'Partida
''          ado_datos14.Recordset!munic_codigo = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Municipio
'          'ado_datos14.Recordset!fecha_inicio = VAR_FECHA1   'DTPFechaInicio.Value                         'Fecha de Inicio de la Tarea
'          'ado_datos14.Recordset!fecha_fin = VAR_FECHA2      'DTPFechaFin.Value                               'Fecha de Finalización de la Tarea
'          'ado_datos14.Recordset!hora_inicio = ""                            'Hora de Inicio de la Tarea
'          'ado_datos14.Recordset!hora_fin = ""                               'Hora de Fin de la Tarea
'          VAR_DIAS = Val(VAR_FECHA2) - Val(VAR_FECHA1)
'          ado_datos14.Recordset!bien_tiempo_dias = VAR_DIAS                       ' Tiempo Estimado en Días
'          ado_datos14.Recordset!bien_cantidad_por_empaque = IIf(txt_dias.Text <> "2" And txt_dias.Text <> "4", "2", txt_dias.Text)
'          ado_datos14.Recordset!bien_codigo1 = IIf(IsNull(dtc_codigo6.Text), "", dtc_codigo6.Text)
'          ado_datos14.Recordset!bien_codigo2 = IIf(IsNull(dtc_codigo6A.Text), "", dtc_codigo6A.Text)
'          ado_datos14.Recordset!bien_codigo3 = IIf(IsNull(dtc_codigo6B.Text), "", dtc_codigo6B.Text)
'          ado_datos14.Recordset!bien_codigo4 = IIf(IsNull(dtc_codigo6C.Text), "", dtc_codigo6C.Text)
'          ado_datos14.Recordset!bien_codigo5 = IIf(IsNull(dtc_codigo6D.Text), "", dtc_codigo6D.Text)
'          ado_datos14.Recordset!cantidad1 = IIf(Txt_cant1.Text = "", "0", Txt_cant1.Text)
'          ado_datos14.Recordset!cantidad2 = IIf(Txt_cant2.Text = "", "0", Txt_cant2.Text)
'          ado_datos14.Recordset!cantidad3 = IIf(Txt_cant3.Text = "", "0", Txt_cant3.Text)
'          ado_datos14.Recordset!cantidad4 = IIf(Txt_cant4.Text = "", "0", Txt_cant4.Text)
'          ado_datos14.Recordset!cantidad5 = IIf(Txt_cant5.Text = "", "0", Txt_cant5.Text)
'          ado_datos14.Recordset!carta = IIf(txt_carta.Text = "", "NO", txt_carta.Text)
'
'          ado_datos14.Recordset!estado_codigo = "REG"
'          ado_datos14.Recordset!usr_codigo = glusuario
'          ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'          ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'          ado_datos14.Recordset.Update
'        'db.CommitTrans
        
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        'SSTab1.TabEnabled(2) = False
        FraNavega.Enabled = True
        FrmDetalle.Enabled = True
        FrmDetalle.Visible = True
        FrmEdita.Enabled = False
        fraOpciones.Enabled = True
        FrmABMDet.Visible = True
        FrmABMDet2.Visible = True
        Call ABRIR_DETALLE
'        If Ado_datos.Recordset("estado_codigo") = "REG" Then
'            Call OptFilGral1_Click
'        Else
'            Call OptFilGral2_Click
'        End If
        swnuevo = 0
'    Else
'        MsgBox "Saldo Insuficiente en Almacen Origen, debe realizar Transferencia de otro Almacen, Luego Intente nuevamente !..."
'    End If
  'Else
  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
  'End If
End Sub

Private Sub BtnImprimir2_Click()
  If ado_datos14.Recordset.RecordCount > 0 Then
    Dim iResult As Variant  ', i%, y%
    CryR01.ReportFileName = App.Path & "\reportes\tecnico\tr_R-362_Instalaciones.rpt"
    CryR01.WindowShowRefreshBtn = True
    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!tec_plan_codigo
    'Literal por el Total de la Compra
    var_literal = "UNO"    'Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
    CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
    'CryR01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
    
    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!tec_plan_codigo & "' "
    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
        Case "DNINS"
            var_titulo = "Módulo Instalaciones"
        Case "DNAJS"
            var_titulo = "Módulo Ajustes"
        Case "DNMAN"
            var_titulo = "Módulo Mantenimiento"
        Case "DNREP"
            var_titulo = "Módulo Reparaciones"
        Case "DNEME"
            var_titulo = "Módulo Emergencias"
        Case "DNMOD"
            var_titulo = "Módulo Modernización"
    End Select
    CryR01.Formulas(3) = "titulo = '" & var_titulo & "' "
    CryR01.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "
    iResult = CryR01.PrintReport
    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
'     ado_datos14.Recordset.Delete
'     ado_datos14.Recordset.Update
'     rs_datos14.Requery
'     ado_datos14.Refresh
'     'cerea
'     ado_datos14.Refresh
      db.Execute "update ao_ventas_detalle set ao_ventas_detalle.estado_codigo = 'ANL' Where ao_ventas_detalle.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_detalle.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_detalle.venta_codigo_det = " & ado_datos14.Recordset("venta_codigo_det") & " "
   End If
  Else
    MsgBox "Los Bienes del registro Aprobado o Anulado, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModDetalle_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    FraNavega.Enabled = False
    FrmDetalle.Enabled = False
    fraOpciones.Enabled = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
    
    swgrabar = 0
    swnuevo = 2
    marca1 = Ado_datos.Recordset.Bookmark
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(0) = False

    FrmEdita.Visible = True
    FrmEdita.Enabled = True
    TxtNroVenta.Text = Ado_datos.Recordset!tec_plan_codigo
    txt_carta.Text = "NO"
    Txt_cant1.Text = IIf(Txt_cant1.Text = "", "0.2", Txt_cant1.Text)
    Txt_cant2.Text = IIf(Txt_cant2.Text = "", "0.25", Txt_cant2.Text)
    Txt_cant3.Text = IIf(Txt_cant3.Text = "", "0.25", Txt_cant3.Text)
    Txt_cant4.Text = IIf(Txt_cant4.Text = "", "0", Txt_cant4.Text)
    Txt_cant5.Text = IIf(Txt_cant5.Text = "", "0", Txt_cant5.Text)
    If dtc_codigo6.Text = "" Or dtc_codigo6.Text <> "4211" Then
        dtc_codigo6.Text = "4211"
        dtc_desc6.BoundText = dtc_codigo6.BoundText
    End If
    If dtc_codigo6A.Text = "" Or dtc_codigo6A.Text <> "479" Then
        dtc_codigo6A.Text = "479"
        dtc_desc6A.BoundText = dtc_codigo6A.BoundText
    End If
    If dtc_codigo6B.Text = "" Or dtc_codigo6B.Text <> "500" Then
        dtc_codigo6B.Text = "500"
        dtc_desc6B.BoundText = dtc_codigo6B.BoundText
    End If
    If dtc_codigo6C.Text = "" Or dtc_codigo6C.Text <> "4529" Then
        dtc_codigo6C.Text = "4529"
        dtc_desc6C.BoundText = dtc_codigo6C.BoundText
    End If
    If dtc_codigo6D.Text = "" Then
        dtc_codigo6D.Text = "3113"
        dtc_desc6D.BoundText = dtc_codigo6D.BoundText
    End If
'    Set rs_datos14 = New ADODB.Recordset
'    If rs_datos14.State = 1 Then rs_datos14.Close
'    rs_datos14.Open "select * from tv_tecnico_cronograma_detalle_inst where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' and unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "'  AND tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & " ", db, adOpenKeyset, adLockOptimistic
'    Set ado_datos14.Recordset = rs_datos14
  Else
    MsgBox "Los datos del registro Aprobado o Anulado, NO pueden ser modificados !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo2.BoundText = Dtc_aux2.BoundText
    dtc_desc2.BoundText = Dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_aux4.BoundText
    dtc_desc4.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo15_LostFocus()
'    dtc_codigo15.Text = VAR_COD4
'    dtc_desc15.BoundText = dtc_codigo15.BoundText
''    dtc_unimed15.BoundText = dtc_codigo15.BoundText
'    dtc_grupo15.BoundText = dtc_codigo15.BoundText
'    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
'    Dtc_partida15.BoundText = dtc_codigo15.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    Dtc_aux2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_horaA5.BoundText = dtc_codigo5.BoundText
    dtc_horaB5.BoundText = dtc_codigo5.BoundText
    dtc_horaC5.BoundText = dtc_codigo5.BoundText
    dtc_horaD5.BoundText = dtc_codigo5.BoundText
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6A_Click(Area As Integer)
    dtc_desc6A.BoundText = dtc_codigo6A.BoundText
End Sub

Private Sub dtc_codigo6B_Click(Area As Integer)
    dtc_desc6B.BoundText = dtc_codigo6B.BoundText
End Sub

Private Sub dtc_codigo6C_Click(Area As Integer)
    dtc_desc6C.BoundText = dtc_codigo6C.BoundText
End Sub

Private Sub dtc_codigo6D_Click(Area As Integer)
    dtc_desc6D.BoundText = dtc_codigo6D.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    Dtc_aux2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    dtc_aux4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_horaA5.BoundText = dtc_desc5.BoundText
    dtc_horaB5.BoundText = dtc_desc5.BoundText
    dtc_horaC5.BoundText = dtc_desc5.BoundText
    dtc_horaD5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo15_Click(Area As Integer)
    dtc_desc15.BoundText = dtc_codigo15.BoundText
'    dtc_unimed15.BoundText = dtc_codigo15.BoundText
    dtc_grupo15.BoundText = dtc_codigo15.BoundText
    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
    Dtc_partida15.BoundText = dtc_codigo15.BoundText
    VAR_COD4 = dtc_codigo15.Text
End Sub

Private Sub dtc_desc6A_Click(Area As Integer)
    dtc_codigo6A.BoundText = dtc_desc6A.BoundText
End Sub

Private Sub dtc_desc6B_Click(Area As Integer)
    dtc_codigo6B.BoundText = dtc_desc6B.BoundText
End Sub

Private Sub dtc_desc6C_Click(Area As Integer)
    dtc_codigo6C.BoundText = dtc_desc6C.BoundText
End Sub

Private Sub dtc_desc6D_Click(Area As Integer)
    dtc_codigo6D.BoundText = dtc_desc6D.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub dtc_horaA5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_horaA5.BoundText
    dtc_horaB5.BoundText = dtc_horaA5.BoundText
    dtc_horaC5.BoundText = dtc_horaA5.BoundText
    dtc_horaD5.BoundText = dtc_horaA5.BoundText
    dtc_desc5.BoundText = dtc_horaA5.BoundText
End Sub

Private Sub dtc_horaB5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_horaB5.BoundText
    dtc_horaA5.BoundText = dtc_horaB5.BoundText
    dtc_horaC5.BoundText = dtc_horaB5.BoundText
    dtc_horaD5.BoundText = dtc_horaB5.BoundText
    dtc_desc5.BoundText = dtc_horaB5.BoundText
End Sub

Private Sub dtc_horaC5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_horaC5.BoundText
    dtc_horaA5.BoundText = dtc_horaC5.BoundText
    dtc_horaB5.BoundText = dtc_horaC5.BoundText
    dtc_horaD5.BoundText = dtc_horaC5.BoundText
    dtc_desc5.BoundText = dtc_horaC5.BoundText
End Sub

Private Sub dtc_horaD5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_horaD5.BoundText
    dtc_horaA5.BoundText = dtc_horaD5.BoundText
    dtc_horaB5.BoundText = dtc_horaD5.BoundText
    dtc_horaC5.BoundText = dtc_horaD5.BoundText
    dtc_desc5.BoundText = dtc_horaD5.BoundText
End Sub

Private Sub dtc_subgrupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_subgrupo15.BoundText
    dtc_desc15.BoundText = dtc_subgrupo15.BoundText
'    dtc_unimed15.BoundText = dtc_subgrupo15.BoundText
    dtc_grupo15.BoundText = dtc_subgrupo15.BoundText
    Dtc_partida15.BoundText = dtc_subgrupo15.BoundText
End Sub

Private Sub dtc_partida15_Click(Area As Integer)
    dtc_desc15.BoundText = Dtc_partida15.BoundText
'    dtc_unimed15.BoundText = Dtc_partida15.BoundText
    dtc_grupo15.BoundText = Dtc_partida15.BoundText
    dtc_subgrupo15.BoundText = Dtc_partida15.BoundText
    dtc_codigo15.BoundText = Dtc_partida15.BoundText
End Sub

Private Sub dtc_desc15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_desc15.BoundText
'    dtc_unimed15.BoundText = dtc_desc15.BoundText
    dtc_grupo15.BoundText = dtc_desc15.BoundText
    dtc_subgrupo15.BoundText = dtc_desc15.BoundText
    Dtc_partida15.BoundText = dtc_desc15.BoundText
End Sub

Private Sub dtc_desc2_LostFocus()
'    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
'    If Dtc_deudor2.Text = "SI" Then
'        Dtc_deudor2.backColor = &HFF&
'    Else
'        Dtc_deudor2.backColor = &H80000010
'    End If
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_grupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_grupo15.BoundText
    dtc_desc15.BoundText = dtc_grupo15.BoundText
'    dtc_unimed15.BoundText = dtc_grupo15.BoundText
    dtc_subgrupo15.BoundText = dtc_grupo15.BoundText
    Dtc_partida15.BoundText = dtc_grupo15.BoundText
End Sub

Private Sub dtc_desc12_LostFocus()
'  If GlSistema = "A" Then       'Or GlSistema = "Z"
'    If dtc_codigo12.Text = "10" Then
'        TxtPrecioU.Text = dtc_precioventabase15.Text
'    Else
'        TxtPrecioU.Text = dtc_precioventafinal15.Text
'    End If
'  Else
'    'If lblventa_tipo.Caption = "E" Then
'    '    TxtPrecioU.Text = dtc_precioventafinal15.Text
'    'Else
'    '    TxtPrecioU.Text = dtc_precioventabase15.Text
'    'End If
'    If Val(dtc_codigo12.Text) > 19 Then
'        TxtPrecioU.Text = dtc_precioventafinal15.Text
'    Else
'        TxtPrecioU.Text = dtc_precioventabase15.Text
'    End If
'    If Val(dtc_codigo12.Text) = 100 Then
'        TxtPrecioU.Text = dtc_preciocompra15.Text
'    End If
'    If Val(dtc_codigo12.Text) = 200 Then
'        TxtPrecioU.Text = "0"
'    End If
'  End If
End Sub

'Private Sub dtc_unimed15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_unimed15.BoundText
'    dtc_desc15.BoundText = dtc_unimed15.BoundText
''    dtc_stocktotal15.BoundText = dtc_unimed15.BoundText
'    dtc_grupo15.BoundText = dtc_unimed15.BoundText
'    dtc_subgrupo15.BoundText = dtc_unimed15.BoundText
'    Dtc_partida15.BoundText = dtc_unimed15.BoundText
''    dtc_precioventafinal15.BoundText = dtc_unimed15.BoundText
''    dtc_precioventabase15.BoundText = dtc_unimed15.BoundText
''    dtc_preciocompra15.BoundText = dtc_unimed15.BoundText
'End Sub

Private Sub DTPFechaFin_LostFocus()
    txt_dias.Text = Val(DTPfechaFin.Value) - Val(DTPFechaInicio.Value)
    'txt_dias.Text = CStr(Val(DTPFechaInicio.Value) - Val(DTPFechaFin.Value))
End Sub

Private Sub DTPfechasol_LostFocus()
    Set rs_TipoCambio = New ADODB.Recordset
    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
    End If
    'Ado_datos4.Refresh
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux8.RecordCount > 0 Then
        usuario2 = rs_aux8!beneficiario_codigo
        VAR_DA = rs_aux8!da_codigo
        VAR_DPTOC = rs_aux8!depto_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.3"
        VAR_DPTOC = "2"
    End If
    If Aux = "DNINS" Then
        Select Case VAR_DPTOC
            Case "1"    ' Chuquisaca
                VAR_UORIGEN = "DINSC"
            Case "2"    'La Paz - Tecnico
                VAR_UORIGEN = "DNINS"
            Case "3"    'Cochabamba
                VAR_UORIGEN = "DINSB"
                'VAR_DPTOC = "3"
            Case "7"    'Santa Cruz
                VAR_UORIGEN = "DINSS"
                'VAR_DPTOC = "7"
            Case "4"    'Oruro - Tecnico
                VAR_UORIGEN = "DINSB"
                'VAR_DPTOC = "2"
            Case "5"    ' Potosi
                VAR_UORIGEN = "DINSC"
            Case "6"    ' Tarija
                VAR_UORIGEN = "DNINS"
            Case "8"    ' Beni
                VAR_UORIGEN = "DINSS"
            Case "9"    ' Pando
                VAR_UORIGEN = "DINSS"
            Case Else    ' TODO
                VAR_UORIGEN = "DNINS"
                VAR_DPTOC = "0"
         End Select
    End If
    
    parametro = Aux
    db.Execute "update to_cronograma_inst set to_cronograma_inst.zpiloto_codigo = tc_zona_piloto_edif_inst.zpiloto_codigo from to_cronograma_inst  INNER JOIN tc_zona_piloto_edif_inst ON to_cronograma_inst.edif_codigo = tc_zona_piloto_edif_inst.edif_codigo WHERE to_cronograma_inst.unidad_codigo_tec = '" & VAR_UORIGEN & "' OR to_cronograma_inst.unidad_codigo_tec = '" & Aux & "' "
    db.Execute "update to_cronograma_inst set to_cronograma_inst.zona_edif_orden = tc_zona_piloto_edif_inst.zona_edif_orden from to_cronograma_inst  INNER JOIN tc_zona_piloto_edif_inst ON to_cronograma_inst.zpiloto_codigo = tc_zona_piloto_edif_inst.zpiloto_codigo AND to_cronograma_inst.edif_codigo = tc_zona_piloto_edif_inst.edif_codigo WHERE (to_cronograma_inst.unidad_codigo_tec = '" & VAR_UORIGEN & "' OR to_cronograma_inst.unidad_codigo_tec = '" & Aux & "')  AND to_cronograma_inst.estado_codigo <> 'ANL'"
    'db.Execute "update to_cronograma_inst set to_cronograma_inst.ges_gestion = YEAR(to_cronograma_inst.fecha_inicio_tec ) where to_cronograma_inst.ges_gestion <> YEAR(to_cronograma_inst.fecha_inicio_tec ) and to_cronograma_inst.unidad_codigo_tec = '" & VAR_UORIGEN & "'  OR to_cronograma_inst.unidad_codigo_tec = '" & Aux & "' "
    db.Execute "update to_cronograma_detalle_inst set to_cronograma_detalle_inst.ges_gestion = to_cronograma_inst.ges_gestion from to_cronograma_detalle_inst inner join to_cronograma_inst on to_cronograma_detalle_inst.unidad_codigo_tec = to_cronograma_inst.unidad_codigo_tec  and to_cronograma_detalle_inst.tec_plan_codigo = to_cronograma_inst.tec_plan_codigo AND dbo.to_cronograma_inst.ges_gestion <> dbo.to_cronograma_detalle_inst.ges_gestion "
    db.Execute "update to_cronograma_inst set to_cronograma_inst.beneficiario_codigo_resp = tc_zonas_piloto_inst.beneficiario_codigo from to_cronograma_inst inner join tc_zonas_piloto_inst on to_cronograma_inst.zpiloto_codigo  = tc_zonas_piloto_inst.zpiloto_codigo WHERE (to_cronograma_inst.unidad_codigo_tec = '" & VAR_UORIGEN & "'  OR to_cronograma_inst.unidad_codigo_tec = '" & Aux & "' )  and to_cronograma_inst.estado_codigo = 'REG' "
    'db.Execute "update ao_ventas_cabecera set edif_codigo_corto= substring(edif_codigo,7,len(edif_codigo)) where edif_codigo_corto is null "
    
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    If glusuario = "VPAREDES" Then
        BtnModificar.Visible = False
        BtnEliminar.Visible = False
        BtnAprobar.Visible = False
        BtnModDetalle.Visible = False
        BtnAddDetalle2.Visible = False
        BtnAnlDetalle2.Visible = False
        CmdGrabaDet.Visible = False
    End If

    'txt_codigo.Enabled = True
    mbDataChanged = False
    FrmCabecera.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
'    BtnImprimir2.Visible = False
'    BtnImprimir3.Visible = False
'    FrmEdita.Enabled = False
'    FrmCobros.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'    SSTab1.TabVisible(2) = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
'    Chk_plazo.Value = 0
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas (Cliente, Proveedor, etc.)
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset     'Beneficiario Funcionario CGI (Vendedor, Cobrador, Adm, etc.)
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    'VAR_UORIGEN
    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & VAR_UORIGEN & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from rc_horarios ", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    'dtc_desc5.BoundText = dtc_codigo5.BoundText

    'INSUMOS
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "select distinct * from av_bienes_vs_venta_detalle where par_codigo = '33100' or par_codigo = '34110' ORDER BY bien_descripcion ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText

    'ZONAS PILOTO
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "select * from tc_zonas_piloto_inst order by zpiloto_descripcion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    rs_datos15.Open "Select * from ac_bienes where par_codigo = '43340' ", db, adOpenStatic
    Set ado_datos15.Recordset = rs_datos15
    dtc_desc15.BoundText = dtc_codigo15.BoundText
'    dtc_unimed15.BoundText = dtc_codigo15.BoundText
    dtc_grupo15.BoundText = dtc_codigo15.BoundText
    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
    Dtc_partida15.BoundText = dtc_codigo15.BoundText
End Sub

Private Sub grabar()
  'db.BeginTrans
    If swgrabar = 1 Then
'      Dim rstdestino As New ADODB.Recordset
'      Set rstdestino = New ADODB.Recordset
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select tipo_tramite, numero_correlativo from fc_correl WHERE tipo_tramite='ventas'", db, adOpenDynamic, adLockOptimistic
'      If rstdestino.RecordCount <> 0 Then
'        Ado_datos.Recordset("venta_codigo") = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino!numero_correlativo = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino.Update
'      Else
'        Ado_datos.Recordset("venta_codigo") = 1
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'      'Ado_datos.Recordset("venta_codigo") = Ado_datos.Recordset.RecordCount
'      'rstdestino.AddNew
'       Ado_datos.Recordset("ges_gestion") = GlGestion       'CStr(Year(DTPfechasol.Value))
'       Ado_datos.Recordset("unidad_codigo_tec") = dtc_codigo1.Text
'       Ado_datos.Recordset("tec_plan_codigo") = txt_codigo.Caption
    End If
       VAR_UNITEC = Ado_datos.Recordset!unidad_codigo_tec
       correldet2 = Ado_datos.Recordset!tec_plan_codigo
       Select Case cmd_unimed2.Text
           Case "MES"
               UMED_NRO = 1
           Case "BMES"
               UMED_NRO = 2
           Case "TMES"
               UMED_NRO = 3
           Case "CMES"
               UMED_NRO = 4
           Case "5MES"
               UMED_NRO = 5
           Case "SMES"
               UMED_NRO = 6
           Case "7MES"
               UMED_NRO = 7
           Case "8MES"
               UMED_NRO = 8
           Case "9MES"
               UMED_NRO = 9
           Case "10MES"
               UMED_NRO = 10
           Case "11MES"
               UMED_NRO = 11
           Case "ANUAL"
               UMED_NRO = 12
       End Select
       Select Case RTrim(cmb_mes_ini.Text)
           Case "ENERO"
               VAR_MESINI2 = 1
           Case "FEBRERO"
               VAR_MESINI2 = 2
           Case "MARZO"
               VAR_MESINI2 = 3
           Case "ABRIL"
               VAR_MESINI2 = 4
           Case "MAYO"
               VAR_MESINI2 = 5
           Case "JUNIO"
               VAR_MESINI2 = 6
           Case "JULIO"
               VAR_MESINI2 = 7
           Case "AGOSTO"
               VAR_MESINI2 = 8
           Case "SEPTIEMBRE"
               VAR_MESINI2 = 9
           Case "OCTUBRE"
               VAR_MESINI2 = 10
           Case "NOVIEMBRE"
               VAR_MESINI2 = 11
           Case "DICIEMBRE"
               VAR_MESINI2 = 12
       End Select
       'ini
       Select Case VAR_UNITEC
            Case "COMEX"        'INI GRABA COMEX
                
            Case "DVTA", "DCOMS", "DCOMB", "DCOMC"                        'INI GRABA VENTAS
                VAR_PROC = "TEC"
                VAR_SUB = "TEC-01"
                VAR_ETAPA = "TEC-01-02"
                VAR_CLASIF = "TEC"
                VAR_DOC = "R-362"
                VAR_DOCNUM = correldet2
                VAR_POA = "3.2.2"
                VAR_SOLTIPO = 3
            
            Case "DNINS", "DINSS", "DINSB", "DINSC"                        'INI GRABA INSTALACIONES
                VAR_PROC = "COM"
                VAR_SUB = "COM-03"
                VAR_ETAPA = "COM-03-02"
                VAR_CLASIF = "TEC"
                VAR_DOC = "R-362"
                VAR_DOCNUM = correldet2
                VAR_POA = "3.2.2"
                VAR_SOLTIPO = 4
                          
            Case "DNAJS"
                VAR_PROC = "COM"
                VAR_SUB = "COM-04"
                VAR_ETAPA = "COM-04-02"
                VAR_CLASIF = "TEC"
                VAR_DOC = "R-362"
                VAR_DOCNUM = correldet2
                VAR_POA = "3.2.6"
                VAR_SOLTIPO = 5
                                
            Case "DNMAN", "DMANS", "DMANB", "DMANC"                'MANTENIMIENTO
                VAR_PROC = "TEC"
                VAR_SUB = "TEC-02"
                VAR_ETAPA = "TEC-02-03"
                VAR_CLASIF = "TEC"
                VAR_DOC = "R-362"
                VAR_DOCNUM = correldet2
                VAR_POA = "3.2.3"
                VAR_SOLTIPO = 10
    
            Case Else
                VAR_PROC = "TEC"
                VAR_SUB = "TEC-02"
                VAR_ETAPA = "TEC-02-03"
                VAR_CLASIF = "TEC"
                VAR_DOC = "R-362"
                VAR_DOCNUM = correldet2
                VAR_POA = "3.2.3"
                VAR_SOLTIPO = 10
         End Select
       'fin
       
       cmb_lunes.Text = "SI"
       cmb_primero.Text = "SI"
       If dtc_codigo5.Text = "" Then
            dtc_codigo5.Text = "0"
       End If
       If cmb_dia.Text = "" Then
            cmb_dia.Text = "AUTOMATICO"
       End If
       db.Execute "Update to_cronograma_inst Set zpiloto_codigo = " & dtc_codigo7.Text & ", beneficiario_codigo_resp = '" & dtc_codigo4.Text & "', tec_plan_fecha = '" & DTPfechasol.Value & "', tec_cantidad_unidades= " & Val(txt_cant.Text) & ", lunes_cambia = '" & cmb_lunes.Text & "', primero_mes='" & cmb_primero.Text & "', unimed_codigo_nro= " & UMED_NRO & ", mes_inicio_crono_nro= " & VAR_MESINI2 & ", " & _
                " proceso_codigo='" & VAR_PROC & "', subproceso_codigo='" & VAR_SUB & "', etapa_codigo= '" & VAR_ETAPA & "', clasif_codigo='" & VAR_CLASIF & "', doc_codigo = '" & VAR_DOC & "', doc_numero= " & VAR_DOCNUM & ", poa_codigo = '" & VAR_POA & "', solicitud_tipo = " & VAR_SOLTIPO & ",  " & _
                " unimed_codigo='" & cmd_unimed2.Text & "', fecha_inicio_tec='" & lbl_fecha_ini.Value & "', fecha_fin_tec= '" & lbl_fecha_fin.Value & "', mes_inicio_crono='" & cmb_mes_ini.Text & "', horario_codigo = '" & dtc_codigo5.Text & "', dia_nombre='" & cmb_dia.Text & "', usr_codigo = '" & glusuario & "', fecha_registro = '" & Date & "'  " & _
                " Where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "   "

       db.Execute "Update to_cronograma_detalle_inst Set beneficiario_codigo = '" & dtc_codigo4.Text & "' Where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "  "
'    If Ado_datos.Recordset("estado_codigo") = "REG" Then
'        Call OptFilGral1_Click
'    Else
'        Call OptFilGral2_Click
'    End If
      
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

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
  'Actualiza zona piloto de la gestion
'    db.Execute "update to_cronograma set to_cronograma.zona_edif_orden = tc_zona_piloto_edif.zona_edif_orden from to_cronograma  INNER JOIN tc_zona_piloto_edif ON to_cronograma.edif_codigo = tc_zona_piloto_edif.edif_codigo WHERE (to_cronograma.zpiloto_codigo <> tc_zona_piloto_edif.zpiloto_codigo and to_cronograma.ges_gestion = '" & Trim(OptFilGral1.Caption) & "' ) AND to_cronograma.estado_codigo <> 'ANL'"
'    db.Execute "update to_cronograma set to_cronograma.zpiloto_codigo = tc_zona_piloto_edif.zpiloto_codigo from to_cronograma  INNER JOIN tc_zona_piloto_edif ON to_cronograma.edif_codigo = tc_zona_piloto_edif.edif_codigo WHERE to_cronograma.zpiloto_codigo <> tc_zona_piloto_edif.zpiloto_codigo and to_cronograma.ges_gestion = '" & Trim(OptFilGral1.Caption) & "'  "
'
'    db.Execute "update to_cronograma set to_cronograma.ges_gestion = YEAR(to_cronograma.fecha_inicio_tec) where to_cronograma.ges_gestion <> YEAR(to_cronograma.fecha_inicio_tec) and to_cronograma.unidad_codigo_tec = '" & VAR_UORIGEN & "'  OR to_cronograma.unidad_codigo_tec = '" & Aux & "' "
'    db.Execute "update to_cronograma_detalle set to_cronograma_detalle.ges_gestion = to_cronograma.ges_gestion from to_cronograma_detalle inner join to_cronograma on to_cronograma_detalle.unidad_codigo_tec = to_cronograma.unidad_codigo_tec  and to_cronograma_detalle.tec_plan_codigo = to_cronograma.tec_plan_codigo AND dbo.to_cronograma.ges_gestion <> dbo.to_cronograma_detalle.ges_gestion "
'    db.Execute "update to_cronograma set to_cronograma.beneficiario_codigo_resp = tc_zonas_piloto.beneficiario_codigo from to_cronograma inner join tc_zonas_piloto on to_cronograma.zpiloto_codigo  = tc_zonas_piloto.zpiloto_codigo WHERE (to_cronograma.unidad_codigo_tec = '" & VAR_UORIGEN & "'  OR to_cronograma.unidad_codigo_tec = '" & Aux & "' )  and to_cronograma.estado_codigo = 'REG' "

    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE  (estado_codigo = 'REG') "
'    Select Case VAR_DPTOC
'        Case "1"    ' Chuquisaca
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE  ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "' or left(edif_codigo,1) = '5' or left(edif_codigo,1) = '6' )) AND ges_gestion = '2021' ) "
'        Case "2"    'La Paz - Tecnico
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "' or left(edif_codigo,1) = '4' )) AND ges_gestion = '2021' ) "
'        Case "3"    'Cochabamba
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND left(edif_codigo,1) = '" & VAR_DPTOC & "') AND ges_gestion = '2021' ) "
'        Case "7"    'Santa Cruz
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9' or zpiloto_codigo = '34')) AND ges_gestion = '2021' ) "
'        Case "4"    'Oruro - Tecnico
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND left(edif_codigo,1) = '" & VAR_DPTOC & "') AND ges_gestion = '2021' ) "
'        Case "5"    ' Potosi
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND left(edif_codigo,1) = '" & VAR_DPTOC & "') AND ges_gestion = '2021' ) "
'        Case "6"    ' Tarija
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND left(edif_codigo,1) = '" & VAR_DPTOC & "') AND ges_gestion = '2021' ) "
'        Case "8"    ' Beni
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND left(edif_codigo,1) = '" & VAR_DPTOC & "') AND ges_gestion = '2021' ) "
'        Case "9"    ' Pando
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((estado_codigo = 'REG' AND (unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND left(edif_codigo,1) = '" & VAR_DPTOC & "') AND ges_gestion = '2021' ) "
'        Case Else    ' TODO
'            queryinicial = "select * From tv_cronograma_edificaciones_inst where ( ges_gestion = '2021' ) "
'     End Select

    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "zpiloto_codigo, zona_edif_orden"       'ges_gestion,
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From tv_cronograma_edificaciones_inst  "
    
'    Select Case VAR_DPTOC
'        Case "1"    ' Chuquisaca
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "' or left(edif_codigo,1) = '5' or left(edif_codigo,1) = '6')) "
'        Case "2"    'La Paz - Tecnico
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "' or left(edif_codigo,1) = '4' )) "
'        Case "3"    'Cochabamba
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "')) "
'        Case "7"    'Santa Cruz
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9' or zpiloto_codigo = '34')) "
'        Case "4"    'Oruro - Tecnico
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "')) "
'        Case "5"    ' Potosi
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "')) "
'        Case "6"    ' Tarija
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "')) "
'        Case "8"    ' Beni
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "')) "
'        Case "9"    ' Pando
'            queryinicial = "select * From tv_cronograma_edificaciones_inst WHERE ((unidad_codigo_tec = '" & parametro & "' OR unidad_codigo_tec= '" & VAR_UORIGEN & "') AND (left(edif_codigo,1) = '" & VAR_DPTOC & "')) "
'        Case Else    ' TODO
'            queryinicial = "select * From tv_cronograma_edificaciones_inst  "
'     End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "zpiloto_codigo, zona_edif_orden"           'ges_gestion,
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

'Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
' If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub

'Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
'    'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

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
'  txt_venta = " "
'  dtc_codigo4.Text = " "
'  TxtMontoUs = "0"
'  TxtConcepto = ""
'  dtc_codigo2 = ""
'  dtc_desc2 = ""
'  txtTDC.Text = GlTipoCambioOficial
'
''  txt_venta = ""
''  txtterref = ""
End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        'SSTab1.TabEnabled(0) = True
        'SSTab1.TabEnabled(1) = False
    Else
'           FrmEditaDet.Visible = False
'           DtGLista.Visible = False
'           adoao_solicitud_lista.Visible = False
    End If

End Sub

'Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
'End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ABRIR_TABLA_DET()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from to_cronograma_diario where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "  and year(dia_fecha)= '" & Ado_datos.Recordset!ges_gestion & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    FrmDetalle.Caption = "Detalle del Equipos del Cronograma Nro. " + Str((Ado_datos.Recordset("tec_plan_codigo")))
End Sub




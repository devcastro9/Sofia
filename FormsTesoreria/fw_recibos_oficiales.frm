VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_recibos_oficiales 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tesoreria - Recibos de Ingresos"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   10815
   Icon            =   "fw_recibos_oficiales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   1.08741e5
   ScaleMode       =   0  'User
   ScaleWidth      =   2.17485e7
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmDetalle2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE COBRANZAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2025
      Left            =   1560
      TabIndex        =   103
      Top             =   7440
      Width           =   16935
      Begin MSDataGridLib.DataGrid DtGLista11 
         Bindings        =   "fw_recibos_oficiales.frx":0A02
         Height          =   1740
         Left            =   120
         TabIndex        =   104
         Top             =   240
         Width           =   16740
         _ExtentX        =   29528
         _ExtentY        =   3069
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "correl_cobro"
            Caption         =   "Correl."
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
            DataField       =   "cmpbte_deposito"
            Caption         =   "Cheq/Cmpbte."
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
            DataField       =   "cmpbte_fecha"
            Caption         =   "Fecha.Cmpbte"
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
            DataField       =   "cta_codigo"
            Caption         =   "Cta.Banco/Caja"
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
            DataField       =   "cobranza_nro_factura"
            Caption         =   "#Factura"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   764.787
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
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   675.213
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   3974.74
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column13 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   4305.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   4305
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   17.688
      ScaleMode       =   4  'Character
      ScaleWidth      =   11.625
      TabIndex        =   101
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton CmdOK100 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Height          =   555
         Left            =   0
         Picture         =   "fw_recibos_oficiales.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Busca en Detalle de Cobranzas"
         Top             =   192
         Width           =   1365
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000015&
         Height          =   525
         Left            =   0
         Picture         =   "fw_recibos_oficiales.frx":11D1
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Acepta envio a Recibo"
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Height          =   525
         Left            =   0
         Picture         =   "fw_recibos_oficiales.frx":19BF
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Modifica Detalle Recibo Elegido"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000015&
         Height          =   525
         Left            =   0
         Picture         =   "fw_recibos_oficiales.frx":22D4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H80000018&
         Height          =   525
         Left            =   0
         Picture         =   "fw_recibos_oficiales.frx":2A20
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprime Kardex del Bien"
         Top             =   1830
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   93
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3720
         Picture         =   "fw_recibos_oficiales.frx":32ED
         ScaleHeight     =   735
         ScaleWidth      =   1320
         TabIndex        =   9
         ToolTipText     =   "Aprueba Entrega de Insumos"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3720
         Picture         =   "fw_recibos_oficiales.frx":3B20
         ScaleHeight     =   735
         ScaleWidth      =   1320
         TabIndex        =   120
         ToolTipText     =   "Aprueba Entrega de Insumos"
         Top             =   20
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5040
         Picture         =   "fw_recibos_oficiales.frx":4517
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   10
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2520
         Picture         =   "fw_recibos_oficiales.frx":4CCC
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   8
         ToolTipText     =   "Anula Zona elegida"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17400
         Picture         =   "fw_recibos_oficiales.frx":5418
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   95
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1185
         Picture         =   "fw_recibos_oficiales.frx":5BDA
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   7
         ToolTipText     =   "Modifica datos de la Zona elegida"
         Top             =   0
         Width           =   1430
      End
      Begin VB.CommandButton BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Caption         =   "Digitaliza"
         Height          =   710
         Left            =   9000
         Picture         =   "fw_recibos_oficiales.frx":64EF
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "fw_recibos_oficiales.frx":6931
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   6360
         Picture         =   "fw_recibos_oficiales.frx":70F0
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   15
         ToolTipText     =   "Comprobante de Salida de Almacenes"
         Top             =   0
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TESORERIA"
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
         Left            =   13035
         TabIndex        =   96
         Top             =   195
         Width           =   1425
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
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "fw_recibos_oficiales.frx":79BD
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   91
         Top             =   0
         Width           =   1280
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "fw_recibos_oficiales.frx":8193
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   90
         Top             =   0
         Width           =   1455
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
         Left            =   13095
         TabIndex        =   92
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   1935
      Left            =   12000
      TabIndex        =   83
      Top             =   5640
      Visible         =   0   'False
      Width           =   6135
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
         ScaleWidth      =   5880
         TabIndex        =   97
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3360
            Picture         =   "fw_recibos_oficiales.frx":8A7F
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   99
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "fw_recibos_oficiales.frx":936B
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   98
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   0
            Width           =   1455
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
            TabIndex        =   100
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.PictureBox DTP_Finicio 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   1425
         TabIndex        =   84
         Top             =   1560
         Width           =   1485
      End
      Begin VB.PictureBox DTP_Ffin 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   1425
         TabIndex        =   85
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE FIN"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3600
         TabIndex        =   87
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE INICIO"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   840
         TabIndex        =   86
         Top             =   1200
         Width           =   1620
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4290
      Left            =   6600
      TabIndex        =   25
      Top             =   765
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   7567
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
      TabCaption(0)   =   "RECIBOS INGRESOS DE TESORERIA"
      TabPicture(0)   =   "fw_recibos_oficiales.frx":9C38
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DETALLE DE RECIBOS DE INGRESOS"
      TabPicture(1)   =   "fw_recibos_oficiales.frx":9C54
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEdita"
      Tab(1).ControlCount=   1
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
         Height          =   3990
         Left            =   -75000
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   11895
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   6225
            TabIndex        =   88
            Top             =   1825
            Width           =   375
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   8460
            TabIndex        =   70
            Top             =   2415
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   5220
            TabIndex        =   69
            Top             =   2415
            Width           =   255
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   2415
            Width           =   255
         End
         Begin VB.PictureBox FraGrabarDet 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   900
            Left            =   0
            ScaleHeight     =   840
            ScaleWidth      =   11880
            TabIndex        =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   11940
            Begin VB.CommandButton CmdGrabaDet 
               BackColor       =   &H80000015&
               Height          =   650
               Left            =   4515
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   120
               Width           =   1365
            End
            Begin VB.CommandButton CmdCancelaDet 
               BackColor       =   &H80000016&
               Height          =   650
               Left            =   5850
               MaskColor       =   &H00000000&
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Cancelar"
               Top             =   120
               Width           =   1365
            End
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   10980
            TabIndex        =   47
            Top             =   3015
            Width           =   255
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   8520
            TabIndex        =   46
            Top             =   1825
            Width           =   255
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   10980
            TabIndex        =   45
            Top             =   3975
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   10980
            TabIndex        =   44
            Top             =   1830
            Width           =   255
         End
         Begin VB.TextBox txt_descripcion_venta 
            CausesValidation=   0   'False
            DataField       =   "concepto_venta"
            DataSource      =   "ado_datos14"
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            MaxLength       =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   3360
            Width           =   8505
         End
         Begin VB.TextBox TxtNroVenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "venta_codigo"
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
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox TxtCantidad 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "venta_det_cantidad"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "0"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox TxtDescuento 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "bien_cantidad_por_empaque"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6600
            TabIndex        =   18
            Text            =   "0"
            Top             =   2520
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo dtc_codigo15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6600
            TabIndex        =   30
            Top             =   1815
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   240
            TabIndex        =   16
            Top             =   1815
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc13 
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   5520
            TabIndex        =   19
            Top             =   1080
            Visible         =   0   'False
            Width           =   5730
            _ExtentX        =   10107
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_descripcion"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo13 
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7560
            TabIndex        =   43
            Top             =   840
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "almacen_codigo"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_partida15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7320
            TabIndex        =   48
            Top             =   2400
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "par_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Grupo"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3240
            TabIndex        =   67
            Top             =   2400
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Partida"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   6480
            TabIndex        =   59
            Top             =   2415
            Width           =   645
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Almacen de Origen:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3960
            TabIndex        =   35
            Top             =   1095
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Actual (Total Nacional)"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   9675
            TabIndex        =   42
            Top             =   3360
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Medida"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9795
            TabIndex        =   41
            Top             =   1560
            Width           =   1515
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Venta:"
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
            Height          =   240
            Left            =   240
            TabIndex        =   40
            Top             =   1095
            Width           =   1500
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción y/o Características Complementarias"
            ForeColor       =   &H00000000&
            Height          =   120
            Left            =   240
            TabIndex        =   39
            Top             =   3075
            Width           =   4425
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Entregada"
            ForeColor       =   &H00000000&
            Height          =   120
            Left            =   4560
            TabIndex        =   38
            Top             =   2520
            Width           =   1800
         End
         Begin VB.Label lbl_des_bien 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Bien"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   37
            Top             =   1560
            Width           =   1860
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Bien"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6600
            TabIndex        =   36
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Solicitada"
            ForeColor       =   &H00000000&
            Height          =   120
            Left            =   225
            TabIndex        =   34
            Top             =   2520
            Width           =   1770
         End
         Begin VB.Label Label24 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   240
            TabIndex        =   33
            Top             =   2415
            Width           =   690
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Actual (Almacen Origen)"
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   9660
            TabIndex        =   32
            Top             =   2520
            Width           =   1785
         End
      End
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
         Height          =   3870
         Left            =   60
         TabIndex        =   27
         Top             =   360
         Width           =   11895
         Begin MSComCtl2.DTPicker DTPfechasol 
            DataField       =   "fecha_recibo"
            DataSource      =   "Ado_datos"
            Height          =   300
            Left            =   1920
            TabIndex        =   121
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            Format          =   110493697
            CurrentDate     =   44419
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   7035
            TabIndex        =   117
            Top             =   390
            Width           =   270
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "----------------------------- DESTINO "
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
            Height          =   2085
            Left            =   5960
            TabIndex        =   78
            Top             =   1755
            Width           =   5895
            Begin VB.TextBox Text12 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   290
               Left            =   2805
               TabIndex        =   119
               Top             =   980
               Width           =   270
            End
            Begin MSDataListLib.DataCombo dtc_desc5 
               Bindings        =   "fw_recibos_oficiales.frx":9C70
               DataField       =   "beneficiario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1395
               TabIndex        =   2
               Top             =   420
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo5 
               Bindings        =   "fw_recibos_oficiales.frx":9C89
               DataField       =   "beneficiario_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1395
               TabIndex        =   3
               Top             =   960
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_desc22 
               DataField       =   "depto_codigo_dR"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1395
               TabIndex        =   5
               Top             =   1440
               Visible         =   0   'False
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "depto_descripcion"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo22 
               DataField       =   "depto_codigo_dR"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4755
               TabIndex        =   82
               Top             =   1440
               Visible         =   0   'False
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin VB.Label lbl_Rdestino 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Regional "
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
               TabIndex        =   81
               Top             =   1485
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.Label lbl_Adestino 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "C.I.:"
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
               Left            =   930
               TabIndex        =   80
               Top             =   960
               Width           =   315
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Entregado a:"
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
               Left            =   120
               TabIndex        =   79
               Top             =   480
               Width           =   1155
            End
         End
         Begin VB.Frame Fra_datos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "-------------------------------- ORIGEN "
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
            Height          =   2085
            Left            =   40
            TabIndex        =   72
            Top             =   1755
            Width           =   5895
            Begin VB.TextBox Text11 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   290
               Left            =   2790
               TabIndex        =   118
               Top             =   980
               Width           =   270
            End
            Begin VB.ComboBox cmd_unimed2 
               DataField       =   "unimed_codigo_cobr"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   6210
               TabIndex        =   73
               Text            =   "ANUAL"
               Top             =   1080
               Visible         =   0   'False
               Width           =   555
            End
            Begin MSDataListLib.DataCombo dtc_desc4 
               Bindings        =   "fw_recibos_oficiales.frx":9CA2
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1380
               TabIndex        =   0
               Top             =   420
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo4 
               Bindings        =   "fw_recibos_oficiales.frx":9CBB
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1380
               TabIndex        =   1
               Top             =   960
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_desc21 
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1380
               TabIndex        =   4
               Top             =   1440
               Visible         =   0   'False
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "depto_descripcion"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo21 
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4320
               TabIndex        =   74
               Top             =   1440
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Recibido de:"
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
               Left            =   120
               TabIndex        =   77
               Top             =   480
               Width           =   1275
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "C.I.:"
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
               Left            =   765
               TabIndex        =   76
               Top             =   975
               Width           =   435
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Regional"
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
               TabIndex        =   75
               Top             =   1485
               Visible         =   0   'False
               Width           =   945
            End
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   8625
            TabIndex        =   58
            Top             =   390
            Width           =   270
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "fw_recibos_oficiales.frx":9CD4
            DataField       =   "doc_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7305
            TabIndex        =   57
            Top             =   375
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "doc_codigo"
            BoundColumn     =   "doc_codigo"
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
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "fw_recibos_oficiales.frx":9CED
            DataField       =   "doc_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1725
            TabIndex        =   17
            Top             =   375
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "doc_descripcion"
            BoundColumn     =   "doc_codigo"
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
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "fw_recibos_oficiales.frx":9D06
            DataField       =   "doc_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6960
            TabIndex        =   56
            Top             =   240
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "clasif_codigo"
            BoundColumn     =   "doc_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label21 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Entrega"
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
            Left            =   180
            TabIndex        =   71
            Top             =   1200
            Width           =   1710
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFF80&
            X1              =   11880
            X2              =   0
            Y1              =   915
            Y2              =   915
         End
         Begin VB.Label lbl_cerrado 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "RECIBO CONCILIADO !!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   3480
            TabIndex        =   66
            Top             =   -30
            Width           =   4875
         End
         Begin VB.Label txt_venta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "total_dol"
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
            Left            =   10500
            TabIndex        =   65
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo.Tes."
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
            Index           =   13
            Left            =   9120
            TabIndex        =   64
            Top             =   375
            Width           =   1290
         End
         Begin VB.Label txt_campo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "correl_doc"
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
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   10395
            TabIndex        =   63
            Top             =   375
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Documento ISO"
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
            Left            =   180
            TabIndex        =   55
            Top             =   375
            Width           =   1650
         End
         Begin VB.Label txt_codigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "Total_bs"
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
            Left            =   6120
            TabIndex        =   54
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Total Bs."
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
            Left            =   4560
            TabIndex        =   53
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label Label15 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Total Dolares"
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
            Height          =   285
            Left            =   8580
            TabIndex        =   28
            Top             =   1200
            Width           =   1845
         End
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00C00000&
      Height          =   4320
      Left            =   120
      TabIndex        =   49
      Top             =   720
      Width           =   6465
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3840
         TabIndex        =   52
         Top             =   3915
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   51
         Top             =   3915
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "fw_recibos_oficiales.frx":9D1F
         Height          =   3570
         Left            =   75
         TabIndex        =   50
         Top             =   240
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   6297
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "clasif_codigo"
            Caption         =   "Clasificacion"
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
            DataField       =   "doc_codigo"
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
         BeginProperty Column02 
            DataField       =   "correl_doc"
            Caption         =   "Recibo.Tes."
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
            DataField       =   "fecha_recibo"
            Caption         =   "Fecha.Recibo"
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
            DataField       =   "total_bs"
            Caption         =   "Total.Bs."
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "CI_Entrega"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "CI_Recibe"
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
            DataField       =   "total_dol"
            Caption         =   "Total.Dolares"
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
            DataField       =   "fecha_registro"
            Caption         =   "Fecha.Registro"
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
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   75
         Top             =   3840
         Width           =   6345
         _ExtentX        =   11192
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
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE COBRANZAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2265
      Left            =   1560
      TabIndex        =   26
      Top             =   5100
      Width           =   16935
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "fw_recibos_oficiales.frx":9D37
         Height          =   1980
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   16740
         _ExtentX        =   29528
         _ExtentY        =   3493
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "correl_cobro"
            Caption         =   "Correl"
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
            DataField       =   "cmpbte_deposito"
            Caption         =   "Cheq/Cmpbte."
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
            DataField       =   "cmpbte_fecha"
            Caption         =   "Fecha.Cmpbte."
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
            DataField       =   "cta_codigo"
            Caption         =   "Cta.Banco/Caja"
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
            DataField       =   "cobranza_nro_factura"
            Caption         =   "#Factura"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   3960
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   4305.26
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdOK0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "TODOS"
         Height          =   315
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   240
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame FraBuscaDet 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   107
         Top             =   180
         Visible         =   0   'False
         Width           =   16455
         Begin VB.CommandButton CmdOK3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "OK"
            Height          =   315
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   80
            Width           =   405
         End
         Begin VB.CommandButton CmdOK2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "OK"
            Height          =   315
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   80
            Width           =   405
         End
         Begin VB.CommandButton CmdOK1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "OK"
            Height          =   315
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   109
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   80
            Width           =   405
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "fw_recibos_oficiales.frx":9D51
            DataField       =   "correl_cobro"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   0
            TabIndex        =   108
            Top             =   80
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "correl_cobro"
            BoundColumn     =   "correl_cobro"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_fecha7 
            Bindings        =   "fw_recibos_oficiales.frx":9D6A
            DataField       =   "correl_cobro"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   1080
            TabIndex        =   110
            Top             =   80
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "cobranza_fecha"
            BoundColumn     =   "correl_cobro"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_recibo7 
            Bindings        =   "fw_recibos_oficiales.frx":9D83
            DataField       =   "correl_cobro"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   2160
            TabIndex        =   111
            Top             =   80
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "doc_numero"
            BoundColumn     =   "correl_cobro"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_edificio7 
            Bindings        =   "fw_recibos_oficiales.frx":9D9C
            DataField       =   "correl_cobro"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   14520
            TabIndex        =   113
            Top             =   80
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "edif_codigo_corto"
            BoundColumn     =   "correl_cobro"
            Text            =   "0"
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todas las cobranzas"
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
         TabIndex        =   106
         Top             =   2040
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todas las cobranzas"
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
         TabIndex        =   105
         Top             =   2040
         Visible         =   0   'False
         Width           =   2355
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   0
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6720
      Top             =   8760
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
      Left            =   2160
      Top             =   8760
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
      Left            =   11280
      Top             =   9720
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
      Left            =   9000
      Top             =   9120
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
      Left            =   0
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
   Begin MSAdodcLib.Adodc Ado_datos16 
      Height          =   330
      Left            =   13560
      Top             =   9120
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
      Caption         =   "Ado_datos16"
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
      Left            =   6720
      Top             =   9120
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   11280
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2160
      Top             =   9120
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
      Left            =   4440
      Top             =   9120
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
      Left            =   13560
      Top             =   8760
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
      Left            =   4440
      Top             =   8760
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
      Top             =   8760
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
   Begin MSAdodcLib.Adodc ado_datos7 
      Height          =   330
      Left            =   9000
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
      Caption         =   "ado_datos7"
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
      Left            =   480
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos20 
      Height          =   330
      Left            =   -120
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
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   2160
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
      Caption         =   "Ado_datos21"
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
   Begin MSAdodcLib.Adodc Ado_datos22 
      Height          =   330
      Left            =   4440
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
      Caption         =   "Ado_datos22"
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
   Begin MSAdodcLib.Adodc AdoAux9 
      Height          =   330
      Left            =   6720
      Top             =   9480
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
      Caption         =   "AdoAux9"
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
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "fw_recibos_oficiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Ventas
Dim rs_datos As New ADODB.Recordset     'VENTAS
Dim rs_datos1 As New ADODB.Recordset    'UNIDAD EJECUTORA
Dim rs_datos2 As New ADODB.Recordset    'Beneficiario Personas Nat. y Juridicas (menos de CGI)
Dim rs_datos3 As New ADODB.Recordset    'Proyecto de Edificacion
Dim rs_datos4 As New ADODB.Recordset    'Beneficiario Funcionario de CGI (Vendedor, Cobrador, Admin, etc.)
Dim rs_datos5 As New ADODB.Recordset    'rv_unidad_vs_responsable
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset    'Auxiliar Origen Cobranzas por Cobrador
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset
'Dim rs_datos16 As New ADODB.Recordset   'Libre
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset

Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
Dim rs_datos20 As New ADODB.Recordset
Dim rs_datos21 As New ADODB.Recordset
Dim rs_datos22 As New ADODB.Recordset

'AUXILIARES
Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset
Dim rs_precio As New ADODB.Recordset

Dim rsNada As ADODB.Recordset
'OTROS
'Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir, accion As String
'Dim queryinicial As String
Dim queryinicial2 As String

'Almacenes
Dim descri_bien As String
Dim VAR_OPCION As String
Dim VAR_FECHA As String
Dim VAR_DOC As String
Dim VAR_DA As String
Dim VAR_EDIF As String
Dim VAR_ORIGEN As String
Dim VAR_DOCI, VAR_DOCR, VAR_DOCH, VAR_DOCA As String
Dim VAR_BENI, VAR_BENR, VAR_BENH, VAR_BENA As String
Dim VAR_BENDI, VAR_BENDR, VAR_BENDH, VAR_BENDA As String
Dim VAR_NUMI, VAR_NUMR, VAR_NUMH, VAR_NUMA As String
Dim VAR_DET As String

Dim Cant_Alm, VAR_CANT As Integer
Dim correlativo1, VAR_RECIBO As Integer
Dim VAR_RECIBOS As Integer

'VARIABLES
Dim marca1 As Variant

Dim swgrabar, swnuevo, deta2, CONT_MED As Integer
Dim nroventa, correlv, correldet2, corrprog As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SOL, CANTOT, var_cod5 As Integer
Dim CONT2, CONT3, CONT4, VAR_TIPO As Integer
Dim fdia, fmes, fanio, Dias_Mes, TimeD  As Integer
Dim VAR_COBR1, VAR_COBR2, VAR_CONTR As Integer
Dim VAR_NUM, var_cod, VAR_COD2 As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double

Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_GLOSA2, VAR_MONEDA As String
Dim VAR_BEND, VAR_EDIFD, VARG_ORGD, VAR_CTAD, VAR_UNID, VAR_DPTO, VAR_DPTOD As String
Dim VAR_COD1, VAR_BIEN2, VAR_COD3, VAR_COD4 As String
Dim VAR_MED, VAR_MED2 As String
Dim VAR_TIPOV, VAR_VAL As String
Dim VAR_FEC2, MControl, VAR_MES2 As String
Dim VAR_BEN2, VAR_BEN3, VAR_ALM As String
Dim VAR_BIEN, VAR_R As String
Dim VAR_N1, VAR_N2, VAR_N3, VAR_POA As String

Dim FInicio, FFin, FControl, FVenta As Date
Dim precio_tot, precio_uni As Double


Private Sub CmdDetalle_Click()
'    FrmCobranza.Visible = True
End Sub

'Private Sub adosalalm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
'        Select Case pRecordset.EditMode
'        Case adEditNone
'            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
'            Set DataGrid2.DataSource = Nothing
'            Set DataGrid2.DataSource = rstdetsalalm
'            DataGrid2.ReBind
'        End Select
'End Sub

Private Sub Adodetallesolicitud_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not adoDetalleSolicitud.Recordset.BOF) And (Not adoDetalleSolicitud.Recordset.EOF) Then
        If Not IsNull(adoDetalleSolicitud.Recordset("correlativo_solicitud")) Then
            txtnosolicitud1.Text = adoDetalleSolicitud.Recordset("correlativo_solicitud")
            txtcorrdet.Text = adoDetalleSolicitud.Recordset("correlativo_detalle")
        Else
            txtnosolicitud1.Text = Ado_datos.Recordset("codigo_solicitud")
            txtcorrdet.Text = " "
            dtccodpar.Text = " "
            dtcdescripar.Text = " "
            txtsolpeso.Text = 0
        End If
    End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Dim descri_bien As String
  Dim Cant_Alm As Integer
  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then

'    If parametro <> Ado_datos.Recordset!unidad_codigo Then
'        BtnAnlDetalle.Visible = False
'    Else
'        BtnAnlDetalle.Visible = True
'    End If
    If Not IsNull(Ado_datos.Recordset!IdRecibo) Then
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
        CmdOK100.Visible = True
        If (Ado_datos.Recordset!estado_codigo = "REG") Then
            BtnAprobar.Visible = True
            BtnDesAprobar.Visible = False
            BtnModificar.Visible = True
'            BtnEliminar.Visible = True
            lbl_cerrado.Caption = ""
            FrmABMDet.Visible = True
            'DtGLista.Enabled = False
            BtnAddDetalle.Visible = True
            BtnAnlDetalle.Visible = False
            FraBuscaDet.Visible = False
            CmdOK100.Visible = True
        Else
            BtnAprobar.Visible = False
            BtnDesAprobar.Visible = True
            BtnModificar.Visible = False
'            BtnEliminar.Visible = False
            FrmABMDet.Visible = False
            BtnAddDetalle.Visible = False
            BtnAnlDetalle.Visible = False
            FraBuscaDet.Visible = False
            CmdOK100.Visible = False
        End If
        
        Call AbrirDetalle
        
        'FrmDetalle.Caption = "ORIGEN - Detalle de Cobrnzas Pendientes NRO. " + Str((IIf(IsNull(Ado_datos.Recordset!correl_doc), 0, Ado_datos.Recordset!correl_doc)))
        Option2.Caption = dtc_desc4.Text
        FrmDetalle.Caption = "ORIGEN - Detalle de Registros realizados por COBRANZAS"
        FrmDetalle2.Caption = "DESTINO - Detalle del Recibo Nro. " + Str((IIf(IsNull(Ado_datos.Recordset!correl_doc), 0, Ado_datos.Recordset!correl_doc)))
        
    End If
        FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
  Else
        FrmABMDet.Visible = False
        FrmDetalle.Visible = False
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
  End If
'        BtnEliminar.Visible = True
End Sub

Private Sub AbrirDetalle()
    'ORIGEN - DETALLE DE COBRANZAS
    If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        deta2 = 1
        DtGLista.Visible = True
'        FraBuscaDet.Visible = True
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        Select Case VAR_OPCION
            Case "0"
                'rs_datos14.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
                queryinicial2 = "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   "
            Case "1"
                'rs_datos14.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND  cobranza_fecha = '" & VAR_FECHA & "' order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
                queryinicial2 = "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND  cobranza_fecha = '" & VAR_FECHA & "'   "
            Case "2"
                'rs_datos14.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND doc_numero = " & VAR_RECIBOS & "  order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
                queryinicial2 = "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND doc_numero = " & VAR_RECIBOS & "  "
            Case "3"
                'rs_datos14.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND edif_codigo_corto= '" & VAR_EDIF & "'  order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
                queryinicial2 = "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND edif_codigo_corto= '" & VAR_EDIF & "'   "
            Case Else
                'rs_datos14.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
                queryinicial2 = "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   "
        End Select
        rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        rs_datos14.Sort = "cobranza_fecha desc"
        Set ado_datos14.Recordset = rs_datos14.DataSource
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            DtGLista.Visible = True
'            FraBuscaDet.Visible = True
            Set DtGLista.DataSource = ado_datos14.Recordset
'            If VAR_OPCION = 0 Then
'                Call AbreOrigen
'            End If
        Else
            deta2 = 0
            DtGLista.Visible = False
'            FraBuscaDet.Visible = False
        End If
    
    Else
        deta2 = 0
        DtGLista.Visible = False
'        FraBuscaDet.Visible = False
    End If
    
    'DESTINO - DETALLE DEL RECIBO
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        'rs_datos11.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo_tes = 'APR' and estado_codigo_cont = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
        rs_datos11.Open "select * from fv_ventas_cobranza_det_recibos where beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND idRecibo = " & Ado_datos.Recordset!IdRecibo & " order by doc_numero ", db, adOpenKeyset, adLockOptimistic
    Else
        rs_datos11.Open "select * from fv_ventas_cobranza_det_recibos where idRecibo = " & Ado_datos.Recordset!IdRecibo & " order by doc_numero ", db, adOpenKeyset, adLockOptimistic
        'beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND
    End If
    'rs_datos11.Sort = "doc_numero "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
        'Call AbreAlmacen
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If

End Sub

Private Sub AbreOrigen()
    'ORIGEN - DETALLE DE COBRANZAS
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
    '    rs_datos7.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
    'Else
        rs_datos7.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic
    'End If
    'rs_datos7.Sort = "cobranza_fecha desc"
    Set Ado_datos7.Recordset = rs_datos7.DataSource
    'ado_datos7.Recordset.Requery
    'dtc_desc7.BoundText = dtc_codigo7.BoundText
    dtc_fecha7.BoundText = dtc_codigo7.BoundText
    dtc_recibo7.BoundText = dtc_codigo7.BoundText
    dtc_edificio7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub Ado_datos11_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not Ado_datos11.Recordset.BOF) And (Not Ado_datos11.Recordset.EOF) Then
        If Ado_datos11.Recordset.RecordCount = 0 Then
            BtnAnlDetalle.Visible = False
        Else
            BtnAnlDetalle.Visible = True
        End If
    End If
End Sub

Private Sub Ado_datos14_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 If (Not ado_datos14.Recordset.BOF) And (Not ado_datos14.Recordset.EOF) Then
    If ado_datos14.Recordset.RecordCount > 0 Then
        'BtnModDetalle2.Visible = False
    Else
    
    End If
 Else
 
 End If
End Sub

Private Sub BtnAddDetalle_Click()
On Error GoTo UpdateErr
If glusuario = "ASANTIVAÑEZ" Or glusuario = "TCASTILLO" Or glusuario = "LMORALES" Or glusuario = "RGIL" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "SPAREDES" Or glusuario = "PLOPEZ" Or glusuario = "MCOARITY" Or glusuario = "VPAREDES" Or glusuario = "EVILLALOBOS" Or glusuario = "MVALDIVIA" Or glusuario = "CSALINAS" Then        '
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        If ado_datos14.Recordset.RecordCount > 0 Then         '<> "" Then
            'REGISTROS CERRADOS QUE NO SE PUEDEN APROBAR
             If (ado_datos14.Recordset!trans_codigo = "F" Or ado_datos14.Recordset!trans_codigo = "T" Or ado_datos14.Recordset!trans_codigo = "O") Then
                If CDate(ado_datos14.Recordset!cmpbte_fecha) <= CDate("31/12/2022") Then
                    If glusuario = "ADMIN" Or glusuario = "PLOPEZ" Then
                    Else
                        MsgBox "No se puede ACEPTAR una cobranza con fecha de Comprobante menor al 31-DICIEMBRE-2022, porque se encuentra CERRADA, consulte con Contabilidad ... ", , "Atención"
                        Exit Sub
                    End If
                End If
            End If
            If (ado_datos14.Recordset!trans_codigo <> "E") And (IsNull(ado_datos14.Recordset!cmpbte_fecha) Or (ado_datos14.Recordset!cmpbte_fecha = "01/01/1900")) Then
                MsgBox "No se puede ACEPTAR, verifique la fecha de Cheque, Transferencia o Comprobante y vuelva a intentar ...", , "Atención"
                Exit Sub
            End If
            If (ado_datos14.Recordset!estado_codigo_rbo = "REG") Or (IsNull(ado_datos14.Recordset!estado_codigo_rbo)) Then
                'GRABA RECIBO DETALLE
                db.Execute "INSERT INTO fo_recibos_detalle (IdRecibo, correl_cobro, cta_codigo, cmpbte_deposito, doc_numero, cobranza_bs, cobranza_dol, estado_codigo, usr_codigo, fecha_registro, hora_registro, trans_codigo) " & _
                " values (" & Ado_datos.Recordset!IdRecibo & ", " & ado_datos14.Recordset!correl_cobro & ", '" & ado_datos14.Recordset!cta_codigo & "', '" & ado_datos14.Recordset!cmpbte_deposito & "', " & ado_datos14.Recordset!doc_numero & ", " & ado_datos14.Recordset!cobranza_bs & ", " & ado_datos14.Recordset!cobranza_dol & ",  " & _
                "  'APR', '" & glusuario & "', '" & Date & "', '', '" & ado_datos14.Recordset!trans_codigo & "'  ) "
    
                'ACTUALIZA APRUEBA ao_ventas_cobranza_det
                db.Execute "UPDATE ao_ventas_cobranza_det SET estado_codigo_tes = 'APR'  WHERE cobranza_codigo = " & ado_datos14.Recordset!cobranza_codigo & " and cobranza_detalle = " & ado_datos14.Recordset!cobranza_detalle & " "
                
                ' ACTUALIZA TOTALES fo_recibos_oficiales
                db.Execute "update fo_recibos_oficiales set total_bs = (select sum(fo_recibos_detalle.cobranza_bs) from fo_recibos_detalle where fo_recibos_detalle.IdRecibo = " & Ado_datos.Recordset!IdRecibo & ")   " & _
                " from fo_recibos_oficiales inner join fo_recibos_detalle on  fo_recibos_oficiales.IdRecibo = fo_recibos_detalle.IdRecibo where fo_recibos_oficiales.IdRecibo = " & Ado_datos.Recordset!IdRecibo & " "
            
                Call AbrirDetalle
                CmdOK100.Visible = True
            Else
                MsgBox "El Recibo ya fue ELEGIDO, elija OTRO y vuelva a intentar ...", , "Atención"
            End If
        Else
            MsgBox "Debe elegir un registro cobrado,  vuelva a intentar ...", , "Atención"
        End If
    Else
        MsgBox "El registro ya se encuentra procesado, vuelva a intentar ...", , "Atención"
    End If
 Else
    MsgBox "Debe elegir un registro para procesarlo,  vuelva a intentar ...", , "Atención"
 End If
Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
End If
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnAñadir_Click()
accion = "NEW"
    
On Error GoTo UpdateErr
If glusuario = "ASANTIVAÑEZ" Or glusuario = "TCASTILLO" Or glusuario = "LMORALES" Or glusuario = "RGIL" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "SPAREDES" Or glusuario = "PLOPEZ" Or glusuario = "MCOARITY" Or glusuario = "VPAREDES" Or glusuario = "EVILLALOBOS" Or glusuario = "MVALDIVIA" Or glusuario = "CSALINAS" Then
    'Ado_datos.Recordset.AddNew
    dtc_codigo3.Text = VAR_R
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
    
    'dtc_desc3.backColor = &H80000005
    'dtc_desc3.ForeColor = &H80000008
    
    'txt_campo1.Caption = "0"
    'dtc_desc3.Locked = False
    'dtc_desc3.Width = 5955
    
    DTPfechasol.Value = Date
    swgrabar = 1
    FrmCabecera.Enabled = True
    FrmDetalle.Visible = False
    FraNavega.Enabled = False
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    Fra_datos.Enabled = True

    FrmABMDet.Visible = False
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    dtc_desc4.SetFocus
    Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
    End If
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
 On Error GoTo UpdateErr
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    If glusuario = "ASANTIVAÑEZ" Or glusuario = "TCASTILLO" Or glusuario = "LMORALES" Or glusuario = "RGIL" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "PLOPEZ" Or glusuario = "MCOARITY" Or glusuario = "VPAREDES" Or glusuario = "EVILLALOBOS" Then
    'Or glusuario = "SPAREDES" Or glusuario = "MVALDIVIA" Or glusuario = "CSALINAS"
    
      VAR_RECIBO = Ado_datos.Recordset!IdRecibo
      'Actualiza Totales
      'db.Execute "update fo_recibos_oficiales set fo_recibos_oficiales.total_bs = fv_recibos_detalle_acumulado.totalCobro_Bs,  -.total_dol = fv_recibos_detalle_acumulado.totalCobro_dol from fo_recibos_oficiales inner join fv_recibos_detalle_acumulado on fo_recibos_oficiales.IdRecibo = fv_recibos_detalle_acumulado.IdRecibo WHERE fo_recibos_oficiales.IdRecibo =  " & VAR_RECIBO & "  "
      db.Execute "update fo_recibos_oficiales set fo_recibos_oficiales.total_bs = fv_recibos_detalle_acumulado.totalCobro_bs, fo_recibos_oficiales.total_dol = fv_recibos_detalle_acumulado.totalCobro_dol from fo_recibos_oficiales inner join fv_recibos_detalle_acumulado on fo_recibos_oficiales.IdRecibo = fv_recibos_detalle_acumulado.IdRecibo WHERE fo_recibos_oficiales.IdRecibo =  " & VAR_RECIBO & ""
      'Actualiza Detalle
      db.Execute "update fo_recibos_detalle set estado_codigo = 'APR', estado_destino = 'APR', estado_aprueba='REG'  WHERE fo_recibos_detalle.IdRecibo =  " & VAR_RECIBO & "  "
      'APRUEBA ao_ventas_cobranza_det
      db.Execute "update ao_ventas_cobranza_det set ao_ventas_cobranza_det.estado_codigo = 'APR', ao_ventas_cobranza_det.estado_codigo_tes = 'APR', ao_ventas_cobranza_det.estado_codigo_cont = 'REG' from ao_ventas_cobranza_det inner join fo_recibos_detalle on ao_ventas_cobranza_det.correl_cobro = fo_recibos_detalle.correl_cobro WHERE fo_recibos_detalle.IdRecibo =  " & VAR_RECIBO & "  "
      'APRUEBA fo_recibos_oficiales
      db.Execute "update fo_recibos_oficiales set estado_codigo = 'APR' where IdRecibo = " & VAR_RECIBO & " "
      
      OptFilGral2_Click
      
      If (dg_datos.SelBookmarks.Count <> 0) Then
          dg_datos.SelBookmarks.Remove 0
      End If
      If Ado_datos.Recordset.RecordCount > 0 Then
          rs_datos.Find "IdRecibo = " & VAR_RECIBO & "   ", , , 1
          dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
          ' If rs_det1.RecordCount > 0 Then
          ' rs_det1.MoveLast
          'End If
      Else
          rs_datos.MoveLast
      End If
    Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
    End If
  Else
      MsgBox "No se puede aprobar el registro actual"
  End If

Exit Sub
UpdateErr:
MsgBox Err.Description

End Sub

Private Sub GENERA_COMPRA()
'    If rs_datos!estado_cotiza = "REG" Then
'      VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'      VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'        ' MANTENIMIENTO PREVENTIVO - INSUMOS y/o COMPRAS BB y SS
'                'EQUIPO
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & parametro & "'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux2.RecordCount > 0 Then
'                       rs_aux2!correl_negocia = rs_aux2!correl_negocia + 1
'                       correldetalle = rs_aux2!correl_negocia
'                       rs_aux2.Update
'                    End If
'                    'WWWWWWWWWWWWWWW
'                    'correlv = Ado_datos.Recordset!venta_codigo
'                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux3.RecordCount = 0 Then
'                    'beneficiario_codigo_resp,'doc_numero,estado_codigo_tra, estado_codigo_nac, estado_codigo_des, hora_registro, usr_codigo_aprueba,'                      fecha_registro_aprueba
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion     'Year(Date)
'                        'rs_aux3!compra_codigo = 0      'Autonumerico
'                        rs_aux3!unidad_codigo_adm = parametro
'                        rs_aux3!solicitud_codigo_adm = correldetalle
'                        rs_aux3!unidad_codigo = VAR_COD4
'                        rs_aux3!solicitud_codigo = VAR_SOL
'                        rs_aux3!edif_codigo = VAR_PROY2
'                        rs_aux3!beneficiario_codigo = VAR_BENEF
'                        rs_aux3!solicitud_tipo = Ado_datos.Recordset!solicitud_tipo       '"10"
'                        rs_aux3!venta_tipo = "E"
'                        rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
'                        rs_aux3!compra_fecha = Date
'                        rs_aux3!compra_descripcion = "COMPRA POR: " + lbl_titulo.Caption
'                        rs_aux3!compra_observaciones = "Edificio: " + Trim(dtc_desc3.Text)
'                        rs_aux3!compra_cantidad_total = 1   'Ado_datos.Recordset!venta_cantidad_total
'                        rs_aux3!compra_monto_bs = 0     'VAR_BS2
'                        rs_aux3!tipo_moneda = "BOB"
'                        rs_aux3!compra_monto_dol = 0        'VAR_DOL2
'                        rs_aux3!proceso_codigo = "TEC"
'                        rs_aux3!subproceso_codigo = "TEC-06"
'                        rs_aux3!etapa_codigo = "TEC-06-01"
'                        rs_aux3!clasif_codigo = "ADM"
'                        rs_aux3!doc_codigo = "R-114"
'                        rs_aux3!poa_codigo = "3.2.8"
'                        rs_aux3!estado_codigo_eqp = "REG"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3!fecha_registro = Date
'                        rs_aux3.Update
'
'                        'DETALLE Carga ao_ventas_detalle
'                        Set rstdestino = New ADODB.Recordset
'                        If rstdestino.State = 1 Then rstdestino.Close
'                        rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rstdestino.RecordCount > 0 Then
'                        End If
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        'rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & rs_aux3!compra_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
'                        rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & VAR_SOL & "  and grupo_codigo = '30000' ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                            VAR_REG = 1
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                              If rs_aux4!grupo_codigo = "30000" Then
'                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigoR , usr_usuario, fecha_registro) " & _
'                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
'
'                                db.Execute "Update ao_compra_detalle SET ao_compra_detalle.compra_concepto  = ac_bienes.bien_descripcion From ao_compra_detalle INNER JOIN ac_bienes ON ao_compra_detalle.bien_codigo = ac_bienes.bien_codigo where ao_compra_detalle.compra_codigo = " & rs_aux3!compra_codigo & " and ao_compra_detalle.bien_codigo = '" & rs_aux4!bien_codigo & "' "
'                                VAR_REG = VAR_REG + 1
'                              End If
'                               rs_aux4.MoveNext
'                           Wend
'                        End If
'                        If rstdestino.State = 1 Then rstdestino.Close
'                    End If
'                    'WWWWWWWWWW
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            Txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = Txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        rs_datos!estado_cotiza = "APR"
'        rs_datos!fecha_aprueba = Date
'        rs_datos!usr_codigo_aprueba = glusuario
'        rs_datos.UpdateBatch adAffectAll
'      End If
'
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
End Sub



Private Sub BtnBuscar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      buscados = 1
      PosibleApliqueFiltro = False
      
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
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    OptFilGral1.Visible = True
    OptFilGral2.Visible = True
  End If
End Sub

Private Sub BtnCancelar_Click()
On Error GoTo UpdateErr
  If swgrabar = 2 Then
    var_cod5 = Ado_datos.Recordset!venta_codigo
  End If
  'Ado_datos.Refresh
  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  'marca1 = Ado_datos.Recordset.Bookmark
  FraNavega.Enabled = True
  FrmCabecera.Enabled = False
  Fra_datos.Enabled = True
  FrmDetalle.Visible = True

'  Fra_Total.Visible = True
  dg_datos.Visible = True
  FrmABMDet.Visible = True
  dtc_desc3.backColor = &H80000008
  dtc_desc3.ForeColor = &H80000005
  
  'Refrescar Grid
  If OptFilGral1.Value = True Then
       Call OptFilGral1_Click        'Pendientes
  Else
       Call OptFilGral2_Click        'TODOS
  End If
  If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
  End If
  If Ado_datos.Recordset.RecordCount > 0 And swgrabar = 2 Then
       rs_datos.Find "venta_codigo = " & var_cod5 & "   ", , , 1
       dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
  Else
       rs_datos.MoveLast
  End If
  swgrabar = 0
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True
  accion = ""
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnCancelar3_Click()
        Fra_reporte.Visible = False
End Sub

Private Sub btnEliminar_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    If Ado_datos.Recordset("estado_almacen") = "REG" Then
'      sino = MsgBox("Esta seguro de ANULAR la venta registrada ?", vbYesNo, "Confirmando")
'      If sino = vbYes Then
'          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'ANL' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  "
'          'Dim rstdestino As New ADODB.Recordset
'          'Set rstdestino = New ADODB.Recordset
'          'If rstdestino.State = 1 Then rstdestino.Close
'          'rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  ", db, adOpenDynamic, adLockOptimistic
'          'If Not rstdestino.BOF Then rstdestino.MoveFirst
'          'If Not rstdestino.BOF And Not rstdestino.EOF Then
'          '    rstdestino("estado_codigo") = "E"
'          '    rstdestino.Update
'          'End If
'          'If rstdestino.State = 1 Then rstdestino.Close
'          marca1 = Ado_datos.Recordset.Bookmark
'          'Ado_datos.Recordset.Requery
'          'Ado_datos.Refresh
'          Call OptFilGral1_Click
'          Ado_datos.Recordset.Move marca1 - 1
'      End If
'    Else
'      MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
'    End If
'  Else
'    MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
     If rs_datos!estado_almacen = "REG" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
'     If ExisteReg(Ado_datos.Recordset!unidad_codigo_sol, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
          rs_datos!estado_almacen = "ANL"
'          rs_datos!fecha_registro = Date
'          rs_datos!usr_codigo = glusuario
'           Ado_datos.Recordset.Requery
'           Ado_datos.Refresh
           db.Execute "ap_ventas_grla 1 ,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!edif_codigo & "'," & Ado_datos.Recordset!venta_codigo & ",'" & Ado_datos.Recordset!beneficiario_codigo_alm & "','" & Ado_datos.Recordset!fecha_verif & "'," & ado_datos14.Recordset!bien_cantidad_por_empaque & "," & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "'," & precio_uni & ""
           Call AbrirDetalle
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

Private Sub BtnGrabar_Click()
'    ' CIERRE TEMPORAL DE COBRANZAS GESTION 2021
'    If CDate(DTPfechasol.Value) >= CDate("01/01/2021") And CDate(DTPfechasol.Value) <= CDate("31/12/2021") Then
'        If Ado_datos.Recordset!unidad_codigo = "DVTA" Or Ado_datos.Recordset!unidad_codigo = "DCOMS" Or Ado_datos.Recordset!unidad_codigo = "DCOMB" Or Ado_datos.Recordset!unidad_codigo = "DCOMC" Then
'            MsgBox "El registro para la Gestión 2021, será CERRADO el 31-mar-2022, consulte con Contabilidad ... ", , "Atención"
'        Else
'            MsgBox "No se puede Registrar un Recibo de Tesorería con fecha de la Gestión 2021, esta se encuentra CERRADA, consulte con Contabilidad ... ", , "Atención"
'            Exit Sub
'        End If
'    End If

On Error GoTo UpdateErr
  VAR_VAL = "OK"

  Call valida_campos
  If VAR_VAL = "OK" Then
    If swgrabar = 2 Then
        var_cod5 = Ado_datos.Recordset!IdRecibo
        'FInicio = IIf(IsNull(Ado_datos.Recordset!venta_fecha_inicio), Date, Ado_datos.Recordset!venta_fecha_inicio)
        'CANTOT = IIf(IsNull(Ado_datos.Recordset!venta_cantidad_total), 1, Ado_datos.Recordset!venta_cantidad_total)
        'gestion0 = IIf(IsNull(Ado_datos.Recordset!ges_gestion), glGestion, Ado_datos.Recordset!ges_gestion)
        VAR_BENEF = IIf(IsNull(Ado_datos.Recordset!beneficiario_codigo), "0", Ado_datos.Recordset!beneficiario_codigo)
        corrprog = Ado_datos.Recordset!correl_doc
        'VAR_MED = Ado_datos.Recordset!unimed_codigo
        'VAR_UNI = Ado_datos.Recordset!unidad_codigo
        'FControl = IIf(IsNull(Ado_datos.Recordset!fecha_verif), Date, Ado_datos.Recordset!fecha_verif)
        'Ado_datos.Recordset("fecha_verif") = DTPfechasol.Value
        '        rs_datos!fecha_verif = Date
    End If
    FrmCabecera.Enabled = False
    Call grabar
    '
    'db.Execute "update ao_almacen_salidas set concepto = '" & TxtConcepto.Text & "' WHERE venta_codigo = " & var_cod5
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    FrmCabecera.Enabled = False
    Fra_datos.Enabled = True
    dg_datos.Visible = True
    FrmDetalle.Visible = True
    dtc_desc3.backColor = &H80000008
    dtc_desc3.ForeColor = &H80000005
'    Fra_Total.Visible = True
    FrmABMDet.Visible = True
    'Refrescar Grid
    If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     'var_cod5 = Ado_datos.Recordset!IdRecibo       'Or swgrabar = 1
     If Ado_datos.Recordset.RecordCount > 0 And swgrabar = 2 Then
        rs_datos.Find "IdRecibo = " & var_cod5 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
     swgrabar = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
  End If
    accion = ""
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub valida_campos()

'  If dtc_codigo2 = "" Then
'    MsgBox "Debe Elejir La Unidad Destino, Vuelva a Intentar ...", vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  
  If dtc_codigo4 = "" Then
    MsgBox "Debe Elejir Responsable de la entrega ORIGEN, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If dtc_codigo11 = "" Then
'    MsgBox "Debe Elejir el Almacen!! , Vuelva a Intentar ...", vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If dtc_codigo5 = "" Then
    MsgBox "Debe Elejir ... Entregado a:, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3 = "" Then
    MsgBox "Debe Registrar el Documento ISO, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If dtc_codigo21 = "" Then
'    MsgBox "Debe Elejir Regional ORIGEN, Vuelva a Intentar ...", vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub BtnImprimir_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        VAR_RECIBO = Ado_datos.Recordset!IdRecibo
        CryV01.Reset
        CryV01.WindowState = crptMaximized
        CryV01.WindowShowSearchBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.WindowShowPrintSetupBtn = True
        
        Dim iResult As Integer
        If Ado_datos11.Recordset.RecordCount > 10 Then
            CryV01.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_recibos_oficiales_tesoreria_pag1.rpt"
        Else
            CryV01.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_recibos_oficiales_tesoreria.rpt"
        End If
            var_titulo = "RECIBO INGRESOS TESORERIA"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Ado_datos.Recordset!IdRecibo
        'CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
        CryV01.Formulas(0) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(1) = "subtitulo = 'DETALLE DE COBRANZAS' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
End Sub

Private Sub BtnImprimir1_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        CryV01.Reset
        CryV01.WindowState = crptMaximized
        CryV01.WindowShowSearchBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.WindowShowPrintSetupBtn = True
        
        Dim iResult As Integer
            CryV01.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_recibos_oficiales_tesoreria.rpt"
            var_titulo = "RECIBO DE TESORERIA"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Ado_datos.Recordset!IdRecibo
        'CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
        CryV01.Formulas(0) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(1) = "subtitulo = 'DETALLE DE COBRNZAS' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
End Sub

Private Sub BtnModificar_Click()
On Error GoTo UpdateErr
If glusuario = "ASANTIVAÑEZ" Or glusuario = "TCASTILLO" Or glusuario = "LMORALES" Or glusuario = "RGIL" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "SPAREDES" Or glusuario = "PLOPEZ" Or glusuario = "MCOARITY" Or glusuario = "VPAREDES" Or glusuario = "EVILLALOBOS" Or glusuario = "MVALDIVIA" Or glusuario = "CSALINAS" Then
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
        accion = "MOD"
        FrmCabecera.Enabled = True
        FrmDetalle.Visible = False
        FraNavega.Enabled = False
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        'If dtc_desc4.Text = "" Or dtc_desc11.Text = "" Or dtc_desc21.Text = "" Then
            Fra_datos.Enabled = True
        'Else
        '    Fra_datos.Enabled = False
        'End If
'        Fra_Total.Visible = False
        FrmABMDet.Visible = False
        swgrabar = 2
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        dtc_desc4.SetFocus
        
    Else
      MsgBox "NO se puede MODIFICAR, porque el registro ya fue Aprobado, Anulado o Cerrado.", , "Atencion"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
End If
  Exit Sub
UpdateErr:
    MsgBox Err.Description
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
  FrmCobros.Enabled = False
  'swgrabar = 0
  'Call cerea
  swnuevo = 0
  If Ado_datos.Recordset("estado_codigo") = "REG" Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
'    FrmABMDet2.Visible = True
End Sub

Private Sub CmdCancelaDet_Click()
  'TxtNroVenta.Enabled = True
  FrmEdita.Enabled = False
  swgrabar = 0
  swnuevo = 0
  'cmdElige.Enabled = False
  marca1 = Ado_datos.Recordset.Bookmark
'  If Ado_datos.Recordset("estado_almacen") = "REG" Then
'    Call OptFilGral1_Click
'  Else
'    Call OptFilGral2_Click
'  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    FraNavega.Enabled = True
    FrmDetalle.Enabled = True
    FrmABMDet.Visible = True
    
'     Call AbrirDetalle
  ado_datos14.Recordset.Cancel
  Call AbrirDetalle
  'Ado_datos.Recordset.Move marca1 - 1
  accion = ""
End Sub

Private Sub BtnDesAprobar_Click()
 On Error GoTo UpdateErr
  If Ado_datos.Recordset!estado_codigo = "APR" Then
    sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        VAR_RECIBO = Ado_datos.Recordset!IdRecibo
        '    VERIFICA SI YA SE REGISTRO EN TESORERIA        'Count(*) AS Cuantos    '(dbo.fo_recibos_detalle.IdTraspasoBancos <> 0 OR dbo.fo_recibos_detalle.IdTraspasoBancos IS NOT NULL) AND (dbo.fo_recibos_detalle.IdRecibo = 288)
        Set rs_datos1 = New ADODB.Recordset
        If rs_datos1.State = 1 Then rs_datos1.Close
        'rs_datos1.Open "Select * from fv_ventas_cobranza_det_traspasos where (IdRecibo = " & VAR_RECIBO & " AND (IdTraspasoBancos <> 0 OR IdTraspasoBancos IS NOT NULL)) ", db, adOpenStatic
        rs_datos1.Open "Select * from fv_ventas_cobranza_det_traspasos where (IdRecibo = " & VAR_RECIBO & " AND (IdTraspasoBancos = 0 OR IdTraspasoBancos IS NULL)) ", db, adOpenStatic
        If rs_datos1.RecordCount > 0 Then
            'MsgBox "- No se puede DESAPROBAR, porque el registro seleccionado está siendo utilizado en un TRASPASO CUENTAS BANCARIAS... " & vbCrLf & " - Retorne el registro, desde el DESTINO al ORIGEN DEL TRASPASO ...", vbExclamation, "Atención!"
            'Exit Sub
        'Else
            'Actualiza Totales
            'db.Execute "update fo_recibos_oficiales set fo_recibos_oficiales.total_bs = fv_recibos_detalle_acumulado.totalCobro_bs, fo_recibos_oficiales.total_dol = fv_recibos_detalle_acumulado.totalCobro_dol from fo_recibos_oficiales inner join fv_recibos_detalle_acumulado on fo_recibos_oficiales.IdRecibo = fv_recibos_detalle_acumulado.IdRecibo WHERE fo_recibos_oficiales.IdRecibo =  " & VAR_RECIBO & ""
            db.Execute "update fo_recibos_oficiales set fo_recibos_oficiales.total_bs = '0', fo_recibos_oficiales.total_dol = '0' WHERE fo_recibos_oficiales.IdRecibo =  " & VAR_RECIBO & ""
            'Actualiza Detalle
            'db.Execute "update fo_recibos_detalle set estado_codigo = 'APR', estado_destino = 'APR', estado_aprueba='REG'  WHERE fo_recibos_detalle.IdRecibo =  " & VAR_RECIBO & "  "
            db.Execute "update fo_recibos_detalle set estado_codigo = 'REG', estado_destino = 'REG', estado_aprueba='REG'  WHERE fo_recibos_detalle.IdRecibo =  " & VAR_RECIBO & "  "
            'APRUEBA ao_ventas_cobranza_det
            'db.Execute "update ao_ventas_cobranza_det set ao_ventas_cobranza_det.estado_codigo = 'APR', ao_ventas_cobranza_det.estado_codigo_tes = 'APR' from ao_ventas_cobranza_det inner join fo_recibos_detalle on ao_ventas_cobranza_det.correl_cobro = fo_recibos_detalle.correl_cobro WHERE fo_recibos_detalle.IdRecibo =  " & VAR_RECIBO & "  "
            'db.Execute "update ao_ventas_cobranza_det set ao_ventas_cobranza_det.estado_codigo = 'REG', ao_ventas_cobranza_det.estado_codigo_tes = 'REG' from ao_ventas_cobranza_det inner join fo_recibos_detalle on ao_ventas_cobranza_det.correl_cobro = fo_recibos_detalle.correl_cobro WHERE fo_recibos_detalle.IdRecibo =  " & VAR_RECIBO & "  "
            'APRUEBA fo_recibos_oficiales
            'db.Execute "update fo_recibos_oficiales set estado_codigo = 'APR' where IdRecibo = " & VAR_RECIBO & " "
            db.Execute "update fo_recibos_oficiales set estado_codigo = 'REG' where IdRecibo = " & VAR_RECIBO & " "
            
            OptFilGral2_Click
            
            If (dg_datos.SelBookmarks.Count <> 0) Then
                dg_datos.SelBookmarks.Remove 0
            End If
            If Ado_datos.Recordset.RecordCount > 0 Then
                rs_datos.Find "IdRecibo = " & VAR_RECIBO & "   ", , , 1
                dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
                ' If rs_det1.RecordCount > 0 Then
                ' rs_det1.MoveLast
                'End If
            Else
                rs_datos.MoveLast
            End If
        Else
            MsgBox "- No se puede DESAPROBAR, porque el registro seleccionado está siendo utilizado en un TRASPASO CUENTAS BANCARIAS... " & vbCrLf & " - Retorne el registro, desde el DESTINO al ORIGEN DEL TRASPASO ...", vbExclamation, "Atención!"
            Exit Sub
        End If
    Else
      MsgBox "No se puede aprobar el registro actual"
    End If
   Else
   
   End If
Exit Sub
UpdateErr:
MsgBox Err.Description

'
'On Error GoTo UpdateErr
'  If Ado_datos.Recordset.RecordCount > 0 Then
'     If rs_datos!estado_almacen = "APR" Then
'       sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'       If sino = vbYes Then
''     If ExisteReg(Ado_datos.Recordset!unidad_codigo_sol, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
'
''          rs_datos!fecha_registro = Date
''          rs_datos!usr_codigo = glusuario
''           Ado_datos.Recordset.Requery
''           Ado_datos.Refresh
'          db.Execute "ap_ventas_grla 3,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!edif_codigo & "'," & Ado_datos.Recordset!venta_codigo & ",'" & Ado_datos.Recordset!beneficiario_codigo_alm & "','" & Ado_datos.Recordset!fecha_verif & "'," & ado_datos14.Recordset!bien_cantidad_por_empaque & "," & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "'," & precio_uni & ""
'            rs_datos!estado_almacen = "REG"
'           Call AbrirDetalle
'          rs_datos.UpdateBatch adAffectAll
'       End If
'    Else
'       MsgBox "No se puede DESPROBAR un registro Aulado(ANL) o Registrado (REG) ...", vbExclamation, "Validación de Registro"
'    End If
'  Else
'      MsgBox "NO se puede DESAPROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
'  Exit Sub
'
'UpdateErr:
'  MsgBox Err.Description
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

Private Sub Contabiliza_venta()
    Call graba_proyecto
    Call graba_ingreso
  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  'If sino = vbYes Then
    ' INI CORRECCION 18-JUN-2014
    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DEI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
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
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                Exit Sub
              End If
            End If
            If rs_aux1.State = 1 Then rs_aux1.Close

        Case "DYR"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DES"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "ANI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DVI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

            '' 02/07/2014 VERIFICAR
            'If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
            'If rstdestino2.State = 1 Then rstdestino2.Close
            'rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "" Then
'        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
'        Exit Sub
'      End If

    ' INI CORRECCION 18-JUN-2014
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' 02/07/2014 VERIFICAR
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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

    fte_codigo1 = VAR_FTE
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
      If rs_aux2.RecordCount > 0 Then
        yacontabilizo = 1
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
      rs_aux2("ges_gestion") = glGestion    'Year(Date)
      'rstdestino("Num_Respaldo") = Ado_datos.Recordset("numero_documento")
      If yacontabilizo = 0 Then
        rs_aux2("Fecha_transacion") = Date
      End If
      rs_aux2("beneficiario_codigo") = VAR_BENEF
      rs_aux2("glosa") = VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4       'Ado_datos.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = Ado_datos.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE

      rs_aux2("proceso_codigo") = "FIN"
      rs_aux2("subproceso_codigo") = "FIN-02"
      Select Case VAR_CODTIPO
        Case "DEI"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "REC"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "DYR"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "DES"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "ANI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "DVI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
      End Select

      rs_aux2("clasif_codigo") = "ADM"
      rs_aux2("doc_codigo") = "R-640"
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
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then     'Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then     'Devengado y Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then     'Desafectado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then     'Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then     'Desafectado y Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If

'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' VERIFICAR SI SE ESTA CONTROLANDA con el DYR
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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
      '2115
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
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
        d_cta_nombre_1 = rs_aux1("NombreCta")
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = rs_aux1("NombreCta")
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
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = Trim(d_cta_nombre_1) ' CAMPO PARA ELIMINAR
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
                rstdestino2("D_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                rstdestino2("D_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                rstdestino2("D_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = VAR_DPTO
                rstdestino2("D_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                rstdestino2("D_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
        End Select

        Select Case d_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                rstdestino2("D_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                rstdestino2("D_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                rstdestino2("D_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = VAR_DPTO
                rstdestino2("D_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                rstdestino2("D_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
        End Select

        Select Case d_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                rstdestino2("D_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                rstdestino2("D_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                rstdestino2("D_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = VAR_DPTO
                rstdestino2("D_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                rstdestino2("D_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        ' CORREGIR MONTOS JQA 2014-JUL-08
        If j > 1 Then
            If i = 1 Then
                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
            Else
                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
            End If
        Else
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
        End If
        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        'AQUI MONEDA 02/07/01
        'rstdestino2("D_Cambio") = GlTipoCambioMercado
        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
        rstdestino2("H_Cuenta") = cta_credito1
        rstdestino2("H_Nombre") = Trim(h_cta_nombre_1) ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_cred11
        rstdestino2("H_SubCta2") = Subcta_cred21
        rstdestino2("H_Aux1") = h_aux1_1
        rstdestino2("H_Aux2") = h_aux2_1
        rstdestino2("H_Aux3") = h_aux3_1
        'rstdestino2("H_Cta_Aux1") = ""
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
                rstdestino2("H_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                rstdestino2("H_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                rstdestino2("H_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = VAR_DPTO
                rstdestino2("H_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                rstdestino2("H_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
        End Select

        Select Case h_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                rstdestino2("H_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                rstdestino2("H_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                rstdestino2("H_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = VAR_DPTO
                rstdestino2("H_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                rstdestino2("H_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
        End Select

        Select Case h_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                rstdestino2("H_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                rstdestino2("H_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                rstdestino2("H_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = VAR_DPTO
                rstdestino2("H_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                rstdestino2("H_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
        End Select

'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        If j > 1 Then
            If i = 1 Then
                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
            Else
                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
            End If
        Else
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
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
                rstdestino2("D_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                rstdestino2("D_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                rstdestino2("D_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = VAR_DPTO
                rstdestino2("D_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                rstdestino2("D_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
        End Select

        Select Case h_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                rstdestino2("D_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                rstdestino2("D_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                rstdestino2("D_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = VAR_DPTO
                rstdestino2("D_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                rstdestino2("D_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
        End Select

        Select Case h_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                rstdestino2("D_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                rstdestino2("D_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                rstdestino2("D_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = VAR_DPTO
                rstdestino2("D_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                rstdestino2("D_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
        End Select
'        If h_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        If i = 1 Then
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
        Else
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
        End If
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
                rstdestino2("H_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                rstdestino2("H_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                rstdestino2("H_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = VAR_DPTO
                rstdestino2("H_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                rstdestino2("H_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
        End Select

        Select Case d_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                rstdestino2("H_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                rstdestino2("H_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                rstdestino2("H_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = VAR_DPTO
                rstdestino2("H_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                rstdestino2("H_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
        End Select

        Select Case d_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                rstdestino2("H_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                rstdestino2("H_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                rstdestino2("H_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = VAR_DPTO
                rstdestino2("H_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                rstdestino2("H_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        If i = 1 Then
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
        Else
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
        End If
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
        VAR_CTAD = rstdestino!cta_descripcion
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
  'marca1 = Ado_datos.Recordset.Bookmark
  'rs_datos.Update
  'rs_datos.Requery
  Call OptFilGral1_Click
  'Set Ado_datos.Recordset = rs_datos
  'If rs_datos.RecordCount > 0 Then
    Ado_datos.Recordset.Move marca1 - 1
  'End If
  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"

End Sub

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
'           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & glGestion & ", 'APR', '" & glusuario & "', '" & Date & "')"
'    End If
    '
End Sub

Private Sub graba_ingreso()
'    '======= Ini grabado de datos
'   'swgraba = 0
'   'Call valida
'   VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'   VAR_CODTIPO = "DEI"
'   Select Case VAR_COD4
'        Case "DVTA"              'INI COMERCIAL
'            VAR_ORG = "111"
'            VAR_PARTIDA = "11310"
'        Case "COMEX"            'INI COMEX
'            VAR_ORG = "111"
'            VAR_PARTIDA = "11310"
'        Case "DNINS"            'INI INSTALACIONES
'            VAR_ORG = "111"
'            VAR_PARTIDA = "11350"
'        Case "DNAJS"            'INI AJUSTE
'            VAR_ORG = "113"
'            VAR_PARTIDA = "11350"
'        Case "DNMAN"            'INI MANTENIMIENTO
'            VAR_ORG = "112"
'            VAR_PARTIDA = "11320"
'        Case "DNREP"            'INI REPARACIONES
'            VAR_ORG = "113"
'            VAR_PARTIDA = "11330"
'        Case "DNMOD"            'INI MODERNIZACION
'            VAR_ORG = "114"
'            VAR_PARTIDA = "11340"
'        Case "DNEME"            'INI EMERGENCIAS
'            VAR_ORG = "113"
'            VAR_PARTIDA = "11330"
'        Case Else               'INI CREDITO
'            VAR_ORG = "311"
'            VAR_PARTIDA = "11350"
'   End Select
''   If swgraba = 1 Then
''      FraOpciones2.Visible = False
''      fraOpciones.Visible = True
''      FraIngresosNav.Enabled = True
''      FraIngresosDat.Enabled = False
'
'      'If v_añadir = 1 Then
'        'EFECTIVO o a CREDITO
'         'db.BeginTrans
'         Call add_correl
'         Set rstdestino = New ADODB.Recordset
'         rstdestino.Open "select * from fo_ingresos_cabecera order by org_codigo, ingreso_codigo   ", db, adOpenDynamic, adLockOptimistic
'         rstdestino.AddNew
'         rstdestino("Ges_Gestion") = glGestion      'Year(Date)     'Ado_datos.Recordset("ges_gestion")
'         rstdestino("ingreso_codigo") = correlativo1
'         VAR_CODANT = correlativo1
'         'CAMBIAR org_codigo
'         rstdestino("org_codigo") = VAR_ORG
'         'CAMBIAR org_codigo
'         'CAMBIAR COD ingreso_codigo_anterior
'         rstdestino("ingreso_codigo_anterior") = correlativo1
'         'CAMBIAR COD ingreso_codigo_anterior
'         'CAMBIAR DEI O REC
'         'VAR_CODTIPO = "DEI"
'         rstdestino("Codigo_tipo") = VAR_CODTIPO    '"DEI"
'         'VAR_CODTIPO = "DEI"
'         'CAMBIAR DEI O REC
'         rstdestino("proceso_codigo") = "FIN"
'         rstdestino("subproceso_codigo") = "FIN-01"
'         rstdestino("etapa_codigo") = "FIN-01-01"
'         rstdestino("clasif_codigo") = "ADM"
'         rstdestino("doc_codigo") = "R-110"
'         rstdestino("doc_numero") = correlativo1
'         rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos.Recordset("unidad_codigo")
'         rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos.Recordset("solicitud_codigo")
'         rstdestino("solicitud_tipo") = VAR_TIPO    '"10"
'
'         rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
'         'VAR_BENEF = Ado_datos.Recordset("beneficiario_codigo")
'         rstdestino("fecha_ingreso") = Date
'         rstdestino("tipo_cambio") = GlTipoCambioOficial 'GlTipoCambioMercado
'         rstdestino("tipo_moneda") = "BOB"
'         VAR_MONEDA = "BOB"
'         rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA2  'Ado_datos.Recordset("venta_descripcion")
'         VAR_GLOSA = "INGRESO POR: " + VAR_GLOSA2       'Ado_datos.Recordset("venta_descripcion")
'         If Ado_datos.Recordset("venta_tipo") = "E" Then
'            rstdestino("tipo_comp") = "DYR"
'         Else
'            rstdestino("tipo_comp") = "DEI"
'         End If
'         'CAMBIAR FTE
'         Select Case VAR_ORG
'             Case "111"              'INI SERVICIOS DE PROVISION E INSTALACION
'                 VAR_FTE = "10"
'             Case "112"            'INI SERVICIO DE MANTENIMIENTO - MANTENIMIENTO PREVENTIVO
'                 VAR_FTE = "10"
'             Case "113"            'INI SERVICIO DE REPARACIONES - MANTENIMIENTO CORRECTIVO
'                 VAR_FTE = "10"
'             Case "114"            'INI SERVICIO DE MODERNIZACION
'                 VAR_FTE = "10"
'             Case "211"            'INI APORTES DE CAPITAL
'                 VAR_FTE = "20"
'             Case "311"            'INI BANCO MERCANTIL SANTA CRUZ
'                 VAR_FTE = "30"
'             Case "312"            'INI BANCO DE CREDITO
'                 VAR_FTE = "30"
'             Case "411"            'INI AMT - REPOSICION DE PIEZAS Y PARTES
'                 VAR_FTE = "40"
'             Case Else               'INI OTROS
'                 VAR_FTE = "10"
'        End Select
'         rstdestino("fte_codigo") = VAR_FTE
'         'CAMBIAR FTE
'         'CAMBIAR RUBROS    'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww ya pues
'         'rstdestino("rubro_codigo") = "11200"
'         'VAR_PARTIDA = "11200"
'         'VAR_PARTIDA = "11320"
'         rstdestino("rubro_codigo") = VAR_PARTIDA
'         'CAMBIAR RUBROS
'         rstdestino("cheque_o_trf") = ""
'         rstdestino("Bco_codigo") = "NN"
'         'CAMBIAR CTA
'         rstdestino("cta_codigo") = "NN"
'         VAR_CTA = "NN"
'         'CAMBIAR CTA
'         rstdestino("numero_documento") = "0"
'         rstdestino("unidad_codigo_ant") = VAR_CITE     'Ado_datos.Recordset("unidad_codigo_ant")
'         'VAR_CITE = Ado_datos.Recordset("unidad_codigo_ant")
'         rstdestino("monto_dolares") = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'         VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'         rstdestino("monto_bolivianos") = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'         VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'         rstdestino("monto_recaudado_dolares") = 0
'         rstdestino("monto_recaudado_bolivianos") = 0
'         rstdestino("convenio_codigo") = "NN"
'         rstdestino("pro_codigo_det") = Ado_datos.Recordset("edif_codigo")
'         VAR_PROY2 = Ado_datos.Recordset("edif_codigo")
'         rstdestino("estado_CODIGO") = "APR"
'         'rstdestino("estado_codigo_dr") = "DEI"
'
'         rstdestino("usr_CODIGO") = glusuario
'         rstdestino("fecha_registro") = Date
'         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
'
'         rstdestino.Update
'         If rstdestino.State = 1 Then rstdestino.Close
'        'db.CommitTrans
'
''          If rstIngresos.State = 1 Then rstIngresos.Close
''          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
''          rstIngresos.Sort = "ingreso_codigo"
''          rstIngresos.Requery
'
''          rstIngresos.Requery
''          Set AdoIngresos.Recordset = rstIngresos
''          AdoIngresos.Refresh
''          AdoIngresos.Recordset.Find "ultimo = 'S'"
''          If Not (AdoIngresos.Recordset.EOF) Then
''            marca1 = AdoIngresos.Recordset.Bookmark
''            AdoIngresos.Recordset("ultimo") = "N"
''            AdoIngresos.Recordset.Update
''          End If
'
''          AdoIngresos.Recordset.Move marca1 - 1
'
''          marca1 = 0
'      'End If
''   Else
''      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
'''      FraOpciones2.Visible = False
'''      FraOpciones.Visible = True
'''      FraIngresosNav.Enabled = True
'''      FraIngresosDat.Enabled = False
'''      AdoIngresos.Refresh
''   End If
''   LblAccion = ""
''AAQQQQQUIIIIIIIIII    JQA

End Sub

Private Sub add_correl()
  'FALTAAAAA!! org_codigo JQA 2014-07-10
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "'", db, adOpenDynamic, adLockOptimistic
  If rstcorrel_ing.RecordCount = 0 Then
     rstcorrel_ing.AddNew
     rstcorrel_ing("org_codigo") = VAR_ORG
     rstcorrel_ing("ges_gestion") = glGestion       'Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
     'rstcorrel_ing("correlativo") = 1
     rstcorrel_ing("correlativo_ingreso") = 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
  Else
     VARG_ORGD = rstcorrel_ing!org_descripcion
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
  End If
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close

End Sub


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
                rstpagos("ges_gestion") = glGestion     'Ado_datos.Recordset("ges_gestion")
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

Private Sub CmdGrabaDet_Click()
'On Error GoTo UpdateErr
'  If dtc_codigo15 = "" Then
'     MsgBox "Debe Elejir un Bien para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'      'ado_datos14.Recordset.CancelBatch
'      'Call AbrirDetalle
'    Exit Sub
'  End If
'  If TxtDescuento.Text = "" Or TxtDescuento.Text = "0" Then
'    MsgBox "Debe Registrar la Cantidad Entregada, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    ' ado_datos14.Recordset.CancelBatch
'    'Call AbrirDetalle
'    Exit Sub
'  End If
'
''  If dtc_desc13 = "" Then
''    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atención"
''    Exit Sub
''  End If
'
''    If Ado_datos.Recordset!unidad_codigo <> "DNREP" And Ado_datos.Recordset!unidad_codigo <> "UALMR" Then
'    If CDbl(TxtDescuento.Text) > CDbl(IIf(Dtc_Stock13.Text = "", "0", Dtc_Stock13.Text)) Then
'    '        'VAR_PARTIDA = "OK"
'       MsgBox "Saldo Insuficiente en Stock (no se guardara este registro)!..."
'       ' ado_datos14.Recordset.CancelBatch
'       ' Call AbrirDetalle
'       Exit Sub
'    End If
''    End If
'
''    If Dtc_Stock13.Text = "10" Then
''    MsgBox ("El saldo en el almacen es de 10")
''    End If
''
''    If Dtc_Stock13.Text = "2" Then
''    MsgBox ("ATENCION El saldo en el almacen es de 2 ")
'''      ado_datos14.Recordset.CancelBatch
'''      Call AbrirDetalle
'''       Exit Sub
''    End If
'
'    'VARIABLES DE LA CABECERA
'    VAR_ALMX = Ado_datos.Recordset!almacen_codigoR
'    correlv = Ado_datos.Recordset!venta_codigo
'    VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'    VAR_BEN3 = Ado_datos.Recordset!beneficiario_codigo_almR
'    VAR_DOC = Ado_datos.Recordset!doc_codigo_almR
'    VAR_ALMD = IIf(IsNull(Ado_datos.Recordset!almacen_codigo_dR), "0", Ado_datos.Recordset!almacen_codigo_dR)
'
'    'If CDbl(Dtc_Stock13.Text) >= CDbl(TxtDescuento.Text) Then
'        If swnuevo = 1 Then
'            Set rs_aux8 = New ADODB.Recordset
'            If rs_aux8.State = 1 Then rs_aux8.Close
'            rs_aux8.Open "select * from ao_ventas_detalle where venta_codigo= " & correlv & "  and bien_codigo = '" & dtc_codigo15.Text & "'", db, adOpenKeyset, adLockBatchOptimistic
'            If rs_aux8.RecordCount > 0 Then
'                MsgBox "Error, El bien ya fue registrado vuelva a intentar...", , "Atención"
'                'ado_datos14.Recordset.CancelBatch
'                'Call AbrirDetalle
'                Exit Sub
'            Else
'                'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
'                ado_datos14.Recordset!venta_codigo = Ado_datos.Recordset!venta_codigo
'                ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset!ges_gestion
'                ado_datos14.Recordset!estado_codigo = "APR"
'                ado_datos14.Recordset!usr_codigo = glusuario
'                ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'                ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'            End If
'        End If
'        VAR_BIEN2 = Trim(dtc_codigo15.Text)                                     'Codigo Bien (Equipo, Producto, etc)
'        ado_datos14.Recordset!bien_codigo = Trim(VAR_BIEN2)                     'Codigo Bien (Equipo, Producto, etc)
'            ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
'            ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
'            ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
'    '       ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
'            ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW FALTA PARAMETRIZAR WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'
'            ado_datos14.Recordset!almacen_codigo = IIf(Ado_datos.Recordset!almacen_codigoR = "", "2", Ado_datos.Recordset!almacen_codigoR)
'            ado_datos14.Recordset!doc_codigo_alm = IIf(Ado_datos.Recordset!doc_codigo_alm = "", "R-115", Ado_datos.Recordset!doc_codigo_alm)
'            ado_datos14.Recordset!doc_numero_alm = IIf(Ado_datos.Recordset!doc_numero_almR = "", "1", Ado_datos.Recordset!doc_numero_almR)
'            VAR_COD2 = IIf(Ado_datos.Recordset!doc_numero_almR = "", "1", Ado_datos.Recordset!doc_numero_almR)
'            If TxtCantidad.Text = "" Then
'                TxtCantidad.Text = TxtDescuento.Text
'            End If
''
'            ado_datos14.Recordset!almacen_tipo = VAR_ALMT                              'Tipo de Almacen I, R, H, A
'            ado_datos14.Recordset!venta_det_cantidad = CDbl(IIf(TxtCantidad.Text = "", 1, TxtCantidad))  'Cantidad Solicitada
'            ado_datos14.Recordset!bien_cantidad_por_empaque = CDbl(IIf(TxtDescuento = "", 1, TxtDescuento)) 'Cantidad Entregada
'            ado_datos14.Recordset!estado_almacen = "REG"
'            ado_datos14.Recordset.Update
'            'ACTUALIZA MONTOS DEL BIEN
'            If swnuevo = 1 Then
'                db.Execute "UPDATE ao_ventas_detalle SET venta_descuento_bs = '0', venta_descuento_dol = '0' WHERE venta_codigo = " & correlv & " AND bien_codigo = '" & VAR_BIEN2 & "' "
'
'                db.Execute "UPDATE ao_ventas_detalle SET ao_ventas_detalle.venta_precio_unitario_bs  = ac_bienes.bien_precio_venta_final, ao_ventas_detalle.venta_precio_total_bs = ac_bienes.bien_precio_venta_final FROM ao_ventas_detalle INNER JOIN ac_bienes " & _
'                    " ON ao_ventas_detalle.bien_codigo  = ac_bienes.bien_codigo WHERE ao_ventas_detalle.venta_codigo = " & correlv & " AND ao_ventas_detalle.bien_codigo = '" & VAR_BIEN2 & "' "
'
'                db.Execute "UPDATE ao_ventas_detalle SET ao_ventas_detalle.venta_precio_unitario_dol  = ac_bienes.bien_precio_venta_final / " & GlTipoCambioOficial & ", ao_ventas_detalle.venta_precio_total_dol = ac_bienes.bien_precio_venta_final / " & GlTipoCambioOficial & " FROM ao_ventas_detalle INNER JOIN ac_bienes " & _
'                    " ON ao_ventas_detalle.bien_codigo  = ac_bienes.bien_codigo WHERE ao_ventas_detalle.venta_codigo = " & correlv & " AND ao_ventas_detalle.bien_codigo = '" & VAR_BIEN2 & "' "
'            End If
'
'
'            Set rs_datos15 = New ADODB.Recordset
'            If rs_datos15.State = 1 Then rs_datos15.Close
'            rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'R' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
'            Set ado_datos15.Recordset = rs_datos15
'
'            Call AbrirDetalle
'            SSTab1.Tab = 0
'            SSTab1.TabEnabled(0) = True
'            SSTab1.TabEnabled(1) = False
'            FraNavega.Enabled = True
'            FrmDetalle.Enabled = True
'            FrmABMDet.Visible = True
'            FrmEdita.Enabled = False
'            'Call OptFilGral1_Click
'            swnuevo = 0
''        End If
'    'Else
'    '    MsgBox "Saldo Insuficiente en Stock registrado en Almacenes, verifique y luego intente nuevamente !..."
'    'End If
'  'Else
'  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
'  'End If
'  accion = ""
'  Exit Sub
'UpdateErr:
'    MsgBox Err.Description
End Sub


Private Sub BtnImprimir2_Click()
'    If ado_datos14.Recordset.RecordCount > 0 Then
'         Dim iResult As Integer
'        'Dim co As New ADODB.Command
'        'CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
'        CryR01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado.rpt" '
'        CryR01.WindowShowPrintSetupBtn = True
'        CryR01.WindowShowRefreshBtn = True
'        'CryR01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
'        CryR01.StoredProcParam(0) = ado_datos14.Recordset!bien_codigo
'        CryR01.StoredProcParam(1) = Trim(Str(ado_datos14.Recordset!almacen_codigo))            'dtc_codigo1.Text
'        CryR01.StoredProcParam(2) = Format(DTP_Finicio.Value, "dd/mm/yyyy")
'        CryR01.StoredProcParam(3) = Format(DTP_Ffin.Value, "dd/mm/yyyy")
'        CryR01.Formulas(0) = "almace = '" & dtc_desc1.Text & "' "
'        'CryR01.Formulas(2) = "DEL_AL = '' "
'        'CryR01.Formulas(3) = "fechafin = '" & DTP_Ffin.Value & "' "
'
'        iResult = CryR01.PrintReport
'        If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'        CryR01.WindowState = crptMaximized
'        Fra_reporte.Visible = False
'    Else
'        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
'    End If
'    Fra_reporte.Visible = True
End Sub

Private Sub BtnAnlDetalle_Click()
On Error GoTo UpdateErr
 If Ado_datos11.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo = "REG" Then
          'sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + ado_datos14.Recordset!cobranza_codigo, vbYesNo + vbQuestion, "Atención")
          sino = MsgBox("Está Seguro de ANULAR el Registro Activo --> ", vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            'BORRA RECIBO DETALLE
            db.Execute "delete fo_recibos_detalle Where IdRecibo = " & Ado_datos.Recordset!IdRecibo & " and correl_cobro = " & Ado_datos11.Recordset!correl_cobro & "  "

            'ACTUALIZA DES-APRUEBA ao_ventas_cobranza_det
            db.Execute "UPDATE ao_ventas_cobranza_det SET estado_codigo_tes = 'REG'  WHERE cobranza_codigo = " & Ado_datos11.Recordset!cobranza_codigo & " and cobranza_detalle = " & Ado_datos11.Recordset!cobranza_detalle & " "

            ' ACTUALIZA TOTALES fo_recibos_oficiales
            db.Execute "update fo_recibos_oficiales set total_bs = (select sum(fo_recibos_detalle.cobranza_bs) from fo_recibos_detalle where fo_recibos_detalle.IdRecibo = " & Ado_datos.Recordset!IdRecibo & ")   " & _
            " from fo_recibos_oficiales inner join fo_recibos_detalle on  fo_recibos_oficiales.IdRecibo = fo_recibos_detalle.IdRecibo where fo_recibos_oficiales.IdRecibo = " & Ado_datos.Recordset!IdRecibo & " "

'            sino = MsgBox("Desea Desaprobar el Registro del Supervisor de Cobranzas ?", vbYesNo, "Confirmando")
'                If sino = vbYes Then
'                    db.Execute "update ao_ventas_cobranza_det set ao_ventas_cobranza_det.estado_codigo = 'REG', ao_ventas_cobranza_det.estado_codigo_tes = 'REG' from ao_ventas_cobranza_det inner join fo_recibos_detalle on ao_ventas_cobranza_det.correl_cobro = fo_recibos_detalle.correl_cobro WHERE fo_recibos_detalle.IdRecibo =  " & VAR_RECIBO & "  "
'                End If
            Call AbrirDetalle
          End If
       Else
          MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
 Else
     MsgBox "No se puede BORRAR, el registro ya fue BORRADO o APROBADO (APR), Verifique por favor ...", vbExclamation, "Validación de Registro"
 End If
 'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
  Exit Sub
UpdateErr:
    MsgBox Err.Description

End Sub

Private Sub BtnModDetalle_Click()
' If ado_datos14.Recordset.RecordCount > 0 Then
'  If ado_datos14.Recordset!estado_almacen = "REG" Then
'    If IsNull(Ado_datos.Recordset!almacen_codigoR) Then
'        MsgBox "El Almacen Origen NO esta registrado, verifique y vuelva a intentar... !! ", vbExclamation, "Atención!"
'        Exit Sub
'    End If
'    FraNavega.Enabled = False
'    FrmDetalle.Enabled = False
'    swnuevo = 2
'
'    marca1 = Ado_datos.Recordset.Bookmark
'    TxtNroVenta.Text = Ado_datos.Recordset!venta_codigo  'txt_venta.Text
'    TxtNroVenta.Locked = True
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'
'    FrmEdita.Visible = True
'    FrmEdita.Enabled = True
'    FrmABMDet.Visible = False
'
'    If parametro <> Ado_datos.Recordset!unidad_codigo Then
'    dtc_desc15.backColor = &HC0C0C0
'    dtc_desc15.Locked = False
'    Text9.Visible = True
'    Else
'     dtc_desc15.backColor = &HFFFFFF
'    dtc_desc15.Locked = True
'    Text9.Visible = False
'    End If
'
'    If ado_datos14.Recordset!par_codigo = "43340" Then
'        dtc_codigo13.Text = "0"
'        dtc_desc13.BoundText = dtc_codigo13.BoundText
'        dtc_desc13.backColor = &H80000013
'        dtc_desc13.ForeColor = &HFFFFFF
'    Else
'        dtc_desc13.backColor = &HFFFFFF
'        dtc_desc13.ForeColor = &H80000008
'        If ado_datos14.Recordset!bien_cantidad_por_empaque = "0" Then
'            TxtDescuento.Text = ado_datos14.Recordset!venta_det_cantidad
'        End If
'    End If
'    dtc_desc15.BoundText = dtc_codigo15.BoundText
'    dtc_unimed15.BoundText = dtc_codigo15.BoundText
'    dtc_stocktotal15.BoundText = dtc_codigo15.BoundText
'    dtc_grupo15.BoundText = dtc_codigo15.BoundText
'    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
'    Dtc_partida15.BoundText = dtc_codigo15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_codigo15.BoundText
'    dtc_precioventabase15.BoundText = dtc_codigo15.BoundText
'    dtc_preciocompra15.BoundText = dtc_codigo15.BoundText
'  Else
'    MsgBox "Los registros Aprobado o Entregado, NO pueden ser modificados !! ", vbExclamation, "Atención!"
'  End If
'
'    Else
'     MsgBox "No se puede Modificar, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
'   End If
End Sub

Private Sub CmdOK0_Click()
    VAR_OPCION = "0"
    Call AbrirDetalle
    BtnAddDetalle.Visible = False
    DtGLista.Enabled = False
End Sub

Private Sub CmdOK1_Click()
    VAR_OPCION = "1"
    Call AbrirDetalle
    BtnAddDetalle.Visible = True
    DtGLista.Enabled = True
End Sub

Private Sub CmdOK100_Click()
    'CmdOK0.Visible = True
    'DtGLista.Enabled = False
    'Call AbreOrigen
    'BtnAddDetalle.Visible = False
    'FraBuscaDet.Visible = True
    
 If ado_datos14.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      buscados = 1
      PosibleApliqueFiltro = False
      
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = DtGLista
      ClBuscaGrid.QueryUtilizado = queryinicial2
      Set ClBuscaGrid.RecordsetTrabajo = ado_datos14.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    OptFilGral1.Visible = True
    OptFilGral2.Visible = True
  End If

End Sub

Private Sub CmdOK100_LostFocus()
    CmdOK100.Visible = False
End Sub

Private Sub CmdOK2_Click()
    VAR_OPCION = "2"
    Call AbrirDetalle
    BtnAddDetalle.Visible = True
    DtGLista.Enabled = True
End Sub

Private Sub CmdOK3_Click()
    VAR_OPCION = "3"
    Call AbrirDetalle
    BtnAddDetalle.Visible = True
    DtGLista.Enabled = True
End Sub


Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub


Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
'    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_fecha7.BoundText = dtc_codigo7.BoundText
    dtc_recibo7.BoundText = dtc_codigo7.BoundText
    dtc_edificio7.BoundText = dtc_codigo7.BoundText
End Sub

'Private Sub dtc_desc15_Change()
'If accion <> "NEW" Then
'If Ado_datos.Recordset!almacen_codigoR <> "NULL" Then
'Set rs_aux9 = New ADODB.Recordset
'    If rs_aux9.State = 1 Then rs_aux9.Close
'    rs_aux9.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo = " & Ado_datos.Recordset!almacen_codigoR & " and bien_codigo ='" & dtc_codigo15.Text & "'", db, adOpenStatic
'   ' Set AdoAux9.Recordset = rs_aux9
'   If rs_aux9.RecordCount > 0 Then
'    Dtc_Stock13.Text = IIf(IsNull(rs_aux9!stock_actual), 0, rs_aux9!stock_actual)
'    End If
'  End If
'  Else
'  Dtc_Stock13.Text = "0"
'End If
'End Sub

Private Sub dtc_desc15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
TxtDescuento.SetFocus
End If
End Sub

'Private Sub pDeptoD(CodigoA As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_departamento where depto_codigo  = '" & CodigoA & "'"
'
'   Set dtc_codigo22.RowSource = Nothing
'   Set dtc_codigo22.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo22.ReFill
'   dtc_codigo22.BoundText = Empty
'
'   Set dtc_desc22.RowSource = Nothing
'   Set dtc_desc22.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc22.ReFill
'   dtc_desc22.BoundText = Empty
'End Sub


Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
'    dtc_aux4.BoundText = dtc_desc4.BoundText
    VAR_BEN2 = dtc_codigo4.Text
    'Call pAlmacen(dtc_codigo4.BoundText)
    'dtc_desc11.Enabled = True
End Sub

'Private Sub pAlmacen(CodigoA As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from ac_almacenes where beneficiario_codigo = '" & CodigoA & "'"
'
'   Set dtc_codigo11.RowSource = Nothing
'   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo11.ReFill
'   dtc_codigo11.BoundText = Empty
'
'   Set dtc_desc11.RowSource = Nothing
'   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc11.ReFill
'   dtc_desc11.BoundText = Empty
'
'   Set dtc_Aux11.RowSource = Nothing
'   Set dtc_Aux11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_Aux11.ReFill
'   dtc_Aux11.BoundText = Empty
'
'End Sub

Private Sub dtc_desc4_LostFocus()
    dtc_codigo4.Text = VAR_BEN2
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
'    If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
'        Call pAlmacenD(dtc_codigo5.BoundText)
'        dtc_desc20.Enabled = True
'    End If
End Sub

Private Sub pAlmacenD(CodigoA As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from ac_almacenes where beneficiario_codigo = '" & CodigoA & "'"
   
   Set dtc_codigo20.RowSource = Nothing
   Set dtc_codigo20.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo20.ReFill
   dtc_codigo20.BoundText = Empty
   
   Set dtc_desc20.RowSource = Nothing
   Set dtc_desc20.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc20.ReFill
   dtc_desc20.BoundText = Empty

End Sub

'Private Sub dtc_codigo13_Click(Area As Integer)
'    dtc_desc13.BoundText = dtc_codigo13.BoundText
'    Dtc_Stock13.BoundText = dtc_codigo13.BoundText
'End Sub

Private Sub dtc_codigo2A_Click(Area As Integer)
    dtc_desc2A.BoundText = dtc_codigo2A.BoundText
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

Private Sub cmdVerifica_existencia_Click()
' verifica existencia  del almacen
Cant_Alm = 0
AlFrmExistencia_Almacen.Show

DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
Txtcant_alm = Cant_Alm
If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
    End If
End Sub

'Private Sub dtc_codigo11_Click(Area As Integer)
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
'    dtc_Aux11.BoundText = dtc_codigo11.BoundText
'End Sub

'Private Sub dtc_desc11_Click(Area As Integer)
'    dtc_codigo11.BoundText = dtc_desc11.BoundText
'    dtc_Aux11.BoundText = dtc_desc11.BoundText
'    Call pDepto(dtc_Aux11.Text)
'    dtc_desc21.Enabled = True
'End Sub

'Private Sub pDepto(CodigoA As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_departamento where depto_codigo  = '" & CodigoA & "'"
'
'   Set dtc_codigo21.RowSource = Nothing
'   Set dtc_codigo21.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo21.ReFill
'   dtc_codigo21.BoundText = Empty
'
'   Set dtc_desc21.RowSource = Nothing
'   Set dtc_desc21.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc21.ReFill
'   'dtc_desc21.BoundText = Empty
'End Sub

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


'Private Sub dtc_partida15_Click(Area As Integer)
'    dtc_desc15.BoundText = Dtc_partida15.BoundText
'    dtc_unimed15.BoundText = Dtc_partida15.BoundText
'    dtc_stocktotal15.BoundText = Dtc_partida15.BoundText
'    dtc_grupo15.BoundText = Dtc_partida15.BoundText
'    dtc_subgrupo15.BoundText = Dtc_partida15.BoundText
'    dtc_codigo15.BoundText = Dtc_partida15.BoundText
''    dtc_precioventafinal15.BoundText = Dtc_partida15.BoundText
''    dtc_precioventabase15.BoundText = Dtc_partida15.BoundText
''    dtc_preciocompra15.BoundText = Dtc_partida15.BoundText
'End Sub

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

'Private Sub dtc_desc2_LostFocus()
'    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
'    If Dtc_deudor2.Text = "SI" Then
'        Dtc_deudor2.backColor = &HFF&
'    Else
'        Dtc_deudor2.backColor = &H80000010
'    End If
'
'End Sub

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

Private Sub dtc_desc15_LostFocus()
    'txt_descripcion_venta.Text = dtc_desc15.Text
End Sub

Private Sub dtc_desc2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_desc2A.BoundText
End Sub

Private Sub dtc_edificio7_Click(Area As Integer)
    dtc_fecha7.BoundText = dtc_edificio7.BoundText
    dtc_recibo7.BoundText = dtc_edificio7.BoundText
    dtc_codigo7.BoundText = dtc_edificio7.BoundText
    VAR_EDIF = dtc_edificio7.Text
End Sub

Private Sub dtc_fecha7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_fecha7.BoundText
    dtc_recibo7.BoundText = dtc_fecha7.BoundText
    dtc_edificio7.BoundText = dtc_fecha7.BoundText
    VAR_FECHA = dtc_fecha7.Text
End Sub

Private Sub dtc_recibo7_Click(Area As Integer)
    dtc_fecha7.BoundText = dtc_recibo7.BoundText
    dtc_codigo7.BoundText = dtc_recibo7.BoundText
    dtc_edificio7.BoundText = dtc_recibo7.BoundText
    VAR_RECIBOS = dtc_recibo7.Text
End Sub


Private Sub DtGLista_Click()
    VAR_BENDA = ""
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

'Private Sub DTPfechasol_LostFocus()
'    Set rs_TipoCambio = New ADODB.Recordset
'    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
'    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
'    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
'    End If
'    'Ado_datos4.Refresh

'    DTPfechaIni.Value = DTPfechasol.Value
''    'validar fecha solicitud OJO JQA 31/12/2014
''    Set rs_TipoCambio = New ADODB.Recordset
''    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
''    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
''    If rs_TipoCambio.RecordCount > 0 Then
''        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
''    End If
'End Sub

Private Sub Form_Load()
    buscados = 0
    swnuevo = 0
    accion = ""
    VAR_SW = ""
    lbl_cerrado = ""
    VAR_OPCION = "0"
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_BENEF = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_BENEF = "0"
        VAR_DA = "1.3"
    End If
    VAR_ORIGEN = Aux
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            VAR_DPTO = "3"
            parametro = "DADMB"
        Case "1.7"    'Santa Cruz
            VAR_DPTO = "7"
            parametro = "DADMS"
        Case "1.3", "1.4"    'La Paz
            VAR_DPTO = "2"
            parametro = "DCOBR"
        Case "1.9"    ' Chuquisaca
            VAR_DPTO = "1"
            parametro = "DADMC"
        Case Else    ' OTRO
            VAR_DPTO = "2"
            parametro = "DCOBR"
     End Select
    
    'REVISAR PARA TODOS LOS DOCS................
    '
    VAR_R = Aux
    
    'Call CARGAPARAM
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'Usuario
    lbl_cerrado.Caption = ""
    
    'Call ABRIR_TABLAS_AUX
    FrmDetalle.Caption = "DETALLE DE COBRANZAS - RECIBO NRO. 0"         '+ VAR_BIEN
    'aw_almacen_salida.Caption = "" + VAR_BIEN
    
    mbDataChanged = False
    FrmCabecera.Enabled = False
    dg_datos.Enabled = True
    GlNombFor = "F04"

    marca1 = 1
    deta2 = 0
    swgrabar = 0
    swnuevo = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False

    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption

  
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
'    'UNIDAD EJECUTORA
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
'    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText

'    'Beneficiario Personas Nat. y Juridicas
'     Set rs_datos2 = New ADODB.Recordset
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    rs_datos2.Open "select * from gc_unidad_ejecutora where estado_codigo = 'APR' AND da_codigo = '" & VAR_DA & "'", db, adOpenStatic
'    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    
    'Documentos de Respaldo                 OK
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_documentos_respaldo ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText

    'Beneficiario Funcionario - Quien Entrega       OK
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "Select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' order by beneficiario_denominacion", db, adOpenStatic
    rs_datos4.Open "Select * from gc_beneficiario where tipoben_codigo = '1' and estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText

    'Beneficiario Funcionario - Quien Recibe        OK
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' AND estado_codigo_resp = 'APR' order by beneficiario_denominacion", db, adOpenStatic
    'rs_datos5.Open "select * from gc_beneficiario where tipoben_codigo = '1' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText

'    'gc_departamento - Origen
'    Set rs_datos21 = New ADODB.Recordset
'    If rs_datos21.State = 1 Then rs_datos21.Close
'    rs_datos21.Open "select * from gc_departamento   ", db, adOpenStatic
'    'rs_datos21.Open "select * from gc_departamento where depto_codigo = '" & VAR_DPTO & "'  ", db, adOpenStatic      ''4273257'    'beneficiario_codigo= '" & dtc_codigo4.Text & "'
'    Set Ado_datos21.Recordset = rs_datos21
''    dtc_desc21.BoundText = dtc_codigo21.BoundText
    
'    'gc_departamento - Destino
'    Set rs_datos22 = New ADODB.Recordset
'    If rs_datos22.State = 1 Then rs_datos22.Close
'    rs_datos22.Open "select * from gc_departamento  ", db, adOpenStatic
'    'rs_datos22.Open "select * from gc_departamento where depto_codigo <>  '" & VAR_DPTO & "'  ", db, adOpenStatic       ''4273257'    'beneficiario_codigo= '" & dtc_codigo4.Text & "'
'    Set Ado_datos22.Recordset = rs_datos22
'    dtc_desc22.BoundText = dtc_codigo22.BoundText
    
End Sub

Private Sub grabar()
  'db.BeginTrans
  If VAR_ORIGEN = "R-640" Then
    VAR_R = VAR_ORIGEN
  Else
    VAR_ORIGEN = "R-640"
    VAR_R = VAR_ORIGEN
  End If
    If swgrabar = 1 Then
        Set rs_aux4 = New ADODB.Recordset
        SQL_FOR = "Select max(correl_doc) as Codigo from fo_recibos_oficiales where doc_codigo = '" & VAR_ORIGEN & "' "
        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If Not rs_aux4.EOF Then
            var_cod = IIf(IsNull(rs_aux4!Codigo), 1, rs_aux4!Codigo + 1)
            db.Execute "Update gc_documentos_respaldo Set correl_doc = " & var_cod & " Where doc_codigo = '" & VAR_ORIGEN & "'   "
        Else
            var_cod = 1
        End If
        'CREA CABECERA
       
        db.Execute "INSERT INTO fo_recibos_oficiales (clasif_codigo, doc_codigo, correl_doc, beneficiario_codigo_resp, beneficiario_codigo, unidad_codigo, unidad_codigo_resp, total_bs, total_dol, " & _
            " fecha_recibo , estado_conciliado, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
            " values ('" & dtc_aux3 & "', '" & VAR_R & "', " & var_cod & ", '" & dtc_codigo4 & "', '" & dtc_codigo5 & "', '" & parametro & "', '" & parametro & "', '0', '0',  " & _
            " '" & DTPfechasol & "', 'REG', 'REG', '" & glusuario & "', '" & Date & "', ''  ) "
    End If
    If swgrabar = 2 Then
        If Ado_datos.Recordset.RecordCount > 0 Then
            'INI ACTUALIZA
            db.Execute "UPDATE fo_recibos_oficiales SET beneficiario_codigo_resp = '" & dtc_codigo4 & "', usr_codigo = '" & glusuario & "', fecha_recibo = '" & DTPfechasol & "', beneficiario_codigo = '" & dtc_codigo5.Text & "'  WHERE IdRecibo = " & Ado_datos.Recordset!IdRecibo & " "
        End If

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
'  End If
'  glPersNew = "N"

End Sub

Private Sub OptFilGral1_Click()
   '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_datos6.RecordCount > 0 Then
        VAR_BENI = rs_datos6!beneficiario_codigo
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case glusuario
        Case "ADMIN", "VPAREDES", "SQUISPE", "CSALINAS"
            queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG') "
        Case "MPEÑARANDA", "SPAREDES", "PLOPEZ", "MVALDIVIA", "MCOARITY"
            'queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG' AND  (beneficiario_codigo ='" & VAR_BENI & "' OR beneficiario_codigo ='6962804')) "
            queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG') "
        Case "FCABRERA", "FDELGADILLO", "ASANTIVAÑEZ"
            queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG' AND  beneficiario_codigo ='" & VAR_BENI & "') "
        Case "TCASTILLO", "RGIL", "LMORALES"
            queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG' AND  (beneficiario_codigo ='" & VAR_BENI & "' OR beneficiario_codigo='2375079')) "
        Case "EVILLALOBOS"
            queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG' AND  beneficiario_codigo ='" & VAR_BENI & "') "
        'Case "PRODAS"
        '    queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG' AND  beneficiario_codigo ='" & VAR_BENI & "') "
        Case Else
            queryinicial = "select * From fo_recibos_oficiales WHERE (estado_codigo = 'REG' AND  (beneficiario_codigo ='" & VAR_BENI & "' OR beneficiario_codigo ='6962804')) "
    End Select
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    queryinicial = "select * From fo_recibos_oficiales WHERE estado_codigo = 'REG' AND  beneficiario_codigo ='' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "IdRecibo"
    'rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

Private Sub OptFilGral2_Click()
 '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_datos6.RecordCount > 0 Then
        VAR_BENI = rs_datos6!beneficiario_codigo
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case glusuario
        Case "ADMIN", "VPAREDES", "SQUISPE", "CSALINAS"
            queryinicial = "select * From fo_recibos_oficiales  "
        Case "MPEÑARANDA", "SPAREDES", "PLOPEZ", "MVALDIVIA", "MCOARITY"
            'queryinicial = "select * From fo_recibos_oficiales WHERE (beneficiario_codigo ='" & VAR_BENI & "' OR beneficiario_codigo ='6962804') "
            queryinicial = "select * From fo_recibos_oficiales  "
        Case "FCABRERA", "FDELGADILLO", "ASANTIVAÑEZ"
            queryinicial = "select * From fo_recibos_oficiales WHERE (beneficiario_codigo ='" & VAR_BENI & "') "
        Case "TCASTILLO", "RGIL", "LMORALES"
            queryinicial = "select * From fo_recibos_oficiales WHERE (beneficiario_codigo ='" & VAR_BENI & "' OR beneficiario_codigo='2375079') "
        Case "EVILLALOBOS"
            queryinicial = "select * From fo_recibos_oficiales WHERE (beneficiario_codigo ='" & VAR_BENI & "') "
        'Case "PRODAS"
        '    queryinicial = "select * From fo_recibos_oficiales WHERE (beneficiario_codigo ='" & VAR_BENI & "') "
        Case Else
            queryinicial = "select * From fo_recibos_oficiales WHERE (beneficiario_codigo ='" & VAR_BENI & "' OR beneficiario_codigo ='6962804') "
    End Select
    'Set rs_datos = New Recordset
    'If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From fo_recibos_oficiales   "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "IdRecibo"
    'rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

'Private Sub Option1_Click()
'    Fra_Total.Visible = True
'End Sub
'
'Private Sub Option2_Click()
'    FrmCobranza.Visible = True
'End Sub

Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares_contra.Text = 0
    End If
  End If
End Sub

Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares.Text = 0
    End If
  End If

End Sub

Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos_contra.Text = 0
    End If
  End If
End Sub

Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos.Text = 0
    End If
  End If
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtsolpeso_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then

    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtterref_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Then
        KeyAscii = Asc(UCase(Chr(0)))
    Else
        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "N" Or KeyAscii = 8 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            KeyAscii = Asc(UCase(Chr(0)))
            MsgBox "Debe escribir solo 'N' o 'S'", vbOKOnly, "Error..."
        End If
    End If
End Sub

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
  txtTDC.Text = GlTipoCambioOficial

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
db.Execute "drop view vwF04"

db.Execute "create view vwF04 as " & _
            "select  ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.tipoben_codigo, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, ao_solicitud_lista.telefono, ao_solicitud_lista.razon_s, ao_solicitud.codigo_solicitud, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_numero, ao_solicitud.por_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.caracteristicas, ao_solicitud.duracion_estimada_tiempo, " & _
            "ao_solicitud.tr_adjuntos AS docAdjunta, " & _
            "ao_solicitud.codigo_bien, ac_bienes.bie_descripcion , ao_solicitud.observaciones, fc_unidad_ejecutora.uni_descripcion_larga, ao_solicitud.fecha_solicitud, " & _
            "(rc_personal.paterno) + ' ' + (rc_personal.materno) + ' ' +(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
            "from ao_solicitud_lista  ,     " & _
                 "ao_solicitud       ,     " & _
                 "fc_unidad_ejecutora,     " & _
                 "rc_personal,             " & _
                 "ac_bienes                " & _
            "where  ao_solicitud_lista.ges_Gestion       = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
                    "ao_solicitud_lista.codigo_unidad    = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
                    "ao_solicitud_lista.codigo_solicitud =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
                    "ao_solicitud_lista.ges_Gestion      = ao_solicitud.ges_gestion            and " & _
                    "ao_solicitud_lista.codigo_unidad    = ao_solicitud.codigo_unidad          and " & _
                    "ao_solicitud_lista.codigo_solicitud = ao_solicitud.codigo_solicitud       and " & _
                    "ao_solicitud.codigo_unidad          = fc_unidad_ejecutora.codigo_unidad   and " & _
                    "ao_solicitud.codigo_bien            = ac_bienes.codigo_bien               and " & _
                    "ao_solicitud.ci                     = rc_personal.ci                      " & _
            "GROUP BY ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.tipoben_codigo, " & _
            "ao_solicitud.codigo_solicitud, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.razon_s, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, " & _
            "ao_solicitud_lista.telefono, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.nacional_extranjero, ao_solicitud.por_tiempo, ao_solicitud.codigo_bien, ac_bienes.bie_descripcion, ao_solicitud.duracion_estimada_numero, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.esparaRH, ao_solicitud.tr_adjuntos, ao_solicitud.observaciones, ao_solicitud.caracteristicas, fc_unidad_ejecutora.Uni_descripcion_larga, ao_solicitud.fecha_solicitud, (rc_personal.paterno)+' '+(rc_personal.materno)+' '+(rc_personal.nombres)+' ['+ao_solicitud.ci+']', ao_solicitud_lista.id_beneficiario "

'            "trim$(rc_personal.paterno) + ' ' + trim$(rc_personal.materno) + ' ' +trim$(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _

'''db.Execute "create view vwF05 as " & _
'''            "select  ao_solicitud_lista.* " & _
'''            "from ao_solicitud_lista"
End Sub

Sub CREAVISTAF11()
db.Execute "drop view VWF11"
db.Execute "create view VWF11 as " & _
    "SELECT ao_Solicitud.Ges_Gestion, ao_Solicitud.codigo_unidad, " & _
    "ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, " & _
    "ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, " & _
    "ao_Solicitud.fecha_solicitud, ao_Solicitud.codigo_bien, " & _
    "ac_bienes_grupo.DescGrupo, RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres, " & _
    "ao_Solicitud.observaciones, ao_Solicitud.caracteristicas, " & _
    "ao_Solicitud.tr_adjuntos, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, " & _
    "ao_Solicitud.duracion_estimada_numero, ao_Solicitud.duracion_estimada_tiempo, " & _
    "ao_solicitud_lista.codDetalle AS ci_material,  ao_solicitud_lista.profesion, ao_solicitud_lista.Aplanilla, " & _
    "ao_solicitud_lista.razon_s, ao_solicitud_lista.Nro_pagos, ao_solicitud_lista.Monto_solicitud_dl, ao_solicitud_lista.AUnidad " & _
"FROM ao_Solicitud, ao_Solicitud_detalle, ac_bienes_grupo, RC_Personal, ao_solicitud_lista " & _
"WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    "(ao_Solicitud.codigo_unidad) = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
    "(ao_Solicitud.codigo_solicitud) =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
    "ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_lista.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_lista.codigo_solicitud AND " & _
    "ao_Solicitud.CodGrupo = ac_bienes_grupo.CodGrupo AND " & _
    "ao_Solicitud.ci = RC_Personal.ci"
End Sub

Private Sub acumulaMont(ges, Nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  Set rs_datos19 = New ADODB.Recordset
  If rs_datos19.State = 1 Then rs_datos19.Close
'  LblGestion
'  lblcorrelVenta
'  lblNroVenta
  'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as VAR_COBR2 from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot0 from ao_ventas_detalle where venta_codigo = " & Nro & " and par_codigo = '43340'", db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!cantot0
  End If

  'rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza_prog where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If

  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.venta_codigo = " & Nro & " "

  TxtMontoBs.Text = VAR_AUX
  TxtCobrado.Text = Cobrobs
  TxtBstotal.Text = VAR_Bs

'  If IsNull(Ado_datos.Recordset!venta_monto_cobrado_bs) Then
'    Ado_datos.Recordset!venta_monto_cobrado_bs = 0
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs
'  Else
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs
'  End If
'  If VAR_AUX > 0 Then
'        VAR_AUX2 = VAR_AUX / Ado_datos.Recordset!venta_tipo_cambio
'  Else
'        VAR_AUX2 = 0
'  End If
'  'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.monto_total_Bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.monto_cobrado = " & rstacumdet!totbs & ", ao_ventas_cabecera.monto_total_Us = " & rstacumdet!totdl & ", ao_ventas_cabecera.cantidad_total_vendida = " & rstacumdet!cantot & ", ao_ventas_cabecera.saldo_p_cobrar = ao_ventas_cabecera.monto_total_Bs - ao_ventas_cabecera.deuda_cobrada Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.venta_monto_total_dol = " & rstacumdet!totdl & ", ao_ventas_cabecera.venta_cantidad_total = " & rstacumdet!cantot & ", ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_AUX & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_AUX2 & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'
'  TxtMontoBs = rstacumdet!totbs
'  TxtCobrado = rs_datos19!totbs2    'IIf(IsNull(Ado_datos.Recordset("venta_monto_cobrado_bs")), 0, Ado_datos.Recordset("venta_monto_cobrado_bs"))
'  If IsNull(Ado_datos.Recordset("venta_saldo_p_cobrar_bs")) Then
'    Text2 = VAR_AUX 'Ado_datos.Recordset("venta_monto_total_bs") - Ado_datos.Recordset("venta_monto_cobrado_bs")
'    Ado_datos.Recordset("venta_saldo_p_cobrar_bs") = VAR_AUX
'  Else
'    Text2 = Ado_datos.Recordset("venta_saldo_p_cobrar_bs")
'  End If

  If rstacumdet.State = 1 Then rstacumdet.Close

  'Print ado_datos14.Recordset!ges_gestion
  'Print ado_datos14.Recordset!correl_venta
  'Print ado_datos14.Recordset!venta_codigo
  'ado_datos14.Recordset!monto_Bolivianos = rstacumdet!totbs
  'ado_datos14.Recordset!monto_dolares = rstacumdet!totdl
  'ado_datos14.Recordset.Update
'  Set rstdestino = New ADODB.Recordset
'  If rstdestino.State = 1 Then rstdestino.Close
'  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & ges & "' and correl_venta = '" & corr & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
'  If rstdestino.RecordCount > 0 Then
'    rstdestino!monto_total_Bs = rstacumdet!totbs
'    rstdestino!monto_cobrado = rstacumdet!totbs
'    rstdestino!monto_total_Us = rstacumdet!totdl
'    rstdestino!cantidad_total_vendida = rstacumdet!cantot
'    rstdestino!saldo_p_cobrar = 0
'    rstdestino.Update
'  End If
'  'Set Ado_datos.Recordset = rstdestino
'  If rstdestino.State = 1 Then rstdestino.Close
'  If rstacumdet.State = 1 Then rstacumdet.Close
End Sub

Private Sub Option1_Click()
    'ORIGEN - DETALLE DE COBRANZAS
    Set rs_datos14 = New ADODB.Recordset
    If rs_datos14.State = 1 Then rs_datos14.Close
    If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        rs_datos14.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo = 'APR' and estado_codigo_tes = 'REG'  ", db, adOpenKeyset, adLockOptimistic  ' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'     ' order by  cobranza_fecha DESC
    Else
        rs_datos14.Open "select * from fv_ventas_cobranza_det_recibos where  idRecibo = " & Ado_datos.Recordset!IdRecibo & " order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic   ' beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND
    End If
    rs_datos14.Sort = "cobranza_fecha desc"
    Set ado_datos14.Recordset = rs_datos14.DataSource
    ado_datos14.Recordset.Requery
    If ado_datos14.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista.Visible = True
        FraBuscaDet.Visible = True
        Set DtGLista.DataSource = ado_datos14.Recordset
        'Call AbreAlmacen
    Else
        deta2 = 0
        DtGLista.Visible = False
        FraBuscaDet.Visible = False
    End If
    
    'DESTINO - DETALLE DEL RECIBO
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        'rs_datos11.Open "select * from fv_ventas_cobranza_det_recibos WHERE estado_codigo_tes = 'APR' and estado_codigo_cont = 'REG' AND beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "'   order by  cobranza_codigo DESC ", db, adOpenKeyset, adLockOptimistic
        rs_datos11.Open "select * from fv_ventas_cobranza_det_recibos where  idRecibo = " & Ado_datos.Recordset!IdRecibo & " order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic   ' beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND
    Else
        rs_datos11.Open "select * from fv_ventas_cobranza_det_recibos where  idRecibo = " & Ado_datos.Recordset!IdRecibo & " order by  cobranza_fecha DESC ", db, adOpenKeyset, adLockOptimistic   ' beneficiario_codigo_resp = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND
    End If
    rs_datos11.Sort = "doc_numero"
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
        'Call AbreAlmacen
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If

End Sub

Private Sub Option2_Click()
    Call AbrirDetalle
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

Private Sub txt_descripcion_venta_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'KeyAscii = 0
'Call CmdGrabaDet_Click
'Call BtnAddDetalle_Click
'End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtCantidad_LostFocus()
  If (TxtCantidad.Text) = "" Then
    TxtCantidad.Text = 1
  End If
  If dtc_codigo11.Text = "E" Then
    If (dtc_codigo12.Text) = "" Or IsNull(dtc_codigo12.Text) Then
        TxtDescuento.Text = "0"
    Else
        TxtDescuento.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) * CDbl(Dtc_aux12.Text))
    End If
    'TxtPrecioU.Text = dtc_precioventabase15.Text
    'TxtTotal.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) - CDbl(TxtDescuento.Text))
  End If
  If dtc_codigo11.Text = "C" Then
     TxtDescuento.Text = "0"
     'TxtDescuento.Text = CDbl(Dtc_aux12) * (CDbl(TxtCantidad) * CDbl(TxtPrecioU))
     TxtPrecioU.Text = dtc_precioventafinal15.Text
  End If
  If (dtc_codigo11.Text <> "E" And dtc_codigo11.Text <> "C") Then
     TxtDescuento.Text = "0"
     TxtPrecioU.Text = "0"
  End If
  TxtTotal.Text = (CDbl(TxtCantidad.Text) * CDbl(TxtPrecioU.Text)) - CDbl(TxtDescuento.Text)

End Sub

Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Then      '(KeyAscii = 8) Or '(0..9)
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtDsctoTot_LostFocus()
    If TxtDsctoTot.Text = "" Or TxtDsctoTot.Text = "0" Or TxtDsctoTot.Text = "0.00" Then
        TxtMonto.Text = "0"
    Else
        TxtMonto.Text = Round(CDbl(TxtDsctoTot.Text) * GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtDsctoTot.Text = "0"
    Else
        TxtDsctoTot.Text = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub

'Private Sub TxtDescuento_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        Call CmdGrabaDet_Click
'        Call BtnAddDetalle_Click
'        'txt_descripcion_venta.SetFocus
'    End If
'
'    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46 Then
'        Exit Sub
'    Else
'        KeyAscii = 0
'    End If
'
'End Sub


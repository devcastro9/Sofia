VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_requerimiento_compra 
   BackColor       =   &H00000000&
   Caption         =   "Procesos Administrativos - Compras - Requerimiento de Compra"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "aw_requerimiento_compra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmABMDet3 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   120
      Picture         =   "aw_requerimiento_compra.frx":0A02
      ScaleHeight     =   1425
      ScaleWidth      =   2715
      TabIndex        =   94
      Top             =   8280
      Width           =   2775
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1320
         Picture         =   "aw_requerimiento_compra.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Imprime Detalle de Registro"
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   60
         Picture         =   "aw_requerimiento_compra.frx":6D301
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Anula Detalle de Registro"
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1305
         Picture         =   "aw_requerimiento_compra.frx":6DA4D
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Modifica Detalle de Registro"
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   60
         Picture         =   "aw_requerimiento_compra.frx":6E362
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Adiciona Detalle de Registro"
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   120
      Picture         =   "aw_requerimiento_compra.frx":6EB21
      ScaleHeight     =   1425
      ScaleWidth      =   2715
      TabIndex        =   89
      Top             =   6720
      Width           =   2775
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1300
         Picture         =   "aw_requerimiento_compra.frx":DAB53
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Imprime Detalle de Registro"
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   60
         Picture         =   "aw_requerimiento_compra.frx":DB420
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Anula Detalle de Registro"
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1300
         Picture         =   "aw_requerimiento_compra.frx":DBB6C
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Modifica Detalle de Registro"
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   60
         Picture         =   "aw_requerimiento_compra.frx":DC481
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Adiciona Detalle de Registro"
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   120
      Picture         =   "aw_requerimiento_compra.frx":DCC40
      ScaleHeight     =   1425
      ScaleWidth      =   2715
      TabIndex        =   84
      Top             =   5040
      Width           =   2775
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1285
         Picture         =   "aw_requerimiento_compra.frx":148C72
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Imprime Detalle de Registro"
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton BtnAnlDetalle2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   60
         Picture         =   "aw_requerimiento_compra.frx":14953F
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Anula Detalle de Registro"
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1285
         Picture         =   "aw_requerimiento_compra.frx":149C8B
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Modifica Detalle de Registro"
         Top             =   180
         Width           =   1370
      End
      Begin VB.CommandButton BtnAddDetalle2 
         BackColor       =   &H00FFC0C0&
         Height          =   645
         Left            =   60
         Picture         =   "aw_requerimiento_compra.frx":14A5A0
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Adiciona Detalle de Registro"
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.PictureBox fra_opciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   71
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   15960
         Picture         =   "aw_requerimiento_compra.frx":14AD5F
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "aw_requerimiento_compra.frx":14AF69
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   79
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
         Left            =   5505
         Picture         =   "aw_requerimiento_compra.frx":14B72B
         ScaleHeight     =   615
         ScaleWidth      =   1365
         TabIndex        =   78
         ToolTipText     =   "Imprime Lista de Registros"
         Top             =   0
         Width           =   1365
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "aw_requerimiento_compra.frx":14BFF8
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   77
         ToolTipText     =   "Buscar Registros"
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
         Picture         =   "aw_requerimiento_compra.frx":14C7AD
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   76
         ToolTipText     =   "Aprueba el Registro Seleccionado"
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
         Picture         =   "aw_requerimiento_compra.frx":14CFE0
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   75
         ToolTipText     =   "Anula el Registro Seleccionado"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "aw_requerimiento_compra.frx":14D72C
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   74
         ToolTipText     =   "Modifica el Registro Seleccionado"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "aw_requerimiento_compra.frx":14E041
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   73
         ToolTipText     =   "Adiciona un Nuevo Registro"
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   15600
         Picture         =   "aw_requerimiento_compra.frx":14E800
         Style           =   1  'Graphical
         TabIndex        =   72
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
         TabIndex        =   81
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
      TabIndex        =   67
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
         Picture         =   "aw_requerimiento_compra.frx":14EC42
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   69
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "aw_requerimiento_compra.frx":14F418
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   68
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
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   13215
         TabIndex        =   70
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00000000&
      Caption         =   "Elija Parámetros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3720
      Left            =   6120
      TabIndex        =   57
      Top             =   840
      Visible         =   0   'False
      Width           =   7860
      Begin VB.CommandButton BtnCancelar3 
         BackColor       =   &H80000015&
         Height          =   615
         Left            =   4080
         Picture         =   "aw_requerimiento_compra.frx":14FD04
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Cancela sin Guardar"
         Top             =   2880
         Width           =   1485
      End
      Begin VB.PictureBox btn_imprimir_reporte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2040
         Picture         =   "aw_requerimiento_compra.frx":1505F0
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   82
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   2880
         Width           =   1400
      End
      Begin MSDataListLib.DataCombo dtc_desc_dpto 
         Bindings        =   "aw_requerimiento_compra.frx":150EBD
         Height          =   315
         Left            =   1920
         TabIndex        =   58
         Top             =   675
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "depto_descripcion"
         BoundColumn     =   "depto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc_unidad 
         Bindings        =   "aw_requerimiento_compra.frx":150ED7
         Height          =   315
         Left            =   1920
         TabIndex        =   59
         Top             =   1440
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc_z_piloto 
         Bindings        =   "aw_requerimiento_compra.frx":150EF1
         Height          =   315
         Left            =   1920
         TabIndex        =   60
         Top             =   2160
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_cod_dpto 
         Bindings        =   "aw_requerimiento_compra.frx":150F0B
         DataField       =   "depto_codigo"
         Height          =   315
         Left            =   840
         TabIndex        =   61
         Top             =   675
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "depto_codigo"
         BoundColumn     =   "depto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_cod_unidad 
         Bindings        =   "aw_requerimiento_compra.frx":150F25
         DataField       =   "unidad_codigo"
         Height          =   315
         Left            =   840
         TabIndex        =   62
         Top             =   1440
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_cod_z_piloto 
         Bindings        =   "aw_requerimiento_compra.frx":150F3F
         DataField       =   "zpiloto_codigo"
         Height          =   315
         Left            =   840
         TabIndex        =   63
         Top             =   2160
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "zpiloto_codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lbl_campo5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento (Regional)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   1920
         TabIndex        =   66
         Top             =   405
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   1920
         TabIndex        =   65
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Zona Piloto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   1920
         TabIndex        =   64
         Top             =   1920
         Width           =   990
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   3120
      TabIndex        =   49
      Top             =   6540
      Width           =   12735
      Begin MSDataGridLib.DataGrid dg_det1 
         Height          =   1215
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2143
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "bitacora_codigo"
            Caption         =   "Correl"
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
            DataField       =   "negocia_forma"
            Caption         =   "Tipo.Evento"
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
            DataField       =   "negocia_fecha_real"
            Caption         =   "Fecha Evento"
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
            DataField       =   "negocia_hora_real"
            Caption         =   "Hora Evento"
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
            DataField       =   "negocia_gasto_estimado"
            Caption         =   "Gasto Estimado"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cliente Contactado"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Personal CGI"
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
            DataField       =   "negocia_tarea_realizada"
            Caption         =   "Tema Tratado"
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
            DataField       =   "negocia_observaciones"
            Caption         =   "Conclusiones u Observaciones"
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
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   3734.929
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   2294.929
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00000000&
      Caption         =   "COTIZACION"
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   3105
      TabIndex        =   30
      Top             =   8220
      Visible         =   0   'False
      Width           =   12735
      Begin MSDataGridLib.DataGrid dg_det3 
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2143
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "CodGrupo"
            Caption         =   "Grupo"
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
            DataField       =   "cod_montador"
            Caption         =   "Sub-Grupo"
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
            DataField       =   "CodDetalle"
            Caption         =   "Codigo Bien"
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
            DataField       =   "DescDetalle"
            Caption         =   "Denominación del Producto"
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
            DataField       =   "cantidad"
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
         BeginProperty Column05 
            DataField       =   "precio_compra"
            Caption         =   "Precio.Actual"
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
            DataField       =   "Total_compra"
            Caption         =   "Total Actual"
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
            DataField       =   "precio_venta"
            Caption         =   "Precio.c/Dscto."
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
            DataField       =   "Total_venta"
            Caption         =   "Total c/Dscto."
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
            DataField       =   "profesion"
            Caption         =   "Caracteristicas del Bien"
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
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2355.024
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00000000&
      Caption         =   "DETALLE DE LA SOLICITUD"
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   3120
      TabIndex        =   25
      Top             =   4920
      Width           =   12735
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "aw_requerimiento_compra.frx":150F59
         Height          =   1215
         Left            =   75
         TabIndex        =   26
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo de Bien"
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
            DataField       =   "marca_codigo"
            Caption         =   "Marca del Bien"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo del Bien"
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
            DataField       =   "bien_cantidad"
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
         BeginProperty Column04 
            DataField       =   "unimed_codigo"
            Caption         =   "Unidad Medida"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "fosa_dimension_frente"
            Caption         =   "Dim.Fosa Frente"
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
            DataField       =   "fosa_dimension_fondo"
            Caption         =   "Dim.Fosa Fondo"
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
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   5790.047
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   4080
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   5895
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
         Left            =   1200
         TabIndex        =   44
         Top             =   3720
         Value           =   -1  'True
         Width           =   1575
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
         Left            =   3600
         TabIndex        =   45
         Top             =   3720
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3600
         Width           =   5625
         _ExtentX        =   9922
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "aw_requerimiento_compra.frx":150F74
         Height          =   3255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
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
            DataField       =   "solicitud_fecha_solicitud"
            Caption         =   "Fecha.Req."
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
            DataField       =   "estado_cotiza"
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
         BeginProperty Column05 
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
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   4080
      Left            =   6000
      TabIndex        =   12
      Top             =   720
      Width           =   9855
      Begin VB.TextBox Text3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DataSource      =   "Ado_datos"
         Enabled         =   0   'False
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   4720
         TabIndex        =   54
         Top             =   1165
         Width           =   345
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "solicitud_fecha_solicitud"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   5625
         TabIndex        =   1
         Top             =   3660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   110231553
         CurrentDate     =   41678
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   9195
         TabIndex        =   53
         Top             =   1990
         Width           =   270
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7080
         TabIndex        =   48
         Top             =   500
         Width           =   270
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DataSource      =   "Ado_datos"
         Enabled         =   0   'False
         Height          =   285
         Left            =   9195
         TabIndex        =   43
         Top             =   1175
         Width           =   270
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   1605
      End
      Begin MSDataListLib.DataCombo dtc_aux11 
         Bindings        =   "aw_requerimiento_compra.frx":150F8C
         DataField       =   "beneficiario_codigo_resp2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2640
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "aw_requerimiento_compra.frx":150FA6
         DataField       =   "beneficiario_codigo_resp2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3360
         TabIndex        =   33
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
         Bindings        =   "aw_requerimiento_compra.frx":150FC0
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   32
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
         Bindings        =   "aw_requerimiento_compra.frx":150FD9
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8280
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "solicitud_tipo"
         BoundColumn     =   "solicitud_tipo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "aw_requerimiento_compra.frx":150FF2
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2820
         TabIndex        =   4
         Top             =   2880
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "aw_requerimiento_compra.frx":15100C
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "aw_requerimiento_compra.frx":151025
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   18
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
         Bindings        =   "aw_requerimiento_compra.frx":15103E
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_justificacion"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2400
         Width           =   8265
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "aw_requerimiento_compra.frx":151057
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1155
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   0
         ForeColor       =   16777215
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "aw_requerimiento_compra.frx":151070
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5340
         TabIndex        =   16
         Top             =   1155
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   0
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "aw_requerimiento_compra.frx":151089
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4560
         TabIndex        =   20
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
         Bindings        =   "aw_requerimiento_compra.frx":1510A2
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5115
         TabIndex        =   21
         Top             =   1980
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "solicitud_tipo_descripcion"
         BoundColumn     =   "solicitud_tipo"
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
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "aw_requerimiento_compra.frx":1510BB
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   5565
         _ExtentX        =   9816
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
         Bindings        =   "aw_requerimiento_compra.frx":1510D4
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Top             =   2880
         Width           =   1245
         _ExtentX        =   2196
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
         Bindings        =   "aw_requerimiento_compra.frx":1510EE
         DataField       =   "beneficiario_codigo_resp2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   1980
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
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
         Left            =   180
         TabIndex        =   52
         Top             =   3660
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Trámite"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   1
         Left            =   5160
         TabIndex        =   51
         Top             =   1710
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite del Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   6
         Left            =   7830
         TabIndex        =   47
         Top             =   240
         Width           =   1530
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
         Left            =   7740
         TabIndex        =   46
         Top             =   510
         Width           =   1695
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Edificio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Left            =   180
         TabIndex        =   42
         Top             =   885
         Width           =   660
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   41
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Actividad del POA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   40
         Top             =   2940
         Width           =   1635
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   3380
         Width           =   1755
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Responsable del Proceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   38
         Top             =   1710
         Width           =   2415
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Representante Legal / Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Left            =   5340
         TabIndex        =   37
         Top             =   885
         Width           =   2655
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad Ejecutora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   1845
         TabIndex        =   36
         Top             =   225
         Width           =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9840
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9840
         Y1              =   1620
         Y2              =   1620
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
         TabIndex        =   29
         Top             =   510
         Width           =   1335
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
         Left            =   2820
         TabIndex        =   28
         Top             =   3660
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. de Documento Respaldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   13
         Left            =   2280
         TabIndex        =   24
         Top             =   3375
         Width           =   2685
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   12
         Left            =   5625
         TabIndex        =   23
         Top             =   3375
         Width           =   1665
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "REG"
         DataField       =   "estado_cotiza"
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
         TabIndex        =   5
         Top             =   3660
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cod.Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   2
         Left            =   7995
         TabIndex        =   13
         Top             =   3375
         Width           =   765
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
      TabIndex        =   6
      Top             =   10260
      Width           =   11280
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   11
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
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
      Top             =   10680
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
      Top             =   10680
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
      Top             =   11040
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
      Top             =   11040
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
   Begin MSAdodcLib.Adodc Ado_datos13 
      Height          =   330
      Left            =   24240
      Top             =   3600
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
   Begin MSAdodcLib.Adodc Ado_datos14 
      Height          =   330
      Left            =   1200
      Top             =   840
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
      Caption         =   "Ado_datos14"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2280
      Top             =   0
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
      Caption         =   "Ado_datos14"
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
   Begin MSAdodcLib.Adodc Ado_datos15 
      Height          =   330
      Left            =   10920
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
      Caption         =   "Ado_datos15"
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
Attribute VB_Name = "aw_requerimiento_compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
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

Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset
Dim rs_datos15 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
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
Dim parametro As String
Dim VAR_DET As String

Dim VAR_REG As Integer
Dim VAR_AUX, VAR_CONT2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub btn_imprimir_reporte_Click()
If (Ado_datos13.Recordset.RecordCount > 0) And (ado_datos14.Recordset.RecordCount > 0) And (ado_datos15.Recordset.RecordCount > 0) Then
'    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\Insumos\ar_listar_id_cliente_almacen.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
          'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
        'Call CREAVISTAF11          'JQA JUN-2008
        If dtc_cod_dpto.Text = "" Then
         CR01.StoredProcParam(0) = "%"
         Else
          CR01.StoredProcParam(0) = dtc_cod_dpto.Text
        End If
        
            If dtc_cod_unidad.Text = "" Then
         CR01.StoredProcParam(1) = "%"
         Else
          CR01.StoredProcParam(1) = dtc_cod_unidad.Text
        End If
        
            If dtc_cod_z_piloto.Text = "" Then
         CR01.StoredProcParam(2) = "%"
         Else
          CR01.StoredProcParam(2) = dtc_cod_z_piloto.Text
        End If
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
        CR01.WindowState = crptMaximized
'    Else
'        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
'    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnAddDetalle_Click()
'  marca1 = Ado_datos.Recordset.Bookmark
'  If rs_datos!estado_cotiza = "REG" Then
'    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'    swnuevo = 1
'    fra_opciones.Enabled = False
'    FraNavega.Enabled = False
'    FraDet1.Enabled = False
'    FrmABMDet.Enabled = False
'    FraDet2.Enabled = False
'    FrmABMDet2.Enabled = False
'    Fra_datos.Enabled = False
'    Call ABRIR_TABLA_DET
'    aw_solicitud_bienes_insumos.txt_codigo.Caption = Me.txt_codigo.Caption
'    aw_solicitud_bienes_insumos.txt_campo1.Caption = Me.dtc_codigo1.Text
'    aw_solicitud_bienes_insumos.Txt_descripcion.Caption = Me.dtc_desc1.Text
'    aw_solicitud_bienes_insumos.txt_correl.Caption = 0    'rs_datos!correl_bitacora + 1
'    aw_solicitud_bienes_insumos.Txt_estado.Caption = "REG"
''    aw_compra_bitacora.lbl_bitacora.Caption = Me.FraDet1.Caption
'    Ado_detalle1.Recordset.AddNew
'    frm_ao_solicitud_bitacora.Show vbModal
'
'    Call ABRIR_TABLA_DET
'
'    swnuevo = 0
'    fra_opciones.Enabled = True
'    FraNavega.Enabled = True
''    FraDet1.Enabled = True
'    FrmABMDet.Enabled = True
'    FraDet2.Enabled = True
'    FrmABMDet2.Enabled = True
'    'Fra_datos.Enabled = True
'  Else
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
'  End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_cotiza = "REG" Then
    swnuevo = 1
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
            Call ABRIR_TABLA_DET
            If VAR_DET = "30000" Then
                Ado_detalle2.Recordset.AddNew
                aw_solicitud_bienes_insumos.txt_codigo.Caption = Me.txt_codigo.Caption
                aw_solicitud_bienes_insumos.txt_campo1.Caption = Me.dtc_codigo1.Text
                aw_solicitud_bienes_insumos.Txt_descripcion.Caption = Me.dtc_desc1.Text
                aw_solicitud_bienes_insumos.lbl_edif.Caption = dtc_codigo3.Text
                aw_solicitud_bienes_insumos.lbl_det.Caption = VAR_DET     '"34110"
                aw_solicitud_bienes_insumos.Txt_estado.Caption = "REG"
                aw_solicitud_bienes_insumos.Show vbModal
            End If
            If VAR_DET = "39800" Then
                Ado_detalle2.Recordset.AddNew
                frm_solicitud_bienes5A.txt_codigo.Caption = Me.txt_codigo.Caption
                frm_solicitud_bienes5A.txt_campo1.Caption = Me.dtc_codigo1.Text
                frm_solicitud_bienes5A.Txt_descripcion.Caption = Me.dtc_desc1.Text
                frm_solicitud_bienes5A.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes5A.lbl_det.Caption = VAR_DET     '"34110"
                frm_solicitud_bienes5A.Txt_estado.Caption = "REG"
                frm_solicitud_bienes5A.Show vbModal
            End If
            If VAR_DET = "34800" Then
                Ado_detalle2.Recordset.AddNew
                frm_solicitud_bienes6.txt_codigo.Caption = Me.txt_codigo.Caption
                frm_solicitud_bienes6.txt_campo1.Caption = Me.dtc_codigo1.Text
                frm_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
                frm_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes6.lbl_det.Caption = VAR_DET     '"34110"
                frm_solicitud_bienes6.Txt_estado.Caption = "REG"
                frm_solicitud_bienes6.Show vbModal
            End If

        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            

        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            Call ABRIR_TABLA_DET
            Ado_detalle1.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            'mw_solicitud_edificacion.Txt_Correl.Caption = 0
'            mw_solicitud_edificacion.dtc_codigo1.Text = Me.dtc_codigo3.Text
'            mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.Show vbModal
        Case "5"    ' SERVICIO MODERNIZACION
        
        
    End Select
    
    swnuevo = 0
    fra_opciones.Enabled = True
    FraNavega.Enabled = True
'    FraDet2.Enabled = True
    
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
    Call ABRIR_TABLA_DET
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
    FraDet2.Enabled = True
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
  If Ado_detalle1.Recordset.RecordCount > 0 Then
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo --> " + Str(Ado_detalle1.Recordset!bitacora_codigo), vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
        Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
 Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
 End If
End Sub

Private Sub BtnAnlDetalle2_Click()
   If Ado_detalle2.Recordset.RecordCount > 0 Then
       If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle2.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
'        MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
'        Exit Sub
'   End If
'   Set rs_aux2 = New ADODB.Recordset
'   rs_aux2.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'   If rs_aux2.RecordCount > 0 Then
'        VAR_CONT2 = rs_aux2.RecordCount
'   End If
   If rs_datos!estado_cotiza = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      VAR_COD4 = Ado_datos.Recordset!unidad_codigo
      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
      VAR_PROY2 = Ado_datos.Recordset!edif_codigo
      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
      
      If sino = vbYes Then
        Select Case dtc_codigo2.Text
            'Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "1", "10"    ' MANTENIMIENTO PREVENTIVO - INSUMOS y/o COMPRAS BB y SS
                'EQUIPO
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & parametro & "'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux2.RecordCount > 0 Then
                       rs_aux2!correl_negocia = rs_aux2!correl_negocia + 1
                       correldetalle = rs_aux2!correl_negocia
                       rs_aux2.Update
                    End If
                    'WWWWWWWWWWWWWWW
                    'correlv = Ado_datos.Recordset!venta_codigo
                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
           
                    Set rs_aux3 = New ADODB.Recordset
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux3.RecordCount = 0 Then
                    'beneficiario_codigo_resp,'doc_numero,estado_codigo_tra, estado_codigo_nac, estado_codigo_des, hora_registro, usr_codigo_aprueba,'                      fecha_registro_aprueba
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = glGestion     'Year(Date)
                        'rs_aux3!compra_codigo = 0      'Autonumerico
                        rs_aux3!unidad_codigo_adm = parametro
                        rs_aux3!solicitud_codigo_adm = correldetalle
                        rs_aux3!unidad_codigo = VAR_COD4
                        rs_aux3!solicitud_codigo = VAR_SOL
                        rs_aux3!edif_codigo = VAR_PROY2
                        rs_aux3!beneficiario_codigo = VAR_BENEF
                        rs_aux3!solicitud_tipo = Ado_datos.Recordset!solicitud_tipo       '"10"
                        rs_aux3!venta_tipo = "E"
                        rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
                        rs_aux3!compra_fecha = Date
                        rs_aux3!compra_DESCRIPCION = "COMPRA POR: " + lbl_titulo.Caption
                        rs_aux3!compra_observaciones = "Edificio: " + Trim(dtc_desc3.Text)
                        rs_aux3!compra_cantidad_total = 1   'Ado_datos.Recordset!venta_cantidad_total
                        rs_aux3!compra_monto_bs = 0     'VAR_BS2
                        rs_aux3!tipo_moneda = "BOB"
                        rs_aux3!compra_monto_DOL = 0        'VAR_DOL2
                        rs_aux3!proceso_codigo = "TEC"
                        rs_aux3!subproceso_codigo = "TEC-06"
                        rs_aux3!etapa_codigo = "TEC-06-01"
                        rs_aux3!clasif_codigo = "ADM"
                        rs_aux3!doc_codigo = "R-114"
                        rs_aux3!poa_codigo = "3.2.8"
                        rs_aux3!estado_codigo_eqp = "REG"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3!Fecha_Registro = Date
                        rs_aux3.Update
                        
                        'DETALLE Carga ao_ventas_detalle
                        Set rstdestino = New ADODB.Recordset
                        If rstdestino.State = 1 Then rstdestino.Close
                        rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
                        If rstdestino.RecordCount > 0 Then
                        End If
                        Set rs_aux4 = New ADODB.Recordset
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        'rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & rs_aux3!compra_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                        rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & VAR_SOL & "  and grupo_codigo = '30000' ", db, adOpenKeyset, adLockBatchOptimistic
                        If rs_aux4.RecordCount > 0 Then
                            VAR_REG = 1
                           rs_aux4.MoveFirst
                           While Not rs_aux4.EOF
                              If rs_aux4!grupo_codigo = "30000" Then
                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_DESCRIPCION & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
                                
                                db.Execute "Update ao_compra_detalle SET ao_compra_detalle.compra_concepto  = ac_bienes.bien_descripcion From ao_compra_detalle INNER JOIN ac_bienes ON ao_compra_detalle.bien_codigo = ac_bienes.bien_codigo where ao_compra_detalle.compra_codigo = " & rs_aux3!compra_codigo & " and ao_compra_detalle.bien_codigo = '" & rs_aux4!bien_codigo & "' "
                                VAR_REG = VAR_REG + 1
                              End If
'Update ao_ventas_cabecera SET ao_ventas_cabecera.venta_monto_total_bs  = (select SUM(ao_ventas_detalle.venta_precio_total_bs) from ao_ventas_detalle WHERE ao_ventas_cabecera.venta_codigo = ao_ventas_detalle.venta_codigo and par_codigo='43340')
'From ao_ventas_cabecera     INNER JOIN ao_ventas_detalle     ON (ao_ventas_cabecera.venta_codigo = ao_ventas_detalle.venta_codigo ) where ao_ventas_cabecera.unidad_codigo = 'DNMAN'
                                rs_aux4.MoveNext
                           Wend
                        End If
                        If rstdestino.State = 1 Then rstdestino.Close
                    End If
                    'WWWWWWWWWW
            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            Case "5"    ' SERVICIO MODERNIZACION
        End Select
        Set rs_aux2 = New ADODB.Recordset
        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
            txt_campo1.Caption = rs_aux2!correl_doc
            rs_aux2.Update
        End If
        rs_datos!doc_numero = txt_campo1.Caption
        'REVISAR !!! JQA 2014_07_08
        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!estado_cotiza = "APR"
        rs_datos!fecha_aprueba = Date
        rs_datos!usr_codigo_aprueba = glusuario
        rs_datos.UpdateBatch adAffectAll
  End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
   End If
  Else
      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
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
   rs_datos.CancelBatch
    dtc_desc3.Locked = False
    dtc_desc3.backColor = &H0&
    Text3.Visible = True
     dtc_desc3.ForeColor = &HFFFFFF
    
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
        fra_opciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        dtc_desc3.Locked = True
        'txt_codigo.Enabled = True
        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub BtnCancelar3_Click()
Fra_reporte.Visible = False
fra_opciones.Enabled = True
dg_datos.Enabled = True
Ado_datos.Enabled = True
Fra_datos.Enabled = True
FrmABMDet2.Enabled = True
dg_det2.Enabled = True
FrmABMDet.Enabled = True
dg_det1.Enabled = True
FrmABMDet3.Enabled = True
dg_det3.Enabled = True
End Sub

Private Sub btnEliminar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
     If rs_datos!estado_cotiza = "APR" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
     If ExisteReg(Ado_datos.Recordset!unidad_codigo_sol, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
          rs_datos!estado_cotiza = "ANL"
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

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
'  Select Case parametro
'        Case "UALMI"    'INSUMOS
'            VAR_UNI = "UALMI"
'        Case "UALMR"    'REPUESTOS
'            VAR_UNI = "UALMR"
'        Case "UALMH"    'HERRAMIENTAS
'            codigo1 = "UALMH"
'        Case Else
'            VAR_UNI = parametro
'   End Select
  dtc_codigo1.Text = parametro
  If VAR_VAL = "OK" Then
    VAR_UNI = parametro ' dtc_codigo1.Text  Corregir
    VAR_COD4 = Ado_datos.Recordset!unidad_codigo
    If VAR_SW = "ADD" Then
        Ado_datos.Recordset!unidad_codigo = parametro
        VAR_COD4 = Ado_datos.Recordset!unidad_codigo
        'VAR_UNI = dtc_codigo1.Text
        'var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        'SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & var_cod & "  "
        SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
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
'     If VAR_SW = "MOD" Then
'        var_cod = rs_datos!solicitud_codigo
'     End If
     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
     rs_datos!solicitud_tipo = IIf(dtc_codigo2.Text = "", "1", dtc_codigo2.Text)
     Select Case parametro
        Case "DVTA"        'INI COMERCIAL
            rs_datos!solicitud_tipo = 3
        Case "COMEX"        'INI COMEX
           rs_datos!solicitud_tipo = 3
        Case "DNINS"                        'INI GRABA INSTALACIONES
            '
            rs_datos!solicitud_tipo = 4
        Case "DNAJS"
            '
            rs_datos!solicitud_tipo = 4
        Case "DNMAN"
            'UALMR
            rs_datos!solicitud_tipo = 4
        Case "UALMR", "UALMI", "UALMH"
'            dtc_codigo2.Text = 1
             rs_datos!solicitud_tipo = IIf(dtc_codigo2.Text = "", "1", "1")
           
        Case Else
            rs_datos!solicitud_tipo = 5
    End Select
     rs_datos!edif_codigo = dtc_codigo3.Text
     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
        rs_datos!beneficiario_codigo = dtc_aux3.Text
     Else
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     End If
     rs_datos!solicitud_justificacion = Txt_descripcion.Text
     
     Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
            Select Case parametro
                Case "UALMI"    'INSUMOS
                    rs_datos!etapa_codigo = "TEC-06-01"
                    rs_datos!doc_codigo = "R-126"
                    rs_datos!unidad_codigo_sol = "UALMI"
                Case "UALMR"    'REPUESTOS
                    rs_datos!etapa_codigo = "TEC-07-01"
                    rs_datos!doc_codigo = "R-126"
                      rs_datos!unidad_codigo_sol = "UALMR"
                Case "UALMH"    'HERRAMIENTAS
                    rs_datos!unidad_codigo_sol = "TEC-08-01"
                    rs_datos!doc_codigo = "R-126"
                    rs_datos!unidad_codigo_sol = "UALMH"
            End Select
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
            If VAR_UNI = "DNREP" Then
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
        Case Else
            Select Case parametro
                Case "UALMI"    'INSUMOS
                    rs_datos!etapa_codigo = "TEC-06-01"
                    rs_datos!doc_codigo = "R-126"
                   
                Case "UALMR"    'REPUESTOS
                    rs_datos!etapa_codigo = "TEC-06-01"
                    rs_datos!doc_codigo = "R-126"
                    
                Case "UALMH"    'HERRAMIENTAS
                    rs_datos!etapa_codigo = "TEC-06-01"
                    rs_datos!doc_codigo = "R-126"
                    
            End Select
     End Select
     rs_datos!poa_codigo = dtc_codigo10.Text
     rs_datos!solicitud_observaciones = txt_obs.Text
     rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
     
     rs_datos!ges_gestion_ant = glGestion       'Year(Date)
     If VAR_SW = "MOD" Then
        var_cod = rs_datos!solicitud_codigo
     End If
     If var_cod < 10 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00000" + Trim(txt_codigo)
     End If
     If var_cod > 9 And var_cod < 100 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0000" + Trim(txt_codigo)
     End If
     If var_cod > 99 And var_cod < 1000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-000" + Trim(txt_codigo)
     End If
     If var_cod > 999 And var_cod < 10000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00" + Trim(txt_codigo)
     End If
     If var_cod > 9999 And var_cod < 100000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0" + Trim(txt_codigo)
     End If
     If var_cod > 99999 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-" + Trim(txt_codigo)
     End If
'     rs_datos!solicitud_codigo_ant = 0
     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_aprueba = Date
     rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     rs_datos!unidad_codigo = VAR_COD4
     rs_datos!Fecha_Registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update 'Batch   adAffectAll
    
'     If Ado_datos.Recordset!estado_cotiza = "REG" Then
'        Call OptFilGral1_Click
'     Else
'        Call OptFilGral2_Click
'     End If
     rs_datos.MoveLast
     mbDataChanged = False
          
     Fra_datos.Enabled = False
     fra_opciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     
        Call OptFilGral1_Click
       dtc_desc3.Locked = True
'     dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
'     dtc_codigo9.Enabled = True
'     Call OptFilGral1_Click
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
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
  If (dtc_codigo8.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'    If (dtc_codigo9.Text = "") Then
'      MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'      VAR_VAL = "ERR"
'      Exit Sub
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
Fra_reporte.Visible = True
fra_opciones.Enabled = False
dg_datos.Enabled = False
Ado_datos.Enabled = False
Fra_datos.Enabled = False
FrmABMDet2.Enabled = False
dg_det2.Enabled = False
FrmABMDet.Enabled = False
dg_det1.Enabled = False
FrmABMDet3.Enabled = False
dg_det3.Enabled = False
dtc_desc_dpto.Enabled = True
dtc_desc_unidad.Enabled = False
dtc_desc_z_piloto.Enabled = False
'LIMPIAR COMBOS
dtc_desc_dpto.Text = ""
dtc_desc_unidad.Text = ""
dtc_desc_z_piloto.Text = ""
End Sub

Private Sub BtnImprimir1_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_identificacion_cliente.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
        CR01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnModDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And rs_datos!estado_cotiza = "REG" And Ado_detalle1.Recordset.RecordCount > 0 Then
    swnuevo = 2
    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
   FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
        
    frm_ao_solicitud_bitacora.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
    frm_ao_solicitud_bitacora.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
    frm_ao_solicitud_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
    frm_ao_solicitud_bitacora.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
    'frm_ao_solicitud_bitacora.Txt_estado.Caption = "REG"
    'Ado_detalle1.Recordset.AddNew
     
    frm_ao_solicitud_bitacora.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("negocia_forma")
    frm_ao_solicitud_bitacora.DTPfecha1.Value = Me.Ado_detalle1.Recordset("negocia_fecha_real")
    frm_ao_solicitud_bitacora.Txt_campo2.Value = Me.Ado_detalle1.Recordset("negocia_hora_real")
    frm_ao_solicitud_bitacora.Txt_monto1.Text = Me.Ado_detalle1.Recordset("negocia_gasto_estimado")
    frm_ao_solicitud_bitacora.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo")
    frm_ao_solicitud_bitacora.dtc_codigo3.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo_resp")
    frm_ao_solicitud_bitacora.Txt_campo3.Text = Me.Ado_detalle1.Recordset("negocia_tarea_realizada")
    frm_ao_solicitud_bitacora.Txt_campo4.Text = Me.Ado_detalle1.Recordset("negocia_observaciones")
    frm_ao_solicitud_bitacora.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bitacora_cite")
    If swnuevo = 2 Then
        frm_ao_solicitud_bitacora.dtc_desc1.BoundText = frm_ao_solicitud_bitacora.dtc_codigo1.BoundText
'        frm_ao_solicitud_bitacora.dtc_desc2.BoundText = frm_ao_solicitud_bitacora.dtc_codigo2.BoundTextdtc_desc10
        frm_ao_solicitud_bitacora.dtc_desc3.BoundText = frm_ao_solicitud_bitacora.dtc_codigo3.BoundText
'    frm_ao_solicitud_bitacora.HH = Left(frm_ao_solicitud_bitacora.Txt_campo2.Text, 2)
'        frm_ao_solicitud_bitacora.MM = Right(frm_ao_solicitud_bitacora.Txt_campo2.Text, 2)
    End If
    
    frm_ao_solicitud_bitacora.Show vbModal
    
    Call ABRIR_TABLA_DET
    
    swnuevo = 0
    fra_opciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
    Call OptFilGral1_Click
  Else
    MsgBox "No se puede Modificar el registro, verifique si está Aprobado o fue correctamente identificado !! ", vbExclamation
  End If

End Sub

Private Sub BtnModDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And rs_datos!estado_cotiza = "REG" Then
    swnuevo = 2
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False

'    Select Case dtc_codigo2.Text
'        Case "1"    'SOLO COMPRAS BB y SS
'            If VAR_DET = "30000" Then
''                marca1 = Ado_detalle2.Recordset.Bookmark
'                aw_solicitud_bienes_insumos.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
'                aw_solicitud_bienes_insumos.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
'                aw_solicitud_bienes_insumos.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                aw_solicitud_bienes_insumos.lbl_edif.Caption = dtc_codigo3.Text
'                aw_solicitud_bienes_insumos.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
'
'
'                aw_solicitud_bienes_insumos.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("bien_codigo")
'
'                aw_solicitud_bienes_insumos.dtc_desc1.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
'                aw_solicitud_bienes_insumos.dtc_aux1.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
'                aw_solicitud_bienes_insumos.dtc_aux2.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
'                aw_solicitud_bienes_insumos.dtc_aux3.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
'                aw_solicitud_bienes_insumos.Txt_campo2.Text = Me.Ado_detalle2.Recordset("marca_codigo")
'                aw_solicitud_bienes_insumos.Txt_campo3.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
'                aw_solicitud_bienes_insumos.Txt_campo4.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
'                aw_solicitud_bienes_insumos.Txt_campo18.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
''
'                aw_solicitud_bienes_insumos.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion), "-", Me.Ado_detalle2.Recordset!bien_descripcion)
'                aw_solicitud_bienes_insumos.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle2.Recordset!bien_descripcion_anterior)
'                aw_solicitud_bienes_insumos.Txt_campo8.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!marca_codigo), "S/M", Me.Ado_detalle2.Recordset!marca_codigo)
'                aw_solicitud_bienes_insumos.Txt_campo9.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!modelo_codigo), "S/M", Me.Ado_detalle2.Recordset!modelo_codigo)
'
'                aw_solicitud_bienes_insumos.Txt_campo16.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_cantidad), "1", Me.Ado_detalle2.Recordset!bien_cantidad)
'                aw_solicitud_bienes_insumos.Txt_campo10.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_precio_venta_base), "0", Me.Ado_detalle2.Recordset!bien_precio_venta_base)
'                aw_solicitud_bienes_insumos.Txt_campo11.Caption = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_total_venta), "0", Me.Ado_detalle2.Recordset!bien_total_venta)
'
'                aw_solicitud_bienes_insumos.Txt_campo14.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!unimed_codigo), "1", Me.Ado_detalle2.Recordset!unimed_codigo)
''               aw_solicitud_bienes_insumos.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
'                aw_solicitud_bienes_insumos.Txt_campo15.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!fosa_dimension_frente), "0", Me.Ado_detalle2.Recordset!fosa_dimension_frente)
''               aw_solicitud_bienes_insumos.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
'                aw_solicitud_bienes_insumos.dtc_codigo2.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!unimed_codigo), "1", Me.Ado_detalle2.Recordset!unimed_codigo)
'
'                aw_solicitud_bienes_insumos.dtc_desc2.BoundText = aw_solicitud_bienes_insumos.dtc_codigo2.BoundText
'                aw_solicitud_bienes_insumos.lbl_det.Caption = VAR_DET
'                aw_solicitud_bienes_insumos.Show vbModal
'                'Ado_detalle2.Recordset.Move marca1 - 1
'            End If
'            If VAR_DET = "39800" Then
'                frm_solicitud_bienes5A.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
'                frm_solicitud_bienes5A.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes5A.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                frm_solicitud_bienes5A.lbl_edif.Caption = dtc_codigo3.Text
'                frm_solicitud_bienes5A.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
'
'                frm_solicitud_bienes5A.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion), "-", Me.Ado_detalle2.Recordset!bien_descripcion)
'                frm_solicitud_bienes5A.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle2.Recordset!bien_descripcion_anterior)
'                frm_solicitud_bienes5A.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
'                frm_solicitud_bienes5A.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
'
'                frm_solicitud_bienes5A.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
'                frm_solicitud_bienes5A.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
'                frm_solicitud_bienes5A.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
'
'                frm_solicitud_bienes5A.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
'                frm_solicitud_bienes5A.Txt_campo15.Text = Me.Ado_detalle2.Recordset("fosa_dimension_frente")
'
'                frm_solicitud_bienes5A.lbl_det.Caption = VAR_DET
'                frm_solicitud_bienes5A.Show vbModal
'            End If
'            If VAR_DET = "34800" Then
'                frm_solicitud_bienes6.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
'                frm_solicitud_bienes6.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                frm_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
'                frm_solicitud_bienes6.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
'
'                frm_solicitud_bienes6.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion), "-", Me.Ado_detalle2.Recordset!bien_descripcion)
'                frm_solicitud_bienes6.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle2.Recordset!bien_descripcion_anterior)
'                frm_solicitud_bienes6.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
'                frm_solicitud_bienes6.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
'
'                frm_solicitud_bienes6.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
'                frm_solicitud_bienes6.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
'                frm_solicitud_bienes6.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
'
'                frm_solicitud_bienes6.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
'                frm_solicitud_bienes6.Txt_campo15.Text = Me.Ado_detalle2.Recordset("fosa_dimension_frente")
'
'                frm_solicitud_bienes6.lbl_det.Caption = VAR_DET
'                frm_solicitud_bienes6.Show vbModal
'            End If
'
'        Case "2"    'SOLO VENTA DE BIENES
'        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'            Call ABRIR_TABLA_DET
'            mw_solicitud_edificacion.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'            mw_solicitud_edificacion.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'            mw_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            'mw_solicitud_edificacion.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
'            'mw_solicitud_edificacion.Txt_estado.Caption = "REG"
'            mw_solicitud_edificacion.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("edif_codigo")
'            mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'
'            mw_solicitud_edificacion.Txt_campo2.Text = Me.Ado_detalle1.Recordset("edif_area_total_m2")
'            mw_solicitud_edificacion.Txt_campo3.Text = Me.Ado_detalle1.Recordset("edif_area_util_m2")
'            mw_solicitud_edificacion.Txt_campo4.Text = Me.Ado_detalle1.Recordset("edif_num_pisos")
'            mw_solicitud_edificacion.Txt_campo5.Text = Me.Ado_detalle1.Recordset("edif_num_salas_may_200m")
'            mw_solicitud_edificacion.Txt_campo6.Text = Me.Ado_detalle1.Recordset("edif_num_salas_men_200m")
'            mw_solicitud_edificacion.Txt_campo7.Text = Me.Ado_detalle1.Recordset("edif_num_habit_libres")
'            mw_solicitud_edificacion.Txt_campo8.Text = Me.Ado_detalle1.Recordset("edif_num_habit_ocupadas")
'            mw_solicitud_edificacion.Txt_campo9.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_2")
'            mw_solicitud_edificacion.Txt_campo10.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_3")
'            mw_solicitud_edificacion.Txt_campo11.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_4")
'            mw_solicitud_edificacion.Txt_campo12.Caption = Me.Ado_detalle1.Recordset("edif_indicador_min_trafico")
'            mw_solicitud_edificacion.Txt_campo13.Caption = Me.Ado_detalle1.Recordset("edif_capacidad_min_trafico")
'
'            mw_solicitud_edificacion.Show vbModal
'        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'        Case "5"    ' SERVICIO MODERNIZACION
'
'    End Select
aw_solicitud_bienes_insumos.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                aw_solicitud_bienes_insumos.txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                aw_solicitud_bienes_insumos.Txt_descripcion.Caption = Me.dtc_desc1.Text
                
                aw_solicitud_bienes_insumos.lbl_edif.Caption = dtc_codigo3.Text
                aw_solicitud_bienes_insumos.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                aw_solicitud_bienes_insumos.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                aw_solicitud_bienes_insumos.dtc_desc5.BoundText = aw_solicitud_bienes_insumos.dtc_codigo5.BoundText
                
                aw_solicitud_bienes_insumos.dtc_desc1.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                aw_solicitud_bienes_insumos.dtc_aux1.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                aw_solicitud_bienes_insumos.dtc_aux2.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                aw_solicitud_bienes_insumos.dtc_aux3.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                aw_solicitud_bienes_insumos.Txt_campo2.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                aw_solicitud_bienes_insumos.Txt_campo3.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                aw_solicitud_bienes_insumos.Txt_campo4.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                aw_solicitud_bienes_insumos.Txt_campo18.BoundText = aw_solicitud_bienes_insumos.dtc_codigo1.BoundText
                
                aw_solicitud_bienes_insumos.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion), "-", Me.Ado_detalle2.Recordset!bien_descripcion)
                aw_solicitud_bienes_insumos.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle2.Recordset!bien_descripcion_anterior)
                aw_solicitud_bienes_insumos.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                aw_solicitud_bienes_insumos.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                aw_solicitud_bienes_insumos.Txt_campo14.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!unimed_codigo), "1", Me.Ado_detalle2.Recordset!unimed_codigo)
                aw_solicitud_bienes_insumos.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
'                aw_solicitud_bienes_insumos.Txt_campo15.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!fosa_dimension_frente), "0", Me.Ado_detalle2.Recordset!fosa_dimension_frente)
                
                aw_solicitud_bienes_insumos.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                aw_solicitud_bienes_insumos.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_compra")
                aw_solicitud_bienes_insumos.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_compra")
                
                aw_solicitud_bienes_insumos.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                aw_solicitud_bienes_insumos.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                aw_solicitud_bienes_insumos.dtc_desc2.BoundText = aw_solicitud_bienes_insumos.dtc_codigo2.BoundText
              
                aw_solicitud_bienes_insumos.Txt_campo15.Text = Me.Ado_detalle2.Recordset("fosa_dimension_frente")
                
                aw_solicitud_bienes_insumos.lbl_det.Caption = VAR_DET
                   aw_solicitud_bienes_insumos.Show vbModal

    swnuevo = 0
    fra_opciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
    Call ABRIR_TABLA_DET
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
     FrmABMDet2.Enabled = True
  End If

End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_cotiza = "REG" Then
        Fra_datos.Enabled = True
        fra_opciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
'        dtc_desc3.Locked = True
'        dtc_desc3.Enabled = True
        
        VAR_SW = "MOD"
        DTPfecha1.Value = Ado_datos.Recordset!solicitud_fecha_solicitud
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        dtc_desc4.SetFocus
    '    BtnVer.Visible = True
'        dtc_codigo9.Enabled = False
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


Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_cod_dpto_Click(Area As Integer)
 dtc_desc_dpto.BoundText = dtc_cod_dpto.BoundText
End Sub

Private Sub dtc_cod_unidad_Click(Area As Integer)
dtc_desc_unidad.BoundText = dtc_cod_unidad.BoundText
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
    dtc_Aux11.BoundText = dtc_codigo11.BoundText
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

Private Sub dtc_desc_dpto_Click(Area As Integer)
  dtc_cod_dpto.BoundText = dtc_desc_dpto.BoundText
  dtc_desc_unidad.Enabled = True
 
End Sub

Private Sub dtc_desc_dpto_LostFocus()
'unidad Ejecutora
          Set rs_datos14 = New ADODB.Recordset
    If rs_datos14.State = 1 Then rs_datos14.Close
    rs_datos14.Open "gv_unidad_ejecutora_insumos where depto_codigo = '" & dtc_cod_dpto & "' ", db, adOpenStatic
    Set ado_datos14.Recordset = rs_datos14
    dtc_desc_unidad.BoundText = dtc_cod_unidad.BoundText
    
         'Zonas Pilotos
          Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    rs_datos15.Open "tv_zonas_piloto_insumos where depto_codigo = '" & dtc_cod_dpto & "' ", db, adOpenStatic
    Set ado_datos15.Recordset = rs_datos15
    dtc_desc_z_piloto.BoundText = dtc_cod_z_piloto.BoundText
End Sub

Private Sub dtc_desc_unidad_Click(Area As Integer)
dtc_cod_unidad.BoundText = dtc_desc_unidad.BoundText
dtc_desc_z_piloto.Enabled = True
End Sub

Private Sub dtc_desc_z_piloto_Click(Area As Integer)
dtc_cod_z_piloto.BoundText = dtc_desc_z_piloto.BoundText
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
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub
   
Private Sub pnivel1(codigo1 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
   Select Case parametro
        Case "UALMI"    'INSUMOS
            codigo1 = "UALMI"
        Case "UALMR"    'REPUESTOS
            codigo1 = "UALMR"
        Case "UALMH"    'HERRAMIENTAS
            codigo1 = "UALMH"
        Case Else
            codigo1 = parametro
   End Select
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

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
    Call pnivel1(parametro)
    dtc_desc10.Enabled = True
  
End Sub
 
Private Sub dtc_desc3_LostFocus()
    dtc_codigo4.Text = dtc_aux3.Text
    Txt_descripcion.Text = lbl_titulo + " - " + dtc_desc3.Text
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
   dtc_codigo4.BoundText = dtc_desc4.BoundText
     dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
    
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
    'parametro = "DVTA"
    '
   db.Execute "UPDATE ao_solicitud set unidad_codigo_sol = 'UALMI' from ao_solicitud INNER JOIN ao_solicitud_bienes ON ao_solicitud.unidad_codigo = ao_solicitud_bienes.unidad_codigo AND  ao_solicitud.solicitud_codigo = ao_solicitud_bienes.solicitud_codigo " & _
   "WHERE     (ao_solicitud_bienes.grupo_codigo = '30000') AND (ao_solicitud_bienes.par_codigo <> '39810') AND (ao_solicitud_bienes.par_codigo <> '39820') AND (estado_cotiza = 'REG') AND (ao_solicitud_bienes.par_codigo <> '34800') "
   
   db.Execute " UPDATE ao_solicitud set unidad_codigo_sol = 'UALMI' WHERE unidad_codigo = 'UALMI'"
   
  db.Execute " UPDATE ao_solicitud set solicitud_tipo = '1' WHERE solicitud_tipo  IS NULL"
   

    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    FrmABMDet.Visible = False
    FraDet1.Visible = False
   Command4.Visible = False
   
   FrmABMDet3.Visible = False

    
    'JQA 2014-JUL-14
    'db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'gc_tipo_solicitud
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_tipo_solicitud order by solicitud_tipo", db, adOpenStatic
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic'
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_edificaciones order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
'    Set rs_datos5 = New ADODB.Recordset
'    If rs_datos5.State = 1 Then rs_datos5.Close
'    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
'    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
'    Set Ado_datos5.Recordset = rs_datos5
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
'
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    'rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
'    rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
'    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'
'    Set rs_datos7 = New ADODB.Recordset
'    If rs_datos7.State = 1 Then rs_datos7.Close
'    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
'    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
'    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'
'    Set rs_datos8 = New ADODB.Recordset
'    If rs_datos8.State = 1 Then rs_datos8.Close
'    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
'    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
'    Set Ado_datos8.Recordset = rs_datos8
''    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
'    'gc_documentos_respaldo
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
'    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
'    Set Ado_datos9.Recordset = rs_datos9
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
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
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    'rs_aux2.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Select Case parametro
        Case "UALMI"    'INSUMOS
            rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39800' and par_codigo <> '34800'))  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and grupo_codigo = '30000' and (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800') ", db, adOpenKeyset, adLockOptimistic, adCmdText
            VAR_DET = "30000"
        Case "UALMR"    'REPUESTOS
            rs_aux2.Open "select * from av_solicitud_bienes3 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and (par_codigo = '39810' or par_codigo = '39820')   ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and (par_codigo = '39800' )  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            VAR_DET = "39800"
        Case "UALMH"    'HERRAMIENTAS
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            VAR_DET = "34800"
    End Select
    
    If rs_aux2.RecordCount > 0 Then
    
        Set Ado_detalle2.Recordset = rs_aux2
        Set dg_det2.DataSource = Ado_detalle2.Recordset
        dg_det2.Visible = True
    Else
         Set dg_det2.DataSource = rsNada
          dg_det2.Visible = False
    End If
    
        'Departamento
      Set rs_datos13 = New ADODB.Recordset
    If rs_datos13.State = 1 Then rs_datos13.Close
    rs_datos13.Open "gc_departamento", db, adOpenStatic
    Set Ado_datos13.Recordset = rs_datos13
    dtc_desc_dpto.BoundText = dtc_cod_dpto.BoundText
    
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
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
    'Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    'Image2 = Img_Foto
'    If Ado_datos.Recordset!archivo_foto_cargado = "S" Then
'        'BtnVer.Visible = True
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    Else
'        'BtnVer.Visible = False
'        'chkEstado.Value = vbUnchecked
'    End If
    If dtc_codigo2.Text = "" Then
        dtc_codigo2.Text = "1"
    End If
    If VAR_SW <> "ADD" Then
        Select Case dtc_codigo2.Text        'rs_datos!solicitud_tipo
            Case "1"    'SOLO COMPRAS BB y SS
                'DTPfecha1.Value = Ado_datos.Recordset!solicitud_fecha_solicitud
                Call ABRIR_TABLA_DET
            Case "2"    'SOLO VENTA DE BIENES
            Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
                Call ABRIR_TABLA_DET
            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
                Call ABRIR_TABLA_DET
            Case "5"    ' SERVICIO MODERNIZACION
            
            Case Else
                Call ABRIR_TABLA_DET
        End Select
        Call ABRIR_TABLA_AUX2
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    FraDet1.Caption = "BITÁCORA DE: " + dtc_desc1.Text
'    txt_aux9.Text = dtc_desc9.Text
    If Ado_datos.Recordset!estado_cotiza = "REG" Then
         FrmABMDet2.Enabled = True
         FrmABMDet2.Visible = True
    Else
            FrmABMDet2.Visible = False
    End If
  End If
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
    VAR_SW = "ADD"
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Enabled = True
    fra_opciones.Visible = False
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
    dtc_desc3.Locked = False
    dtc_desc3.backColor = &HFFFFFF
     Text3.Visible = False
     dtc_desc3.ForeColor = &H0&
'    Select Case parametro
'        Case "DVTA"        'INI COMERCIAL
'            dtc_codigo2.Text = 3
'        Case "COMEX"        'INI COMEX
'            dtc_codigo2.Text = 3
'        Case "DNINS"                        'INI GRABA INSTALACIONES
'            '
'            dtc_codigo2.Text = 4
'        Case "DNAJS"
'            '
'            dtc_codigo2.Text = 4
'        Case "DNMAN"
'            'UALMR
'            dtc_codigo2.Text = 4
'        Case "UALMR", "UALMI", "UALMH"
'            dtc_codigo2.Text = 1
'        Case Else
'            dtc_codigo2.Text = 5
'    End Select
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

Private Function ExisteReg(Unidad As String, Codigo As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_compra_cabecera WHERE unidad_codigo = '" & Unidad & "' and solicitud_codigo=" & Codigo & " and estado_codigo= 'APR'    "
'    <> 'ANL'
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub Fra_reporte_DragDrop(Source As Control, x As Single, Y As Single)
dtc_desc_z_piloto.BoundText = dtc_cod_z_piloto.BoundText

End Sub
Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case parametro
        Case "UALMI"    'INSUMOS
            queryinicial = "Select * from ao_solicitud where estado_cotiza = 'REG' AND unidad_codigo_sol = '" & parametro & "'    "
           
        Case "UALMR"    'REPUESTOS
            'queryinicial = "Select * from av_solicitud_repuestos where estado_cotiza = 'REG' "     'AND unidad_codigo = '" & parametro & "'
            queryinicial = "Select * from ao_solicitud where estado_cotiza = 'REG' AND unidad_codigo = '" & parametro & "'"
        Case "UALMH"    'HERRAMIENTAS
            queryinicial = "Select * from av_solicitud_herramientas where estado_cotiza = 'REG' "     'AND unidad_codigo = '" & parametro & "'
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case parametro
        Case "UALMI"    'INSUMOS
            queryinicial = "Select * from ao_solicitud where unidad_codigo_sol = '" & parametro & "'"
        Case "UALMR"    'REPUESTOS
            queryinicial = "Select * from av_solicitud_repuestos "
        Case "UALMH"    'HERRAMIENTAS
            queryinicial = "Select * from av_solicitud_herramientas "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub







Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

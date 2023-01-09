VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form mw_solicitud 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Identificacion del Cliente (Oportunidades de Negocio)"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "mw_solicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame fra_reportes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Elija una de las opciones ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   6000
      TabIndex        =   98
      Top             =   3360
      Visible         =   0   'False
      Width           =   5895
      Begin VB.OptionButton ob_opcion4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OTROS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   102
         Top             =   1560
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton ob_opcion3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NEGOCIACIONES AGRUPADAS POR ZONAS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   101
         Top             =   1200
         Width           =   4215
      End
      Begin VB.OptionButton ob_opcion1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NEGOCIACIONES AGRUPADAS POR REGIONAL"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   100
         Top             =   480
         Width           =   4095
      End
      Begin VB.OptionButton ob_opcion2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NEGOCIACIONES AGRUPADAS POR VENDEDOR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   99
         Top             =   840
         Width           =   4215
      End
      Begin Crystal.CrystalReport Cr_otros 
         Left            =   120
         Top             =   600
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
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   120
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   87
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "mw_solicitud.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "mw_solicitud.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   95
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
         Left            =   0
         Picture         =   "mw_solicitud.frx":104E
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   94
         ToolTipText     =   "Adiciona un Nuevo Registro"
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
         Picture         =   "mw_solicitud.frx":180D
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   93
         ToolTipText     =   "Modifica el Registro Seleccionado"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "mw_solicitud.frx":2122
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   92
         ToolTipText     =   "Anula el Registro Seleccionado"
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
         Picture         =   "mw_solicitud.frx":286E
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   91
         ToolTipText     =   "Aprueba el Registro Seleccionado"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "mw_solicitud.frx":30A1
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   90
         ToolTipText     =   "Buscar Registros"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         Picture         =   "mw_solicitud.frx":3856
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   89
         ToolTipText     =   "Imprime Lista de Registros"
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
         Picture         =   "mw_solicitud.frx":4123
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   88
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROYECTOS"
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
         Left            =   12990
         TabIndex        =   97
         Top             =   240
         Width           =   1545
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
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4515
         Picture         =   "mw_solicitud.frx":48E5
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   85
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3240
         Picture         =   "mw_solicitud.frx":51D1
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   84
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROYECTOS"
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
         Left            =   12945
         TabIndex        =   86
         Top             =   195
         Width           =   1545
      End
   End
   Begin VB.CommandButton BtnImprimir1 
      BackColor       =   &H80000015&
      Height          =   620
      Left            =   1530
      Picture         =   "mw_solicitud.frx":59A7
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Formulario de Datos para la Cotización y Cálculo de Tráfico"
      Top             =   6620
      Width           =   1365
   End
   Begin VB.CommandButton BtnImprimir2 
      BackColor       =   &H80000015&
      Height          =   620
      Left            =   1530
      Picture         =   "mw_solicitud.frx":642A
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Imprime Nota de Venta"
      Top             =   8505
      Width           =   1365
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   120
      ScaleHeight     =   1815
      ScaleWidth      =   2775
      TabIndex        =   60
      Top             =   7470
      Width           =   2835
      Begin VB.CommandButton BtnEliminar2 
         BackColor       =   &H80000015&
         Height          =   620
         Left            =   40
         Picture         =   "mw_solicitud.frx":6D62
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   1005
         Width           =   1365
      End
      Begin VB.CommandButton BtnAñadir2 
         BackColor       =   &H80000015&
         Height          =   620
         Left            =   40
         Picture         =   "mw_solicitud.frx":74AE
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Adiciona Detalle"
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton BtnModificar2 
         BackColor       =   &H80000015&
         Height          =   620
         Left            =   1380
         Picture         =   "mw_solicitud.frx":7C6D
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   1540
      Left            =   120
      ScaleHeight     =   1485
      ScaleWidth      =   2775
      TabIndex        =   56
      Top             =   5790
      Width           =   2835
      Begin VB.CommandButton BtnEliminar1 
         BackColor       =   &H80000015&
         Height          =   620
         Left            =   40
         Picture         =   "mw_solicitud.frx":8582
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   800
         Width           =   1365
      End
      Begin VB.CommandButton BtnAñadir1 
         BackColor       =   &H80000015&
         Height          =   620
         Left            =   40
         Picture         =   "mw_solicitud.frx":8CCE
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Adiciona Detalle"
         Top             =   100
         Width           =   1365
      End
      Begin VB.CommandButton BtnModificar1 
         BackColor       =   &H80000015&
         Height          =   620
         Left            =   1380
         Picture         =   "mw_solicitud.frx":948D
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   100
         Width           =   1365
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BITACORA DE NEGOCIACIONES"
      ForeColor       =   &H00C00000&
      Height          =   1980
      Left            =   2925
      TabIndex        =   38
      Top             =   7380
      Width           =   15015
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "mw_solicitud.frx":9DA2
         Height          =   1695
         Left            =   75
         TabIndex        =   61
         Top             =   240
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
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
            Caption         =   "Tipo Negoc."
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
            DataField       =   "bitacora_cite"
            Caption         =   "CITE"
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
            DataField       =   "negocia_fecha_real"
            Caption         =   "Fecha Negoc."
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
            DataField       =   "negocia_hora_real"
            Caption         =   "Hora Negoc."
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
         BeginProperty Column06 
            DataField       =   "beneficiario_nombre_ref"
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
         BeginProperty Column07 
            DataField       =   "beneficiario_codigo_cgi"
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               ColumnWidth     =   2505.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE LA EDIFICACION"
      ForeColor       =   &H00C00000&
      Height          =   1620
      Left            =   2925
      TabIndex        =   32
      Top             =   5700
      Width           =   15015
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "mw_solicitud.frx":9DBD
         Height          =   1335
         Left            =   75
         TabIndex        =   33
         Top             =   225
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   2355
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
            DataField       =   "edif_codigo"
            Caption         =   "Codigo Edificio"
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
            DataField       =   "edif_capacidad_min_trafico"
            Caption         =   "Cap.Min.Traf."
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
            DataField       =   "edif_area_total_m2"
            Caption         =   "Area Total mt2"
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
            DataField       =   "edif_area_util_m2"
            Caption         =   "Area Util mt2"
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
            DataField       =   "edif_num_pisos"
            Caption         =   "Nro.de.Pisos"
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
            DataField       =   "edif_num_salas_may_200m"
            Caption         =   "Nro.Salas >200mt."
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
            DataField       =   "edif_num_salas_men_200m"
            Caption         =   "Nro.Salas < 200.mt"
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
            DataField       =   "edif_num_habit_libres"
            Caption         =   "Nro.Habit.Libres"
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
         BeginProperty Column08 
            DataField       =   "edif_num_habit_ocupadas"
            Caption         =   "Nro.Hab.Ocup/1.D"
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
            DataField       =   "edif_num_habit_dorm_2"
            Caption         =   "Habit.de.2.Dorm."
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
            DataField       =   "edif_num_habit_dorm_3"
            Caption         =   "Habit.de.3.Dorm."
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
            DataField       =   "edif_num_habit_dorm_4"
            Caption         =   "Habit.>= 4.Dorm."
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
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00C00000&
      Height          =   4920
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   8175
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "mw_solicitud.frx":9DD8
         Height          =   4170
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   7355
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Tramite"
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
            DataField       =   "observacion_proy"
            Caption         =   "Nombre.Edificio"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Contrato"
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
            DataField       =   "solicitud_fecha_solicitud"
            Caption         =   "Fecha.Solic."
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
         BeginProperty Column07 
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
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3390.236
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
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
         Left            =   2040
         TabIndex        =   51
         Top             =   4520
         Value           =   -1  'True
         Width           =   1455
      End
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
         Left            =   4800
         TabIndex        =   52
         Top             =   4520
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4440
         Width           =   7905
         _ExtentX        =   13944
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
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000040&
      Height          =   4920
      Left            =   8385
      TabIndex        =   12
      Top             =   720
      Width           =   9555
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7665
         TabIndex        =   108
         Text            =   "/"
         Top             =   4320
         Width           =   270
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3345
         TabIndex        =   78
         Top             =   4155
         Visible         =   0   'False
         Width           =   280
      End
      Begin VB.CommandButton BtnAux3 
         BackColor       =   &H00C0FFFF&
         Height          =   600
         Left            =   6120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "mw_solicitud.frx":9DF0
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Registrar un NUEVO edificio."
         Top             =   2055
         Width           =   1360
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7210
         TabIndex        =   77
         Top             =   2310
         Width           =   255
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6315
         TabIndex        =   42
         Top             =   4515
         Width           =   280
      End
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "mw_solicitud.frx":AAD1
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1575
         TabIndex        =   17
         Top             =   4500
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "etapa_descripcion"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "mw_solicitud.frx":AAEA
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   4140
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "mw_solicitud.frx":AB04
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4620
         TabIndex        =   75
         Top             =   2715
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "mw_solicitud.frx":AB1E
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   74
         Top             =   2295
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   ""
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   5895
         TabIndex        =   55
         Top             =   495
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "mw_solicitud.frx":AB37
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   480
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "unidad_codigo_ant"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6480
         TabIndex        =   76
         Text            =   "0"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame fra_cliente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CLIENTE (Registra una de las 3 alternativas)"
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
         Height          =   1120
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   9255
         Begin VB.TextBox txt_ci 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   4800
            TabIndex        =   73
            Top             =   165
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton BtnAux2 
            BackColor       =   &H00C0FFFF&
            Height          =   580
            Left            =   7690
            MaskColor       =   &H00FFFFFF&
            Picture         =   "mw_solicitud.frx":AB50
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   170
            Width           =   1380
         End
         Begin VB.TextBox txt_obs 
            BackColor       =   &H00FFFFFF&
            DataField       =   "solicitud_observaciones"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   720
            Width           =   5250
         End
         Begin VB.TextBox txt_nombre 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   3075
            TabIndex        =   68
            Top             =   60
            Visible         =   0   'False
            Width           =   4455
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "mw_solicitud.frx":B927
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3045
            TabIndex        =   67
            Top             =   285
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "mw_solicitud.frx":B940
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6000
            TabIndex        =   72
            Top             =   285
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "3. Datos Referenciales (Nombre, Telef...)"
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
            TabIndex        =   69
            Top             =   720
            Width           =   3630
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "1. Existente en la Base de Datos"
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
            TabIndex        =   66
            Top             =   280
            Width           =   2880
         End
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_justificacion"
         DataSource      =   "Ado_datos"
         Height          =   480
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   3165
         Width           =   7905
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2385
         TabIndex        =   54
         Top             =   3920
         Width           =   350
      End
      Begin MSDataListLib.DataCombo dtc_aux11 
         Bindings        =   "mw_solicitud.frx":B959
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6480
         TabIndex        =   44
         Top             =   2715
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "mw_solicitud.frx":B973
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7800
         TabIndex        =   43
         Top             =   2715
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7275
         TabIndex        =   41
         Top             =   3210
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "mw_solicitud.frx":B98D
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3360
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo3"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7275
         TabIndex        =   40
         Top             =   3210
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "mw_solicitud.frx":B9A6
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5640
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "mw_solicitud.frx":B9BF
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   25
         Top             =   1935
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
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "mw_solicitud.frx":B9D8
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2760
         TabIndex        =   24
         Top             =   4140
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "clasif_codigo"
         BoundColumn     =   "clasif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "mw_solicitud.frx":B9F1
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5160
         TabIndex        =   23
         Top             =   4500
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "etapa_codigo"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5160
         TabIndex        =   22
         Top             =   4500
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "subproceso_codigo"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         DataField       =   "proceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7800
         TabIndex        =   21
         Top             =   3240
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "proceso_codigo"
         BoundColumn     =   "proceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "mw_solicitud.frx":BA0A
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6120
         TabIndex        =   20
         Top             =   2295
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         DataField       =   "proceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2820
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "proceso_descripcion"
         BoundColumn     =   "proceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2820
         TabIndex        =   18
         Top             =   3195
         Visible         =   0   'False
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "subproceso_descripcion"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "mw_solicitud.frx":BA23
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2100
         TabIndex        =   2
         Top             =   4500
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "clasif_descripcion"
         BoundColumn     =   "clasif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "mw_solicitud.frx":BA3C
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   26
         Top             =   4320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   128
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
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
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "mw_solicitud.frx":BA55
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6480
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   -2147483629
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
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
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "mw_solicitud.frx":BA6E
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   3900
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "mw_solicitud.frx":BA87
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2940
         TabIndex        =   29
         Top             =   3945
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "mw_solicitud.frx":BAA0
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   35
         Top             =   4140
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "solicitud_fecha_solicitud"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   7600
         TabIndex        =   79
         Top             =   2295
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   118226945
         CurrentDate     =   44232
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   7920
         TabIndex        =   109
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   315
         Left            =   8790
         TabIndex        =   107
         Top             =   4320
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   240
         Left            =   7710
         TabIndex        =   106
         Top             =   4560
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lblGestion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8880
         TabIndex        =   105
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lblNeocia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "correl_bitacora"
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
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   7800
         TabIndex        =   104
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblUniSol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "unidad_codigo_sol"
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
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   6960
         TabIndex        =   103
         Top             =   4560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   8580
         TabIndex        =   13
         Top             =   225
         Width           =   645
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cite Contrato"
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
         Index           =   6
         Left            =   6765
         TabIndex        =   53
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   50
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Justificación"
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
         TabIndex        =   49
         Top             =   3285
         Width           =   1095
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Trámite (NT)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6900
         TabIndex        =   48
         Top             =   3900
         Width           =   2520
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro ISO"
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
         TabIndex        =   47
         Top             =   3900
         Width           =   1140
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Personal de CGI Reponsable de la Negociacion"
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
         TabIndex        =   46
         Top             =   2745
         Width           =   4335
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1725
         TabIndex        =   45
         Top             =   225
         Width           =   2160
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   6780
         X2              =   6780
         Y1              =   3735
         Y2              =   4915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   6780
         Y1              =   4335
         Y2              =   4335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9500
         Y1              =   3735
         Y2              =   3735
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
         TabIndex        =   37
         Top             =   480
         Width           =   1215
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
         Left            =   5460
         TabIndex        =   36
         Top             =   3900
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo ISO"
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
         Index           =   13
         Left            =   3975
         TabIndex        =   31
         Top             =   3900
         Width           =   1365
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitud"
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
         Index           =   12
         Left            =   7545
         TabIndex        =   30
         Top             =   2040
         Width           =   1620
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Etapa Proceso"
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
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Top             =   4485
         Width           =   1350
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
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
         Height          =   300
         Left            =   8595
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Num.Tramite"
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
         Left            =   180
         TabIndex        =   14
         Top             =   225
         Width           =   1155
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
      Top             =   9480
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
      Left            =   12360
      Top             =   9720
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
      Left            =   2160
      Top             =   9480
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
      Left            =   4200
      Top             =   9480
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
      Left            =   6240
      Top             =   9480
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
      Left            =   8280
      Top             =   9480
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
      Left            =   10320
      Top             =   9480
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
      Left            =   12360
      Top             =   9480
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
      Top             =   9840
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
      Left            =   2160
      Top             =   9840
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
      Left            =   4200
      Top             =   9840
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
      Left            =   6240
      Top             =   9840
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
      Left            =   8280
      Top             =   9840
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
      Left            =   10320
      Top             =   9840
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
   Begin Crystal.CrystalReport CR02 
      Left            =   12840
      Top             =   9720
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
   Begin Crystal.CrystalReport CR03 
      Left            =   13320
      Top             =   9720
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
End
Attribute VB_Name = "mw_solicitud"
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

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI, VAR_UORIGEN As String
Dim sino As String
Dim VAR_BENEF As String
Dim VAR_CITE As String
Dim VAR_ARCH As String
Dim VAR_DA As String
Dim VAR_DPTO As String

Dim VAR_AUX, VAR_CONT2 As Double

Dim VAR_VALI As Integer
Dim VAR_SOLA As Long

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAñadir1_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If

On Error GoTo AddErr
  'marca1 = Ado_datos.Recordset.Bookmark
  'If rs_datos!estado_codigo = "REG" Then
    VAR_SOLA = Ado_datos.Recordset!solicitud_codigo_ant
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    GlUnidad = Ado_datos.Recordset!unidad_codigo
    GlSolicitud = Ado_datos.Recordset!solicitud_codigo
    GlEdificio = Ado_datos.Recordset!edif_codigo
    glGestion = Ado_datos.Recordset!ges_gestion
    mw_solicitud_edificacion.Show vbModal
'    Select Case dtc_codigo2.Text
'        Case "1"
'        Case "2"
'        Case "3"
'
''            Call ABRIR_TABLA_DET3
''            Ado_detalle1.Recordset.AddNew
''            GlEdificio = Me.dtc_codigo3.Text
''            mw_solicitud_edificacion.txt_codigo.Caption = Me.txt_codigo.Caption
''            mw_solicitud_edificacion.txt_campo1.Caption = Me.dtc_codigo1.Text
''            mw_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
''            mw_solicitud_edificacion.txt_gestion.Caption = Ado_datos.Recordset!ges_gestion
''
''            mw_solicitud_edificacion.Txt_campo18.Caption = Me.dtc_codigo3.Text
''            mw_solicitud_edificacion.Txt_campo19.Caption = Trim(Me.dtc_desc3.Text)
''            mw_solicitud_edificacion.Txt_campo20.Caption = Me.dtc_aux3.Text
''
'''            mw_solicitud_edificacion.dtc_codigo1.Text = GlEdificio 'Me.dtc_codigo3.Text
'''            mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'''            mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'''            mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'''            mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
''            mw_solicitud_edificacion.Txt_estado.Caption = "REG"
''            mw_solicitud_edificacion.Show vbModal
'        Case "9"
'            Call ABRIR_TABLA_DET3
'            Ado_detalle1.Recordset.AddNew
'            GlEdificio = Me.dtc_codigo3.Text
'            mw_solicitud_edificacion.txt_codigo.Caption = Me.txt_codigo.Caption
'            mw_solicitud_edificacion.txt_campo1.Caption = Me.dtc_codigo1.Text
'            mw_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            mw_solicitud_edificacion.txt_gestion.Caption = Ado_datos.Recordset!ges_gestion
'
'            mw_solicitud_edificacion.Txt_campo18.Caption = Me.dtc_codigo3.Text
'            mw_solicitud_edificacion.Txt_campo19.Caption = Me.dtc_desc3.Text
'            mw_solicitud_edificacion.Txt_campo20.Caption = Me.dtc_aux3.Text
'
'            mw_solicitud_edificacion.Txt_estado.Caption = "REG"
'            mw_solicitud_edificacion.Show vbModal
'        Case "4"
'    End Select
    Call ABRIR_TABLA_DET3
    If OptFilGral1.Value = True Then
       Call OptFilGral1_Click        'Pendientes
    Else
       Call OptFilGral2_Click        'TODOS
    End If
    If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
    End If
    If Ado_datos.Recordset.RecordCount > 0 Then
       rs_datos.Find "solicitud_codigo_ant = " & VAR_SOLA & "   ", , , 1
       dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
    Else
       rs_datos.MoveLast
    End If
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    Fra_datos.Enabled = True
  'Else
  '  MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  'End If
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAñadir2_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If

On Error GoTo AddErr
  'marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo <> "ERR" Then
    VAR_SOLA = Ado_datos.Recordset!solicitud_codigo_ant
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    Call ABRIR_TABLA_DET3
    aw_p_ao_negociacion_bitacora.txt_codigo.Caption = Me.txt_codigo.Caption
    aw_p_ao_negociacion_bitacora.txt_campo1.Caption = Me.dtc_codigo1.Text
    aw_p_ao_negociacion_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
    aw_p_ao_negociacion_bitacora.Txt_Correl.Caption = 0
    aw_p_ao_negociacion_bitacora.Txt_estado.Caption = "REG"
    aw_p_ao_negociacion_bitacora.txt_cliente.Text = txt_obs
    Ado_detalle2.Recordset.AddNew
    aw_p_ao_negociacion_bitacora.Show vbModal
    
    Call ABRIR_TABLA_DET3
    If OptFilGral1.Value = True Then
       Call OptFilGral1_Click        'Pendientes
    Else
       Call OptFilGral2_Click        'TODOS
    End If
    If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
    End If
    If Ado_datos.Recordset.RecordCount > 0 Then
       rs_datos.Find "solicitud_codigo_ant = " & VAR_SOLA & "   ", , , 1
       dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
    Else
       rs_datos.MoveLast
    End If
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
    'Ado_datos.Recordset.Move marca1 - 1
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este fue Anulado!! ", vbExclamation
  End If
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnEliminar1_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
   
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
            Call ABRIR_TABLA_DET3
          End If
       Else
            MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
       End If
    Else
        MsgBox "No se puede ANULAR, debe identificar algun registro ...", vbExclamation, "Validación de Registro"
    End If
End Sub

Private Sub BtnEliminar2_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  
  If Ado_detalle2.Recordset.RecordCount > 0 Then
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle2.Recordset.Delete 'adAffectAll
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

Private Sub BtnAprobar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  
  On Error GoTo UpdateErr
'   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
'        MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
'        Exit Sub
'   End If
   If rs_datos!estado_codigo = "REG" Then
     Set rs_aux2 = New ADODB.Recordset
     If rs_aux2.State = 1 Then rs_aux2.Close
     rs_aux2.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & parametro & "'  and solicitud_codigo = " & GlSolicitud & "   ", db, adOpenStatic
     If rs_aux2.RecordCount > 0 Then
        Set rs_aux4 = New ADODB.Recordset
        If rs_aux4.State = 1 Then rs_aux4.Close
        rs_aux4.Open "SELECT * From ao_negociacion_bitacora WHERE unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
        If rs_aux4.RecordCount > 0 Then
            VAR_CONT2 = rs_aux2.RecordCount
            'If rs_datos!estado_codigo = "REG" And Ado_datos.Recordset!correl_edificacion > 0 Then
            sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
            If sino = vbYes Then
              Select Case dtc_codigo2.Text
                  Case "1"
                  Case "2"
                  Case "3", "9"
                      Set rs_aux1 = New ADODB.Recordset
                      'SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_detalle1.Recordset!edif_codigo & "'  "
                      'SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
                      SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
                      rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                      'If rs_aux1.RecordCount > 0 Then
                      '    MsgBox "El código ya existe, consulte con el administrador del Sistema..."
                      '    var_cod = 0
                      '    Exit Sub
                      'Else
                          Set rs_aux2 = New ADODB.Recordset
                          If rs_aux2.State = 1 Then rs_aux2.Close
                          'rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                          rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' ", db, adOpenStatic
                          If Not rs_aux2.EOF Then
                              var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                          End If
                          Set rs_aux2 = New ADODB.Recordset
                          If rs_aux2.State = 1 Then rs_aux2.Close
                          rs_aux2.Open "Select edif_capacidad_min_trafico as Codigo from ao_solicitud_edificacion where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                          If Not rs_aux2.EOF Then
                              VAR_AUX = rs_aux2!Codigo
                          End If
                          rs_aux1.AddNew
                          'var_cod = rs_aux1.RecordCount + 1
                          rs_aux1!ges_gestion = Year(Date)
                          rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                          rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                          rs_aux1!edif_codigo = Ado_detalle1.Recordset!edif_codigo
                          rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                          rs_aux1!trafico_codigo = var_cod
                          rs_aux1!h_capacidad_trafico_parametro = Round(VAR_AUX, 2)
                          rs_aux1!estado_codigo = "REG"
                          rs_aux1!fecha_registro = Date
                          rs_aux1!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
                          rs_aux1!usr_codigo = glusuario
                          rs_aux1.Update
                          db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
                      'End If
                      'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
                  Case "4"
              End Select
              Set rs_aux2 = New ADODB.Recordset
              SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
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
              rs_datos!estado_codigo = "APR"
              rs_datos!estado_cotiza = "REG"
              rs_datos!fecha_registro = Date
              rs_datos!usr_codigo = glusuario
              rs_datos.UpdateBatch adAffectAll
            End If
        Else
          MsgBox "No se puede APROBAR, registre la " + FraDet2.Caption + ", luego, vuelva a intentar...", vbExclamation, "Validación de Registro"
        End If
     Else
          MsgBox "No se puede APROBAR, registre el " + FraDet1.Caption + ", luego,  vuelva a intentar...", vbExclamation, "Validación de Registro"
     End If
   
   Else
       MsgBox "Error, el Registro ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAux3_Click()
    'Validacion 1
'    If dtc_codigo3 = "" Or dtc_codigo3 = "0" Then
'        MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
'        VAR_VAL = "ERR"
'        Exit Sub
'    End If
    glPersNew = "NEWE"
    gw_p_gc_edificaciones_aux.Show vbModal
    'Fra_ABM.Enabled = False
    'Fra_aux1.Visible = True
End Sub

Private Sub BtnAux2_Click()
    glPersNew = "NEWC"
    txt_nombre.Visible = False
    gw_p_gc_beneficiario_aux.Show vbModal
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Call OptFilGral2_Click
        buscados = 1
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False

        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
        MsgBox "NO se puede Procesar !!. Verifique si existe registro. ", vbExclamation, "Atención!"
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        VAR_SOLA = Ado_datos.Recordset!solicitud_codigo_ant
        rs_datos.CancelUpdate
        If OptFilGral1.Value = True Then
           Call OptFilGral1_Click        'Pendientes
        Else
           Call OptFilGral2_Click        'TODOS
        End If
        If (dg_datos.SelBookmarks.Count <> 0) Then
           dg_datos.SelBookmarks.Remove 0
        End If
        If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
           rs_datos.Find "solicitud_codigo_ant = " & VAR_SOLA & "   ", , , 1
           dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
        Else
           rs_datos.MoveLast
        End If
        mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        'txt_codigo.Enabled = True
        VAR_SW = ""
        BtnAux3.Visible = False
        BtnAux2.Visible = False
        dtc_codigo9.Enabled = True
        
        FrmABMDet.Visible = True
        FrmABMDet2.Visible = True
        FraDet1.Visible = True
        FraDet2.Visible = True
        BtnImprimir1.Visible = True
        BtnImprimir2.Visible = True
     
        fra_cliente.Caption = "CLIENTE"
        
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub BtnEliminar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  
  On Error GoTo UpdateErr
   'If ExisteReg(Ado_datos.Recordset!edif_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   If ExisteReg(Ado_datos.Recordset!unidad_codigo, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   If rs_datos!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ANL"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        rs_datos!estado_codigo = "ERR"
        rs_datos!fecha_registro = Date
        rs_datos!usr_codigo = glusuario
        rs_datos.UpdateBatch adAffectAll
      'MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
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
         rs_datos!fecha_registro = Date
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
  Call valida_campos
  If VAR_VAL = "OK" Then
     If VAR_SW = "ADD" Then
        VAR_UNI = dtc_codigo1.Text
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "select * from ao_solicitud where edif_codigo = '" & dtc_codigo3 & "' AND unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
           sino = MsgBox("Un proceso anterior con este EDIFICIO ya existe, si desea vender un nuevo EQUIPO elija <SI>, caso contrario elija <NO> y cambie el edificio. ", vbYesNo + vbQuestion, "Atención")
           If sino = vbYes Then
               VAR_VALI = 1
           Else
               VAR_VALI = 0
               Exit Sub
           End If
        Else
           VAR_VALI = 1
        End If
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        If VAR_VALI = 1 Then
            SQL_FOR = "Select max(solicitud_codigo) as Codigo from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            If Not rs_aux1.EOF Then
                var_cod = IIf(IsNull(rs_aux1!Codigo), 1, rs_aux1!Codigo + 1)
            Else
                var_cod = 1
            End If
            txt_codigo.Caption = var_cod
'            rs_datos!solicitud_codigo = var_cod
'            rs_datos!estado_codigo = "REG"      'no cambia
'            rs_datos!ges_gestion = Year(Date)   'no cambia
'            rs_datos!unidad_codigo = VAR_UNI
            'Actualiza correaltivo ...
            db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
'            rs_datos!doc_numero = "0"    'txt_campo1.Caption
'            'rs_datos!correl_edificacion = 0
'            rs_datos!archivo_respaldo = "sin_nombre"
'            rs_datos!archivo_respaldo_cargado = "N"
            'WWWWWWWWWWWWWWWWWWWWWWWWWW
            Set rs_aux5 = New ADODB.Recordset
            If rs_aux5.State = 1 Then rs_aux5.Close
            rs_aux5.Open "Select * from gc_unidad_ejecutora correl_negocia where unidad_codigo = '" & VAR_UNI & "' ", db, adOpenStatic
            If rs_aux5.RecordCount > 0 Then
                Select Case VAR_UNI
                        Case "DVTA"                            'LA PAZ - NACIONAL
                            'Txt_campo5 = "COM-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                            lblUniSol.Caption = "COM"
                            lblNeocia.Caption = rs_aux5!correl_negocia + 1
                            lblGestion.Caption = Year(Date)
                        Case "DCOMS"                            'SANTA CRUZ
                            'Txt_campo5 = "COMS-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                            lblUniSol.Caption = "COMS"
                            lblNeocia.Caption = rs_aux5!correl_negocia + 1
                            lblGestion.Caption = Year(Date)
                        Case "DCOMB"                            'CBBA
                            'Txt_campo5 = "COMB-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                            lblUniSol.Caption = "COMB"
                            lblNeocia.Caption = rs_aux5!correl_negocia + 1
                            lblGestion.Caption = Year(Date)
                         Case "DCOMC"                            'CHUQUISACA
                            'Txt_campo5 = "COMC-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                            lblUniSol.Caption = "COMC"
                            lblNeocia.Caption = rs_aux5!correl_negocia + 1
                            lblGestion.Caption = Year(Date)
                        Case Else
                            'Txt_campo5 = "COM-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                            lblUniSol.Caption = "COM"
                            lblNeocia.Caption = rs_aux5!correl_negocia + 1
                            lblGestion.Caption = Year(Date)
                End Select
                'CITE = 1
                db.Execute "update gc_unidad_ejecutora set correl_negocia =  correl_negocia + 1 where unidad_codigo = '" & VAR_UNI & "'  "
                'BtnGrabar2.Enabled = False
            End If
            'WWWWWWWWWWWWWWWWWWWw
            If dtc_codigo4.Text = "" Then
               VAR_BENEF = txt_ci.Text
            Else
               VAR_BENEF = dtc_codigo4.Text
            End If
            
            If Left(Txt_campo2, 4) = "36NB" Or Left(Txt_campo2, 4) = "OA36" Or Left(Txt_campo2, 4) = "36NO" Then
               rs_datos!unidad_codigo_ant = Trim(Txt_campo2.Text)
            Else
               If var_cod < 10 Then
                  rs_datos!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
               End If
               If var_cod > 9 And var_cod < 100 Then
                  rs_datos!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
               End If
               If var_cod > 99 And var_cod < 1000 Then
                  rs_datos!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
               End If
               If var_cod > 999 And var_cod < 10000 Then
                  rs_datos!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
               End If
               If var_cod > 9999 And var_cod < 100000 Then
                  rs_datos!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
               End If
               If var_cod > 99999 Then
                  rs_datos!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
               End If
            End If
            VAR_ARCH = "COM-R-234-" + Trim(txt_codigo) + ".JPG"
            If usuario2 = "0" Then
                usuario2 = dtc_codigo11.Text
            End If
            If VAR_UNI = "DNMOD" Then
                lblUniSol.Caption = "MOD"
                db.Execute "insert into ao_solicitud (ges_gestion, unidad_codigo, solicitud_codigo, solicitud_fecha_solicitud, solicitud_fecha_recepción, solicitud_tipo, edif_codigo, beneficiario_codigo, beneficiario_codigo_resp, beneficiario_codigo_resp2, unidad_codigo_sol, " & _
                " solicitud_justificacion, solicitud_observaciones, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, poa_codigo, ges_gestion_ant, unidad_codigo_ant, correl_detalle, correl_edificacion, correl_calculo, correl_persona, correl_cotiza, correl_bitacora, " & _
                " archivo_respaldo, archivo_respaldo_cargado, estado_codigo, estado_etapa2, estado_cotiza, fecha_registro,  usr_codigo, observacion_proy)  " & _
                " Values ('" & Year(Date) & "', '" & VAR_UNI & "', " & var_cod & ", '" & DTPfecha1.Value & "', '" & DTPfecha1.Value & "', " & dtc_codigo2.Text & ", '" & dtc_codigo3.Text & "', '" & VAR_BENEF & "', '" & dtc_codigo11.Text & "', '" & usuario2 & "', '" & lblUniSol.Caption & "', " & _
                " '" & Trim(Txt_descripcion.Text) & "', '" & txt_obs.Text & "', 'TEC', 'TEC-05', 'TEC-05-01', 'TEC', 'R-313', '0', '3.2.7', '" & Year(Date) & "', '" & VAR_CITE & "', '0', '0', '0', '0', '0', " & lblNeocia.Caption & ", " & _
                " '" & VAR_ARCH & "', 'N', 'REG', 'REG', 'REG', '" & Date & "', '" & glusuario & "', '" & dtc_desc3.Text & "' )"
            Else
                db.Execute "insert into ao_solicitud (ges_gestion, unidad_codigo, solicitud_codigo, solicitud_fecha_solicitud, solicitud_fecha_recepción, solicitud_tipo, edif_codigo, beneficiario_codigo, beneficiario_codigo_resp, beneficiario_codigo_resp2, unidad_codigo_sol, " & _
                " solicitud_justificacion, solicitud_observaciones, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, poa_codigo, ges_gestion_ant, unidad_codigo_ant, correl_detalle, correl_edificacion, correl_calculo, correl_persona, correl_cotiza, correl_bitacora, " & _
                " archivo_respaldo, archivo_respaldo_cargado, estado_codigo, estado_etapa2, estado_cotiza, fecha_registro,  usr_codigo, observacion_proy)  " & _
                " Values ('" & Year(Date) & "', '" & VAR_UNI & "', " & var_cod & ", '" & DTPfecha1.Value & "', '" & DTPfecha1.Value & "', " & dtc_codigo2.Text & ", '" & dtc_codigo3.Text & "', '" & VAR_BENEF & "', '" & dtc_codigo11.Text & "', '" & usuario2 & "', '" & lblUniSol.Caption & "', " & _
                " '" & Trim(Txt_descripcion.Text) & "', '" & txt_obs.Text & "', 'COM', 'COM-01', 'COM-01-02', 'COM', 'R-220', '0', '3.1.1', '" & Year(Date) & "', '" & VAR_CITE & "', '0', '0', '0', '0', '0', " & lblNeocia.Caption & ", " & _
                " '" & VAR_ARCH & "', 'N', 'REG', 'REG', 'REG', '" & Date & "', '" & glusuario & "', '" & dtc_desc3.Text & "' )"
            End If
            'VAR_SOLA = Ado_datos.Recordset!solicitud_codigo_ant
        End If
     End If
     If VAR_SW = "MOD" Then
        VAR_UNI = dtc_codigo1.Text
        VAR_SOLA = Ado_datos.Recordset!solicitud_codigo_ant
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "select * from ao_solicitud where edif_codigo = '" & dtc_codigo3 & "' AND unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 1 Then
           sino = MsgBox("Un proceso anterior con este EDIFICIO ya existe, si desea vender un nuevo EQUIPO elija <SI>, caso contrario elija <NO> y cambie el edificio. ", vbYesNo + vbQuestion, "Atención")
           If sino = vbYes Then
               VAR_VALI = 1
           Else
               VAR_VALI = 0
               Exit Sub
           End If
        Else
           VAR_VALI = 1
        End If
        var_cod = rs_datos!solicitud_codigo
        If dtc_codigo4.Text = "" Then
            VAR_BENEF = txt_ci.Text
         Else
            VAR_BENEF = dtc_codigo4.Text
         End If
         db.Execute "Update ao_solicitud Set unidad_codigo_ant = '" & Txt_campo2.Text & "', solicitud_fecha_solicitud = '" & DTPfecha1.Value & "', beneficiario_codigo = '" & VAR_BENEF & "', edif_codigo = '" & dtc_codigo3.Text & "', beneficiario_codigo_resp = '" & dtc_codigo11.Text & "', solicitud_justificacion = '" & Txt_descripcion.Text & "', fecha_registro = '" & Date & "', usr_codigo= '" & glusuario & "', observacion_proy = '" & dtc_desc3.Text & "'  Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  "
     End If
     
'     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
'     rs_datos!solicitud_tipo = dtc_codigo2.Text
'     rs_datos!edif_codigo = dtc_codigo3.Text
'     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
'        VAR_BENEF = txt_ci.Text
'        dtc_codigo4.Text = txt_ci.Text
'        rs_datos!beneficiario_codigo = txt_ci   'dtc_aux3.Text
'     Else
'        VAR_BENEF = dtc_codigo4.Text
'        rs_datos!beneficiario_codigo = dtc_codigo4.Text
'        txt_ci = dtc_codigo4.Text
'     End If
'     rs_datos!solicitud_justificacion = Trim(Txt_descripcion.Text)
'     Select Case dtc_codigo2.Text
'        Case "1"    'Solo Compras
'        Case "2"    'Solo Ventas
'        Case "3"    'Compras y Ventas de Equipos
'            rs_datos!proceso_codigo = IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
'            rs_datos!subproceso_codigo = IIf(dtc_codigo6.Text = "", "COM-01", dtc_codigo6.Text)
'            rs_datos!etapa_codigo = IIf(dtc_codigo7.Text = "", "COM-01-01", dtc_codigo7.Text)
'            rs_datos!clasif_codigo = IIf(dtc_codigo8.Text = "", "COM", dtc_codigo8.Text)
'            rs_datos!doc_codigo = IIf(dtc_codigo9.Text = "", "R-220", dtc_codigo9.Text)
'            rs_datos!doc_codigo2 = "R-233"       'IIf(dtc_codigo9.Text = "", "R-220", dtc_codigo9.Text)
'        Case "4"    ' Ventas Nuevas
'            rs_datos!proceso_codigo = IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
'            rs_datos!subproceso_codigo = IIf(dtc_codigo6.Text = "", "COM-01", dtc_codigo6.Text)
'            rs_datos!etapa_codigo = IIf(dtc_codigo7.Text = "", "COM-01-02", dtc_codigo7.Text)
'            rs_datos!clasif_codigo = IIf(dtc_codigo8.Text = "", "COM", dtc_codigo8.Text)
'            rs_datos!doc_codigo = IIf(dtc_codigo9.Text = "", "R-234", dtc_codigo9.Text)
'        Case "5"    'Modernizacion
'     End Select
'     rs_datos!poa_codigo = IIf(dtc_codigo10.Text = "", "3.1.1", dtc_codigo10.Text)
'     '5978372
'     rs_datos!solicitud_observaciones = txt_obs.Text
'     rs_datos!observacion_proy = dtc_desc3.Text
'     rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
'     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
'     rs_datos!beneficiario_codigo_resp2 = usuario2
'     rs_datos!ges_gestion_ant = Year(Date)
'     If Left(Txt_campo2, 4) = "36NO" Or Left(Txt_campo2, 4) = "OA36" Then
'        rs_datos!unidad_codigo_ant = Trim(Txt_campo2.Text)
'     Else
'        If var_cod < 10 Then
'           rs_datos!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
'        End If
'        If var_cod > 9 And var_cod < 100 Then
'           rs_datos!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
'        End If
'        If var_cod > 99 And var_cod < 1000 Then
'           rs_datos!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
'        End If
'        If var_cod > 999 And var_cod < 10000 Then
'           rs_datos!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
'        End If
'        If var_cod > 9999 And var_cod < 100000 Then
'           rs_datos!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
'        End If
'        If var_cod > 99999 Then
'           rs_datos!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
'        End If
'     End If
''     rs_datos!solicitud_codigo_ant = 0
'     rs_datos!usr_codigo_aprueba = ""
'     rs_datos!fecha_aprueba = Date
'     rs_datos!hora_aprueba = ""
'     'rs_datos!Foto = Date
'     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
'     'rs_datos!archivo_foto_cargado = "N"
'     'hora_registro
'     'rs_datos!beneficiario_codigo = txt_ci
'     rs_datos!fecha_registro = Date     'no cambia
'     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
'     rs_datos.Update    'Batch 'adAffectAll
     If dtc_codigo4.Text = "" And IsNull(dtc_codigo4.Text) Then
        db.Execute "update gc_edificaciones set beneficiario_codigo = '0' where edif_codigo = '" & dtc_codigo3.Text & "' "
     Else
        db.Execute "update gc_edificaciones set beneficiario_codigo = '" & dtc_codigo4.Text & "' where edif_codigo = '" & dtc_codigo3.Text & "' "
     End If
'     If Ado_datos.Recordset("estado_codigo") = "REG" And VAR_SW = "ADD" Then
'        Call OptFilGral1_Click
'        rs_datos.MoveLast
'     'Else
'     '   Call OptFilGral2_Click
'     End If
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
        rs_datos.Find "solicitud_codigo_ant = " & VAR_SOLA & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
'     mbDataChanged = False
     'db.Execute "Update ao_solicitud Set beneficiario_codigo = '" & txt_ci & "' Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
'     dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
     BtnAux3.Visible = False
     BtnAux2.Visible = False
     dtc_codigo9.Enabled = True
     FrmABMDet.Visible = True
     FrmABMDet2.Visible = True
     FraDet1.Visible = True
     FraDet2.Visible = True
     BtnImprimir1.Visible = True
     BtnImprimir2.Visible = True
     fra_cliente.Caption = "CLIENTE"
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
  If (dtc_codigo8.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo9.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo10.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        fra_reportes.Visible = True
        
'        Dim iResult As Integer
'        'Dim co As New ADODB.Command
'        CR03.ReportFileName = App.Path & "\Reportes\comercial\ar_listar_id_cliente_com.rpt"
'        CR03.WindowShowPrintSetupBtn = True
'        CR03.WindowShowRefreshBtn = True
'        'MsgBox rs.RecordCount
'          'CR03.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
'          'CR03.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
'        'Call CREAVISTAF11          'JQA JUN-2008
''        CR03.StoredProcParam() = Me.Ado_datos.Recordset!ges_gestion
'        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
''        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
'        iResult = CR03.PrintReport
'        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
'        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
End If
    
End Sub

Private Sub BtnImprimir2_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR02.ReportFileName = App.Path & "\Reportes\Comercial\ar_bitacora_negociaciones.rpt"
        CR02.WindowShowPrintSetupBtn = True
        CR02.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
        CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
        CR02.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR02.PrintReport
        If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
        CR02.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos de la Bitacora ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir1_Click()
    If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\comercial\ar_R220.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
          'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
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

Private Sub BtnModificar1_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If

On Error GoTo EditErr
  'marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    VAR_SOLA = Ado_datos.Recordset!solicitud_codigo_ant
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    GlUnidad = Ado_datos.Recordset!unidad_codigo
    GlSolicitud = Ado_datos.Recordset!solicitud_codigo
    GlEdificio = Ado_datos.Recordset!edif_codigo
    glGestion = Ado_datos.Recordset!ges_gestion
'    Select Case dtc_codigo2.Text
'        Case "1"
'        Case "2"
'        Case "3", "9"
'            Call ABRIR_TABLA_DET3
'            mw_solicitud_edificacion.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'            mw_solicitud_edificacion.Txt_campo1.Text = Me.Ado_detalle1.Recordset("unidad_codigo")   'Unidad
'            mw_solicitud_edificacion.Txt_descripcion.Text = Me.dtc_desc1.Text
'            'mw_solicitud_edificacion.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
'            'mw_solicitud_edificacion.Txt_estado.Caption = "REG"
'            GlEdificio = Me.Ado_detalle1.Recordset("edif_codigo")
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

            mw_solicitud_edificacion.Show vbModal
'        Case "4"
'
'    End Select
    Call ABRIR_TABLA_DET3
    If OptFilGral1.Value = True Then
       Call OptFilGral1_Click        'Pendientes
    Else
       Call OptFilGral2_Click        'TODOS
    End If
    If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
    End If
    If Ado_datos.Recordset.RecordCount > 0 And swnuevo = 2 Then
       rs_datos.Find "solicitud_codigo_ant = " & VAR_SOLA & "   ", , , 1
       dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
    Else
       rs_datos.MoveLast
    End If
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If

  Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnModificar2_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If

On Error GoTo QError
  'marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And Ado_detalle2.Recordset!estado_codigo = "REG" And Ado_detalle2.Recordset.RecordCount > 0 Then
    VAR_SOLA = Ado_datos.Recordset!solicitud_codigo_ant
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    
    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
    Aux = Ado_datos.Recordset!unidad_codigo               'Unidad
    aw_p_ao_negociacion_bitacora.txt_campo1.Caption = Aux  'Unidad
    aw_p_ao_negociacion_bitacora.txt_codigo.Caption = VAR_SOL  'Tramite
      
    'aw_p_ao_negociacion_bitacora.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'solicitud_codigo
    'aw_p_ao_negociacion_bitacora.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
    aw_p_ao_negociacion_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
    aw_p_ao_negociacion_bitacora.Txt_Correl.Caption = Me.Ado_detalle2.Recordset("bitacora_codigo")
    'aw_p_ao_negociacion_bitacora.Txt_estado.Caption = "REG"
    'Ado_detalle1.Recordset.AddNew
     
    aw_p_ao_negociacion_bitacora.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("negocia_forma")
    aw_p_ao_negociacion_bitacora.DTPfecha1.Value = IIf(IsNull(Me.Ado_detalle2.Recordset!negocia_fecha_real), Date, Me.Ado_detalle2.Recordset!negocia_fecha_real)        'Fecha
    aw_p_ao_negociacion_bitacora.Txt_campo2.Value = IIf(IsNull(Me.Ado_detalle2.Recordset!negocia_hora_real) Or (Me.Ado_detalle2.Recordset!negocia_hora_real = "0"), Str(Time), Me.Ado_detalle2.Recordset!negocia_hora_real) 'Hora
    aw_p_ao_negociacion_bitacora.Txt_monto1.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!negocia_gasto_estimado), 0, Me.Ado_detalle2.Recordset!negocia_gasto_estimado)
    aw_p_ao_negociacion_bitacora.dtc_codigo2.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!beneficiario_codigo), "0", Me.Ado_detalle2.Recordset!beneficiario_codigo)
    aw_p_ao_negociacion_bitacora.dtc_codigo3.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!beneficiario_codigo_cgi), "0", Me.Ado_detalle2.Recordset!beneficiario_codigo_cgi)
    aw_p_ao_negociacion_bitacora.Txt_campo3.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!negocia_tarea_realizada), "NINGUNA", Me.Ado_detalle2.Recordset!negocia_tarea_realizada)
    aw_p_ao_negociacion_bitacora.Txt_campo4.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!negocia_observaciones), "", Me.Ado_detalle2.Recordset!negocia_observaciones)
    aw_p_ao_negociacion_bitacora.Txt_campo5.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bitacora_cite), "-", Me.Ado_detalle2.Recordset!bitacora_cite)
    If swnuevo = 2 Then
        aw_p_ao_negociacion_bitacora.dtc_desc1.BoundText = aw_p_ao_negociacion_bitacora.dtc_codigo1.BoundText
        aw_p_ao_negociacion_bitacora.dtc_desc2.BoundText = aw_p_ao_negociacion_bitacora.dtc_codigo2.BoundText
        aw_p_ao_negociacion_bitacora.dtc_desc3.BoundText = aw_p_ao_negociacion_bitacora.dtc_codigo3.BoundText
    End If
    
    aw_p_ao_negociacion_bitacora.Show vbModal
    
    Call ABRIR_TABLA_DET3
    If OptFilGral1.Value = True Then
       Call OptFilGral1_Click        'Pendientes
    Else
       Call OptFilGral2_Click        'TODOS
    End If
    If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
    End If
    If Ado_datos.Recordset.RecordCount > 0 And swnuevo = 2 Then
       rs_datos.Find "solicitud_codigo_ant = " & VAR_SOLA & "   ", , , 1
       dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
    Else
       rs_datos.MoveLast
    End If
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar un registro Aprobado o verifique si fue correctamente identificado !! ", vbExclamation
  End If
  
Exit Sub
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub BtnModificar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
        BtnAux3.Visible = True
        BtnAux2.Visible = True
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        dtc_desc3.SetFocus
    '    BtnVer.Visible = True
        dtc_codigo9.Enabled = False
        fra_cliente.Caption = "CLIENTE (Registra una de las 3 alternativas)"
        glBenef = dtc_codigo4.Text
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
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
      'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
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
      'solicitud_codigo, unidad_codigo, negocia_fecha_inicio as fecha1, negocia_descripcion, estado_codigo, fecha_registro, usr_codigo, solicitud_tipo as codigo2, edif_codigo as codigo3, beneficiario_codigo as codigo4, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, hora_registro, ges_gestion, archivo_respaldo, archivo_respaldo_cargado
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
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
Exit Sub
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

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
    dtc_desc11.Enabled = True
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
  
Private Sub pnivel11(codigo1 As String)
   Dim strConsultaF As String
   'strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
   strConsultaF = "Select * from rv_unidad_vs_responsable where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"
   
   Set dtc_codigo11.RowSource = Nothing
   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo11.ReFill
   dtc_codigo11.BoundText = Empty
   
   Set dtc_desc11.RowSource = Nothing
   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc11.ReFill
   dtc_desc11.BoundText = Empty
End Sub

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
'    Call pnivel11(parametro)
    dtc_desc11.Enabled = True
End Sub
 
Private Sub dtc_desc3_LostFocus()
    dtc_codigo4.Text = dtc_aux3.Text
    If Txt_descripcion.Text = "" Then
        'Txt_descripcion.Text = "IDENTIFICACION DEL CLIENTE - " + Trim(dtc_desc3.Text)
        Txt_descripcion.Text = lbl_titulo + " para el Edificio: " + Trim(dtc_desc3.Text)
    Else
        Txt_descripcion.Text = lbl_titulo + " para el Edificio: " + Trim(dtc_desc3.Text)
    '    txt_obs.Text = Txt_descripcion.Text
    End If
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
End Sub
   
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

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
'    Call pnivel6(dtc_codigo6.BoundText)
'    dtc_desc7.Enabled = True
End Sub
  
Private Sub pnivel6(codigo6 As String)
   Dim strConsultaF As String
   strConsultaF = "select * from gc_proceso_nivel3 where subproceso_codigo = '" & codigo6 & "'"
   
   Set dtc_codigo7.RowSource = Nothing
   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo7.RowSource = db.Execute("EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
   dtc_codigo7.ReFill
   dtc_codigo7.BoundText = Empty
   
   Set dtc_desc7.RowSource = Nothing
   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo7.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
   dtc_desc7.ReFill
   dtc_desc7.BoundText = Empty
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    Call pnivel8(dtc_codigo8.BoundText)
    'dtc_desc9.Enabled = True
    dtc_codigo9.Enabled = True
End Sub
   
Private Sub pnivel8(codigo8 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_documentos_respaldo where clasif_codigo = '" & codigo8 & "'"
   
   Set dtc_codigo9.RowSource = Nothing
   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo9.ReFill
   dtc_codigo9.BoundText = Empty
   
   Set dtc_desc9.RowSource = Nothing
   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc9.ReFill
   dtc_desc9.BoundText = Empty
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
    dtc_codigo9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    txt_ci.Text = "0"
    BtnAux3.Visible = False
    BtnAux2.Visible = False
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    VAR_UORIGEN = Aux
    Select Case VAR_DA
        Case "1.8"    'Cochabamba - Comercial
            Aux = "DCOMB"
            VAR_DPTO = "3"
        Case "1.7"    'Santa Cruz - Comercial
            Aux = "DCOMS"
            VAR_DPTO = "7"
        Case "1.2"    'La Paz - Comercial
            Aux = "DVTA"
            VAR_DPTO = "2"
        Case "1.9"    ' Chuquisaca - Comercial
            Aux = "DCOMC"
            VAR_DPTO = "1"
        Case "1.3"    ' La Paz - Modernizacion
            Aux = "DNMOD"
            VAR_DPTO = "2"
        Case "0"    ' TODO
            'If glusuario = "ASANTIVAÑEZ" Then
            '    Aux = "DNMOD"
            '    VAR_DPTO = "3"
            'Else
                Aux = "DVTA"
                VAR_DPTO = "0"
            'End If
     End Select
    parametro = Aux
    Call ABRIR_TABLA_AUX2
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    If Ado_datos.Recordset.RecordCount > 0 Then
    GlSolicitud = Ado_datos.Recordset!solicitud_codigo
        glGestion = Ado_datos.Recordset!ges_gestion
        Call ABRIR_TABLA_DET3
    End If
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "Select * from gc_tipo_solicitud order by solicitud_tipo", db, adOpenStatic
    rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    Select Case VAR_DPTO
        Case "3"    'Cochabamba
            rs_datos3.Open "Select * from gc_edificaciones where (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '4') and estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
        Case "7"    'Santa Cruz
            rs_datos3.Open "Select * from gc_edificaciones where estado_codigo = 'APR' and (depto_codigo = '" & VAR_DPTO & "' or depto_codigo = '8' or depto_codigo = '9' ) order by edif_descripcion", db, adOpenStatic
        Case "2"    'La Paz
            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "BINFANTE" Or glusuario = "AURBINA" Or glusuario = "GSOLIZ" Or glusuario = "CSALINAS" Then
                rs_datos3.Open "Select * from gc_edificaciones where estado_codigo = 'APR'  order by edif_descripcion", db, adOpenStatic
            Else
                rs_datos3.Open "Select * from gc_edificaciones where estado_codigo = 'APR' and (depto_codigo = '" & VAR_DPTO & "' or depto_codigo = '4' ) order by edif_descripcion", db, adOpenStatic
            End If
        Case "1"    ' Chuquisaca
            rs_datos3.Open "Select * from gc_edificaciones where estado_codigo = 'APR' and (depto_codigo = '" & VAR_DPTO & "' or depto_codigo = '5' or depto_codigo = '6' ) order by edif_descripcion", db, adOpenStatic
        Case Else   ' TODO
            rs_datos3.Open "Select * from gc_edificaciones where estado_codigo = 'APR'  order by edif_descripcion", db, adOpenStatic
     End Select
'    rs_datos3.Open "Select * from gc_edificaciones where estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "Select * from gc_beneficiario where (tipoben_codigo < 20 and tipoben_codigo <> 1) order by beneficiario_denominacion", db, adOpenStatic
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    'rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
    rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
          
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
End Sub

Private Sub ABRIR_TABLA_DET3()
    
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from ao_solicitud_edificacion where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    
    Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    'rs_det2.Open "SELECT ges_gestion, bitacora_codigo, estado_codigo, fecha_registro, hora_registro, usr_codigo, unidad_codigo, solicitud_codigo as codigo2, negocia_forma  as codigo3, beneficiario_codigo  as codigo4, beneficiario_codigo_cgi  as codigo5, negocia_tarea_realizada as descripcion, negocia_observaciones as campo1, negocia_fecha_prevista As fecha1, negocia_fecha_real As fecha2, negocia_hora_prevista As campo2, negocia_hora_real As campo3, negocia_gasto_estimado As monto1, bitacora_cite, beneficiario_nombre_ref From ao_negociacion_bitacora WHERE unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det2.Open "SELECT * From ao_negociacion_bitacora WHERE unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle2.Recordset = rs_det2
    Set dg_det2.DataSource = Ado_detalle2.Recordset
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockOptimistic
    'rs_datos11.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
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
  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
      If Ado_datos.Recordset.RecordCount > 0 Then
        GlSolicitud = Ado_datos.Recordset!solicitud_codigo
        glGestion = Ado_datos.Recordset!ges_gestion
        If VAR_SW = "MOD" Then
            parametro = Ado_datos.Recordset!unidad_codigo
        End If
        'Esto mostrará la posición de registro actual para este Recordset
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
        If VAR_SW <> "ADD" Then
            parametro = Ado_datos.Recordset!unidad_codigo
            'Select Case dtc_codigo2.Text
            '    Case "1"
            '    Case "2"
            '    Case "3", "9"
                    
                    Call ABRIR_TABLA_DET3
'                    If Ado_detalle1.Recordset.RecordCount > 0 Then
'                        BtnAñadir1.Visible = False
'                    Else
'                        BtnAñadir1.Visible = True
'                    End If
             '   Case "4"
               
            'End Select
            
            txt_nombre.Visible = False
            Call ABRIR_TABLA_AUX2
            fra_cliente.Caption = "CLIENTE"
        Else
            If VAR_SW <> "MOD" Then
                fra_cliente.Caption = "CLIENTE"
            Else
            
            End If
            'If Ado_detalle1.Recordset.RecordCount > 0 Then
            '    BtnAñadir1.Visible = False
            'Else
            '    BtnAñadir1.Visible = True
            'End If
            'Set rs_det1 = New ADODB.Recordset
            Set dg_det1.DataSource = rsNada
            Set dg_det2.DataSource = rsNada
            'Set DtgLaborales.DataSource = rsNada
        End If
    '    txt_aux9.Text = dtc_desc9.Text
        If dtc_codigo2.Text <> "" Then
            GlSolicitud = Ado_datos.Recordset!solicitud_codigo
            If Ado_datos.Recordset!estado_codigo = "APR" Then
                FrmABMDet.Visible = False
                FrmABMDet2.Visible = True
            Else
                FrmABMDet.Visible = True
                FrmABMDet2.Visible = True
            End If
        End If
      End If
    End If
End Sub

'Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'Aquí se coloca el código de validación
'  'Se llama a este evento cuando ocurre la siguiente acción
'  Dim bCancel As Boolean
'
'  Select Case adReason
'  Case adRsnAddNew
'  Case adRsnClose
'  Case adRsnDelete
'  Case adRsnFirstChange
'  Case adRsnMove
'  Case adRsnRequery
'  Case adRsnResynch
'  Case adRsnUndoAddNew
'  Case adRsnUndoDelete
'  Case adRsnUndoUpdate
'  Case adRsnUpdate
'  End Select
'
'  If bCancel Then adStatus = adStatusCancel
'End Sub

Private Sub BtnAñadir_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo AddErr
    VAR_SW = "ADD"
    BtnAux2.Visible = True
    BtnAux3.Visible = True
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    'txt_codigo.Enabled = False
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    dtc_desc3.SetFocus
    'dtc_desc1.BackColor = &H80000005
    dtc_codigo1.Text = parametro    '"DVTA"
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
    If parametro = "DNMOD" Then
        dtc_codigo2.Text = 9
        dtc_desc2.BoundText = dtc_codigo2.BoundText
        dtc_codigo5.Text = "TEC"
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_codigo6.Text = "TEC-05"
        dtc_desc6.BoundText = dtc_codigo6.BoundText
        dtc_codigo7.Text = "TEC-05-01"
        dtc_desc7.BoundText = dtc_codigo7.BoundText
        dtc_codigo8.Text = "TEC"
        dtc_desc8.BoundText = dtc_codigo8.BoundText
        dtc_codigo9.Text = "R-313"
        dtc_desc9.BoundText = dtc_codigo9.BoundText
        dtc_codigo10.Text = "3.2.7"
        dtc_desc10.BoundText = dtc_codigo10.BoundText
    Else
        dtc_codigo2.Text = 3
        dtc_desc2.BoundText = dtc_codigo2.BoundText
        dtc_codigo5.Text = "COM"
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_codigo6.Text = "COM-01"
        dtc_desc6.BoundText = dtc_codigo6.BoundText
        dtc_codigo7.Text = "COM-01-01"
        dtc_desc7.BoundText = dtc_codigo7.BoundText
        dtc_codigo8.Text = "COM"
        dtc_desc8.BoundText = dtc_codigo8.BoundText
        dtc_codigo9.Text = "R-220"      '"R-233"
        dtc_desc9.BoundText = dtc_codigo9.BoundText
        dtc_codigo10.Text = "3.1.1"
        dtc_desc10.BoundText = dtc_codigo10.BoundText
        'DTPfecha1.Value = Date       'Format(Date, "dd/mm/aaaa")
    End If
'    BtnVer.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
    FraDet1.Visible = False
    FraDet2.Visible = False
    BtnImprimir1.Visible = False
    BtnImprimir2.Visible = False
    fra_cliente.Caption = "CLIENTE (Registra una de las 3 alternativas)"
    dtc_codigo9.Enabled = False
    glBenef = "0"
    txt_ci.Text = "0"
    
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
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_ventas_cabecera WHERE unidad_codigo = '" & Unidad & "' and solicitud_codigo=" & Codigo & " and estado_codigo = 'APR'   "
'    <> 'ANL'
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

'Private Function ExisteReg(Unidad As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
'    rs.Open GlSqlAux, db, adOpenStatic
'    ExisteReg = rs!Cuantos > 0
'End Function

Private Sub ob_opcion2_Click()
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR03.ReportFileName = App.Path & "\Reportes\comercial\ar_listar_id_cliente_vendedor.rpt"
    CR03.WindowShowPrintSetupBtn = True
    CR03.WindowShowRefreshBtn = True
    CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
    iResult = CR03.PrintReport
    If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
    CR03.WindowState = crptMaximized
    fra_reportes.Visible = False
End Sub

Private Sub ob_opcion3_Click()
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR03.ReportFileName = App.Path & "\Reportes\comercial\ar_listar_id_cliente_zonas.rpt"
    CR03.WindowShowPrintSetupBtn = True
    CR03.WindowShowRefreshBtn = True
    CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
    iResult = CR03.PrintReport
    If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
    CR03.WindowState = crptMaximized
    fra_reportes.Visible = False
End Sub

Private Sub ob_opcion4_Click()
    fra_reportes.Visible = False
End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
     Set rs_datos = New Recordset
     If rs_datos.State = 1 Then rs_datos.Close
     Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
'            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Then           'CBB
'                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND ((unidad_codigo = 'DVTA') OR unidad_codigo = 'DCOMB' )) "
'                'queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
'            Else
'                'queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'  AND beneficiario_codigo_resp2 = '" & usuario2 & "' ) "
'                queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
'            End If
        Case "1.7"    'Santa Cruz
            If glusuario = "CURDININEA" Then        'SCZ
                queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9' or left(edif_codigo,1) = '3' ) )) "
            Else
                queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' ) )) "
            End If
'            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "CPAREDES" Or glusuario = "AACOSTA" Or glusuario = "GSOLIZ" Or glusuario = "CURDININEA" Then        'SCZ
'                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' )) "
'            Else
'                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'  AND beneficiario_codigo_resp2 = '" & usuario2 & "') "
'            End If
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then            'LPZ
                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMC')) "
            Else
                'queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'  AND beneficiario_codigo_resp2 = '" & usuario2 & "') "
                queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "CCOLODRO" Or glusuario = "CSALINAS" Then
                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD') )"
            Else
                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "') OR (unidad_codigo = '" & VAR_UORIGEN & "'  AND left(edif_codigo,1) = '" & VAR_DPTO & "' )"      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "'  AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5' or left(edif_codigo,1) = '6' ) )) "
'            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "EVILLALOBOS" Or glusuario = "GSOLIZ" Then            'CHQ
'                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMC')) "
'            Else
'                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'  AND beneficiario_codigo_resp2 = '" & usuario2 & "' )"
'            End If
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            Else
                queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From ao_solicitud WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
     End Select
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "unidad_codigo, solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    'wwwwwwwwwwwwwwwwwwwwwwwwww
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
        Case "1.7"    'Santa Cruz
            If glusuario = "CURDININEA" Then        'SCZ
                queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9' or left(edif_codigo,1) = '3' )))  "
            Else
                queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' )))  "
            End If
            
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then            'LPZ
                queryinicial = "select * From ao_solicitud WHERE (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMC') "
            Else
                queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9' )))  "
                'queryinicial = "select * From ao_solicitud WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "CCOLODRO" Or glusuario = "CSALINAS" Then
                queryinicial = "select * From ao_solicitud WHERE (unidad_codigo = 'DNMOD') "
            Else
                queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' ))) "      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5' or left(edif_codigo,1) = '6' )))  "
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
     End Select
'    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "unidad_codigo, solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    'wwwwwwwwwwwwwwwwwwwwwwwwww
End Sub

Private Sub ob_opcion1_Click()
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR03.ReportFileName = App.Path & "\Reportes\comercial\ar_listar_id_cliente_com.rpt"
    CR03.WindowShowPrintSetupBtn = True
    CR03.WindowShowRefreshBtn = True
    CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
    iResult = CR03.PrintReport
    If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
    CR03.WindowState = crptMaximized
    fra_reportes.Visible = False
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

'Private Sub OptFilGral1_Click()
'    parametro = "estado_codigo" + " = " + "'REG'"
'    Call ABRIR_TABLA
'End Sub
'
'Private Sub OptFilGral2_Click()
'    parametro = "estado_codigo" + " <> " + "'0'"
'    Call ABRIR_TABLA
'End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

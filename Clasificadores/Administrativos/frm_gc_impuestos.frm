VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_gc_impuestos 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Modalidad de Contratación"
   ClientHeight    =   7335
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14835
   Icon            =   "frm_gc_impuestos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   29
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "frm_gc_impuestos.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   38
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
         Picture         =   "frm_gc_impuestos.frx":11C4
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   37
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
         Picture         =   "frm_gc_impuestos.frx":1A91
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   36
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
         Picture         =   "frm_gc_impuestos.frx":2246
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   35
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
         Picture         =   "frm_gc_impuestos.frx":2A79
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   34
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1305
         Picture         =   "frm_gc_impuestos.frx":31C5
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   33
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
         Picture         =   "frm_gc_impuestos.frx":3ADA
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   32
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "frm_gc_impuestos.frx":4299
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "frm_gc_impuestos.frx":46DB
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
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
         TabIndex        =   39
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
      TabIndex        =   25
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
         Picture         =   "frm_gc_impuestos.frx":48E5
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   27
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
         Picture         =   "frm_gc_impuestos.frx":50BB
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   26
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
         TabIndex        =   28
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
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
      Height          =   6000
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   6255
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "frm_gc_impuestos.frx":59A7
         Height          =   5175
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   9128
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "imp_codigo"
            Caption         =   "Código"
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
            DataField       =   "imp_descripcion"
            Caption         =   "Denominación"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "cuantia_inicial"
            Caption         =   "Cuantía Inicial"
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
            DataField       =   "cuantia_limite"
            Caption         =   "Cuantía Límite"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3974.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5520
         Width           =   5985
         _ExtentX        =   10557
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
         Caption         =   " <-- Inicio                        Gerencia General                          Fin -->"
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
      BackColor       =   &H00000000&
      Height          =   6000
      Left            =   6600
      TabIndex        =   17
      Top             =   1200
      Width           =   8200
      Begin VB.ComboBox txt_adm 
         DataField       =   "ADM"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_gc_impuestos.frx":59BF
         Left            =   6240
         List            =   "frm_gc_impuestos.frx":59C9
         TabIndex        =   8
         Text            =   "N"
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox txt_tecnico 
         DataField       =   "TEC"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_gc_impuestos.frx":59D3
         Left            =   4920
         List            =   "frm_gc_impuestos.frx":59DD
         TabIndex        =   7
         Text            =   "N"
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox txt_otros 
         DataField       =   "OTR"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_gc_impuestos.frx":59E7
         Left            =   360
         List            =   "frm_gc_impuestos.frx":59F1
         TabIndex        =   9
         Text            =   "N"
         Top             =   3360
         Width           =   615
      End
      Begin VB.ComboBox txt_contabilidad 
         DataField       =   "CTB"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_gc_impuestos.frx":59FB
         Left            =   3480
         List            =   "frm_gc_impuestos.frx":5A05
         TabIndex        =   6
         Text            =   "N"
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox txt_comex 
         DataField       =   "CMX"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_gc_impuestos.frx":5A0F
         Left            =   1920
         List            =   "frm_gc_impuestos.frx":5A19
         TabIndex        =   5
         Text            =   "N"
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox txt_comercial 
         DataField       =   "COM"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_gc_impuestos.frx":5A23
         Left            =   360
         List            =   "frm_gc_impuestos.frx":5A2D
         TabIndex        =   4
         Text            =   "N"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txt_porcentaje 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "imp_porcentaje"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   1635
         Width           =   1815
      End
      Begin VB.TextBox txt_porcentaje2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "imp_porcentaje2"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   1635
         Width           =   1815
      End
      Begin VB.TextBox Txt_descripcion 
         DataField       =   "imp_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Text            =   "-"
         Top             =   800
         Width           =   4935
      End
      Begin VB.TextBox txt_codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "imp_codigo"
         DataSource      =   "Ado_datos"
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
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Top             =   800
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "comex"
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
         Left            =   1920
         TabIndex        =   45
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "contabilidad"
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
         Left            =   3360
         TabIndex        =   44
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tecnico"
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
         Left            =   4920
         TabIndex        =   43
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "otros"
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
         Left            =   360
         TabIndex        =   42
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Aministracion"
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
         Left            =   6120
         TabIndex        =   41
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lbl_parametro_mayor 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Porcentaje 2"
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
         Left            =   2880
         TabIndex        =   40
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label lbl_observacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Comercial"
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
         Left            =   360
         TabIndex        =   24
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label lbl_parametro_menor 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Porcentaje"
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
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Denominación"
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
         Left            =   2880
         TabIndex        =   22
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código "
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
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   10
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado Registro"
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
         Index           =   2
         Left            =   6240
         TabIndex        =   18
         Top             =   3060
         Width           =   1455
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
      TabIndex        =   11
      Top             =   7335
      Width           =   14835
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   16
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2400
      Top             =   6480
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
      Top             =   6480
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
End
Attribute VB_Name = "frm_gc_impuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "APR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
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
        Call ABRIR_TABLA
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        txt_codigo.Enabled = True
'        dtc_desc1.Enabled = True
    End If
End Sub

Private Sub BtnEliminar_Click()
 On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!imp_codigo) Then
      sino = MsgBox("No se puede ELIMINAR porque el Registro ya fue utilizado. Desea marcar como ERRADO ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ERR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
      sino = MsgBox("Está Seguro de ELIMINAR fisicamente el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         db.Execute "DELETE gc_impuestos where imp_codigo = '" & Ado_datos.Recordset!imp_codigo & "'  "
      End If
   End If
   Call ABRIR_TABLA
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
'        Set rs_aux1 = New ADODB.Recordset
'        'Busca en la tabla actual el codigo del padre
'        SQL_FOR = "select * from gc_documentos_respaldo where clasif_codigo = '" & dtc_codigo1.Text & "'  "
'        'Set rs_aux1.DataSource = db.Execute(" EXEC gp_listar_mediante_codigo_gc_direccion_general '" & txt_codigo.Text & "' ")
'        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
''            MsgBox " CODIGO DUPLICADO, Vuelva a intentar..."
''            Exit Sub
'            var_cod = rs_aux1.RecordCount + 1
'        Else
'            var_cod = 1
'        End If
''        rs_datos!doc_codigo = RTrim(RTrim(dtc_codigo1.Text) + ".") + LTrim(Str(Val(var_cod)))
        
        rs_datos!imp_codigo = txt_codigo.Text ' Esto para codigos trascritos
        rs_datos!estado_codigo = "REG"  ' no cambia
'        rs_datos!correl_etapa = 0
'        rs_datos!proceso_codigo = dtc_codigo1.Text   'Codigo del padre
        'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
'        db.Execute "Update gc_direccion_general Set correl_da = CAST('" & var_cod & "' AS INT) + 1 Where dgral_codigo= '" & dtc_codigo1.Text & "' "
     End If
     rs_datos!imp_descripcion = Txt_descripcion.Text
     rs_datos!imp_porcentaje = txt_porcentaje.Text
     rs_datos!imp_porcentaje2 = txt_porcentaje2.Text
     rs_datos!COM = txt_comercial.Text
     rs_datos!CMX = txt_comex.Text
     rs_datos!CTB = txt_contabilidad.Text
     rs_datos!TEC = txt_tecnico.Text
     rs_datos!ADM = txt_adm.Text
     rs_datos!OTR = txt_otros.Text
     rs_datos!fecha_registro = Date     ' no cambia
     rs_datos!usr_codigo = glusuario    ' no cambia
     rs_datos.UpdateBatch adAffectAll
    
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
      
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      txt_codigo.Enabled = True
'      dtc_desc1.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'habilitar codigo cuando se transcribe
  If txt_codigo.Text = "" Then
    MsgBox "Debe registrar el " + lbl_codigo.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_porcentaje.Text <= 100 Then
    MsgBox "Debe registrar: " + lbl_enlace1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  cr01.WindowShowPrintSetupBtn = True
  cr01.WindowShowRefreshBtn = True
  cr01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_impuestos.rpt"
  iResult = cr01.PrintReport
  If iResult <> 0 Then
      MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  cr01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
          Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
    txt_codigo.Enabled = False
'    dtc_desc1.Enabled = False
   Else
        MsgBox "No se puede MODIFICAR un registro APROBADO o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  
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

Private Sub DtcUE_Click(Area As Integer)
    DtcUE_Des.BoundText = DtcUE.BoundText
End Sub

Private Sub DtcUE_Des_Click(Area As Integer)
    DtcUE.BoundText = DtcUE_Des.BoundText
End Sub

'Private Sub dtc_codigo1_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'End Sub

'Private Sub dtc_desc1_Click(Area As Integer)
'    dtc_codigo1.BoundText = dtc_desc1.BoundText
'End Sub

Private Sub Form_Load()
'    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from gc_impuestos "
  'queryinicial = "select  da_codigo, da_descripcion, dgral_codigo, proceso_codigo, estado_codigo, fecha_registro, usr_codigo, correl_unidad as correl from gc_direccion_administrativa  "
  'queryinicial = "gp_listar_gc_direccion_general "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

'Private Sub ABRIR_TABLAS_AUX()
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    rs_datos1.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'End Sub

'Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
      Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
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
    Call ABRIR_TABLA
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
'    txt_codigo.Enabled = False
'    Txt_descripcion.SetFocus
    txt_codigo.SetFocus
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
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ro_datos_educacionales WHERE imp_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Private Sub txt_margen_preferencia_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

'Private Sub txt_nal_ext_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

'Private Sub txt_nro_dias_calendario_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

'Private Sub txt_nro_dias_habiles_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub txt_observacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_parametro_mayor_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
Validar txt_porcentaje.Text, KeyAscii

End Sub

Public Sub Validar(DatosActuales As String, Caracter As Integer)
  ' Salimos si se ha pulsado la tecla de Retroceso
  If Caracter = 8 Then Exit Sub
  ' Salimos si es de 0 a 9
  If InStr("0123456789", Chr$(Caracter)) Then Exit Sub
  ' Si es punto y no está en el contenido salimos
  If Caracter = 46 And InStr(DatosActuales, ".") = 0 Then Exit Sub
  ' Borramos el Caracter introducido
  Caracter = 0
End Sub

Private Sub txt_porcentaje_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txt_porcentaje.Text) < 0 Or Val(txt_porcentaje.Text) > 100 Then
MsgBox "rango invalido solo de 0 a 100"
txt_porcentaje.Text = ""
End If
End Sub

Private Sub txt_porcentaje2_KeyPress(KeyAscii As Integer)
Validar txt_porcentaje2.Text, KeyAscii
End Sub

Private Sub txt_porcentaje2_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txt_porcentaje2.Text) < 0 Or Val(txt_porcentaje2.Text) >= 100 Then
MsgBox "rango invalido solo de 0 a 100"
txt_porcentaje2.Text = ""
End If
End Sub

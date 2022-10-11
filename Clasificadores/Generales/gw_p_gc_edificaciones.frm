VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form gw_p_gc_edificaciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Oportunidades de Negocio - Proyectos de Edificación"
   ClientHeight    =   9120
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "gw_p_gc_edificaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
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
         Left            =   16560
         Picture         =   "gw_p_gc_edificaciones.frx":0A02
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
         Picture         =   "gw_p_gc_edificaciones.frx":0C0C
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
         Left            =   5520
         Picture         =   "gw_p_gc_edificaciones.frx":13CE
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   78
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         Picture         =   "gw_p_gc_edificaciones.frx":1C9B
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   77
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6840
         Picture         =   "gw_p_gc_edificaciones.frx":2450
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   76
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         Picture         =   "gw_p_gc_edificaciones.frx":2C83
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   75
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1200
         Picture         =   "gw_p_gc_edificaciones.frx":33CF
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   74
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
         Picture         =   "gw_p_gc_edificaciones.frx":3CE4
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   73
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8280
         Picture         =   "gw_p_gc_edificaciones.frx":44A3
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   72
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
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
         Left            =   13320
         TabIndex        =   81
         Top             =   195
         Width           =   885
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
         Left            =   4680
         Picture         =   "gw_p_gc_edificaciones.frx":4FA6
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
         Picture         =   "gw_p_gc_edificaciones.frx":577C
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
         Caption         =   "TITULO"
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
         Left            =   13275
         TabIndex        =   70
         Top             =   195
         Width           =   885
      End
   End
   Begin VB.PictureBox Fra_aux1 
      BackColor       =   &H00808080&
      FillColor       =   &H00FFFFFF&
      Height          =   1260
      Left            =   6600
      ScaleHeight     =   1200
      ScaleWidth      =   8295
      TabIndex        =   61
      Top             =   7560
      Width           =   8350
      Begin VB.ComboBox dtc_codigo11 
         Height          =   315
         ItemData        =   "gw_p_gc_edificaciones.frx":6068
         Left            =   120
         List            =   "gw_p_gc_edificaciones.frx":607B
         TabIndex        =   16
         Text            =   "CALLE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Txt_descripcion11 
         DataField       =   "calle_denominacion"
         Height          =   645
         Left            =   2160
         TabIndex        =   17
         Text            =   "-"
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton CmdCancelaDet 
         BackColor       =   &H80000018&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   6585
         MaskColor       =   &H00000000&
         Picture         =   "gw_p_gc_edificaciones.frx":609C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancelar"
         Top             =   315
         Width           =   765
      End
      Begin VB.CommandButton CmdGrabaDet 
         BackColor       =   &H80000018&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   7440
         Picture         =   "gw_p_gc_edificaciones.frx":62A6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   315
         Width           =   765
      End
      Begin VB.Label lbl_descripcion11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación Via de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   2400
         TabIndex        =   63
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label lbl_enlace11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Vía de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   2070
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TITULO"
      ForeColor       =   &H00C00000&
      Height          =   7695
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   6255
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "gw_p_gc_edificaciones.frx":64B0
         Height          =   6810
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   12012
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
            DataField       =   "edif_codigo"
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
            DataField       =   "edif_descripcion"
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
         BeginProperty Column05 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Propietario"
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
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3525.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   675.213
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
         Top             =   7200
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
         BackColor       =   16777152
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
         Caption         =   " <-- Inicio                   Viviendas - Edificaciones                   Fin -->"
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
      Height          =   7695
      Left            =   6520
      TabIndex        =   27
      Top             =   720
      Width           =   8535
      Begin VB.TextBox txt_campo6 
         DataField       =   "texto_borrar"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "gw_p_gc_edificaciones.frx":64C8
         Top             =   5640
         Width           =   6780
      End
      Begin VB.CommandButton BtnAux1 
         BackColor       =   &H00808000&
         Caption         =   "Nueva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   7700
         MaskColor       =   &H00FFFFFF&
         Picture         =   "gw_p_gc_edificaciones.frx":64CA
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4900
         Width           =   660
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Ubicación del Proyecto"
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
         Height          =   1815
         Left            =   100
         TabIndex        =   49
         Top             =   2340
         Width           =   8320
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "gw_p_gc_edificaciones.frx":6C7B
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "pais_descripcion"
            BoundColumn     =   "pais_codigo"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo dtc_campo2 
            Bindings        =   "gw_p_gc_edificaciones.frx":6C94
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   58
            Top             =   1080
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_sigla"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_aux2 
            Bindings        =   "gw_p_gc_edificaciones.frx":6CAD
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6120
            TabIndex        =   57
            Top             =   360
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "correl_edif"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "gw_p_gc_edificaciones.frx":6CC6
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5160
            TabIndex        =   56
            Top             =   1080
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_codigo"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "gw_p_gc_edificaciones.frx":6CDF
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   50
            Top             =   435
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_codigo"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo9 
            Bindings        =   "gw_p_gc_edificaciones.frx":6CF8
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   51
            Top             =   1035
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "prov_codigo"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc9 
            Bindings        =   "gw_p_gc_edificaciones.frx":6D11
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   1275
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "gw_p_gc_edificaciones.frx":6D2A
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   52
            Top             =   435
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "gw_p_gc_edificaciones.frx":6D43
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4080
            TabIndex        =   3
            Top             =   600
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "gw_p_gc_edificaciones.frx":6D5C
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4080
            TabIndex        =   5
            Top             =   1275
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label lbl_titulo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Municipio"
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
            Left            =   4080
            TabIndex        =   59
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Provincia"
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
            Left            =   120
            TabIndex        =   55
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Departamento"
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
            Left            =   4080
            TabIndex        =   54
            Top             =   340
            Width           =   1290
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "País"
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
            Left            =   120
            TabIndex        =   53
            Top             =   340
            Width           =   405
         End
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "gw_p_gc_edificaciones.frx":6D75
         DataField       =   "codigo6"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   45
         Top             =   6240
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
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "gw_p_gc_edificaciones.frx":6D8E
         DataField       =   "codigo5"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   44
         Top             =   5880
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
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "gw_p_gc_edificaciones.frx":6DA7
         DataField       =   "codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   43
         Top             =   4440
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
         Bindings        =   "gw_p_gc_edificaciones.frx":6DC0
         DataField       =   "codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2880
         TabIndex        =   42
         Top             =   4440
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
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "gw_p_gc_edificaciones.frx":6DD9
         DataField       =   "codigo1"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5280
         TabIndex        =   41
         Top             =   1920
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
      Begin VB.PictureBox Img_Foto 
         Height          =   2055
         Left            =   6360
         ScaleHeight     =   1995
         ScaleWidth      =   1995
         TabIndex        =   40
         Top             =   250
         Width           =   2055
         Begin VB.Image Image2 
            DataField       =   "foto"
            DataSource      =   "Ado_datos"
            Height          =   2002
            Left            =   0
            Picture         =   "gw_p_gc_edificaciones.frx":6DF2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2002
         End
      End
      Begin VB.TextBox txt_campo3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "campo3"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "-"
         Top             =   7240
         Width           =   1575
      End
      Begin VB.TextBox txt_campo4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "campo4"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4080
         TabIndex        =   14
         Text            =   "-"
         Top             =   7240
         Width           =   1575
      End
      Begin VB.TextBox txt_campo5 
         BackColor       =   &H00FFFFFF&
         DataField       =   "campo5"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   6360
         TabIndex        =   15
         Text            =   "-"
         Top             =   7240
         Width           =   1575
      End
      Begin VB.TextBox txt_campo2 
         DataField       =   "campo2"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "gw_p_gc_edificaciones.frx":9A9C
         Top             =   5160
         Width           =   5655
      End
      Begin VB.TextBox txt_campo1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "campo1"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "-"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox Txt_descripcion 
         Appearance      =   0  'Flat
         DataField       =   "edif_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Text            =   "gw_p_gc_edificaciones.frx":9A9E
         Top             =   1100
         Width           =   5895
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "gw_p_gc_edificaciones.frx":9AA0
         DataField       =   "codigo1"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1300
         TabIndex        =   1
         Top             =   1875
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "gw_p_gc_edificaciones.frx":9AB9
         DataField       =   "codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         TabIndex        =   7
         Top             =   4500
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "gw_p_gc_edificaciones.frx":9AD2
         DataField       =   "codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   4500
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "gw_p_gc_edificaciones.frx":9AEB
         DataField       =   "codigo5"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2470
         TabIndex        =   11
         Top             =   6045
         Width           =   5900
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "gw_p_gc_edificaciones.frx":9B04
         DataField       =   "codigo6"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2115
         TabIndex        =   12
         Top             =   6540
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Observaciones"
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
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   5640
         Width           =   1380
      End
      Begin VB.Label lbl_calle 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Via de Acceso (Calle, Av, etc.)"
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
         Left            =   4200
         TabIndex        =   65
         Top             =   4240
         Width           =   2685
      End
      Begin VB.Label lbl_zona 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Zona / Barrio"
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
         Left            =   120
         TabIndex        =   64
         Top             =   4240
         Width           =   1155
      End
      Begin VB.Label lbl_titulo1 
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Edificación"
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
         Height          =   480
         Left            =   240
         TabIndex        =   48
         Top             =   1695
         Width           =   1020
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Denominación del Proyecto de Edificación"
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
         Left            =   240
         TabIndex        =   47
         Top             =   830
         Width           =   3810
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "edif_codigo"
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
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   450
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Altura Nivel del Mar"
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
         Index           =   14
         Left            =   6360
         TabIndex        =   39
         Top             =   6980
         Width           =   1740
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Longitud"
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
         Index           =   13
         Left            =   4080
         TabIndex        =   38
         Top             =   6980
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Latitud"
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
         Index           =   12
         Left            =   1920
         TabIndex        =   37
         Top             =   6980
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Geo Referencia"
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
         Index           =   11
         Left            =   240
         TabIndex        =   36
         Top             =   7260
         Width           =   1425
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Propietario/ Responsable"
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
         Height          =   360
         Index           =   10
         Left            =   120
         TabIndex        =   35
         Top             =   6050
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Empresa o Institución"
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
         Index           =   9
         Left            =   120
         TabIndex        =   34
         Top             =   6555
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Ubicación Referencial Cercana"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   33
         Top             =   4920
         Width           =   2805
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. del Edificio"
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
         Index           =   7
         Left            =   120
         TabIndex        =   32
         Top             =   4920
         Width           =   1410
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código Proy.Edificación"
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
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   200
         Width           =   2160
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
         Left            =   4680
         TabIndex        =   28
         Top             =   195
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
      ScaleWidth      =   21360
      TabIndex        =   21
      Top             =   12930
      Width           =   21360
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   26
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   8160
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
      Left            =   2040
      Top             =   8040
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
      Top             =   8160
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
      Left            =   4440
      Top             =   8160
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
      Left            =   120
      Top             =   8520
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
      Left            =   2280
      Top             =   8520
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
      Left            =   4440
      Top             =   8520
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
      Left            =   6600
      Top             =   8160
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
      Left            =   8760
      Top             =   8160
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
      Left            =   11040
      Top             =   8160
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
End
Attribute VB_Name = "gw_p_gc_edificaciones"
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
Dim rs_datos11 As New ADODB.Recordset

Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_COD2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Dim RUTA1, RUTA2 As String
         'RUTA1 = "BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset("dtc_campo2")) + "\" + Trim(Ado_datos.Recordset("edif_codigo"))
         RUTA1 = "BIENES\EDIFICIOS\" + Trim(dtc_campo2) + "\" + Trim(txt_codigo)
         MsgBox "Se esta creando la carpeta: " + RUTA1
         MkDir RUTA1
        '        MkDir RUTA1 + "\CONTRATOS"
         rs_datos!estado_codigo = "APR"
         'rs_datos!fecha_registro = Date
         rs_datos!fecha_aprueba = Date
         'rs_datos!usr_codigo = glusuario
         rs_datos!usr_codigo_aprueba = glusuario
         rs_datos.UpdateBatch 'adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAux1_Click()
    'Validacion 1
    If dtc_codigo3 = "" Or dtc_codigo3 = "0" Then
        MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
        VAR_VAL = "ERR"
        Exit Sub
    End If
    Fra_ABM.Enabled = False
    Fra_aux1.Visible = True
End Sub

Private Sub BtnBuscar_Click()
    buscados = 1
    'OptFilGral2.Visible = False
    'OptFilGral1.Visible = False
    Call ABRIR_TABLA
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
    End If
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!Codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado en otro proceso ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   If rs_datos!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ERR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR)...", vbExclamation, "Validación de Registro"
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
        var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        'var_cod = RTrim(RTrim(left(dtc_codigo2.Text,1)) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        Set rstbeneaux = New ADODB.Recordset
        SQL_FOR = "select * from gc_edificaciones where edif_codigo = '" & var_cod & "'  "
        rstbeneaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rstbeneaux.RecordCount > 0 Then
            MsgBox " CODIGO DUPLICADO, Vuelva a intentar..."
            Exit Sub
        End If
        txt_codigo.Caption = var_cod
        rs_datos!edif_codigo = var_cod
        rs_datos!estado_codigo = "REG"
        rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
        rs_datos!archivo_foto_cargado = "N"
        'rs_datos!ges_gestion = Year(Date)
        'rs_datos!correl_da = 0
        db.Execute "Update gc_municipio Set correl_edif = CAST('" & dtc_aux2.Text & "' AS INT) + 1 Where munic_codigo= '" & dtc_codigo2.Text & "' "
        db.Execute "Update gc_departamento Set correl_edif = CAST('" & dtc_aux2.Text & "' AS INT) + 1 Where depto_codigo= '" & Left(var_cod, 1) & "' "
     End If
     
     rs_datos!edif_descripcion = RTrim(Txt_descripcion.Text)
     rs_datos!campo1 = txt_campo1
     rs_datos!campo2 = RTrim(txt_campo2)
     
     rs_datos!codigo1 = dtc_codigo1.Text
     rs_datos!munic_codigo = IIf(dtc_codigo2.Text = "", "NN", dtc_codigo2.Text)
     rs_datos!codigo3 = IIf(dtc_codigo3.Text = "", "0", dtc_codigo3.Text)
     rs_datos!codigo4 = IIf(dtc_codigo4.Text = "", "0", dtc_codigo4.Text)
     rs_datos!codigo5 = IIf(dtc_codigo5.Text = "", "0", dtc_codigo5.Text)
     rs_datos!codigo6 = IIf(dtc_codigo6.Text = "", "0", dtc_codigo6.Text)

     rs_datos!campo3 = IIf(txt_campo3.Text = "", "0", txt_campo3.Text)
     rs_datos!campo4 = IIf(txt_campo4.Text = "", "0", txt_campo4.Text)
     rs_datos!campo5 = IIf(txt_campo5.Text = "", "0", txt_campo5.Text)
     rs_datos!campo6 = "0"
     rs_datos!campo7 = "0"
     
     If rs_datos!ARCHIVO_Foto = ".JPG" Or rs_datos!ARCHIVO_Foto = "" Then
        rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
     End If
     
     rs_datos!fecha_registro = Date
     rs_datos!usr_codigo = glusuario
     rs_datos.UpdateBatch adAffectAll
    
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
      
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      txt_codigo.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar el " + lbl_titulo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo2.Text = "" Then
    MsgBox "Debe registrar la " + lbl_titulo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
     CR01.WindowShowPrintSetupBtn = True
     CR01.WindowShowRefreshBtn = True
     CR01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_edificaciones.rpt"
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  'If glusuario = "RVEIZAGA" Or Ado_datos.Recordset!estado_codigo = "REG" Then
  If Ado_datos.Recordset!estado_codigo = "REG" Or glusuario = "HLIZARRO" Then
'    If Ado_datos.Recordset!estado_codigo = "REG" Then
    '  lblStatus.Caption = "Modificar registro"
        Fra_ABM.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
        txt_codigo.Enabled = False
        dtc_desc8.Enabled = False
        dtc_desc9.Enabled = False
        dtc_desc2.Enabled = False
        dtc_desc3.Enabled = False
        dtc_desc4.Enabled = False
    '    BtnVer.Visible = True
    Else
        MsgBox "No se puede MODIFICAR un registro Aprobado (APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
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
    If Ado_datos.Recordset!archivo_foto_cargado = "N" Then
      'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!dtc_campo2) & "\"
      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(dtc_campo2.Text) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FEDF"
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
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!codigo1) & "\"
          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(dtc_campo2.Text) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FEDF"
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
      End If
    End If
    If SW0 = 1 Then
    '    If GlServidor = "SRVPRO" Then
    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
    '    Else
            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(dtc_campo2.Text) + "\" + Trim(Ado_datos.Recordset("edif_codigo")) + "\" + Trim(Ado_datos.Recordset("edif_codigo")) + ".JPG"
            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!codigo1) + "\" + Trim(Ado_datos.Recordset!edif_codigo) + ".JPG"
    '    End If
        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
        CodBien = Ado_datos.Recordset!edif_codigo
        If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
            MsgBox "Se cargo la Imagen Correctamente !!"
        Else
            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
        End If
    Else
        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
        Image2 = Img_Foto
    End If
  Else
    MsgBox "Debe Aprobar el registro, para crear la carpeta correspondiente..."
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub CmdCancelaDet_Click()
    Fra_ABM.Enabled = True
    Fra_aux1.Visible = False
End Sub

Private Sub CmdGrabaDet_Click()
  'Validacion
  If Txt_descripcion11.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo11.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3 = "" Or dtc_codigo3 = "0" Then
    MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  'INI Graba Calle
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select max(calle_codigo) as Codigo from gc_calles where zona_codigo = " & dtc_codigo3.Text & "    ", db, adOpenStatic
    If rs_aux2.RecordCount > 0 Then
        If IsNull(rs_aux2!Codigo) Then
            VAR_COD2 = (Val(dtc_codigo3.Text) * 100) + 1
        Else
            VAR_COD2 = Round(CDbl(rs_aux2!Codigo) + 1, 0)
        End If
    Else
        VAR_COD2 = 1
    End If
    db.Execute "insert into gc_calles(zona_codigo, calle_codigo, calle_denominacion, calle_tipo, correl, estado_codigo, fecha_registro, usr_codigo)" & _
    "values ('" & dtc_codigo3.Text & "', " & VAR_COD2 & ", '" & Txt_descripcion11.Text & "', '" & dtc_codigo11.Text & "', '0', 'APR', '" & Date & "', '" & glusuario & "') "
    
   'FIN Graba Calle
    'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
    db.Execute "Update gc_zonas Set correl = " & VAR_COD2 & " Where zona_codigo= '" & dtc_codigo3.Text & "' "
    'gc_calles
    Call pnivel3(dtc_codigo3.BoundText)
    dtc_desc4.Enabled = True
    
    dtc_codigo4.Text = VAR_COD2
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Fra_ABM.Enabled = True
    Fra_aux1.Visible = False
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_aux2.BoundText
    dtc_codigo8.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_campo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_campo2.BoundText
'    dtc_aux2.BoundText = dtc_campo2.BoundText
    dtc_codigo2.BoundText = dtc_campo2.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
'    dtc_aux2.BoundText = dtc_codigo2.BoundText
    dtc_campo2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
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
    dtc_aux2.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
'    dtc_aux2.BoundText = dtc_desc2.BoundText
    dtc_campo2.BoundText = dtc_desc2.BoundText
    Call pnivel2(dtc_codigo2.BoundText)
    dtc_desc3.Enabled = True
End Sub
   
Private Sub pnivel2(codigo2 As String)
   'Dim strConsultaF As String
     
   'strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo2 & "'"
   'strConsultaF = "select zona_codigo as codigo, zona_denominacion as descripcion, estado_codigo, fecha_registro, usr_codigo, correl as correl, pais_codigo As codigo1, depto_codigo As codigo2, prov_codigo As codigo3, munic_codigo As codigo4, comun_codigo As codigo5, hora_registro From gc_zonas where estado_codigo='APR' AND munic_codigo = '" & codigo2 & "' order by descripcion "
   'strConsultaF = "select zona_codigo as codigo, zona_denominacion as descripcion, estado_codigo, munic_codigo As codigo4 From gc_zonas where estado_codigo='APR' AND munic_codigo = '" & codigo2 & "' order by descripcion "
      
   Set dtc_codigo3.RowSource = Nothing
   'Set dtc_codigo3.RowSource = db.Execute(strConsultaF, "codigo2", adCmdText)
   Set dtc_codigo3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_codigo3.ReFill
   dtc_codigo3.BoundText = Empty
   
   Set dtc_desc3.RowSource = Nothing
   'Set dtc_desc3.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_desc3.ReFill
   dtc_desc3.BoundText = Empty

End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    Call pnivel3(dtc_codigo3.BoundText)
    dtc_desc4.Enabled = True
End Sub
   
Private Sub pnivel3(codigo3 As String)
   'Dim strConsultaF As String
   
   'strConsultaF = "select * from gc_calles where zona_codigo = '" & codigo3 & "'"
   'strConsultaF = "select calle_codigo as codigo, calle_denominacion as descripcion, estado_codigo, fecha_registro, usr_codigo, correl as correl, zona_codigo As codigo1, calle_tipo As codigo2, hora_registro From gc_calles where estado_codigo='APR' AND zona_codigo = '" & codigo3 & "' order by descripcion "
   'strConsultaF = "select calle_codigo as codigo, calle_denominacion as descripcion, estado_codigo, zona_codigo As codigo1 From gc_calles where estado_codigo='APR' AND zona_codigo = '" & codigo3 & "' order by descripcion "

   Set dtc_codigo4.RowSource = Nothing
   'Set dtc_codigo4.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo4.ReFill
   dtc_codigo4.BoundText = Empty
   
   Set dtc_desc4.RowSource = Nothing
   'Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc4.ReFill
   dtc_desc4.BoundText = Empty

End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
    Call pnivel7(dtc_codigo7.BoundText)
    dtc_desc8.Enabled = True
End Sub
   
Private Sub pnivel7(codigo7 As String)
   Dim strConsultaF As String
     
   strConsultaF = "select * from gc_departamento where pais_codigo = '" & codigo7 & "'"
   Set dtc_codigo8.RowSource = Nothing
   Set dtc_codigo8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_codigo8.ReFill
   dtc_codigo8.BoundText = Empty
   
   Set dtc_desc8.RowSource = Nothing
   Set dtc_desc8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_desc8.ReFill
   dtc_desc8.BoundText = Empty

End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    dtc_aux2.BoundText = dtc_desc8.BoundText
    Call pnivel8(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
End Sub

Private Sub pnivel8(codigo8 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo8 & "'"
   Set dtc_codigo9.RowSource = Nothing
   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo9.ReFill
   dtc_codigo9.BoundText = Empty
   
   Set dtc_desc9.RowSource = Nothing
   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc9.ReFill
   dtc_desc9.BoundText = Empty
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
    dtc_codigo9.BoundText = dtc_desc9.BoundText
    Call pnivel9(dtc_codigo9.BoundText)
    dtc_desc2.Enabled = True
End Sub
  
Private Sub pnivel9(codigo9 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo9 & "'"
   Set dtc_codigo2.RowSource = Nothing
   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo2.ReFill
   dtc_codigo2.BoundText = Empty
   
   Set dtc_desc2.RowSource = Nothing
   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc2.ReFill
   dtc_desc2.BoundText = Empty
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
    Fra_aux1.Visible = False
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_edificacion_tipo order by edif_tipo_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_gc_edificacion_tipo", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_municipio order by munic_descripcion", db, adOpenStatic
    'rs_datos2.Open "gp_listar_gc_municipio", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_zonas order by zona_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_gc_zonas", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "Select * from gc_calles order by calle_denominacion", db, adOpenStatic
    rs_datos4.Open "gp_listar_gc_calles", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    'rs_datos5.Open "Select * from gc_beneficiario where (tipoben_codigo < 20 and tipoben_codigo <> 1) order by beneficiario_denominacion", db, adOpenStatic
    rs_datos5.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    'rs_datos6.Open "Select * from gc_beneficiario where (tipoben_codigo > 19) order by beneficiario_denominacion", db, adOpenStatic
    rs_datos6.Open "gp_listar_gc_beneficiario_empresas", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    'gc_pais
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from gc_pais where estado_codigo = 'APR' and pais_continente = 'AMERICA' ", db, adOpenKeyset
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'gc_Departamento  '<>
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_departamento order by depto_descripcion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    'gc_provincia
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_provincia ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  'queryinicial = "select  edif_codigo, edif_descripcion, estado_codigo, fecha_registro, usr_codigo, correl, edif_nro as campo1, edif_referencia as campo2, edif_tipo as codigo1, pais_codigo, depto_codigo, prov_codigo, munic_codigo, zona_codigo as codigo3, calle_codigo as codigo4, beneficiario_codigo as codigo5, beneficiario_codigo_empresa as codigo6, latitud As campo3, longitud As campo4, altitud_snm As campo5, foto, archivo_foto, archivo_foto_cargado, hora_registro, loc_eje_X As campo6, loc_eje_Y As campo7, texto_borrar From gc_edificaciones "
  queryinicial = "select  *, edif_nro as campo1, edif_referencia as campo2, edif_tipo as codigo1, zona_codigo as codigo3, calle_codigo as codigo4, beneficiario_codigo as codigo5, beneficiario_codigo_empresa as codigo6, latitud As campo3, longitud As campo4, altitud_snm As campo5, loc_eje_X As campo6, loc_eje_Y As campo7 From gc_edificaciones "
  'queryinicial = "select * From gc_edificaciones"
  'queryinicial = "gp_listar_gc_edificaciones"
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
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
    Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    If Ado_datos.Recordset.AbsolutePosition >= 0 Then
        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
        Image2 = Img_Foto
    End If
'    If Ado_datos.Recordset!archivo_foto_cargado = "S" Then
'        'BtnVer.Visible = True
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    Else
'        'BtnVer.Visible = False
'        'chkEstado.Value = vbUnchecked
'    End If
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
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
    txt_codigo.Enabled = False
    dtc_desc8.Enabled = False
    dtc_desc9.Enabled = False
    dtc_desc2.Enabled = False
    dtc_desc3.Enabled = False
    dtc_desc4.Enabled = False
    
    Txt_descripcion.SetFocus
'    BtnVer.Visible = False
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

Private Function ExisteReg(codigoe As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_negociacion_cabecera WHERE edif_codigo = '" & codigoe & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub txt_campo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

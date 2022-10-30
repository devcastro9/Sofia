VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form gw_p_gc_beneficiario_empresa 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clasificadores - Registro de Empresas"
   ClientHeight    =   9495
   ClientLeft      =   495
   ClientTop       =   1905
   ClientWidth     =   16980
   Icon            =   "gw_p_gc_beneficiario_empresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   16980
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   81
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5040
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   83
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3720
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":11B7
         ScaleHeight     =   735
         ScaleWidth      =   1320
         TabIndex        =   85
         ToolTipText     =   "Aprueba Entrega de Insumos"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17400
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":19EA
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   90
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
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":21AC
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   89
         ToolTipText     =   "Modifica datos de la Zona elegida"
         Top             =   0
         Width           =   1430
      End
      Begin VB.CommandButton BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Height          =   855
         Left            =   3795
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":2AC1
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   -60
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.CommandButton BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Caption         =   "Digitaliza"
         Height          =   710
         Left            =   7800
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":34B8
         Style           =   1  'Graphical
         TabIndex        =   87
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
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":38FA
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   86
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2520
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":40B9
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   84
         ToolTipText     =   "Anula Zona elegida"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   6360
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":4805
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   82
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   0
         Width           =   1400
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13305
         TabIndex        =   91
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
      TabIndex        =   77
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
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":50D2
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   79
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
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":59BE
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   78
         Top             =   0
         Width           =   1280
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13155
         TabIndex        =   80
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.PictureBox Fra_aux1 
      BackColor       =   &H00808080&
      FillColor       =   &H00FFFFFF&
      Height          =   1300
      Left            =   7320
      ScaleHeight     =   1245
      ScaleWidth      =   9195
      TabIndex        =   68
      Top             =   7110
      Width           =   9255
      Begin VB.CommandButton CmdGrabaDet 
         BackColor       =   &H80000015&
         Height          =   555
         Left            =   7680
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":6194
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   75
         Width           =   1365
      End
      Begin VB.CommandButton CmdCancelaDet 
         BackColor       =   &H80000015&
         Height          =   555
         Left            =   7680
         MaskColor       =   &H00000000&
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":696A
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Cancelar"
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox Txt_descripcion11 
         DataField       =   "calle_denominacion"
         Height          =   525
         Left            =   2160
         TabIndex        =   70
         Text            =   "-"
         Top             =   360
         Width           =   5175
      End
      Begin VB.ComboBox dtc_codigo11 
         Height          =   315
         ItemData        =   "gw_p_gc_beneficiario_empresa.frx":7256
         Left            =   120
         List            =   "gw_p_gc_beneficiario_empresa.frx":7269
         TabIndex        =   69
         Text            =   "CALLE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl_enlace11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Vía de Acceso"
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
         Height          =   240
         Left            =   120
         TabIndex        =   74
         Top             =   120
         Width           =   1785
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2160
         TabIndex        =   73
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00FF0000&
      Height          =   7815
      Left            =   120
      TabIndex        =   42
      Top             =   840
      Width           =   6975
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   7215
         Width           =   6825
         _ExtentX        =   12039
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":728A
         Height          =   6855
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Denominación Entidad"
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
            DataField       =   "beneficiario_nit"
            Caption         =   "NIT / C.I."
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Código_Beneficiario"
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
            DataField       =   "tipoben_codigo"
            Caption         =   "Tipo_Benef"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   4080.189
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraDatos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      FillColor       =   &H00C0C0C0&
      Height          =   7815
      Left            =   7200
      ScaleHeight     =   7755
      ScaleWidth      =   9525
      TabIndex        =   24
      Top             =   840
      Width           =   9585
      Begin VB.TextBox txt_campo1 
         DataField       =   "beneficiario_primer_apellido"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1440
         TabIndex        =   92
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CommandButton BtnAux1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8460
         Picture         =   "gw_p_gc_beneficiario_empresa.frx":72A2
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Registra Nueva Calle, Av, etc."
         Top             =   6240
         Width           =   900
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Localización de la Empresa"
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
         Height          =   1815
         Left            =   60
         TabIndex        =   50
         Top             =   4320
         Width           =   9270
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7D7A
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   51
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
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7D93
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   52
            Top             =   1035
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
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7DAC
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4560
            TabIndex        =   15
            Top             =   1275
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7DC5
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   53
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
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7DDE
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   1275
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7DF7
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   54
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
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7E10
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4560
            TabIndex        =   13
            Top             =   600
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7E29
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_descripcion"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   4560
            TabIndex        =   57
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4560
            TabIndex        =   56
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   55
            Top             =   1035
            Width           =   840
         End
      End
      Begin VB.TextBox txt_campo7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "beneficiario_telefono_Cel"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   6360
         TabIndex        =   9
         Text            =   "-"
         Top             =   3250
         Width           =   2775
      End
      Begin VB.TextBox txt_campo5 
         DataField       =   "beneficiario_telefono_fijo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   720
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "-"
         Top             =   3250
         Width           =   2700
      End
      Begin VB.TextBox txt_campo6 
         DataField       =   "beneficiario_telefono_Of"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Text            =   "-"
         Top             =   3250
         Width           =   2835
      End
      Begin VB.TextBox txt_campo11 
         DataField       =   "beneficiario_edif_piso_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   23
         Top             =   7215
         Width           =   915
      End
      Begin VB.TextBox txt_campo9 
         DataField       =   "beneficiario_email_of"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4620
         TabIndex        =   11
         Text            =   "-"
         Top             =   3960
         Width           =   4515
      End
      Begin VB.TextBox txt_campo12 
         BackColor       =   &H00FFFFFF&
         DataField       =   "beneficiario_edif_depto_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8115
         TabIndex        =   21
         Top             =   7215
         Width           =   1020
      End
      Begin VB.TextBox txt_campo8 
         DataField       =   "beneficiario_email"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "-"
         Top             =   3960
         Width           =   4440
      End
      Begin VB.TextBox txt_campo10 
         DataField       =   "beneficiario_edif_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5640
         TabIndex        =   22
         Top             =   7215
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   1545
         Left            =   60
         TabIndex        =   25
         Top             =   825
         Width           =   9270
         Begin VB.TextBox Txt_descripcion 
            DataField       =   "beneficiario_denominacion"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   1125
            Width           =   9000
         End
         Begin VB.TextBox txt_campo4 
            BackColor       =   &H00FFFFFF&
            DataField       =   "beneficiario_nit"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   1
            Top             =   480
            Width           =   2220
         End
         Begin VB.TextBox LblInicial 
            BackColor       =   &H00FFFFFF&
            DataField       =   "beneficiario_iniciales"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   2805
            MaxLength       =   15
            TabIndex        =   2
            Top             =   480
            Width           =   1245
         End
         Begin VB.TextBox txt_codigo 
            BackColor       =   &H00FFFFFF&
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   65
            Top             =   480
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.TextBox txt_campo2 
            DataField       =   "beneficiario_segundo_apellido"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   2985
            TabIndex        =   19
            Top             =   1125
            Visible         =   0   'False
            Width           =   2550
         End
         Begin VB.TextBox txt_campo3 
            DataField       =   "beneficiario_nombres"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   5625
            TabIndex        =   20
            Top             =   1125
            Visible         =   0   'False
            Width           =   2670
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7E42
            DataField       =   "tipodoc_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5520
            TabIndex        =   3
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "tipodoc_codigo"
            BoundColumn     =   "tipodoc_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7E5B
            DataField       =   "depto_sigla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7920
            TabIndex        =   4
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_sigla"
            BoundColumn     =   "depto_sigla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7E74
            DataField       =   "tipodoc_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4320
            TabIndex        =   59
            Top             =   360
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "tipodoc_descripcion"
            BoundColumn     =   "tipodoc_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7E8D
            DataField       =   "depto_sigla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6960
            TabIndex        =   60
            Top             =   360
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_sigla"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "-"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   2520
            TabIndex        =   66
            Top             =   480
            Width           =   60
         End
         Begin VB.Label lbl_campo6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Expedido en"
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
            Left            =   7920
            TabIndex        =   64
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo.Documento"
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
            Left            =   5280
            TabIndex        =   63
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Iniciales Sucursal"
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
            Index           =   14
            Left            =   2640
            TabIndex        =   41
            Top             =   225
            Width           =   1560
         End
         Begin VB.Label lbl_titulo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Documento / NIT"
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
            Left            =   105
            TabIndex        =   29
            Top             =   225
            Width           =   1875
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombres"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5745
            TabIndex        =   28
            Top             =   855
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Denominación de la Entidad"
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
            Left            =   105
            TabIndex        =   27
            Top             =   855
            Width           =   2535
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Segundo Apellido"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3105
            TabIndex        =   26
            Top             =   855
            Visible         =   0   'False
            Width           =   1620
         End
      End
      Begin MSComCtl2.DTPicker DTP_Fecha1 
         DataField       =   "beneficiario_fecha_nacimiento"
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
         Height          =   315
         Left            =   7140
         TabIndex        =   6
         Top             =   2520
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   118554625
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7EA6
         DataField       =   "tipoben_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipoben_descripcion"
         BoundColumn     =   "tipoben_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7EBF
         DataField       =   "tipoben_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3240
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "tipoben_codigo"
         BoundColumn     =   "tipoben_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7ED8
         DataField       =   "zona_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   44
         Top             =   6240
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "zona_codigo"
         BoundColumn     =   "zona_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7EF1
         DataField       =   "zona_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   6480
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "zona_denominacion"
         BoundColumn     =   "zona_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7F0A
         DataField       =   "calle_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7200
         TabIndex        =   45
         Top             =   6240
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "calle_codigo"
         BoundColumn     =   "calle_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7F23
         DataField       =   "calle_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4500
         TabIndex        =   17
         Top             =   6480
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "calle_denominacion"
         BoundColumn     =   "calle_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7F3C
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2940
         TabIndex        =   46
         Top             =   7080
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7F56
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   7200
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux4 
         Bindings        =   "gw_p_gc_beneficiario_empresa.frx":7F70
         DataField       =   "pais_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   3240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "pais_cod_telefonico"
         BoundColumn     =   "pais_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lbl_calle 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Calle, Av, Plaza, etc."
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
         Left            =   4560
         TabIndex        =   76
         Top             =   6225
         Width           =   1800
      End
      Begin VB.Label lbl_zona 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   75
         Top             =   6220
         Width           =   1155
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Empresas"
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
         TabIndex        =   62
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Depto."
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
         Left            =   8100
         TabIndex        =   49
         Top             =   6960
         Width           =   1020
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Piso"
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
         Left            =   7005
         TabIndex        =   48
         Top             =   6960
         Width           =   855
      End
      Begin VB.Label lbl_campo17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Vivienda / Edificación"
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
         TabIndex        =   47
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   7680
         TabIndex        =   39
         Top             =   120
         Width           =   1455
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
         Height          =   315
         Left            =   7680
         TabIndex        =   38
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono Celular"
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
         Left            =   6360
         TabIndex        =   37
         Top             =   3000
         Width           =   1485
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono Fijo (Central)"
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
         Left            =   720
         TabIndex        =   36
         Top             =   3000
         Width           =   1980
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono(s) FAX"
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
         Left            =   3480
         TabIndex        =   35
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Creación"
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
         Left            =   5280
         TabIndex        =   34
         Top             =   2520
         Width           =   1710
      End
      Begin VB.Label LlbCargo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Vivienda"
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
         Left            =   5625
         TabIndex        =   33
         Top             =   6960
         Width           =   1245
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electrónico Contacto"
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
         Left            =   120
         TabIndex        =   32
         Top             =   3720
         Width           =   2505
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Página WEB"
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
         Left            =   4635
         TabIndex        =   31
         Top             =   3720
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla Entidad"
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
         TabIndex        =   30
         Top             =   2535
         Width           =   1200
      End
   End
   Begin Crystal.CrystalReport CR01 
      Left            =   720
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   10800
      Top             =   8400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   12960
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2160
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   0
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4320
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   2160
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6480
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   4320
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   8640
      Top             =   8400
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
End
Attribute VB_Name = "gw_p_gc_beneficiario_empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento de Beneficiarios
Option Explicit
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

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset

Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim SQL_FOR As String
Dim RUTA1 As String
Dim VAR_COD3 As String
Dim VAR_DIR As String
Dim VAR_COD2 As Double
'Dim SW As Boolean
'Dim CORREL, accion As Integer
'Dim swnuevo As Boolean
Dim CodBenef As String
Dim sino As String       ', NombreCarpeta, e


Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  If Ado_datos.Recordset.EOF Or Ado_datos.Recordset.BOF Then
'      BtnModificar.Enabled = False
'     ' BtnEliminar.Enabled = False
'      'TxtTipo.Text = Empty
'      txtCodigo.Text = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
'      txtDenominacion.Text = Empty
'      Exit Sub
'  End If
  If Ado_datos.Recordset.RecordCount > 0 Then
'    Select Case Ado_datos.Recordset.EditMode
'      Case adEditInProgress
'        Frame2.Enabled = False            'Verif. Nombre Proveedor JQA NOV-2009
'
'      Case adEditNone
'      Case adEditDelete
'      Case adEditAdd
'        Frame2.Enabled = True            'Verif. Nombre Proveedor JQA NOV-2009
'    End Select

'    If VAR_SW = "ADD" Or VAR_SW = "MOD" Then
'        txt_campo4.Visible = True
'        txt_codigo.Visible = False
'    Else
'        txt_campo4.Visible = False
'        txt_codigo.Visible = True
'    End If
    Ado_datos.Caption = CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & CStr(Ado_datos.Recordset.RecordCount)
  End If
End Sub
   
Private Sub BtnAux1_Click()
    'Validacion 1
    If dtc_codigo8 = "" Or dtc_codigo8 = "0" Then
        MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
        VAR_VAL = "ERR"
        Exit Sub
    End If
    fraDatos.Enabled = False
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

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     If VAR_SW = "ADD" Then
       var_cod = txt_campo4.Text + "-" + LblInicial
       Set rs_aux1 = New ADODB.Recordset
       SQL_FOR = "select * from gc_beneficiario where beneficiario_codigo = '" & var_cod & "'  "
       rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
       If rs_aux1.RecordCount > 0 Then
'                SW = True
                MsgBox " CODIGO DUPLICADO"
                Txt_descripcion.SetFocus
                Exit Sub
       End If
       If LblInicial.Text = "" Then
            VAR_COD3 = txt_campo4.Text
            rs_datos!beneficiario_iniciales = ""
       Else
            VAR_COD3 = Trim(txt_campo4.Text + "-" + Trim(LblInicial))     '
            rs_datos!beneficiario_iniciales = LblInicial
       End If
       rs_datos!beneficiario_codigo = VAR_COD3
       rs_datos!beneficiario_nit = txt_campo4.Text  'IIf(txt_campo4.Text = "", txt_codigo, txt_campo4.Text)
       rs_datos!estado_codigo = "REG"
        'rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
        'rs_datos!archivo_foto_cargado = "N"
        'rs_datos!ges_gestion = Year(Date)
        'rs_datos!correl_da = 0
        'db.Execute "Update gc_municipio Set correl_edif = CAST('" & dtc_aux2.Text & "' AS INT) + 1 Where munic_codigo= '" & dtc_codigo2.Text & "' "
    '     If txt_campo2.Text = "" Then txt_campo2.Text = "-"
    '     LblInicial = Trim(Left(txt_campo1.Text, 1)) + Trim(Left(txt_campo2.Text, 1)) + Trim(Left(txt_campo3.Text, 1))
    '     var_cod = Trim(txt_campo1.Text) + " " + Trim(txt_campo2.Text) + " " + Trim(txt_campo3.Text)
         rs_datos!depto_sigla = dtc_codigo3.Text
         rs_datos!tipodoc_codigo = dtc_codigo2.Text
         rs_datos!tipoben_codigo = dtc_codigo1.Text
         rs_datos!beneficiario_primer_apellido = IIf(txt_campo1.Text = "", "-", txt_campo1.Text)
         rs_datos!beneficiario_segundo_apellido = IIf(txt_campo2.Text = "", "-", txt_campo2.Text)
         rs_datos!beneficiario_nombres = IIf(txt_campo3.Text = "", "-", txt_campo3.Text)
         rs_datos!beneficiario_denominacion = Txt_descripcion.Text
         rs_datos!beneficiario_fecha_nacimiento = DTP_Fecha1.Value
         rs_datos!beneficiario_telefono_fijo = IIf(txt_campo5.Text = "", "0", txt_campo5.Text)
         rs_datos!beneficiario_telefono_Of = IIf(txt_campo6.Text = "", "0", txt_campo6.Text)
         rs_datos!beneficiario_telefono_Cel = IIf(txt_campo7.Text = "", "0", txt_campo7.Text)
         rs_datos!beneficiario_email = IIf(txt_campo8.Text = "", "-", txt_campo8.Text)
         rs_datos!beneficiario_email_of = IIf(txt_campo9.Text = "", "-", txt_campo9.Text)
         rs_datos!pais_codigo = IIf(dtc_codigo4.Text = "", "BOL", dtc_codigo4.Text)
         rs_datos!depto_codigo = dtc_codigo5.Text
         rs_datos!prov_codigo = dtc_codigo6.Text
         rs_datos!munic_codigo = dtc_codigo7.Text
         rs_datos!zona_codigo = IIf(dtc_codigo8.Text = "", "0", dtc_codigo8.Text)
         rs_datos!calle_codigo = IIf(dtc_codigo9.Text = "", "0", dtc_codigo9.Text)
         rs_datos!EDIF_CODIGO = dtc_codigo10.Text
         rs_datos!beneficiario_edif_nro = IIf(txt_campo10.Text = "", "0", txt_campo10.Text)
         rs_datos!beneficiario_edif_piso_nro = IIf(txt_campo11.Text = "", "0", txt_campo11.Text)
         rs_datos!beneficiario_edif_depto_nro = IIf(txt_campo12.Text = "", "0", txt_campo12.Text)
         rs_datos!beneficiario_domicilio_legal = dtc_desc8.Text + " - " + dtc_desc9.Text + " Nro. " + txt_campo10.Text
    '     If rs_datos!ARCHIVO_Foto = ".JPG" Or rs_datos!ARCHIVO_Foto = "" Then
    '        rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
    '     End If
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
     End If
     'If VAR_SW = "MOD" And (glusuario = "SQUISPE" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Or glusuario = "ADMIN") Then
     '   rs_datos!beneficiario_nit = Txt_campo4.Text
     'End If
     If VAR_SW = "MOD" Then
       VAR_COD3 = Ado_datos.Recordset!beneficiario_codigo   'Codigo Llave de la Tabla
       VAR_DIR = dtc_desc8.Text + " - " + dtc_desc9.Text + " Nro. " + txt_campo10.Text
       db.Execute "Update gc_beneficiario Set depto_sigla = '" & dtc_codigo3.Text & "', tipodoc_codigo = '" & dtc_codigo2.Text & "', tipoben_codigo = " & dtc_codigo1.Text & ", beneficiario_primer_apellido = '" & txt_campo1.Text & "', beneficiario_segundo_apellido = '" & txt_campo2.Text & "', beneficiario_nombres = '" & txt_campo3.Text & "', beneficiario_denominacion = '" & Txt_descripcion.Text & "', " & _
        " beneficiario_fecha_nacimiento = '" & DTP_Fecha1.Value & "', beneficiario_telefono_fijo = '" & txt_campo5.Text & "', beneficiario_telefono_Of = '" & txt_campo6.Text & "', beneficiario_telefono_Cel = '" & txt_campo7.Text & "', beneficiario_email = '" & txt_campo8.Text & "', " & _
        " beneficiario_email_of = '" & txt_campo9.Text & "', pais_codigo = '" & dtc_codigo4.Text & "', depto_codigo = '" & dtc_codigo5.Text & "', prov_codigo = '" & dtc_codigo6.Text & "', munic_codigo = '" & dtc_codigo7.Text & "', zona_codigo = '" & dtc_codigo8.Text & "', " & _
        " calle_codigo = '" & dtc_codigo9.Text & "', edif_codigo = '" & dtc_codigo10.Text & "', beneficiario_edif_nro = '" & txt_campo10.Text & "', beneficiario_edif_piso_nro = '" & txt_campo11.Text & "', beneficiario_edif_depto_nro = '" & txt_campo12.Text & "', beneficiario_nit = '" & txt_campo4.Text & "', " & _
        " beneficiario_domicilio_legal = '" & VAR_DIR & "', Fecha_Registro = '" & Date & "', usr_codigo = '" & glusuario & "' " & _
       " Where beneficiario_codigo= '" & VAR_COD3 & "'  "

     End If
     Call ABRIR_TABLA
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "beneficiario_codigo = '" & VAR_COD3 & "' ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     fraDatos.Enabled = False
     dg_datos.Enabled = True
     txt_campo4.Enabled = True
     LblInicial.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
    
End Sub

Private Sub valida_campos()
  If (txt_campo4.Text = "") Then
    MsgBox "Debe registrar el " + lbl_titulo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If txt_campo2.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If txt_campo3.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo2.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub
 
Private Sub graba_persona()
    Set rs_aux1 = New ADODB.Recordset
    rs_aux1.Open "select * from ro_personal_contratado where beneficiario_codigo = '" & txt_codigo.Text & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount = 0 Then
        rs_aux1.AddNew
        rs_aux1!beneficiario_codigo = txt_codigo.Text
        'rs_aux1!idfuncionario = CORREL
    'Else
        'MsgBox " YA EXISTE EL CODIGO ..."
    End If
        rs_aux1!ARCHIVO_Foto = Trim(LblInicial) + Ado_datos.Recordset("beneficiario_codigo") + ".JPG"
        rs_aux1!archivo_foto_cargado = "N"
        rs_aux1!archivo_hojavida = Trim(LblInicial) + Ado_datos.Recordset("beneficiario_codigo") + "_HV.PDF"
        rs_aux1!archivo_hojavida_cargado = "N"
        rs_aux1!archivo_respaldo = Trim(LblInicial) + Ado_datos.Recordset("beneficiario_codigo") + "_DOC.PDF"
        rs_aux1!archivo_respaldo_cargado = "N"
        rs_aux1!archivo_vacaciones = Trim(LblInicial) + Ado_datos.Recordset("beneficiario_codigo") + "_VAC.PDF"
        rs_aux1!archivo_vacaciones_cargado = "N"
        rs_aux1!archivo_otros = Trim(LblInicial) + Ado_datos.Recordset("beneficiario_codigo") + "_OTR.PDF"
        rs_aux1!archivo_otros_cargado = "N"
        rs_aux1!usr_codigo = glusuario 'frmLogin.txtUserName.Text
        rs_aux1!fecha_registro = Date
        'rs_aux1!hora_registro = Format(Time, "hh:mm:ss")
        rs_aux1!estado_codigo = "REG"
        rs_aux1.Update
End Sub

Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    fraDatos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
    txt_campo4.SetFocus
'    BtnVer.Visible = False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         If (Ado_datos.Recordset("tipoben_codigo") = 21) Then
            'db.Execute "insert into gc_usuarios(usr_codigo, beneficiario_codigo, usr_nombres, usr_primer_apellido, usr_segundo_apellido, usr_clave, IdNivelAcceso, estado_codigo, fecha_registro, dgral_codigo, da_codigo, unidad_codigo, ocup_codigo, usr_observaciones)" & _
            '"values ('" & Left(Ado_datos.Recordset("beneficiario_nombres"), 1) & "' + '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "', '" & Ado_datos.Recordset("beneficiario_codigo") & "','" & Ado_datos.Recordset("beneficiario_nombres") & "', '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "','" & Ado_datos.Recordset("beneficiario_segundo_apellido") & "','" & Ado_datos.Recordset("beneficiario_codigo") & "', '1', 'REG', '" & Date & "', '0', '0', '0', '0', '0') "
            RUTA1 = "CLIENTES\" + Trim(Ado_datos.Recordset("beneficiario_primer_apellido")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
         Else
            RUTA1 = "PROVEEDORES\" + Trim(Ado_datos.Recordset("beneficiario_primer_apellido")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
         End If
            MsgBox RUTA1
            MkDir RUTA1
            MkDir RUTA1 + "\CONTRATOS"
            MkDir RUTA1 + "\RESPALDOS"
            MkDir RUTA1 + "\HOJA_VIDA"
            MkDir RUTA1 + "\OTROS"
'            Call graba_persona
         
         rs_datos!estado_codigo = "APR"
         rs_datos!fecha_aprueba = Date
         rs_datos!usr_codigo_aprueba = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnEliminar_Click()
   If ExisteBenef(Ado_datos.Recordset!beneficiario_codigo) Then MsgBox "No se puede ANULAR un Beneficiario que ya fue utilizado ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   If ExisteBenef2(Ado_datos.Recordset!beneficiario_codigo) Then MsgBox "No se puede ANULAR un Beneficiario que ya fue utilizado ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   sino = MsgBox("Está Seguro de ANULAR el Registro?", vbYesNo + vbQuestion, "Atención")
   If Ado_datos.Recordset("estado_codigo") = "APR" Then
      If sino = vbYes Then
        Ado_datos.Recordset("estado_codigo") = "ERR"
        Ado_datos.Recordset("fecha_registro") = Date
        Ado_datos.Recordset("usr_codigo") = glusuario
        Ado_datos.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        If VAR_SW = "MOD" Then
            VAR_COD3 = Ado_datos.Recordset!beneficiario_codigo   'Codigo Llave de la Tabla
        End If
        rs_datos.CancelUpdate
        Call ABRIR_TABLAS_AUX
        Call ABRIR_TABLA
        If (dg_datos.SelBookmarks.Count <> 0) Then
           dg_datos.SelBookmarks.Remove 0
        End If
        If Ado_datos.Recordset.RecordCount > 0 Then
           rs_datos.Find "beneficiario_codigo = '" & VAR_COD3 & "' ", , , 1
           dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
        Else
           rs_datos.MoveLast
        End If
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        fraDatos.Enabled = False
        dg_datos.Enabled = True
        txt_campo4.Enabled = True
        LblInicial.Enabled = True
    End If
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If glusuario = "SQUISPE" Or glusuario = "RGIL" Or glusuario = "FDELGADILLO" Or glusuario = "ADMIN" Or glusuario = "MARTEAGA" Or glusuario = "APALACIOS" Or glusuario = "JCASTRO" Or glusuario = "CSALINAS" Or glusuario = "GSOLIZ" Or glusuario = "EVILLALOBOS" Or glusuario = "HMARIN" Or glusuario = "LVEDIA" Then
    fraDatos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    txt_campo4.Enabled = True
    LblInicial.Enabled = False
    VAR_SW = "MOD"
  Else
      If Ado_datos.Recordset!estado_codigo = "REG" Then
    '  lblStatus.Caption = "Modificar registro"
        fraDatos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        txt_campo4.Enabled = False
        LblInicial.Enabled = False
        VAR_SW = "MOD"
    '    BtnVer.Visible = True
      Else
         MsgBox "No se puede MODIFICAR un registro Aprobado (APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
      End If
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description

End Sub

Private Sub BtnVer_Click()
   On Error GoTo QError
    If Ado_datos.Recordset!ARCHIVO_Foto = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FOT"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(ado_datos.Recordset!iniciales) & "-" & Trim(ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FOT"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(ado_datos.Recordset!iniciales) & "-" & Trim(ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(ado_datos.Recordset!iniciales) + "-" + Trim(ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        ARCH_FOTO = App.Path + "\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
'    End If
    'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + ado_datos.Recordset!codigo_beneficiario + "\" + ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    CodBenef = Ado_datos.Recordset!codigo_beneficiario
    If Guardar_Imagen(db, "Select Foto From fc_beneficiario Where codigo_beneficiario= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
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

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
     CR01.WindowShowPrintSetupBtn = True
     CR01.WindowShowRefreshBtn = True
     CR01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_beneficiario_Empresa.rpt"
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

CR01.WindowState = crptMaximized
End Sub

Private Sub BtnSalir_Click()
'  If glPersNew = "CMP" Then
'    frmComprasDirectas.DtcProv = Ado_datos.Recordset("codigo_beneficiario")
'    frmComprasDirectas.cboListaProv2 = Ado_datos.Recordset("denominacion_Beneficiario")
'    frmComprasDirectas.txtProv = Ado_datos.Recordset("codigo_beneficiario")
'    frmComprasDirectas.txtProveedor = Ado_datos.Recordset("denominacion_Beneficiario")
'  End If
  Unload Me
End Sub

Private Sub CmdCancelaDet_Click()
    fraDatos.Enabled = True
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
  If dtc_codigo8 = "" Or dtc_codigo8 = "0" Then
    MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  'INI Graba Calle
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select max(calle_codigo) as Codigo from gc_calles where zona_codigo = " & dtc_codigo8.Text & "    ", db, adOpenStatic
    If rs_aux2.RecordCount > 0 And Not IsNull(rs_aux2!Codigo) Then
        VAR_COD2 = Round(CDbl(rs_aux2!Codigo) + 1, 0)
    Else
        VAR_COD2 = 1
        VAR_COD2 = (Val(dtc_codigo8.Text) * 100) + 1
    End If
    db.Execute "insert into gc_calles(zona_codigo, calle_codigo, calle_denominacion, calle_tipo, correl, estado_codigo, fecha_registro, usr_codigo)" & _
    "values ('" & dtc_codigo8.Text & "', " & VAR_COD2 & ", '" & Txt_descripcion11.Text & "', '" & dtc_codigo11.Text & "', '0', 'APR', '" & Date & "', '" & glusuario & "') "
    
   'FIN Graba Calle
    'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
    db.Execute "Update gc_zonas Set correl = " & VAR_COD2 & " Where zona_codigo= '" & dtc_codigo8.Text & "' "
    'gc_calles
    Call pnivel6(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
    
    dtc_codigo9.Text = VAR_COD2
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    fraDatos.Enabled = True
    Fra_aux1.Visible = False

End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_aux4.BoundText
    dtc_codigo4.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux7_Click(Area As Integer)

End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
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

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
    Call pnivel5(dtc_codigo7.BoundText)
    dtc_desc8.Enabled = True
    Call pnivel7(dtc_codigo7.BoundText)
    dtc_desc10.Enabled = True
End Sub
   
Private Sub pnivel5(codigo7 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo7 & "' order by zona_denominacion"
   Set dtc_codigo8.RowSource = Nothing
   Set dtc_codigo8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo8.ReFill
   dtc_codigo8.BoundText = Empty
   
   Set dtc_desc8.RowSource = Nothing
   Set dtc_desc8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc8.ReFill
   dtc_desc8.BoundText = Empty
End Sub

Private Sub pnivel7(codigo9 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_edificaciones where munic_codigo = '" & codigo9 & "' order by edif_descripcion"
   Set dtc_codigo10.RowSource = Nothing
   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty
   
   Set dtc_desc10.RowSource = Nothing
   Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub
   
Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    Call pnivel6(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
End Sub

Private Sub pnivel6(codigo8 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_calles where zona_codigo = '" & codigo8 & "'"
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
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    'txt_codigo.Enabled = True
'    mbDataChanged = False
    fraDatos.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
    Fra_aux1.Visible = False
End Sub

Private Sub ABRIR_TABLA()
   Set rs_datos = New ADODB.Recordset
   If rs_datos.State = 1 Then rs_datos.Close
   queryinicial = "select * from gc_beneficiario WHERE beneficiario_codigo <> ' ' and beneficiario_codigo <> '0' and tipoben_codigo > 19 "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rs_datos.Sort = "beneficiario_denominacion"
   Set Ado_datos.Recordset = rs_datos
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'     FrmVentas.DtcNIT = Ado_datos.Recordset("codigo_beneficiario")
'     FrmVentas.DtcdesNIT = Ado_datos.Recordset("denominacion_Beneficiario")
'  End If
'
'  glPersNew = "N"
   
   If (rs_datos.State = adStateClosed) Then rs_datos.Close
   'Set rs_datos = Nothing
End Sub

Private Sub ABRIR_TABLAS_AUX()
  'carga    fc_tipo_beneficiario
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "SELECT * FROM gc_tipo_beneficiario WHERE tipoben_codigo > 19 ORDER BY tipoben_descripcion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
      'gc_tipo_documento_id     'Tipo Doc. de Id.
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "select * from gc_tipo_documento_id", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'gc_Departamento    'Expedido en...
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_departamento order by depto_sigla", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    'gc_pais
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_pais  where estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'gc_Departamento  '<>
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from gc_departamento order by depto_descripcion", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    'gc_provincia
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_provincia ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    'gc_municipio
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from gc_municipio ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'gc_zonas
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_zonas ", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    'gc_calles
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_calles ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    'gc_edificaciones
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificaciones order by edif_descripcion", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
End Sub

Private Sub txt_campo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_campo1_LostFocus()
'    Txt_descripcion.Text = txt_campo1.Text + " " + txt_campo2.Text + " " + txt_campo3.Text
End Sub

Private Sub txt_campo3_LostFocus()
'    Txt_descripcion.Text = txt_campo1.Text + " " + txt_campo2.Text + " " + txt_campo3.Text
End Sub

Private Function ExisteBenef(CodBenef As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo_resp = '" & CodBenef & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteBenef = rs!Cuantos > 0
End Function

Private Function ExisteBenef2(CodBenef As String) As Boolean
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo = '" & CodBenef & "'"
    rs2.Open GlSqlAux, db, adOpenStatic
    ExisteBenef2 = rs2!Cuantos > 0
End Function


Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
'    dtc_aux2.BoundText = dtc_codigo2.BoundText
'    dtc_campo2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    dtc_aux4.BoundText = dtc_desc4.BoundText
    Call pnivel2(dtc_codigo4.BoundText)
    dtc_desc5.Enabled = True
End Sub
   
Private Sub pnivel2(codigo4 As String)
   Dim strConsultaF As String
     
   strConsultaF = "select * from gc_departamento where pais_codigo = '" & codigo4 & "'"
   Set dtc_codigo5.RowSource = Nothing
   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_codigo5.ReFill
   dtc_codigo5.BoundText = Empty
   
   Set dtc_desc5.RowSource = Nothing
   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_desc5.ReFill
   dtc_desc5.BoundText = Empty

End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    Call pnivel3(dtc_codigo5.BoundText)
    dtc_desc6.Enabled = True
End Sub
   
Private Sub pnivel3(codigo5 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo5 & "'"
   Set dtc_codigo6.RowSource = Nothing
   Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo6.ReFill
   dtc_codigo6.BoundText = Empty
   
   Set dtc_desc6.RowSource = Nothing
   Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc6.ReFill
   dtc_desc6.BoundText = Empty
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
    Call pnivel4(dtc_codigo6.BoundText)
    dtc_desc7.Enabled = True
End Sub
   
Private Sub pnivel4(codigo6 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo6 & "'"
   Set dtc_codigo7.RowSource = Nothing
   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo7.ReFill
   dtc_codigo7.BoundText = Empty
   
   Set dtc_desc7.RowSource = Nothing
   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc7.ReFill
   dtc_desc7.BoundText = Empty
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

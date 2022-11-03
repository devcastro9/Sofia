VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_gc_zonas 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Zonas"
   ClientHeight    =   7455
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   12945
   Icon            =   "frm_gc_zonas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   12945
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_gc_zonas.frx":0A02
      ScaleHeight     =   960
      ScaleWidth      =   12120
      TabIndex        =   20
      Top             =   120
      Width           =   12180
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_gc_zonas.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_gc_zonas.frx":6CC3E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_gc_zonas.frx":6CE48
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_gc_zonas.frx":6D400
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_gc_zonas.frx":6D9BD
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_gc_zonas.frx":6DBC7
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_gc_zonas.frx":6E891
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_gc_zonas.frx":6EE71
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZONAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   8700
         TabIndex        =   23
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00FFFFC0&
      Height          =   5655
      Left            =   120
      TabIndex        =   25
      Top             =   1200
      Width           =   6255
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "frm_gc_zonas.frx":6F495
         Height          =   4815
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   8493
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "zona_codigo"
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
            DataField       =   "zona_denominacion"
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
            DataField       =   "pais_codigo"
            Caption         =   "Pais"
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
            DataField       =   "depto_codigo"
            Caption         =   "Depto"
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
            DataField       =   "prov_codigo"
            Caption         =   "Provincia"
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
            DataField       =   "munic_codigo"
            Caption         =   "Municipio"
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
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3990.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   689.953
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
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5160
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
         Caption         =   " <-- Inicio                              Navegar                                Fin -->"
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
      Height          =   5655
      Left            =   6520
      TabIndex        =   18
      Top             =   1200
      Width           =   5775
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Ubicacion Geográfica"
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
         Height          =   3255
         Left            =   40
         TabIndex        =   29
         Top             =   2280
         Width           =   5670
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "frm_gc_zonas.frx":6F4AD
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   30
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
            Bindings        =   "frm_gc_zonas.frx":6F4C6
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3000
            TabIndex        =   31
            Top             =   2595
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
            Bindings        =   "frm_gc_zonas.frx":6F4DF
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   240
            TabIndex        =   32
            Top             =   2715
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "frm_gc_zonas.frx":6F4F8
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   33
            Top             =   1755
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
            Bindings        =   "frm_gc_zonas.frx":6F511
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   240
            TabIndex        =   34
            Top             =   1995
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "frm_gc_zonas.frx":6F52A
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3000
            TabIndex        =   35
            Top             =   1155
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
            Bindings        =   "frm_gc_zonas.frx":6F543
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "frm_gc_zonas.frx":6F55C
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   240
            TabIndex        =   37
            Top             =   600
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_descripcion"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
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
            Left            =   240
            TabIndex        =   41
            Top             =   1740
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
            Left            =   240
            TabIndex        =   40
            Top             =   1060
            Width           =   1290
         End
         Begin VB.Label lblLabels 
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
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   2460
            Width           =   855
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
            Left            =   240
            TabIndex        =   38
            Top             =   340
            Width           =   405
         End
      End
      Begin VB.TextBox Txt_descripcion 
         DataField       =   "zona_denominacion"
         DataSource      =   "Ado_datos"
         Height          =   645
         Left            =   360
         TabIndex        =   0
         Text            =   "-"
         Top             =   1500
         Width           =   4935
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "REG"
         DataField       =   "zona_codigo"
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
         Left            =   360
         TabIndex        =   28
         Top             =   680
         Width           =   1815
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Denominación de la Zona y Barrio"
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
         TabIndex        =   27
         Top             =   1200
         Width           =   3045
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código de la Zona"
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
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   1665
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
         Left            =   3960
         TabIndex        =   9
         Top             =   720
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
         Left            =   3880
         TabIndex        =   19
         Top             =   420
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
      ScaleWidth      =   12945
      TabIndex        =   12
      Top             =   7455
      Width           =   12945
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   17
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2400
      Top             =   6840
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
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_gc_zonas.frx":6F575
      ScaleHeight     =   960
      ScaleWidth      =   12120
      TabIndex        =   21
      Top             =   120
      Width           =   12180
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_gc_zonas.frx":DB5A7
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_gc_zonas.frx":DB7B1
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GERENCIA GENERAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   7290
         TabIndex        =   22
         Top             =   300
         Width           =   3900
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   120
      Top             =   6960
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
      Left            =   2880
      Top             =   6960
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
      Left            =   5280
      Top             =   6960
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
      Left            =   7680
      Top             =   6960
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
End
Attribute VB_Name = "frm_gc_zonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String

Dim mvBookMark, marca1 As Variant
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
        dtc_desc1.Enabled = True
    End If
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!calle_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If ExisteReg2(Ado_datos.Recordset!calle_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If rs_datos!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ERR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
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
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "Select max(zona_codigo) as Codigo from gc_zonas where munic_codigo = " & dtc_codigo7.Text & "    ", db, adOpenStatic
        If rs_aux2!Codigo <> "NULL" Then
            var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
        Else
            var_cod = Val(dtc_codigo7.Text) * 10
        End If
        
        rs_datos!zona_codigo = var_cod
        txt_codigo = var_cod            ' Esto para codigos correlativos
        rs_datos!estado_codigo = "REG"  ' no cambia
        rs_datos!CORREL = 0
        rs_datos!pais_codigo = dtc_codigo4.Text   'Codigo del padre 1
        rs_datos!depto_codigo = dtc_codigo5.Text   'Codigo del padre 2
        rs_datos!prov_codigo = dtc_codigo6.Text   'Codigo del padre 3
        rs_datos!munic_codigo = dtc_codigo7.Text   'Codigo del padre 4
        'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
        'db.Execute "Update gc_zonas Set correl = " & var_cod & " Where zona_codigo= '" & dtc_codigo1.Text & "' "
     End If
     rs_datos!zona_denominacion = Txt_descripcion.Text
     rs_datos!fecha_registro = Date     ' no cambia
     rs_datos!usr_codigo = glusuario    ' no cambia
     rs_datos.UpdateBatch adAffectAll
     marca1 = Ado_datos.Recordset.Bookmark
     Call ABRIR_TABLA
     'rs_datos.MoveLast
     mbDataChanged = False
     Ado_datos.Recordset.Move marca1 - 1
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      'txt_codigo.Enabled = True
      'dtc_desc1.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'habilitar codigo cuando se transcribe
'  If txt_codigo = "" Then
'    MsgBox "Debe registrar el " + lbl_codigo.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo6.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo7 = "" Then
    MsgBox "Debe registrar: " + lbl_enlace2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  CR01.WindowShowPrintSetupBtn = True
  CR01.WindowShowRefreshBtn = True
  CR01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_zonas.rpt"
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If rs_datos!estado_codigo = "REG" Then
'  lblStatus.Caption = "Modificar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
    Frame1.Enabled = False
    'txt_codigo.Enabled = False
    'dtc_desc1.Enabled = False
  Else
      MsgBox "No se puede MODIFICAR un registro Aprobado(APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
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

Private Sub DtcUE_Click(Area As Integer)
    DtcUE_Des.BoundText = DtcUE.BoundText
End Sub

Private Sub DtcUE_Des_Click(Area As Integer)
    DtcUE.BoundText = DtcUE_Des.BoundText
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

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    Call pnivel2(dtc_codigo4.BoundText)
    dtc_desc5.Enabled = True
End Sub
   
Private Sub pnivel2(codigo4 As String)
   Dim strConsultaF As String
   strConsultaF = "select * from gc_departamento where pais_codigo = '" & codigo4 & "'"
   
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
   dtc_codigo7.ReFill
   dtc_codigo7.BoundText = Empty
   
   Set dtc_desc7.RowSource = Nothing
   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc7.ReFill
   dtc_desc7.BoundText = Empty
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from gc_zonas  "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLAS_AUX()
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    rs_datos1.Open "Select * from gc_zonas order by zona_codigo", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'
'    Set rs_datos2 = New ADODB.Recordset
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    rs_datos2.Open "Select * from gc_calle_tipo order by calle_tipo_nombre", db, adOpenStatic
'    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_pais
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_pais where estado_codigo = 'APR' ", db, adOpenStatic
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
    Txt_descripcion.SetFocus
    'txt_codigo.SetFocus
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
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM gc_beneficiario WHERE estado_codigo = 'APR' and calle_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Function ExisteReg2(Unidad As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM gc_edificaciones WHERE estado_codigo = 'APR' and calle_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg2 = rs!Cuantos > 0
End Function

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

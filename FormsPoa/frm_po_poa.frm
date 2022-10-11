VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_po_poa 
   BackColor       =   &H00000000&
   Caption         =   "Planificación - Plan de Operaciones - POA Consolidado"
   ClientHeight    =   9765
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "frm_po_poa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_po_poa.frx":0A02
      ScaleHeight     =   960
      ScaleWidth      =   16395
      TabIndex        =   62
      Top             =   120
      Width           =   16455
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_po_poa.frx":1982B
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_po_poa.frx":19A35
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_po_poa.frx":1A059
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_po_poa.frx":1A639
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_po_poa.frx":1B303
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_po_poa.frx":1B50D
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_po_poa.frx":1BACA
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_po_poa.frx":1C082
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_po_poa.frx":1C28C
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO1"
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
         Left            =   11400
         TabIndex        =   72
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_po_poa.frx":1C6CE
      ScaleHeight     =   915
      ScaleWidth      =   16395
      TabIndex        =   58
      Top             =   120
      Width           =   16455
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_po_poa.frx":354F7
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_po_poa.frx":35701
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
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
         Left            =   11460
         TabIndex        =   61
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   9000
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   5895
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "frm_po_poa.frx":3590B
         Height          =   8250
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   14552
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "poa_codigo"
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
            DataField       =   "poa_descripcion"
            Caption         =   "Descripción"
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
         BeginProperty Column04 
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Responsable Ejecución"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Autoridad Aprueba"
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
            DataField       =   "doc_codigo"
            Caption         =   "Código Registro"
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
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3660.095
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   705.26
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
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   33
         Top             =   8625
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   34
         Top             =   8625
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   8565
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
      BackColor       =   &H00000000&
      Height          =   9000
      Left            =   6120
      TabIndex        =   11
      Top             =   1200
      Width           =   10455
      Begin MSDataListLib.DataCombo dtc_codigo12 
         Bindings        =   "frm_po_poa.frx":35923
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   117
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "insumo_codigo"
         BoundColumn     =   "poa_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "frm_po_poa.frx":3593D
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   95
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tarea_codigo"
         BoundColumn     =   "poa_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "frm_po_poa.frx":35956
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   88
         Top             =   2175
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "actividad_codigo"
         BoundColumn     =   "poa_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "frm_po_poa.frx":3596F
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   79
         Top             =   1485
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "especifico_codigo"
         BoundColumn     =   "poa_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc12 
         Bindings        =   "frm_po_poa.frx":35988
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4515
         TabIndex        =   118
         Top             =   3525
         Visible         =   0   'False
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "poa_descripcion"
         BoundColumn     =   "poa_codigo"
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
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "frm_po_poa.frx":359A2
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4515
         TabIndex        =   97
         Top             =   2865
         Visible         =   0   'False
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "poa_descripcion"
         BoundColumn     =   "poa_codigo"
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
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "frm_po_poa.frx":359BB
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4515
         TabIndex        =   89
         Top             =   2220
         Visible         =   0   'False
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "poa_descripcion"
         BoundColumn     =   "poa_codigo"
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
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "frm_po_poa.frx":359D4
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4515
         TabIndex        =   80
         Top             =   1455
         Visible         =   0   'False
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "poa_descripcion"
         BoundColumn     =   "poa_codigo"
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
      Begin VB.TextBox txt_aux12 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1500
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   122
         Top             =   3465
         Width           =   8715
      End
      Begin VB.TextBox txt_aux7 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1500
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   121
         Top             =   2790
         Width           =   8715
      End
      Begin VB.TextBox txt_aux6 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1500
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   120
         Top             =   2115
         Width           =   8715
      End
      Begin VB.TextBox txt_aux5 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1500
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   119
         Top             =   1440
         Width           =   8715
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_po_poa.frx":359ED
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4560
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "poa_descripcion"
         BoundColumn     =   "poa_codigo"
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
      Begin VB.TextBox txt_campo4 
         Alignment       =   2  'Center
         DataField       =   "ges_gestion"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   9000
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   8355
         TabIndex        =   36
         Top             =   6300
         Width           =   285
      End
      Begin VB.TextBox txt_campo3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "poa_fuente_verificacion"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   7785
         Width           =   6360
      End
      Begin MSDataListLib.DataCombo dtc_aux11 
         Bindings        =   "frm_po_poa.frx":35A06
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6360
         TabIndex        =   26
         Top             =   5040
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
         Bindings        =   "frm_po_poa.frx":35A20
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   25
         Top             =   5040
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "frm_po_poa.frx":35A3A
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6360
         TabIndex        =   24
         Top             =   4080
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
         Bindings        =   "frm_po_poa.frx":35A53
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "estrategico_codigo"
         BoundColumn     =   "poa_codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "fecha_registro"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   4725
         TabIndex        =   1
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   83558401
         CurrentDate     =   41678
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_po_poa.frx":35A6C
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   16
         Top             =   4560
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_descripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1500
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   765
         Width           =   8715
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_po_poa.frx":35A85
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4035
         TabIndex        =   15
         Top             =   4755
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_po_poa.frx":35A9E
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   17
         Top             =   4080
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
         Bindings        =   "frm_po_poa.frx":35AB7
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   0
         Top             =   4275
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "frm_po_poa.frx":35AD0
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4035
         TabIndex        =   2
         Top             =   5235
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
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "frm_po_poa.frx":35AEA
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5700
         TabIndex        =   40
         Top             =   5880
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
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "frm_po_poa.frx":35B03
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3045
         TabIndex        =   41
         Top             =   5805
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "clasif_descripcion"
         BoundColumn     =   "clasif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "frm_po_poa.frx":35B1C
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1995
         TabIndex        =   43
         Top             =   6285
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "frm_po_poa.frx":35B35
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3045
         TabIndex        =   44
         Top             =   6285
         Width           =   5595
         _ExtentX        =   9869
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
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_po_poa.frx":35B4E
         DataField       =   "org_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2925
         TabIndex        =   75
         Top             =   6885
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "org_descripcion"
         BoundColumn     =   "org_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_po_poa.frx":35B67
         DataField       =   "org_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   76
         Top             =   6720
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "org_codigo"
         BoundColumn     =   "org_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "frm_po_poa.frx":35B80
         DataField       =   "par_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2925
         TabIndex        =   105
         Top             =   7335
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "par_descripcion"
         BoundColumn     =   "par_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "frm_po_poa.frx":35B9A
         DataField       =   "par_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   106
         Top             =   7200
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "par_codigo"
         BoundColumn     =   "par_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "frm_po_poa.frx":35BB4
         DataField       =   "org_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6240
         TabIndex        =   123
         Top             =   6720
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "fte_codigo"
         BoundColumn     =   "org_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Obj. General"
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
         Left            =   180
         TabIndex        =   74
         Top             =   785
         Width           =   1140
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   3480
         TabIndex        =   116
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   3120
         TabIndex        =   115
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   2760
         TabIndex        =   114
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   2400
         TabIndex        =   113
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label lbl_ins1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   3600
         TabIndex        =   112
         Top             =   3360
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbl_tar2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   3240
         TabIndex        =   111
         Top             =   3360
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbl_act3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         TabIndex        =   110
         Top             =   3360
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbl_objesp4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   2520
         TabIndex        =   109
         Top             =   3360
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbl_objgral5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         TabIndex        =   108
         Top             =   3735
         Width           =   225
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Partida por Objeto del Gasto"
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
         TabIndex        =   107
         Top             =   7365
         Width           =   2550
      End
      Begin VB.Label lbl_tar1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   3240
         TabIndex        =   104
         Top             =   2715
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   3120
         TabIndex        =   103
         Top             =   2715
         Width           =   45
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   2760
         TabIndex        =   102
         Top             =   2715
         Width           =   45
      End
      Begin VB.Label lbl_act2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         TabIndex        =   101
         Top             =   2715
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbl_objesp3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   2520
         TabIndex        =   100
         Top             =   2715
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   2400
         TabIndex        =   99
         Top             =   2715
         Width           =   45
      End
      Begin VB.Label lbl_objgral4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         TabIndex        =   98
         Top             =   3060
         Width           =   225
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Tarea POA"
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
         Index           =   16
         Left            =   180
         TabIndex        =   96
         Top             =   2810
         Width           =   1095
      End
      Begin VB.Label lbl_act1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   1080
         TabIndex        =   94
         Top             =   2325
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbl_objesp2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   720
         TabIndex        =   93
         Top             =   2325
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   960
         TabIndex        =   92
         Top             =   2325
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   600
         TabIndex        =   91
         Top             =   2325
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lbl_objgral3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         TabIndex        =   90
         Top             =   2385
         Width           =   225
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   8760
         X2              =   8760
         Y1              =   4100
         Y2              =   9000
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Actividad POA"
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
         Index           =   15
         Left            =   180
         TabIndex        =   87
         Top             =   2135
         Width           =   1380
      End
      Begin VB.Label lbl_objesp1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         Left            =   840
         TabIndex        =   86
         Top             =   1680
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
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
         Left            =   765
         TabIndex        =   85
         Top             =   1680
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lbl_objgral2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         TabIndex        =   84
         Top             =   1710
         Width           =   225
      End
      Begin VB.Label lbl_objgral1 
         BackColor       =   &H00000000&
         Caption         =   "0"
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
         TabIndex        =   83
         Top             =   1035
         Width           =   225
      End
      Begin VB.Label DTPfecha3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_fecha_final"
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
         Left            =   8880
         TabIndex        =   82
         Top             =   8205
         Width           =   1455
      End
      Begin VB.Label DTPfecha2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_fecha_inicio"
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
         Left            =   8880
         TabIndex        =   81
         Top             =   7320
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Obj.Específico"
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
         Index           =   14
         Left            =   180
         TabIndex        =   78
         Top             =   1460
         Width           =   1305
      End
      Begin VB.Label lbl_campo3a 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Organismo Financiador"
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
         TabIndex        =   77
         Top             =   6900
         Width           =   2100
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite Documento"
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
         Left            =   8925
         TabIndex        =   73
         Top             =   5925
         Width           =   1440
      End
      Begin VB.Label txt_monto4 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_monto_vigente_bs"
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
         Left            =   7080
         TabIndex        =   57
         Top             =   8565
         Width           =   1455
      End
      Begin VB.Label txt_monto3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_monto_modif"
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
         Left            =   4800
         TabIndex        =   56
         Top             =   8565
         Width           =   1455
      End
      Begin VB.Label txt_monto2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_monto_adiciones"
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
         Left            =   2520
         TabIndex        =   55
         Top             =   8565
         Width           =   1455
      End
      Begin VB.Label txt_monto1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_monto_origen"
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
         TabIndex        =   54
         Top             =   8565
         Width           =   1455
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fuente de Verificación"
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
         TabIndex        =   53
         Top             =   7800
         Width           =   1995
      End
      Begin VB.Label txt_campo5 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_nivel"
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
         Left            =   7680
         TabIndex        =   42
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nivel POA"
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
         Index           =   11
         Left            =   6645
         TabIndex        =   52
         Top             =   285
         Width           =   930
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Monto Vigente Bs."
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
         Index           =   10
         Left            =   6960
         TabIndex        =   51
         Top             =   8280
         Width           =   1620
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Modificaciones Bs."
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
         Index           =   9
         Left            =   4800
         TabIndex        =   50
         Top             =   8280
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Adiciones Bs."
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
         Index           =   8
         Left            =   2640
         TabIndex        =   49
         Top             =   8280
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Monto Origen Bs."
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
         Index           =   7
         Left            =   180
         TabIndex        =   48
         Top             =   8280
         Width           =   1530
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Final"
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
         Index           =   5
         Left            =   9000
         TabIndex        =   46
         Top             =   7920
         Width           =   1170
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Inicio"
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
         Index           =   4
         Left            =   8880
         TabIndex        =   45
         Top             =   7005
         Width           =   1455
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   10815
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Valor de la Meta"
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
         Index           =   3
         Left            =   8880
         TabIndex        =   39
         Top             =   4995
         Width           =   1470
      End
      Begin VB.Label lbl_campo8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Clasificación de Documentos"
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
         Top             =   5820
         Width           =   2610
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gestión Creación"
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
         Left            =   8835
         TabIndex        =   37
         Top             =   4245
         Width           =   1545
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
         DataField       =   "poa_num_respaldo"
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
         Left            =   8820
         TabIndex        =   35
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label lbl_descripcion 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo POA"
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
         Left            =   180
         TabIndex        =   32
         Top             =   3485
         Width           =   1185
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   31
         Top             =   6300
         Width           =   1755
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Responsable de la Ejecución del Objetivo:"
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
         TabIndex        =   30
         Top             =   4785
         Width           =   3825
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Autoridad que Aprueba su Cumplimiento:"
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
         TabIndex        =   29
         Top             =   5265
         Width           =   3645
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad Responsable de la Ejecución"
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
         TabIndex        =   28
         Top             =   4305
         Width           =   3360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   10815
         Y1              =   5655
         Y2              =   5655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   10815
         Y1              =   4100
         Y2              =   4100
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "poa_codigo"
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
         Left            =   1380
         TabIndex        =   23
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "poa_meta"
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
         Left            =   9000
         TabIndex        =   22
         Top             =   5280
         Width           =   1245
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   12
         Left            =   3000
         TabIndex        =   20
         Top             =   285
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
         Height          =   300
         Left            =   9360
         TabIndex        =   4
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Codigo POA"
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
         TabIndex        =   13
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label lblLabels 
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
         Left            =   8595
         TabIndex        =   12
         Top             =   285
         Width           =   645
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
      ScaleWidth      =   15120
      TabIndex        =   5
      Top             =   9765
      Width           =   15120
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   10
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
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
      Left            =   13680
      Top             =   10680
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
      Left            =   11280
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
      Top             =   11040
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
      Top             =   11040
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   6720
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
End
Attribute VB_Name = "frm_po_poa"
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
Dim rs_datos12 As New ADODB.Recordset

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
Dim VAR_UNI As String
Dim sino As String
Dim parametro As String
Dim VAR_AUX, VAR_CONT2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
   If Ado_datos.Recordset!org_codigo = "" Or Ado_datos.Recordset!par_codigo = "" Or Ado_datos.Recordset!fte_codigo = "" Then
        'MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
        MsgBox "No se puede APROBAR, Verifique los Datos y Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        Exit Sub
   End If
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
'        Select Case dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Case "5"    ' SERVICIO MODERNIZACION
'        End Select
        
         Dim cmdppto As ADODB.Command
         Set cmdppto = New ADODB.Command
         db.BeginTrans
        ' Set cmdppto = New ADODB.Command
         With cmdppto
           .ActiveConnection = db
           .CommandType = adCmdStoredProc
           .CommandText = "fp_insertar_fo_ppto_formulacion_gasto"
           .Parameters("@ges_gestion") = glGestion
           .Parameters("@dgral_codigo") = Ado_datos.Recordset!dgral_codigo
           .Parameters("@da_codigo") = Ado_datos.Recordset!da_codigo
           
           .Parameters("@pro_codigo") = Ado_datos.Recordset!pro_codigo
           .Parameters("@fte_codigo") = Ado_datos.Recordset!fte_codigo
           .Parameters("@org_codigo") = Ado_datos.Recordset!org_codigo
           
           .Parameters("@par_codigo") = Ado_datos.Recordset!par_codigo
           .Parameters("@ppto_formulado") = IIf(IsNull(Ado_datos.Recordset!poa_monto_vigente_bs), 0, Ado_datos.Recordset!poa_monto_vigente_bs)
           .Parameters("@estado_codigo") = "REG"
           
           .Parameters("@fecha_registro") = Date
           .Parameters("@hora_registro") = Format(Time, "hh:mm:ss")
           .Parameters("@usr_usuario") = glusuario
           .Execute
         End With
         db.CommitTrans
  
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = txt_campo1.Caption
        'REVISAR !!! JQA 2014_07_08
        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!estado_codigo = "APR"
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
        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
    If ExisteReg(Ado_datos.Recordset!estrategico_codigo, Ado_datos.Recordset!especifico_codigo, Ado_datos.Recordset!actividad_codigo, Ado_datos.Recordset!tarea_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
    If rs_datos!estado_codigo = "APR" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
          rs_datos!estado_codigo = "ERR"
          rs_datos!fecha_registro = Date
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
        'VAR_UNI = dtc_codigo1.Text
        'var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        SQL_FOR = "select * from pc_poa_tarea where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " AND especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " AND actividad_codigo = " & dtc_codigo6.Text & "  AND tarea_codigo = " & dtc_codigo7.Text & " "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            var_cod = rs_aux1!correl_insumo + 1
            'Actualiza correaltivo ...
            db.Execute "Update pc_poa_tarea Set correl_insumo = " & var_cod & " where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " AND especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " AND actividad_codigo = " & dtc_codigo6.Text & "  AND tarea_codigo = " & dtc_codigo7.Text & " "
        Else
            var_cod = 1
        End If
        rs_datos!estrategico_codigo = dtc_codigo2.Text
        rs_datos!especifico_codigo = dtc_codigo5.Text
        rs_datos!actividad_codigo = dtc_codigo6.Text
        rs_datos!tarea_codigo = dtc_codigo7.Text
        txt_codigo.Caption = Trim(dtc_codigo2.Text) + "." + Trim(dtc_codigo5.Text) + "." + Trim(dtc_codigo6.Text) + "." + Trim(dtc_codigo7.Text) + "." + var_cod
        rs_datos!insumo_codigo = var_cod
        rs_datos!poa_codigo = txt_codigo.Caption
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = glGestion    ' Year(Date)   'no cambia
        
        rs_datos!poa_meta = "100"
        rs_datos!poa_codigo_anterior = txt_codigo.Caption
        'rs_datos!archivo_respaldo = "sin_nombre"
        'rs_datos!archivo_respaldo_cargado = "N"
        'rs_datos!correl_insumo = 0
        'rs_datos!doc_numero = 0
        rs_datos!poa_nivel = 5
        rs_datos!poa_monto_origen = 0       'txt_monto1.Text
        rs_datos!poa_monto_adiciones = 0    'txt_monto2.Text
        rs_datos!poa_monto_modif = 0        'txt_monto3.Text
        rs_datos!poa_monto_vigente_bs = 0   'txt_monto4.Text
     End If
     rs_datos!ges_gestion = txt_campo4.Text
     'rs_datos!poa_descripcion = Txt_descripcion.Text
     rs_datos!unidad_codigo = dtc_codigo1.Text
     rs_datos!fecha_registro = DTPfecha1.Value
     rs_datos!beneficiario_codigo = dtc_codigo4.Text
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
     rs_datos!poa_fuente_verificacion = txt_campo3.Text
     rs_datos!clasif_codigo = dtc_codigo8.Text
     rs_datos!doc_codigo = dtc_codigo9.Text
     'rs_datos!pro_codigo = IIf(dtc_codigo3.Text = "", 0, dtc_codigo3.Text)
     rs_datos!fte_codigo = IIf(dtc_aux3.Text = "", "10", dtc_aux3.Text)
     rs_datos!org_codigo = dtc_codigo3.Text
     rs_datos!par_codigo = dtc_codigo10.Text
     rs_datos!usr_codigo_aprueba = "0"
     rs_datos!fecha_aprueba = Date
     rs_datos!poa_num_respaldo = dtc_codigo8.Text + "-" + dtc_codigo9.Text
     'hora_registro
     rs_datos!fecha_registro = Date     'no cambia
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
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo4.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
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
  If (txt_campo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3a.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\poa\pr_poa_insumo.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
          'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
        'Call CREAVISTAF11          'JQA JUN-2008
'        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'        CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
'        CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
        CR01.WindowState = crptMaximized
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
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
        VAR_SW = "MOD"
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        Txt_descripcion.SetFocus
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

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

'Private Sub dtc_codigo12_Click(Area As Integer)
'    dtc_desc12.BoundText = dtc_codigo12.BoundText
'End Sub

'Private Sub dtc_codigo2_Click(Area As Integer)
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
'End Sub

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

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

'Private Sub dtc_desc12_Click(Area As Integer)
'    dtc_codigo12.BoundText = dtc_desc12.BoundText
'End Sub

Private Sub dtc_desc2_LostFocus()
    lbl_objgral1.Caption = Ado_datos.Recordset!estrategico_codigo   'dtc_codigo2.Text
End Sub

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
'    Call pnivel6(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
'End Sub
'
'Private Sub pnivel6(codigo6 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_actividad where especifico_codigo = '" & codigo6 & "'"
'
'   Set dtc_codigo6.RowSource = Nothing
'   Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo6.ReFill
'   dtc_codigo6.BoundText = Empty
'
'   Set dtc_desc6.RowSource = Nothing
'   Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc6.ReFill
'   dtc_desc6.BoundText = Empty
'End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc4.Enabled = True
    Call pnivel11(dtc_codigo1.BoundText)
    dtc_desc11.Enabled = True
End Sub
   
Private Sub pnivel1(codigo1 As String)
   Dim strConsultaF As String
   strConsultaF = "select * from rv_unidad_vs_responsable where unidad_codigo = '" & codigo1 & "'"

   Set dtc_codigo4.RowSource = Nothing
   Set dtc_codigo4.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo4.ReFill
   dtc_codigo4.BoundText = Empty

   Set dtc_desc4.RowSource = Nothing
   Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc4.ReFill
   dtc_desc4.BoundText = Empty
End Sub
  
Private Sub pnivel11(codigo1 As String)
   Dim strConsultaF As String
    strConsultaF = "select * from rv_unidad_vs_responsable where unidad_codigo = '" & codigo1 & "'"

   Set dtc_codigo11.RowSource = Nothing
   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo11.ReFill
   dtc_codigo11.BoundText = Empty

   Set dtc_desc11.RowSource = Nothing
   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc11.ReFill
   dtc_desc11.BoundText = Empty
End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

'Private Sub dtc_desc2_Click(Area As Integer)
'    dtc_codigo2.BoundText = dtc_desc2.BoundText
'    Call pnivel5(dtc_codigo2.BoundText)
'    dtc_desc5.Enabled = True
'End Sub

'Private Sub pnivel5(codigo5 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_especifico where estrategico_codigo = '" & codigo5 & "'"
'
'   Set dtc_codigo5.RowSource = Nothing
'   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo5.ReFill
'   dtc_codigo5.BoundText = Empty
'
'   Set dtc_desc5.RowSource = Nothing
'   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc5.ReFill
'   dtc_desc5.BoundText = Empty
'End Sub
 
Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_LostFocus()
    lbl_objgral2.Caption = Ado_datos.Recordset!estrategico_codigo
    lbl_objesp1.Caption = Ado_datos.Recordset!especifico_codigo
End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'    dtc_codigo6.BoundText = dtc_desc6.BoundText
'    Call pnivel7(dtc_codigo6.BoundText)
'    dtc_desc7.Enabled = True
'End Sub
'
'Private Sub pnivel7(codigo7 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_tarea where actividad_codigo = '" & codigo7 & "'"
'
'   Set dtc_codigo7.RowSource = Nothing
'   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo7.ReFill
'   dtc_codigo7.BoundText = Empty
'
'   Set dtc_desc7.RowSource = Nothing
'   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc7.ReFill
'   dtc_desc7.BoundText = Empty
'End Sub

Private Sub dtc_desc6_LostFocus()
    lbl_objgral3.Caption = Ado_datos.Recordset!estrategico_codigo
    lbl_objesp2.Caption = Ado_datos.Recordset!especifico_codigo
    lbl_act1.Caption = IIf(IsNull(Ado_datos.Recordset!actividad_codigo), dtc_codigo6.Text, Ado_datos.Recordset!actividad_codigo)
End Sub

'Private Sub dtc_desc7_Click(Area As Integer)
'    dtc_codigo7.BoundText = dtc_desc7.BoundText
'End Sub

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
'    parametro = AUX
    'parametro = "estado_codigo" + " = " + "'REG'"
    '
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'JQA 2014-JUL-14
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'pc_poa_estrategico
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from pc_poa_estrategico where estado_codigo= 'APR' order by poa_codigo", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'fc_organismo_financiamiento
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from fc_estructura_programatica WHERE pro_nivel <> '1' order by pro_descripcion", db, adOpenStatic
    rs_datos3.Open "Select * from fc_organismo_financiamiento order by org_descripcion", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "select * from rv_unidad_vs_responsable ORDER BY beneficiario_denominacion ", db, adOpenStatic
    'rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    'pc_poa_especifico
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from pc_poa_especifico where estado_codigo= 'APR' order by poa_codigo", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText

    'pc_poa_actividad
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from pc_poa_actividad where estado_codigo= 'APR' order by poa_codigo", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText

    'pc_poa_tarea
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from pc_poa_tarea where estado_codigo= 'APR' order by poa_codigo", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText

    'gc_documentos_clasificacion
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
    'gc_documentos_respaldo
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    'fc_partida_gasto
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from fc_partida_gasto order by par_descripcion", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
    'gc_beneficiario (Personal CGI)
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
    
    'pc_poa_insumo
    Set rs_datos12 = New ADODB.Recordset
    If rs_datos12.State = 1 Then rs_datos12.Close
    rs_datos12.Open "Select * from pc_poa_insumo where estado_codigo= 'APR' order by poa_codigo", db, adOpenStatic
    Set Ado_datos12.Recordset = rs_datos12
    dtc_desc12.BoundText = dtc_codigo12.BoundText
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
    If VAR_SW <> "ADD" Then
        Select Case Ado_datos.Recordset!poa_nivel
            Case "1"    'pc_poa_estrategico
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "Select * from pc_poa_estrategico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " ", db, adOpenStatic
                If rs_aux1.RecordCount > 0 Then
                    Txt_descripcion = rs_aux1!poa_descripcion
                    lbl_objgral1.Caption = rs_aux1!estrategico_codigo
                Else
                    Txt_descripcion = ""
                End If
                txt_aux5 = ""
                txt_aux6 = ""
                txt_aux7 = ""
                txt_aux12 = ""
                lbl_objgral2 = ""
                lbl_objgral3 = ""
                lbl_objgral4 = ""
                lbl_objgral5 = ""
                BtnAprobar.Visible = False
            Case "2"    'pc_poa_especifico
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "Select * from pc_poa_estrategico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " ", db, adOpenStatic
                If rs_aux1.RecordCount > 0 Then
                    Txt_descripcion = rs_aux1!poa_descripcion
                    lbl_objgral1.Caption = rs_aux1!estrategico_codigo
                Else
                    Txt_descripcion = ""
                End If
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                rs_aux2.Open "Select * from pc_poa_especifico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " ", db, adOpenStatic
                If rs_aux2.RecordCount > 0 Then
                    txt_aux5 = rs_aux2!poa_descripcion
                    lbl_objgral2.Caption = Str(rs_aux2!estrategico_codigo) + "." + Str(rs_aux2!especifico_codigo)
                End If
                txt_aux6 = ""
                txt_aux7 = ""
                txt_aux12 = ""
                lbl_objgral3 = ""
                lbl_objgral4 = ""
                lbl_objgral5 = ""
                BtnAprobar.Visible = False
            Case "3"    'pc_poa_actividad
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "Select * from pc_poa_estrategico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " ", db, adOpenStatic
                If rs_aux1.RecordCount > 0 Then
                    Txt_descripcion = rs_aux1!poa_descripcion
                    lbl_objgral1.Caption = rs_aux1!estrategico_codigo
                Else
                    Txt_descripcion = ""
                End If
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                rs_aux2.Open "Select * from pc_poa_especifico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " ", db, adOpenStatic
                If rs_aux2.RecordCount > 0 Then
                    txt_aux5 = rs_aux2!poa_descripcion
                    lbl_objgral2.Caption = Str(rs_aux2!estrategico_codigo) + "." + Str(rs_aux2!especifico_codigo)
                End If
                Set rs_aux3 = New ADODB.Recordset
                If rs_aux3.State = 1 Then rs_aux3.Close
                rs_aux3.Open "Select * from pc_poa_actividad where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " and actividad_codigo = " & Ado_datos.Recordset!actividad_codigo & " ", db, adOpenStatic
                If rs_aux3.RecordCount > 0 Then
                    txt_aux6 = rs_aux3!poa_descripcion
                    lbl_objgral3.Caption = Str(rs_aux3!estrategico_codigo) + "." + Str(rs_aux3!especifico_codigo) + "." + Str(rs_aux3!actividad_codigo)
                End If
                txt_aux7 = ""
                txt_aux12 = ""
                lbl_objgral4 = ""
                lbl_objgral5 = ""
                BtnAprobar.Visible = False
            Case "4"    'pc_poa_tarea
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "Select * from pc_poa_estrategico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " ", db, adOpenStatic
                If rs_aux1.RecordCount > 0 Then
                    Txt_descripcion = rs_aux1!poa_descripcion
                    lbl_objgral1.Caption = rs_aux1!estrategico_codigo
                Else
                    Txt_descripcion = ""
                End If
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                rs_aux2.Open "Select * from pc_poa_especifico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " ", db, adOpenStatic
                If rs_aux2.RecordCount > 0 Then
                    txt_aux5 = rs_aux2!poa_descripcion
                    lbl_objgral2.Caption = Str(rs_aux2!estrategico_codigo) + "." + Str(rs_aux2!especifico_codigo)
                End If
                Set rs_aux3 = New ADODB.Recordset
                If rs_aux3.State = 1 Then rs_aux3.Close
                rs_aux3.Open "Select * from pc_poa_actividad where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " and actividad_codigo = " & Ado_datos.Recordset!actividad_codigo & " ", db, adOpenStatic
                If rs_aux3.RecordCount > 0 Then
                    txt_aux6 = rs_aux3!poa_descripcion
                    lbl_objgral3.Caption = Str(rs_aux3!estrategico_codigo) + "." + Str(rs_aux3!especifico_codigo) + "." + Str(rs_aux3!actividad_codigo)
                End If
                Set rs_aux4 = New ADODB.Recordset
                If rs_aux4.State = 1 Then rs_aux4.Close
                rs_aux4.Open "Select * from pc_poa_tarea where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " and actividad_codigo = " & Ado_datos.Recordset!actividad_codigo & " and tarea_codigo = " & Ado_datos.Recordset!tarea_codigo & " ", db, adOpenStatic
                If rs_aux4.RecordCount > 0 Then
                    txt_aux7 = rs_aux4!poa_descripcion
                    lbl_objgral4.Caption = Str(rs_aux4!estrategico_codigo) + "." + Str(rs_aux4!especifico_codigo) + "." + Str(rs_aux4!actividad_codigo) + "." + Str(rs_aux4!tarea_codigo)
                End If
                txt_aux12 = ""
                lbl_objgral5 = ""
                BtnAprobar.Visible = False
            Case "5"    'pc_poa_insumo
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "Select * from pc_poa_estrategico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " ", db, adOpenStatic
                If rs_aux1.RecordCount > 0 Then
                    Txt_descripcion = rs_aux1!poa_descripcion
                    lbl_objgral1.Caption = rs_aux1!estrategico_codigo
                Else
                    Txt_descripcion = ""
                End If
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                rs_aux2.Open "Select * from pc_poa_especifico where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " ", db, adOpenStatic
                If rs_aux2.RecordCount > 0 Then
                    txt_aux5 = rs_aux2!poa_descripcion
                    lbl_objgral2.Caption = Str(rs_aux2!estrategico_codigo) + "." + Str(rs_aux2!especifico_codigo)
                End If
                Set rs_aux3 = New ADODB.Recordset
                If rs_aux3.State = 1 Then rs_aux3.Close
                rs_aux3.Open "Select * from pc_poa_actividad where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " and actividad_codigo = " & Ado_datos.Recordset!actividad_codigo & " ", db, adOpenStatic
                If rs_aux3.RecordCount > 0 Then
                    txt_aux6 = rs_aux3!poa_descripcion
                    lbl_objgral3.Caption = Str(rs_aux3!estrategico_codigo) + "." + Str(rs_aux3!especifico_codigo) + "." + Str(rs_aux3!actividad_codigo)
                End If
                Set rs_aux4 = New ADODB.Recordset
                If rs_aux4.State = 1 Then rs_aux4.Close
                rs_aux4.Open "Select * from pc_poa_tarea where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " and actividad_codigo = " & Ado_datos.Recordset!actividad_codigo & " and tarea_codigo = " & Ado_datos.Recordset!tarea_codigo & " ", db, adOpenStatic
                If rs_aux4.RecordCount > 0 Then
                    txt_aux7 = rs_aux4!poa_descripcion
                    lbl_objgral4.Caption = Str(rs_aux4!estrategico_codigo) + "." + Str(rs_aux4!especifico_codigo) + "." + Str(rs_aux4!actividad_codigo) + "." + Str(rs_aux4!tarea_codigo)
                End If
                Set rs_aux5 = New ADODB.Recordset
                If rs_aux5.State = 1 Then rs_aux5.Close
                rs_aux5.Open "Select * from pc_poa_insumo where estrategico_codigo = " & Ado_datos.Recordset!estrategico_codigo & " and especifico_codigo = " & Ado_datos.Recordset!especifico_codigo & " and actividad_codigo = " & Ado_datos.Recordset!actividad_codigo & " and tarea_codigo = " & Ado_datos.Recordset!tarea_codigo & " and insumo_codigo = " & Ado_datos.Recordset!insumo_codigo & " ", db, adOpenStatic
                If rs_aux5.RecordCount > 0 Then
                    txt_aux12 = rs_aux5!poa_descripcion
                    lbl_objgral5.Caption = Str(rs_aux5!estrategico_codigo) + "." + Str(rs_aux5!especifico_codigo) + "." + Str(rs_aux5!actividad_codigo) + "." + Str(rs_aux5!tarea_codigo) + "." + Str(rs_aux5!insumo_codigo)
                End If
                BtnAprobar.Visible = True
        End Select
'        lbl_objgral1.Caption = Ado_datos.Recordset!estrategico_codigo
'
'        lbl_objgral2.Caption = Ado_datos.Recordset!estrategico_codigo
'        lbl_objesp1.Caption = Ado_datos.Recordset!especifico_codigo
'
'        lbl_objgral3.Caption = Ado_datos.Recordset!estrategico_codigo
'        lbl_objesp2.Caption = Ado_datos.Recordset!especifico_codigo
'        lbl_act1.Caption = Ado_datos.Recordset!actividad_codigo
'
'        lbl_objgral4.Caption = Ado_datos.Recordset!estrategico_codigo
'        lbl_objesp3.Caption = Ado_datos.Recordset!especifico_codigo
'        lbl_act2.Caption = Ado_datos.Recordset!actividad_codigo
'        lbl_tar1.Caption = Ado_datos.Recordset!tarea_codigo
        
'        Txt_descripcion = dtc_desc2.Text    ' Ado_datos.Recordset!estrategico_codigo
'        txt_aux5 = dtc_desc5.Text    'Ado_datos.Recordset!especifico_codigo
'        txt_aux6 = dtc_desc6.Text    'Ado_datos.Recordset!actividad_codigo
'        txt_aux7 = dtc_desc7.Text    'Ado_datos.Recordset!tarea_codigo
'        txt_aux12 = dtc_desc12.Text    'Ado_datos.Recordset!insumo_codigo
       
    Else
        'Set rs_det1 = New ADODB.Recordset
'        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
'    FraDet1.Caption = FraDet1.Caption + dtc_aux1.Text
'    txt_aux9.Text = dtc_desc9.Text
    If Ado_datos.Recordset!estado_codigo = "APR" Then
            
    Else
            
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
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    'txt_codigo.Enabled = False
'    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
'    rs_datos.AddNew
    Ado_datos.Recordset.AddNew
    Txt_descripcion.SetFocus
    'dtc_desc1.BackColor = &H80000005
    dtc_desc4.Enabled = False
    dtc_desc11.Enabled = False
    lbl_objgral1.Caption = "0"
    lbl_objgral2.Caption = "0"
    lbl_objesp1.Caption = "0"
    lbl_objgral3.Caption = "0"
    lbl_objesp2.Caption = "0"
    lbl_act1.Caption = "0"
    
    lbl_objgral4.Caption = "0"
    lbl_objesp3.Caption = "0"
    lbl_act2.Caption = "0"
    lbl_tar1.Caption = "0"
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

Private Function ExisteReg(Obj1 As String, obj2 As String, activ As String, tarea2 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM pc_poa_insumo WHERE estrategico_codigo = " & Obj1 & " and especifico_codigo = " & obj2 & " and actividad_codigo = " & activ & " and tarea_codigo = " & tarea2 & " "
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from po_poa where estado_codigo = 'REG' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from po_poa "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub txt_campo4_LostFocus()
    If Val(txt_campo4.Text) > 2104 Or Val(txt_campo4.Text) < 1980 Then
        MsgBox "Debe Registrar una Gestión Válida !!. Vuelva a Intentar. ", vbExclamation, "Atención!"
        txt_campo4.SetFocus
    End If
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

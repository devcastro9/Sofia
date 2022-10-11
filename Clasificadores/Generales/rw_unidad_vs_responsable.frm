VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rw_unidad_vs_responsable 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Clasificadores - Unidad Responsable"
   ClientHeight    =   7770
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14880
   Icon            =   "rw_unidad_vs_responsable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00E0E0E0&
      Height          =   5775
      Left            =   7680
      TabIndex        =   22
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox txt_cod_estado 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
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
         Height          =   315
         Left            =   600
         TabIndex        =   24
         Top             =   3720
         Width           =   1245
      End
      Begin VB.CheckBox ckEsresponsable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "    Es responsable de unidad"
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
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   4800
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "rw_unidad_vs_responsable.frx":0A02
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   600
         TabIndex        =   25
         Top             =   2160
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "nombres"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "rw_unidad_vs_responsable.frx":0A1B
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   600
         TabIndex        =   26
         Top             =   1080
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "rw_unidad_vs_responsable.frx":0A34
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   29
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "rw_unidad_vs_responsable.frx":0A4D
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   31
         Top             =   1800
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblesresponsable 
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
         Left            =   600
         TabIndex        =   32
         Top             =   3360
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
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
         Left            =   600
         TabIndex        =   28
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable"
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
         Left            =   600
         TabIndex        =   27
         Top             =   1800
         Width           =   1230
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      Height          =   1020
      Left            =   120
      ScaleHeight     =   960
      ScaleWidth      =   14280
      TabIndex        =   16
      Top             =   120
      Width           =   14340
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H80000015&
         Caption         =   "Anular"
         Height          =   720
         Left            =   2280
         Picture         =   "rw_unidad_vs_responsable.frx":0A66
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H80000015&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   3360
         Picture         =   "rw_unidad_vs_responsable.frx":1730
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H80000015&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   3360
         Picture         =   "rw_unidad_vs_responsable.frx":193A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H80000015&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   4440
         Picture         =   "rw_unidad_vs_responsable.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H80000015&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   6120
         Picture         =   "rw_unidad_vs_responsable.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H80000015&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   13080
         Picture         =   "rw_unidad_vs_responsable.frx":26B9
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H80000015&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   1200
         Picture         =   "rw_unidad_vs_responsable.frx":28C3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H80000015&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "rw_unidad_vs_responsable.frx":2EA3
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   1125
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "rw_unidad_vs_responsable.frx":34C7
         DataField       =   "depto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "depto_codigo"
         BoundColumn     =   "depto_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   330
         Left            =   8745
         TabIndex        =   19
         Top             =   300
         Width           =   1230
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00800000&
      Height          =   5775
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   7455
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "rw_unidad_vs_responsable.frx":34E0
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   8705
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "correl_ur"
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
         BeginProperty Column02 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Responsable"
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
            DataField       =   "esresponsable"
            Caption         =   "Responsable Estado"
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
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5280
         Width           =   7185
         _ExtentX        =   12674
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
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   14880
      TabIndex        =   10
      Top             =   7770
      Width           =   14880
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   15
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2400
      Top             =   6960
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
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   120
      ScaleHeight     =   960
      ScaleWidth      =   12600
      TabIndex        =   17
      Top             =   120
      Width           =   12660
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "rw_unidad_vs_responsable.frx":34F8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   1560
         Picture         =   "rw_unidad_vs_responsable.frx":3DE4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1245
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
         Left            =   8595
         TabIndex        =   18
         Top             =   300
         Width           =   1290
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
End
Attribute VB_Name = "rw_unidad_vs_responsable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
'BUSCADOR
'Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean


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
'    Set ClBuscaGrid = New ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = db
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = dg_datos
'    ClBuscaGrid.QueryUtilizado = queryinicial
'    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
'    ClBuscaGrid.Ejecutar
    
  PosibleApliqueFiltro = False
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = dg_datos
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
   
        If VAR_SW = "ADD" Then
          rs_datos.Delete
        Else
          rs_datos.CancelUpdate
        End If
        

        Call ABRIR_TABLA
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
       ' txt_codigo_estado.Enabled = True
        dtc_desc1.Enabled = True
    End If
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   'If ExisteReg(Ado_datos.Recordset!calle_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   'If ExisteReg2(Ado_datos.Recordset!calle_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
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
  
  If VAR_SW = "ADD" Then
          Call ValidarExisteRegistro
  End If
  
  If VAR_VAL = "OK" Then
     
     rs_datos!unidad_codigo = dtc_codigo4
     rs_datos!beneficiario_codigo = dtc_codigo7
     rs_datos!estado_codigo = "REG"  ' no cambia
     If ckEsresponsable.Value = 1 Then
       rs_datos!estado_codigo_resp = "APR"
     Else
       rs_datos!estado_codigo_resp = "REG"
     End If

     rs_datos!fecha_registro = Date     ' no cambia
     rs_datos!usr_codigo = glusuario    ' no cambia
     rs_datos.UpdateBatch adAffectAll
    
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
     
     Ado_datos.Recordset.Move marca1 - 1
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()

  If dtc_codigo4 = "" Then
    MsgBox "Debe seleccionar una Unidad"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo7 = "" Then
    MsgBox "Debe seleccionar un responsable"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
End Sub

Private Sub ValidarExisteRegistro()
    If VAR_VAL <> "ERR" Then
    Dim cuantosReg As Integer
    Dim rsCant As New ADODB.Recordset
    rsCant.Open "SELECT COUNT(*) As cuantos FROM rc_unidad_vs_responsable WHERE unidad_codigo = '" + dtc_codigo4 + "' AND beneficiario_codigo = '" + dtc_codigo7 + "' ", db, adOpenStatic
    rsCant.MoveFirst
    cuantosReg = rsCant![Cuantos]
    
    If cuantosReg > 0 Then
        MsgBox " REGISTRO EXISTENTE UNI:" + dtc_desc4 + " RESP:" + dtc_desc5
        VAR_VAL = "ERR"
    End If
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  cr01.WindowShowPrintSetupBtn = True
  cr01.WindowShowRefreshBtn = True
  cr01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_provincias.rpt"
  iResult = cr01.PrintReport
  If iResult <> 0 Then
      MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  cr01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If rs_datos!estado_codigo = "REG" Then

    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
   
  Else
      MsgBox "No se puede MODIFICAR un registro Aprobado(APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()

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

'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo7.BoundText
'End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
   dtc_codigo7.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    'Call pnivel2(dtc_codigo4.BoundText)
    'dtc_desc5.Enabled = True
End Sub


Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc5.BoundText
   ' Call pnivel3(dtc_codigo7.BoundText)
   ' dtc_desc6.Enabled = True
End Sub
   
Private Sub pnivel3(codigo5 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo5 & "'"
'   Set dtc_desc4.RowSource = Nothing
'   Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc4.ReFill
'   dtc_desc4.BoundText = Empty
'
'   Set dtc_desc5.RowSource = Nothing
'   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc5.ReFill
'   dtc_desc5.BoundText = Empty
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
    'ckEsresponsable.Value = 0
    txt_cod_estado.Text = "REG"
   
    
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = " SELECT *, (CASE estado_codigo_resp WHEN 'APR' THEN 'SI' ELSE 'NO' END) AS esresponsable FROM rc_unidad_vs_responsable ORDER BY unidad_codigo, beneficiario_codigo ASC"   'correl_ur
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
  
End Sub

Private Sub ABRIR_TABLAS_AUX()
    ' Unidad
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "SELECT * FROM gc_unidad_ejecutora WHERE estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    ' Responsable
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open " SELECT (beneficiario_primer_apellido + ' ' + beneficiario_segundo_apellido + ' ' + beneficiario_nombres) AS nombres, * FROM gc_beneficiario WHERE tipoben_codigo = 1 ORDER BY beneficiario_primer_apellido, beneficiario_segundo_apellido, beneficiario_nombres", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo7.BoundText
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
      Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
      
       If Ado_datos.Recordset!estado_codigo_resp = "APR" Then
           ckEsresponsable.Value = 1
        Else
           ckEsresponsable.Value = 0
        End If
        
        If Ado_datos.Recordset!estado_codigo <> Nulo Then
            txt_cod_estado.Text = Ado_datos.Recordset!estado_codigo
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
    Call ABRIR_TABLA
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
    ckEsresponsable.Value = 0
    txt_cod_estado.Text = "REG"
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

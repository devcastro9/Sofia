VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_rc_planilla_grupo 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Grupo Planilla"
   ClientHeight    =   9255
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   16
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7080
         Picture         =   "frm_rc_planilla_grupo.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   29
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   0
         Width           =   1400
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   9840
         Picture         =   "frm_rc_planilla_grupo.frx":08CD
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   120
         Picture         =   "frm_rc_planilla_grupo.frx":0AD7
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   22
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1545
         Picture         =   "frm_rc_planilla_grupo.frx":1296
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   21
         ToolTipText     =   "Modifica Datos Cabecera"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3000
         Picture         =   "frm_rc_planilla_grupo.frx":1BAB
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   20
         ToolTipText     =   "Anular Cronograma"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4320
         Picture         =   "frm_rc_planilla_grupo.frx":22F7
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   19
         ToolTipText     =   "Aprueba Cronograma"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5760
         Picture         =   "frm_rc_planilla_grupo.frx":2B2A
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   18
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17760
         Picture         =   "frm_rc_planilla_grupo.frx":32DF
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   17
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
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
         Left            =   12255
         TabIndex        =   24
         Top             =   200
         Width           =   1815
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
      ForeColor       =   &H00FFFF00&
      Height          =   6015
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   6255
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   5175
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "planilla_codigo"
            Caption         =   "Código"
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
            DataField       =   "planilla_descripcion"
            Caption         =   "Descripción"
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
            DataField       =   "depto_codigo"
            Caption         =   "Depto Codigo"
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
            DataField       =   "registro_patronal"
            Caption         =   "Reg Patronal"
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
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3075.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2954.835
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1544.882
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
         Caption         =   " <-- Inicio                        Grupo Planilla                          Fin -->"
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
      ForeColor       =   &H00FFFF80&
      Height          =   6015
      Left            =   6480
      TabIndex        =   10
      Top             =   720
      Width           =   5775
      Begin VB.TextBox txt_reg_patronal 
         Appearance      =   0  'Flat
         DataField       =   "registro_patronal"
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
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Text            =   "-"
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Txt_planilla_descripcion 
         DataField       =   "planilla_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   525
         Left            =   360
         TabIndex        =   2
         Text            =   "-"
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txt_planilla_codigo 
         Appearance      =   0  'Flat
         DataField       =   "planilla_codigo"
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Text            =   "-"
         Top             =   720
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dtc_descripcion 
         Bindings        =   "frm_rc_planilla_grupo.frx":3AA1
         DataField       =   "depto_codigo"
         Height          =   315
         Left            =   1200
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "depto_descripcion"
         BoundColumn     =   "depto_codigo"
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
      Begin MSDataListLib.DataCombo dtc_codigo 
         Bindings        =   "frm_rc_planilla_grupo.frx":3ABA
         DataField       =   "depto_codigo"
         Height          =   315
         Left            =   360
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "depto_codigo"
         BoundColumn     =   "depto_codigo"
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
      Begin VB.TextBox txt_depto_codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DataField       =   "depto_codigo"
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
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "-"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txt_depto_descripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "-"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Registro Patronal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   3120
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Departamento Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   2280
         Width           =   1830
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Descripción Planilla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   1275
         Width           =   1680
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código Planilla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   420
         Width           =   1260
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   3960
         Width           =   1365
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
      TabIndex        =   4
      Top             =   9255
      Width           =   11280
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   9
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   7560
      Top             =   7320
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
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4275
         Picture         =   "frm_rc_planilla_grupo.frx":3AD3
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   27
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "frm_rc_planilla_grupo.frx":43BF
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   26
         Top             =   0
         Width           =   1300
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
         Left            =   14175
         TabIndex        =   28
         Top             =   195
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   240
      Top             =   7680
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
Attribute VB_Name = "frm_rc_planilla_grupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset

'BUSCADOR

Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

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
        'txt_codigo.Enabled = True
        dtc_desc1.Enabled = True
        
        dtc_codigo.Visible = False
        dtc_descripcion.Visible = False
        txt_depto_codigo.Visible = True
        txt_depto_descripcion.Visible = True
      
        
    End If
End Sub

Private Sub BtnEliminar_Click()
    On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!planilla_codigo) Then
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
         db.Execute "DELETE rc_planilla_grupo where planilla_codigo = '" & Ado_datos.Recordset!planilla_codigo & "'"
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
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "select * from rc_planilla_grupo where planilla_codigo = '" & txt_planilla_codigo.Text & "'  "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            MsgBox " CODIGO DUPLICADO, Vuelva a intentar..."
            Exit Sub
        End If
        rs_datos!planilla_codigo = txt_planilla_codigo.Text
        rs_datos!planilla_descripcion = Txt_planilla_descripcion.Text
        rs_datos!depto_codigo = dtc_codigo.Text
        rs_datos!registro_patronal = txt_reg_patronal.Text
        
        rs_datos!estado_codigo = "REG"  ' no cambia
        rs_datos!fecha_registro = Date     ' no cambia
        rs_datos!usr_codigo = glusuario    ' no cambia
        rs_datos.UpdateBatch adAffectAll
       
     End If
    
     'rs_datos!planilla_codigo = txt_planilla_codigo.Text
     rs_datos!planilla_descripcion = Txt_planilla_descripcion.Text
     rs_datos!registro_patronal = txt_reg_patronal.Text
     rs_datos!estado_codigo = "REG"  ' no cambia
     rs_datos!fecha_registro = Date     ' no cambia
     rs_datos!usr_codigo = glusuario    ' no cambia
     rs_datos.UpdateBatch adAffectAll
     
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
      
      
      dtc_codigo.Visible = False
    dtc_descripcion.Visible = False
    txt_depto_codigo.Visible = True
    txt_depto_descripcion.Visible = True
      
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      txt_planilla_codigo.Enabled = True
      'dtc_desc1.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'  If txt_codigo.Text = "" Then
'    MsgBox "Debe registrar el " + lbl_codigo.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If Txt_planilla_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  CR01.WindowShowPrintSetupBtn = True
  CR01.WindowShowRefreshBtn = True
  CR01.ReportFileName = App.Path & "\REPORTES\clasificadores\rr_planilla_grupo.rpt"
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
   On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
   txt_planilla_codigo.Locked = True
    dtc_codigo.Visible = False
    dtc_descripcion.Visible = False
    txt_depto_codigo.Visible = True
    txt_depto_descripcion.Visible = True
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
    txt_planilla_codigo.Enabled = False
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

Private Sub dtc_codigo_Click(Area As Integer)
dtc_descripcion.BoundText = dtc_codigo.BoundText
Txt_planilla_descripcion.Text = "PLANILLA " & dtc_descripcion.Text
End Sub

Private Sub dtc_descripcion_Click(Area As Integer)
 dtc_codigo.BoundText = dtc_descripcion.BoundText
 Txt_planilla_descripcion.Text = "PLANILLA " & dtc_descripcion.Text
End Sub

Private Sub Form_Load()
   
    Call ABRIR_TABLA
    txt_planilla_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
End Sub
Private Sub ABRIR_TABLAS_AUX()
'    gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos4.Close
    rs_datos1.Open "SELECT * FROM gc_departamento", db, adOpenStatic
    
    Set Ado_datos1.Recordset = rs_datos1
    dtc_descripcion.BoundText = dtc_codigo.BoundText
    
        
    
'    Set rs_datos5 = New ADODB.Recordset
'    If rs_datos5.State = 1 Then rs_datos5.Close
'    rs_datos5.Open "av_ro_peronal_vs_gc_beneficiario where unidad_codigo = '" & rs_datos1!unidad_codigo_pla & "' order by beneficiario_denominacion", db, adOpenStatic
'    Set Ado_datos5.Recordset = rs_datos5
'    dtc_descripcion.BoundText = dtc_codigo.BoundText
'
'    'Beneficiario Funcionario CGI (Vendedor, Cobrador, Adm, etc.)
'    Set rs_datos4 = New ADODB.Recordset
'    If rs_datos4.State = 1 Then rs_datos4.Close
'    'rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = 'DNMAN' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'    Set Ado_datos4.Recordset = rs_datos4
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'
End Sub
Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select * from rc_planilla_grupo"
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
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
     Call ABRIR_TABLAS_AUX
    dtc_codigo.Visible = True
    dtc_descripcion.Visible = True
    txt_depto_codigo.Visible = False
    txt_depto_descripcion.Visible = False
     txt_planilla_codigo.Locked = False
    'txt_planilla_codigo.Enabled = True
    txt_planilla_codigo.SetFocus
    'Txt_descripcion.SetFocus
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

Private Function ExisteReg(planilla_codigo As String) As Boolean
        Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM rc_planilla_sub_grupo WHERE planilla_codigo = '" & txt_planilla_codigo & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
    
End Function

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_depto_codigo_Change()
 Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "SELECT * FROM gc_departamento where depto_codigo = '" & txt_depto_codigo & "'", db, adOpenStatic
    If rs_aux2.RecordCount > 0 Then
      txt_depto_descripcion.Text = rs_aux2!depto_descripcion
      End If
End Sub

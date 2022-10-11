VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_co_plan_cuentas 
   BackColor       =   &H00000000&
   Caption         =   "Contabilidad - Plan de Cuentas"
   ClientHeight    =   9165
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   62
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "frm_co_plan_cuentas.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   71
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
         Picture         =   "frm_co_plan_cuentas.frx":07C2
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   70
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
         Picture         =   "frm_co_plan_cuentas.frx":108F
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   69
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
         Picture         =   "frm_co_plan_cuentas.frx":1844
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   68
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "frm_co_plan_cuentas.frx":2077
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1305
         Picture         =   "frm_co_plan_cuentas.frx":27C3
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "frm_co_plan_cuentas.frx":30D8
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   8640
         Picture         =   "frm_co_plan_cuentas.frx":3897
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   9600
         Picture         =   "frm_co_plan_cuentas.frx":3CD9
         Style           =   1  'Graphical
         TabIndex        =   63
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
         TabIndex        =   72
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
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "frm_co_plan_cuentas.frx":3EE3
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   60
         Top             =   0
         Width           =   1300
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4275
         Picture         =   "frm_co_plan_cuentas.frx":46B9
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   59
         Top             =   0
         Width           =   1400
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
         TabIndex        =   61
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   3360
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   14652
      Begin VB.OptionButton OptFilGral4 
         BackColor       =   &H80000018&
         Caption         =   "Cuentas de Detalle"
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
         Left            =   10920
         TabIndex        =   39
         Top             =   2970
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton OptFilGral3 
         BackColor       =   &H80000018&
         Caption         =   "Cuentas de Sub Título"
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
         Left            =   7560
         TabIndex        =   38
         Top             =   2970
         Width           =   2175
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H80000018&
         Caption         =   "Cuentas de Título"
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
         Left            =   4200
         TabIndex        =   37
         Top             =   2970
         Width           =   1815
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H80000018&
         Caption         =   "Todas"
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
         Left            =   1800
         TabIndex        =   36
         Top             =   2970
         Width           =   915
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "frm_co_plan_cuentas.frx":4FA5
         Height          =   2532
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "Cuenta"
            Caption         =   "Cuenta"
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
            DataField       =   "SubCta1"
            Caption         =   "SubCta1"
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
            DataField       =   "SubCta2"
            Caption         =   "SubCta2"
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
            DataField       =   "NombreCta"
            Caption         =   "Nombre de la Cuenta"
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
            DataField       =   "Aux1"
            Caption         =   "Aux1"
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
            DataField       =   "Aux2"
            Caption         =   "Aux2"
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
            DataField       =   "Aux3"
            Caption         =   "Aux3"
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
            DataField       =   "Mov"
            Caption         =   "Titulo/Mov"
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
            DataField       =   "nivel"
            Caption         =   "Nivel"
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   6630.236
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   336
         Left            =   120
         Top             =   2880
         Width           =   14268
         _ExtentX        =   25162
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
         BackColor       =   -2147483624
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
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00000000&
      Height          =   4200
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   14652
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_co_plan_cuentas.frx":4FBD
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4560
         TabIndex        =   2
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreCta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "frm_co_plan_cuentas.frx":4FD6
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   27
         Top             =   600
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SubCta2"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "frm_co_plan_cuentas.frx":4FEF
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2640
         TabIndex        =   26
         Top             =   600
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SubCta1"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_co_plan_cuentas.frx":5008
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   600
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Cuenta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "frm_co_plan_cuentas.frx":5021
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   1080
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Cuenta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "frm_co_plan_cuentas.frx":503A
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2640
         TabIndex        =   28
         Top             =   1080
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SubCta1"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "frm_co_plan_cuentas.frx":5053
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   29
         Top             =   1080
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SubCta2"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_co_plan_cuentas.frx":506C
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4560
         TabIndex        =   23
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreCta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin VB.TextBox txt_desc2 
         DataSource      =   "Ado_datos"
         Height          =   288
         Left            =   4560
         TabIndex        =   57
         Text            =   "-"
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox txt_desc1 
         DataSource      =   "Ado_datos"
         Height          =   288
         Left            =   4560
         TabIndex        =   56
         Text            =   "-"
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox txt_Tcta2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1680
         TabIndex        =   55
         Text            =   "-"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt_Tscta22 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3600
         TabIndex        =   54
         Text            =   "-"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt_Tscta12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2640
         TabIndex        =   53
         Text            =   "-"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt_Tscta1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2640
         TabIndex        =   52
         Text            =   "-"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txt_Tscta2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3600
         TabIndex        =   51
         Text            =   "-"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txt_Tcta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1680
         TabIndex        =   50
         Text            =   "-"
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame Fra_Aux 
         BackColor       =   &H00000000&
         Caption         =   "AUXILIARES"
         ForeColor       =   &H00FFFF80&
         Height          =   1695
         Left            =   600
         TabIndex        =   40
         Top             =   2160
         Width           =   11175
         Begin VB.CheckBox Chkaux1 
            BackColor       =   &H00000000&
            Caption         =   "Auxiliar 1"
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
            TabIndex        =   43
            Top             =   260
            Width           =   1092
         End
         Begin VB.CheckBox Chkaux2 
            BackColor       =   &H00000000&
            Caption         =   "Auxiliar 2"
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
            TabIndex        =   42
            Top             =   740
            Width           =   1092
         End
         Begin VB.CheckBox Chkaux3 
            BackColor       =   &H00000000&
            Caption         =   "Auxiliar 3"
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
            TabIndex        =   41
            Top             =   1220
            Width           =   1092
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "frm_co_plan_cuentas.frx":5085
            DataField       =   "Aux1"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   44
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "aux"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "frm_co_plan_cuentas.frx":509E
            DataField       =   "Aux2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   45
            Top             =   720
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "aux"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "frm_co_plan_cuentas.frx":50B7
            DataField       =   "Aux3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   46
            Top             =   1200
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "aux"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "frm_co_plan_cuentas.frx":50D0
            DataField       =   "Aux1"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   3360
            TabIndex        =   47
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripcion"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "frm_co_plan_cuentas.frx":50E9
            DataField       =   "Aux2"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   3360
            TabIndex        =   48
            Top             =   720
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripcion"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "frm_co_plan_cuentas.frx":5102
            DataField       =   "Aux3"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   3360
            TabIndex        =   49
            Top             =   1200
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripcion"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "nivel"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   13080
         TabIndex        =   25
         Text            =   "-"
         Top             =   3120
         Width           =   852
      End
      Begin VB.TextBox txt_parametro_menor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Mov"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   13080
         TabIndex        =   3
         Text            =   "-"
         Top             =   1920
         Width           =   852
      End
      Begin VB.TextBox Txt_descripcion 
         DataField       =   "NombreCta"
         DataSource      =   "Ado_datos"
         Height          =   288
         Left            =   2640
         TabIndex        =   0
         Text            =   "-"
         Top             =   2040
         Visible         =   0   'False
         Width           =   7815
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_co_plan_cuentas.frx":511B
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   1560
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Cuenta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_co_plan_cuentas.frx":5134
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4560
         TabIndex        =   24
         Top             =   1560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreCta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "frm_co_plan_cuentas.frx":514D
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2640
         TabIndex        =   30
         Top             =   1560
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SubCta1"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo12 
         Bindings        =   "frm_co_plan_cuentas.frx":5166
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   31
         Top             =   1560
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SubCta2"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Detalle"
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
         Left            =   480
         TabIndex        =   35
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Sub Título"
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
         Left            =   480
         TabIndex        =   34
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Título"
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
         Left            =   480
         TabIndex        =   33
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nombre de la Cuenta"
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
         Left            =   4560
         TabIndex        =   32
         Top             =   240
         Width           =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   12360
         X2              =   12360
         Y1              =   120
         Y2              =   4200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "SubCta2"
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
         TabIndex        =   20
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl_nro_dias_calendario 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nivel"
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
         Left            =   13120
         TabIndex        =   19
         Top             =   2760
         Width           =   750
      End
      Begin VB.Label lbl_observacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Titulo/Movimiento"
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
         Left            =   12640
         TabIndex        =   18
         Top             =   1560
         Width           =   1710
      End
      Begin VB.Label lbl_parametro_menor 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "SubCta1"
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
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl_enlace1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cuenta"
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
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nombre de la Cuenta"
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
         Left            =   480
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   2040
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
         Left            =   12960
         TabIndex        =   1
         Top             =   720
         Width           =   1095
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
         Left            =   12760
         TabIndex        =   11
         Top             =   360
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
      ScaleWidth      =   15120
      TabIndex        =   4
      Top             =   9165
      Width           =   15120
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
      Left            =   2520
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
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
      Left            =   6600
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
      Left            =   8760
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
      Left            =   10920
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
Attribute VB_Name = "frm_co_plan_cuentas"
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
Dim queryinicial As String

Dim var_cod, VAR_COD2, VAR_COD3 As String
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
        Call OptFilGral2_Click
        
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
   If ExisteReg(Ado_datos.Recordset!subproceso_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
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
        
        rs_datos!subproceso_codigo = txt_codigo.Text ' Esto para codigos trascritos
        rs_datos!estado_codigo = "REG"  ' no cambia
        rs_datos!correl_etapa = 0
        rs_datos!proceso_codigo = dtc_codigo1.Text   'Codigo del padre
        'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
'        db.Execute "Update gc_direccion_general Set correl_da = CAST('" & var_cod & "' AS INT) + 1 Where dgral_codigo= '" & dtc_codigo1.Text & "' "
     End If
     rs_datos!subproceso_descripcion = Txt_descripcion.Text
     rs_datos!fecha_registro = Date     ' no cambia
     rs_datos!usr_codigo = glusuario    ' no cambia
     rs_datos.UpdateBatch adAffectAll
    
     Call OptFilGral2_Click
     rs_datos.MoveLast
     mbDataChanged = False
      
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      txt_codigo.Enabled = True
      dtc_desc1.Enabled = True
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
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  cr01.WindowShowPrintSetupBtn = True
  cr01.WindowShowRefreshBtn = True
  cr01.ReportFileName = App.Path & "\REPORTES\contabilidad\cr_plan_cuentas.rpt"
  iResult = cr01.PrintReport
  If iResult <> 0 Then
      MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  cr01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
    dtc_codigo1.Enabled = False
    dtc_desc1.Enabled = False
    dtc_codigo2.Enabled = False
    dtc_desc2.Enabled = False
    dtc_codigo3.Enabled = False
    dtc_desc3.Enabled = False
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

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_codigo7.BoundText = dtc_codigo1.BoundText
    dtc_codigo10.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo10.BoundText
    dtc_codigo7.BoundText = dtc_codigo10.BoundText
    dtc_codigo1.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo11.BoundText
    dtc_codigo8.BoundText = dtc_codigo11.BoundText
    dtc_codigo2.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo12_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo12.BoundText
    dtc_codigo9.BoundText = dtc_codigo12.BoundText
    dtc_codigo3.BoundText = dtc_codigo12.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    dtc_codigo8.BoundText = dtc_codigo2.BoundText
    dtc_codigo11.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_codigo9.BoundText = dtc_codigo3.BoundText
    dtc_codigo12.BoundText = dtc_codigo3.BoundText
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
    dtc_desc1.BoundText = dtc_codigo7.BoundText
    dtc_codigo1.BoundText = dtc_codigo7.BoundText
    dtc_codigo10.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo8.BoundText
    dtc_codigo2.BoundText = dtc_codigo8.BoundText
    dtc_codigo11.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo9.BoundText
    dtc_codigo3.BoundText = dtc_codigo9.BoundText
    dtc_codigo12.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_codigo7.BoundText = dtc_desc1.BoundText
    dtc_codigo10.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    dtc_codigo8.BoundText = dtc_desc2.BoundText
    dtc_codigo11.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_codigo9.BoundText = dtc_desc3.BoundText
    dtc_codigo12.BoundText = dtc_desc3.BoundText
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

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click
    'Call ABRIR_TABLA
'    txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
    txt_Tcta.Visible = False
    txt_Tscta1.Visible = False
    txt_Tscta2.Visible = False
    txt_desc1.Visible = False
    
    txt_Tcta2.Visible = False
    txt_Tscta12.Visible = False
    txt_Tscta22.Visible = False
    txt_desc2.Visible = False
End Sub

Private Sub OptFilGral2_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from CC_Plan_Cuentas "
  'queryinicial = "select  da_codigo, da_descripcion, dgral_codigo, proceso_codigo, estado_codigo, fecha_registro, usr_codigo, correl_unidad as correl from gc_direccion_administrativa  "
  'queryinicial = "gp_listar_gc_direccion_general "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral1_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from CC_Plan_Cuentas where mov= 'T' "
  'queryinicial = "select  da_codigo, da_descripcion, dgral_codigo, proceso_codigo, estado_codigo, fecha_registro, usr_codigo, correl_unidad as correl from gc_direccion_administrativa  "
  'queryinicial = "gp_listar_gc_direccion_general "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral3_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from CC_Plan_Cuentas  where mov= 'S' "
  'queryinicial = "select  da_codigo, da_descripcion, dgral_codigo, proceso_codigo, estado_codigo, fecha_registro, usr_codigo, correl_unidad as correl from gc_direccion_administrativa  "
  'queryinicial = "gp_listar_gc_direccion_general "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral4_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from CC_Plan_Cuentas  where mov= 'D' "
  'queryinicial = "select  da_codigo, da_descripcion, dgral_codigo, proceso_codigo, estado_codigo, fecha_registro, usr_codigo, correl_unidad as correl from gc_direccion_administrativa  "
  'queryinicial = "gp_listar_gc_direccion_general "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 = '00' AND SubCta2 = '00' order by Cuenta ", db, adOpenStatic
    rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'T' order by Cuenta ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 = '00' order by Cuenta ", db, adOpenStatic
    rs_datos2.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'S' order by Cuenta ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 <> '00' order by Cuenta ", db, adOpenStatic
    rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'D' order by Cuenta ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
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
  If Ado_datos.Recordset.RecordCount > 0 Then
     ' Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
     If Ado_datos.Recordset!mov = "T" Then
        lbl_observacion.Caption = "TITULO"
        Fra_Aux.Visible = False
        dtc_codigo1.Visible = True
        dtc_codigo7.Visible = True
        dtc_codigo10.Visible = True
        dtc_desc1.Visible = True
        
        dtc_codigo2.Visible = False
        dtc_codigo8.Visible = False
        dtc_codigo11.Visible = False
        dtc_desc2.Visible = False
        
        dtc_codigo3.Visible = False
        dtc_codigo9.Visible = False
        dtc_codigo12.Visible = False
        dtc_desc3.Visible = False
        
        txt_Tcta.Visible = False
        txt_Tscta1.Visible = False
        txt_Tscta2.Visible = False
        txt_desc1.Visible = False
        
        txt_Tcta2.Visible = False
        txt_Tscta12.Visible = False
        txt_Tscta22.Visible = False
        txt_desc2.Visible = False
     Else
        If Ado_datos.Recordset!mov = "S" Then
           Fra_Aux.Visible = False
           lbl_observacion.Caption = "SUB TITULO"
           dtc_codigo1.Visible = False
            dtc_codigo7.Visible = False
            dtc_codigo10.Visible = False
            dtc_desc1.Visible = False
            
            dtc_codigo2.Visible = True
            dtc_codigo8.Visible = True
            dtc_codigo11.Visible = True
            dtc_desc2.Visible = True
            
            dtc_codigo3.Visible = False
            dtc_codigo9.Visible = False
            dtc_codigo12.Visible = False
            dtc_desc3.Visible = False
            
            var_cod = Ado_datos.Recordset!cuenta
            txt_Tcta.Visible = True
            txt_Tscta1.Visible = True
            txt_Tscta2.Visible = True
            txt_desc1.Visible = True
            
'            If txt_Tcta.Text <> "" Then
              If rs_aux2.State = 1 Then rs_aux2.Close
              'rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & txt_Tcta & "' and SubCta1 = '" & txt_Tscta1 & "' and SubCta2 = '" & txt_Tscta2 & "' ", db, adOpenKeyset, adLockOptimistic
              'rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & dtc_codigo1 & "' and SubCta1 = '00' and SubCta2 = '00' ", db, adOpenKeyset, adLockOptimistic
              rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & var_cod & "'  ", db, adOpenKeyset, adLockOptimistic
              If rs_aux2.RecordCount > 0 Then
                txt_Tcta.Text = rs_aux2("Cuenta")
                txt_Tscta1.Text = rs_aux2("SubCta1")
                txt_Tscta2.Text = rs_aux2("SubCta2")
                txt_desc1.Text = rs_aux2("NombreCta")
              End If
'            Else
'                txt_desc1.Text = dtc_desc1.Text
'            End If
        
                
        '    If DtCCta_codigo.Text <> "01" Then
        '      If rstdestino.State = 1 Then rstdestino.Close
        '      rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
        '      If Not rstFc_cuenta_bancaria.EOF Then
        '        fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
        '      Else
        '      End If
        '    Else
        '        fte_codigo1 = Me.DtCFte_codigo.Text
        '    End If
        Else
           Fra_Aux.Visible = True
           lbl_observacion.Caption = "DETALLE"
           dtc_codigo1.Visible = False
            dtc_codigo7.Visible = False
            dtc_codigo10.Visible = False
            dtc_desc1.Visible = False
            
            dtc_codigo2.Visible = False
            dtc_codigo8.Visible = False
            dtc_codigo11.Visible = False
            dtc_desc2.Visible = False
            
            dtc_codigo3.Visible = True
            dtc_codigo9.Visible = True
            dtc_codigo12.Visible = True
            dtc_desc3.Visible = True
            
            txt_Tcta.Visible = True
            txt_Tscta1.Visible = True
            txt_Tscta2.Visible = True
            txt_desc1.Visible = True
            
            txt_Tcta2.Visible = True
            txt_Tscta12.Visible = True
            txt_Tscta22.Visible = True
            txt_desc2.Visible = True
            
            var_cod = Ado_datos.Recordset!cuenta
            VAR_COD2 = Ado_datos.Recordset!subcta1
            
'            If txt_Tcta2.Text <> "" Then
              If rs_aux1.State = 1 Then rs_aux1.Close
              'rs_aux1.Open "select  * from CC_Plan_Cuentas where mov= 'S' and cuenta = '" & dtc_codigo3 & "' and SubCta1 = '" & dtc_codigo9 & "' and SubCta2 = '00' ", db, adOpenKeyset, adLockOptimistic
              rs_aux1.Open "select  * from CC_Plan_Cuentas where mov= 'S' and cuenta = '" & var_cod & "' and subcta1 = '" & VAR_COD2 & "'  ", db, adOpenKeyset, adLockOptimistic
              If rs_aux1.RecordCount > 0 Then
                txt_Tcta2.Text = rs_aux1("Cuenta")
                txt_Tscta12.Text = rs_aux1("SubCta1")
                txt_Tscta22.Text = rs_aux1("SubCta2")
                txt_desc2.Text = rs_aux1("NombreCta")
              End If
              
              If rs_aux2.State = 1 Then rs_aux2.Close
              'rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & txt_Tcta & "' and SubCta1 = '00' and SubCta2 = '00' ", db, adOpenKeyset, adLockOptimistic
              rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & var_cod & "' ", db, adOpenKeyset, adLockOptimistic
              If rs_aux2.RecordCount > 0 Then
                  txt_Tcta.Text = rs_aux2("Cuenta")
                  txt_Tscta1.Text = rs_aux2("SubCta1")
                  txt_Tscta2.Text = rs_aux2("SubCta2")
                  txt_desc1.Text = rs_aux2("NombreCta")
              End If
            If Ado_datos.Recordset!aux1 = "00" Then
                Chkaux1.Value = 0
                dtc_codigo4.Visible = False
                dtc_desc4.Visible = False
            Else
                Chkaux1.Value = 1
                dtc_codigo4.Visible = True
                dtc_desc4.Visible = True
            End If
            If Ado_datos.Recordset!AUX2 = "00" Then
                Chkaux2.Value = 0
                dtc_codigo5.Visible = False
                dtc_desc5.Visible = False
            Else
                Chkaux2.Value = 1
                dtc_codigo5.Visible = True
                dtc_desc5.Visible = True
            End If
            If Ado_datos.Recordset!aux3 = "00" Then
                Chkaux3.Value = 0
                dtc_codigo6.Visible = False
                dtc_desc6.Visible = False
            Else
                Chkaux3.Value = 1
                dtc_codigo6.Visible = True
                dtc_desc6.Visible = True
            End If
        End If
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
    Call OptFilGral2_Click
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
    dtc_codigo1.SetFocus
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
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM pc_poa_estrategico WHERE estado_codigo = 'APR' and dgral_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


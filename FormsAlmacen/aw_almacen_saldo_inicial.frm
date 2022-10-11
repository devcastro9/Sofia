VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_almacen_saldo_inicial 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clasificadores - RR.HH. - Cargos"
   ClientHeight    =   6030
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   5175
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   6255
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
         Left            =   1440
         TabIndex        =   34
         Top             =   4800
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
         Left            =   3720
         TabIndex        =   33
         Top             =   4800
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4680
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
         Caption         =   " "
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
         Height          =   4335
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   7646
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "almacen_descripcion"
            Caption         =   "Almacen"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Bien Descripcion"
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
            DataField       =   "cantidad_ingreso"
            Caption         =   "Cantidad Ingreso"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   5175
      Left            =   6360
      TabIndex        =   0
      Top             =   720
      Width           =   6615
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3285
         TabIndex        =   36
         Top             =   2895
         Width           =   255
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6285
         TabIndex        =   31
         Top             =   1820
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5445
         TabIndex        =   28
         Top             =   740
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "cantidad_ingreso"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Text            =   "0"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox TxtSueldo 
         BackColor       =   &H00FFFFFF&
         DataField       =   "sueldo_basico"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   "0"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dtc_alm_desc 
         Bindings        =   "aw_almacen_saldo_inicial.frx":0000
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   360
         TabIndex        =   25
         Top             =   720
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "almacen_descripcion"
         BoundColumn     =   "almacen_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_alm_cod 
         Bindings        =   "aw_almacen_saldo_inicial.frx":001A
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   26
         Top             =   720
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "almacen_codigo"
         BoundColumn     =   "almacen_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_bien_desc 
         Bindings        =   "aw_almacen_saldo_inicial.frx":0034
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   1800
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_bien_cod 
         Bindings        =   "aw_almacen_saldo_inicial.frx":004D
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   30
         Top             =   1800
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_alm_resp 
         Bindings        =   "aw_almacen_saldo_inicial.frx":0066
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4680
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "almacen_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_alm_tipo 
         Bindings        =   "aw_almacen_saldo_inicial.frx":0080
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "almacen_tipo"
         BoundColumn     =   "almacen_codigo"
         Text            =   "%"
      End
      Begin MSDataListLib.DataCombo dtc_medida 
         Bindings        =   "aw_almacen_saldo_inicial.frx":009A
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1800
         TabIndex        =   37
         Top             =   2880
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "unimed_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad Ingreso"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   3855
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   4680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado Registro:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   4365
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sueldo Basico:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   3525
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Denominación del Bien"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   1420
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Almacen"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   120
      Top             =   5880
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
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   10
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "aw_almacen_saldo_inicial.frx":00B3
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   19
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
         Picture         =   "aw_almacen_saldo_inicial.frx":0875
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   18
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
         Left            =   4080
         Picture         =   "aw_almacen_saldo_inicial.frx":1142
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "aw_almacen_saldo_inicial.frx":18F7
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   16
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
         Picture         =   "aw_almacen_saldo_inicial.frx":212A
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   15
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
         Picture         =   "aw_almacen_saldo_inicial.frx":2876
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   14
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
         Picture         =   "aw_almacen_saldo_inicial.frx":318B
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   13
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "aw_almacen_saldo_inicial.frx":394A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "aw_almacen_saldo_inicial.frx":3D8C
         Style           =   1  'Graphical
         TabIndex        =   11
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
         TabIndex        =   20
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
      TabIndex        =   21
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
         Picture         =   "aw_almacen_saldo_inicial.frx":3F96
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   23
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
         Picture         =   "aw_almacen_saldo_inicial.frx":476C
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   22
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
         TabIndex        =   24
         Top             =   195
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc Ado_almacen 
      Height          =   330
      Left            =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
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
      Caption         =   "Ado_almacen"
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
   Begin MSAdodcLib.Adodc Ado_bienes 
      Height          =   330
      Left            =   2040
      Top             =   5880
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
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
      Caption         =   "Ado_bienes"
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
Attribute VB_Name = "aw_almacen_saldo_inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Dim rs_almacen As New ADODB.Recordset
Dim rs_bienes As New ADODB.Recordset
Dim rs_verifica As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
'BUSCADOR
'Dim ClBuscaex As ClBuscaEnGridExterno

'Dim ClBuscaEx As ClBuscaEnGridExterno

Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As Integer
Dim VAR_VAL As String
Dim VAR_SWC, var_OP As String
Dim gestion, almacen, bien, estado, tipo_alm As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_datos!estado_codigo <> "ANL" Then
      If sino = vbYes Then
       var_OP = Ado_datos.Recordset!bien_codigo
        db.Execute "ap_almacen_saldo_inicial 6 ,'" & glGestion & "', '" & dtc_alm_cod.Text & "', '" & "R-114" & "', '" & "0" & "', '" & dtc_bien_cod.Text & "', '" & "20101-3" & "', " & "0" & ", '" & dtc_alm_resp.Text & "', '" & Date & "', '" & " " & "', '" & Date & "', '" & Text1.Text & "'," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
        
        db.Execute "update ac_bienes set ac_bienes.bien_stock_ingreso = total_ingresos_js.cantidad_ingreso from total_ingresos_js Where ac_bienes.bien_codigo = total_ingresos_js.bien_codigo"
       
        db.Execute "update ac_bienes set bien_stock_actual = bien_stock_ingreso - bien_stock_salida"
        
        
        Set rs_aux6 = New ADODB.Recordset
        If rs_aux6.State = 1 Then rs_aux6.Close
        rs_aux6.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo =" & rs_datos!almacen_codigo & " AND bien_codigo = '" & rs_datos!bien_codigo & "'", db, adOpenStatic
        If rs_aux6.RecordCount > 0 Then
        'db.Execute "update ao_almacen_totales set stock_ingreso  =" & CDbl(rs_det1A!adjudica_cantidad + rs_aux6!stock_ingreso) & ", total_compra_bs =" & rs_det1A!bien_total_adjudica_bs + rs_aux6!total_compra_bs & " WHERE almacen_codigo =" & rs_det1A!almacen_codigo & " AND bien_codigo = '" & rs_det1A!bien_codigo & "'"
        db.Execute "update ao_almacen_totales set stock_ingreso  = totales_almacen.cantidad_ingreso FROM totales_almacen WHERE totales_almacen.bien_codigo = ao_almacen_totales.bien_codigo and totales_almacen.almacen_codigo = ao_almacen_totales.almacen_codigo"
        db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida"
        
        Else
       db.Execute "INSERT INTO ao_almacen_totales (                   almacen_codigo,                  bien_codigo,                    stock_ingreso,    stock_salida,                stock_actual, total_compra_bs, total_venta_bs, utilidad_Bs, Total_compra_dol,total_venta_dol, utilidad_dol,estado_codigo, fecha_registro, usr_codigo)" & _
                                                 "Values(" & rs_datos!almacen_codigo & ", '" & rs_datos!bien_codigo & "', " & rs_datos!cantidad_ingreso & ", 0" & ", " & rs_datos!cantidad_ingreso & ", 0 , 0, 0, 0, 0, 0, 'REG', '" & Date & "', '" & glusuario & "')"
        End If
      
      
      
         rs_datos!estado_codigo = "APR"
         rs_datos!Fecha_Registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.Update
         Call OptFilGral2_Click
           OptFilGral2.Value = True



     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "bien_codigo = '" & var_OP & "' ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
          'Call ABRIR_TABLA_DET
          
        VAR_SW = ""
     Else
        VAR_SW = ""
        rs_datos.MoveLast
     End If
          
    End If
       
 
    
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
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
    
''    Set ClBuscaEx = New BuscadorSistema.ClBuscaEnGridExterno
'  Set ClBuscaex.Conexión = db
'  Set ClBuscaex.RecordsetTrabajo = rs_cargo
'  Set ClBuscaex.GridTrabajo = dg_datos
'  ClBuscaex.QueryUtilizado = queryinicial
'  ClBuscaex.EsTdbGrid = True
'  ClBuscaex.Ejecutar
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
   If VAR_SWC = "ADD" Then
   Call OptFilGral1_Click
   End If
   Call ABRIR_TABLA_AUX
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        FraNavega.Enabled = True
        dg_datos.Enabled = True
        rs_datos.Cancel
    End If
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_CARGO!estado_codigo = "S" Then
      If sino = vbYes Then
         rs_CARGO!estado_codigo = "L"
         rs_CARGO!Fecha_Registro = Date
         rs_CARGO!usr_codigo = glusuario
         rs_CARGO.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_CARGO!estado_codigo = "S" Then
      If sino = vbYes Then
         rs_CARGO!estado_codigo = "N"
         rs_CARGO!Fecha_Registro = Date
         rs_CARGO!usr_codigo = glusuario
         rs_CARGO.UpdateBatch adAffectAll
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
     If VAR_SWC = "ADD" Then

    Set rs_verifica = New ADODB.Recordset
    If rs_verifica.State = 1 Then rs_verifica.Close
    rs_verifica.Open "SELECT * FROM ao_almacen_ingresos  WHERE almacen_codigo = '" & dtc_alm_cod.Text & "' AND bien_codigo = '" & dtc_bien_cod.Text & "' and ges_gestion = " & glGestion & " and doc_codigo = 'R-114' and doc_numero = '0'", db, adOpenStatic
    If rs_verifica.RecordCount > 0 Then
    sino = MsgBox("Ya existe un saldo inicial de este bien para este almacen", vbInformation, "SOFIA")
'    Ado_datos.Recordset.Cancel
'    Call OptFilGral1_Click
    Exit Sub
    End If
    'rs_verifica.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
     db.Execute "ap_almacen_saldo_inicial 4 ,'" & glGestion & "', '" & dtc_alm_cod.Text & "', '" & "R-114" & "', '" & "0" & "', '" & dtc_bien_cod.Text & "', '" & "20101-3" & "', " & "0" & ", '" & dtc_alm_resp.Text & "', '" & Date & "', '" & " " & "', '" & Date & "', '" & Text1.Text & "'," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
     Ado_datos.Recordset.Cancel
     Else
     db.Execute "ap_almacen_saldo_inicial 5 ,'" & glGestion & "', '" & dtc_alm_cod.Text & "', '" & "R-114" & "', '" & "0" & "', '" & dtc_bien_cod.Text & "', '" & "20101-3" & "', " & "0" & ", '" & dtc_alm_resp.Text & "', '" & Date & "', '" & " " & "', '" & Date & "', '" & Text1.Text & "'," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
     End If
     
Call ABRIR_TABLA_AUX
Call OptFilGral1_Click

      FraNavega.Enabled = True
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      
'      If (dg_datos.SelBookmarks.Count <> 0) Then
'        dg_datos.SelBookmarks.Remove 0
'     End If
'     If Ado_datos.Recordset.RecordCount > 0 Then
'        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
'        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
'         If rs_det1.RecordCount > 0 Then
    '         rs_det1.MoveLast
'        End If
'     Else
'        rs_datos.MoveLast
'     End If
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If dtc_alm_cod.Text = "" Then
    MsgBox "Debe registrar el ALMACEN ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If dtc_bien_cod.Text = "" Then
    MsgBox "Debe registrar el BIEN ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If Text1.Text = "" Or Text1.Text <= 0 Then
    MsgBox "El saldo inicial no puede ser Menor o Igual a 0 ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  cr01.WindowShowPrintSetupBtn = True
  cr01.WindowShowRefreshBtn = True
  cr01.ReportFileName = App.Path & "\REPORTES\clasificadores\rr_cargos.rpt"
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
    VAR_SWC = "MOD"
    dtc_alm_desc.Enabled = False
   dtc_bien_desc.Enabled = False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
  If glPersOtro = "CGO" Then
    frmmc_personal.dtc_cargo = rs_CARGO!cargo_codigo
    frmmc_personal.Dtc_cargoDes = rs_CARGO!cargo_descripcion
  End If
  glPersOtro = "N"
  Unload Me
End Sub

Private Sub dtc_alm_cod_Click(Area As Integer)
dtc_alm_desc.BoundText = dtc_alm_cod.BoundText
dtc_alm_resp.BoundText = dtc_alm_cod.BoundText
dtc_alm_tipo.BoundText = dtc_alm_cod.BoundText
End Sub

Private Sub dtc_alm_desc_Change()
'Set rs_bienes = New ADODB.Recordset
'    If rs_bienes.State = 1 Then rs_bienes.Close
'    'rs_datos4.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo_pla = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo <> 'ANL' order by beneficiario_denominacion", db, adOpenStatic
'     sql = "ap_almacen_saldo_inicial 3," & glGestion & ", '" & dtc_alm_cod.Text & "', '" & "R-114" & "', " & "0" & ", " & dtc_bien_cod.Text & ", '" & dtc_alm_tipo.Text & "', " & "0" & ", '" & dtc_alm_resp.Text & "', '" & Date & "', '" & " " & "', '" & Date & "', " & "0" & "," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
'    rs_bienes.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
'   Set Ado_bienes.Recordset = rs_bienes
'   dtc_bien_cod.BoundText = dtc_bien_desc.BoundText
End Sub

Private Sub dtc_alm_desc_Click(Area As Integer)
dtc_alm_cod.BoundText = dtc_alm_desc.BoundText
dtc_alm_resp.BoundText = dtc_alm_desc.BoundText
dtc_alm_tipo.BoundText = dtc_alm_desc.BoundText

Set rs_bienes = New ADODB.Recordset
    If rs_bienes.State = 1 Then rs_bienes.Close
    'rs_datos4.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo_pla = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo <> 'ANL' order by beneficiario_denominacion", db, adOpenStatic
     sql = "ap_almacen_saldo_inicial 3," & glGestion & ", '" & "0" & "', '" & "R-114" & "', " & "0" & ", " & "0" & ", '" & IIf(dtc_alm_tipo.Text = "", "%", dtc_alm_tipo.Text) & "', " & "0" & ", " & "0" & ", '" & Date & "', '" & "0" & "', '" & Date & "', " & "0" & "," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
    rs_bienes.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_bienes.Recordset = rs_bienes
   dtc_bien_cod.BoundText = dtc_bien_desc.BoundText

End Sub
Private Sub ABRIR_TABLA()

  Set rs_datos = New ADODB.Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'rs_datos4.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo_pla = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo <> 'ANL' order by beneficiario_denominacion", db, adOpenStatic
     sql = "ap_almacen_saldo_inicial 1 ,'" & estado & "', " & "0" & ", '" & "R-114" & "', " & "0" & ", " & "0" & ", '" & "20101-3" & "', " & "0" & ", '" & usuario2 & "', '" & Date & "', '" & " " & "', '" & Date & "', '" & "0" & "'," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
     queryinicial = "ap_almacen_saldo_inicial 1 ,'" & estado & "', " & "0" & ", '" & "R-114" & "', " & "0" & ", " & "0" & ", '" & "20101-3" & "', " & "0" & ", '" & usuario2 & "', '" & Date & "', '" & " " & "', '" & Date & "', '" & "0" & "'," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
    rs_datos.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_datos.Recordset = rs_datos
'   dtc_descripcion.BoundText = dtc_codigo.BoundText
  Set dg_datos.DataSource = rs_datos

  
'  Set ClBuscaex = New ClBuscaEnGridExterno
'
'  mbDataChanged = False
'  Fra_ABM.Enabled = False
'  dg_datos.Enabled = True
End Sub

Private Sub dtc_bien_cod_Click(Area As Integer)
dtc_bien_desc.BoundText = dtc_bien_cod.BoundText
dtc_medida.BoundText = dtc_bien_cod.BoundText
End Sub

Private Sub dtc_bien_desc_Click(Area As Integer)
dtc_bien_cod.BoundText = dtc_bien_desc.BoundText
dtc_medida.BoundText = dtc_bien_desc.BoundText
End Sub

Private Sub Form_Load()
Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_BENEF = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_BENEF = "0"
        VAR_DA = "1.3"
    End If
  Call ABRIR_TABLA_AUX
  Call OptFilGral1_Click
  
End Sub

Private Sub ABRIR_TABLA_AUX()

    Set rs_almacen = New ADODB.Recordset
    If rs_almacen.State = 1 Then rs_almacen.Close
    'rs_datos4.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo_pla = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo <> 'ANL' order by beneficiario_denominacion", db, adOpenStatic
    sql = "ap_almacen_saldo_inicial 2 ," & glGestion & ", " & dtc_alm_cod.Text & ", '" & "R-114" & "', " & "0" & ", '" & dtc_bien_cod.Text & "', '" & "20101-3" & "', " & "0" & ", " & usuario2 & ", '" & Date & "', '" & " " & "', '" & Date & "', " & Text1.Text & "," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
    rs_almacen.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_almacen.Recordset = rs_almacen
   dtc_alm_cod.BoundText = dtc_alm_desc.BoundText
   
   Set rs_bienes = New ADODB.Recordset
    If rs_bienes.State = 1 Then rs_bienes.Close
    'rs_datos4.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo_pla = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo <> 'ANL' order by beneficiario_denominacion", db, adOpenStatic
     sql = "ap_almacen_saldo_inicial 3," & glGestion & ", " & "0" & ", '" & "R-114" & "', " & "0" & ", '" & dtc_bien_cod.Text & "', '" & "%" & "', " & "0" & ", " & dtc_alm_resp.Text & ", '" & Date & "', '" & " " & "', '" & Date & "', " & Text1.Text & "," & "0" & "," & "0" & "," & "REG" & ", '" & Date & "', '" & Format(Time, "HH:mm:ss") & "', " & glusuario
    rs_bienes.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_bienes.Recordset = rs_bienes
   dtc_bien_cod.BoundText = dtc_bien_desc.BoundText
   
End Sub


Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set ClBuscaGrid = Nothing
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
   'rs_datos.MoveLast
   rs_datos.AddNew
   dtc_alm_desc.Enabled = True
   dtc_bien_desc.Enabled = True
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = False
    dg_datos.Enabled = False
    
    VAR_SWC = "ADD"
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_CARGO.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub txt_cargo_nivel_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub

Private Sub OptFilGral1_Click()
estado = "REG"
Call ABRIR_TABLA

End Sub

Private Sub OptFilGral2_Click()
estado = "%"
Call ABRIR_TABLA
End Sub

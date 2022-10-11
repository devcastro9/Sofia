VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ac_NoObjecion_c1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos  Administrativos  - Contratacion de  Personal - Registro Inicial"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11790
   Icon            =   "ac_NoObjecion_c1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraOpciones1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   114
      Top             =   720
      Width           =   8340
      Begin VB.CommandButton CmdCopiar 
         Caption         =   "Copiar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   7050
         Picture         =   "ac_NoObjecion_c1.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Copia el comprobante de Ingreso a uno nuevo"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   2535
         MaskColor       =   &H8000000F&
         Picture         =   "ac_NoObjecion_c1.frx":10D4
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Busca un Comprobante de Ingreso"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   1005
         Picture         =   "ac_NoObjecion_c1.frx":12DE
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Modifica el comprobante de Ingreso"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdAñadir 
         Caption         =   "Adicionar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   240
         Picture         =   "ac_NoObjecion_c1.frx":14E8
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Adiciona un comprobante de Ingreso"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "Anular"
         Enabled         =   0   'False
         Height          =   720
         Left            =   1770
         Picture         =   "ac_NoObjecion_c1.frx":17F2
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Anula el comprobante de Ingreso"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   720
         Left            =   5670
         Picture         =   "ac_NoObjecion_c1.frx":1EDC
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Sale del Formulario de Ingresos"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   720
         Left            =   4905
         Picture         =   "ac_NoObjecion_c1.frx":20E6
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Imprime el comprobante de Ingreso"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Confirma Envío"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4065
         Picture         =   "ac_NoObjecion_c1.frx":27D0
         TabIndex        =   116
         ToolTipText     =   "Imprime el comprobante de Ingreso"
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton cmdAPublicacion 
         Caption         =   "Publicación"
         Enabled         =   0   'False
         Height          =   720
         Left            =   3300
         Picture         =   "ac_NoObjecion_c1.frx":2EBA
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Habilita para publicación"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Bindings        =   "ac_NoObjecion_c1.frx":30C4
      Height          =   7455
      Left            =   8400
      TabIndex        =   3
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
      HeadLines       =   2
      RowHeight       =   23
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
         DataField       =   "ges_gestion"
         Caption         =   "Gestion"
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
         DataField       =   "codigo_unidad"
         Caption         =   "unidad"
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
         DataField       =   "codigo_solicitud"
         Caption         =   "solicitud"
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
         DataField       =   "Enviado"
         Caption         =   "Enviado"
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
         Locked          =   -1  'True
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   689.953
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cr 
      Left            =   1080
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCancelBtn=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame FrameP 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   8175
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trámite"
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   21
         Left            =   2760
         TabIndex        =   66
         Top             =   240
         Width           =   855
      End
      Begin VB.Label labNumero_consultoria 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3720
         TabIndex        =   65
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   9
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "codigo_unidad"
         DataSource      =   "adoMain"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "uni_descripcion_larga"
         DataSource      =   "adoMain"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lab 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "codigo_solicitud"
         DataSource      =   "adoMain"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Solicitud"
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label labFormulario 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "formulario"
         DataSource      =   "adoMain"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de formulario"
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label labGes_gestion 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ges_gestion"
         DataSource      =   "adoMain"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión"
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   330
      Left            =   8400
      Top             =   8160
      Width           =   3375
      _ExtentX        =   5953
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
      BackColor       =   12648447
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
      Caption         =   "adoMain"
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
   Begin TabDlg.SSTab ssTab 
      Height          =   6855
      Left            =   0
      TabIndex        =   14
      Top             =   1680
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   635
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Datos de la Solicitud"
      TabPicture(0)   =   "ac_NoObjecion_c1.frx":30DA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "adoDetalleSolicitud"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dgDetalleSolicitud"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Datos de la Contratación"
      TabPicture(1)   =   "ac_NoObjecion_c1.frx":30F6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "cboCorrelativo"
      Tab(1).Control(2)=   "cboModalidadContratacion"
      Tab(1).Control(3)=   "cboModalidadSeleccion"
      Tab(1).Control(4)=   "adoListaCorta"
      Tab(1).Control(5)=   "dgListacorta"
      Tab(1).Control(6)=   "txtDuracion_Estimada_Tiempo"
      Tab(1).Control(7)=   "txtJustifSeleccion"
      Tab(1).Control(8)=   "txtObjetivo"
      Tab(1).Control(9)=   "txtSolNoObjecion"
      Tab(1).Control(10)=   "frlicitacionpara"
      Tab(1).Control(11)=   "dtpFEstInicioCons"
      Tab(1).Control(12)=   "labCorrelativo_Consultoria"
      Tab(1).Control(13)=   "Label1(22)"
      Tab(1).Control(14)=   "labParaPublicacion"
      Tab(1).Control(15)=   "lblmodalidadlicita(5)"
      Tab(1).Control(16)=   "lblmodalidadlicita(3)"
      Tab(1).Control(17)=   "lblobservaciones(3)"
      Tab(1).Control(18)=   "lblobservaciones(2)"
      Tab(1).Control(19)=   "lblmodalidadlicita(6)"
      Tab(1).Control(20)=   "lblmodalidadlicita(0)"
      Tab(1).Control(21)=   "lblobservaciones(1)"
      Tab(1).Control(22)=   "lblobservaciones(0)"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "Datos de la Autorización"
      TabPicture(2)   =   "ac_NoObjecion_c1.frx":3112
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCitesMuestra"
      Tab(2).Control(1)=   "fraAdiCite"
      Tab(2).Control(2)=   "adoMotivos"
      Tab(2).Control(3)=   "fracites"
      Tab(2).Control(4)=   "adoCite"
      Tab(2).Control(5)=   "adoDocAdjunta"
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame7 
         Height          =   385
         Left            =   -74880
         TabIndex        =   111
         Top             =   4880
         Visible         =   0   'False
         Width           =   2775
         Begin VB.OptionButton opFiles 
            Caption         =   "Varios files"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   113
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton opFiles 
            Caption         =   "Un file"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   112
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboCorrelativo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   4920
         Width           =   3615
      End
      Begin VB.ComboBox cboModalidadContratacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   4560
         Width           =   3855
      End
      Begin VB.ComboBox cboModalidadSeleccion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71040
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   92
         Top             =   6000
         Width           =   6615
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nro. Comprobante"
            Height          =   255
            Index           =   3
            Left            =   3960
            TabIndex        =   96
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   95
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            DataField       =   "codigo_pago"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   5400
            TabIndex        =   94
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "fecha_egreso"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   93
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc adoDocAdjunta 
         Height          =   330
         Left            =   -70560
         Top             =   6360
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "adoDocAdjunta"
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
      Begin MSAdodcLib.Adodc adoCite 
         Height          =   330
         Left            =   -74640
         Top             =   6360
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "adoCite"
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
      Begin MSAdodcLib.Adodc adoListaCorta 
         Height          =   330
         Left            =   -70080
         Top             =   6480
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "adoListaCorta"
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
      Begin MSDataGridLib.DataGrid dgListacorta 
         Bindings        =   "ac_NoObjecion_c1.frx":312E
         Height          =   1215
         Left            =   -74880
         TabIndex        =   51
         Top             =   5520
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2143
         _Version        =   393216
         BackColor       =   12632319
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
            DataField       =   "descripcion_correlativo"
            Caption         =   "Tipo File"
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
            DataField       =   "correlativo_consultoria"
            Caption         =   "Nro File"
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
            DataField       =   "ci"
            Caption         =   "Doc. Identidad"
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
            DataField       =   "paterno"
            Caption         =   "Primer apellido"
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
            DataField       =   "materno"
            Caption         =   "Segundo apellido"
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
            DataField       =   "nombres"
            Caption         =   "Nombres"
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
            DataField       =   "aunidad"
            Caption         =   "Incorporar a Unidad"
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
            DataField       =   "aplanilla"
            Caption         =   "Incorporar a Planilla"
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
            Locked          =   -1  'True
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDuracion_Estimada_Tiempo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73800
         TabIndex        =   73
         Text            =   "txtDuracion_Estimada_Tiempo"
         Top             =   3980
         Width           =   3135
      End
      Begin VB.TextBox txtJustifSeleccion 
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73800
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   70
         Text            =   "ac_NoObjecion_c1.frx":314A
         Top             =   3480
         Width           =   7095
      End
      Begin VB.Frame fracites 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   -74880
         TabIndex        =   54
         Top             =   1680
         Width           =   8175
         Begin VB.OptionButton optenviados 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ambos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   4920
            TabIndex        =   81
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optenviados 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recibidos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2880
            TabIndex        =   80
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optenviados 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enviados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   720
            TabIndex        =   55
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtObjetivo 
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73800
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   50
         Text            =   "ac_NoObjecion_c1.frx":315F
         Top             =   3000
         Width           =   7095
      End
      Begin VB.TextBox txtSolNoObjecion 
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73800
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   49
         Text            =   "ac_NoObjecion_c1.frx":316D
         Top             =   2520
         Width           =   7095
      End
      Begin VB.Frame frlicitacionpara 
         Height          =   855
         Left            =   -74880
         TabIndex        =   46
         Top             =   1560
         Width           =   8175
         Begin VB.TextBox txtdesConsultoria 
            Enabled         =   0   'False
            Height          =   525
            Left            =   1560
            MaxLength       =   200
            TabIndex        =   47
            Top             =   210
            Width           =   6375
         End
         Begin VB.Label lbldesclicita 
            Alignment       =   1  'Right Justify
            Caption         =   "Descripción del Trámite"
            Height          =   495
            Left            =   120
            TabIndex        =   48
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   120
         TabIndex        =   36
         Top             =   4920
         Width           =   6615
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "par_descripcion_larga"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   2040
            TabIndex        =   44
            Top             =   600
            Width           =   4455
         End
         Begin VB.Label Label20 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "pro_descripcion_larga"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1920
            TabIndex        =   43
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label labPro_actividad 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "pro_actividad"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1680
            TabIndex        =   42
            Top             =   240
            Width           =   255
         End
         Begin VB.Label labPro_proyecto 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "pro_proyecto"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1440
            TabIndex        =   41
            Top             =   240
            Width           =   255
         End
         Begin VB.Label labPro_programa 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "pro_programa"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   40
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "par_codigo"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Partida"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Proyecto"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   6615
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "denominacion_categoria"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   2040
            TabIndex        =   35
            Top             =   600
            Width           =   4455
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "denominacion_convenio"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   2040
            TabIndex        =   34
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Convenio"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Categoría"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.Label labCodigo_categoria 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "codigo_categoria"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   31
            Top             =   600
            Width           =   855
         End
         Begin VB.Label labCodigo_convenio 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "codigo_convenio"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   6615
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "descripcion_poa"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   2280
            TabIndex        =   28
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "codigo_poa"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Solicitud."
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   6615
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Bs"
            Height          =   255
            Index           =   16
            Left            =   6120
            TabIndex        =   126
            Top             =   960
            Width           =   255
         End
         Begin VB.Label labMonto_dolares_ext 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "monto_dolares"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   106
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "fte_descripcion_larga"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   2040
            TabIndex        =   23
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "org_descripcion_larga"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   2040
            TabIndex        =   22
            Top             =   600
            Width           =   4455
         End
         Begin VB.Label labMonto_bolivianos_ext 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "monto_bolivianos"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   4800
            TabIndex        =   24
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label labFte_codigo_ext 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "fte_codigo"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.Label labOrg_codigo_ext 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "org_codigo"
            DataSource      =   "adoDetalleSolicitud"
            Height          =   255
            Left            =   1200
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "$us"
            Height          =   255
            Index           =   8
            Left            =   2520
            TabIndex        =   19
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Monto"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   18
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Financiamien."
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Financ."
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid dgDetalleSolicitud 
         Bindings        =   "ac_NoObjecion_c1.frx":3180
         Height          =   4695
         Left            =   6840
         TabIndex        =   45
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   8281
         _Version        =   393216
         BackColor       =   12648384
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
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "codigo_poa"
            Caption         =   "Solicitud"
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
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoDetalleSolicitud 
         Height          =   375
         Left            =   6840
         Top             =   6360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         BackColor       =   12648384
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
         Caption         =   "adoDetalleSolicitud"
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
      Begin MSComCtl2.DTPicker dtpFEstInicioCons 
         Height          =   375
         Left            =   -68400
         TabIndex        =   74
         Top             =   3980
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   55181313
         CurrentDate     =   36749
      End
      Begin MSAdodcLib.Adodc adoMotivos 
         Height          =   330
         Left            =   -72720
         Top             =   6360
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "adoMotivos"
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
      Begin VB.Frame fraAdiCite 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   56
         Top             =   2280
         Visible         =   0   'False
         Width           =   8175
         Begin VB.Frame fraRespuestaDeRecibidos 
            Height          =   975
            Left            =   240
            TabIndex        =   98
            Top             =   600
            Visible         =   0   'False
            Width           =   7695
            Begin VB.ComboBox cboSelCite 
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   104
               Top             =   240
               Width           =   5775
            End
            Begin VB.OptionButton opNoObj 
               Caption         =   "Si"
               Height          =   195
               Index           =   1
               Left            =   5040
               TabIndex        =   100
               Top             =   640
               Width           =   615
            End
            Begin VB.OptionButton opNoObj 
               Caption         =   "No"
               Height          =   195
               Index           =   0
               Left            =   4320
               TabIndex        =   99
               Top             =   640
               Width           =   615
            End
            Begin VB.Label lblcite 
               Alignment       =   1  'Right Justify
               Caption         =   "Respuesta al CITE:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   105
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label labPreguntaNoObjecion 
               Alignment       =   1  'Right Justify
               Caption         =   "La Autorización fue rechazada ?"
               Height          =   255
               Left            =   240
               TabIndex        =   101
               Top             =   640
               Width           =   3975
            End
         End
         Begin VB.ListBox lstMotivos 
            Height          =   1410
            ItemData        =   "ac_NoObjecion_c1.frx":31A2
            Left            =   240
            List            =   "ac_NoObjecion_c1.frx":31A4
            Style           =   1  'Checkbox
            TabIndex        =   87
            Top             =   3000
            Width           =   3855
         End
         Begin VB.ListBox lstDocAdjunta 
            Height          =   1410
            Left            =   4200
            Style           =   1  'Checkbox
            TabIndex        =   58
            Top             =   3000
            Width           =   3735
         End
         Begin VB.TextBox txtcite 
            Height          =   285
            Left            =   2520
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dtpFCite 
            Height          =   375
            Left            =   6000
            TabIndex        =   59
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   55181313
            CurrentDate     =   36739
         End
         Begin VB.Label labPrefijo 
            Alignment       =   1  'Right Justify
            Caption         =   "/"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1080
            TabIndex        =   107
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblcite 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha del cite:"
            Height          =   255
            Index           =   1
            Left            =   4680
            TabIndex        =   97
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblmotivo 
            Caption         =   "Motivo(s) de la Autorización"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   86
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label labOrg_cargo 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   85
            Top             =   2400
            Width           =   6255
         End
         Begin VB.Label labOrg_representante 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   84
            Top             =   2040
            Width           =   6255
         End
         Begin VB.Label labOrg_descripcion 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2520
            TabIndex        =   83
            Top             =   1680
            Width           =   5415
         End
         Begin VB.Label labOrg_Codigo 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   82
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblentidad 
            Alignment       =   1  'Right Justify
            Caption         =   "Entidad/Unidad:"
            Height          =   255
            Left            =   360
            TabIndex        =   64
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblmotivo 
            Caption         =   "Documentación adjunta:"
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   63
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label lblcargo 
            Alignment       =   1  'Right Justify
            Caption         =   "Cargo:"
            Height          =   255
            Left            =   480
            TabIndex        =   62
            Top             =   2430
            Width           =   1095
         End
         Begin VB.Label lblrepresentante 
            Alignment       =   1  'Right Justify
            Caption         =   "Representante:"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   2070
            Width           =   1335
         End
         Begin VB.Label lblcite 
            Caption         =   "Nro. CITE:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame fraCitesMuestra 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   88
         Top             =   2280
         Width           =   7695
         Begin MSDataGridLib.DataGrid dgcite 
            Bindings        =   "ac_NoObjecion_c1.frx":31A6
            Height          =   1815
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3201
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12632319
            HeadLines       =   2
            RowHeight       =   19
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
               DataField       =   "nro_cite"
               Caption         =   "CITE"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Fecha_cite"
               Caption         =   "Fecha cite"
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
               DataField       =   "enviado"
               Caption         =   "Enviado / Recibido"
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
               DataField       =   "fecha_envio_recepcion"
               Caption         =   "Fecha Envio / Recepción"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "org_descripcion_larga"
               Caption         =   "Financiamiento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "org_representante"
               Caption         =   "Representante"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "org_cargo"
               Caption         =   "Cargo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "AceptaronNoObjecion"
               Caption         =   "Vo.Bo. Financiador"
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
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column07 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgMotivos 
            Bindings        =   "ac_NoObjecion_c1.frx":31BC
            Height          =   975
            Left            =   120
            TabIndex        =   90
            Top             =   2160
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   1720
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "denominacion_motivo"
               Caption         =   "Motivos"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgDocAdjunta 
            Bindings        =   "ac_NoObjecion_c1.frx":31D5
            Height          =   855
            Left            =   120
            TabIndex        =   91
            Top             =   3240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   1508
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "denominacion_doc_adjunta"
               Caption         =   "Documentación Adjunta"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label labCorrelativo_Consultoria 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70320
         TabIndex        =   110
         Top             =   4920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Correlativo de files :"
         Height          =   255
         Index           =   22
         Left            =   -71880
         TabIndex        =   108
         Top             =   4950
         Width           =   1575
      End
      Begin VB.Label labParaPublicacion 
         Caption         =   "labParaPublicacion"
         Height          =   255
         Left            =   -70320
         TabIndex        =   102
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblmodalidadlicita 
         Caption         =   "Modalidad de contratación"
         Height          =   255
         Index           =   5
         Left            =   -74805
         TabIndex        =   79
         Top             =   4335
         Width           =   1935
      End
      Begin VB.Label lblmodalidadlicita 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha estimada de inicio"
         Height          =   255
         Index           =   3
         Left            =   -70560
         TabIndex        =   75
         Top             =   4020
         Width           =   1935
      End
      Begin VB.Label lblobservaciones 
         Caption         =   "Duración . . ."
         Height          =   255
         Index           =   3
         Left            =   -74805
         TabIndex        =   72
         Top             =   4020
         Width           =   1095
      End
      Begin VB.Label lblobservaciones 
         Caption         =   "Justificación de selección"
         Height          =   495
         Index           =   2
         Left            =   -74805
         TabIndex        =   71
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblmodalidadlicita 
         Caption         =   "Consultor(es) seleccionado(s)"
         Height          =   200
         Index           =   6
         Left            =   -74760
         TabIndex        =   69
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Label lblmodalidadlicita 
         Alignment       =   1  'Right Justify
         Caption         =   "Modalidad de Selección"
         Height          =   255
         Index           =   0
         Left            =   -71160
         TabIndex        =   68
         Top             =   4335
         Width           =   1935
      End
      Begin VB.Label lblobservaciones 
         Caption         =   "Objetivo . . . ."
         Height          =   255
         Index           =   1
         Left            =   -74805
         TabIndex        =   53
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblobservaciones 
         Caption         =   "Autorización para Solicitud"
         Height          =   570
         Index           =   0
         Left            =   -74805
         TabIndex        =   52
         Top             =   2460
         Width           =   975
      End
   End
   Begin VB.Frame FraOpciones2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   8340
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   360
         MousePointer    =   4  'Icon
         Picture         =   "ac_NoObjecion_c1.frx":31F1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   1200
         MousePointer    =   4  'Icon
         Picture         =   "ac_NoObjecion_c1.frx":34FB
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.Frame fraOpciones0 
      BackColor       =   &H00C0C0C0&
      Height          =   930
      Left            =   0
      TabIndex        =   76
      Top             =   720
      Width           =   8340
      Begin VB.CommandButton cmdSalir1 
         Caption         =   "Salir"
         Height          =   720
         Left            =   240
         Picture         =   "ac_NoObjecion_c1.frx":3805
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Sale del Formulario de Ingresos"
         Top             =   160
         Width           =   765
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   120
      TabIndex        =   125
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   120
      TabIndex        =   124
      Top             =   360
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRATACION DE PERSONAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Index           =   15
      Left            =   6120
      TabIndex        =   103
      Top             =   120
      Width           =   5535
   End
   Begin VB.Image Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   0
      Picture         =   "ac_NoObjecion_c1.frx":3A0F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11730
   End
End
Attribute VB_Name = "ac_NoObjecion_c1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim filtro As String    'para filtrar los cites enviados, recibidos o todos
Dim Adicionar As Boolean 'para saber si se adiciona o se modifica una consultoria


Private Sub adoCite_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Me.adoCite.Recordset.RecordCount > 0 Then
    Call refrescaMotivos
    Call refrescaDocAdjunta
End If
Call CuidaBotones
End Sub

Sub refrescaMotivos()
''refresca la lista de motivos de la no objecion
If Me.adoCite.Recordset.RecordCount > 0 Then
    DE.dbo_edListaMotivosCites Me.adoCite.Recordset!numero_consultoria, Me.adoCite.Recordset!numero_consultoria_cite
    With DE.rsdbo_edListaMotivosCites
        Set Me.adoMotivos.Recordset = .Clone
        .Close
    End With
Else
    DE.dbo_edListaMotivosCites 0, 0
    With DE.rsdbo_edListaMotivosCites
        Set Me.adoMotivos.Recordset = .Clone
        .Close
    End With
End If
End Sub


Sub refrescaDocAdjunta()
'refresca la lista de documentacion adjunta de la no oibjecion
If Me.adoCite.Recordset.RecordCount > 0 Then
    DE.dbo_edListaDocAdjuntaCites Me.adoCite.Recordset!numero_consultoria, Me.adoCite.Recordset!numero_consultoria_cite
    With DE.rsdbo_edListaDocAdjuntaCites
        Set Me.adoDocAdjunta.Recordset = .Clone
        .Close
    End With
Else
    DE.dbo_edListaDocAdjuntaCites 0, 0
    With DE.rsdbo_edListaDocAdjuntaCites
        Set Me.adoDocAdjunta.Recordset = .Clone
        .Close
    End With
End If
End Sub


Private Sub adoMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Call refrescaDetalleComun
Select Case Me.ssTab.Tab
Case 0
    Call refrescaDetalleSolicitud
Case 1
    Call refrescaDetalleConsultoria
Case 2
    Call refrescaDetalleCites
End Select
Call CuidaBotones
End Sub


Private Sub cboSelCite_Click()
'cada vez que selecciona un cite enviado, le busca su lista
'de documentacion adjunta y motivos de la no objecion
Dim i%
Dim rsm As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
rsm.Open "select * from ao_no_objecion_motivo_c where numero_consultoria=" & (Me.labNumero_consultoria) & " and numero_consultoria_cite=" & Val(Mid(Me.cboSelCite, 2, 4)), db, adOpenStatic, adLockReadOnly
rsd.Open "select * from ao_cite_doc_adjunta_c where numero_consultoria=" & (Me.labNumero_consultoria) & " and numero_consultoria_cite=" & Val(Mid(Me.cboSelCite, 2, 4)), db, adOpenStatic, adLockReadOnly
For i = 0 To Me.lstDocAdjunta.ListCount - 1
    Me.lstDocAdjunta.Selected(i) = False
Next
For i = 0 To Me.lstMotivos.ListCount - 1
    Me.lstMotivos.Selected(i) = False
Next
'de la listade motivos
With rsm
    If .RecordCount > 0 Then
        For i = 0 To Me.lstMotivos.ListCount - 1
            .MoveFirst
            .Find "codigo_motivo='" & Left(Me.lstMotivos.List(i), 2) & "'"
            If Not .EOF Then
                Me.lstMotivos.Selected(i) = True
            End If
        Next i
    End If
End With
'de la lista de documentacion adjunta
With rsd
    If .RecordCount > 0 Then
        For i = 0 To Me.lstDocAdjunta.ListCount - 1
            .MoveFirst
            .Find "codigo_doc_adjunta='" & Left(Me.lstDocAdjunta.List(i), 2) & "'"
            If Not .EOF Then
                Me.lstDocAdjunta.Selected(i) = True
            End If
        Next i
    End If
End With

End Sub


Private Sub cmdAPublicacion_Click()
'marca la consultoria para que se habilite en em submodulo de publicacion
If MsgBox("Desea habilitar la consultoria para el proceso de Publicación ?", vbYesNo) = vbYes Then
    DE.dbo_apGeneralSearching "update ao_no_objecion_c set ParaPublicacion='S' where numero_consultoria=" & Val(Me.labNumero_consultoria)
    Me.cmdAPublicacion.Enabled = False
End If
End Sub


Private Sub CmdBorrar_Click()
Select Case Me.ssTab.Tab
Case 0
Case 1
    'borra los datos de la consultoria,
    If MsgBox("Desea eliminar la consultoría indicada ?", vbYesNo) = vbYes Then
        DE.dbo_edBorraConsultoria Val(Me.labNumero_consultoria)
        Call RefrescaMainList
        Call refrescaDetalleConsultoria
    End If
Case 2
    'mata el cite junto con las relaciones a documentacion adjunta y motivos de la no objecion
    If MsgBox("Desea eliminar el cite indicado ?", vbYesNo) = vbYes Then
        With Me.adoCite.Recordset
        DE.dbo_edBorraCite !numero_consultoria, !numero_consultoria_cite
        End With
        Call refrescaDetalleCites
    End If
End Select
End Sub

Private Sub cmdCancelar_Click()
'*******CANCELAR UNA ADICION O MODIFICACION
Select Case ssTab.Tab
    Case 2
    Me.fraAdiCite.Visible = False
    fracites.Enabled = True
End Select
Me.dgMain.Enabled = True
fraOpciones0.Visible = False
FraOpciones1.Visible = True
FraOpciones2.Visible = False
Me.fraAdiCite.Visible = False
Call BloqueaTodosLosCamposEditables
Call Me.refrescaDetalleConsultoria
End Sub


Private Sub CmdEnviar_Click()
Dim Mensa$, fechax$, horax$
'marca el cite enviado =S y le pone la fecha de envio
If Me.optenviados(0).Value = True Then
    Mensa = "Desea confirmar el envío del cite ?"
End If
'marca el cite recibido =S y le pone la fecha de recepcion
If Me.optenviados(1).Value = True Then
    Mensa = "Desea confirmar la recepción ?"
End If

If MsgBox(Chr(13) & Mensa, vbYesNo) = vbYes Then
    DE.dbo_edGetProcessDateTime fechax, horax
    DE.dbo_apGeneralSearching "update ao_no_objecion_cite_c set fecha_envio_recepcion ='" & fechax & "', enviado='S' where numero_consultoria=" & Me.adoCite.Recordset!numero_consultoria & " and numero_consultoria_cite=" & Me.adoCite.Recordset!numero_consultoria_cite
    Call refrescaDetalleCites
End If
End Sub

Private Sub CmdGrabar_Click()
'graba una adiocion o modificacion de la consultoria en el caso = 1
'en el caso=2 graba el cite adicionado o modificado
Dim i%, xcorrelativo_consultoria%
Dim APE_ESPOSO As String
Dim Monto_solicitud_dl As Integer
Dim Nro_pagos As Integer
Dim aunidad As String
Dim aplanilla As Integer
Dim tipo_documento As String
Select Case ssTab.Tab
Case 1
    'GRABA CONSULTORIAS
    Dim numero_consultoria_out As Integer
    Dim rsX As New ADODB.Recordset
    numero_consultoria_out = 0
    If TodoBienConsultoria() Then
        'aqui graba el registro de la consultoria
        DE.dbo_apGeneralSearching "select * from pagos_espera pe where ES_BASE='S' AND formulario='F05' and ges_gestion='" & Me.adoMain.Recordset!ges_gestion & "' and codigo_solicitud=" & Me.adoMain.Recordset!codigo_solicitud & " and codigo_unidad = '" & Me.adoMain.Recordset!codigo_unidad & "'"
        With DE.rsdbo_apGeneralSearching
            If .RecordCount > 0 Then
                If Adicionar = True Then
                    DE.dbo_edGrabaConsultoria !ges_gestion, !org_codigo, !codigo_pago, !codigo_solicitud, !codigo_unidad, 0, Me.cboCorrelativo.ItemData(Me.cboCorrelativo.ListIndex), "N", "", "", Trim(Left(Me.cboModalidadContratacion, 6)), !Codigo_convenio, Trim(Me.labPro_programa), Trim(Me.labPro_proyecto), Trim(Me.labPro_actividad), Trim(Me.txtdesConsultoria), "", Trim(Left(Me.cboModalidadSeleccion, 6)), Trim(Me.txtJustifSeleccion), !duracion_estimada_tiempo, Me.dtpFEstInicioCons, GlUsuario, !formulario, Trim(Me.txtSolNoObjecion), Trim(Me.txtObjetivo), numero_consultoria_out, Me.adoMain.Recordset!Es_planilla, Me.adoMain.Recordset!Planilla_depto, Me.adoMain.Recordset!bco_codigo
                Else
                    DE.dbo_edGrabaConsultoria !ges_gestion, !org_codigo, !codigo_pago, !codigo_solicitud, !codigo_unidad, Val(Me.labNumero_consultoria), Me.cboCorrelativo.ItemData(Me.cboCorrelativo.ListIndex), "N", "", "", Trim(Left(Me.cboModalidadContratacion, 6)), !Codigo_convenio, Trim(Me.labPro_programa), Trim(Me.labPro_proyecto), Trim(Me.labPro_actividad), Trim(Me.txtdesConsultoria), "", Trim(Left(Me.cboModalidadSeleccion, 6)), Trim(Me.txtJustifSeleccion), !duracion_estimada_tiempo, Me.dtpFEstInicioCons, GlUsuario, !formulario, Trim(Me.txtSolNoObjecion), Trim(Me.txtObjetivo), numero_consultoria_out, "", "", ""
                End If
                'aqui cargar lo de la tabla del g- por que no hay registros
                If numero_consultoria_out <> 0 Then
                    rsX.Open "select SL.*, s.nacional_extranjero from ao_solicitud_lista SL, ao_solicitud s where sl.ges_gestion=s.ges_gestion and sl.codigo_unidad= s.codigo_unidad and sl.codigo_solicitud=s.codigo_solicitud and s.ges_gestion='" & Me.adoMain.Recordset!ges_gestion & "' and s.codigo_unidad='" & Me.adoMain.Recordset!codigo_unidad & "' and s.codigo_solicitud=" & Me.adoMain.Recordset!codigo_solicitud & " and s.formulario='F05' order by SL.id_beneficiario", db, adOpenStatic, adLockReadOnly
                    Do While Not rsX.EOF
'                        MsgBox Me.opFiles(0)
                        If Me.opFiles(0).Value = True Then
                            If rsX.Bookmark = 1 Then
                                DE.dbo_edGenCorrCons numero_consultoria_out, xcorrelativo_consultoria
                            Else
                                xcorrelativo_consultoria = 0
                            End If
                        Else
                            DE.dbo_edGenCorrCons numero_consultoria_out, xcorrelativo_consultoria
                        End If
                        If IsNull(rsX!apellido_esposo) Then
                            APE_ESPOSO = "-"
                        Else
                            APE_ESPOSO = rsX!apellido_esposo
                        End If
                        If IsNull(rsX!Monto_solicitud_dl) Then
                            Monto_solicitud_dl = 0
                        Else
                            Monto_solicitud_dl = rsX!Monto_solicitud_dl
                        End If
                        
                        If IsNull(rsX!Nro_pagos) Then
                            Nro_pagos = 0
                        Else
                            Nro_pagos = rsX!Nro_pagos
                        End If
                        If IsNull(rsX!aunidad) Then
                            aunidad = "-"
                        Else
                            aunidad = rsX!aunidad
                        End If
                        
                        If IsNull(rsX!aplanilla) Then
                            aplanilla = 0
                        Else
                            aplanilla = rsX!aplanilla
                        End If
                        If IsNull(rsX!tipo_documento) Or (rsX!tipo_documento) = "" Then
                            tipo_documento = "C"
                        Else
                            tipo_documento = rsX!tipo_documento
                        End If
                        DE.dbo_edGrabaNoObjDetalle numero_consultoria_out, 0, tipo_documento, Left(rsX!doc_identidad, 15), rsX!paterno, rsX!materno, rsX!NombreS, APE_ESPOSO, rsX!nacionalidad, rsX!Telefono, rsX!grado_instruccion, rsX!profesion, rsX!ciudad_postula, GlUsuario, Monto_solicitud_dl, Nro_pagos, rsX!nacional_Extranjero, aunidad, aplanilla, xcorrelativo_consultoria
                        rsX.MoveNext
                    Loop
                End If
            End If
            .Close
        End With
        Me.dgMain.Enabled = True
        fraOpciones0.Visible = False
        FraOpciones1.Visible = True
        FraOpciones2.Visible = False
        dgMain.Enabled = True
        Call BloqueaTodosLosCamposEditables
        ''''''Call refrescaMainList
        Call Me.refrescaDetalleConsultoria
    End If
Case 2
    'GRABA CITES
    Dim numero_consultoria_cite_out As Integer
    Dim rsZ As New ADODB.Recordset
    numero_consultoria_cite_out = 0
    If TodoBienConsultoriaCite() Then
        fracites.Enabled = True
        'aqui graba el registro de la consultoria
        Dim fechaCite As String
        If IsNull(Me.dtpFCite) Then
            fechaCite = ""
        Else
            fechaCite = CStr(Format(Me.dtpFCite, "dd/mm/yyyy"))
        End If
        If Adicionar = True Then
            DE.dbo_edGrabaCite Val(Me.labNumero_consultoria), 0, Trim(Me.txtcite), fechaCite, filtro, Val(Mid(Me.cboSelCite, 2, 4)), Me.labOrg_Codigo, IIf(Me.opNoObj(0).Value = True, "S", "N"), GlUsuario, numero_consultoria_cite_out
        Else
            DE.dbo_edGrabaCite Me.adoCite.Recordset!numero_consultoria, Me.adoCite.Recordset!numero_consultoria_cite, Trim(Me.txtcite), fechaCite, filtro, Val(Mid(Me.cboSelCite, 2, 4)), Me.labOrg_Codigo, IIf(Me.opNoObj(0).Value = True, "S", "N"), GlUsuario, numero_consultoria_cite_out
        End If
        'aqui graba los registros de la documentacion adjunta
        DE.dbo_apGeneralSearching "delete from ao_no_objecion_motivo_c where numero_consultoria =" & Val(Me.labNumero_consultoria) & " and numero_consultoria_cite=" & numero_consultoria_cite_out
        DE.dbo_apGeneralSearching "delete from ao_cite_doc_adjunta_c where numero_consultoria =" & Val(Me.labNumero_consultoria) & " and numero_consultoria_cite=" & numero_consultoria_cite_out
        If numero_consultoria_cite_out <> 0 Then
            For i = 0 To Me.lstMotivos.ListCount - 1
                If Me.lstMotivos.Selected(i) Then
                    DE.dbo_edGrabaCitemotivos Val(Me.labNumero_consultoria), numero_consultoria_cite_out, Left(Me.lstMotivos.List(i), 2), GlUsuario
                End If
            Next
            For i = 0 To Me.lstDocAdjunta.ListCount - 1
                If Me.lstDocAdjunta.Selected(i) Then
                    DE.dbo_edGrabaCiteDocAdjunta Val(Me.labNumero_consultoria), numero_consultoria_cite_out, Left(Me.lstDocAdjunta.List(i), 2), GlUsuario
                End If
            Next
        End If
        fraOpciones0.Visible = False
        FraOpciones1.Visible = True
        FraOpciones2.Visible = False
        'si es entrada de cite no objecion de contratacion o a TRF entonces genera la lista en adjudicacion
        
        dgMain.Enabled = True
        Me.fraAdiCite.Visible = False
        Call BloqueaTodosLosCamposEditables
        Call Me.refrescaDetalleCites
    End If
End Select
End Sub

Function TodoBienConsultoriaCite() As Boolean
TodoBienConsultoriaCite = True
If Me.optenviados(1).Value = True Then
    If Me.cboSelCite.ListIndex < 0 Then
        TodoBienConsultoriaCite = False
        MsgBox "Debe seleccionar el cite al que se registra respuesta", vbCritical, "Atencion"
    End If
End If
End Function


Function TodoBienConsultoria() As Boolean
'valida que todos los datos de la consultoria esten Ok, devuelve TRUE si es así
TodoBienConsultoria = False
If Me.cboCorrelativo.ListIndex < 0 Then
    MsgBox "Debe seleccionar un correlativo", vbInformation
    Me.cboCorrelativo.SetFocus
ElseIf Me.cboModalidadContratacion.ListIndex < 0 Then
    MsgBox "Debe seleccionar una Modalidad de Contratación", vbInformation
    Me.cboModalidadContratacion.SetFocus
ElseIf Me.cboModalidadSeleccion.ListIndex < 0 Then
    MsgBox "Debe seleccionar una Modalidad de Selección", vbInformation
    Me.cboModalidadSeleccion.SetFocus
ElseIf Len(Trim(Me.txtdesConsultoria)) = 0 Then
    MsgBox "Debe indicar la Descripción de la consultoría", vbInformation
    Me.txtdesConsultoria.SetFocus
ElseIf Len(Trim(Me.txtSolNoObjecion)) = 0 Then
    MsgBox "Debe indicar la Solicitud de No Objeción", vbInformation
    Me.txtSolNoObjecion.SetFocus
ElseIf Len(Trim(Me.txtObjetivo)) = 0 Then
    MsgBox "Debe indicar el Objetivo de la consultoría", vbInformation
    Me.txtObjetivo.SetFocus
ElseIf Len(Trim(Me.txtJustifSeleccion)) = 0 Then
    MsgBox "Debe indicar la Justificacion de la selección", vbInformation
    Me.txtJustifSeleccion.SetFocus
Else
    TodoBienConsultoria = True
End If
End Function
Sub CargaListaMotivos()
'carga la lista de motivos de la no objecion
    Dim rs As New ADODB.Recordset
    Dim i%
    'carga lista de Motivos
    rs.Open "select * from ac_no_motivo_c", db, adOpenStatic, adLockReadOnly
    lstMotivos.Clear
    For i = 1 To rs.RecordCount
        lstMotivos.AddItem Left(rs!codigo_motivo & "  ", 2) & " - " & rs!denominacion_motivo
        rs.MoveNext
    Next i
End Sub

Sub CargaListaDocAdjunta()
'carga la lista de documentacion adjunta a la no objecion
    Dim rs As New ADODB.Recordset
    Dim i%
    'carga lista de Documentacion Adjunta
    rs.Open "select * from ac_doc_adjunta_c", db, adOpenStatic, adLockReadOnly
    Me.lstDocAdjunta.Clear
    For i = 1 To rs.RecordCount
        lstDocAdjunta.AddItem Left(rs!codigo_doc_adjunta & "  ", 2) & " - " & rs!denominacion_doc_adjunta
        rs.MoveNext
    Next i
End Sub

Sub refrescaDetalleComun()
'refresca datos comunes a los tres tabs
Dim rs1 As New ADODB.Recordset
If Me.adoMain.Recordset.RecordCount > 0 Then
    rs1.Open "select numero_consultoria from ao_no_objecion_c where ges_Gestion='" & Me.adoMain.Recordset!ges_gestion & "' and codigo_unidad='" & Me.adoMain.Recordset!codigo_unidad & "' and codigo_solicitud=" & Me.adoMain.Recordset!codigo_solicitud, db, adOpenStatic, adLockReadOnly
    If rs1.RecordCount > 0 Then
        Me.labNumero_consultoria = rs1!numero_consultoria
    Else
        Me.labNumero_consultoria = 0
    End If
End If
End Sub


Sub refrescaDetalleSolicitud()
'refresca detalle de la solicitud de contratacion
If Me.adoMain.Recordset.RecordCount > 0 Then
    With Me.adoMain.Recordset
        If DE.rsdbo_apGeneralSearching.State = 1 Then DE.rsdbo_apGeneralSearching.Close
        DE.dbo_apGeneralSearching "select * from edVwSolicitudDetalle where ges_gestion='" & !ges_gestion & "' and codigo_unidad='" & !codigo_unidad & "' and codigo_solicitud=" & !codigo_solicitud
    End With
'    MsgBox DE.rsdbo_apGeneralSearching!CODIGO_POA
    Set Me.adoDetalleSolicitud.Recordset = DE.rsdbo_apGeneralSearching.Clone
    DE.rsdbo_apGeneralSearching.Close
End If
End Sub


Sub refrescaDetalleConsultoria()
'refresca el detalle de la consultoria
Dim i%
If Me.adoMain.Recordset.RecordCount > 0 Then
    If DE.rsdbo_apGeneralSearching.State = 1 Then DE.rsdbo_apGeneralSearching.Close
'    MsgBox Me.adoMain.Recordset!NUMERO_CONSULTORIA
    DE.dbo_edSacaDetalleConsultoria Me.adoMain.Recordset!ges_gestion, Me.adoMain.Recordset!codigo_unidad, Me.adoMain.Recordset!codigo_solicitud
    With DE.rsdbo_edSacaDetalleConsultoria
        If .RecordCount > 0 Then
            Me.labNumero_consultoria = !numero_consultoria
'''''            Me.labCorrelativo_Consultoria = !CORRELATIVO_CONSULTORIA
            Me.txtdesConsultoria = !descripcion_de_la_consultoria
            Me.txtSolNoObjecion = !solnoobjecioncons
            Me.txtObjetivo = !ObjetivoCons
            Me.labParaPublicacion = !parapublicacion
            For i = 0 To Me.cboCorrelativo.ListCount - 1
                Me.cboCorrelativo.ListIndex = i
                If Me.cboCorrelativo.ItemData(i) = !id_correlativo Then Exit For
            Next
            Me.cboModalidadContratacion = Left(!cod_mod_contratacion & Space(6), 6) & "- " & !des_mod_contratacion
            Me.cboModalidadSeleccion = Left(!cod_mod_seleccion & Space(6), 6) & "- " & !des_mod_seleccion
            Me.txtJustifSeleccion = !justif_seleccion
            'Me.txtTiempoMesesCons = !duracion_estimada_numero
            'Me.txtTiempoDiasCons = !por_tiempo
            Me.txtDuracion_Estimada_Tiempo = !duracion_estimada_tiempo
            Me.dtpFEstInicioCons = IIf(IsNull(!fecha_estimada_inicio), Null, !fecha_estimada_inicio)
        Else
            Me.labNumero_consultoria = ""
'''''''            Me.labCorrelativo_Consultoria = "NUEVO"
            Me.txtdesConsultoria = ""
            Me.txtSolNoObjecion = ""
            Me.txtObjetivo = ""
            Me.labParaPublicacion = "N"
            'Me.cboCorrelativo = ""
            'Me.cboModalidadContratacion = ""
            'Me.cboModalidadSeleccion = ""
            Me.txtJustifSeleccion = ""
            'Me.txtTiempoMesesCons = 0
            'Me.txtTiempoDiasCons = 0
            Me.txtDuracion_Estimada_Tiempo = ""
            Me.dtpFEstInicioCons = Date
        End If
        .Close
    End With
End If
Call RefrescaListaCorta
Call CuidaBotones
End Sub


Sub RefrescaListaCorta()
''refresca la listacorta de personas adjuntas a la solicitud de contraTACION
'With Me.adoMain.Recordset
'If DE.rsdbo_apGeneralSearching.State = 1 Then DE.rsdbo_apGeneralSearching.Close
'DE.dbo_apGeneralSearching "select * from edVwListaPersonasConsultoria where numero_consultoria=" & Val(Me.labNumero_consultoria)
'End With
'Set Me.adoListaCorta.Recordset = DE.rsdbo_apGeneralSearching.Clone
'DE.rsdbo_apGeneralSearching.Close
End Sub


Sub refrescaDetalleOrgFinanciador()
'refresca datos del organismo financiador
    Dim rs As New ADODB.Recordset
    'carga el nombre del organismo financiador a donde se dirigirá el cite
    rs.Open "select no.org_codigo,ofi.* from ao_no_objecion_c no, fc_organismo_financiamiento ofi where no.numero_consultoria=" & Val(Me.labNumero_consultoria) & " and no.org_codigo=ofi.org_codigo", db, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        Me.labOrg_Codigo = rs!org_codigo
        Me.labOrg_descripcion = rs!org_descripcion
        Me.labOrg_representante = rs!org_representante
        Me.labOrg_cargo = rs!org_cargo
    Else
        Me.labOrg_Codigo = ""
        Me.labOrg_descripcion = ""
        Me.labOrg_representante = ""
        Me.labOrg_cargo = ""
    End If
End Sub


Sub refrescaDetalleCites()
'refresca detalle de los cites enviados o recibidos
'DE.dbo_edListaCites Val(Me.labNumero_consultoria), filtro
'With DE.rsdbo_edListaCites
'    Set Me.adoCite.Recordset = .Clone
'    .Close
'End With
'Call CuidaBotones
End Sub


Private Sub CmdAñadir_Click()
'prepara los campos para añadir consultoria en casE=1
'prepara los campos para añadir un cite en case=2
Dim i%
fraOpciones0.Visible = False
FraOpciones1.Visible = False
FraOpciones2.Visible = True
Me.dgMain.Enabled = False
Adicionar = True
Select Case ssTab.Tab
Case 0
Case 1
    'habilita campos para añadir consultoria
    Me.cboCorrelativo.Enabled = True
    Me.cboModalidadContratacion.Enabled = True
    Me.cboModalidadSeleccion.Enabled = True
    Me.txtdesConsultoria.Enabled = True
    Me.txtSolNoObjecion.Enabled = True
    Me.txtObjetivo.Enabled = True
    Me.txtJustifSeleccion.Enabled = True
    Me.Frame7.Visible = True
    Me.cboModalidadContratacion.SetFocus
    DE.dbo_apGeneralSearching "select fecha_estimada_inicio, justificacion, Duracion_Estimada_Tiempo from pagos_espera pe where formulario='F05' and ges_gestion='" & Me.adoMain.Recordset!ges_gestion & "' and codigo_solicitud=" & Me.adoMain.Recordset!codigo_solicitud & " and codigo_unidad = '" & Me.adoMain.Recordset!codigo_unidad & "'"
    Me.dtpFEstInicioCons.Enabled = True
'    If IsNull(De.rsdbo_apGeneralSearching!fecha_estimada_inicio) Then
'        Me.dtpFEstInicioCons = Date
'    Else
'        Me.dtpFEstInicioCons = De.rsdbo_apGeneralSearching!fecha_estimada_inicio
'    End If
    Me.dtpFEstInicioCons = IIf(IsNull(DE.rsdbo_apGeneralSearching!fecha_estimada_inicio), Date, DE.rsdbo_apGeneralSearching!fecha_estimada_inicio)
    Me.txtdesConsultoria = IIf(IsNull(DE.rsdbo_apGeneralSearching!justificacion), "-", DE.rsdbo_apGeneralSearching!justificacion)
    txtJustifSeleccion = IIf(IsNull(DE.rsdbo_apGeneralSearching!justificacion), "-", DE.rsdbo_apGeneralSearching!justificacion)
    Me.txtDuracion_Estimada_Tiempo = IIf(IsNull(DE.rsdbo_apGeneralSearching!duracion_estimada_tiempo), "-", DE.rsdbo_apGeneralSearching!duracion_estimada_tiempo)
    DE.rsdbo_apGeneralSearching.Close
Case 2
    'habilita campos para añadir cite a la consultoria
    Call CargaListaMotivos
    Call CargaListaDocAdjunta
    Call refrescaDetalleOrgFinanciador
    fraCitesMuestra.Visible = False
    fraAdiCite.Visible = True
    fracites.Enabled = False
    Me.txtcite = ""
    Me.dtpFCite = Date
    Me.txtcite.SetFocus
    Me.txtcite.SelStart = 0
    Me.txtcite.SelLength = Len(Me.txtcite)
    If Me.optenviados(1).Value = True Then
        Me.cboSelCite.Clear
        DE.dbo_edGeneralSearching "select * from ao_no_objecion_cite_c where numero_consultoria=" & Val(Me.labNumero_consultoria) & " and estado_envio_recepcion='E' and enviado='S' order by numero_consultoria_cite"
        With DE.rsdbo_edGeneralSearching
            If .RecordCount > 0 Then
                Do While Not .EOF
                    Me.cboSelCite.AddItem "[" & Right("0000" & CStr(!numero_consultoria_cite), 4) & "][" & Left(!nro_cite & Space(25), 25) & "] de fecha " & Format(CStr(!fecha_cite), "dd/mm/yyyy")
                    .MoveNext
                Loop
            Else
                MsgBox "No puede registrar un cite como RECIBIDO por que no tiene cites ENVIADOS", vbCritical
                 Call cmdCancelar_Click
            End If
            .Close
        End With
        Me.lstDocAdjunta.Enabled = False
        Me.lstMotivos.Enabled = False
        Me.labPrefijo = ""
    Else
        Me.lstDocAdjunta.Enabled = True
        Me.lstMotivos.Enabled = True
        Me.labPrefijo = "Planalto/ CONSULT/"
    End If
End Select

End Sub


Private Sub CmdImprimir_Click()
'imprime la no objecion
Dim IResult As Variant, i%, Files$
CR.Formulas(0) = "codigos_poa='" & fSacaPOAs(Me.adoCite.Recordset!numero_consultoria) & "'"
CR.Formulas(1) = "docadj01=''"
CR.Formulas(2) = "docadj02=''"
CR.Formulas(3) = "docadj03=''"
CR.Formulas(4) = "docadj04=''"
CR.Formulas(5) = "docadj05=''"
If Me.adoDocAdjunta.Recordset.RecordCount > 0 Then
    Me.adoDocAdjunta.Recordset.MoveFirst
    For i = 1 To Me.adoDocAdjunta.Recordset.RecordCount
        If i > 5 Then Exit For
        CR.Formulas(i) = "docadj0" & Trim(CStr(i)) & "='" & Trim(Me.adoDocAdjunta.Recordset!denominacion_doc_adjunta) & "'"
        Me.adoDocAdjunta.Recordset.MoveNext
    Next
End If
DE.dbo_apGeneralSearching "select nombre= rtrim(nod.paterno) + ' ' + rtrim(nod.materno) + ' ' + rtrim(nod.nombres) from ao_no_objecion_detalle_c nod where numero_consultoria=" & Me.adoCite.Recordset!numero_consultoria
If DE.rsdbo_apGeneralSearching.RecordCount > 0 Then
    If DE.rsdbo_apGeneralSearching.RecordCount = 1 Then
       CR.Formulas(8) = "nombre ='" & DE.rsdbo_apGeneralSearching!Nombre & "'"
    Else
        CR.Formulas(8) = "nombre ='ver planilla adjunta'"
    End If
End If

DE.rsdbo_apGeneralSearching.Close
CR.Formulas(6) = "fecha ='La Paz, " & meses(Month(Me.adoCite.Recordset!fecha_cite)) & " " & CStr(Day(Me.adoCite.Recordset!fecha_cite)) & " del " & CStr(Year(Me.adoCite.Recordset!fecha_cite)) & "'"
CR.Formulas(7) = "nro_cite ='Planalto / CONSULT/" & Trim(Me.adoCite.Recordset!nro_cite) & "'"
DE.dbo_edNoObjFiles Val(Me.adoCite.Recordset!numero_consultoria), Files
CR.Formulas(9) = "files='" & Files & "'"
CR.StoredProcParam(0) = Val(Me.adoCite.Recordset!numero_consultoria)
CR.ReportFileName = App.Path & "\Consultoria RRHH\repNoObjecion1.rpt"
IResult = CR.PrintReport
If IResult <> 0 Then MsgBox CR.LastErrorNumber & " : " & CR.LastErrorString, vbCritical, "Error de impresión"
End Sub


Private Sub CmdModificar_Click()
'habilita campos para modificar o eliminar
'en el case =1 para consultoria
'en el case =2 para el cite
Dim i%
fraOpciones0.Visible = False
FraOpciones1.Visible = False
FraOpciones2.Visible = True
Me.dgMain.Enabled = False
Adicionar = False
Select Case Me.ssTab.Tab
Case 0
Case 1
    'habilita campos para modificar consultoria
    Me.cboModalidadContratacion.Enabled = True
    Me.cboModalidadSeleccion.Enabled = True
    Me.txtdesConsultoria.Enabled = True
    Me.txtSolNoObjecion.Enabled = True
    Me.txtObjetivo.Enabled = True
    Me.txtJustifSeleccion.Enabled = True
    Me.cboModalidadContratacion.SetFocus
    Me.dtpFEstInicioCons.Enabled = True
Case 2
    'habilita campos para modificar el cite
    Call CargaListaMotivos
    With Me.adoMotivos.Recordset
        If .RecordCount > 0 Then
            For i = 0 To Me.lstMotivos.ListCount - 1
                .MoveFirst
                .Find "codigo_motivo='" & Left(Me.lstMotivos.List(i), 2) & "'"
                If Not .EOF Then
                    Me.lstMotivos.Selected(i) = True
                End If
            Next i
        End If
    End With
    Call CargaListaDocAdjunta
    With Me.adoDocAdjunta.Recordset
        If .RecordCount > 0 Then
            For i = 0 To Me.lstDocAdjunta.ListCount - 1
                .MoveFirst
                .Find "codigo_doc_adjunta='" & Left(Me.lstDocAdjunta.List(i), 2) & "'"
                If Not .EOF Then
                    Me.lstDocAdjunta.Selected(i) = True
                End If
            Next i
        End If
    End With
    Call refrescaDetalleOrgFinanciador
    Me.txtcite = Me.adoCite.Recordset!nro_cite
    Me.dtpFCite = IIf(IsNull(Me.adoCite.Recordset!fecha_cite), Null, Me.adoCite.Recordset!fecha_cite)
    If Me.adoCite.Recordset!AceptaronNoObjecion = "S" Then
        Me.opNoObj(0).Value = True
    Else
        Me.opNoObj(1).Value = True
    End If
    fraCitesMuestra.Visible = False
    fraAdiCite.Visible = True
End Select
End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub cmdSalir1_Click()
Unload Me
End Sub


Private Sub Form_Load()
If glProceso = "CONSULTORIA" Then
    Me.Caption = "Procesos Administrativos  - Contratación de  Consultores - Registro Inicial"
Else
    Me.Caption = "Recursos Humanos - Solicitud de No Objeción"
End If
Dim rs As New ADODB.Recordset
'carga combo de seleccion de correlativos
Dim rsc As New ADODB.Recordset
rsc.Open "select * from ac_correlativos_consultoria_c WHERE ID_CORrELATIVO<>0", db, adOpenStatic, adLockReadOnly
Me.cboCorrelativo.Clear
Do While Not rsc.EOF
    Me.cboCorrelativo.AddItem rsc!descripcion
    Me.cboCorrelativo.ItemData(Me.cboCorrelativo.NewIndex) = rsc!id_correlativo
    rsc.MoveNext
Loop
rsc.Close

'carga combo modalidad contratacion
Me.cboModalidadContratacion.Clear
rs.Open "select * from ac_modalidad_contratacion order by cod_mod_contratacion", db, adOpenStatic, adLockReadOnly
With rs
    If .RecordCount > 0 Then
        Do While Not .EOF
            Me.cboModalidadContratacion.AddItem Left(!cod_mod_contratacion & Space(6), 6) & "- " & !des_mod_contratacion
            .MoveNext
        Loop
    End If
End With
rs.Close
'carga combo modalidad seleccion
Me.cboModalidadSeleccion.Clear
rs.Open "select * from ac_modalidad_seleccion  WHERE ACTIVO = 'S' order by cod_mod_seleccion", db, adOpenStatic, adLockReadOnly
With rs
    If .RecordCount > 0 Then
        Do While Not .EOF
            Me.cboModalidadSeleccion.AddItem Left(!cod_mod_seleccion & Space(6), 6) & "- " & !des_mod_seleccion
            .MoveNext
        Loop
    End If
End With
'----------
Call RefrescaMainList
Me.ssTab.Tab = 0
Me.fraOpciones0.Visible = True
Me.FraOpciones1.Visible = False
Me.FraOpciones2.Visible = False
End Sub


Sub RefrescaMainList()
'refresca la lista principal del mosulo de consultoria
If glProceso = "CONSULTORIA" Then
    DE.dbo_apGeneralSearching "select no.numero_consultoria, noc.enviado, s.*, ue.uni_descripcion_larga " & _
       "from ao_solicitud s left outer join ao_no_objecion_c no on s.ges_gestion=no.ges_gestion and s.codigo_unidad=no.codigo_unidad and s.codigo_solicitud=no.codigo_solicitud left outer join ao_no_objecion_cite_c noc on no.numero_consultoria=noc.numero_consultoria and noc.estado_envio_recepcion='E', fc_unidad_ejecutora ue where s.estatus='S' AND s.formulario='F05' and s.codigo_unidad=ue.uni_codigo AND S.espararh='N' order by s.codigo_unidad, s.ges_gestion, s.codigo_solicitud"
Else    'RRHH
    DE.dbo_apGeneralSearching "select no.numero_consultoria, noc.enviado, s.*, ue.uni_descripcion_larga " & _
       "from ao_solicitud s left outer join ao_no_objecion_c no on s.ges_gestion=no.ges_gestion and s.codigo_unidad=no.codigo_unidad and s.codigo_solicitud=no.codigo_solicitud left outer join ao_no_objecion_cite_c noc on no.numero_consultoria=noc.numero_consultoria and noc.estado_envio_recepcion='E', fc_unidad_ejecutora ue where s.estatus='S' AND (s.formulario='F05' OR s.formulario='F10')  and s.codigo_unidad=ue.uni_codigo AND S.espararh='S' order by s.codigo_unidad, s.ges_gestion, s.codigo_solicitud"
End If
Set adoMain.Recordset = DE.rsdbo_apGeneralSearching.Clone
If DE.rsdbo_apGeneralSearching.State = 1 Then DE.rsdbo_apGeneralSearching.Close
If Me.adoMain.Recordset.RecordCount > 0 Then
    Me.ssTab.Enabled = True
    Call CuidaBotones
Else
    'Me.ssTab.Enabled = False
End If
End Sub


Private Sub optenviados_Click(Index As Integer)
'cada vez que seleccione cites ENVIADOS o RECIBIDOS
Select Case Index
Case 0
    'cites enviados
    filtro = "E"
    If Val(Me.labNumero_consultoria) <> 0 Then
        Me.CmdAñadir.Enabled = True
    Else
        Me.CmdAñadir.Enabled = False
    End If
    Me.fraRespuestaDeRecibidos.Visible = False
    Me.CmdEnviar.Caption = "Confirma envío"
Case 1
    'cites recibifdos
    filtro = "R"
    If Val(Me.labNumero_consultoria) <> 0 Then
        Me.CmdAñadir.Enabled = True
    Else
        Me.CmdAñadir.Enabled = False
    End If
    Me.fraRespuestaDeRecibidos.Visible = True
    Me.CmdEnviar.Caption = "Confirma recepcion"
Case 2
    'todos los cites
    filtro = "T"
    Me.CmdAñadir.Enabled = False
End Select
Call refrescaDetalleCites
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
'cambio de tab
Select Case ssTab.Tab
Case 0
    'cuando cambia al tab de los datos de la relacion presupuestaria
    Me.fraOpciones0.Visible = True
    Me.FraOpciones1.Visible = False
    Me.FraOpciones2.Visible = False
    Call refrescaDetalleSolicitud
Case 1
    'cambia a los datos de la consultoiria
    Me.CmdImprimir.Enabled = False
    Me.fraOpciones0.Visible = False
    Me.FraOpciones1.Visible = True
    Me.FraOpciones2.Visible = False
    Call refrescaDetalleConsultoria
Case 2
    'cambia a los datos de cites enviados y recibidos
    Me.CmdImprimir.Enabled = True
    Me.fraOpciones0.Visible = False
    Me.FraOpciones1.Visible = True
    Me.FraOpciones2.Visible = False
    fraCitesMuestra.Visible = True
    fraAdiCite.Visible = False
    If Val(Me.labNumero_consultoria) = 0 Then
        Me.CmdAñadir.Enabled = False
    Else
        Me.CmdAñadir.Enabled = True
    End If
    Call refrescaDetalleCites
End Select
End Sub


Sub BloqueaTodosLosCamposEditables()
'bloquea todsos los campos editables
Me.cboCorrelativo.Enabled = False
Me.cboModalidadContratacion.Enabled = False
Me.cboModalidadSeleccion.Enabled = False
Me.txtdesConsultoria.Enabled = False
Me.txtSolNoObjecion.Enabled = False
Me.txtObjetivo.Enabled = False
Me.txtJustifSeleccion.Enabled = False
Me.dtpFEstInicioCons.Enabled = False

fraCitesMuestra.Visible = True
fraAdiCite.Visible = False
Me.Frame7.Visible = False
End Sub


Sub CuidaBotones()
'cuida botones
Select Case ssTab.Tab
Case 0
Case 1
    'cuida los botones para consultoria
    Me.cmdAPublicacion.Enabled = False
    If Val(Me.labNumero_consultoria) <> 0 Then
        Me.CmdAñadir.Enabled = False
        Me.CmdBorrar.Enabled = True
        Me.CmdModificar.Enabled = True
        If Me.labParaPublicacion = "N" Then
            Me.cmdAPublicacion.Enabled = True
        End If
        
    Else
        Me.CmdAñadir.Enabled = True
        Me.CmdBorrar.Enabled = False
        Me.CmdModificar.Enabled = False
    End If
Case 2
    'cuida los botones para cites
    Me.CmdAñadir.Enabled = False
    Me.CmdEnviar.Enabled = False
    Me.CmdImprimir.Enabled = False
    Me.CmdModificar.Enabled = False
    Me.CmdBorrar.Enabled = False
    If optenviados(0).Value = True Or optenviados(1).Value = True Then
        If Val(Me.labNumero_consultoria) <> 0 Then
            Me.CmdAñadir.Enabled = True
            
            If optenviados(0).Value = True Then
                If Me.adoCite.Recordset.RecordCount > 0 Then
                    If Me.adoCite.Recordset!enviado = "N" Then
                        'Me.cmdEnviar.Enabled = True
                    Else
                        Me.CmdImprimir.Enabled = True
                    End If
                End If
            End If
            If Me.adoCite.Recordset.RecordCount > 0 Then
                If Me.adoCite.Recordset!enviado = "N" Then
                    Me.CmdEnviar.Enabled = True
                    Me.CmdModificar.Enabled = True
                End If
                Me.CmdBorrar.Enabled = True
            End If
        End If
    End If
End Select
End Sub

Private Function fSacaPOAs(ncons As Integer) As String
'obtiene la lista de los POAs relacionados al detalle de la formulacion presupuestaria
Dim rs As New ADODB.Recordset
fSacaPOAs = ""
rs.Open "select DISTINCT pde.codigo_poa from pagos_Espera pe, pago_detalle_espera pde, ao_no_objecion_c no " & _
            " where no.numero_consultoria=" & ncons & " and pe.ges_Gestion=pde.ges_Gestion and pe.org_codigo=pde.org_codigo and pe.codigo_pago=pde.codigo_pago " & _
            " and no.ges_gestion= pe.ges_gestion and no.codigo_unidad=pe.codigo_unidad and " & _
            " no.codigo_solicitud=pe.codigo_solicitud", db, adOpenStatic, adLockReadOnly
Do While Not rs.EOF
    If fSacaPOAs <> Trim(rs!codigo_poa) Then
        fSacaPOAs = fSacaPOAs & " / " & Trim(rs!codigo_poa)
    Else
        fSacaPOAs = Trim(rs!codigo_poa)
    End If
    rs.MoveNext
Loop
fSacaPOAs = Right(fSacaPOAs, Len(fSacaPOAs) - 3)
End Function

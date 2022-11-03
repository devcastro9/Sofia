VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRc_Calendario 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - RRHH - Calendario Laboral"
   ClientHeight    =   7770
   ClientLeft      =   1065
   ClientTop       =   2415
   ClientWidth     =   10800
   Icon            =   "FrmRc_Calendario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmRc_Calendario.frx":0A02
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "FrmRc_Calendario.frx":6CA34
      ScaleHeight     =   960
      ScaleWidth      =   10560
      TabIndex        =   23
      Top             =   120
      Width           =   10620
      Begin VB.CommandButton cmdAdicionar 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "FrmRc_Calendario.frx":D8A66
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "FrmRc_Calendario.frx":D908A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdBorrar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "FrmRc_Calendario.frx":D966A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdIMPRIMIR 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "FrmRc_Calendario.frx":DA334
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmd_busqueda 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "FrmRc_Calendario.frx":DA8F1
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   1800
         Picture         =   "FrmRc_Calendario.frx":DAEA9
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   2640
         Picture         =   "FrmRc_Calendario.frx":DB0B3
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   720
         Left            =   3480
         Picture         =   "FrmRc_Calendario.frx":DB2BD
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   5160
         Picture         =   "FrmRc_Calendario.frx":DB9D8
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CALENDARIO LABORAL"
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
         Left            =   6120
         TabIndex        =   33
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.Frame FraEdicion 
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
      ForeColor       =   &H00008000&
      Height          =   6100
      Left            =   6240
      TabIndex        =   10
      Top             =   1080
      Width           =   4520
      Begin VB.TextBox txt03 
         DataField       =   "descripcion"
         DataSource      =   "adoLista"
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Text            =   "-"
         Top             =   4200
         Width           =   3495
      End
      Begin VB.ComboBox Txt01 
         DataField       =   "dia"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_Calendario.frx":DBBE2
         Left            =   240
         List            =   "FrmRc_Calendario.frx":DBBFB
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "LUNES"
         Top             =   2520
         Width           =   2055
      End
      Begin VB.ComboBox txt02 
         DataField       =   "tipo"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_Calendario.frx":DBC3B
         Left            =   240
         List            =   "FrmRc_Calendario.frx":DBC45
         TabIndex        =   12
         Text            =   "F"
         Top             =   3360
         Width           =   735
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_Calendario.frx":DBC4F
         Left            =   240
         List            =   "FrmRc_Calendario.frx":DBC68
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "2011"
         Top             =   840
         Width           =   900
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         DataField       =   "fecha"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   83361793
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "DIA LABORAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1080
         TabIndex        =   22
         Top             =   3360
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1420
         Width           =   615
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   3940
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Laboral/Feriado:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   3100
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion:"
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
         Index           =   45
         Left            =   240
         TabIndex        =   16
         Top             =   580
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Día de la Semana:"
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
         Index           =   40
         Left            =   240
         TabIndex        =   15
         Top             =   2260
         Width           =   1665
      End
   End
   Begin MSAdodcLib.Adodc adoLista 
      Height          =   330
      Left            =   120
      Top             =   6840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin MSAdodcLib.Adodc AdoTurno 
      Height          =   375
      Left            =   120
      Top             =   7320
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "AdoTurno"
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
   Begin MSDataGridLib.DataGrid grdlista 
      Bindings        =   "FrmRc_Calendario.frx":DBC96
      Height          =   5595
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   6100
      _ExtentX        =   10769
      _ExtentY        =   9869
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Ges_gestion"
         Caption         =   "Gestion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
         DataField       =   "dia"
         Caption         =   "Día"
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
         DataField       =   "tipo"
         Caption         =   "Lab/Fer"
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
         DataField       =   "descripcion"
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
      BeginProperty Column05 
         DataField       =   "fecha_registro"
         Caption         =   "fecha_registro"
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
         DataField       =   "hora_registro"
         Caption         =   "hora_registro"
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
         DataField       =   "usr_usuario"
         Caption         =   "usr_usuario"
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
            WrapText        =   -1  'True
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2340.284
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   10920
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picButtons 
      BackColor       =   &H00C0FFC0&
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdSalir99 
         Caption         =   "Cerrar"
         Height          =   480
         Left            =   9120
         Picture         =   "FrmRc_Calendario.frx":DBCAD
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir de Personas"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton CmdIMPRIMIR99 
         Caption         =   "Imprimir"
         Height          =   480
         Left            =   5040
         Picture         =   "FrmRc_Calendario.frx":DC6AF
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_busqueda99 
         Caption         =   "&Buscar"
         Height          =   480
         Left            =   4080
         Picture         =   "FrmRc_Calendario.frx":DD0B1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdAdicionar99 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         Picture         =   "FrmRc_Calendario.frx":DDAB3
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nuevo Registro"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdEditar99 
         Caption         =   "Modif."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1200
         Picture         =   "FrmRc_Calendario.frx":DE03D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh99 
         Caption         =   "Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3120
         Picture         =   "FrmRc_Calendario.frx":DE5C7
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprueba Registro"
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdBorrar99 
         Caption         =   "AnuLar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         Picture         =   "FrmRc_Calendario.frx":DEB51
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Anula Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdaceptar99 
         Caption         =   "Grabar"
         Height          =   480
         Left            =   6000
         Picture         =   "FrmRc_Calendario.frx":DF553
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar99 
         Caption         =   "Cancelar"
         Height          =   480
         Left            =   6960
         Picture         =   "FrmRc_Calendario.frx":DFADD
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmRc_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstorg As New ADODB.Recordset
Dim rs_turnos As New ADODB.Recordset
Dim rs_PARAMETRO As New ADODB.Recordset

Dim CAMPOS As ADODB.Field
'Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim sql_financiador As String
Dim sw2 As String

Private Sub Adolista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'   If pRecordset.EOF Or pRecordset.BOF Then
'      cmdEditar.Enabled = False
'      cmdBorrar.Enabled = False
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
'      Text4.Text = Empty
'      dtcfu.Text = ""
'      dtcfue.Text = ""
'      Exit Sub
'   End If

   If pRecordset.EOF Or pRecordset.BOF Then
   Else
     If pRecordset("tipo") = "F" Then
        LblDia.Caption = "FERIADO - DIA NO LABORAL"
        LblDia.ForeColor = &HC0&
     Else
        LblDia.Caption = "DIA LABORAL"
        LblDia.ForeColor = &H8000&
     End If
   End If
   
'   cmdEditar.Enabled = True
'   cmdBorrar.Enabled = True
'
'   Select Case pRecordset.EditMode
'      Case adEditInProgress
'      Case adEditNone
'         Txt01.Text = IIf(IsNull(pRecordset("Dia_control")), "", pRecordset("Dia_control"))
'         TxtGestion.Text = IIf(IsNull(pRecordset("ges_gestion")), "", pRecordset("ges_gestion"))
'         txt02.Text = IIf(IsNull(pRecordset("minutos_tolerancia")), "", pRecordset("minutos_tolerancia"))
'         txt03.Value = IIf(IsNull(pRecordset("primer_ingreso")), "", pRecordset("primer_ingreso"))
'         txt04.Value = IIf(IsNull(pRecordset("primera_salida")), "", pRecordset("primera_salida"))
'         txt05.Value = IIf(IsNull(pRecordset("tope_primer_ingreso")), "", pRecordset("tope_primer_ingreso"))
'         Txt06.Text = IIf(IsNull(pRecordset("Hora_extra_AM")), "", pRecordset("Hora_extra_AM"))
'         txt07.Value = IIf(IsNull(pRecordset("segundo_ingreso")), "", pRecordset("segundo_ingreso"))
'         txt08.Value = IIf(IsNull(pRecordset("segunda_salida")), "", pRecordset("segunda_salida"))
'         txt09.Value = IIf(IsNull(pRecordset("tope_segundo_ingreso")), "", pRecordset("tope_segundo_ingreso"))
'         txt10.Text = IIf(IsNull(pRecordset("Hora_extra_PM")), "", pRecordset("Hora_extra_PM"))
'         DTPFec_Inicio.Value = IIf(IsNull(pRecordset("fecha_registro")), "", pRecordset("fecha_registro"))
'         TxtHora.Text = IIf(IsNull(pRecordset("hora_registro")), "", pRecordset("hora_registro"))
''         TxtUsuario.Text = IIf(IsNull(pRecordset("usr_usuario")), "", pRecordset("usr_usuario"))
'      Case adEditDelete
'      Case adEditAdd
'   End Select
   adoLista.Caption = CStr(adoLista.Recordset.AbsolutePosition) & " de " & CStr(adoLista.Recordset.RecordCount)
End Sub
   
Private Sub cmdAceptar_Click()
On Error GoTo errorAceptar
Dim SW As Boolean
Dim SQL_FOR As String
Dim RSTORAUX As New ADODB.Recordset

   With adoLista
        If TxtGestion = "" Then
              MsgBox "INTRODUZCA DATOS"
              TxtGestion.SetFocus
              Exit Sub
        End If
        If Txt01 = "" Then
              MsgBox "INTRODUZCA DATOS"
              Txt01.SetFocus
              Exit Sub
        End If
        If Txt02 = "" Then
              MsgBox "INTRODUZCA DATOS"
              Txt02.SetFocus
              Exit Sub
        End If
        If txt03 = "" Then
              MsgBox "INTRODUZCA DATOS"
              txt03.SetFocus
              Exit Sub
        End If
'        If txt03 > txt04 Then
'              MsgBox "La Hora de INGRESO, NO puede ser mayor a la fecha de SALIDA.."
'              txt03.SetFocus
'              Exit Sub
'        End If
'        If DTPFec_Inicio > DTPFec_Fin Then
'              MsgBox "La fecha inicial de Control, NO puede ser mayor a la fecha Final de Control.."
'              DTPFec_Inicio.SetFocus
'              Exit Sub
'        End If
'        If (txt05 < txt03) Or (txt05 > txt04) Then
'              MsgBox "La fecha tope debe ser mayor a la de INGRESO y menor a la de SALIDA.."
'              txt05.SetFocus
'              Exit Sub
'        End If
'        If (txt09 < txt07) Or (txt09 > txt08) Then
'              MsgBox "La fecha tope debe ser mayor a la de INGRESO y menor a la de SALIDA.."
'              txt09.SetFocus
'              Exit Sub
'        End If
'
'    Set RSTORAUX = New ADODB.Recordset
'    SQL_FOR = "select * from Fc_ORGANISMO_FINANCIAMIENTO where ORG_CODIGO = '" & Text1.Text & "'"
'    RSTORAUX.Open SQL_FOR, DB, adOpenKeyset, adLockOptimistic, adCmdText
'    If RSTORAUX.RecordCount > 0 And Text1.Enabled Then
'      sw = True
'      MsgBox " CODIGO DUPLICADO"
'      Text1.SetFocus
'      Exit Sub
'    End If
    '
    'DB.BeginTrans
    SW = False
    If sw2 = "ADD" Then
'        .Recordset.AddNew
'        .Recordset("ges_gestion").Value = TxtGestion.Text
'        .Recordset("turno") = Dtc_Par.Text
    End If
    .Recordset("tipo").Value = Txt02.Text
    .Recordset("descripcion").Value = txt03.Text

    .Recordset("usr_usuario").Value = glusuario
    .Recordset("fecha_registro").Value = Date   'Format(Date, "dd/mm/aaaa")
    .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
    .Recordset.Update
    '.Recordset.Requery
    'DB.CommitTrans
            
   End With
   
   sw2 = "XX"
   cmdAdicionar.Visible = True
   cmdEditar.Visible = True
   cmdBorrar.Visible = True
   cmdRefresh.Visible = True
   cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   CmdSalir.Visible = True
   cmdaceptar.Visible = False
   cmdCancelar.Visible = False
   adoLista.Enabled = True
   grdlista.Enabled = True
   TxtGestion.Enabled = True
   FraEdicion.Enabled = False
   
   Call ABRIR_TABLA
   
   Exit Sub

errorAceptar:
   
   Call pErrorRst(db.Errors)
   
   adoLista.Recordset.CancelUpdate
   
   'DB.RollbackTrans
End Sub
 Private Sub Cmdadicionar_Click()
'   SW2 = "ADD"
'   cmdAdicionar.Visible = False
'   cmdEditar.Visible = False
'   cmdBorrar.Visible = False
'   cmdRefresh.Visible = False
'   cmd_busqueda.Visible = False
'   'CmdIMPRIMIR.Visible = False
'   cmdSalir.Visible = False
'
'   cmdaceptar.Visible = True
'   cmdCancelar.Visible = True
'   adoLista.Recordset.AddNew
'   adoLista.Enabled = False
'   grdlista.Enabled = False
'   'TxtGestion.Text = Empty
'   'Txt01.Text = Empty
''   Txt02.Text = Empty
''   Txt06.Text = Empty
''   txt10.Text = Empty
'Dim dateValue As Date = #6/11/2008#
'Console.WriteLine(dateValue.ToString("dddd", New CultureInfo("es-ES"))     ' Displays miércoles.

'Weekday(date, [firstdayofweek])
'•vbUseSystemDayOfWeek = 0 (el del sistema)
'•vbSunday = 1
'•vbMonday = 2
'•vbTuesday = 3
'•vbWednesday = 4
'•vbThursday = 5
'•vbFriday = 6
'•vbSaturday = 7
    
FraEdicion.Enabled = False
Dim dia2 As String
Dim fechita As Date
Dim mes2 As String
'dia = WeekdayName(Weekday(Date))
'MsgBox dia
Dim i As String
i = Year(Date)
Dim rsCalendar As New ADODB.Recordset
   Set rsCalendar = New ADODB.Recordset
   rsCalendar.Open "select * from gc_calendario WHERE ges_gestion = '" & i & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'   Set AdoPais.Recordset = rsCalendar
  If rsCalendar.RecordCount > 0 Then
    MsgBox "Ya fue creado el calendario de la Gestión: " + i, vbInformation + vbOKOnly, "Atención"
  Else
    fechita = CDate("01/01/" + i)
    While fechita >= CDate("01/01/" + i) And fechita <= CDate("31/12/" + i)
       rsCalendar.AddNew
       dia2 = WeekdayName(Weekday(fechita))
       rsCalendar!Fecha = fechita   'Format(fechita, "dd/mm/aaaa")
       rsCalendar!ges_gestion = Year(fechita)
       mes2 = MonthName(Month(fechita))
       rsCalendar!mes = UCase(mes2)
'       Select Case mes2
'            Case 2:                     Dias_Del_Mes = IIf(saltarYear(fecha), 29, 28)
'            Case 1, 3, 5, 7, 8, 10, 12: Dias_Del_Mes = 31
'            Case 4, 6, 9, 11:           Dias_Del_Mes = 30
'       End Select

       rsCalendar!dia = UCase(dia2)
       If dia2 = "sábado" Or dia2 = "domingo" Then
          rsCalendar!tipo = "F"
          rsCalendar!descripcion = "DIA FERIADO - NO LABORAL"
       Else
          rsCalendar!tipo = "L"
          rsCalendar!descripcion = "DIA LABORAL"
       End If
       rsCalendar.Update
       fechita = fechita + 1
    Wend
    MsgBox "Se ha creado con éxito el calendario de la Gestión: " + i, vbInformation + vbOKOnly, "Atención"
  End If
  Call ABRIR_TABLA
End Sub

'Private Sub Cmdborrar_Click()
'   Dim Mensaje As String
'
'On Error GoTo errorDelete
'
'   Mensaje = "¿Borrar: " & _
'               Text1.Text & " " & _
'               Trim(Text3.Text) & "?"
'   If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar:") = vbYes Then
'      db.BeginTrans
'      adoLista.Recordset.Delete
'      db.CommitTrans
'   End If
'
'   Exit Sub
'errorDelete:
'
'   Dim e As ADODB.Error
'
'   For Each e In db.Errors
'      MsgBox "Error No. " & e.Number & " " & e.Description
'   Next
'
'   db.RollbackTrans
'
'End Sub

Private Sub Cmd_Busqueda_Click()
''BUSQUEDA.Visible = True
''fradatos.Enabled = True
' Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = DB
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = grdlista
'    ClBuscaGrid.QueryUtilizado = sql_financiador
'    Set ClBuscaGrid.RecordsetTrabajo = adoLista.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar

End Sub

Private Sub cmdCancelar_Click()
  On Error Resume Next
   sw2 = "XX"
   cmdAdicionar.Visible = True
   cmdEditar.Visible = True
   cmdBorrar.Visible = True
   cmdRefresh.Visible = True
   cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   CmdSalir.Visible = True

   cmdaceptar.Visible = False
   cmdCancelar.Visible = False
   adoLista.Enabled = True
   grdlista.Enabled = True
   TxtGestion.Enabled = True
   FraEdicion.Enabled = False
   Call ABRIR_TABLA
End Sub

Private Sub cmdEditar_Click()
   sw2 = "MOD"
   cmdAdicionar.Visible = False
   cmdEditar.Visible = False
   cmdBorrar.Visible = False
   cmdRefresh.Visible = False
   cmd_busqueda.Visible = False
   'CmdIMPRIMIR.Visible = False
   CmdSalir.Visible = False

   cmdaceptar.Visible = True
   cmdCancelar.Visible = True
   adoLista.Enabled = False
   grdlista.Enabled = False
   FraEdicion.Enabled = True

   TxtGestion.Enabled = False
   Txt01.Enabled = False
   Txt02.SetFocus
End Sub

Private Sub CmdImprimir_Click()
  Dim iResult As Integer
     CrystalReport1.WindowShowPrintSetupBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
         CrystalReport1.ReportFileName = App.Path & "\Reportes\RRHH\rr_calendario_laboral.rpt"
  'CrystalReport1.ReportFileName = "\SAF-2000\Reportes\RRHH\rr_calendario_laboral.rpt"
  iResult = CrystalReport1.PrintReport
  If iResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

CrystalReport1.WindowState = crptMaximized
'REPORGFIN.Show

'   rptModalidadSeleccion.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
'  If adoLista.Recordset!turno <> "AM" And adoLista.Recordset!turno <> "PM" Then
'    If adoLista.Recordset!turno = "AM" Then
'        GlHora1 = adoLista.Recordset!tope_hora_ingreso
'    End If
'    If adoLista.Recordset!turno = "PM" Then
'        GlHora2 = adoLista.Recordset!tope_hora_ingreso
'    End If
'    Set rs_PARAMETRO = New ADODB.Recordset
'    rs_PARAMETRO.Open "select * from gc_parametros_sistema where estado_codigo = 'APR' ", cnn, adOpenDynamic, adLockReadOnly
'    If rs_PARAMETRO.RecordCount > 0 Then
'        'rs_PARAMETRO.MoveFirst
'        rs_PARAMETRO!Hora_Ingreso1 = GlHora1
'        rs_PARAMETRO!Hora_Ingreso2 = GlHora2
'    End If
'  End If
'  adoLista.Recordset!estado_codigo = "SI"

'Dim dateValue As Date = #6/11/2008#
'Console.WriteLine(dateValue.ToString("dddd", New CultureInfo("es-ES"))     ' Displays miércoles.

'Weekday(date, [firstdayofweek])
'•vbUseSystemDayOfWeek = 0 (el del sistema)
'•vbSunday = 1
'•vbMonday = 2
'•vbTuesday = 3
'•vbWednesday = 4
'•vbThursday = 5
'•vbFriday = 6
'•vbSaturday = 7

Dim dia As String
Dim fechita As Date
'dia = WeekdayName(Weekday(Date))
'MsgBox dia

Dim rsCalendar As New ADODB.Recordset
   Set rsCalendar = New ADODB.Recordset
   rsCalendar.Open "select * from gc_calendario ", db, adOpenKeyset, adLockOptimistic, adCmdText
'   Set AdoPais.Recordset = rsCalendar
Dim i As Integer
  i = 1
  fechita = CDate("01/01/2016")
  While fechita >= "01/01/2016" And fechita <= "31/12/2016"
     rsCalendar.AddNew
     dia = WeekdayName(Weekday(fechita))
     rsCalendar!Fecha = fechita 'Format(fechita, "dd/mm/aaaa")
     rsCalendar!ges_gestion = Year(fechita)
     If dia = "sábado" Or dia = "domingo" Then
        rsCalendar!tipo = "F"
     Else
        rsCalendar!tipo = "H"
     End If
     fechita = fechita + 1
  Wend
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

'Private Sub Dtc_Par_Click(Area As Integer)
''    Dtc_ParDes.BoundText = Dtc_Par.BoundText
'End Sub
'
'Private Sub Dtc_ParDes_Click(Area As Integer)
''    Dtc_Par.BoundText = Dtc_ParDes.BoundText
'End Sub

Private Sub Form_Load()
   'Dim sql_fuente As String
   sw2 = "XX"
   cmdAdicionar.Visible = True
   cmdEditar.Visible = True
   cmdBorrar.Visible = True
   cmdRefresh.Visible = True
   cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   CmdSalir.Visible = True

   cmdaceptar.Visible = False
   cmdCancelar.Visible = False
   adoLista.Enabled = True
   grdlista.Enabled = True
   FraEdicion.Enabled = False
   
   Call ABRIR_TABLA
   
'   Set rs_turnos = New ADODB.Recordset
'   'sql_fuente = "select * from rc_turnos" ' order by fte_codigo"
'   rs_turnos.Open "select * from rc_turnos", DB, adOpenKeyset, adLockOptimistic, adCmdText
'   rs_turnos.Sort = "correl"
''  ' MsgBox rstfue.RecordCount
'   Set AdoTurno.Recordset = rs_turnos
'
   
'   Set rstorg = New ADODB.Recordset
'   sql_financiador = "select * from rc_horarios" 'order by org_codigo"
'   rstorg.Open sql_financiador, DB, adOpenKeyset, adLockOptimistic, adCmdText
''   rstorg.Sort = "Dia_control"
'   Set adoLista.Recordset = rstorg
'   'Set ClBuscaGrid = Nothing
  
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
    Set rstorg = New ADODB.Recordset
    If rstorg.State = 1 Then rstorg.Close
    sql_financiador = "select * from gc_calendario "
    rstorg.Open sql_financiador, db, adOpenKeyset, adLockOptimistic, adCmdText
    rstorg.Sort = "fecha"
    Set adoLista.Recordset = rstorg
    Set grdlista.DataSource = adoLista.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (rstorg.State = adStateClosed) Then rstorg.Close
   'Set rstorg = Nothing
End Sub

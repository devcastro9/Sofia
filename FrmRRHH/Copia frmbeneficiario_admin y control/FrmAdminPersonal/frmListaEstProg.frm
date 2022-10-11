VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmListaEstProg 
   Caption         =   "Estructura Programatica"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   645
      Left            =   4995
      Picture         =   "frmListaEstProg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2940
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   645
      Left            =   5745
      Picture         =   "frmListaEstProg.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2940
      Width           =   765
   End
   Begin MSDataGridLib.DataGrid dGrEstProg 
      Bindings        =   "frmListaEstProg.frx":074C
      Height          =   2610
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   4604
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
      Caption         =   "Estructura Programatica"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Pro_programa"
         Caption         =   "Prog"
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
         DataField       =   "Pro_subprograma"
         Caption         =   "SubProg"
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
         DataField       =   "Pro_proyecto"
         Caption         =   "Proy"
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
         DataField       =   "Pro_actividad"
         Caption         =   "Actv."
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
         DataField       =   "Pro_descripcion_larga"
         Caption         =   "Descripcion"
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
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   5715.213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoEstrProg 
      Height          =   330
      Left            =   180
      Top             =   2805
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6059
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
      Caption         =   "Adodc1"
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
Attribute VB_Name = "frmListaEstProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As New Connection
Public tFc_fuente_financiamiento As New ADODB.Recordset
Public tFc_organismo_financiamiento As New ADODB.Recordset
Public tFc_convenios As New ADODB.Recordset
Public tFc_estructura_programatica As New ADODB.Recordset

Private Sub cmdAceptar_Click()
  GetEstructura
  Unload Me
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub GetEstructura()
  frmRepPresupuesto.txtProg.Text = dGrEstProg.Columns(0)
  frmRepPresupuesto.txtSubProg = dGrEstProg.Columns(1)
  frmRepPresupuesto.txtProy = dGrEstProg.Columns(2)
  frmRepPresupuesto.txtAct = dGrEstProg.Columns(3)
End Sub

Private Sub dGrEstProg_DblClick()
  GetEstructura
  Unload Me
End Sub

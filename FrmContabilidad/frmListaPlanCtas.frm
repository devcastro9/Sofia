VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmListaPlanCtas 
   Caption         =   "Elige una Cta Contable ..."
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   13005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Elegir"
      Height          =   645
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2700
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   645
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2700
      Width           =   765
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
   Begin MSDataGridLib.DataGrid dGrEstProg 
      Bindings        =   "frmListaPlanCtas.frx":0000
      Height          =   2610
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4604
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
      Caption         =   "PLAN DE CUENTAS"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Cuenta"
         Caption         =   "Cuenta"
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
         DataField       =   "SubCta1"
         Caption         =   "SubCta1"
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
         DataField       =   "SubCta2"
         Caption         =   "SubCta2"
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
         DataField       =   "Aux1"
         Caption         =   "Aux1"
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
         DataField       =   "Aux2"
         Caption         =   "Aux2"
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
         DataField       =   "Aux3"
         Caption         =   "Aux3"
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
         DataField       =   "NombreCta"
         Caption         =   "NombreCta"
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
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   5655.118
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListaPlanCtas"
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
  FrmConta_BalApertura.cbocta.Text = dGrEstProg.Columns(0)
  FrmConta_BalApertura.cbosubcta1 = dGrEstProg.Columns(1)
  FrmConta_BalApertura.cbosubcta2 = dGrEstProg.Columns(2)
  FrmConta_BalApertura.txtax1 = dGrEstProg.Columns(3)
  FrmConta_BalApertura.Txtax2 = dGrEstProg.Columns(4)
  FrmConta_BalApertura.txtax3 = dGrEstProg.Columns(5)
  FrmConta_BalApertura.DtcCtaNom = dGrEstProg.Columns(6)
End Sub

Private Sub dGrEstProg_DblClick()
  GetEstructura
  Unload Me
End Sub


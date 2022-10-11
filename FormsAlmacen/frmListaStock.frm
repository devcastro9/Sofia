VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmListaStock 
   Caption         =   "Stock de Productos por Almacen"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   10905
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   645
      Left            =   9075
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2820
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   645
      Left            =   10065
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2820
      Width           =   765
   End
   Begin MSDataGridLib.DataGrid dGrEstProg 
      Height          =   2610
      Left            =   15
      TabIndex        =   0
      Top             =   165
      Width           =   10860
      _ExtentX        =   19156
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
      Caption         =   "Stock de Productos"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "CodDestino"
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
         DataField       =   "DescDetalle"
         Caption         =   "Nombre del Producto"
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
         DataField       =   "CodDetalle"
         Caption         =   "Producto"
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
         DataField       =   "nro_licitacion"
         Caption         =   "Nro.Compra"
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
         DataField       =   "Nro_Lote"
         Caption         =   "Nro_Lote"
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
         DataField       =   "fechaVenc"
         Caption         =   "Fecha.Venc."
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
         DataField       =   "StockActual"
         Caption         =   "StockActual"
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
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4334.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoEstrProg 
      Height          =   330
      Left            =   60
      Top             =   2925
      Visible         =   0   'False
      Width           =   8595
      _ExtentX        =   15161
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=DIMOSUD;Data Source=SERVIDOR"
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=DIMOSUD;Data Source=SERVIDOR"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "av_Stock_Almacenes"
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
Attribute VB_Name = "frmListaStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As New Connection
Public tFc_fuente_financiamiento As New ADODB.Recordset
Public tFc_organismo_financiamiento As New ADODB.Recordset
Public tFc_convenios As New ADODB.Recordset
Public tFc_estructura_programatica As New ADODB.Recordset

Private Sub CmdAceptar_Click()
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



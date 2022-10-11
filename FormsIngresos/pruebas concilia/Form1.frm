VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   2715
      TabIndex        =   1
      Top             =   4485
      Width           =   1035
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   4185
      Left            =   165
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   0
      Top             =   270
      Width           =   2760
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   4515
      Width           =   2475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Data1.Recordset.MoveFirst
  While Not (Data1.Recordset.EOF)
    If IsNull(Data1.Recordset("fecha")) Then
      
    Else
      MsgBox Data1.Recordset("fecha")
    End If
    Data1.Recordset.MoveNext
  Wend
End Sub

Private Sub Form_Load()
Dim xxx As Recordset
  Data1.Connect = "Excel 8.0"
  Data1.DatabaseName = "c:\mis documentos\grecocon.xls"    ' "c:\mis documentos\grecocon.xls"
  Data1.RecordSource = "Hoja1$"
'  Data1.Database.Connection.Execute
  TDBGrid1.DataSource = Data1
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Data1.Connect = ""
  Data1.DatabaseName = ""
  Data1.RecordSource = ""
  Print Data1.DatabaseName
  Print Data1.RecordSource
  
End Sub


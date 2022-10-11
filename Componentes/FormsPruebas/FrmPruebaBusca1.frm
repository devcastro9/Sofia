VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn As ADODB.Connection
Dim ClBusca As CompBusquedas.ClVentanaBuscaEnQuery

Private Sub CmdBuscar_Click()
  Set ClBusca.Conexión = cnn
  ClBusca.QueryUtilizado = "SELECT * FROM Pagos"
  ClBusca.Ejecutar
End Sub

Private Sub Form_Load()
  Set cnn = New ADODB.Connection
  cnn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000;Data Source=SERSIS"
  
  Set ClBusca = New CompBusquedas.ClVentanaBuscaEnQuery
End Sub

Private Sub Form_Unload(Cancel As Integer)
  cnn.Close
  Set cnn = Nothing
  Set ClBusca = Nothing
End Sub

VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtValor 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   12
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox TxtValor 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox TxtValor 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuscarBinaria 
      Caption         =   "&Buscar Binaria"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CmdBuscarSecuencial 
      Caption         =   "&Buscar Secuencial"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton CmdBusca2 
      Caption         =   "&Buscar 2"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin TrueOleDBGrid60.TDBGrid TdbgBusca 
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "FrmPruebaBusca.frx":0000
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar 1"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valores:"
      Height          =   195
      Left            =   4920
      TabIndex        =   8
      Top             =   3240
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Campos:"
      Height          =   195
      Left            =   4920
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn As ADODB.Connection
Dim rsTabla As ADODB.Recordset
'JQA
'Dim ClBusca As CompBusquedas.ClBuscaEnGridPropio
'JQA
Dim ClBuscaEx As CompBusquedas.ClBuscaEnGridExterno
Dim ClBuscaSec As CompBusquedas.ClBuscaSecuencialEnRS
Dim ClBuscaBin As CompBusquedas.ClBuscaBinariaEnRS

Private Sub CmdBusca2_Click()
  Set ClBuscaEx.Conexi�n = cnn
  Set ClBuscaEx.RecordsetTrabajo = rsTabla
  Set ClBuscaEx.GridTrabajo = TdbgBusca
  ClBuscaEx.QueryUtilizado = "SELECT * FROM CO_Comprobante_m"
  ClBuscaEx.EsTdbGrid = True
  ClBuscaEx.Ejecutar
End Sub

Private Sub CmdBuscar_Click()
'JQA
'  Set ClBusca.Conexi�n = cnn
'  ClBusca.FiltrosMultiples = True
'  ClBusca.QueryUtilizado = "SELECT * FROM CO_Comprobante_m"
'  ClBusca.Ejecutar
'JQA
End Sub

Private Sub cmdBuscarBinaria_Click()
  Set ClBuscaBin.Recordset = rsTabla
  ClBuscaBin.Campo = TxtCampo(0).Text
  ClBuscaBin.ValorCampo = TxtValor(0)
  ClBuscaBin.Posicionar = True
  ClBuscaBin.Ejecutar
End Sub

Private Sub CmdBuscarSecuencial_Click()
  Set ClBuscaSec.Recordset = rsTabla
  ClBuscaSec.Campo1 = TxtCampo(0)
  ClBuscaSec.ValorCampo1 = TxtValor(0)
  ClBuscaSec.Campo2 = TxtCampo(1)
  ClBuscaSec.ValorCampo2 = TxtValor(1)
  ClBuscaSec.Campo3 = TxtCampo(2)
  ClBuscaSec.ValorCampo3 = TxtValor(2)
  ClBuscaSec.Ejecutar
End Sub

Private Sub Form_Load()
On Error GoTo QueError
  Set cnn = New ADODB.Connection
  cnn.CursorLocation = adUseClient
  cnn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;" & _
           "User ID=sa;Initial Catalog=QUEIROS;Data Source=SERSIS"
  Set rsTabla = New ADODB.Recordset
  rsTabla.Open "SELECT * FROM CO_Comprobante_m", cnn, adOpenStatic
  Set TdbgBusca.DataSource = rsTabla
'JQA
'  Set ClBusca = New CompBusquedas.ClBuscaEnGridPropio
'JQA
  Set ClBuscaEx = New CompBusquedas.ClBuscaEnGridExterno
  Set ClBuscaSec = New CompBusquedas.ClBuscaSecuencialEnRS
  Set ClBuscaBin = New CompBusquedas.ClBuscaBinariaEnRS
  Exit Sub
QueError:
  MsgBox "Error al Cargar: " & Err.Description, vbInformation + vbOKOnly, "Atenci�n"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsTabla.Close
  cnn.Close
  Set cnn = Nothing
  'JQA
'  Set ClBusca = Nothing
  'JQA
  Set ClBuscaEx = Nothing
  Set ClBuscaSec = Nothing
  Set ClBuscaBin = Nothing
End Sub


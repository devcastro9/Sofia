VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ALFrmMateriales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de Materiales"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList iml 
      Left            =   3015
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALFrmMateriales.frx":0000
            Key             =   "Raiz"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALFrmMateriales.frx":031A
            Key             =   "Cerrado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALFrmMateriales.frx":0BF4
            Key             =   "Abierto"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALFrmMateriales.frx":14CE
            Key             =   "Detalle"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trv 
      Height          =   4905
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   8652
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "iml"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2190
      TabIndex        =   2
      Top             =   5085
      Width           =   1305
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   270
      TabIndex        =   1
      Top             =   5085
      Width           =   1305
   End
End
Attribute VB_Name = "ALFrmMateriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public QResp As Boolean
Public QCodigo As String
Public QItem As String
Public QGrupo As String
Dim rsgrupo As ADODB.Recordset
Dim RsDet As ADODB.Recordset
Public Sub ALPrincipal()
    '--
    LlenaArbol
    '--
    Me.Show vbModal
End Sub
Private Sub LlenaArbol()
Dim Nodo As Node
    Set rsgrupo = New ADODB.Recordset
    Set RsDet = New ADODB.Recordset
    GlSqlAux = "SELECT CodGrupo, DescGrupo " & _
               "FROM ALCLGrupo " & _
               "WHERE   cast(codgrupo as int) >100 OR cast(codgrupo as int) <5" & _
               "ORDER BY cast(CodGrupo as int)"
    rsgrupo.Open GlSqlAux, DB, adOpenStatic
    If rsgrupo.RecordCount > 0 Then
        Set Nodo = trv.Nodes.Add(, , "M", "Lista de Materiales", "Raiz")
        Nodo.Bold = True
        Nodo.Expanded = True
        While Not rsgrupo.EOF
            Set Nodo = trv.Nodes.Add("M", tvwChild, "G" & rsgrupo!CodGrupo, rsgrupo!CodGrupo & " - " & rsgrupo!descgrupo, "Cerrado", "Abierto")
            rsgrupo.MoveNext
        Wend
        
        GlSqlAux = "SELECT CodGrupo, CodDetalle, DescDetalle " & _
                   "FROM ALCLDetalle " & _
                   "WHERE Estado = 1 and cast(codgrupo as int) >100 OR cast(codgrupo as int) <5" & _
                   "ORDER BY CodGrupo, CodDetalle"
                   '"WHERE Estado = 1 AND CodGrupo IN ()" & _'
        RsDet.Open GlSqlAux, DB, adOpenStatic
        While Not RsDet.EOF
            Set Nodo = trv.Nodes.Add("G" & RsDet!CodGrupo, tvwChild, "D" & RsDet!CodGrupo & "-" & RsDet!codDetalle, RsDet!CodGrupo & "-" & RsDet!codDetalle & " : " & RsDet!descdetalle, "Detalle")
            RsDet.MoveNext
        Wend
        Cmdaceptar.Enabled = True
    Else
        Set Nodo = trv.Nodes.Add(, , "M", "No Existe Lista de Materiales")
        Nodo.Bold = True
        Nodo.Expanded = True
        Cmdaceptar.Enabled = False
    End If
End Sub
Private Sub CmdAceptar_Click()
    If Trim(QCodigo) <> "" Then
        QResp = True
        Unload Me
    End If
End Sub
Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    QResp = False
End Sub
Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
    If InStr(Node.Key, "D") > 0 Then
        QCodigo = Mid(Node.Key, 2)
        QItem = Mid(Node.Text, InStr(Node.Text, ":") + 1)
    Else
        QCodigo = ""
        QItem = ""
    End If
End Sub

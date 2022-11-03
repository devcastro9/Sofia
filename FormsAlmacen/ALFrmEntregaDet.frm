VERSION 5.00
Begin VB.Form ALFrmEntregaDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrega Detalle"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   Icon            =   "ALFrmEntregaDet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fra 
      Height          =   1170
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   -60
      Width           =   4950
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   195
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   750
         Width           =   4245
      End
      Begin VB.CommandButton cmdElige 
         Caption         =   "..."
         Height          =   225
         Left            =   195
         TabIndex        =   0
         Top             =   150
         Width           =   255
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   195
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Item"
         Height          =   195
         Left            =   510
         TabIndex        =   12
         Top             =   180
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   420
      Left            =   3555
      TabIndex        =   4
      Top             =   2055
      Width           =   1320
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   90
      TabIndex        =   3
      Top             =   2055
      Width           =   1320
   End
   Begin VB.Frame Fra 
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   1155
      Width           =   4950
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Precio Total"
         Height          =   195
         Left            =   2700
         TabIndex        =   1
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Peso Total Caja(s)"
         Height          =   195
         Left            =   3480
         TabIndex        =   2
         Top             =   60
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Precio Unitario"
         Height          =   195
         Left            =   2550
         TabIndex        =   5
         Top             =   1590
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Peso Unitario Caja(s)"
         Height          =   195
         Left            =   1905
         TabIndex        =   6
         Top             =   60
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   1485
         TabIndex        =   7
         Top             =   1650
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   60
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Caja(s)"
         Height          =   195
         Left            =   1155
         TabIndex        =   14
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ejemp.(s)"
         Height          =   195
         Left            =   1440
         TabIndex        =   19
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Kgs."
         Height          =   195
         Left            =   2985
         TabIndex        =   18
         Top             =   375
         Width           =   315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "$us."
         Height          =   195
         Left            =   3495
         TabIndex        =   17
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Kgs."
         Height          =   195
         Left            =   4545
         TabIndex        =   16
         Top             =   375
         Width           =   315
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "$us."
         Height          =   195
         Left            =   2895
         TabIndex        =   15
         Top             =   1395
         Width           =   300
      End
   End
End
Attribute VB_Name = "ALFrmEntregaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public QResp As Boolean
Public CodItem As String
Public Item As String
Public CantCaja As Long
Public CantEjem As Long
Public PesoKgs As Currency
Public PrecioSus As Currency
Public PesoTotal As Currency
Public PrecioTotal As Currency
Public estado As Integer
'--
Dim rsProv As ADODB.Recordset
Private Sub CmdAceptar_Click()
    If valida Then
        QResp = True
        GrabaDatos
        Unload Me
    End If
End Sub
Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdElige_Click()
    With ALFrmMateriales
        .ALPrincipal
        If .QResp Then
            txtCodigo.Text = .QCodigo
            txtDesc.Text = .QItem
        End If
    End With
End Sub
Private Sub Form_Load()
    QResp = False
    If estado = 1 Then ' Agregar
        VaciaDatos
    ElseIf estado = 2 Then ' Modifcar
        LlenaDatos
    Else
        Fra(0).Enabled = False
        Fra(1).Enabled = False
        Cmdaceptar.Enabled = False
    End If
	Call SeguridadSet(Me)
End Sub
Public Sub LlenaDatos()
    txtCodigo.Text = CodItem
    txtDesc.Text = Item
    tdbnCantCaja.Value = CantCaja
'    tdbnCantEjem.Value = CantEjem
    tdbnPesoKgs.Value = PesoKgs
'    tdbnPrecioSus.Value = PrecioSus
    tdbnPesoTotal.Value = PesoTotal
'    tdbnPrecioTotal.Value = PrecioTotal
End Sub
Public Sub VaciaDatos()
    txtCodigo.Text = ""
    txtDesc.Text = ""
    tdbnCantCaja.Value = 0
    tdbnCantEjem.Value = 0
    tdbnPesoKgs.Value = 0
    tdbnPrecioSus.Value = 0
    tdbnPesoTotal.Value = 0
    tdbnPrecioTotal.Value = 0
End Sub
Public Sub GrabaDatos()
    CodItem = txtCodigo.Text
    Item = txtDesc.Text
    CantCaja = tdbnCantCaja.Value
'    CantEjem = tdbnCantEjem.Value
    PesoKgs = tdbnPesoKgs.Value
'    PrecioSus = tdbnPrecioSus.Value
    PesoTotal = tdbnPesoTotal.Value
'    PrecioTotal = tdbnPrecioTotal.Value
End Sub
Private Function valida() As Boolean
    valida = False
    If txtCodigo.Text = "" Then
        MsgBox "Elija el Item.", vbExclamation + vbOKOnly, "Atención"
        cmdElige.SetFocus
        Exit Function
    End If
    If tdbnCantCaja.Value <= 0 Then
        MsgBox "La Cantidad de Cajas del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantCaja.SetFocus
        Exit Function
    End If
'    If tdbnCantEjem.Value <= 0 Then
'        MsgBox "La Cantidad de Ejemplares del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
'        tdbnCantEjem.SetFocus
'        Exit Function
'    End If
'    If tdbnCantCaja.Value > tdbnCantEjem.Value Then
'        MsgBox "La Cantidad de Cajas del Item no debe ser Mayor a la Cantidad de Ejemplares.", vbExclamation + vbOKOnly, "Atención"
'        tdbnCantCaja.SetFocus
'        Exit Function
'    End If
''    If tdbnCantCaja.Value <> CantidadCajas(txtCodigo.Text, tdbnCantEjem.Value) Then
''        MsgBox "La Cantidad de Cajas no tiene relación con el Número de Ejemplares." & vbCrLf & "1 Caja = " & CantidadEjm(txtCodigo.Text, 1) & " Ejemplares", vbExclamation + vbOKOnly, "Atención"
''        tdbnCantCaja.SetFocus
''        Exit Function
''    End If
'    If (tdbnCantEjem.Value Mod tdbnCantCaja.Value) > 0 Then
'        MsgBox "La Cantidad de Cajas no tiene relación con el Número de Ejemplares." & vbCrLf & "1 Caja = " & tdbnCantEjem.Value / CCur(tdbnCantCaja.Value) & " Ejemplares ?", vbExclamation + vbOKOnly, "Atención"
'        tdbnCantCaja.SetFocus
'        Exit Function
'    End If
    If tdbnPesoKgs.Value <= 0 Then
        MsgBox "El Peso Unitario del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnPesoKgs.SetFocus
        Exit Function
    End If
'    If tdbnPrecioSus.Value <= 0 Then
'        MsgBox "El Precio Unitario del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
'        tdbnPrecioSus.SetFocus
'        Exit Function
'    End If
    If tdbnPesoTotal.Value <= 0 Then
        MsgBox "El Peso Total del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
'        tdbnPesoTotal.SetFocus
        Exit Function
    End If
'    If tdbnPrecioTotal.Value <= 0 Then
'        MsgBox "El Precio Total del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
'        tdbnPrecioTotal.SetFocus
'        Exit Function
'    End If
    If tdbnPesoKgs.Value > tdbnPesoTotal.Value Then
        MsgBox "El Peso Unitario del Item no debe ser Mayor al Peso Total.", vbExclamation + vbOKOnly, "Atención"
        tdbnPesoKgs.SetFocus
        Exit Function
    End If
'    If tdbnPrecioSus.Value > tdbnPrecioTotal.Value Then
'        MsgBox "El Precio Unitario del Item no debe ser Mayor al Precio Total.", vbExclamation + vbOKOnly, "Atención"
'        tdbnPrecioSus.SetFocus
'        Exit Function
'    End If
    valida = True
End Function

Private Sub tdbnCantCaja_Change()
  tdbnPesoTotal.Value = tdbnCantCaja.Value * tdbnPesoKgs.Value
End Sub

Private Sub tdbnCantEjem_Change()
    tdbnPesoTotal.Value = tdbnCantEjem.Value * tdbnPesoKgs.Value
    tdbnPrecioTotal.Value = tdbnCantEjem.Value * tdbnPrecioSus.Value
End Sub
Private Sub tdbnPesoKgs_Change()
    tdbnPesoTotal.Value = tdbnCantCaja.Value * tdbnPesoKgs.Value
End Sub
Private Sub tdbnPrecioSus_Change()
    tdbnPrecioTotal.Value = tdbnCantEjem.Value * tdbnPrecioSus.Value
End Sub


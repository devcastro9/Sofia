VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Begin VB.Form ALFrmIngresoDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso del Detalle"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "ALFrmIngresoDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fra 
      Height          =   1470
      Index           =   0
      Left            =   15
      TabIndex        =   19
      Top             =   -60
      Width           =   6300
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1065
         Width           =   1590
      End
      Begin VB.CommandButton cmdElige 
         Caption         =   "..."
         Height          =   225
         Left            =   180
         TabIndex        =   1
         Top             =   795
         Width           =   255
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1845
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1065
         Width           =   4245
      End
      Begin VB.TextBox txtProv 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1845
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   390
         Width           =   4245
      End
      Begin TrueOleDBList60.TDBCombo tdbcProv 
         Height          =   300
         Left            =   180
         OleObjectBlob   =   "ALFrmIngresoDet.frx":6852
         TabIndex        =   0
         Top             =   375
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Item"
         Height          =   195
         Left            =   495
         TabIndex        =   23
         Top             =   825
         Width           =   300
      End
   End
   Begin VB.Frame Fra 
      BorderStyle     =   0  'None
      Height          =   2070
      Index           =   1
      Left            =   15
      TabIndex        =   11
      Top             =   1305
      Width           =   6300
      Begin VB.Label lblUnidadCaja 
         AutoSize        =   -1  'True
         Caption         =   "Unidades x Caja"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   1755
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Al Tipo de Cambio "
         Height          =   195
         Left            =   3840
         TabIndex        =   3
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "$us."
         Height          =   195
         Left            =   4920
         TabIndex        =   4
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Kgs."
         Height          =   195
         Left            =   4920
         TabIndex        =   5
         Top             =   780
         Width           =   315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "$us."
         Height          =   195
         Left            =   3270
         TabIndex        =   6
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Kgs."
         Height          =   195
         Left            =   3270
         TabIndex        =   7
         Top             =   780
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Unid.(s)"
         Height          =   195
         Left            =   1215
         TabIndex        =   8
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Caja(s)"
         Height          =   195
         Left            =   1200
         TabIndex        =   12
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   465
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Peso Unitario Caja"
         Height          =   195
         Left            =   2190
         TabIndex        =   16
         Top             =   465
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Precio Unitario"
         Height          =   195
         Left            =   2190
         TabIndex        =   15
         Top             =   1125
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Peso Total Caja(s)"
         Height          =   195
         Left            =   3855
         TabIndex        =   14
         Top             =   465
         Width           =   1290
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Precio Total"
         Height          =   195
         Left            =   3855
         TabIndex        =   13
         Top             =   1125
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   435
      Left            =   150
      TabIndex        =   9
      Top             =   3405
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   4845
      TabIndex        =   10
      Top             =   3405
      Width           =   1320
   End
End
Attribute VB_Name = "ALFrmIngresoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public QResp As Boolean
Public CodProv As String
Public NomProv As String
Public CodItem As String
Public Item As String
Public CantCaja As Long
Public CantEjem As Long
Public PesoKgs As Currency
Public PrecioSus As Currency
Public PesoTotal As Currency
Public PrecioTotal As Currency
Public TipoCambio As Currency
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

Private Sub Form_Activate()
    If estado = 2 Then tdbcProv.Text = CodProv
End Sub

Private Sub Form_Load()
    QResp = False
    Set rsProv = New ADODB.Recordset
    GlSqlAux = "SELECT * FROM ac_Proveedor ORDER BY Ruc_Id"
    rsProv.Open GlSqlAux, DB, adOpenStatic
    Set tdbcProv.RowSource = rsProv
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
    tdbcProv.Text = CodProv
    rsProv.MoveFirst
    rsProv.Find "Ruc_Id = '" & CodProv & "'"
    txtProv.Text = ""
    If Not rsProv.EOF Then txtProv.Text = rsProv!Descripcion_Larga & ""
    txtCodigo.Text = CodItem
    txtDesc.Text = Item
    tdbnCantCaja.Value = CantCaja
    tdbnCantEjem.Value = CantEjem
    tdbnPesoKgs.Value = PesoKgs
    tdbnPrecioSus.Value = PrecioSus
    tdbnPesoTotal.Value = PesoTotal
    tdbnPrecioTotal.Value = PrecioTotal
    tdbnTipoCambio.Value = TipoCambio
    If tdbnCantCaja.Value > 0 Then
        lblUnidadCaja.Caption = (tdbnCantEjem.Value / CCur(tdbnCantCaja.Value)) & " Unidades x Caja"
    Else
        lblUnidadCaja.Caption = "Unidades x Caja"
    End If
    'lblUnidadCaja.Caption = IIf(tdbnCantEjem.Value > 0, (tdbnCantEjem.Value / CCur(tdbnCantEjem.Value)) & " Unidades x Caja", "Unidades x Caja")
End Sub
Public Sub VaciaDatos()
    tdbcProv.Text = ""
    txtProv.Text = ""
    txtCodigo.Text = ""
    txtDesc.Text = ""
    tdbnCantCaja.Value = 0
    tdbnCantEjem.Value = 0
    tdbnPesoKgs.Value = 0
    tdbnPrecioSus.Value = 0
    tdbnPesoTotal.Value = 0
    tdbnPrecioTotal.Value = 0
    tdbnTipoCambio.Value = 0
    lblUnidadCaja.Caption = "Unidades x Caja"
End Sub
Public Sub GrabaDatos()
    CodProv = tdbcProv.Text
    NomProv = txtProv.Text
    CodItem = txtCodigo.Text
    Item = txtDesc.Text
    CantCaja = tdbnCantCaja.Value
    CantEjem = tdbnCantEjem.Value
    PesoKgs = tdbnPesoKgs.Value
    PrecioSus = tdbnPrecioSus.Value
    PesoTotal = tdbnPesoTotal.Value
    PrecioTotal = tdbnPrecioTotal.Value
    TipoCambio = tdbnTipoCambio.Value
End Sub
Private Function valida() As Boolean
    valida = False
    If tdbcProv.Text = "" Then
        MsgBox "Elija a el Proveedor del Item.", vbExclamation + vbOKOnly, "Atención"
        tdbcProv.SetFocus
        Exit Function
    End If
    If txtCodigo.Text = "" Then
        MsgBox "Elija el Item.", vbExclamation + vbOKOnly, "Atención"
        tdbcProv.SetFocus
        Exit Function
    End If
    If tdbnCantCaja.Value <= 0 Then
        MsgBox "La Cantidad de Cajas del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantCaja.SetFocus
        Exit Function
    End If
    If tdbnCantEjem.Value <= 0 Then
        MsgBox "La Cantidad de Unidades del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantEjem.SetFocus
        Exit Function
    End If
    If tdbnCantCaja.Value > tdbnCantEjem.Value Then
        MsgBox "La Cantidad de Cajas del Item no debe ser Mayor a la Cantidad de Unidades.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantCaja.SetFocus
        Exit Function
    End If
'    If tdbnCantCaja.Value <> CantidadCajas(txtCodigo.Text, tdbnCantEjem.Value) Then
'        MsgBox "La Cantidad de Cajas no tiene relación con el Número de Unidades." & vbCrLf & "1 Caja = " & CantidadEjm(txtCodigo.Text, 1) & " Unidades", vbExclamation + vbOKOnly, "Atención"
'        tdbnCantCaja.SetFocus
'        Exit Function
'    End If
    If (tdbnCantEjem.Value Mod tdbnCantCaja.Value) > 0 Then
        MsgBox "La Cantidad de Cajas no tiene relación con el Número de Unidades." & vbCrLf & "1 Caja = " & tdbnCantEjem.Value / CCur(tdbnCantCaja.Value) & " Unidades ?", vbExclamation + vbOKOnly, "Atención"
        tdbnCantCaja.SetFocus
        Exit Function
    End If
    If tdbnPesoKgs.Value <= 0 Then
        MsgBox "El Peso Unitario del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnPesoKgs.SetFocus
        Exit Function
    End If
    If tdbnPrecioSus.Value <= 0 Then
        MsgBox "El Precio Unitario del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnPrecioSus.SetFocus
        Exit Function
    End If
    If tdbnPesoTotal.Value <= 0 Then
        MsgBox "El Peso Total del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnPesoTotal.SetFocus
        Exit Function
    End If
    If tdbnPrecioTotal.Value <= 0 Then
        MsgBox "El Precio Total del Item no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnPrecioTotal.SetFocus
        Exit Function
    End If
    If tdbnPesoKgs.Value > tdbnPesoTotal.Value Then
        MsgBox "El Peso Unitario del Item no debe ser Mayor al Peso Total.", vbExclamation + vbOKOnly, "Atención"
        tdbnPesoKgs.SetFocus
        Exit Function
    End If
    If tdbnPrecioSus.Value > tdbnPrecioTotal.Value Then
        MsgBox "El Precio Unitario del Item no debe ser Mayor al Precio Total.", vbExclamation + vbOKOnly, "Atención"
        tdbnPrecioSus.SetFocus
        Exit Function
    End If
    If tdbnTipoCambio.Value <= 0 Then
        MsgBox "El Tipo de Cambio no debe ser CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnTipoCambio.SetFocus
        Exit Function
    End If
    valida = True
End Function
Private Sub tdbcProv_ItemChange()
    txtProv.Text = tdbcProv.Columns("descripcion_larga").Value
End Sub
Private Sub tdbcProv_NotInList(NewEntry As String, Retry As Integer)
    tdbcProv.Text = ""
    txtProv.Text = ""
End Sub
Private Sub tdbnCantCaja_Change()
    If tdbnCantCaja.Value > 0 Then
        lblUnidadCaja.Caption = (tdbnCantEjem.Value / CCur(tdbnCantCaja.Value)) & " Unidades x Caja"
    Else
        lblUnidadCaja.Caption = "Unidades x Caja"
    End If
    'lblUnidadCaja.Caption = IIf(tdbnCantCaja.Value > 0, (tdbnCantEjem.Value / CCur(tdbnCantCaja.Value)) & " Unidades x Caja", "Unidades x Caja")
End Sub
Private Sub tdbnCantEjem_Change()
    If tdbnCantCaja.Value > 0 Then
        lblUnidadCaja.Caption = (tdbnCantEjem.Value / CCur(tdbnCantCaja.Value)) & " Unidades x Caja"
    Else
        lblUnidadCaja.Caption = "Unidades x Caja"
    End If
    'lblUnidadCaja.Caption = IIf(tdbnCantCaja.Value > 0, (tdbnCantEjem.Value / CCur(tdbnCantCaja.Value)) & " Unidades x Caja", "Unidades x Caja")
    tdbnPesoTotal.Value = tdbnCantCaja.Value * tdbnPesoKgs.Value
    tdbnPrecioTotal.Value = tdbnCantEjem.Value * tdbnPrecioSus.Value
End Sub
Private Sub tdbnPesoKgs_Change()
    tdbnPesoTotal.Value = tdbnCantCaja.Value * tdbnPesoKgs.Value
End Sub
Private Sub tdbnPrecioSus_Change()
    tdbnPrecioTotal.Value = tdbnCantEjem.Value * tdbnPrecioSus.Value
End Sub

VERSION 5.00
Begin VB.Form ALFrmEntDet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entrega Manual"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   Icon            =   "ALFrmEntDet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2768
      TabIndex        =   3
      Top             =   2955
      Width           =   1590
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   435
      Left            =   353
      TabIndex        =   2
      Top             =   2955
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   15
      TabIndex        =   4
      Top             =   -75
      Width           =   4680
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   75
         TabIndex        =   0
         Top             =   2025
         Width           =   4515
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   75
         TabIndex        =   10
         Top             =   1110
         Width           =   4515
      End
      Begin VB.TextBox txtNoLici 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   735
         Width           =   1170
      End
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   165
         Width           =   4515
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Unidad(es)"
         Height          =   195
         Left            =   3660
         TabIndex        =   1
         Top             =   2490
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Caja(s)"
         Height          =   195
         Left            =   1365
         TabIndex        =   7
         Top             =   2490
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Unidad(es)"
         Height          =   195
         Left            =   3660
         TabIndex        =   8
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Caja(s)"
         Height          =   195
         Left            =   1365
         TabIndex        =   15
         Top             =   1530
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   2370
         TabIndex        =   14
         Top             =   2190
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   75
         TabIndex        =   13
         Top             =   2190
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Left            =   2370
         TabIndex        =   12
         Top             =   1230
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   1230
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Licitación"
         Height          =   195
         Left            =   75
         TabIndex        =   9
         Top             =   510
         Width           =   975
      End
   End
End
Attribute VB_Name = "ALFrmEntDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IdIngreso As Long
Public NoLici As String
Public Codigo As String
Public Item As String
Public SaldoCaja As Long
Public SaldoEj As Long
Public CantCaja As Long
Public CantEjm As Long
Public UnidadCaja As Integer
Public estado As Integer ' 0 Navegar, 1 Agregar, 2 Editar
Public QResp As Boolean
Private Sub VaciaCampos()
    txtItem.Text = ""
    txtNoLici.Text = ""
    tdbnSaldoCaja.Value = 0
    tdbnSaldoEj.Value = 0
    tdbnCantCaja.Value = 0
    tdbnCantEjm.Value = 0
End Sub
Private Sub LlenaCampos()
    txtItem.Text = Codigo & " : " & Item
    txtNoLici.Text = NoLici
    tdbnSaldoCaja.Value = SaldoCaja
    tdbnSaldoEj.Value = SaldoEj
    tdbnCantCaja.Value = CantCaja
    tdbnCantEjm.Value = CantEjm
End Sub
Public Sub GrabaCampos()
    CantCaja = tdbnCantCaja.Value
    CantEjm = tdbnCantEjm.Value
End Sub
Private Sub CmdAceptar_Click()
    If valida Then
        GrabaCampos
        QResp = True
        Unload Me
    End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    QResp = False
    If estado = 1 Then
        VaciaCampos
    Else
        LlenaCampos
    End If
	Call SeguridadSet(Me)
End Sub
Private Function valida() As Boolean
    valida = False
    If tdbnCantCaja.Value < 0 Then
        MsgBox "La cantidad de Cajas No debe ser Menor a CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantCaja.SetFocus
        Exit Function
    End If
    If tdbnCantEjm.Value < 0 Then
        MsgBox "La cantidad de Ejemplares No debe ser Menor a CERO.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantCaja.SetFocus
        Exit Function
    End If
    If tdbnCantCaja.Value > tdbnSaldoCaja.Value Then
        MsgBox "La cantidad de Cajas No debe ser Mayor al Saldo de Cajas.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantCaja.SetFocus
        Exit Function
    End If
    If tdbnCantEjm.Value > tdbnSaldoEj.Value Then
        MsgBox "La cantidad de Ejemplares No debe ser Mayor al Saldo de Ejemplares.", vbExclamation + vbOKOnly, "Atención"
        tdbnCantEjm.SetFocus
        Exit Function
    End If
'    If tdbnCantCaja.Value <> CantidadCajas(Codigo, tdbnCantEjm.Value) Then
'        MsgBox "La Cantidad de Cajas no tiene relación con el Número de Ejemplares." & vbCrLf & "1 Caja = " & CantidadEjm(Codigo, 1) & " Ejemplares", vbExclamation + vbOKOnly, "Atención"
'        tdbnCantCaja.SetFocus
'        Exit Function
'    End If
    
    valida = True
End Function
Private Sub tdbnCantCaja_Change()
    tdbnCantEjm.Value = tdbnCantCaja.Value * UnidadCaja
End Sub


VERSION 5.00
Begin VB.Form FrmPresentacion 
   Caption         =   "Conciliacion Bancaria"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   -30
      TabIndex        =   0
      Top             =   -105
      Width           =   5085
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         CausesValidation=   0   'False
         Height          =   480
         Left            =   645
         TabIndex        =   3
         Top             =   1470
         Width           =   3630
      End
      Begin VB.CommandButton CmdTransferencias 
         Caption         =   "Transferencias"
         Height          =   450
         Left            =   645
         TabIndex        =   2
         Top             =   975
         Width           =   3630
      End
      Begin VB.CommandButton CmdCheques 
         Caption         =   "Cheques"
         Height          =   480
         Left            =   660
         TabIndex        =   1
         Top             =   450
         Width           =   3630
      End
   End
End
Attribute VB_Name = "FrmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCheques_Click()
    swConciliacion = "CHEQUE"
    FrmEleccion.LblTitulo(1).Caption = "Cheques"
    FrmEleccion.LblTitulo(0).Caption = "Cheques"
    FrmEleccion.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTransferencias_Click()
    swConciliacion = "TRANSFERENCIA"
    FrmEleccion.LblTitulo(1).Caption = "Transferencias"
    FrmEleccion.LblTitulo(0).Caption = "Transferencias"
    FrmEleccion.Show
End Sub

Private Sub Command1_Click()

End Sub


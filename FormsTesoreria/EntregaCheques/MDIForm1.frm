VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Ctrl 
      Caption         =   "Control de Cheques"
      Begin VB.Menu CE 
         Caption         =   "Cheques a Entregar"
      End
      Begin VB.Menu tributosFiscales 
         Caption         =   "Tributos Fiscales"
      End
      Begin VB.Menu sl 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CE_Click()
    FrmActivacionCheques.Show
End Sub

Private Sub sl_Click()
    Unload Me
    End
End Sub

Private Sub tributosFiscales_Click()
    FrmTributosFiscales.Show
End Sub

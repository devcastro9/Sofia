VERSION 5.00
Begin VB.Form frmmenu 
   Caption         =   "Adquisiones"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4890
   Icon            =   "frmmenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuadquisiones 
      Caption         =   "Adquisiones"
      Begin VB.Menu mnuantecedentes 
         Caption         =   "Solicitud de Desembolso"
      End
      Begin VB.Menu mnuline 
         Caption         =   "Planilla de Deducciones"
      End
      Begin VB.Menu mnusale 
         Caption         =   "salir"
      End
   End
   Begin VB.Menu mnusalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuadjudiacion_Click()
 'frmadjudicacion.Show vbModal
End Sub

Private Sub mnuantecedentes_Click()
  frmSoesMain.frmSoesMain_procesar "ABM_SOES"
End Sub

Private Sub mnuapertura_Click()
 'frmaperturasobres.Show vbModal
End Sub

Private Sub mnunoObjecion_Click()
 frmSoesMain.Show vbModal
End Sub

Private Sub mnupliegos_Click()
 'frmpliegos.Show vbModal
End Sub

Private Sub mnupublicacion_Click()
 'frmpublicacion.Show vbModal
End Sub

Private Sub mnurecepcion_Click()
 'frmrecepcionsobres.Show vbModal
End Sub

Private Sub mnuline_Click()
  frmSoesMain.frmSoesMain_procesar "DEDUCCIONES"
End Sub

Private Sub mnusalir_Click()
 End
End Sub

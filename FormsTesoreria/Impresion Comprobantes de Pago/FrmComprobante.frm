VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Begin VB.Form FrmComprobante 
   Caption         =   "Impresión Comrpobante de Pago"
   ClientHeight    =   6465
   ClientLeft      =   1785
   ClientTop       =   3030
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   6660
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   120
      TabIndex        =   0
      Top             =   15
      Width           =   5805
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
End
Attribute VB_Name = "FrmComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Dim Report As New CryComprobante
'CryComprobante.Comprobante2 = FrmControlPagos.TxtNC.Text
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
'TOMAR NOTA DE  REFRECAR
'CRViewer1.EnableRefreshButton = True
'CRViewer1.Refresh
'iMPRIMIR DIRECTAMENTE
'CRViewer1.PrintReport
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

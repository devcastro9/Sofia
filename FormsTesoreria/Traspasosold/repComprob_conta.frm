VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Begin VB.Form RepComprob_Conta 
   Caption         =   "   "
   ClientHeight    =   7065
   ClientLeft      =   1815
   ClientTop       =   2745
   ClientWidth     =   9075
   LinkTopic       =   "Form2"
   ScaleHeight     =   7065
   ScaleWidth      =   9075
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6996
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8676
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
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
End
Attribute VB_Name = "RepComprob_Conta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CryComprob_conta

Private Sub Form_Load()
Set Report = New CryComprob_conta
CRViewer1.ReportSource = Report
CRViewer1.ViewReport

	Call SeguridadSet(Me)
End Sub

Private Sub Form_Resize()
'Dim Report As New CrtComprobante
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

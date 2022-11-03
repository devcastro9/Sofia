VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form RepOperacionesCheques 
   Caption         =   "Form3"
   ClientHeight    =   7845
   ClientLeft      =   1545
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   7200
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControl=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertControl=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
   End
End
Attribute VB_Name = "RepOperacionesCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim Report As New CryOpCheques
    Report.FormulaFields(1).Text = "'" & FrmDesplegado.DTPFechaInicio.Value & "'"
    Report.FormulaFields(2).Text = "'" & FrmDesplegado.DTPFechaFin.Value & "'"
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
	Call SeguridadSet(Me)
End Sub
Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
End Sub

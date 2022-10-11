VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form RepCuentaBancaria 
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
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
Attribute VB_Name = "RepCuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim Report As New CryCuentaBancaria
    
    Report.FormulaFields(1).Text = "'" & FrmCuentaBancaria.DTPFechaInicio.Value & "'"
    '
    'Report.FormulaFields(1).Text = "'" & FrmCuentaBancaria.DTPFechaInicio.Value & "'"
    'Report.FormulaFields(4).Text = "'" & FrmCuentaBancaria.DTPFechaFin.Value & "'"
    CRViewer1.ReportSource = Report
    'CryCuentaBancaria.TxtTitulo = "hola"
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

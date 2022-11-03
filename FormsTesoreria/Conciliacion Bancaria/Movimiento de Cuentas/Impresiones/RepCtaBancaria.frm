VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form RepCtaBancaria 
   Caption         =   "Form2"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   4875
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   5805
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
Attribute VB_Name = "RepCtaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim Report As New CryCtaBancaria
Private Sub Form_Load()
Dim Report As New CryCtaBancaria

'Report.FormulaFields(2).Text = "'" & FrmCuentaBancaria.DTPFechaInicio.Value & "'"
If swMes = 1 Then
    Report.FormulaFields(5).Text = "'" & FrmCuentaBancaria.CmbMes.Text & "'"
End If
If swFecha = 1 Then
    De = "De"
    A = "A"
    Report.FormulaFields(2).Text = "'" & FrmCuentaBancaria.DTPFechaInicio.Value & "'"
    Report.FormulaFields(3).Text = "'" & FrmCuentaBancaria.DTPFechaFin.Value & "'"
    Report.FormulaFields(4).Text = "'" & De & "' "
    Report.FormulaFields(5).Text = "'" & A & "' "
End If

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

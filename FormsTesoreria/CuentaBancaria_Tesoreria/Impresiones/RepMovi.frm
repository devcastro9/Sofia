VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Begin VB.Form RepMovi 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   7860
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
Attribute VB_Name = "RepMovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'Dim Report As New CryMovi
Dim De As String
Dim A As String
Dim TC As String
Dim Cadena As String
If FrmCuentaBancaria.OptFechaPago.Value = True Then
     Cadena = "REPORTE POR FECHA DE PAGO"
Else
     Cadena = "REPORTE POR FECHA DE IMPRESION"
End If

If FrmCuentaBancaria.OptFechaPago.Value = True Then Report.FormulaFields(8).Text = "'" & Cadena & "'"
If FrmCuentaBancaria.OptFechaImpresion.Value = True Then Report.FormulaFields(8).Text = "'" & Cadena & "'"

If FrmCuentaBancaria.OptUnaCuenta.Value = True Then
    Report.FormulaFields(6).Text = "'" & FrmCuentaBancaria.DtCCuentaOrigen.Text & "'"
    Report.FormulaFields(7).Text = "'" & FrmCuentaBancaria.DtCDescripcion.Text & "'"
    
End If
If FrmCuentaBancaria.OptTodasCuentas.Value = True Then
    TC = "Todas las cuentas"
    Report.FormulaFields(6).Text = "'" & TC & "'"
End If


If swMes = 1 Then
    Report.FormulaFields(1).Text = "'" & FrmCuentaBancaria.CmbMes.Text & "'"
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

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

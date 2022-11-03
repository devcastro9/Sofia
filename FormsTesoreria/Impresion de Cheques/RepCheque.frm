VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Begin VB.Form RepCheque 
   Caption         =   "Form2"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   1350
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   ScaleHeight     =   6270
   ScaleWidth      =   7680
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
Attribute VB_Name = "RepCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsComprobante As New ADODB.Recordset

Private Sub Form_Load()

'   For i = 0 To LstChequesCodigo.ListCount - 2
'        LstChequesCodigo.ListIndex = i
'        NrosChequeImprimir = NrosChequeImprimir & "pago_detalle.numero_cheque_trf= " & "'" & LstChequesCodigo.Text & "'" & " Or "
'        Next i
'
'    LstChequesCodigo.ListIndex = i
'    NrosChequeImprimir = NrosChequeImprimir + "pago_detalle.numero_cheque_trf = " & "'" & LstChequesCodigo.Text & "'"
    'MsgBox NrosChequeImprimir
    
'        If rsComprobante.State = 1 Then rsComprobante.Close
'        Set rsComprobante = New ADODB.Recordset
'        rsComprobante.Open "SELECT fc_beneficiario.denominacion_beneficiario,pago_detalle.numero_cheque_trf,Pagos.codigo_pago, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, fc_beneficiario.denominacion_beneficiario,  pago_detalle.cheque_o_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal " & _
'        "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'C' and  " & NrosChequeImprimir & "", db, adOpenKeyset, adLockOptimistic
'        MsgBox rsComprobante.RecordCount
    'Set DtGCheques.DataSource = rsComprobante
    'DtGCheques.Refresh
    
'    MsgBox "Imprimiendo..."
    
    
    
'Report.Database.SetDataSource rsComprobante
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

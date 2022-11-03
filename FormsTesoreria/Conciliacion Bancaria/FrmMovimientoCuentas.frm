VERSION 5.00
Begin VB.Form FrmMovimientoCuentas 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmMovimientoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Aquí se unen las tablas pago_detalle, Co_MovimientoPCO  y



db.Execute = "INSERT INTO fc_datosgtz(Nro_Cmpte, Organismo, Fecha_Pago, Monto, " & _
          "Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Estado_Conciliacion) " & _
          "SELECT pago_detalle.codigo_pago, pago_detalle.org_codigo, " & _
          "pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, " & _
          "pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario,  " & _
          "pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf, " & _
          "pago_detalle.Cta_codigo , fc_cuenta_bancaria.Bco_codigo, pago_detalle.Estado_Conciliacion  " & _
          "FROM pago_detalle INNER JOIN fc_cuenta_bancaria ON pago_detalle.Cta_codigo = fc_cuenta_bancaria.Cta_codigo "
          
db.Execute = "INSERT INTO fc_datosgtz(Nro_Cmpte, Organismo, Fecha_Pago, Monto, " & _
          "Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Estado_Conciliacion) " & _
          "SELECT pago_detalle.codigo_pago, pago_detalle.org_codigo, " & _
          "pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, " & _
          "pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario,  " & _
          "pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf, " & _
          "pago_detalle.Cta_codigo , fc_cuenta_bancaria.Bco_codigo, pago_detalle.Estado_Conciliacion  " & _
          "FROM pago_detalle INNER JOIN fc_cuenta_bancaria ON pago_detalle.Cta_codigo = fc_cuenta_bancaria.Cta_codigo "
          
	Call SeguridadSet(Me)
End Sub

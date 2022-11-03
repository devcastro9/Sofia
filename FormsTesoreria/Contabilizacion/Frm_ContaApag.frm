VERSION 5.00
Begin VB.Form Frm_ContaApag 
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Pagado 
      Caption         =   "Contabiliza Pagado"
      Height          =   516
      Left            =   732
      TabIndex        =   0
      Top             =   756
      Width           =   2280
   End
End
Attribute VB_Name = "Frm_ContaApag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Pagado_Click()
Dim sw As Boolean
Dim Sw_Fuente As Boolean
Dim Cont_Comp As Long
Dim aux_T As String


db.BeginTrans



'*******************************************************
'******************** Contabilizar Pagos ***************'
'********************************************************
'************** Para inicializar el contador ******************'

'*********** Para obtenerr en el recordset recsetAuxComp losdatos necesarios para almacenar*********"
If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
recSetAuxcomp.Open "SELECT distinct pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion,Estado_Pagado,Pago_Detalle.Cta_Codigo,Pago_Detalle.tipo_cambio," & _
" Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_Bolivianos,estado_Devengado,Pagos.Org_Codigo,Pagos.Codigo_Orden,Pagos.Codigo_Documento," & _
" pago_detalle.Monto_Dolares,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and " & _
" pago_detalle.Ges_Gestion = pagos.Ges_Gestion and  pago_detalle.estado_aprobacion='S'  AND Pagos.Tipo_comp='PAC' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'AND pago_detalle.estado_aprobacion = 'S'
If recSetAuxcomp.RecordCount > 0 Then
recSetAuxcomp.MoveFirst
End If


'************Abrimos un record set para adicionar datos*********************'
Set recSetAuxActualizar = New ADODB.Recordset
If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
recSetAuxActualizar.Open " select * from CO_Comprobante_M ", db, adOpenDynamic, adLockOptimistic, adCmdText

Set recSetAuxActualizar1 = New ADODB.Recordset
If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar.Close
recSetAuxActualizar1.Open " select * from CO_Diario ", db, adOpenDynamic, adLockOptimistic, adCmdText
Dim Aux_Cont As String

aux_T = "select * from Co_comprobante_M"

While Not (recSetAuxcomp.EOF)

If Not Buscar(aux_T, recSetAuxcomp!codigo_pago, recSetAuxcomp!org_codigo, recSetAuxcomp!Ges_gestion, "PAC", recSetAuxcomp!codigo_pago_detalle) Then
       
    Select Case recSetAuxcomp!Fte_codigo
    
    Case "10"
   
    If recSetPartida.State = 1 Then recSetPartida.Close
    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H, CC_Cuentas_D" & _
    " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst= 'PAG' and CC_Cuenta_H.Inst= 'PAG' and " & _
    " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=1 AND " & _
    " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
    Sw_Fuente = True
    
    Case "70"
    If recSetPartida.State = 1 Then recSetPartida.Close
    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H, CC_Cuentas_D" & _
    " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PAG' and CC_Cuenta_H.Inst='PAG' and " & _
    " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=2 AND " & _
    " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
    Sw_Fuente = True
    
    Case "80"
    If recSetPartida.State = 1 Then recSetPartida.Close
    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H, CC_Cuentas_D" & _
    " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PAG' and CC_Cuenta_H.Inst='PAG' and " & _
    " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=3 and  " & _
    " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
    Sw_Fuente = True
    
    Case Else
    'MsgBox "No esta asociado a ninguna fuente ... partida no relacionada "
    Sw_Fuente = False
    
    End Select
  If Sw_Fuente Then
   
    recSetAuxActualizar.AddNew
    recSetAuxActualizar1.AddNew
    'recSetAuxActualizar!Cod_Comp = Cont_Comp
    recSetAuxActualizar!Cod_Trans = recSetAuxcomp!codigo_pago
    recSetAuxActualizar!cod_trans_detalle = recSetAuxcomp!codigo_pago_detalle
    recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
    recSetAuxActualizar!Codigo_Beneficiario = recSetAuxcomp!Codigo_Beneficiario
    recSetAuxActualizar!Ges_gestion = recSetAuxcomp!Ges_gestion
    recSetAuxActualizar!Num_Respaldo = recSetAuxcomp!codigo_orden
    recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
    
    recSetAuxActualizar!Fecha_A = recSetAuxcomp!fecha_pago
    recSetAuxActualizar!Glosa = recSetAuxcomp!justificacion
    'recSetAuxActualizar!codigo_solicitud = recSetAuxcomp!codigo_solicitud
    recSetAuxActualizar!Tipo_Comp = "PAC"
      
    recSetAuxActualizar!Status = "S"
    recSetAuxActualizar1!Tipo_Comp = "PAC"
    recSetAuxActualizar1!d_cuenta = recSetPartida!cuenta
    recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
    recSetAuxActualizar1!d_subcta1 = recSetPartida!Subcta1
    recSetAuxActualizar1!d_subcta2 = recSetPartida!Subcta2
    recSetAuxActualizar1!d_Aux1 = recSetPartida!Aux1
    recSetAuxActualizar1!d_Aux2 = recSetPartida!Aux2
    recSetAuxActualizar1!d_Aux3 = recSetPartida!Aux3
    
'************* CONTABILIZA AUXILIAARES DEBITO
    Select Case recSetPartida!Aux1
    Case "01"
            Set recsetAdicion = New ADODB.Recordset
            If recsetAdicion.State = 1 Then recsetAdicion.Close
            recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_Beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
            recSetAuxActualizar1!D_Cta_Larga = recsetAdicion!Codigo_Beneficiario
            recSetAuxActualizar1!D_Des_Larga = recsetAdicion!denominacion_beneficiario
            
    Case "02"
            Set recsetAdicion = New ADODB.Recordset
            If recsetAdicion.State = 1 Then recsetAdicion.Close
            recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!Cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
            recSetAuxActualizar1!D_Cta_Larga = recsetAdicion!Cta_codigo
            recSetAuxActualizar1!D_Des_Larga = recsetAdicion!Cta_descripcion_larga

    Case Else
    End Select
''****************** finaliza sesion de auxiliares
           
       
    recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
    recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
    recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
    
'************* CONTABILIZA AUXILIAARES DEBITO

    Select Case recSetPartida!h_Aux1
    Case "01"
            Set recsetAdicion = New ADODB.Recordset
            If recsetAdicion.State = 1 Then recsetAdicion.Close
        
            recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_Beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
            recSetAuxActualizar1!H_Cta_Larga = recsetAdicion!Codigo_Beneficiario
            recSetAuxActualizar1!H_Des_Larga = recsetAdicion!denominacion_beneficiario
            
    Case "02"
            Set recsetAdicion = New ADODB.Recordset
            If recsetAdicion.State = 1 Then recsetAdicion.Close
            
            recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!Cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
            'recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
            recSetAuxActualizar1!H_Cta_Larga = recsetAdicion!Cta_codigo
            recSetAuxActualizar1!H_Des_Larga = recsetAdicion!Cta_descripcion_larga

    Case Else
    End Select
''****************** finaliza sesion de auxiliares

    recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
    recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
    recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
    recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2
    recSetAuxActualizar1!D_MontoBs = recSetAuxcomp!monto_Bolivianos
    recSetAuxActualizar1!D_MontoDl = recSetAuxcomp!monto_Dolares
    recSetAuxActualizar1!D_MontoDl = recSetAuxcomp!monto_Dolares
    recSetAuxActualizar1!D_Cambio = recSetAuxcomp!tipo_cambio
    
    recSetAuxActualizar1!H_MontoBs = recSetAuxcomp!monto_Bolivianos
    recSetAuxActualizar1!H_MontoDl = recSetAuxcomp!monto_Dolares
    recSetAuxActualizar1!H_MontoDl = recSetAuxcomp!monto_Dolares
    recSetAuxActualizar1!H_Cambio = recSetAuxcomp!tipo_cambio
''************ GENERA EL CODIGO DE COMPROBANTE**********

            Set recSetGenera = New ADODB.Recordset
            recSetGenera.CursorLocation = adUseClient
            If recSetGenera.State = 1 Then recSetGenera.Close
            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
            If recSetGenera.RecordCount > 0 Then
             Cont_Comp = Val(recSetGenera!numero_correlativo)
             Cont_Comp = Cont_Comp + 1
             recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
             

     
'************TERMINA GENERACION DE COMPROBANTE********
             recSetAuxActualizar!Cod_Comp = Cont_Comp
             recSetAuxActualizar1!Cod_Comp = Cont_Comp
             recSetAuxActualizar1.Update
             recSetAuxActualizar.Update
             recSetGenera.Update

            End If
    
   Else
   MsgBox "No esta asociado a ninguna fuente ... partida no relacionada "
  
   End If
Else
MsgBox "Existe registro....."
End If
    'Cont_Comp = Cont_Comp + 1
    recSetAuxcomp.MoveNext

Wend
'recSetGenera!Numero_correlativo = Cont_Comp
'recSetGenera.Update
db.CommitTrans
MsgBox "CONTABILIZO CORRECTAMENTE "

End Sub

Private Sub Form_Load()
Set recSetBusqueda = New ADODB.Recordset
Set recSetAuxActualizar = New ADODB.Recordset
Set recSetAuxActualizar1 = New ADODB.Recordset

Set recsetAdicion = New ADODB.Recordset

Set recSetAuxcomp = New ADODB.Recordset
Set recSetAuxcomp1 = New ADODB.Recordset
Set recSetPartida = New ADODB.Recordset
Set recSetPartida1 = New ADODB.Recordset

Set recSetAuxcomp1 = New ADODB.Recordset
Set recSetAuxbenefi1 = New ADODB.Recordset
Set recSetPartid1 = New ADODB.Recordset



recSetPartida.CursorLocation = adUseClient
recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
recSetAuxActualizar.CursorLocation = adUseClient


	Call SeguridadSet(Me)
End Sub

Attribute VB_Name = "ModContabilizacion"
Public permite As String
Public existecomp As Integer
'---(1) si el comprobante se genero, (0) problemas en la generacion
Public regANL999 As Integer
Public regDEV999 As Integer
Public regRVT999 As Integer
Public regANLTRP  As Integer
Public regDAC As Integer
Public regANL As Integer
Public regPCC As Integer
Public regRVT As Integer
'---numeros generados de comprobantes
Public numRVT999 As Integer
Public numDEV999 As Integer
Public numANL999 As Integer
Public numANLTRP As Integer
Public numPCC As Integer
'--
Public numANL As Integer
'---
Public anul999 As Integer
Public rever999 As Integer
Public generoTRP As Integer
Public numero As Integer
'Option Explicit
Public Sub Cmd_Pagado(P_codigo_pago As String, P_codigo_pago_detalle As String, P_org_codigo As String, P_ges_gestion As String)
  Dim Sw As Boolean
  Dim Sw_Fuente As Boolean
  Dim Cont_Comp As Long
  Dim aux_T As String
  
  Dim v_Cuenta As String
  Dim v_SubCta1 As String
  Dim v_SubCta2 As String
  Dim v_NombreCta As String
  Dim v_H_Cuenta As String
  Dim v_H_SubCta1 As String
  Dim v_H_SubCta2 As String
  Dim v_H_NombCta As String
  Dim v_Aux1 As String
  Dim v_Aux2 As String
  Dim v_Aux3 As String
  Dim v_H_Aux1 As String
  Dim v_H_Aux2 As String
  Dim v_H_Aux3 As String
  Dim Aux_Cont As String
  Dim rstipopy As ADODB.Recordset
  Set rstipopy = New ADODB.Recordset
'On Error GoTo errorPag

  db.BeginTrans
  MsgBox "Contabilizando............", vbInformation + vbOKOnly, "Contabilización"
  Set recSetAuxcomp = New ADODB.Recordset
  recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
  'If Me.DtCCuentaOrigen.Text = "" Then
  '  MsgBox "ERROR, NO SE CONTABILIZO", vbDefaultButton1 + vbOKOnly
  '       Exit Sub
    'End If
  If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
  recSetAuxcomp.Open "SELECT distinct pago_detalle.pro_programa,pago_Detalle.pro_subprograma,pago_Detalle.pro_proyecto,pago_detalle.pro_actividad,pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion,Estado_Pagado,Pago_Detalle.Cta_Codigo,Pago_Detalle.tipo_cambio," & _
        " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_Total as Monto_Bolivianos,estado_Devengado,Pagos.Org_Codigo,Pagos.Codigo_Orden,Pagos.Codigo_Documento," & _
        " pago_detalle.Monto_Dolares,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and   " & _
        " pago_Detalle.Org_codigo= '" & P_org_codigo & "' and  pago_detalle.Ges_Gestion='" & P_ges_gestion & "' and pago_detalle.codigo_Pago=" & Val(P_codigo_pago) & " and " & _
        " pago_detalle.Ges_Gestion = pagos.Ges_Gestion  and pago_detalle.codigo_pago_detalle='" & P_codigo_pago_detalle & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
  If recSetAuxcomp.RecordCount > 0 Then
      recSetAuxcomp.MoveFirst
  Else
      MsgBox "ERROR EN LA CONTABILIZACION", vbCritical + vbDefaultButton1
      Exit Sub
  End If
  While Not (recSetAuxcomp.EOF)
      If rstipopy.State = 1 Then rstipopy.Close
      'Dim sqlpy  As String
      Dim tipopy As String
      rstipopy.Open "select tipo_proyecto from fc_estructura_programatica where Pro_programa='" & recSetAuxcomp!pro_programa & "' and  Pro_subprograma='" & recSetAuxcomp!pro_subprograma & "' and Pro_proyecto='" & recSetAuxcomp!pro_proyecto & "' and Pro_actividad='" & recSetAuxcomp!pro_actividad & "'", db, adOpenKeyset, adLockReadOnly
      If rstipopy.RecordCount <> 0 Then
          tipopy = rstipopy!tipo_proyecto
      Else
         ' MsgBox "La Categoria Programática elegida no existe"
         MsgBox "Error en la contabilización, No se encontró la Estructura Programática", vbExclamation + vbDefaultButton1
         Exit Sub
      End If
      'VERIFICA FUENTE
      Select Case recSetAuxcomp!fte_codigo
        Case "10", "41"
            Select Case tipopy
              Case "N"
                    Set recSetPartida = New ADODB.Recordset
                    recSetPartida.CursorLocation = adUseClient
                    If recSetPartida.State = 1 Then recSetPartida.Close
                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
                      " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst= 'PAG' and CC_Cuenta_H1.Inst= 'PAG' and " & _
                      " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=1 AND " & _
                      " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    Sw_Fuente = True
              Case "S"
                    Set recSetPartida = New ADODB.Recordset
                    recSetPartida.CursorLocation = adUseClient
                    If recSetPartida.State = 1 Then recSetPartida.Close
                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
                      " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst= 'PSP' and CC_Cuenta_H1.Inst= 'PSP' and " & _
                      " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=1 AND " & _
                      " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    Sw_Fuente = True
              Case "F"
                    Set recSetPartida = New ADODB.Recordset
                    recSetPartida.CursorLocation = adUseClient
                    If recSetPartida.State = 1 Then recSetPartida.Close
                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
                      " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst= 'PFP' and CC_Cuenta_H1.Inst= 'PFP' and " & _
                      " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=1 AND " & _
                      " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    Sw_Fuente = True
            End Select
    'Asignacion a variables
        Case "70", "43"
            Select Case tipopy
              Case "N"
                    Set recSetPartida = New ADODB.Recordset
                    recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
                    If recSetPartida.State = 1 Then recSetPartida.Close
                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
                      " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PAG' and CC_Cuenta_H1.Inst='PAG' and " & _
                      " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=2 AND " & _
                      " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    Sw_Fuente = True
              Case "S"
                    Set recSetPartida = New ADODB.Recordset
                    recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
                    If recSetPartida.State = 1 Then recSetPartida.Close
                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
                      " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PSP' and CC_Cuenta_H1.Inst='PSP' and " & _
                      " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=2 AND " & _
                      " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    Sw_Fuente = True
              Case "F"
                    Set recSetPartida = New ADODB.Recordset
                    recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
                    If recSetPartida.State = 1 Then recSetPartida.Close
                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
                    " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PFP' and CC_Cuenta_H1.Inst='PFP' and " & _
                    " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=2 AND " & _
                    " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    Sw_Fuente = True
            End Select
    Case "80"
      Select Case tipopy
       Case "N"
          Set recSetPartida = New ADODB.Recordset
          recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
          If recSetPartida.State = 1 Then recSetPartida.Close
          recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H1, CC_Cuentas_D1" & _
          " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PAG' and CC_Cuenta_H1.Inst='PAG' and " & _
          " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=3 and  " & _
          " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
          Sw_Fuente = True
       Case "S"
          Set recSetPartida = New ADODB.Recordset
          recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
          If recSetPartida.State = 1 Then recSetPartida.Close
          recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H1, CC_Cuentas_D1" & _
          " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PSP' and CC_Cuenta_H1.Inst='PSP' and " & _
          " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=3 and  " & _
          " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
          Sw_Fuente = True
       Case "F"
          Set recSetPartida = New ADODB.Recordset
          recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
          If recSetPartida.State = 1 Then recSetPartida.Close
          recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H1, CC_Cuentas_D1" & _
          " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PFP' and CC_Cuenta_H1.Inst='PFP' and " & _
          " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=3 and  " & _
          " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
          Sw_Fuente = True
      End Select
    Case Else
        Sw_Fuente = False
        MsgBox "No esta asociado a ninguna fuente ... partida no relacionada "
        Exit Sub
    End Select
    If Sw_Fuente Then
        'Asignacion a variables
        v_Cuenta = recSetPartida!cuenta
        v_SubCta1 = recSetPartida!subcta1
        v_SubCta2 = recSetPartida!subcta2
        v_NombreCta = recSetPartida!NombreCta
        v_H_Cuenta = recSetPartida!h_cuenta
        v_H_SubCta1 = recSetPartida!h_subcta1
        v_H_SubCta2 = recSetPartida!h_subcta2
        v_H_NombCta = recSetPartida!H_NombCta
        
        v_Aux1 = recSetPartida!aux1
        v_Aux2 = recSetPartida!aux2
        v_Aux3 = recSetPartida!aux3
        
        v_H_Aux1 = recSetPartida!h_Aux1
        v_H_Aux2 = recSetPartida!h_Aux2
        v_H_Aux3 = recSetPartida!h_Aux3
        If recSetPartida.State = 1 Then recSetPartida.Close
        '************Abrimos un record set para adicionar datos*********************'
        Set recSetAuxActualizar = New ADODB.Recordset
        If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
        recSetAuxActualizar.Open " select * from CO_Comprobante_M  where Cod_Trans='" & P_codigo_pago & "' and Org_Codigo='" & P_org_codigo & "' " & _
          " and Ges_Gestion='" & P_ges_gestion & "' and Tipo_comp='PAC' and Cod_Trans_Detalle='" & P_codigo_pago_detalle & _
          "' and isnull(estado,'')<>'ANL' and isnull(estado,'')<>'DVL'", db, adOpenDynamic, adLockOptimistic, adCmdText
        '---- no importa que exista un comprobante
        If Not recSetAuxActualizar.BOF Then recSetAuxActualizar.MoveFirst
        If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
            '************* GENERA EL CODIGO DE COMPROBANTE**********
            Set recSetGenera = New ADODB.Recordset
            recSetGenera.CursorLocation = adUseClient
            If recSetGenera.State = 1 Then recSetGenera.Close
            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
            If recSetGenera.RecordCount > 0 Then
                Cont_Comp = Val(recSetGenera!numero_correlativo)
                Cont_Comp = Cont_Comp + 1
                recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
                recSetGenera.Update
            End If
            If recSetGenera.State = 1 Then recSetGenera.Close
            '************TERMINA GENERACION DE COMPROBANTE********
            ' Datos Para co_Comprobante
            recSetAuxActualizar.AddNew
            recSetAuxActualizar!Cod_Comp = Cont_Comp
            recSetAuxActualizar!cod_trans = recSetAuxcomp!codigo_pago
            recSetAuxActualizar!cod_trans_detalle = recSetAuxcomp!codigo_pago_detalle
            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
            recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
            recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
            recSetAuxActualizar!num_respaldo = recSetAuxcomp!Codigo_orden
            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
            recSetAuxActualizar!fecha_A = CDate(recSetAuxcomp!fecha_pago)
            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
            recSetAuxActualizar!tipo_comp = "PAC"
            recSetAuxActualizar!Status = "S"
            recSetAuxActualizar.Update
            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
            ' Datos Para co_Diario
            Set recSetAuxActualizar1 = New ADODB.Recordset
            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
            recSetAuxActualizar1.Open " select * from CO_Diario where  cod_Comp = " & Cont_Comp & " ", db, adOpenDynamic, adLockOptimistic, adCmdText
            If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
                recSetAuxActualizar1.AddNew
                recSetAuxActualizar1!tipo_comp = "PAC"
                recSetAuxActualizar1!d_cuenta = v_Cuenta
                recSetAuxActualizar1!D_Nombre = IIf(IsNull(v_NombreCta), "", v_NombreCta)
                recSetAuxActualizar1!d_subcta1 = v_SubCta1
                recSetAuxActualizar1!d_subcta2 = v_SubCta2
                recSetAuxActualizar1!d_Aux1 = v_Aux1
                recSetAuxActualizar1!d_Aux2 = v_Aux2
                recSetAuxActualizar1!d_Aux3 = v_Aux3
            '************* CONTABILIZA AUXILIAARES DEBITO
                Select Case v_Aux1
                  Case "01"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
                    recSetAuxActualizar1!d_des_Larga = IIf(IsNull(recsetAdicion!denominacion_beneficiario), "", recsetAdicion!denominacion_beneficiario)
                        
                  Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!d_des_Larga = IIf(IsNull(recsetAdicion!cta_descripcion_larga), "", recsetAdicion!cta_descripcion_larga)
                  Case Else
                End Select
                ''****************** finaliza sesion de auxiliares
                recSetAuxActualizar1!h_Aux1 = v_H_Aux1
                recSetAuxActualizar1!h_Aux2 = v_H_Aux2
                recSetAuxActualizar1!h_Aux3 = v_H_Aux3
                '************* CONTABILIZA AUXILIAARES CREDITO
                Select Case v_H_Aux1
                  Case "01"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
                    recSetAuxActualizar1!h_des_Larga = IIf(IsNull(recsetAdicion!denominacion_beneficiario), "", recsetAdicion!denominacion_beneficiario)
                  Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    'recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!h_des_Larga = IIf(IsNull(recsetAdicion!cta_descripcion_larga), "", recsetAdicion!cta_descripcion_larga)
                  Case Else
                End Select
                ''****************** finaliza sesion de auxiliares
                recSetAuxActualizar1!h_cuenta = v_H_Cuenta
                recSetAuxActualizar1!H_Nombre = IIf(IsNull(v_H_NombCta), "", v_H_NombCta)
                recSetAuxActualizar1!h_subcta1 = v_H_SubCta1
                recSetAuxActualizar1!h_subcta2 = v_H_SubCta2
                recSetAuxActualizar1!d_montoBs = recSetAuxcomp!monto_Bolivianos
                recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_dolares
                recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
                
                recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_Bolivianos
                recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_dolares
                recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
                recSetAuxActualizar1!Cod_Comp = Cont_Comp
                recSetAuxActualizar1.Update
            End If
        Else
          MsgBox "Ya fue contabilizado anteriormente...  ", vbOKOnly, "contabilizando...  "
          ' buscar el que ya existe y reemplazar los datos
          If (Not recSetAuxActualizar.BOF) Then recSetAuxActualizar.MoveFirst
          Cont_Comp = recSetAuxActualizar!Cod_Comp
          recSetAuxActualizar!cod_trans = recSetAuxcomp!codigo_pago
          recSetAuxActualizar!cod_trans_detalle = recSetAuxcomp!codigo_pago_detalle
          recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
          recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
          recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
          recSetAuxActualizar!num_respaldo = recSetAuxcomp!Codigo_orden
          recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
          recSetAuxActualizar!fecha_A = CDate(recSetAuxcomp!fecha_pago)
          recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
'            recSetAuxActualizar!Tipo_Comp = "PAC"
          recSetAuxActualizar!Status = "S"
          recSetAuxActualizar.Update
          If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
        ' Datos Para co_Diario
          Set recSetAuxActualizar1 = New ADODB.Recordset
          If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
          recSetAuxActualizar1.Open " select * from CO_Diario where  cod_Comp = " & Cont_Comp & " ", db, adOpenDynamic, adLockOptimistic, adCmdText
          If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
                recSetAuxActualizar1.AddNew
                recSetAuxActualizar1!tipo_comp = "PAC"
                recSetAuxActualizar1!Cod_Comp = Cont_Comp
          Else
                If (Not recSetAuxActualizar1.BOF) Then recSetAuxActualizar1.MoveFirst
          End If
          recSetAuxActualizar1!d_cuenta = v_Cuenta
          recSetAuxActualizar1!D_Nombre = v_NombreCta
          recSetAuxActualizar1!d_subcta1 = v_SubCta1
          recSetAuxActualizar1!d_subcta2 = v_SubCta2
          recSetAuxActualizar1!d_Aux1 = v_Aux1
          recSetAuxActualizar1!d_Aux2 = v_Aux2
          recSetAuxActualizar1!d_Aux3 = v_Aux3
          '************* CONTABILIZA AUXILIAARES DEBITO
          Select Case v_Aux1
             Case "01"
                 Set recsetAdicion = New ADODB.Recordset
                 If recsetAdicion.State = 1 Then recsetAdicion.Close
                 recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                 recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
                 recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
                Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
                Case Else
                End Select
''****************** finaliza sesion de auxiliares
                recSetAuxActualizar1!h_Aux1 = v_H_Aux1
                recSetAuxActualizar1!h_Aux2 = v_H_Aux2
                recSetAuxActualizar1!h_Aux3 = v_H_Aux3
'************* CONTABILIZA AUXILIAARES CREDITO
           
                Select Case v_H_Aux1
                Case "01"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                
                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
                        
                Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga
            
                Case Else
                End Select
''****************** finaliza sesion de auxiliares
               
                recSetAuxActualizar1!h_cuenta = v_H_Cuenta
                recSetAuxActualizar1!H_Nombre = v_H_NombCta
                recSetAuxActualizar1!h_subcta1 = v_H_SubCta1
                recSetAuxActualizar1!h_subcta2 = v_H_SubCta2
                recSetAuxActualizar1!d_montoBs = recSetAuxcomp!monto_Bolivianos
                recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_dolares
                recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
                
                recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_Bolivianos
                recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_dolares
                recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
                recSetAuxActualizar1.Update
        End If
    Else
           MsgBox "No esta asociado a ninguna fuente ...  "
    End If
    recSetAuxcomp.MoveNext
MsgBox "Contabilizacion exitosa...... ", vbOKOnly, "Contabilizacion"
Wend
db.CommitTrans


    Set recSetAuxcomp = New ADODB.Recordset
    recSetAuxcomp.CursorLocation = adUseClient
    If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
    
    Set recSetAuxActualizar = New ADODB.Recordset
    If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
    
    Set recSetAuxActualizar1 = New ADODB.Recordset
    If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
    
    Set recSetPartida = New ADODB.Recordset
    recSetPartida.CursorLocation = adUseClient
    If recSetPartida.State = 1 Then recSetPartida.Close
Exit Sub
errorPag:
db.RollbackTrans
MsgBox "No se contabilizó ... "

End Sub
Public Sub Anulacion_DAC(pagos)
Dim comanulDAC As ADODB.Command
 Set comanulDAC = New ADODB.Command ' para obtener los saldos
        With comanulDAC
            .CommandType = adCmdStoredProc
            .CommandText = "Anulacion"
            .Parameters.Append comanulDAC.CreateParameter("cod", adInteger, adParamInput)
            .Parameters.Append comanulDAC.CreateParameter("org", adVarChar, adParamInput, 3)
            .Parameters.Append comanulDAC.CreateParameter("gestion", adVarChar, adParamInput, 4)
            .Parameters.Append comanulDAC.CreateParameter("usr", adVarChar, adParamInput, 40)
            .Parameters.Append comanulDAC.CreateParameter("hora", adVarChar, adParamInput, 8)
            .Parameters.Append comanulDAC.CreateParameter("registro", adInteger, adParamOutput)
            .Parameters.Append comanulDAC.CreateParameter("numero", adInteger, adParamOutput)
            .Parameters("Cod") = pagos("Nro_Comprobante_Anterior")
            .Parameters("org") = pagos("org_codigo")
            .Parameters("gestion") = pagos("ges_gestion")
            .Parameters("usr") = GlUsuario
            .Parameters("hora") = Format(Time, "hh:mm:ss")
'           .Parameters("registro")=
            .ActiveConnection = db
            .Execute
            regANL = IIf(IsNull(.Parameters("registro")), 0, .Parameters("registro"))
            numANL = IIf(IsNull(.Parameters("numero")), 0, .Parameters("numero"))
        End With
End Sub

'Public Sub Anulacion_DAC(pagos)
'    'Comprobantes PAC
''  db.BeginTrans
'    Dim rsCoCoM As ADODB.Recordset
'    Set rsCoCoM = New ADODB.Recordset
'    If rsCoCoM.State = 1 Then rsCoCoM.Close
'    rsCoCoM.CursorLocation = adUseClient
'    rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & pagos("Nro_Comprobante_Anterior") & "' and org_codigo='" & pagos("org_codigo") & "' and Tipo_Comp='PAC'", db, adOpenKeyset, adLockOptimistic
'    If rsCoCoM.RecordCount > 0 Then
'        '             Set rsCoCoM = New ADODB.Recordset
'        '            If rsCoCoM.State = 1 Then rsCoCoM.Close
'        '            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC'", db, adOpenKeyset, adLockOptimistic
'        '            If rsCoCoM.RecordCount > 0 Then
''               'Creación de la cabecera o registros maestro en CO_COMPROBANTE_M
'    'Recuperando datos de co_comprobante_m
'    cocmCod_CompDiario = IIf(IsNull(rsCoCoM("Cod_Comp")), " ", rsCoCoM("Cod_Comp"))
''    cocmTipo_Comp = IIf(IsNull(rscocom("Tipo_Comp")), " ", rscocom("Tipo_Comp"))
'    cocmTipo_Comp = "ANL"
'    cocmCod_Trans = IIf(IsNull(rsCoCoM("Cod_Trans")), " ", rsCoCoM("cod_trans"))
'    cocmCod_Trans_Detalle = IIf(IsNull(rsCoCoM("Cod_Trans_Detalle")), "", (rsCoCoM("Cod_Trans_Detalle")))
'    cocmOrg_Codigo = IIf(IsNull(rsCoCoM("Org_Codigo")), "", rsCoCoM("Org_Codigo"))
'    cocmGes_Gestion = IIf(IsNull(rsCoCoM("Ges_Gestion")), "", rsCoCoM("Ges_Gestion"))
'    cocmNum_Respaldo = IIf(IsNull(rsCoCoM("Num_Respaldo")), "", rsCoCoM("Num_Respaldo"))
'    cocmFecha_A = CDate(IIf(IsNull(rsCoCoM("Fecha_A")), CDate(Date), rsCoCoM("Fecha_A")))
'    cocmCodigo_Beneficiario = IIf(IsNull(rsCoCoM("Codigo_Beneficiario")), "", rsCoCoM("Codigo_Beneficiario"))
'    cocmCodigo_Documento = IIf(IsNull(rsCoCoM("Codigo_Documento")), "", rsCoCoM("Codigo_Documento"))
'    cocmGlosa = IIf(IsNull(rsCoCoM("Glosa")), "", rsCoCoM("Glosa"))
'    cocmStatus = IIf(IsNull(rsCoCoM("Status")), "", rsCoCoM("Status"))
'    cocmUsr_usuario = IIf(IsNull(rsCoCoM("Usr_Usuario")), "", rsCoCoM("Usr_Usuario"))
'    'Adicionando un nuevo registro
'    'Generando nuevo código
'    'Segunda genera*********
'            Set rsCorr = New ADODB.Recordset
'            If rsCorr.State = 1 Then rsCorr.Close
'            rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'            If rsCorr.RecordCount > 0 Then
'                cocmCod_Comp = rsCorr("numero_correlativo") + 1
'                rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'                rsCorr.Update
'            End If
'            rsCorr.Close
'            'MsgBox "NUMERO DE 1era. CUENTA PAC" & cocmCod_Comp
'    rsCoCoM.AddNew
'
'        rsCoCoM("Cod_Comp") = cocmCod_Comp
'        rsCoCoM("Tipo_Comp") = Trim(cocmTipo_Comp)
'        rsCoCoM("Cod_Trans") = Trim(cocmCod_Trans)
'        rsCoCoM("Cod_Trans_Detalle") = Trim(cocmCod_Trans_Detalle)
'        rsCoCoM("org_codigo") = Trim(cocmOrg_Codigo)
'        rsCoCoM("Ges_Gestion") = Trim(cocmGes_Gestion)
'        rsCoCoM("Num_Respaldo") = Trim(cocmNum_Respaldo)
'        rsCoCoM("Fecha_A") = CDate(cocmFecha_A)
'        rsCoCoM("Codigo_Beneficiario") = Trim(cocmCodigo_Beneficiario)
'        rsCoCoM("Codigo_Documento") = Trim(cocmCodigo_Documento)
'        rsCoCoM("Glosa") = Trim(cocmGlosa)
'        rsCoCoM("Status") = Trim(cocmStatus)
'        rsCoCoM("usr_usuario") = GlUsuario
'        rsCoCoM("fecha_registro") = CDate(Format(Date, "dd/mm/yyyy"))
'        rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
'    rsCoCoM.Update
'        Set rsdiario = New ADODB.Recordset
'        If rsdiario.State = 1 Then rsdiario.Close
'        'rsDiario.Open "select * from co_Diario where Cod_Comp=" & rsCoCoM("Cod_Comp") & "", db, adOpenKeyset, adLockOptimistic
'        rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
'        If rsdiario.RecordCount > 0 Then
''                        'Recuperando datos
''                        Set rsCorr = New ADODB.Recordset
''                        If rsCorr.State = 1 Then rsCorr.Close
''                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
''                        If rsCorr.RecordCount > 0 Then
''                            AuxCod_Comp = rsCorr("numero_correlativo") + 1
''                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
''                            rsCorr.Update
''                        End If
'            'AuxCod_Comp_C = rsDiario("Cod_Comp_C")
'            AuxCod_Comp = cocmCod_Comp
'            AuxTipo_Comp = "ANL"
'            'AuxTipo_Comp = IIf(IsNull(rsdiario("Tipo_Comp")), "", rsdiario("Tipo_Comp"))
'            AuxCod_Comp_C = IIf(IsNull(cocmCod_Comp_C), 0, cocmCod_Comp_C)
'            AuxD_Cuenta = rsdiario("D_Cuenta")
'            AuxD_Nombre = IIf(IsNull(rsdiario("D_Nombre")), "", rsdiario("D_Nombre"))
'            AuxD_SubCta1 = rsdiario("D_SubCta1")
'            AuxD_SubCta2 = rsdiario("D_SubCta2")
'            AuxD_Aux1 = rsdiario("D_Aux1")
'            AuxD_Aux2 = rsdiario("D_Aux2")
'            AuxD_Aux3 = rsdiario("D_Aux3")
'            AuxD_Cta_Larga = IIf(IsNull(rsdiario("D_Cta_Larga")), "", rsdiario("D_Cta_Larga"))
'            AuxD_Des_Larga = IIf(IsNull(rsdiario("D_Des_Larga")), "", rsdiario("D_Des_Larga"))
'            AuxD_MontoBs = rsdiario("D_MontoBs")
'            AuxD_MontoDL = rsdiario("D_MontoDL")
'            AuxD_Cambio = rsdiario("D_Cambio")
'
'            AuxH_Cuenta = rsdiario("H_Cuenta")
'            AuxH_Nombre = IIf(IsNull(rsdiario("H_Nombre")), "", rsdiario("H_Nombre"))
'            AuxH_SubCta1 = rsdiario("H_SubCta1")
'            AuxH_SubCta2 = rsdiario("H_SubCta2")
'            AuxH_Aux1 = rsdiario("H_Aux1")
'            AuxH_Aux2 = rsdiario("H_Aux2")
'            AuxH_Aux3 = rsdiario("H_Aux3")
'            AuxH_Cta_Larga = IIf(IsNull(rsdiario("H_Cta_Larga")), "", rsdiario("H_Cta_Larga"))
'            AuxH_Des_Larga = IIf(IsNull(rsdiario("H_Des_Larga")), "", rsdiario("H_Des_Larga"))
'            AuxH_MontoBs = rsdiario("H_MontoBs")
'            AuxH_MontoDL = rsdiario("H_MontoDL")
'            AuxH_Cambio = rsdiario("H_Cambio")
'
'            AuxUsr_Usuario = IIf(IsNull(rsdiario("Usr_Usuario")), "", rsdiario("Usr_Usuario"))
'            AuxFecha_Registro = CDate(IIf(IsNull(rsdiario("Fecha_Registro")), CDate(Date), rsdiario("Fecha_Registro")))
'            AuxHora_Registro = IIf(IsNull(rsdiario("Hora_Registro")), Time, rsdiario("Hora_Registro"))
'
'            'Adicionando una copia del registro
'            rsdiario.AddNew
'            rsdiario("Cod_Comp") = AuxCod_Comp
'            rsdiario("Tipo_Comp") = Trim(AuxTipo_Comp)
'            rsdiario("Cod_Comp_C") = AuxCod_Comp_C
'
'            rsdiario("D_Cuenta") = AuxH_Cuenta
'            rsdiario("D_Nombre") = IIf(IsNull(AuxH_Nombre), "", AuxH_Nombre)
'            rsdiario("D_SubCta1") = AuxH_SubCta1
'            rsdiario("D_SubCta2") = AuxH_SubCta2
'            rsdiario("D_Aux1") = AuxH_Aux1
'            rsdiario("D_Aux2") = AuxH_Aux2
'            rsdiario("D_Aux3") = AuxH_Aux3
'            rsdiario("D_Cta_Larga") = IIf(IsNull(AuxH_Cta_Larga), "", AuxH_Cta_Larga)
'            rsdiario("D_Des_Larga") = IIf(IsNull(AuxH_Des_Larga), "", AuxH_Des_Larga)
'            rsdiario("D_MontoBs") = AuxH_MontoBs
'            rsdiario("D_MontoDL") = AuxH_MontoDL
'            rsdiario("D_Cambio") = AuxH_Cambio
'
'            rsdiario("H_Cuenta") = AuxD_Cuenta
'            rsdiario("H_Nombre") = IIf(IsNull(AuxD_Nombre), "", AuxD_Nombre)
'            rsdiario("H_SubCta1") = AuxD_SubCta1
'            rsdiario("H_SubCta2") = AuxD_SubCta2
'            rsdiario("H_Aux1") = AuxD_Aux1
'            rsdiario("H_Aux2") = AuxD_Aux2
'            rsdiario("H_Aux3") = AuxD_Aux3
'            rsdiario("H_Cta_Larga") = IIf(IsNull(AuxD_Cta_Larga), "", AuxD_Cta_Larga)
'            rsdiario("H_Des_Larga") = IIf(IsNull(AuxD_Des_Larga), "", AuxD_Des_Larga)
'            rsdiario("H_MontoBs") = AuxD_MontoBs
'            rsdiario("H_MontoDL") = AuxD_MontoDL
'            rsdiario("H_Cambio") = AuxD_Cambio
'
'            rsdiario("Usr_Usuario") = AuxUsr_Usuario
'            rsdiario("Fecha_Registro") = CDate(AuxFecha_Registro)
'            rsdiario("Hora_Registro") = Format(AuxHora_Registro, "hh:mm:ss")
'            rsdiario.Update
'    End If
'      Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
'  End If
'' db.CommitTrans
'End Sub

Public Sub Cmd_ContaConf(P_codigo_pago As String, P_org_codigo As String, P_ges_gestion As String)
'Private Sub Cmd_ContaConf_Click()
Dim Sw As Boolean
Dim Sw_Fuente As Boolean
Dim Cont_Comp As Long
Dim rstipopy As ADODB.Recordset
Set rstipopy = New ADODB.Recordset
Dim aux_T As String

'On Error GoTo errorComp
db.BeginTrans


'********* Para obtener en el recordset recsetAuxComp los datos necesarios para almacenar*********"
    Set recSetAuxcomp = New ADODB.Recordset
    recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
    If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
    recSetAuxcomp.Open "SELECT distinct pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion," & _
    " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pagos.Monto_Bolivianos as Monto_total,Pagos.org_Codigo,pagos.Codigo_orden,Pagos.Codigo_documento," & _
    " pago_detalle.pro_programa, pago_detalle.pro_subprograma, pago_detalle.pro_proyecto, pago_detalle.pro_actividad, " & _
    " pagos.Monto_Dolares,pago_detalle.Tipo_Cambio,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and TIPO_COMP='DAC' AND " & _
    " pago_detalle.Ges_Gestion = pagos.Ges_Gestion AND  pagos.codigo_Pago=" & P_codigo_pago & " and pagos.Org_Codigo='" & P_org_codigo & "' and pago_detalle.Ges_Gestion = '" & P_ges_gestion & "'", db, adOpenKeyset, adLockOptimistic

    If recSetAuxcomp.RecordCount > 0 Then recSetAuxcomp.MoveFirst
While Not (recSetAuxcomp.EOF)
      If rstipopy.State = 1 Then rstipopy.Close
      Dim sqlpy  As String
      Dim tipopy As String
      rstipopy.Open "select tipo_proyecto from fc_estructura_programatica where Pro_programa='" & recSetAuxcomp!pro_programa & "' and  Pro_subprograma='" & recSetAuxcomp!pro_subprograma & "' and Pro_proyecto='" & recSetAuxcomp!pro_proyecto & "' and Pro_actividad='" & recSetAuxcomp!pro_actividad & "'", db, adOpenKeyset, adLockReadOnly
      If rstipopy.RecordCount <> 0 Then
          tipopy = rstipopy!tipo_proyecto
      Else
         ' MsgBox "La Categoria Programática elegida no existe"
         MsgBox "Error en la contabilización, No se encontró la Estructura Programática", vbExclamation + vbDefaultButton1
         Exit Sub
      End If
      Set recSetPartida = New ADODB.Recordset
      If recSetPartida.State = 1 Then recSetPartida.Close
      Select Case tipopy
        Case "N"
            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H,CC_Cuentas_D" & _
                " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='DEV' and CC_Cuenta_H.Inst='DEV' and" & _
                " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText

        Case "S"
            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H,CC_Cuentas_D" & _
                " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PSD' and CC_Cuenta_H.Inst='PSD' and" & _
                " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText

        Case "F"
            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H,CC_Cuentas_D" & _
                " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PFD' and CC_Cuenta_H.Inst='PFD' and" & _
                " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
      End Select

    If recSetPartida.BOF And recSetPartida.EOF Then
        MsgBox "No Existe Partida"
    Else


        '************Abrimos un record set para adicionar datos*********************'

        Set recSetAuxActualizar = New ADODB.Recordset
        If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
        recSetAuxActualizar.Open " select * from CO_Comprobante_M  where Cod_Trans=" & recSetAuxcomp!codigo_pago & " and Org_Codigo='" & recSetAuxcomp!org_codigo & "' " & _
        " and Ges_Gestion='" & recSetAuxcomp!ges_gestion & "' and tipo_comp='DAC' and Cod_Trans_Detalle='" & recSetAuxcomp!codigo_pago_detalle & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
        'MsgBox recSetAuxActualizar.RecordCount
        If Not recSetAuxActualizar.BOF Then recSetAuxActualizar.MoveFirst
      If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
        '************* GENERA EL CODIGO DE COMPROBANTE**********
            Set recSetGenera = New ADODB.Recordset
            recSetGenera.CursorLocation = adUseClient
            If recSetGenera.State = 1 Then recSetGenera.Close
            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
            If recSetGenera.RecordCount > 0 Then
                Cont_Comp = Val(recSetGenera!numero_correlativo)
                Cont_Comp = Cont_Comp + 1
                recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
                recSetGenera.Update
            End If
            If recSetGenera.State = 1 Then recSetGenera.Close
'************TERMINA GENERACION DE COMPROBANTE********
' Datos Para co_Comprobante

            recSetAuxActualizar.AddNew

            recSetAuxActualizar!usr_usuario = GlUsuario
            recSetAuxActualizar!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")

            recSetAuxActualizar!Cod_Comp = Cont_Comp
            recSetAuxActualizar!cod_trans = recSetAuxcomp!codigo_pago
            recSetAuxActualizar!cod_trans_detalle = recSetAuxcomp!codigo_pago_detalle
            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
            recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
            recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
            recSetAuxActualizar!num_respaldo = recSetAuxcomp!Codigo_orden
            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
            recSetAuxActualizar!fecha_A = IIf(IsNull(recSetAuxcomp!fecha_pago), (Format(Date, "dd/mm/yyyy")), CDate(recSetAuxcomp!fecha_pago))
            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
            recSetAuxActualizar!tipo_comp = "DAC"
            recSetAuxActualizar!Status = "S"
            recSetAuxActualizar.Update
            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close



        Set recSetAuxActualizar1 = New ADODB.Recordset
        If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
        recSetAuxActualizar1.Open " select * from CO_Diario where cod_Comp = " & Cont_Comp & " ", db, adOpenDynamic, adLockOptimistic, adCmdText
        If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
        recSetAuxActualizar1.AddNew

            recSetAuxActualizar1!usr_usuario = GlUsuario
            recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")

            recSetAuxActualizar1!Cod_Comp = Cont_Comp
            recSetAuxActualizar1!tipo_comp = "DAC"
            recSetAuxActualizar1!d_cuenta = recSetPartida!cuenta
            recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
            recSetAuxActualizar1!d_subcta1 = recSetPartida!subcta1
            recSetAuxActualizar1!d_subcta2 = recSetPartida!subcta2
            recSetAuxActualizar1!d_Aux1 = recSetPartida!aux1
            recSetAuxActualizar1!d_Aux2 = recSetPartida!aux2
            recSetAuxActualizar1!d_Aux3 = recSetPartida!aux3

        ''******* ADICION DE AUXILIARES A DETALLE DEBITO*******
            Select Case recSetPartida!aux1
            Case "01"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
                    recSetAuxActualizar1!d_des_Larga = IIf(IsNull(recsetAdicion!denominacion_beneficiario), " ", recsetAdicion!denominacion_beneficiario)

            Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cTA_cODIGO='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
            Case Else
            End Select
        ''****************** finaliza sesion de auxiliares DEBITO

            recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
            recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
            recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
            recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2

            recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
            recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
            recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
        ''******* ADICION DE AUXILIARES A DETALLE*******
            Select Case recSetPartida!h_Aux1
            Case "01"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
                    recSetAuxActualizar1!h_des_Larga = IIf(IsNull(recsetAdicion!denominacion_beneficiario), "", recsetAdicion!denominacion_beneficiario)

            Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where CTA_CODIGO='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!h_des_Larga = IIf(IsNull(recsetAdicion!cta_descripcion_larga), "", recsetAdicion!cta_descripcion_larga)

            Case Else
            End Select
        ''****************** finaliza sesion de auxiliares


            recSetAuxActualizar1!d_montoBs = recSetAuxcomp!monto_total
            recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_dolares
            recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio

            recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_total
            recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_dolares
            recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
            recSetAuxActualizar1!usr_usuario = GlUsuario
            recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
            recSetAuxActualizar1.Update
            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close

        End If 'Adicion del diario
      Else
            MsgBox "Ya fue contabilizado anteriormente...  ", vbOKOnly, "contabilizando...  "
            'Modifica registro existente
            'recSetAuxActualizar!Cod_Comp = Cont_Comp
            recSetAuxActualizar!usr_usuario = GlUsuario
            Cont_Comp = recSetAuxActualizar!Cod_Comp
            recSetAuxActualizar!cod_trans = recSetAuxcomp!codigo_pago
            recSetAuxActualizar!cod_trans_detalle = recSetAuxcomp!codigo_pago_detalle
            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
            recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
            recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
            recSetAuxActualizar!num_respaldo = recSetAuxcomp!Codigo_orden
            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
'            If IsNull(recSetAuxcomp!fecha_pago) Then
'             FECHA = Date
'            Else
'             FECHA = recSetAuxcomp!fecha_pago
'            End If
            recSetAuxActualizar!fecha_A = IIf(IsNull(recSetAuxcomp!fecha_pago), Format(Date, "dd/mm/yyyy"), CDate(recSetAuxcomp!fecha_pago))
            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
            recSetAuxActualizar!usr_usuario = GlUsuario
            recSetAuxActualizar!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
            'recSetAuxActualizar!Tipo_Comp = "DAC"
            recSetAuxActualizar!Status = "S"
            recSetAuxActualizar.Update
            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
   'Modificacione en el diario

            Set recSetAuxActualizar1 = New ADODB.Recordset
            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
            recSetAuxActualizar1.Open " select * from CO_Diario where cod_Comp = " & Cont_Comp, db, adOpenDynamic, adLockOptimistic  ', adCmdText
            If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
                recSetAuxActualizar1.AddNew
                recSetAuxActualizar1!tipo_comp = "DAC"
                recSetAuxActualizar1!Cod_Comp = Cont_Comp
            Else
                    If (Not recSetAuxActualizar1.BOF) Then recSetAuxActualizar1.MoveFirst
            End If
            recSetAuxActualizar1!usr_usuario = GlUsuario + "MOD"
            recSetAuxActualizar1!d_cuenta = recSetPartida!cuenta
            recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
            recSetAuxActualizar1!d_subcta1 = recSetPartida!subcta1
            recSetAuxActualizar1!d_subcta2 = recSetPartida!subcta2
            recSetAuxActualizar1!d_Aux1 = recSetPartida!aux1
            recSetAuxActualizar1!d_Aux2 = recSetPartida!aux2
            recSetAuxActualizar1!d_Aux3 = recSetPartida!aux3

        ''******* ADICION DE AUXILIARES A DETALLE DEBITO*******
            Select Case recSetPartida!aux1
            Case "01"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
                    recSetAuxActualizar1!d_des_Larga = IsNull(recsetAdicion!denominacion_beneficiario)

            Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cTA_cODIGO='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
            Case Else
            End Select
        ''****************** finaliza sesion de auxiliares DEBITO
            recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
            recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
            recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
            recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2

            recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
            recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
            recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
        ''******* ADICION DE AUXILIARES A DETALLE*******
            Select Case recSetPartida!h_Aux1
            Case "01"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario

            Case "02"
                    Set recsetAdicion = New ADODB.Recordset
                    If recsetAdicion.State = 1 Then recsetAdicion.Close
                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where CTA_CODIGO='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga

            Case Else
            End Select
        ''****************** finaliza sesion de auxiliares


            recSetAuxActualizar1!d_montoBs = recSetAuxcomp!monto_total
            recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_dolares
            recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio

            recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_total
            recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_dolares
            recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
            recSetAuxActualizar1!usr_usuario = GlUsuario
            recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")

            recSetAuxActualizar1.Update
            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close


      End If

    End If ' De Partida
recSetAuxcomp.MoveNext

Wend


db.CommitTrans
MsgBox "Contabilizo....."
'Unload Frm_Cont_Mat


Exit Sub
errorComp:

db.RollbackTrans
MsgBox "error al recuperar datos"
'Unload Frm_Cont_Mat

End Sub
Public Sub Devolucion_PAC_DAC(pagos)
    'Devolución contablemente
    'recogiendo los datos de devolucion Nro de comprobante al que pertenece la devolución
    Dim rsCoCoM As ADODB.Recordset
    Set rsdev = New ADODB.Recordset
    If rsdev.State = 1 Then rsdev.Close
    rsdev.Open "select * from pagos where codigo_pago=" & pagos("codigo_pago") & " and org_codigo='" & pagos("org_codigo") & "' and ges_gestion='" & pagos("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
    If rsdev.RecordCount > 0 Then
            Set rsCoCoM = New ADODB.Recordset
            If rsCoCoM.State = 1 Then rsCoCoM.Close
            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and (Tipo_Comp='DAC') ", db, adOpenKeyset, adLockOptimistic
            If rsCoCoM.RecordCount > 0 Then
                'Creación de la cabecera o registros maestro en CO_COMPROBANTE_M
                'Recuperando datos de co_comprobante_m
                cocmCod_CompDiario = rsCoCoM("Cod_Comp")
                cocmTipo_Comp = "DVL"
                'cocmTipo_Comp = rscocom("Tipo_Comp")
                cocmCod_Trans = rsCoCoM("Cod_Trans") 'pagos("codigo_pago") 'TxtComprobante.text TxtNC.Text '
                cocmCod_Trans_Detalle = rsCoCoM("Cod_Trans_Detalle")
                cocmOrg_Codigo = rsCoCoM("Org_Codigo")
                cocmGes_Gestion = rsCoCoM("Ges_Gestion")
                cocmNum_Respaldo = rsCoCoM("Num_Respaldo")
                cocmFecha_A = CDate(rsCoCoM("Fecha_A"))
                cocmCodigo_Beneficiario = rsCoCoM("Codigo_Beneficiario")
                cocmCodigo_Documento = rsCoCoM("Codigo_Documento")
                cocmGlosa = rsCoCoM("Glosa")
                cocmStatus = rsCoCoM("Status")
                cocmUsr_usuario = rsCoCoM("Usr_Usuario")
                'Adicionando un nuevo registro
                'Generando nuevo código
                        Set rsCorr = New ADODB.Recordset
                        If rsCorr.State = 1 Then rsCorr.Close
                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
                        If rsCorr.RecordCount > 0 Then
                            cocmCod_Comp = rsCorr("numero_correlativo") + 1
                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
                            rsCorr.Update
                        End If
                        'MsgBox "NUMERO DE 1era. CUENTA DAC" & cocmCod_Comp
                        rsCorr.Close
                rsCoCoM.AddNew
                    rsCoCoM("Cod_Comp") = cocmCod_Comp
                    rsCoCoM("Tipo_Comp") = cocmTipo_Comp
                    rsCoCoM("Cod_Trans") = cocmCod_Trans 'pagos("codigo_pago") 'TxtNC.Text 'cocmCod_Trans
                    rsCoCoM("Cod_Trans_Detalle") = cocmCod_Trans_Detalle
                    rsCoCoM("org_codigo") = cocmOrg_Codigo
                    rsCoCoM("Ges_Gestion") = cocmGes_Gestion
                    rsCoCoM("Num_Respaldo") = cocmNum_Respaldo
                    rsCoCoM("Fecha_A") = CDate(cocmFecha_A)
                    rsCoCoM("Codigo_Beneficiario") = cocmCodigo_Beneficiario
                    rsCoCoM("Codigo_Documento") = cocmCodigo_Documento
                    rsCoCoM("Glosa") = cocmGlosa
                    rsCoCoM("Status") = cocmStatus
                    rsCoCoM("usr_usuario") = GlUsuario
                    rsCoCoM("fecha_registro") = CDate(Format(Date, "DD/MM/YYYY"))
                    rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
                rsCoCoM.Update
                
                Set rsdiario = New ADODB.Recordset
                If rsdiario.State = 1 Then rsdiario.Close
                rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
                'rsDiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_Comp & "", db, adOpenKeyset, adLockOptimistic
                If rsdiario.RecordCount > 0 Then
                    AuxCod_Comp = cocmCod_Comp
                    'AuxTipo_Comp = rsdiario("Tipo_Comp")
                    AuxTipo_Comp = "DVL"
                    AuxCod_Comp_C = IIf(IsNull(rsdiario("Cod_Comp_C")), 0, rsdiario("Cod_Comp_C"))
                    AuxD_Cuenta = rsdiario("D_Cuenta")
                    AuxD_Nombre = IIf(IsNull(rsdiario("D_Nombre")), "", rsdiario("D_Nombre"))
                    AuxD_SubCta1 = rsdiario("D_SubCta1")
                    AuxD_SubCta2 = rsdiario("D_SubCta2")
                    AuxD_Aux1 = rsdiario("D_Aux1")
                    AuxD_Aux2 = rsdiario("D_Aux2")
                    AuxD_Aux3 = rsdiario("D_Aux3")
                    AuxD_Cta_Larga = IIf(IsNull(rsdiario("D_Cta_Larga")), "", rsdiario("D_Cta_Larga"))
                    AuxD_Des_Larga = IIf(IsNull(rsdiario("D_Des_Larga")), "", rsdiario("D_Des_Larga"))
                    AuxD_MontoBs = rsdiario("D_MontoBs")
                    AuxD_MontoDL = rsdiario("D_MontoDL")
                    AuxD_Cambio = rsdiario("D_Cambio")
                    
                    AuxH_Cuenta = rsdiario("H_Cuenta")
                    AuxH_Nombre = IIf(IsNull(rsdiario("H_Nombre")), "", rsdiario("H_Nombre"))
                    AuxH_SubCta1 = rsdiario("H_SubCta1")
                    AuxH_SubCta2 = rsdiario("H_SubCta2")
                    AuxH_Aux1 = rsdiario("H_Aux1")
                    AuxH_Aux2 = rsdiario("H_Aux2")
                    AuxH_Aux3 = rsdiario("H_Aux3")
                    AuxH_Cta_Larga = IIf(IsNull(rsdiario("H_Cta_Larga")), "", rsdiario("H_Cta_Larga"))
                    AuxH_Des_Larga = IIf(IsNull(rsdiario("H_Des_Larga")), "", rsdiario("H_Des_Larga"))
                    AuxH_MontoBs = rsdiario("H_MontoBs")
                    AuxH_MontoDL = rsdiario("H_MontoDL")
                    AuxH_Cambio = rsdiario("H_Cambio")
                    
                    AuxUsr_Usuario = rsdiario("Usr_Usuario")
                    AuxFecha_Registro = CDate(Format(Date, "DD/MM/YYYY"))
                    AuxHora_Registro = Format(Time, "hh:mm:ss")
                   
                    'Adicionando una copia del registro
                    rsdiario.AddNew
                    rsdiario("Cod_Comp") = AuxCod_Comp 'AuxCod_Comp_C
                    rsdiario("Tipo_Comp") = AuxTipo_Comp
                    rsdiario("Cod_Comp_C") = AuxCod_Comp_C
                    rsdiario("D_Cuenta") = AuxH_Cuenta
                    'rsdiario("D_Nombre") = AuxH_Nombre
                    rsdiario("D_SubCta1") = AuxH_SubCta1
                    rsdiario("D_SubCta2") = AuxH_SubCta2
                    rsdiario("D_Aux1") = AuxH_Aux1
                    rsdiario("D_Aux2") = AuxH_Aux2
                    rsdiario("D_Aux3") = AuxH_Aux3
                    rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
                    'rsdiario("D_Cta_Larga") = AuxH_Des_Larga
                    rsdiario("D_MontoBs") = AuxH_MontoBs
                    rsdiario("D_MontoDL") = AuxH_MontoDL
                    rsdiario("D_Cambio") = AuxH_Cambio
                    
                    rsdiario("H_Cuenta") = AuxD_Cuenta
                    'rsdiario("H_Nombre") = AuxD_Nombre
                    rsdiario("H_SubCta1") = AuxD_SubCta1
                    rsdiario("H_SubCta2") = AuxD_SubCta2
                    rsdiario("H_Aux1") = AuxD_Aux1
                    rsdiario("H_Aux2") = AuxD_Aux2
                    rsdiario("H_Aux3") = AuxD_Aux3
                    rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
                    'rsdiario("H_Cta_Larga") = AuxD_Des_Larga
                    rsdiario("H_MontoBs") = AuxD_MontoBs
                    rsdiario("H_MontoDL") = AuxD_MontoDL
                    rsdiario("H_Cambio") = AuxD_Cambio
                    
                    rsdiario("Usr_Usuario") = AuxUsr_Usuario
                    rsdiario("Fecha_Registro") = CDate(AuxFecha_Registro)
                    rsdiario("Hora_Registro") = AuxHora_Registro
                    rsdiario.Update
                End If
                'Comprobantes PAC
                If rsCoCoM.State = 1 Then rsCoCoM.Close
                rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='PAC' or Tipo_Comp='CAP'", db, adOpenKeyset, adLockOptimistic
                
                If rsCoCoM.RecordCount > 0 Then
                
'                Set rsCoCoM = New ADODB.Recordset
'                If rsCoCoM.State = 1 Then rsCoCoM.Close
'                rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC'", db, adOpenKeyset, adLockOptimistic
            If rsCoCoM.RecordCount > 0 Then
                'Creación de la cabecera o registros maestro en CO_COMPROBANTE_M
                'Recuperando datos de co_comprobante_m
                cocmCod_CompDiario = rsCoCoM("Cod_Comp")
                cocmTipo_Comp = "DVL"
                ''cocmTipo_Comp = rscocom("Tipo_Comp")
                cocmCod_Trans = rsCoCoM("Cod_Trans") 'pagos("codigo_pago") 'TxtNC.Text 'rsCoCoM("Cod_Trans")
                cocmCod_Trans_Detalle = rsCoCoM("Cod_Trans_Detalle")
                cocmOrg_Codigo = rsCoCoM("Org_Codigo")
                cocmGes_Gestion = rsCoCoM("Ges_Gestion")
                cocmNum_Respaldo = rsCoCoM("Num_Respaldo")
                cocmFecha_A = CDate(rsCoCoM("Fecha_A"))
                cocmCodigo_Beneficiario = rsCoCoM("Codigo_Beneficiario")
                cocmCodigo_Documento = rsCoCoM("Codigo_Documento")
                cocmGlosa = rsCoCoM("Glosa")
                cocmStatus = rsCoCoM("Status")
                cocmUsr_usuario = IIf(IsNull(rsCoCoM("Usr_Usuario")), "", rsCoCoM("Usr_Usuario"))
                'Adicionando un nuevo registro
                'Generando nuevo código
                'Segunda genera*********
                        Set rsCorr = New ADODB.Recordset
                        If rsCorr.State = 1 Then rsCorr.Close
                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
                        If rsCorr.RecordCount > 0 Then
                            cocmCod_Comp = rsCorr("numero_correlativo") + 1
                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
                            rsCorr.Update
                        End If
'                        MsgBox "NUMERO DE 2da. CUENTA PAC " & cocmCod_Comp
                        rsCorr.Close
                rsCoCoM.AddNew
                    
                    rsCoCoM("Cod_Comp") = cocmCod_Comp
                    rsCoCoM("Tipo_Comp") = cocmTipo_Comp
                    rsCoCoM("Cod_Trans") = cocmCod_Trans 'pagos("codigo_pago") 'TxtNC.Text 'cocmCod_Trans
                    rsCoCoM("Cod_Trans_Detalle") = cocmCod_Trans_Detalle
                    rsCoCoM("org_codigo") = cocmOrg_Codigo
                    rsCoCoM("Ges_Gestion") = cocmGes_Gestion
                    rsCoCoM("Num_Respaldo") = cocmNum_Respaldo
                    rsCoCoM("Fecha_A") = CDate(cocmFecha_A)
                    rsCoCoM("Codigo_Beneficiario") = cocmCodigo_Beneficiario
                    rsCoCoM("Codigo_Documento") = cocmCodigo_Documento
                    rsCoCoM("Glosa") = cocmGlosa
                    rsCoCoM("Status") = cocmStatus
                    rsCoCoM("usr_usuario") = GlUsuario
                    rsCoCoM("fecha_registro") = CDate(Date)
                    rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
                rsCoCoM.Update
                    Set rsdiario = New ADODB.Recordset
                    If rsdiario.State = 1 Then rsdiario.Close
                    'rsDiario.Open "select * from co_Diario where Cod_Comp=" & rsCoCoM("Cod_Comp") & "", db, adOpenKeyset, adLockOptimistic
                    rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
                    If rsdiario.RecordCount > 0 Then
'                        'Recuperando datos
'                        Set rsCorr = New ADODB.Recordset
'                        If rsCorr.State = 1 Then rsCorr.Close
'                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
'                        If rsCorr.RecordCount > 0 Then
'                            AuxCod_Comp = rsCorr("numero_correlativo") + 1
'                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
'                            rsCorr.Update
'                        End If
                        'AuxCod_Comp_C = rsDiario("Cod_Comp_C")
                        AuxCod_Comp = cocmCod_Comp
                        'AuxTipo_Comp = rsdiario("Tipo_Comp")
                        AuxTipo_Comp = "DVL"
                        AuxCod_Comp_C = cocmCod_Comp_C
                        AuxD_Cuenta = rsdiario("D_Cuenta")
                        'AuxD_Nombre = IIf(IsNull(rsdiario("D_Nombre")), "", rsdiario("D_Nombre"))
                        AuxD_SubCta1 = rsdiario("D_SubCta1")
                        AuxD_SubCta2 = rsdiario("D_SubCta2")
                        AuxD_Aux1 = rsdiario("D_Aux1")
                        AuxD_Aux2 = rsdiario("D_Aux2")
                        AuxD_Aux3 = rsdiario("D_Aux3")
                        AuxD_Cta_Larga = IIf(IsNull(rsdiario("D_Cta_Larga")), "", rsdiario("D_Cta_Larga"))
                        AuxD_Des_Larga = IIf(IsNull(rsdiario("D_Des_Larga")), "", rsdiario("D_Des_Larga"))
                        AuxD_MontoBs = rsdiario("D_MontoBs")
                        AuxD_MontoDL = rsdiario("D_MontoDL")
                        AuxD_Cambio = rsdiario("D_Cambio")
                        
                        AuxH_Cuenta = rsdiario("H_Cuenta")
                        'AuxH_Nombre = IIf(IsNull(rsdiario("H_Nombre")), "", rsdiario("H_Nombre"))
                        AuxH_SubCta1 = rsdiario("H_SubCta1")
                        AuxH_SubCta2 = rsdiario("H_SubCta2")
                        AuxH_Aux1 = rsdiario("H_Aux1")
                        AuxH_Aux2 = rsdiario("H_Aux2")
                        AuxH_Aux3 = rsdiario("H_Aux3")
                        AuxH_Cta_Larga = IIf(IsNull(rsdiario("H_Cta_Larga")), "", rsdiario("H_Cta_Larga"))
                        AuxH_Des_Larga = IIf(IsNull(rsdiario("H_Des_Larga")), "", rsdiario("H_Des_Larga"))
                        AuxH_MontoBs = rsdiario("H_MontoBs")
                        AuxH_MontoDL = rsdiario("H_MontoDL")
                        AuxH_Cambio = rsdiario("H_Cambio")
                        
                        AuxUsr_Usuario = IIf(IsNull(rsdiario("Usr_Usuario")), "", rsdiario("Usr_Usuario"))
                        AuxFecha_Registro = CDate(IIf(IsNull(rsdiario("Fecha_Registro")), CDate(Date), rsdiario("Fecha_Registro")))
                        AuxHora_Registro = IIf(IsNull(rsdiario("Hora_Registro")), Time, rsdiario("Hora_Registro"))
                       
                        'Adicionando una copia del registro
                        rsdiario.AddNew
                        rsdiario("Cod_Comp") = AuxCod_Comp
                        rsdiario("Tipo_Comp") = AuxTipo_Comp
                        rsdiario("Cod_Comp_C") = AuxCod_Comp_C
                        
                        rsdiario("D_Cuenta") = AuxH_Cuenta
                        'rsdiario("D_Nombre") = AuxH_Nombre
                        rsdiario("D_SubCta1") = AuxH_SubCta1
                        rsdiario("D_SubCta2") = AuxH_SubCta2
                        rsdiario("D_Aux1") = AuxH_Aux1
                        rsdiario("D_Aux2") = AuxH_Aux2
                        rsdiario("D_Aux3") = AuxH_Aux3
                        rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
                        'rsdiario("D_Des_Larga") = AuxH_Des_Larga
                        rsdiario("D_MontoBs") = AuxH_MontoBs
                        rsdiario("D_MontoDL") = AuxH_MontoDL
                        rsdiario("D_Cambio") = AuxH_Cambio
                        
                        rsdiario("H_Cuenta") = AuxD_Cuenta
                        'rsdiario("H_Nombre") = AuxD_Nombre
                        rsdiario("H_SubCta1") = AuxD_SubCta1
                        rsdiario("H_SubCta2") = AuxD_SubCta2
                        rsdiario("H_Aux1") = AuxD_Aux1
                        rsdiario("H_Aux2") = AuxD_Aux2
                        rsdiario("H_Aux3") = AuxD_Aux3
                        rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
                        'rsdiario("H_Des_Larga") = AuxD_Des_Larga
                        rsdiario("H_MontoBs") = AuxD_MontoBs
                        rsdiario("H_MontoDL") = AuxD_MontoDL
                        rsdiario("H_Cambio") = AuxD_Cambio
                        
                        rsdiario("Usr_Usuario") = AuxUsr_Usuario
                        rsdiario("Fecha_Registro") = CDate(AuxFecha_Registro)
                        rsdiario("Hora_Registro") = Format(AuxHora_Registro, "hh:mm:ss")
                        rsdiario.Update
                End If
                  Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
              End If
          Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
    End If
       Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
End If
End If
End Sub
Public Sub Devolucion_DAC(pagos)
    'Devolución contablemente
    'recogiendo los datos de devolucion Nro de comprobante al que pertenece la devolución
    Set rsdev = New ADODB.Recordset
    If rsdev.State = 1 Then rsdev.Close
    rsdev.Open "select * from pagos where codigo_pago='" & pagos("codigo_pago") & "' and org_codigo='" & pagos("org_codigo") & "' and ges_gestion='" & pagos("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
    If rsdev.RecordCount > 0 Then
            Set rsCoCoM = New ADODB.Recordset
            If rsCoCoM.State = 1 Then rsCoCoM.Close
            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC'", db, adOpenKeyset, adLockOptimistic
            If rsCoCoM.RecordCount > 0 Then
                Set rsdiario = New ADODB.Recordset
                If rsdiario.State = 1 Then rsdiario.Close
                rsdiario.Open "select * from co_Diario where Cod_Comp=" & rsCoCoM("Cod_Comp") & "", db, adOpenKeyset, adLockOptimistic
                If rsdiario.RecordCount > 0 Then
                    'Recuperando datos
                    Set rsCorr = New ADODB.Recordset
                    If rsCorr.State = 1 Then rsCorr.Close
                    rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
                    If rsCorr.RecordCount > 0 Then
                        AuxCod_Comp = rsCorr("numero_correlativo") + 1
                        rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
                        rsCorr.Update
                    End If
                    AuxTipo_Comp = rsdiario("Tipo_Comp")
                    AuxCod_Comp_C = rsdiario("Cod_Comp_C")
                    AuxD_Cuenta = rsdiario("D_Cuenta")
                    AuxD_Nombre = rsdiario("D_Nombre")
                    AuxD_SubCta1 = rsdiario("D_SubCta1")
                    AuxD_SubCta2 = rsdiario("D_SubCta2")
                    AuxD_Aux1 = rsdiario("D_Aux1")
                    AuxD_Aux2 = rsdiario("D_Aux2")
                    AuxD_Aux3 = rsdiario("D_Aux3")
                    AuxD_Cta_Larga = rsdiario("D_Cta_Larga")
                    AuxD_Des_Larga = rsdiario("D_Des_Larga")
                    AuxD_MontoBs = rsdiario("D_MontoBs")
                    AuxD_MontoDL = rsdiario("D_MontoDL")
                    AuxD_Cambio = rsdiario("D_Cambio")

                    AuxH_Cuenta = rsdiario("H_Cuenta")
                    AuxH_Nombre = rsdiario("H_Nombre")
                    AuxH_SubCta1 = rsdiario("H_SubCta1")
                    AuxH_SubCta2 = rsdiario("H_SubCta2")
                    AuxH_Aux1 = rsdiario("H_Aux1")
                    AuxH_Aux2 = rsdiario("H_Aux2")
                    AuxH_Aux3 = rsdiario("H_Aux3")
                    AuxH_Cta_Larga = rsdiario("H_Cta_Larga")
                    AuxH_Des_Larga = rsdiario("H_Des_Larga")
                    AuxH_MontoBs = rsdiario("H_MontoBs")
                    AuxH_MontoDL = rsdiario("H_MontoDL")
                    AuxH_Cambio = rsdiario("H_Cambio")

                    AuxUsr_Usuario = rsdiario("Usr_Usuario")
                    AuxFecha_Registro = rsdiario("Fecha_Registro")
                    AuxHora_Registro = rsdiario("Hora_Registro")

                    'Adicionando una copia del registro
                    rsdiario.AddNew
                    rsdiario("Cod_Comp") = AuxCod_Comp
                    rsdiario("Tipo_Comp") = AuxTipo_Comp
                    rsdiario("Cod_Comp_C") = AuxCod_Comp_C

                    rsdiario("D_Cuenta") = AuxH_Cuenta
                    rsdiario("D_Nombre") = AuxH_Nombre
                    rsdiario("D_SubCta1") = AuxH_SubCta1
                    rsdiario("D_SubCta2") = AuxH_SubCta2
                    rsdiario("D_Aux1") = AuxH_Aux1
                    rsdiario("D_Aux2") = AuxH_Aux2
                    rsdiario("D_Aux3") = AuxH_Aux3
                    rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
                    rsdiario("D_Cta_Larga") = AuxH_Des_Larga
                    rsdiario("D_MontoBs") = AuxH_MontoBs
                    rsdiario("D_MontoDL") = AuxH_MontoDL
                    rsdiario("D_Cambio") = AuxH_Cambio

                    rsdiario("H_Cuenta") = AuxD_Cuenta
                    rsdiario("H_Nombre") = AuxD_Nombre
                    rsdiario("H_SubCta1") = AuxD_SubCta1
                    rsdiario("H_SubCta2") = AuxD_SubCta2
                    rsdiario("H_Aux1") = AuxD_Aux1
                    rsdiario("H_Aux2") = AuxD_Aux2
                    rsdiario("H_Aux3") = AuxD_Aux3
                    rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
                    rsdiario("H_Cta_Larga") = AuxD_Des_Larga
                    rsdiario("H_MontoBs") = AuxD_MontoBs
                    rsdiario("H_MontoDL") = AuxD_MontoDL
                    rsdiario("H_Cambio") = AuxD_Cambio

                    rsdiario("Usr_Usuario") = AuxUsr_Usuario
                    rsdiario("Fecha_Registro") = AuxFecha_Registro
                    rsdiario("Hora_Registro") = AuxHora_Registro
                    rsdiario.Update

                End If
          Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
    End If
       Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
End If
End Sub
Public Sub Reversion_DAC(pagos)
    'Devolución contablemente
    'recogiendo los datos de devolucion Nro de comprobante al que pertenece la devolución
    Set rsdev = New ADODB.Recordset
    If rsdev.State = 1 Then rsdev.Close
    rsdev.Open "select * from pagos where codigo_pago=" & pagos("codigo_pago") & " and org_codigo='" & pagos("org_codigo") & "' and ges_gestion='" & pagos("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
    If rsdev.RecordCount > 0 Then
            Set rsCoCoM = New ADODB.Recordset
            If rsCoCoM.State = 1 Then rsCoCoM.Close
            'Verificar en PAC-DAC
            rsCoCoM.Open "select * from co_Comprobante_M where cod_trans='" & rsdev("Nro_Comprobante_Anterior") & "' and org_codigo='" & rsdev("org_codigo") & "' and Tipo_Comp='DAC' ", db, adOpenKeyset, adLockOptimistic
            If rsCoCoM.RecordCount > 0 Then
                'Creación de la cabecera o registros maestro en CO_COMPROBANTE_M
                'Recuperando datos de co_comprobante_m
                cocmCod_CompDiario = rsCoCoM("Cod_Comp")
                cocmTipo_Comp = "RVT"
                'cocmTipo_Comp = rscocom("Tipo_Comp")
                cocmCod_Trans = rsCoCoM("Cod_Trans")
                cocmCod_Trans_Detalle = rsCoCoM("Cod_Trans_Detalle")
                cocmOrg_Codigo = rsCoCoM("Org_Codigo")
                cocmGes_Gestion = rsCoCoM("Ges_Gestion")
                cocmNum_Respaldo = rsCoCoM("Num_Respaldo")
                cocmFecha_A = CDate(rsCoCoM("Fecha_A"))
                cocmCodigo_Beneficiario = rsCoCoM("Codigo_Beneficiario")
                cocmCodigo_Documento = rsCoCoM("Codigo_Documento")
                cocmGlosa = rsCoCoM("Glosa")
                cocmStatus = rsCoCoM("Status")
                cocmUsr_usuario = rsCoCoM("Usr_Usuario")
                'Adicionando un nuevo registro
                'Generando nuevo código
                        Set rsCorr = New ADODB.Recordset
                        If rsCorr.State = 1 Then rsCorr.Close
                        rsCorr.Open "select * from fc_correl where tipo_tramite='cmbte'", db, adOpenKeyset, adLockOptimistic
                        If rsCorr.RecordCount > 0 Then
                            cocmCod_Comp = rsCorr("numero_correlativo") + 1
                            rsCorr("numero_correlativo") = rsCorr("numero_correlativo") + 1
                            rsCorr.Update
                        End If
                        rsCorr.Close
'                        MsgBox "NUMERO DE 1era. CUENTA DAC" & cocmCod_Comp
                rsCoCoM.AddNew
                    rsCoCoM("Cod_Comp") = cocmCod_Comp
                    rsCoCoM("Tipo_Comp") = cocmTipo_Comp
                    rsCoCoM("Cod_Trans") = cocmCod_Trans
                    rsCoCoM("Cod_Trans_Detalle") = cocmCod_Trans_Detalle
                    rsCoCoM("org_codigo") = IIf(IsNull(cocmOrg_Codigo), " ", cocmOrg_Codigo)
                    rsCoCoM("Ges_Gestion") = cocmGes_Gestion
                    rsCoCoM("Num_Respaldo") = IIf(IsNull(cocmNum_Respaldo), "", cocmNum_Respaldo)
                    rsCoCoM("Fecha_A") = IIf(IsNull(cocmFecha_A), CDate(Date), CDate(cocmFecha_A))
                    rsCoCoM("Codigo_Beneficiario") = IIf(IsNull(cocmCodigo_Beneficiario), "", cocmCodigo_Beneficiario)
                    rsCoCoM("Codigo_Documento") = cocmCodigo_Documento
                    rsCoCoM("Glosa") = cocmGlosa
                    rsCoCoM("Status") = cocmStatus
                    rsCoCoM("usr_usuario") = GlUsuario
                    rsCoCoM("fecha_registro") = CDate(Format(Date, "dd/mm/yyyy"))
                    rsCoCoM("hora_registro") = Format(Time, "hh:mm:ss")
                rsCoCoM.Update
                
                Set rsdiario = New ADODB.Recordset
                If rsdiario.State = 1 Then rsdiario.Close
                rsdiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_CompDiario & "", db, adOpenKeyset, adLockOptimistic
                'rsDiario.Open "select * from co_Diario where Cod_Comp=" & cocmCod_Comp & "", db, adOpenKeyset, adLockOptimistic
                If rsdiario.RecordCount > 0 Then
                    AuxCod_Comp = cocmCod_Comp
                    'AuxTipo_Comp = rsdiario("Tipo_Comp")
                    AuxTipo_Comp = "RVT"
                    AuxCod_Comp_C = IIf(IsNull(rsdiario("Cod_Comp_C")), 0, rsdiario("Cod_Comp_C"))
                    AuxD_Cuenta = rsdiario("D_Cuenta")
                    AuxD_Nombre = IIf(IsNull(rsdiario("D_Nombre")), "", rsdiario("D_Nombre"))
                    AuxD_SubCta1 = rsdiario("D_SubCta1")
                    AuxD_SubCta2 = rsdiario("D_SubCta2")
                    AuxD_Aux1 = rsdiario("D_Aux1")
                    AuxD_Aux2 = rsdiario("D_Aux2")
                    AuxD_Aux3 = rsdiario("D_Aux3")
                    AuxD_Cta_Larga = IIf(IsNull(rsdiario("D_Cta_Larga")), "", rsdiario("D_Cta_Larga"))
                    AuxD_Des_Larga = IIf(IsNull(rsdiario("D_Des_Larga")), "", rsdiario("D_Des_Larga"))
                    AuxD_MontoBs = rsdiario("D_MontoBs")
                    AuxD_MontoDL = rsdiario("D_MontoDL")
                    AuxD_Cambio = rsdiario("D_Cambio")
                    
                    AuxH_Cuenta = rsdiario("H_Cuenta")
                    AuxH_Nombre = IIf(IsNull(rsdiario("H_Nombre")), "", rsdiario("H_Nombre"))
                    AuxH_SubCta1 = rsdiario("H_SubCta1")
                    AuxH_SubCta2 = rsdiario("H_SubCta2")
                    AuxH_Aux1 = rsdiario("H_Aux1")
                    AuxH_Aux2 = rsdiario("H_Aux2")
                    AuxH_Aux3 = rsdiario("H_Aux3")
                    AuxH_Cta_Larga = IIf(IsNull(rsdiario("H_Cta_Larga")), "", rsdiario("H_Cta_Larga"))
                    AuxH_Des_Larga = IIf(IsNull(rsdiario("H_Des_Larga")), "", rsdiario("H_Des_Larga"))
                    AuxH_MontoBs = rsdiario("H_MontoBs")
                    AuxH_MontoDL = rsdiario("H_MontoDL")
                    AuxH_Cambio = rsdiario("H_Cambio")
                    
                    AuxUsr_Usuario = rsdiario("Usr_Usuario")
                    AuxFecha_Registro = CDate(IIf(IsNull(rsdiario("Fecha_Registro")), CDate(Date), rsdiario("Fecha_Registro")))
                    AuxHora_Registro = Format(Time, "hh:mm:ss")
                   
                    'Adicionando una copia del registro
                    rsdiario.AddNew
                    rsdiario("Cod_Comp") = AuxCod_Comp 'AuxCod_Comp_C
                    rsdiario("Tipo_Comp") = AuxTipo_Comp
                    rsdiario("Cod_Comp_C") = AuxCod_Comp_C
                    
                    rsdiario("D_Cuenta") = AuxH_Cuenta
                    rsdiario("D_Nombre") = AuxH_Nombre
                    rsdiario("D_SubCta1") = AuxH_SubCta1
                    rsdiario("D_SubCta2") = AuxH_SubCta2
                    rsdiario("D_Aux1") = AuxH_Aux1
                    rsdiario("D_Aux2") = AuxH_Aux2
                    rsdiario("D_Aux3") = AuxH_Aux3
                    rsdiario("D_Cta_Larga") = AuxH_Cta_Larga
                    rsdiario("D_Des_Larga") = AuxH_Des_Larga
                    rsdiario("D_MontoBs") = AuxH_MontoBs
                    rsdiario("D_MontoDL") = AuxH_MontoDL
                    rsdiario("D_Cambio") = AuxH_Cambio
                    
                    rsdiario("H_Cuenta") = AuxD_Cuenta
                    rsdiario("H_Nombre") = AuxD_Nombre
                    rsdiario("H_SubCta1") = AuxD_SubCta1
                    rsdiario("H_SubCta2") = AuxD_SubCta2
                    rsdiario("H_Aux1") = AuxD_Aux1
                    rsdiario("H_Aux2") = AuxD_Aux2
                    rsdiario("H_Aux3") = AuxD_Aux3
                    rsdiario("H_Cta_Larga") = AuxD_Cta_Larga
                    rsdiario("H_Des_Larga") = AuxD_Des_Larga
                    rsdiario("H_MontoBs") = AuxD_MontoBs
                    rsdiario("H_MontoDL") = AuxD_MontoDL
                    rsdiario("H_Cambio") = AuxD_Cambio
                    
                    rsdiario("Usr_Usuario") = AuxUsr_Usuario
                    rsdiario("Fecha_Registro") = CDate(AuxFecha_Registro)
                    rsdiario("Hora_Registro") = AuxHora_Registro
                    rsdiario.Update
                    
                End If
          Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
    End If
       Else: MsgBox "No se contabilizó", vbCritical + vbInformation, "CONTABILIZACION"
End If

End Sub

'Public Sub AnulaTRP(codigo As Integer, org As String)
'  Dim ctacodigoDebe As String
'  Dim ctacodigoHaber As String
'  Dim montoBs As Double
'  Dim montopago As Double
'  Dim liquido As Double
'  Dim rsctabancariaDebe As ADODB.Recordset
'  Dim rsctabancariaHaber As ADODB.Recordset
'  Dim rscomprobanteM As ADODB.Recordset
'  Dim Rsstatus As ADODB.Recordset
'  Dim rsadcomprobanteM As ADODB.Recordset
'  Dim rsaddiario As ADODB.Recordset
'  Dim rspago As ADODB.Recordset
'  Dim rsPAgoDetalle As ADODB.Recordset
'  'Dim rsadpago As ADODB.Recordset
'  'Dim rsadpagodetalle As ADODB.Recordset
'  Set rsctabancariaDebe = New ADODB.Recordset
'  Set rsctabancariaHaber = New ADODB.Recordset
'  Set rspago = New ADODB.Recordset
'  Set rsPAgoDetalle = New ADODB.Recordset
'  'Set rsadpago = New ADODB.Recordset
'  'Set rsadpagodetalle = New ADODB.Recordset
'  Set rsaddiario = New ADODB.Recordset
'  Set rsadcomprobanteM = New ADODB.Recordset
'  Set rscomprobanteM = New ADODB.Recordset
'  Set Rsstatus = New ADODB.Recordset
'  If rspago.State = 1 Then rspago.Close
'  rspago.CursorLocation = adUseClient
'  rspago.Open "select * from pagos where codigo_pago=" & codigo & " and  org_codigo='" & org & "'", db, adOpenKeyset, adLockOptimistic
'  If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
'  rsPAgoDetalle.CursorLocation = adUseClient
'  rsPAgoDetalle.Open "select * from pago_detalle where codigo_pago=" & codigo & " and org_codigo='" & org & "'", db, adOpenKeyset, adLockOptimistic
'  'If rsadpago.State = 1 Then rsadpago.Close
'  'If rsadpagodetalle.State = 1 Then rsadpagodetalle.Close
'  If rscomprobanteM.State = 1 Then rscomprobanteM.Close
'  rscomprobanteM.CursorLocation = adUseClient
'  rscomprobanteM.Open " SELECT Co_Comprobante_M.Cod_Comp," & _
'                      "Co_Comprobante_M.Tipo_Comp,Co_Comprobante_M.cod_trans," & _
'                      "Co_Comprobante_M.cod_trans_detalle,Co_Comprobante_M.org_codigo," & _
'                      "Co_Comprobante_M.ges_gestion,Co_Comprobante_M.Num_Respaldo," & _
'                      "Co_Comprobante_M.Fecha_A,Co_Comprobante_M.codigo_beneficiario," & _
'                      "Co_Comprobante_M.codigo_documento,Co_Comprobante_M.Glosa, Co_Comprobante_M.status," & _
'                      "Co_Comprobante_M.Usr_Usuario,Co_Comprobante_M.codigo_solicitud," & _
'                      "Co_Comprobante_M.tipo_moneda,CO_Diario.Cod_trans_detalle AS diariodetalle," & _
'                      "CO_Diario.Cod_Comp_C, CO_Diario.D_Cuenta, CO_Diario.D_Subcta1,CO_Diario.D_SubCta2," & _
'                      "CO_Diario.D_Aux1, CO_Diario.D_Aux2, CO_Diario.D_Aux3,CO_Diario.D_Cta_Larga, CO_Diario.D_MontoBs," & _
'                      "CO_Diario.D_MontoDl, CO_Diario.D_Cambio,CO_Diario.H_Cuenta, CO_Diario.H_SubCta1," & _
'                      "CO_Diario.H_SubCta2, CO_Diario.H_Aux1, CO_Diario.H_Aux2,CO_Diario.H_Aux3, CO_Diario.H_Cta_Larga," & _
'                      "CO_Diario.H_MontoBs, CO_Diario.H_MontoDl,CO_Diario.H_Cambio " & _
'                      "FROM Co_Comprobante_M INNER JOIN CO_Diario ON Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp " & _
'                      "where co_comprobante_M.cod_comp='" & codigo & "' and org_codigo='" & org & "'and co_comprobante_m.tipo_comp='TRP'", db, adOpenKeyset, adLockReadOnly
''  If rsstatus.State = 1 Then rsstatus.Close
''  rsstatus.Open "select status from co_comprobante_m  where cod_comp=" & codigo & " and org_codigo='" & org & "' and tipo_comp='TRP'", db, adOpenKeyset, adLockOptimistic
''  If rsstatus.RecordCount <> 0 Then
''    rsstatus!Status = "L"
''     rsstatus.Update
''  End If
'  If rsaddiario.State = 1 Then rsaddiario.Close
'  rsaddiario.CursorLocation = adUseClient
'  rsaddiario.Open "select * from co_diario where tipo_comp='YUO'", db, adOpenKeyset, adLockOptimistic
'  If rsadcomprobanteM.State = 1 Then rsadcomprobanteM.Close
'  rsadcomprobanteM.CursorLocation = adUseClient
'  rsadcomprobanteM.Open "select * from co_comprobante_m where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'  'rsadpago.Open "select * from pagos where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'  'rsadpagodetalle.Open "select * from pago_detalle where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'  '***pagos
''  If rspago.RecordCount <> 0 Then
''     rspago!tipo_formulario = "ANC"
''     montopago = rspago!monto_bolivianos
''     liquido = IIf(IsNull(rspago!liquido_pagar), 0, rspago!liquido_pagar)
''     rspago!estado_anulado = "S"
''     rspago!estado_pagado = "L"
''     rspago.Update
''  End If
'    '****comprobanteM
'  If rscomprobanteM.RecordCount <> 0 Then
'    rsadcomprobanteM.AddNew
'    gencodigo
'    rsadcomprobanteM!Cod_Comp = NUMERO 'rscomprobanteM!Cod_Comp
'    rsadcomprobanteM!tipo_comp = "ANC"
'    rsadcomprobanteM!Cod_Trans = rscomprobanteM!Cod_Comp 'IIf(IsNull(rscomprobanteM!Cod_trans), rscomprobanteM!Cod_Comp, rscomprobanteM!Cod_trans)
'    rsadcomprobanteM!Cod_Trans_Detalle = IIf(IsNull(rsadcomprobanteM!Cod_Trans_Detalle), "1", rsadcomprobanteM!Cod_Trans_Detalle)
'    rsadcomprobanteM!org_codigo = IIf(IsNull(rscomprobanteM!org_codigo), "999", rscomprobanteM!org_codigo)
'    rsadcomprobanteM!ges_gestion = IIf(IsNull(rscomprobanteM!ges_gestion), Year(Date), rscomprobanteM!ges_gestion)
'    rsadcomprobanteM!num_respaldo = IIf(IsNull(rscomprobanteM!num_respaldo), "", rscomprobanteM!num_respaldo)
'    rsadcomprobanteM!fecha_A = CDate(rscomprobanteM!fecha_A)
'    rsadcomprobanteM!Codigo_beneficiario = IIf(IsNull(rscomprobanteM!Codigo_beneficiario), "", rscomprobanteM!Codigo_beneficiario)
'    rsadcomprobanteM!codigo_documento = IIf(IsNull(rscomprobanteM!codigo_documento), "", rscomprobanteM!codigo_documento)
'    rsadcomprobanteM!glosa = IIf(IsNull(rscomprobanteM!glosa), "", rscomprobanteM!glosa)
'    rsadcomprobanteM!Status = "N" 'IIf(IsNull(rscomprobanteM!Status), "S", rscomprobanteM!Status)
'    rsadcomprobanteM!usr_usuario = GlUsuario 'rsadcomprobanteM!usr_usuario
'    rsadcomprobanteM!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'    rsadcomprobanteM!hora_registro = Format(Time, "hh:mm:ss")
'    rsadcomprobanteM!tipo_moneda = IIf(IsNull(rscomprobanteM!tipo_moneda), "Bs.", rscomprobanteM!tipo_moneda)
'    rsadcomprobanteM!codigo_solicitud = IIf(IsNull(rscomprobanteM!codigo_solicitud), "", rscomprobanteM!codigo_solicitud)
'    rsaddiario.AddNew
'    rsaddiario!Cod_Comp = NUMERO 'rscomprobanteM!Cod_Comp
'    rsaddiario!tipo_comp = "ANC" 'rscomprobanteM!tipo_comp
'    rsaddiario!Cod_Trans_Detalle = IIf(IsNull(rscomprobanteM!diariodetalle), "1", rscomprobanteM!diariodetalle)
'    rsaddiario!Cod_Comp_C = IIf(IsNull(rscomprobanteM!Cod_Comp_C), 0, rscomprobanteM!Cod_Comp_C)
'    rsaddiario!d_cuenta = rscomprobanteM!h_cuenta
'    rsaddiario!d_subcta1 = rscomprobanteM!h_subcta1
'    rsaddiario!d_subcta2 = rscomprobanteM!h_subcta2
'    rsaddiario!d_Aux1 = rscomprobanteM!h_Aux1
'    rsaddiario!d_Aux2 = rscomprobanteM!h_Aux2
'    rsaddiario!d_Aux3 = rscomprobanteM!h_Aux3
'    rsaddiario!d_cta_larga = IIf(IsNull(rscomprobanteM!h_cta_larga), "", rscomprobanteM!h_cta_larga)
'    ctacodigoDebe = rscomprobanteM!d_cta_larga
'    rsaddiario!d_montobs = IIf(IsNull(rscomprobanteM!h_montoBs), 0, rscomprobanteM!d_montobs)
'    montoBs = IIf(IsNull(rscomprobanteM!h_montoBs), 0, rscomprobanteM!d_montobs)
'    rsaddiario!d_montoDl = IIf(IsNull(rscomprobanteM!h_montoDl), 0, rscomprobanteM!d_montoDl)
'    rsaddiario!d_Cambio = IIf(IsNull(rscomprobanteM!h_Cambio), 0, rscomprobanteM!d_Cambio)
'    rsaddiario!h_cuenta = rscomprobanteM!d_cuenta
'    rsaddiario!h_subcta1 = rscomprobanteM!d_subcta1
'    rsaddiario!h_subcta2 = rscomprobanteM!d_subcta2
'    rsaddiario!h_Aux1 = rscomprobanteM!d_Aux1
'    rsaddiario!h_Aux2 = rscomprobanteM!d_Aux2
'    rsaddiario!h_Aux3 = rscomprobanteM!d_Aux3
'    rsaddiario!h_cta_larga = IIf(IsNull(rscomprobanteM!d_cta_larga), "", rscomprobanteM!d_cta_larga)
'    ctacodigoHaber = rscomprobanteM!h_cta_larga
'    rsaddiario!h_montoBs = rscomprobanteM!d_montobs
'    rsaddiario!h_montoDl = rscomprobanteM!d_montoDl
'    rsaddiario!h_Cambio = rscomprobanteM!d_Cambio
'    rsaddiario!usr_usuario = GlUsuario
'    rsaddiario!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'    rsaddiario!hora_registro = Format(Time, "hh:mm:ss")
'    rsaddiario.Update
'    rsadcomprobanteM.Update
'  End If
'  Dim rsexiste As ADODB.Recordset
'  Set rsexiste = New ADODB.Recordset
'  If rsexiste.State = 1 Then rsexiste.Close
'  rsexiste.Open "select count(*) as numero from co_comprobante_m where cod_trans='" & Trim(codigo) & "' and org_codigo='999' and tipo_comp='ANC'", db, adOpenKeyset, adLockReadOnly
'  If rsexiste.RecordCount <> 0 Then
'    If rsexiste!NUMERO <> 0 Then
'      generoTRP = 1
'    Else
'      generoTRP = 0
'    End If
'  End If
  
'  '****cta del Debe
'  If rsctabancariaDebe.State = 1 Then rsctabancariaDebe.Close
'  rsctabancariaDebe.CursorLocation = adUseClient
'  rsctabancariaDebe.Open "SELECT Cta_Codigo,Cta_Anl_TRP,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & ctacodigoDebe & "'", db, adOpenKeyset, adLockOptimistic
'  If rsctabancariaDebe.RecordCount <> 0 Then
'    If montopago <> 0 Then
'      rsctabancariaDebe!cta_anl_TRP = IIf(IsNull(rsctabancariaDebe!cta_anl_TRP), 0, rsctabancariaDebe!cta_anl_TRP) + montopago
'    Else
'      rsctabancariaDebe!cta_anl_TRP = IIf(IsNull(rsctabancariaDebe!cta_anl_TRP), 0, rsctabancariaDebe!cta_anl_TRP) + montoBs
'    End If
'    rsctabancariaDebe.Update
'  End If
'  '****cta del haber
'  If rsctabancariaHaber.State = 1 Then rsctabancariaHaber.Close
'  rsctabancariaHaber.CursorLocation = adUseClient
'  rsctabancariaHaber.Open "SELECT Cta_Codigo,Cta_Anl_TRP,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & ctacodigoHaber & "'", db, adOpenKeyset, adLockOptimistic
'  If rsctabancariaHaber.RecordCount <> 0 Then
'    If montopago <> 0 Then
'      rsctabancariaHaber!cta_acum_anl = rsctabancariaHaber!cta_acum_anl + montopago
'    Else
'      rsctabancariaHaber!cta_acum_anl = rsctabancariaHaber!cta_acum_anl + montoBs
'    End If
'    rsctabancariaHaber.Update
'  End If
'End Sub

Public Sub gencodigo()
Dim rscorrelativo As New ADODB.Recordset
Set rscorrelativo = New ADODB.Recordset
If rscorrelativo.State = 1 Then rscorrelativo.Close
  rscorrelativo.Open "SELECT numero_correlativo, tipo_tramite FROM fc_correl WHERE (tipo_tramite = 'cmbte')", db, adOpenKeyset, adLockOptimistic
  rscorrelativo.MoveFirst
  numero = rscorrelativo!numero_correlativo + 1
  rscorrelativo!numero_correlativo = rscorrelativo!numero_correlativo + 1
  rscorrelativo.Update
End Sub
'Public Sub Cmd_contabiliza(P_codigo_pago As String, P_org_codigo As String, P_ges_gestion As String)
''On Error GoTo Asiento_Err
'db.BeginTrans
'MsgBox "Contabilizando..........", vbInformation + vbDefaultButton1, "CONTABILIZACION"
'Set recsetaux = New ADODB.Recordset
'recsetaux.CursorLocation = adUseClient
'If recsetaux.State = 1 Then recsetaux.Close
'recsetaux.Open " SELECT  distinct Co_Comprobante_M.Cod_Comp,Co_Comprobante_M.Tipo_Comp,cO_comprobante_M.Num_Respaldo," & _
'  " Co_Comprobante_M.codigo_beneficiario,Co_Comprobante_M.codigo_Documento,Co_Comprobante_M.Fecha_A,Co_Comprobante_M.ges_gestion," & _
'  " Co_Comprobante_M.Glosa,Co_Comprobante_M.status,Co_Comprobante_M.tipo_moneda,Co_Comprobante_M.codigo_solicitud,CO_Diario.D_Aux1," & _
'  "CO_Diario.D_Aux2, CO_Diario.D_Aux3,Co_Diario.d_Cta_Larga,Co_Diario.D_Des_Larga,Co_Comprobante_M.cod_Comp," & _
'  " CO_Diario.D_Cuenta, CO_Diario.D_Subcta1,CO_Diario.D_SubCta2, CO_Diario.D_Nombre,CO_Diario.D_MontoBs,D_Cambio,H_Cambio," & _
'  " CO_Diario.D_MontoDl,CO_Diario.H_SubCta1, CO_Diario.H_SubCta2,CO_Diario.H_Aux1, CO_Diario.H_Aux2,Co_Diario.H_Cta_Larga,Co_Diario.H_Des_Larga," & _
'  " CO_Diario.H_Aux3,CO_Diario.H_Nombre, CO_Diario.H_MontoBs,CO_Diario.H_Montodl,CO_Diario.H_Cuenta " & _
'  " From CO_Diario,CO_Comprobante_M WHERE CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp AND co_Comprobante_M.Cod_Comp=" & Trim(P_codigo_pago) & _
'  " and co_Comprobante_M.Tipo_Comp='PCE' and CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp and status='S' ", db, adOpenDynamic, adLockOptimistic, adCmdText
''If recSetAux.RecordCount > 0 Then
' 'MsgBox recSetAux!Cod_Comp
'  Set recSetAuxActualizar1 = New ADODB.Recordset
'  recSetAuxActualizar1.CursorLocation = adUseClient
'  If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'  recSetAuxActualizar1.Open " select distinct fc_Cuenta_Bancaria.fte_codigo,tipo_comp,Pagos.ORg_Codigo, " & _
'    " Pago_Detalle.ges_Gestion,pago_Detalle.cta_Codigo,fc_Cuenta_Bancaria.cta_codigo_tgn,fc_Cuenta_Bancaria.cta_descripcion_larga," & _
'    " Pago_Detalle.fecha_Pago from Pagos,Pago_Detalle,fc_Cuenta_Bancaria where " & _
'    " Pagos.Ges_Gestion = Pago_Detalle.Ges_Gestion and Pagos.Org_Codigo=Pago_Detalle.Org_Codigo and  Pagos.Codigo_Pago=Pago_Detalle.Codigo_Pago " & _
'    " and Pagos.Tipo_Comp= 'PCE' and Pagos.Codigo_Pago = '" & P_codigo_pago & "' and  Pagos.Org_Codigo='999' and " & _
'    " fc_Cuenta_Bancaria.Cta_Codigo=Pago_Detalle.Cta_Codigo ", db, adOpenDynamic, adLockOptimistic, adCmdText
'  If recSetAuxActualizar1.RecordCount > 0 Then recSetAuxActualizar1.MoveFirst
'  While Not (recSetAuxActualizar1.EOF)
'      v_Fte = recSetAuxActualizar1!fte_codigo
'      If recsetAdicion.State = 1 Then recsetAdicion.Close
'      recsetAdicion.Open " select * from Co_Comprobante_M  where cod_Trans=" & P_codigo_pago & " and Org_Codigo='999' and Ges_Gestion='" & P_ges_gestion & "'", db, adOpenDynamic, adLockOptimistic
'      If Not recsetAdicion.BOF Then recsetAdicion.MoveFirst
'      If (recsetAdicion.BOF) And (recsetAdicion.EOF) Then
'
'    '************* GENERA EL CODIGO DE COMPROBANTE**********
'            Set recSetGenera = New ADODB.Recordset
'            recSetGenera.CursorLocation = adUseClient
'            If recSetGenera.State = 1 Then recSetGenera.Close
'            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'            If recSetGenera.RecordCount > 0 Then
'                Cont_Comp = Val(recSetGenera!numero_correlativo)
'                Cont_Comp = Cont_Comp + 1
'                recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
'                recSetGenera.Update
'            End If
'            If recSetGenera.State = 1 Then recSetGenera.Close
''************TERMINA GENERACION DE COMPROBANTE********
'            recsetAdicion.AddNew
'            recsetAdicion!usr_usuario = GlUsuario
'            recsetAdicion!fecha_registro = Date
'            recsetAdicion!Hora_registro = Format(Time, "hh:mm:ss")
'            recsetAdicion!Cod_Comp = Cont_Comp
'            recsetAdicion!Cod_Trans = recsetaux!Cod_Comp
'            recsetAdicion!cod_trans_detalle = "1"
'            recsetAdicion!org_codigo = P_org_codigo
'            recsetAdicion!tipo_comp = "PCC" 'recsetaux!tipo_comp
'            recsetAdicion!ges_gestion = recSetAuxActualizar1!ges_gestion
'            recsetAdicion!fecha_A = CDate(recSetAuxActualizar1!fecha_pago)
'            Select Case recsetaux!tipo_comp
'              Case "PCE"
'                recsetAdicion!Codigo_beneficiario = recsetaux!Codigo_beneficiario
'              Case "PCO"
'            End Select
'            recsetAdicion!glosa = recsetaux!glosa
'            recsetAdicion!codigo_documento = recsetaux!codigo_documento
'            recsetAdicion!num_respaldo = recsetaux!num_respaldo
'            recsetAdicion!Status = recsetaux!Status
'            recsetAdicion!tipo_moneda = IIf(IsNull(recsetaux!tipo_moneda), "Bs", recsetaux!tipo_moneda)
'            recsetAdicion!codigo_solicitud = IIf(IsNull(recsetaux!codigo_solicitud), "", recsetaux!codigo_solicitud)
'            recsetAdicion!usr_usuario = GlUsuario
'            recsetAdicion!fecha_registro = Format(Date, "dd/mm/yyyy")
'            recsetAdicion!Hora_registro = Format(Time, "hh:mm:ss")
'            recsetAdicion.Update
'            If recsetAdicion.State = 1 Then recsetAdicion.Close
'        '********* adicion Debitos creditos
'            Set recSetAuxActualizar = New ADODB.Recordset
'            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'            recSetAuxActualizar.Open " select * from Co_Diario where  cod_Comp_c=" & recsetaux!Cod_Comp, db, adOpenDynamic, adLockOptimistic, adCmdText
'            If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
'              recSetAuxActualizar.AddNew
'              recSetAuxActualizar!usr_usuario = GlUsuario
'              recSetAuxActualizar!fecha_registro = Format(Date, "dd/mm/yyyy")
'              recSetAuxActualizar!Hora_registro = Format(Time, "hh:mm:ss")
'              'recsetAdicion!Cod_Comp = Cont_Comp
'              recSetAuxActualizar!Cod_Comp = Cont_Comp
'              recSetAuxActualizar!tipo_comp = "PCC" 'recsetaux!tipo_comp
'              recSetAuxActualizar!Cod_Comp_C = recsetaux!Cod_Comp
'              recSetAuxActualizar!d_cuenta = recsetaux!h_cuenta
'              recSetAuxActualizar!d_subcta1 = recsetaux!h_subcta1
'              recSetAuxActualizar!d_subcta2 = recsetaux!h_subcta2
'              recSetAuxActualizar!d_Aux1 = recsetaux!h_Aux1
'              recSetAuxActualizar!d_Aux2 = recsetaux!h_Aux2
'              recSetAuxActualizar!d_Aux3 = recsetaux!h_Aux3
'              recSetAuxActualizar!d_cta_larga = recsetaux!h_cta_larga
'              recSetAuxActualizar!D_Des_Larga = IIf(IsNull(recsetaux!H_Des_Larga), " ", Trim(recsetaux!H_Des_Larga))
'              recSetAuxActualizar!d_montobs = recsetaux!H_MontoBs
'              recSetAuxActualizar!d_montoDl = recsetaux!h_montoDl
'              recSetAuxActualizar!d_Cambio = recsetaux!h_Cambio
'              Select Case v_Fte
'                Case "10", "41"
'                  recSetAuxActualizar!h_cuenta = "1111"
'                  recSetAuxActualizar!h_subcta1 = "02"
'                  recSetAuxActualizar!h_subcta2 = "01"
'                  recSetAuxActualizar!h_Aux1 = "02"
'                  recSetAuxActualizar!h_Aux2 = "00"
'                  recSetAuxActualizar!h_Aux3 = "00"
'                Case "70", "43"
'                  recSetAuxActualizar!h_cuenta = "1111"
'                  recSetAuxActualizar!h_subcta1 = "02"
'                  recSetAuxActualizar!h_subcta2 = "02"
'                  recSetAuxActualizar!h_Aux1 = "02"
'                  recSetAuxActualizar!h_Aux2 = "00"
'                  recSetAuxActualizar!h_Aux3 = "00"
'                Case "80"
'                  recSetAuxActualizar!h_cuenta = "1111"
'                  recSetAuxActualizar!h_subcta1 = "02"
'                  recSetAuxActualizar!h_subcta2 = "03"
'                  recSetAuxActualizar!h_Aux1 = "02"
'                  recSetAuxActualizar!h_Aux2 = "00"
'                  recSetAuxActualizar!h_Aux3 = "00"
'              End Select
'              recSetAuxActualizar!h_cta_larga = recSetAuxActualizar1!cta_codigo
'              recSetAuxActualizar!H_Des_Larga = IIf(IsNull(recSetAuxActualizar1!cta_descripcion_larga), " ", recSetAuxActualizar1!cta_descripcion_larga)
'              recSetAuxActualizar!H_MontoBs = recsetaux!H_MontoBs
'              recSetAuxActualizar!h_montoDl = recsetaux!h_montoDl
'              recSetAuxActualizar!h_Cambio = recsetaux!h_Cambio
'              recSetAuxActualizar!usr_usuario = GlUsuario
'              recSetAuxActualizar!fecha_registro = Format(Date, "dd/mm/yyyy")
'              recSetAuxActualizar!Hora_registro = Format(Time, "hh:mm:ss")
'              recSetAuxActualizar.Update
'              If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
''''************TERMINA GENERACION DE COMPROBANTE********
'         End If 'Adicion del diario
'      Else
'           MsgBox "Ya fue contabilizado anteriormente", vbInformation + vbDefaultButton1, "CONTABILIZACION"
'' ******Modifica registro existente
'            recsetAdicion!fecha_registro = Date
'            recsetAdicion!Hora_registro = Format(Time, "hh:mm:ss")
'            Cont_Comp = recsetAdicion!Cod_Comp
'            recsetAdicion!Cod_Comp = Cont_Comp
'            recsetAdicion!Cod_Trans = recsetaux!Cod_Comp
'            recsetAdicion!cod_trans_detalle = "1"
'            recsetAdicion!org_codigo = recSetAuxActualizar1!org_codigo
'            recsetAdicion!tipo_comp = "PCC" 'recsetaux!tipo_comp
'            recsetAdicion!ges_gestion = recSetAuxActualizar1!ges_gestion
'            recsetAdicion!fecha_A = CDate(recSetAuxActualizar1!fecha_pago)
'            Select Case recsetaux!tipo_comp
'                Case "PCE", "PCC"
'                   recsetAdicion!Codigo_beneficiario = recsetaux!Codigo_beneficiario
'                Case "PCO"
'
'            End Select
'            recsetAdicion!glosa = recsetaux!glosa
'            recsetAdicion!codigo_documento = recsetaux!codigo_documento
'            recsetAdicion!num_respaldo = recsetaux!num_respaldo
'            recsetAdicion!Status = recsetaux!Status
'            recsetAdicion!usr_usuario = GlUsuario
'            recsetAdicion!fecha_registro = Format(Date, "dd/mm/yyyy")
'            recsetAdicion!Hora_registro = Format(Time, "hh:mm:ss")
'            recsetAdicion!tipo_moneda = IIf(IsNull(recsetaux!tipo_moneda), "Bs", recsetaux!tipo_moneda)
'            recsetAdicion!codigo_solicitud = IIf(IsNull(recsetaux!codigo_solicitud), "", recsetaux!codigo_solicitud)
'            recsetAdicion.Update
'            If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'    '******Termina de Modificar
'
'    '******Modifica el Diario
'            Set recSetAuxActualizar = New ADODB.Recordset
'            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'            recSetAuxActualizar.Open " select * from Co_Diario where  cod_Comp=" & Cont_Comp, db, adOpenDynamic, adLockOptimistic
'            If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
'              recSetAuxActualizar.AddNew
'              recSetAuxActualizar!Cod_Comp = Cont_Comp
'              recSetAuxActualizar!tipo_comp = recsetaux!tipo_comp
'            Else
'              If (Not recSetAuxActualizar.BOF) Then recSetAuxActualizar.MoveFirst
'            End If
'
'            recSetAuxActualizar!usr_usuario = GlUsuario
'            recSetAuxActualizar!fecha_registro = Format(Date, "dd/mm/yyyy")
'            recSetAuxActualizar!Hora_registro = Format(Time, "hh:mm:ss")
'            'recsetAdicion!Cod_Comp = Cont_Comp
'            'recSetAuxActualizar!Cod_Comp = Cont_Comp
'            'recSetAuxActualizar!Tipo_comp = recSetAux!Tipo_comp
'            recSetAuxActualizar!Cod_Comp_C = recsetaux!Cod_Comp
'            recSetAuxActualizar!d_cuenta = recsetaux!h_cuenta
'            recSetAuxActualizar!d_subcta1 = recsetaux!h_subcta1
'            recSetAuxActualizar!d_subcta2 = recsetaux!h_subcta2
'
'            recSetAuxActualizar!d_Aux1 = recsetaux!h_Aux1
'            recSetAuxActualizar!d_Aux2 = recsetaux!h_Aux2
'            recSetAuxActualizar!d_Aux3 = recsetaux!h_Aux3
'
'            recSetAuxActualizar!d_cta_larga = recsetaux!h_cta_larga
'            recSetAuxActualizar!D_Des_Larga = IIf(IsNull(recsetaux!H_Des_Larga), " ", recsetaux!H_Des_Larga)
'            recSetAuxActualizar!d_montobs = recsetaux!H_MontoBs
'            recSetAuxActualizar!d_montoDl = recsetaux!h_montoDl
'            recSetAuxActualizar!d_Cambio = recsetaux!h_Cambio
'
'            Select Case v_Fte
'
'               Case "10", "41"
'                  recSetAuxActualizar!h_cuenta = "1111"
'                  recSetAuxActualizar!h_subcta1 = "02"
'                  recSetAuxActualizar!h_subcta2 = "01"
'                  recSetAuxActualizar!h_Aux1 = "02"
'                  recSetAuxActualizar!h_Aux2 = "00"
'                  recSetAuxActualizar!h_Aux3 = "00"
'
'               Case "70", "43"
'                  recSetAuxActualizar!h_cuenta = "1111"
'                  recSetAuxActualizar!h_subcta1 = "02"
'                  recSetAuxActualizar!h_subcta2 = "02"
'                  recSetAuxActualizar!h_Aux1 = "02"
'                  recSetAuxActualizar!h_Aux2 = "00"
'                  recSetAuxActualizar!h_Aux3 = "00"
'
'              Case "80"
'                  recSetAuxActualizar!h_cuenta = "1111"
'                  recSetAuxActualizar!h_subcta1 = "02"
'                  recSetAuxActualizar!h_subcta2 = "03"
'                  recSetAuxActualizar!h_Aux1 = "02"
'                  recSetAuxActualizar!h_Aux2 = "00"
'                  recSetAuxActualizar!h_Aux3 = "00"
'             End Select
'
'                recSetAuxActualizar!h_cta_larga = recSetAuxActualizar1!cta_codigo
'                recSetAuxActualizar!H_Des_Larga = IIf(IsNull(recSetAuxActualizar1!cta_descripcion_larga), "", recSetAuxActualizar1!cta_descripcion_larga)
'                recSetAuxActualizar!H_MontoBs = recsetaux!H_MontoBs
'                recSetAuxActualizar!h_montoDl = recsetaux!h_montoDl
'                recSetAuxActualizar!h_Cambio = recsetaux!h_Cambio
'                recSetAuxActualizar.Update
'             If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'
'       End If '*****Existe comprobante modificaion
'
''******Termina de Modificar el diario
'
''         Else
''         MsgBox "No existen cuentas asociadas................"
''         End If
'    recSetAuxActualizar1.MoveNext
'  Wend
'  db.CommitTrans
'  MsgBox "Contabilizacion exitosa...............", vbInformation + vbDefaultButton1, "CONTABILIZACION"
'Exit Sub
'Asiento_Err:
'    MsgBox "Error al generar contra cuenta"
'    db.RollbackTrans
'    'CmdAgregarDetalle.Enabled = True
'    'Cmd_Modificar.Enabled = True
'    'Cmd_Aprobar.Enabled = True
'    'CmdSalir.Enabled = True
'    'Cmd_GrabaM.Enabled = True
'    'Cmd_Cancelar.Enabled = True
'    'Cmd_Copiar.Enabled = True
'
'End Sub

'Public Sub Reversion999(codigo As Integer, org As String)
'  Dim ctacodigo As String
'  Dim montoBs As Double
'  Dim montopago As Double
'  Dim liquido As Double
'  Dim rsctabancaria As ADODB.Recordset
'  Dim rscomprobanteM As ADODB.Recordset
'  Dim rsadcomprobanteM As ADODB.Recordset
'  Dim rsaddiario As ADODB.Recordset
'  'Dim rspago As ADODB.Recordset
'  'Dim rspagodetalle As ADODB.Recordset
'  Set rsctabancaria = New ADODB.Recordset
'  'Set rspago = New ADODB.Recordset
'  'Set rspagodetalle = New ADODB.Recordset
'  Set rsaddiario = New ADODB.Recordset
'  Set rsadcomprobanteM = New ADODB.Recordset
'  Set rscomprobanteM = New ADODB.Recordset
'  'If rspago.State = 1 Then rspago.Close
'  'rspago.CursorLocation = adUseClient
'  'rspago.Open "select * from pagos where codigo_pago=" & codigo & " and  org_codigo='" & org & "'", db, adOpenKeyset, adLockOptimistic
'  'If rspagodetalle.State = 1 Then rspagodetalle.Close
'  'rspagodetalle.CursorLocation = adUseClient
'  'rspagodetalle.Open "select * from pago_detalle where codigo_pago=" & codigo & " and org_codigo='" & org & "'", db, adOpenKeyset, adLockOptimistic
'
'  If rscomprobanteM.State = 1 Then rscomprobanteM.Close
'  rscomprobanteM.CursorLocation = adUseClient
'  rscomprobanteM.Open " SELECT Co_Comprobante_M.Cod_Comp," & _
'                      "Co_Comprobante_M.Tipo_Comp,Co_Comprobante_M.cod_trans," & _
'                      "Co_Comprobante_M.cod_trans_detalle,Co_Comprobante_M.org_codigo," & _
'                      "Co_Comprobante_M.ges_gestion,Co_Comprobante_M.Num_Respaldo," & _
'                      "Co_Comprobante_M.Fecha_A,Co_Comprobante_M.codigo_beneficiario," & _
'                      "Co_Comprobante_M.codigo_documento,Co_Comprobante_M.Glosa, Co_Comprobante_M.status," & _
'                      "Co_Comprobante_M.Usr_Usuario,Co_Comprobante_M.codigo_solicitud," & _
'                      "Co_Comprobante_M.tipo_moneda,CO_Diario.Cod_trans_detalle AS diariodetalle," & _
'                      "CO_Diario.Cod_Comp_C, CO_Diario.D_Cuenta, CO_Diario.D_Subcta1,CO_Diario.D_SubCta2," & _
'                      "CO_Diario.D_Aux1, CO_Diario.D_Aux2, CO_Diario.D_Aux3,CO_Diario.D_Cta_Larga, CO_Diario.D_MontoBs," & _
'                      "CO_Diario.D_MontoDl, CO_Diario.D_Cambio,CO_Diario.H_Cuenta, CO_Diario.H_SubCta1," & _
'                      "CO_Diario.H_SubCta2, CO_Diario.H_Aux1, CO_Diario.H_Aux2,CO_Diario.H_Aux3, CO_Diario.H_Cta_Larga," & _
'                      "CO_Diario.H_MontoBs, CO_Diario.H_MontoDl,CO_Diario.H_Cambio " & _
'                      "FROM Co_Comprobante_M INNER JOIN CO_Diario ON Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp " & _
'                      "where co_comprobante_M.cod_comp='" & codigo & "' and org_codigo='" & org & _
'                      "' and co_comprobante_m.tipo_comp='PCE'", db, adOpenKeyset, adLockReadOnly
'  If rsaddiario.State = 1 Then rsaddiario.Close
'  rsaddiario.CursorLocation = adUseClient
'  rsaddiario.Open "select * from co_diario where tipo_comp='YUO'", db, adOpenKeyset, adLockOptimistic
'  If rsadcomprobanteM.State = 1 Then rsadcomprobanteM.Close
'  rsadcomprobanteM.CursorLocation = adUseClient
'  rsadcomprobanteM.Open "select * from co_comprobante_m where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'  '***pagos
'  'If rspago.RecordCount <> 0 Then
'   ' rspago!nro_comprobante_anterior = rspago!codigo_pago
'   ' rspago!tipo_formulario = "RVP"
'   ' montopago = rspago!monto_bolivianos
'   ' liquido = IIf(IsNull(rspago!liquido_pagar), 0, rspago!liquido_pagar)
'    'rspago!estado_pagado = "L"
'   ' rspago!estado_contabilidad = "R"
'   ' rspago!estado_aprobacion = "N"
'   ' rspago!Usr_Usuario = GlUsuario
'   ' rspago!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'   ' rspago!Hora_Registro = Format(Time, "hh:mm:ss")
'    'rspago!justificacion = IIf(IsNull(rspago!justificacion), "", Trim(rspago!justificacion))
'   ' rspago.Update
'  'End If
'  '****pagodetalle
'  '****comprobanteM
'  If rscomprobanteM.RecordCount <> 0 Then
'    rsadcomprobanteM.AddNew
'    Call gencodigo
'    rsadcomprobanteM!Cod_Comp = NUMERO 'rscomprobanteM!Cod_Comp
'    rsadcomprobanteM!tipo_comp = "RVT"
'    rsadcomprobanteM!Cod_Trans = codigo 'IIf(IsNull(rscomprobanteM!Cod_Trans), "", rscomprobanteM!Cod_Trans)
'    rsadcomprobanteM!Cod_Trans_Detalle = IIf(IsNull(rsadcomprobanteM!Cod_Trans_Detalle), "1", rsadcomprobanteM!Cod_Trans_Detalle)
'    rsadcomprobanteM!org_codigo = IIf(IsNull(rscomprobanteM!org_codigo), "999", rscomprobanteM!org_codigo)
'    rsadcomprobanteM!ges_gestion = IIf(IsNull(rscomprobanteM!ges_gestion), Year(Date), rscomprobanteM!ges_gestion)
'    rsadcomprobanteM!Num_respaldo = IIf(IsNull(rscomprobanteM!Num_respaldo), "", rscomprobanteM!Num_respaldo)
'    rsadcomprobanteM!fecha_A = CDate(Format(Date, "dd/mm/yyyy")) 'CDate(rscomprobanteM!fecha_A)
'    rsadcomprobanteM!Codigo_beneficiario = IIf(IsNull(rscomprobanteM!Codigo_beneficiario), "", rscomprobanteM!Codigo_beneficiario)
'    rsadcomprobanteM!Codigo_documento = IIf(IsNull(rscomprobanteM!Codigo_documento), "", rscomprobanteM!Codigo_documento)
'    rsadcomprobanteM!glosa = IIf(IsNull(rscomprobanteM!glosa), "", rscomprobanteM!glosa)
'    rsadcomprobanteM!Status = "N" 'IIf(IsNull(rscomprobanteM!Status), "S", rscomprobanteM!Status)
'    rsadcomprobanteM!Usr_Usuario = GlUsuario 'rsadcomprobanteM!usr_usuario
'    rsadcomprobanteM!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'    rsadcomprobanteM!Hora_Registro = Format(Time, "hh:mm:ss")
'    rsadcomprobanteM!tipo_moneda = IIf(IsNull(rscomprobanteM!tipo_moneda), "Bs.", rscomprobanteM!tipo_moneda)
'    rsadcomprobanteM!codigo_solicitud = IIf(IsNull(rscomprobanteM!codigo_solicitud), "", rscomprobanteM!codigo_solicitud)
'    rsaddiario.AddNew
'    rsaddiario!Cod_Comp = NUMERO 'rscomprobanteM!Cod_Comp
'    rsaddiario!tipo_comp = "RVT" 'rscomprobanteM!tipo_comp
'    rsaddiario!Cod_Trans_Detalle = IIf(IsNull(rscomprobanteM!diariodetalle), "1", rscomprobanteM!diariodetalle)
'    rsaddiario!Cod_Comp_C = codigo 'IIf(IsNull(rscomprobanteM!Cod_Comp_C), 0, rscomprobanteM!Cod_Comp_C)
'    rsaddiario!d_cuenta = rscomprobanteM!h_cuenta
'    rsaddiario!d_subcta1 = rscomprobanteM!h_subcta1
'    rsaddiario!d_subcta2 = rscomprobanteM!h_subcta2
'    rsaddiario!d_Aux1 = rscomprobanteM!h_Aux1
'    rsaddiario!d_Aux2 = rscomprobanteM!h_Aux2
'    rsaddiario!d_Aux3 = rscomprobanteM!h_Aux3
'    rsaddiario!d_cta_Larga = IIf(IsNull(rscomprobanteM!h_cta_Larga), "", rscomprobanteM!h_cta_Larga)
'    rsaddiario!d_montobs = IIf(IsNull(rscomprobanteM!h_montoBs), 0, rscomprobanteM!d_montobs)
'    rsaddiario!d_montoDl = IIf(IsNull(rscomprobanteM!h_montoDl), 0, rscomprobanteM!d_montoDl)
'    rsaddiario!d_Cambio = IIf(IsNull(rscomprobanteM!h_Cambio), 0, rscomprobanteM!d_Cambio)
'    rsaddiario!h_cuenta = rscomprobanteM!d_cuenta
'    rsaddiario!h_subcta1 = rscomprobanteM!d_subcta1
'    rsaddiario!h_subcta2 = rscomprobanteM!d_subcta2
'    rsaddiario!h_Aux1 = rscomprobanteM!d_Aux1
'    rsaddiario!h_Aux2 = rscomprobanteM!d_Aux2
'    rsaddiario!h_Aux3 = rscomprobanteM!d_Aux3
'    rsaddiario!h_cta_Larga = IIf(IsNull(rscomprobanteM!d_cta_Larga), "", rscomprobanteM!d_cta_Larga)
'    rsaddiario!h_montoBs = rscomprobanteM!d_montobs
'    rsaddiario!h_montoDl = rscomprobanteM!d_montoDl
'    rsaddiario!h_Cambio = rscomprobanteM!d_Cambio
'    rsaddiario!Usr_Usuario = GlUsuario
'    rsaddiario!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'    rsaddiario!Hora_Registro = Format(Time, "hh:mm:ss")
'    rsaddiario.Update
'    rsadcomprobanteM.Update
'  End If
'  Dim rsrever As ADODB.Recordset
'  Set rsrever = New ADODB.Recordset
'  If rsrever.State = 1 Then rsanul.Close
'  rsrever.CursorLocation = adUseClient
'  rsrever.Open "select count(*) as num from co_comprobante_m where cod_trans='" & codigo & " ' and org_codigo='" & org & "' and tipo_comp='RVT'", db, adOpenKeyset, adLockReadOnly
'  If rsrever.RecordCount <> 0 Then
'    If rsrever!num <> 0 Then
'       rever999 = 1
'    Else
'       rever999 = 0
'    End If
'  End If
'End Sub
'Public Sub Anulacion999(codigo As Integer, org As String)
'  Dim ctacodigo As String
'  Dim montoBs As Double
'  Dim montopago As Double
'  Dim liquido As Double
'  Dim rsctabancaria As ADODB.Recordset
'  Dim rscomprobanteM As ADODB.Recordset
'  Dim rsadcomprobanteM As ADODB.Recordset
'  Dim rsaddiario As ADODB.Recordset
'  Dim rspago As ADODB.Recordset
'  Dim rsPAgoDetalle As ADODB.Recordset
' 'Dim rsadpago As ADODB.Recordset
' ' Dim rsadpagodetalle As ADODB.Recordset
'  Set rsctabancaria = New ADODB.Recordset
'  Set rspago = New ADODB.Recordset
'  Set rsPAgoDetalle = New ADODB.Recordset
'  'Set rsadpago = New ADODB.Recordset
'  'Set rsadpagodetalle = New ADODB.Recordset
'  Set rsaddiario = New ADODB.Recordset
'  Set rsadcomprobanteM = New ADODB.Recordset
'  Set rscomprobanteM = New ADODB.Recordset
'  If rspago.State = 1 Then rspago.Close
'  rspago.CursorLocation = adUseClient
'  rspago.Open "select * from pagos where codigo_pago=" & codigo & " and  org_codigo='" & org & "'", db, adOpenKeyset, adLockOptimistic
'  If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
'  rsPAgoDetalle.CursorLocation = adUseClient
'  rsPAgoDetalle.Open "select * from pago_detalle where codigo_pago=" & codigo & " and org_codigo='" & org & "'", db, adOpenKeyset, adLockOptimistic
'  'If rsadpago.State = 1 Then rsadpago.Close
'  'If rsadpagodetalle.State = 1 Then rsadpagodetalle.Close
'  If rscomprobanteM.State = 1 Then rscomprobanteM.Close
'  rscomprobanteM.CursorLocation = adUseClient
'  rscomprobanteM.Open " SELECT Co_Comprobante_M.Cod_Comp," & _
'                      "Co_Comprobante_M.Tipo_Comp,Co_Comprobante_M.cod_trans," & _
'                      "Co_Comprobante_M.cod_trans_detalle,Co_Comprobante_M.org_codigo," & _
'                      "Co_Comprobante_M.ges_gestion,Co_Comprobante_M.Num_Respaldo," & _
'                      "Co_Comprobante_M.Fecha_A,Co_Comprobante_M.codigo_beneficiario," & _
'                      "Co_Comprobante_M.codigo_documento,Co_Comprobante_M.Glosa, Co_Comprobante_M.status," & _
'                      "Co_Comprobante_M.Usr_Usuario,Co_Comprobante_M.codigo_solicitud," & _
'                      "Co_Comprobante_M.tipo_moneda,CO_Diario.Cod_trans_detalle AS diariodetalle," & _
'                      "CO_Diario.Cod_Comp_C, CO_Diario.D_Cuenta, CO_Diario.D_Subcta1,CO_Diario.D_SubCta2," & _
'                      "CO_Diario.D_Aux1, CO_Diario.D_Aux2, CO_Diario.D_Aux3,CO_Diario.D_Cta_Larga, CO_Diario.D_MontoBs," & _
'                      "CO_Diario.D_MontoDl, CO_Diario.D_Cambio,CO_Diario.H_Cuenta, CO_Diario.H_SubCta1," & _
'                      "CO_Diario.H_SubCta2, CO_Diario.H_Aux1, CO_Diario.H_Aux2,CO_Diario.H_Aux3, CO_Diario.H_Cta_Larga," & _
'                      "CO_Diario.H_MontoBs, CO_Diario.H_MontoDl,CO_Diario.H_Cambio " & _
'                      "FROM Co_Comprobante_M INNER JOIN CO_Diario ON Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp " & _
'                      "where co_comprobante_M.cod_trans='" & codigo & "' and org_codigo='" & org & "'" & _
'                      " and (co_comprobante_m.tipo_comp='PCC' or co_comprobante_m.tipo_comp='PCE')", db, adOpenKeyset, adLockReadOnly
'
'  If rsaddiario.State = 1 Then rsaddiario.Close
'  rsaddiario.CursorLocation = adUseClient
'  rsaddiario.Open "select * from co_diario where tipo_comp='YUO'", db, adOpenKeyset, adLockOptimistic
'  If rsadcomprobanteM.State = 1 Then rsadcomprobanteM.Close
'  rsadcomprobanteM.CursorLocation = adUseClient
'  rsadcomprobanteM.Open "select * from co_comprobante_m where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'  'rsadpago.Open "select * from pagos where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'  'rsadpagodetalle.Open "select * from pago_detalle where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'  '***pagos
'  If rspago.RecordCount <> 0 Then
'    'rsadpago.AddNew
'    'Call gencodigo
'    'rspago!ges_gestion = rspago!ges_gestion
'    'rspago!Org_Codigo = rspago!Org_Codigo
'    'rspago!codigo_pago = NUMERO 'rspago!codigo_pago
'    'rspago!tipo_comp = rspago!tipo_comp
'  ''  rspago!Nro_Comprobante_Anterior = rspago!codigo_pago
'    'rspago!codigo_solicitud = IIf(IsNull(rspago!codigo_solicitud), "", rspago!codigo_solicitud)
'  ''  rspago!tipo_formulario = "ANP"
'    'rspago!codigo_orden = IIf(IsNull(rspago!codigo_orden), "", rspago!codigo_orden)
'    'rspago!codigo_documento = IIf(IsNull(rspago!codigo_documento), "", rspago!codigo_documento)
'    'If IsNull(rspago!fecha_egreso) Then
'    ' FECHA = CDate(Date)
'    'Else
'    '  FECHA = CDate(rspago!fecha_egreso)
'    'End If
'    'rsadpago!fecha_egreso = IIf(IsNull(rspago!fecha_egreso), CDate(Format(Date, "dd/mm/yyyy")), CDate(rspago!fecha_egreso))
'    'rsadpago!fecha_egreso = CDate(FECHA)
'    'rsadpago!tipo_moneda = IIf(IsNull(rspago!tipo_moneda), "Bs.", rspago!tipo_moneda)
'    'rsadpago!monto_bolivianos = IIf(IsNull(rspago!monto_bolivianos), 0, rspago!monto_bolivianos)
'    'rsadpago!monto_Dolares = IIf(IsNull(rspago!monto_Dolares), 0, rspago!monto_Dolares)
'   '' montopago = rspago!monto_bolivianos
'    'rsadpago!liquido_pagar = IIf(IsNull(rspago!liquido_pagar), 0, rspago!liquido_pagar)
'    ''liquido = IIf(IsNull(rspago!liquido_pagar), 0, rspago!liquido_pagar)
'   '' rspago!estado_pagado = "L"
'    'rspago!estado_aprobacion = "N"
'    ''rspago!usr_usuario = GlUsuario
'    ''rspago!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'   '' rspago!hora_registro = Format(Time, "hh:mm:ss")
'    'rspago!justificacion = IIf(IsNull(rspago!justificacion), "", Trim(rspago!justificacion))
'    ''rspago.Update
'  End If
'  '****pagodetalle
'  If rsPAgoDetalle.RecordCount <> 0 Then
'    'rsadpago.Open "select * from ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
''    rsadpagodetalle.Open "select * from pago_detalle where ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
'    'rsadpagodetalle.AddNew
'    'rsadpagodetalle!ges_gestion = rsPAgoDetalle!ges_gestion
'    'rsadpagodetalle!Org_Codigo = "999"
'    'rsadpagodetalle!codigo_pago = NUMERO 'rsadpagodetalle!codigo_pago
'    'rsadpagodetalle!codigo_pago_detalle = IIf(IsNull(rsPAgoDetalle!codigo_pago_detalle), "1", rsPAgoDetalle!codigo_pago_detalle)
'    ''ctacodigo = IIf(IsNull(rsPAgoDetalle!cta_codigo), "", rsPAgoDetalle!cta_codigo)
'    'rsadpagodetalle!cta_codigo = IIf(IsNull(rsPAgoDetalle!cta_codigo), "", rsPAgoDetalle!cta_codigo)
'    'rsadpagodetalle!cheque_o_trf = IIf(IsNull(rsPAgoDetalle!cheque_o_trf), "", rsPAgoDetalle!cheque_o_trf)
'    'rsadpagodetalle!numero_cheque_trf = IIf(IsNull(rsPAgoDetalle!numero_cheque_trf), "", rsPAgoDetalle!numero_cheque_trf)
'    'rsadpagodetalle!Codigo_beneficiario = IIf(IsNull(rsPAgoDetalle!Codigo_beneficiario), "", rsPAgoDetalle!Codigo_beneficiario)
'    'rsadpagodetalle!monto_bolivianos = IIf(IsNull(rsPAgoDetalle!monto_bolivianos), 0, rsPAgoDetalle!monto_bolivianos)
'    'rsadpagodetalle!monto_Dolares = IIf(IsNull(rsPAgoDetalle!monto_Dolares), 0, rsPAgoDetalle!monto_Dolares)
'    'rsadpagodetalle!monto_total = IIf(IsNull(rsPAgoDetalle!monto_total), 0, rsPAgoDetalle!monto_total)
'   '' If rsPAgoDetalle!monto_total <> 0 Then
'    ''  montoBs = IIf(IsNull(rsPAgoDetalle!monto_total), 0, rsPAgoDetalle!monto_total)
'    ''Else
'     '' montoBs = IIf(IsNull(rsPAgoDetalle!monto_bolivianos), 0, rsPAgoDetalle!monto_bolivianos)
'    ''End If
'    'rsadpagodetalle!tipo_cambio = IIf(IsNull(rsPAgoDetalle!tipo_cambio), 0, rsPAgoDetalle!tipo_cambio)
'    'rsadpagodetalle!fecha_pago = IIf(IsNull(rsPAgoDetalle!fecha_pago), CDate(Format(Date, "dd/mm/yyyy")), CDate(rsPAgoDetalle!fecha_pago))
'    ''rsPAgoDetalle!estado_aprobacion = "N"
'    ''rsPAgoDetalle!usr_usuario = GlUsuario
'   '' rsPAgoDetalle!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'   '' rsPAgoDetalle!hora_registro = Format(Time, "hh:mm:ss")
'   '' rsPAgoDetalle.Update
'  End If
'  '****comprobanteM
'  If rscomprobanteM.RecordCount <> 0 Then
'    rsadcomprobanteM.AddNew
'    Call gencodigo
'    rsadcomprobanteM!Cod_Comp = NUMERO 'rscomprobanteM!Cod_Comp
'    rsadcomprobanteM!tipo_comp = "ANP"   'tipo de comprobante de anulacion contable
'    rsadcomprobanteM!Cod_Trans = IIf(IsNull(rscomprobanteM!Cod_Trans), "", rscomprobanteM!Cod_Trans)
'    rsadcomprobanteM!Cod_Trans_Detalle = IIf(IsNull(rsadcomprobanteM!Cod_Trans_Detalle), "1", rsadcomprobanteM!Cod_Trans_Detalle)
'    rsadcomprobanteM!org_codigo = IIf(IsNull(rscomprobanteM!org_codigo), "999", rscomprobanteM!org_codigo)
'    rsadcomprobanteM!ges_gestion = IIf(IsNull(rscomprobanteM!ges_gestion), Year(Date), rscomprobanteM!ges_gestion)
'    rsadcomprobanteM!Num_respaldo = IIf(IsNull(rscomprobanteM!Num_respaldo), "", rscomprobanteM!Num_respaldo)
'    rsadcomprobanteM!fecha_A = CDate(Format(Date, "dd/mm/yyyy")) 'CDate(rscomprobanteM!fecha_A)
'    rsadcomprobanteM!Codigo_beneficiario = IIf(IsNull(rscomprobanteM!Codigo_beneficiario), "", rscomprobanteM!Codigo_beneficiario)
'    rsadcomprobanteM!Codigo_documento = IIf(IsNull(rscomprobanteM!Codigo_documento), "", rscomprobanteM!Codigo_documento)
'    rsadcomprobanteM!glosa = IIf(IsNull(rscomprobanteM!glosa), "", rscomprobanteM!glosa)
'    rsadcomprobanteM!Status = "N" 'IIf(IsNull(rscomprobanteM!Status), "S", rscomprobanteM!Status)
'    rsadcomprobanteM!Usr_Usuario = GlUsuario 'rsadcomprobanteM!usr_usuario
'    rsadcomprobanteM!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'    rsadcomprobanteM!Hora_Registro = Format(Time, "hh:mm:ss")
'    rsadcomprobanteM!tipo_moneda = IIf(IsNull(rscomprobanteM!tipo_moneda), "Bs.", rscomprobanteM!tipo_moneda)
'    rsadcomprobanteM!codigo_solicitud = IIf(IsNull(rscomprobanteM!codigo_solicitud), "", rscomprobanteM!codigo_solicitud)
'    rsaddiario.AddNew
'    rsaddiario!Cod_Comp = NUMERO 'rscomprobanteM!Cod_Comp
'    rsaddiario!tipo_comp = "ANP" 'rscomprobanteM!tipo_comp
'    rsaddiario!Cod_Trans_Detalle = IIf(IsNull(rscomprobanteM!diariodetalle), "1", rscomprobanteM!diariodetalle)
'    rsaddiario!Cod_Comp_C = IIf(IsNull(rscomprobanteM!Cod_Comp_C), 0, rscomprobanteM!Cod_Comp_C)
'    rsaddiario!d_cuenta = rscomprobanteM!h_cuenta
'    rsaddiario!d_subcta1 = rscomprobanteM!h_subcta1
'    rsaddiario!d_subcta2 = rscomprobanteM!h_subcta2
'    rsaddiario!d_Aux1 = rscomprobanteM!h_Aux1
'    rsaddiario!d_Aux2 = rscomprobanteM!h_Aux2
'    rsaddiario!d_Aux3 = rscomprobanteM!h_Aux3
'    rsaddiario!d_cta_Larga = IIf(IsNull(rscomprobanteM!h_cta_Larga), "", rscomprobanteM!h_cta_Larga)
'    rsaddiario!d_montobs = IIf(IsNull(rscomprobanteM!h_montoBs), 0, rscomprobanteM!d_montobs)
'    rsaddiario!d_montoDl = IIf(IsNull(rscomprobanteM!h_montoDl), 0, rscomprobanteM!d_montoDl)
'    rsaddiario!d_Cambio = IIf(IsNull(rscomprobanteM!h_Cambio), 0, rscomprobanteM!d_Cambio)
'    rsaddiario!h_cuenta = rscomprobanteM!d_cuenta
'    rsaddiario!h_subcta1 = rscomprobanteM!d_subcta1
'    rsaddiario!h_subcta2 = rscomprobanteM!d_subcta2
'    rsaddiario!h_Aux1 = rscomprobanteM!d_Aux1
'    rsaddiario!h_Aux2 = rscomprobanteM!d_Aux2
'    rsaddiario!h_Aux3 = rscomprobanteM!d_Aux3
'    rsaddiario!h_cta_Larga = IIf(IsNull(rscomprobanteM!d_cta_Larga), "", rscomprobanteM!d_cta_Larga)
'    rsaddiario!h_montoBs = rscomprobanteM!d_montobs
'    rsaddiario!h_montoDl = rscomprobanteM!d_montoDl
'    rsaddiario!h_Cambio = rscomprobanteM!d_Cambio
'    rsaddiario!Usr_Usuario = GlUsuario
'    rsaddiario!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'    rsaddiario!Hora_Registro = Format(Time, "hh:mm:ss")
'    rsaddiario.Update
'    rsadcomprobanteM.Update
'  End If
'Dim rsanul As ADODB.Recordset
'Set rsanul = New ADODB.Recordset
'If rsanul.State = 1 Then rsanul.Close
'rsanul.CursorLocation = adUseClient
'rsanul.Open "select count(*) as num from co_comprobante_m where cod_trans='" & codigo & " ' and org_codigo='" & org & "' and tipo_comp='ANP'", db, adOpenKeyset, adLockReadOnly
'If rsanul.RecordCount <> 0 Then
'  If rsanul!num <> 0 Then
'     anul999 = 1
'  Else
'     anul999 = 0
'  End If
'End If
''  If rsctabancaria.State = 1 Then rsctabancaria.Close
''  rsctabancaria.CursorLocation = adUseClient
''  rsctabancaria.Open "SELECT Cta_Codigo,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & ctacodigo & "'", db, adOpenKeyset, adLockOptimistic
''  If rsctabancaria.RecordCount <> 0 Then
''    If montoBs <> 0 Then
''      rsctabancaria!cta_acum_anl = IIf(IsNull(rsctabancaria!cta_acum_anl), 0, rsctabancaria!cta_acum_anl) + montoBs
''    Else
''      rsctabancaria!cta_acum_anl = IIf(IsNull(rsctabancaria!cta_acum_anl), 0, rsctabancaria!cta_acum_anl) + montopago
''    End If
''    rsctabancaria.Update
''  End If
'End Sub
Public Sub DEVOLUCION999(Codigo As Integer, org As String, gestion As String)
 Dim comdevol999 As ADODB.Command
 Set comdevol999 = New ADODB.Command ' para obtener los saldos
        With comdevol999
            .CommandType = adCmdStoredProc
            .CommandText = "devolucion999"
            .Parameters.Append comdevol999.CreateParameter("cod", adInteger, adParamInput)
            .Parameters.Append comdevol999.CreateParameter("org", adVarChar, adParamInput, 3)
            .Parameters.Append comdevol999.CreateParameter("gestion", adVarChar, adParamInput, 4)
            .Parameters.Append comdevol999.CreateParameter("usr", adVarChar, adParamInput, 40)
            .Parameters.Append comdevol999.CreateParameter("hora", adVarChar, adParamInput, 8)
            .Parameters.Append comdevol999.CreateParameter("registro", adInteger, adParamOutput)
            .Parameters.Append comdevol999.CreateParameter("numero", adInteger, adParamOutput)
            .Parameters("Cod") = Codigo
            .Parameters("org") = org
            .Parameters("gestion") = gestion
            .Parameters("usr") = GlUsuario
            .Parameters("hora") = Format(Time, "hh:mm:ss")
'           .Parameters("registro")=
            .ActiveConnection = db
            .Execute
            regDEV999 = IIf(IsNull(.Parameters("registro")), 0, .Parameters("registro"))
            numDEV999 = IIf(IsNull(.Parameters("numero")), 0, .Parameters("numero"))
        End With
End Sub
Public Sub buscacomprobante(Codigo As Integer, org As String, gestion As String, tipo As String)
  Dim rsbuscacomp As ADODB.Recordset
  Dim sqlexiste As String
  Set rsbuscacomp = New ADODB.Recordset
  If rsbuscacomp.State = 1 Then rsbuscacomp.Close
  rsbuscacomp.CursorLocation = adUseClient
  Select Case tipo
  Case Is <> "ANL"
      sqlexiste = "select count(*) as numero from co_comprobante_m where cod_trans='" & Codigo & "' and org_codigo='" & org & "' and ges_gestion='" & gestion & "' and tipo_comp='" & tipo & "'"
  Case "ANL"
      sqlexiste = "select count(*) as numero from co_comprobante_m where cod_trans='" & Codigo & "' and org_codigo='" & org & "' and ges_gestion='" & gestion & "' and status<>'S'" & _
                  " and tipo_comp='" & tipo & "'" & " and not(comp_anterior  is null)"
  End Select
      rsbuscacomp.Open sqlexiste, db, adOpenKeyset, adLockReadOnly
      If rsbuscacomp.RecordCount <> 0 Then
        If rsbuscacomp!numero <> 0 Then
          existecomp = 1
        Else
          existecomp = 0
        End If
      End If
  
End Sub
Public Sub Anulacion999(Codigo As Integer, org As String, gestion As String)
Dim comanul999 As ADODB.Command
 Set comanul999 = New ADODB.Command ' para obtener los saldos
        With comanul999
            .CommandType = adCmdStoredProc
            .CommandText = "Anulacion999"
            .Parameters.Append comanul999.CreateParameter("cod", adInteger, adParamInput)
            .Parameters.Append comanul999.CreateParameter("org", adVarChar, adParamInput, 3)
            .Parameters.Append comanul999.CreateParameter("gestion", adVarChar, adParamInput, 4)
            .Parameters.Append comanul999.CreateParameter("usr", adVarChar, adParamInput, 40)
            .Parameters.Append comanul999.CreateParameter("hora", adVarChar, adParamInput, 8)
            .Parameters.Append comanul999.CreateParameter("registro", adInteger, adParamOutput)
            .Parameters.Append comanul999.CreateParameter("numero", adInteger, adParamOutput)
            .Parameters("Cod") = Codigo
            .Parameters("org") = org
            .Parameters("gestion") = gestion
            .Parameters("usr") = GlUsuario
            .Parameters("hora") = Format(Time, "hh:mm:ss")
'           .Parameters("registro")=
            .ActiveConnection = db
            .Execute
            regANL999 = IIf(IsNull(.Parameters("registro")), 0, .Parameters("registro"))
            numANL999 = IIf(IsNull(.Parameters("numero")), 0, .Parameters("numero"))
        End With
End Sub
Public Sub Reversion999(Codigo As Integer, org As String, gestion As String)
Dim comrev999 As ADODB.Command
 Set comrev999 = New ADODB.Command ' para obtener los saldos
        With comrev999
            .CommandType = adCmdStoredProc
            .CommandText = "Reversion999"
            .Parameters.Append comrev999.CreateParameter("cod", adInteger, adParamInput)
            .Parameters.Append comrev999.CreateParameter("org", adVarChar, adParamInput, 3)
            .Parameters.Append comrev999.CreateParameter("gestion", adVarChar, adParamInput, 4)
            .Parameters.Append comrev999.CreateParameter("usr", adVarChar, adParamInput, 40)
            .Parameters.Append comrev999.CreateParameter("hora", adVarChar, adParamInput, 8)
            .Parameters.Append comrev999.CreateParameter("registro", adInteger, adParamOutput)
            .Parameters.Append comrev999.CreateParameter("numero", adInteger, adParamOutput)
            .Parameters("Cod") = Codigo
            .Parameters("org") = org
            .Parameters("gestion") = gestion
            .Parameters("usr") = GlUsuario
            .Parameters("hora") = Format(Time, "hh:mm:ss")
'           .Parameters("registro")=
            .ActiveConnection = db
            .Execute
            regRVT999 = IIf(IsNull(.Parameters("registro")), 0, .Parameters("registro"))
            numRVT999 = IIf(IsNull(.Parameters("numero")), 0, .Parameters("numero"))
        End With
End Sub
Public Sub AnulaTRP(Codigo As Integer, org As String, gestion As String)
Dim comanulTRP As ADODB.Command
 Set comanulTRP = New ADODB.Command ' para obtener los saldos
        With comanulTRP
            .CommandType = adCmdStoredProc
            .CommandText = "AnulaTRP"
            .Parameters.Append comanulTRP.CreateParameter("cod", adInteger, adParamInput)
            .Parameters.Append comanulTRP.CreateParameter("org", adVarChar, adParamInput, 3)
            .Parameters.Append comanulTRP.CreateParameter("gestion", adVarChar, adParamInput, 4)
            .Parameters.Append comanulTRP.CreateParameter("usr", adVarChar, adParamInput, 40)
            .Parameters.Append comanulTRP.CreateParameter("hora", adVarChar, adParamInput, 8)
            .Parameters.Append comanulTRP.CreateParameter("registro", adInteger, adParamOutput)
            .Parameters.Append comanulTRP.CreateParameter("numero", adInteger, adParamOutput)
            .Parameters("Cod") = Codigo
            .Parameters("org") = org
            .Parameters("gestion") = gestion
            .Parameters("usr") = GlUsuario
            .Parameters("hora") = Format(Time, "hh:mm:ss")
'           .Parameters("registro")=
            .ActiveConnection = db
            .Execute
            regANLTRP = IIf(IsNull(.Parameters("registro")), 0, .Parameters("registro"))
            numANLTRP = IIf(IsNull(.Parameters("numero")), 0, .Parameters("numero"))
        End With
End Sub
Public Sub DAC(Codigo As Integer, org As String, gestion As String)
  Dim comDAC As ADODB.Command
  Set comDAC = New ADODB.Command
  With comDAC
    .CommandType = adCmdStoredProc
    .CommandText = "DAC"
    .Parameters.Append comDAC.CreateParameter("codpago", adInteger, adParamInput)
    .Parameters.Append comDAC.CreateParameter("org", adVarChar, adParamInput, 3)
    .Parameters.Append comDAC.CreateParameter("gestion", adVarChar, adParamInput, 4)
    .Parameters.Append comDAC.CreateParameter("usr", adVarChar, adParamInput, 40)
    .Parameters.Append comDAC.CreateParameter("hora", adVarChar, adParamInput, 8)
    .Parameters.Append comDAC.CreateParameter("existe", adInteger, adParamOutput)
    .Parameters("codpago") = Codigo
    .Parameters("org") = org
    .Parameters("gestion") = gestion
    .Parameters("usr") = GlUsuario
    .Parameters("hora") = Format(Time, "hh:mm:ss")
    .ActiveConnection = db
    .Execute
    regDAC = IIf(IsNull(.Parameters("existe")), 0, .Parameters("existe"))
  End With
End Sub
'Public Sub Cmd_ContaConf(codigo As Integer, org As String, gestion As String)
Public Sub Cmd_contabiliza(Codigo As String, org As String, gestion As String)
  Dim comPCC As ADODB.Command
  Set comPCC = New ADODB.Command
  With comPCC
    .CommandType = adCmdStoredProc
    .CommandText = "PCC"
    .Parameters.Append comPCC.CreateParameter("cod", adInteger, adParamInput)
    .Parameters.Append comPCC.CreateParameter("org", adVarChar, adParamInput, 3)
    .Parameters.Append comPCC.CreateParameter("gestion", adVarChar, adParamInput, 4)
    .Parameters.Append comPCC.CreateParameter("usr", adVarChar, adParamInput, 40)
    .Parameters.Append comPCC.CreateParameter("hora", adVarChar, adParamInput, 8)
    .Parameters.Append comPCC.CreateParameter("registro", adInteger, adParamOutput)
    .Parameters.Append comPCC.CreateParameter("numero", adInteger, adParamOutput)
    .Parameters("cod") = Codigo
    .Parameters("org") = org
    .Parameters("gestion") = gestion
    .Parameters("usr") = GlUsuario
    .Parameters("hora") = Format(Time, "hh:mm:ss")
    .ActiveConnection = db
    .Execute
    regPCC = IIf(IsNull(.Parameters("registro")), 0, .Parameters("registro"))
    numPCC = IIf(IsNull(.Parameters("numero")), 0, .Parameters("numero"))
  End With
  If regPCC <> 0 Then
     MsgBox "Contabilización éxitosa", vbExclamation + vbDefaultButton1, "Contabilización"
  Else
     MsgBox "Problemas en la Contabilización", vbExclamation + vbDefaultButton1, "Contabilización"
  End If
End Sub
Public Sub actualizaComp(Codigo As Integer, gestion As String, org As String, Opcion As String)
  Dim comactualiza As ADODB.Command
  Set comactualiza = New ADODB.Command
  With comactualiza
    .CommandType = adCmdStoredProc
    .CommandText = "ActualizaEstado"
    .Parameters.Append comactualiza.CreateParameter("cod", adInteger, adParamInput)
    .Parameters.Append comactualiza.CreateParameter("gestion", adVarChar, adParamInput, 4)
    .Parameters.Append comactualiza.CreateParameter("org", adVarChar, adParamInput, 3)
    .Parameters.Append comactualiza.CreateParameter("opcion", adVarChar, adParamInput, 3)
    .Parameters("cod") = Codigo
    .Parameters("gestion") = gestion
    .Parameters("org") = org
    .Parameters("opcion") = Opcion
    .ActiveConnection = db
    .Execute
  End With
End Sub
Public Sub DevolucionPresup(Codigo As Integer, gestion As String, org As String)
  Dim comdevolucion As ADODB.Command
  Set comdevolucion = New ADODB.Command
  With comdevolucion
    .CommandType = adCmdStoredProc
    .CommandText = "Devolucion"
    .Parameters.Append comdevolucion.CreateParameter("cod", adInteger, adParamInput)
    .Parameters.Append comdevolucion.CreateParameter("org", adVarChar, adParamInput, 3)
    .Parameters.Append comdevolucion.CreateParameter("gestion", adVarChar, adParamInput, 4)
    .Parameters.Append comdevolucion.CreateParameter("usr", adVarChar, adParamInput, 40)
    .Parameters.Append comdevolucion.CreateParameter("hora", adVarChar, adParamInput, 8)
    .Parameters.Append comdevolucion.CreateParameter("registro", adInteger, adParamOutput)
    .Parameters("cod") = Codigo
    .Parameters("org") = org
    .Parameters("gestion") = gestion
    .Parameters("usr") = GlUsuario
    .Parameters("hora") = Format(Time, "hh:mm:ss")
    .ActiveConnection = db
    .Execute
    regRVT = IIf(IsNull(.Parameters("registro")), 0, .Parameters("registro"))
  End With
    
End Sub
Public Sub EstadoDVL(Codigo As Integer)
  '---actualizacion de los estado en CO_comprobante_M 'DVL' o 'DVP' segun corresponda
  Dim comestaDVL As ADODB.Command
  Set comestaDVL = New ADODB.Command
  With comestaDVL
    .CommandType = adCmdStoredProc
    .CommandText = "EstadoDVL"
    .Parameters.Append comestaDVL.CreateParameter("cod", adInteger, adParamInput)
    .Parameters("cod") = Codigo
    .ActiveConnection = db
    .Execute
  End With
End Sub
Public Sub permitectas(cuenta As String, sub1 As String, tipo As String)
  Select Case tipo
    Case "PCE"
      If cuenta = "1111" And sub1 = "02" Then
        MsgBox "Un comprobante PCE no puede manejar cuentas de Bancos", vbExclamation + vbOKOnly, "ERROR EN EL MANEJO DE CUENTAS"
        permite = "1"
      Else
        permite = "0"
      End If
    Case "PCO", "CAM"
  End Select
End Sub

Attribute VB_Name = "ModValPresup"
Option Explicit
Dim rstfc_relacionador_poa_ppto As New ADODB.Recordset
Dim tot_form As Integer
Dim tgn1, ext1 As Double

Public Sub val_presup(adoorigen, GlNombFor)
  Dim rstdestino As New ADODB.Recordset
  Dim rstorigen As New ADODB.Recordset
  Dim rstpagos As New ADODB.Recordset
  Dim rstpago_detalle As New ADODB.Recordset
  Dim rscorrelativo As New ADODB.Recordset
  
  Dim Proyecto1 As String
  Dim Par_Codigo1 As String
  Dim Organismo1 As String
  Dim fte_codigo1 As String
  Dim Org_Codigo1 As String
  Dim pro_Programa1 As String
'  Dim Pro_SubPrograma1 As String
  Dim Pro_Proyecto1 As String
  Dim Pro_Actividad1 As String
  Dim uni_codigo1 As String
  Dim codigo_categoria1 As String
  Dim codigo_convenio1 As String
  
  Dim Fte_contraparte1 As String
  Dim Org_Contraparte1 As String
  
  Dim por_fte_ext1 As Double
  Dim por_fte_nal1 As Double
  Dim codigo_pago1 As Double
  Dim ges_gestion1 As String
  
  Dim swpresup As Integer
  Dim i As Integer
  Dim j As Integer
  Dim v_por_fte(3, 3)

  Dim rectot As Integer
  Dim rstao_solicitud_detalle As New ADODB.Recordset
  Dim rstao_solicitud_recibido As New ADODB.Recordset
  Dim swSubir As String
  Dim tot_reg As Integer
  
  '======== tipo de formualrio F01 ========
  If GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "A") Then
    tot_reg = 0
    Dim Cont_Comp As Integer
    Dim rstdetalle As New ADODB.Recordset
    Set rstdetalle = New ADODB.Recordset
    If rstdetalle.State = 1 Then rstdetalle.Close
    rstdetalle.Open "select * from ao_Solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
    If rstdetalle.RecordCount < 1 Then
      MsgBox "No se puede generar el asiento contable," & vbCrLf & "debido a que el registro no tiene el detalle de montos.", vbOKOnly + vbCritical, "Error al generar el asiento contabl..."
      If rstdetalle.State = 1 Then rstdetalle.Close
      Exit Sub
    Else
      tot_reg = 0
      If rstdetalle!monto_Bolivianos > 0 Then tot_reg = tot_reg + 1
      If rstdetalle!monto_bolivianos_contra > 0 Then tot_reg = tot_reg + 1
    End If
    Set rstao_solicitud_recibido = New ADODB.Recordset
    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
    rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
    db.BeginTrans
    '======== ini registro de co_comprobante_M ========
    Dim rstCodComp As New ADODB.Recordset
    Set rstdestino = New ADODB.Recordset
    For i = 1 To 2 'tot_reg
      If rstdetalle!monto_Bolivianos <= 0 And i = 1 Then
        GoTo etiq
      End If
      If rstdetalle!monto_bolivianos_contra <= 0 And i = 2 Then
        GoTo etiq
      End If
      '======== ini GENERA EL CODIGO DE COMPROBANTE ========
      Set rstCodComp = New ADODB.Recordset
      rstCodComp.CursorLocation = adUseClient
      If rstCodComp.State = 1 Then rstCodComp.Close
      rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'cmbte'", db, adOpenDynamic, adLockOptimistic
      If rstCodComp.RecordCount > 0 Then
        Cont_Comp = Val(rstCodComp!numero_correlativo)
        Cont_Comp = Cont_Comp + 1
        rstCodComp!numero_correlativo = Trim(Str(Cont_Comp))
        rstCodComp.Update
      End If
      If rstCodComp.State = 1 Then rstCodComp.Close
      '======== fin TERMINA GENERACION DE COMPROBANTE ========
      
      '======== ini registro co_comprobantre_m ========
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
      If rstdestino.RecordCount > 0 Then
      End If
      rstdestino.AddNew
      rstdestino!Cod_Comp = Cont_Comp
      rstdestino!cod_trans = ""
      If i = 1 Then
        rstdestino!org_codigo = "999" 'adoorigen!org_codigo_ext
      End If
      If i = 2 Then
        rstdestino!org_codigo = "999"  'rstdestino!org_codigo = adoorigen!org_codigo_contra
      End If
      rstdestino!cod_trans_detalle = 1
      rstdestino!num_respaldo = adoorigen!codigo_unidad & "/" & Str(adoorigen!codigo_solicitud)
      rstdestino!codigo_solicitud = (adoorigen!codigo_solicitud) 'adoorigen!codigo_unidad '& "/" & Str(adoorigen!codigo_solicitud)
      rstdestino!codigo_unidad = (adoorigen!codigo_unidad)
      rstdestino!fecha_A = Format(Date, "dd/mm/yyyy")         'Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
      'rstdestino!codigo_beneficiario = adoorigen!CI_aprueba
      rstdestino!codigo_beneficiario = adoorigen!ci
      rstdestino!Origen = "1"
      'aqui fBuscaFteCorta(fte_1)
      If i = 1 Then
        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_ext) & ": " & Round((rstdetalle!monto_Bolivianos * 100 / (rstdetalle!monto_Bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
      End If
      If i = 2 Then
        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_contra) & ": " & Round((rstdetalle!monto_bolivianos_contra * 100 / (rstdetalle!monto_Bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
      End If
      rstdestino!Status = "N"
      rstdestino!ges_gestion = adoorigen!ges_gestion
      rstdestino!codigo_documento = "D13"
      rstdestino!tipo_comp = "PCE" 'IIf(adoorigen!codigo_tipo = "DEV", "CAD", IIf(adoorigen!codigo_tipo = "REC", "CAR", v_Tipo_Comp(i)))
  '        rstdestino!tipo_moneda = adoorigen!tipo_moneda
      rstdestino!usr_usuario = GlUsuario
      rstdestino!fecha_registro = Date
      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
      rstdestino!tipo_moneda = rstdetalle!tipo_moneda
      rstdestino.Update
      '======== fin registro co_comprobantre_m ========
      
      '======== ini registra CO_diaRIO ========
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from co_diario where Cod_Comp = " & Cont_Comp, db, adOpenKeyset, adLockOptimistic
      If rstdestino.RecordCount > 0 Then
        rstdestino.MoveFirst
      Else
        rstdestino.AddNew
        rstdestino!Cod_Comp = Cont_Comp
      End If
      
      rstdestino!tipo_comp = "PCE"
      rstdestino!d_cuenta = "1127"
      'g--        rstdestino!D_Nombre = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
      rstdestino!d_subcta1 = "02"
      Select Case adoorigen!subcta2
        Case "01" '"Regulares" 'Cargos de Cuenta Regulares
          rstdestino!d_subcta2 = "01"
          rstdestino!d_Aux3 = "00"
        Case "02" '"Otros" 'Cargos de Cuenta Otros
          rstdestino!d_subcta2 = "02"
          rstdestino!d_Aux3 = "00"
        Case "03"  '"PASE" 'Cargos de Cuenta PASE
          rstdestino!d_subcta2 = "03"
          rstdestino!d_Aux3 = "10"
      End Select
      rstdestino!d_Aux1 = "01"
      rstdestino!d_Aux2 = "09"
      ' rstdestino!d_Aux3 = "00"
      rstdestino!d_cta_larga = adoorigen!CI_aprueba
      rstdestino!d_cta_larga = adoorigen!ci             'JQA JUN/2008 ERROR
      rstdestino!d_des_Larga = "-" ' CAMPO PARA ELIMINAR
      If i = 1 Then
        rstdestino!d_montoBs = rstdetalle!monto_Bolivianos
        rstdestino!d_montoDl = rstdetalle!monto_dolares
        rstdestino!d_ctaaux2 = rstdetalle!org_codigo_ext   'g--
      End If
      If i = 2 Then
        rstdestino!d_montoBs = rstdetalle!monto_bolivianos_contra
        rstdestino!d_montoDl = rstdetalle!monto_dolares_contra
        rstdestino!d_ctaaux2 = rstdetalle!org_codigo_contra  'g--
      End If
      rstdestino!d_Cambio = rstdetalle!tipo_cambio
      rstdestino!h_cuenta = "2116"
      'g--        rstdestino!H_Nombre = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
      rstdestino!h_subcta1 = "02"
      rstdestino!h_subcta2 = "00"
      rstdestino!h_Aux1 = "01"
      rstdestino!h_Aux2 = "09"   'g--
      rstdestino!h_Aux3 = "00"
      rstdestino!h_cta_larga = adoorigen!CI_aprueba
      rstdestino!h_des_Larga = "-"   ' CAMPO PARA ELIMINAR
      If i = 1 Then
        rstdestino!h_montoBs = rstdetalle!monto_Bolivianos
        rstdestino!h_montoDl = rstdetalle!monto_dolares
        rstdestino!h_ctaaux2 = rstdetalle!codigo_convenio
        rstdestino!d_ctaaux2 = rstdetalle!codigo_convenio
        rstdestino!d_CtaAux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)
        'rsCo_diario!d_Aux3 = "10"
        'rsCo_diario!d_ctaaux3 = DtCCodigo.Text
      End If
      If i = 2 Then
        rstdestino!h_montoBs = rstdetalle!monto_bolivianos_contra
        rstdestino!h_montoDl = rstdetalle!monto_dolares_contra
        rstdestino!h_ctaaux2 = "FIN_PROPIO" 'rstdetalle!codigo_convenio
        rstdestino!d_ctaaux2 = "FIN_PROPIO" 'rstdetalle!codigo_convenio
        rstdestino!d_CtaAux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)
      End If
      rstdestino!h_Cambio = rstdetalle!tipo_cambio
      'grabar convenios
      'en h_ctaaux2 y en d_ctaaux2
      '      rstdestino!h_ctaaux2 = rstdetalle!codigo_convenio
      '      rstdestino!d_ctaaux2 = rstdetalle!codigo_convenio
      rstdestino!usr_usuario = GlUsuario
      rstdestino!fecha_registro = Date
      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
      '======== fin registra co_diario ========
etiq:
    Next i
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
    If rstdestino.RecordCount > 0 Then
      rstdestino!estado_enviado = "S"
      rstdestino.Update
      rstao_solicitud_recibido.AddNew
      rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
      rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), 0, adoorigen!codigo_solicitud)
      rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
      rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
      rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
      'rstao_solicitud_recibido!swSubir = swSubir
      rstao_solicitud_recibido!usr_usuario = GlUsuario
      rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
      rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
      rstao_solicitud_recibido.Update
    End If
    If rstdestino.State = 1 Then rstdestino.Close
    If rstdetalle.State = 1 Then rstdetalle.Close
    db.CommitTrans
    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
  End If
  '---- fin formulario f01 ----
  
  'If (GlNombFor <> "F01") And (GlNombFor <> "F06") Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
  If ((GlNombFor <> "F02") And (GlNombFor <> "F06")) Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
    Set rstao_solicitud_detalle = New ADODB.Recordset
    If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
    rstao_solicitud_detalle.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
    If rstao_solicitud_detalle.RecordCount > 0 Then
      rectot = rstao_solicitud_detalle.RecordCount
      Fte_contraparte1 = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
      Org_Contraparte1 = rstao_solicitud_detalle!org_codigo_contra
      Dim v_EstPoa(50, 14)
    End If
    If Not (rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
    For i = 1 To rstao_solicitud_detalle.RecordCount
      Set rstfc_relacionador_poa_ppto = New ADODB.Recordset
      If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
      rstfc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & rstao_solicitud_detalle!codigo_poa & "'", db, adOpenKeyset, adLockReadOnly
      If rstfc_relacionador_poa_ppto.RecordCount > 0 Then
      'aqui se puede definir porcentaje
        v_EstPoa(i, 1) = rstao_solicitud_detalle!codigo_poa
        'v_EstPoa(i, 2) = rstfc_relacionador_poa_ppto!Proyecto 'Proyecto1
        v_EstPoa(i, 3) = rstfc_relacionador_poa_ppto!par_codigo 'Par_Codigo1
        v_EstPoa(i, 4) = fBuscaFte(rstfc_relacionador_poa_ppto!org_codigo) 'fte_codigo1
        v_EstPoa(i, 5) = rstfc_relacionador_poa_ppto!org_codigo 'Org_Codigo1
        v_EstPoa(i, 6) = rstfc_relacionador_poa_ppto!pro_programa 'pro_Programa1
        'v_EstPoa(i, 7) = rstfc_relacionador_poa_ppto!pro_subprograma 'Pro_SubPrograma1
        v_EstPoa(i, 8) = rstfc_relacionador_poa_ppto!pro_proyecto 'Pro_Proyecto1
        v_EstPoa(i, 9) = rstfc_relacionador_poa_ppto!pro_actividad 'Pro_Actividad1
        v_EstPoa(i, 10) = rstfc_relacionador_poa_ppto!uni_codigo 'uni_codigo1
        v_EstPoa(i, 11) = IIf(IsNull(rstfc_relacionador_poa_ppto!codigo_Categoria), "xx", rstfc_relacionador_poa_ppto!codigo_Categoria) 'codigo_categoria1
        v_EstPoa(i, 12) = rstfc_relacionador_poa_ppto!codigo_convenio 'codigo_convenio1
        'aqui now consultar con tia la contraparte debe tener estructura.
        If rstao_solicitud_detalle!org_codigo_contra = "" Or rstao_solicitud_detalle!org_codigo_contra = "-" Then
          v_EstPoa(i, 13) = "10"
          v_EstPoa(i, 14) = "111"
        Else
          v_EstPoa(i, 13) = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
          v_EstPoa(i, 14) = rstao_solicitud_detalle!org_codigo_contra
        End If
        If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
        Dim rstfo_formulacion_gasto As New ADODB.Recordset
        Set rstfo_formulacion_gasto = New ADODB.Recordset
        If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
        'rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & pro_Programa1 & "' and pro_subprograma='" & Pro_SubPrograma1 & "' and pro_proyecto='" & Pro_Proyecto1 & "' and pro_actividad='" & Pro_Actividad1 & "' and par_codigo='" & Par_Codigo1 & "' and org_codigo= '" & Org_Codigo1 & "'", db, adOpenKeyset, adLockOptimistic
        rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & v_EstPoa(i, 6) & "' and pro_proyecto='" & v_EstPoa(i, 8) & "' and pro_actividad='" & v_EstPoa(i, 9) & "' and par_codigo='" & v_EstPoa(i, 3) & "' and org_codigo= '" & v_EstPoa(i, 5) & "'", db, adOpenKeyset, adLockOptimistic
        If Not (rstfo_formulacion_gasto.EOF) Then
          If (rstfo_formulacion_gasto!FGS_VIGENTE - rstfo_formulacion_gasto!FGS_compromiso < rstao_solicitud_detalle!monto_Bolivianos) Then  'adoorigen         'adoorigen.adosolicitud.Recordset!monto_dolares ) Then
            'JQA 07/12/01
'            swSubir = "No existe Presup"
'            MsgBox "NO EXISTE Presupuesto para dar curso a la Solicitud ...", vbOKOnly, "ERROR"
'            swpresup = 0
'            Exit Sub
            'JQA 07/12/01
            swpresup = 1    'Borrar despues de habilitar JQA
          Else
            'JQA 07/12/01
            'rstfo_formulacion_gasto!0  = rstfo_formulacion_gasto!fgs_precompromiso  + rstao_solicitud_detalle!monto_bolivianos
            'rstfo_formulacion_gasto.Update
            'JQA 07/12/01
            swpresup = 1
            swSubir = "SI correcto"
          End If
          If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
            swpresup = 1
          Else
            'JQA 07/12/01
'          MsgBox "NO EXISTE Estructura presupuestaria...", vbOKOnly, "ERROR ..."
'          swSubir = "NO Error Estruc.Ppto"
'          swpresup = 0
'          Exit Sub
            'JQA 07/12/01
            swpresup = 1    'Borrar despues de habilitar JQA
          End If
      Else
        MsgBox "NO Existe POA ... ", vbOKOnly, "ERROR ..."
        swSubir = "No existe POA"
        swpresup = 0
        Exit Sub
      End If
'          Else
'            swpresup = 1
'          End If
      rstao_solicitud_detalle.MoveNext
    Next
    
    If swpresup = 1 Then
      'If GlNombFor <> "F01" Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
      If (GlNombFor <> "F02") Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
        If (rstao_solicitud_detalle.RecordCount > 0) And (Not rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
        Set rstao_solicitud_recibido = New ADODB.Recordset
        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
        rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
        db.BeginTrans
        'por_fte_ext
        'por_fte_nal
        For j = 1 To rstao_solicitud_detalle.RecordCount
          'j = 2
          v_por_fte(1, 1) = por_fte_ext1
          v_por_fte(1, 2) = v_EstPoa(j, 4) 'fte_codigo1
          v_por_fte(1, 3) = v_EstPoa(j, 5) 'Org_Codigo1

          v_por_fte(2, 1) = por_fte_nal1
          v_por_fte(2, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
          v_por_fte(2, 3) = v_EstPoa(j, 14) 'Org_Contraparte1

          v_por_fte(3, 1) = por_fte_nal1
          v_por_fte(3, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
          v_por_fte(3, 3) = v_EstPoa(j, 14) 'Org_Contraparte1
          
'          If Trim(v_EstPoa(j, 12)) <> "FIN_PROPIO" Then
'            tot_form = 2
'          Else
'            tot_form = 1
'          End If

          Dim SwEsBase As Integer
          Dim ValEsBase As Double
    '          ValEsBase = v_por_fte(1, 1)
    '          For I = 1 To tot_form
    '            If v_por_fte(I, 1) > ValEsBase Then
    '              SwEsBase = I
    '              ValEsBase = v_por_fte(I, 1)
    '            End If
    '          Next
          
        'AQUI UN SOLO FINANCIADOR
          'For i = 1 To tot_form
          '        Print rstpagos!monto_bolivianos
          'Next
          For i = 1 To tot_form 'dos
            Set rstpagos = New ADODB.Recordset
            If rstpagos.State = 1 Then rstpagos.Close
            If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
              rstpagos.Open "select * from pagos_espera where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
            Else
              rstpagos.Open "select * from pagos where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
            End If
            rstpagos.AddNew
            
            '======== ini GENERA EL CODIGO DE COMPROBANTE ========
             Set rscorrelativo = New ADODB.Recordset
             If rscorrelativo.State = 1 Then rscorrelativo.Close
             rscorrelativo.Open "select * from fc_organismo_financiamiento where org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenDynamic, adLockOptimistic
             If rscorrelativo.RecordCount > 0 Then
                       codigo_pago1 = Val(rscorrelativo!correlativo)
                       codigo_pago1 = codigo_pago1 + 1
                       rscorrelativo!correlativo = Trim(Str(codigo_pago1))
                       rscorrelativo.Update
                       rstpagos!codigo_pago = codigo_pago1
                       rstpagos!nro_comprobante_anterior = codigo_pago1
             End If
             If rscorrelativo.State = 1 Then rscorrelativo.Close
            '======== fin TERMINA GENERACION DE CODIGO DE COMPROBANTE ========
            
            '==== ini generación de correlativo ====
'            Set rscorrelativo = New ADODB.Recordset
'            If rscorrelativo.State = 1 Then rscorrelativo.Close
'            If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
'              rscorrelativo.Open "select * from fc_correlativos_espera", db, adOpenKeyset, adLockOptimistic
'            Else
'              rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
'            End If
'            If v_por_fte(i, 3) = "111" Then  'TGN
'              If Not IsNull(rscorrelativo!correl_org111) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                rscorrelativo!correl_org111 = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "112" Then 'TGNP
'              If Not IsNull(rscorrelativo!correl_org112) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                rscorrelativo!correl_org112 = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "114" Then    'If Org_Codigo1 = "114" Then 'RECON
'              If Not IsNull(rscorrelativo!correl_org114) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                rscorrelativo!correl_org114 = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "344" Then 'UNICEF
''            codigo_pago1 = 1
'              If Not IsNull(rscorrelativo!correl_org344) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                rscorrelativo!correl_org344 = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "381" Then  'FAD
'              If Not IsNull(rscorrelativo!correl_org381) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org381) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org381) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org381) + 1)
'                rscorrelativo!correl_org381 = Val(Val(rscorrelativo!correl_org381) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "411" Then  'BID
'              If Not IsNull(rscorrelativo!correl_org411) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                rscorrelativo!correl_org411 = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "415" Then  'IDA
'              If Not IsNull(rscorrelativo!correl_org415) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                rscorrelativo!correl_org415 = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "516" Then  'KFW
'              If Not IsNull(rscorrelativo!correl_org516) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                rscorrelativo!correl_org516 = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "541" Then  'ALEM
'              If Not IsNull(rscorrelativo!correl_org541) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                rscorrelativo!correl_org541 = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "551" Then  'DIN
'              If Not IsNull(rscorrelativo!correl_org551) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                rscorrelativo!correl_org551 = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "556" Then  'HOL
'              If Not IsNull(rscorrelativo!correl_org556) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                rscorrelativo!correl_org556 = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "565" Then  'SUE
'              If Not IsNull(rscorrelativo!correl_org565) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                rscorrelativo!correl_org565 = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "999" Then  'S/N
'              If Not IsNull(rscorrelativo!correl_org999) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                rscorrelativo!correl_org999 = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "Org14" Then
'              If Not IsNull(rscorrelativo!correl_org14) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org14) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org14) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org14) + 1)
'                rscorrelativo!correl_org14 = CDbl(CDbl(rscorrelativo!correl_org14) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "Org15" Then
'              If Not IsNull(rscorrelativo!correl_org15) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org15) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org15) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org15) + 1)
'                rscorrelativo!correl_org15 = CDbl(CDbl(rscorrelativo!correl_org15) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "Org16" Then
'              If Not IsNull(rscorrelativo!correl_org16) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org16) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org16) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org16) + 1)
'                rscorrelativo!correl_org16 = CDbl(CDbl(rscorrelativo!correl_org16) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "Org17" Then
'              If Not IsNull(rscorrelativo!correl_org17) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org17) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org17) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org17) + 1)
'                rscorrelativo!correl_org17 = CDbl(CDbl(rscorrelativo!correl_org17) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "Org18" Then
'              If Not IsNull(rscorrelativo!correl_org18) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org18) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org18) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org18) + 1)
'                rscorrelativo!correl_org18 = CDbl(CDbl(rscorrelativo!correl_org18) + 1)
'                rscorrelativo.Update
'              Else
'                rscorrelativo!correl_org18 = 0
'                rscorrelativo.Update
'              End If
'            End If
'            If v_por_fte(i, 3) = "514" Then
'              If Not IsNull(rscorrelativo!correl_org514) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                rscorrelativo!correl_org514 = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                rscorrelativo.Update
'              Else
'                rscorrelativo!correl_org514 = 0
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "517" Then  'GTZ
'              If Not IsNull(rscorrelativo!correl_org517) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                rscorrelativo!correl_org517 = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                rscorrelativo.Update
'              End If
'            End If
'
'            If v_por_fte(i, 3) = "528" Then  'AECI
'              If Not IsNull(rscorrelativo!correl_org528) Then
'                rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                rscorrelativo!correl_org528 = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                rscorrelativo.Update
'              End If
'            End If
            '==== fin generación de correlativo ====
            
            'MsgBox "Comprobante : " & codigo_pago1 & vbCrLf & "Organismo :     " & v_por_fte(i, 3), vbInformation + vbOKOnly, " Generando el Comprobante..."
            rstpagos!codigo_pago = codigo_pago1
            rstpagos!org_codigo = v_por_fte(i, 3)
            If i = 1 Then
              rstpagos!uni_codigo = v_EstPoa(j, 10) 'v_EstPoa(I, 10) 'uni_codigo1
              rstpagos!codigo_Categoria = v_EstPoa(j, 11) 'v_EstPoa(I, 11) 'codigo_categoria1
              rstpagos!codigo_convenio = v_EstPoa(j, 12) 'codigo_convenio1
            End If
            If i = 2 Then
'              rstpagos!uni_codigo = v_EstPoa(I - 1, 10) 'uni_codigo1
'              If rstao_solicitud_detalle!por_fte_nal = 100 Or rstao_solicitud_detalle!por_fte_ext = 100 Then
'                rstpagos!codigo_categoria = v_EstPoa(I - 1, 11)
'                rstpagos!codigo_convenio = v_EstPoa(I - 1, 12)
'              Else
'                v_por_fte(I, 3) = "S/C TGNP"
                rstpagos!codigo_convenio = fbusCatConv(v_por_fte(i, 3), 1)  '"S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1
'                Print v_por_fte(j, 3)
                rstpagos!codigo_Categoria = fbusCatConv(v_por_fte(i, 3), 2) '"S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1
'              End If
            End If
            
            If i = 3 Then
              rstpagos!uni_codigo = v_EstPoa(1, 10) 'uni_codigo1
'              rstpagos!codigo_categoria = "S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1
'              rstpagos!codigo_convenio = "S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1
              rstpagos!codigo_convenio = fbusCatConv(v_por_fte(i, 3), 1)  '"S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1
              rstpagos!codigo_Categoria = fbusCatConv(v_por_fte(i, 3), 2) '"S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1
            End If
            'rstpagos!codigo_orden  =         'documento de respaldo
            rstpagos!codigo_solicitud = adoorigen!codigo_solicitud
            rstpagos!codigo_unidad = adoorigen!codigo_unidad 'nuevo
            rstpagos!FTE_codigo = v_por_fte(i, 2)
            rstpagos!justificacion = adoorigen!justificacion_solicitud   'adoorigen.txtjustifica
            rstpagos!tipo_moneda = rstao_solicitud_detalle!tipo_moneda 'adoorigen!tipo_moneda   'DtCDenominacion_moneda.bounttext  '"Bs." 'DtCTipoMoneda.Text
            If i = 1 Then
'              If rstao_solicitud_detalle!por_fte_nal = 100 Or rstao_solicitud_detalle!por_fte_ext = 100 Then
                rstpagos!monto_Bolivianos = IIf(IsNull(rstpagos!monto_Bolivianos), 0, rstpagos!monto_Bolivianos) + (rstao_solicitud_detalle!monto_Bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
                rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + rstao_solicitud_detalle!monto_dolares 'adoorigen!monto_dolares   '- adoorigen!monto_dolares_contra  '* por_fte_ext1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_ext1
'              End If
              
              If rstao_solicitud_detalle!monto_Bolivianos > 0 Then
                rstpagos!es_base = "S"
              Else
                rstpagos!es_base = "N"
              End If
              
            End If
            If i = 2 Then
              If v_EstPoa(j, 12) <> "FIN_PROPIO" Then
                ext1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
                tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
                rstpagos!monto_Bolivianos = IIf(IsNull(rstpagos!monto_Bolivianos), 0, rstpagos!monto_Bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
                rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
              Else
                rstpagos!monto_Bolivianos = Val(rstao_solicitud_detalle!monto_bolivianos_contra)
                rstpagos!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares_contra)
              End If
              If rstao_solicitud_detalle!monto_Bolivianos > 0 Then
                rstpagos!es_base = "N"
              Else
                rstpagos!es_base = "S"
              End If
            End If

            If i = 3 Then
              tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 3)
              rstpagos!monto_Bolivianos = IIf(IsNull(rstpagos!monto_Bolivianos), 0, rstpagos!monto_Bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
              rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
              rstpagos!es_base = "I"
            End If
            'rstpagos!liquido_pagar  = "0" 'Val(TxtLiquido.Text)
            'rstpagos!tipo_formulario = GlNombFor
            rstpagos!formulario = GlNombFor
            'rstpagos!estado_aprobacion  = "X"
            rstpagos!estado_devengado = ""
            'alb28082002
            If GlNombFor = "F11" Then
              rstpagos!estado_compromiso = "S"
              rstpagos!es_licitacion = "D"
              
            'AQUI ULTIMO
            Else
              rstpagos!estado_compromiso = "S"
'              rstpagos!codigo_poa = rstao_solicitud_detalle!codigo_poa 'adoorigen!codigo_poa
            End If

            If GlNombFor = "F05" Or GlNombFor = "F10" Then
              rstpagos!duracion_estimada_tiempo = adoorigen!duracion_estimada_tiempo
              rstpagos!duracion_estimada_numero = adoorigen!duracion_estimada_numero
              rstpagos!por_tiempo = adoorigen!por_tiempo
              rstpagos!estado_compromiso = "S"
              rstpagos!estado_devengado = ""
              rstpagos!fecha_estimada_inicio = IIf(IsNull(adoorigen!fecha_estimada_inicio), Date, Format(adoorigen!fecha_estimada_inicio, "dd/mm/yyyy"))
              rstpagos!Lista_adjunta = adoorigen!Lista_adjunta
              rstpagos!periodo_de_trabajo = ""
            End If
            'rstpagos!estado_devengado  = ""
            'rstpagos!estado_pagado  = ""
            If GlNombFor = "F03" Or GlNombFor = "F12" Then
              rstpagos!tipo_formulario = "CYD"
              rstpagos!estado_compromiso = "N"
              rstpagos!estado_devengado = "N"
            Else
              rstpagos!tipo_formulario = "COM"
            End If
            If (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
              rstpagos!tipo_formulario = "REG"
              rstpagos!estado_compromiso = "N"
              rstpagos!estado_devengado = "N"
              rstpagos!estado_pagado = "N"
            End If
            If (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "A")) Then
              rstpagos!tipo_formulario = "REG"
              rstpagos!estado_compromiso = "S"
              rstpagos!estado_devengado = "S"
              rstpagos!estado_pagado = "N"
            End If
            rstpagos!tipo_comp = "DAC"
            rstpagos!fecha_egreso = Format(Date, "dd/mm/yyyy") 'CDate(adoorigen!fecha_recepcion)   ', "dd/mm/aaaa
            rstpagos!ges_gestion = Year(Date)
            ges_gestion1 = Year(Date)
            
            rstpagos!usr_usuario = GlUsuario
            rstpagos!fecha_registro = Date  ' Format(Date, "dd/mm/aaaa
            rstpagos!hora_registro = Format(Time, "hh:mm:ss")
            rstpagos.Update
            If rstpagos.State = 1 Then rstpagos.Close
            '======== fin graba pagos ========
            
            '======== ini graba pago_detalle ========
            Set rstpago_detalle = New ADODB.Recordset
            If rstpago_detalle.State = 1 Then rstpago_detalle.Close
            If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
                rstpago_detalle.Open "select * from pago_detalle_espera where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
            Else
                rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
            End If
            If rstpago_detalle.RecordCount > 0 Then
                rstpago_detalle.MoveFirst
            Else
                rstpago_detalle.AddNew
            End If
            rstpago_detalle!codigo_pago = codigo_pago1
            rstpago_detalle!ges_gestion = ges_gestion1
            rstpago_detalle!org_codigo = v_por_fte(i, 3)
            rstpago_detalle!codigo_pago_detalle = rstpago_detalle.RecordCount
            
            rstpago_detalle!par_codigo = v_EstPoa(j, 3) 'Par_Codigo1
            rstpago_detalle!pro_programa = v_EstPoa(j, 6) 'pro_Programa1
'            rstpago_detalle!pro_subprograma = v_EstPoa(j, 7) 'Pro_SubPrograma1
            rstpago_detalle!pro_proyecto = v_EstPoa(j, 8) 'Pro_Proyecto1
            rstpago_detalle!pro_actividad = v_EstPoa(j, 9) 'Pro_Actividad1
            rstpago_detalle!codigo_beneficiario = adoorigen!ci      'adoorigen!CI_aprueba
            '==== ini porcentajes ====
            rstpago_detalle!codigo_poa = rstao_solicitud_detalle!codigo_poa 'adoorigen!codigo_poa
            If i = 1 Then
              rstpago_detalle!monto_total = Val(rstao_solicitud_detalle!monto_Bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
              rstpago_detalle!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares)  'adoorigen!monto_dolares   '- adoorigen!monto_dolares_contra  '* por_fte_ext1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_ext1
              'If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
              If GlNombFor <> "F06" And GlNombFor <> "F07" And GlNombFor <> "F02" Then
                rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_ext)
                rstpago_detalle!Porcentaje = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
              End If
            End If
            If i = 2 Then
'              rstpago_detalle!monto_total = rstao_solicitud_detalle!monto_bolivianos_contra  'adoorigen!monto_bolivianos_contra   'adoorigen!monto_bolivianos  * por_fte_nal1 'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_nal1
'              rstpago_detalle!monto_dolares = rstao_solicitud_detalle!monto_dolares_contra 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
              If v_EstPoa(j, 12) = "3096-BO" Then
                ext1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
                tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
                rstpago_detalle!monto_total = ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
                rstpago_detalle!monto_dolares = ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
                'If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
                If GlNombFor <> "F06" And GlNombFor <> "F07" And GlNombFor <> "F02" Then
                  rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_nal)
                End If
              Else
                rstpago_detalle!monto_total = Val(rstao_solicitud_detalle!monto_bolivianos_contra)  'adoorigen!monto_bolivianos_contra   'adoorigen!monto_bolivianos  * por_fte_nal1 'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_nal1
                rstpago_detalle!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares_contra) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
                If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
                  rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_nal)
                  rstpago_detalle!Porcentaje = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
                End If
              End If
'''              rstpago_detalle!monto_total = rstao_solicitud_detalle!monto_bolivianos_contra  'adoorigen!monto_bolivianos_contra   'adoorigen!monto_bolivianos  * por_fte_nal1 'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_nal1
'''              rstpago_detalle!monto_dolares = rstao_solicitud_detalle!monto_dolares_contra 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
            End If
            If i = 3 Then
              tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 3)
              rstpago_detalle!monto_total = ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
              rstpago_detalle!monto_dolares = ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
              'If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
              If GlNombFor <> "F06" And GlNombFor = "F07" And GlNombFor = "F02" Then
                rstpago_detalle!Porcentaje = tgn1
              End If
            End If
            '==== fin porcentajes ====
            
            'rstpago_detalle!Deducciones  = Val(TxtDeducciones.Text)
            'rstpago_detalle!saldo_bolivianos  = Val(TxtSaldo.Text)
            rstpago_detalle!tipo_cambio = Val(rstao_solicitud_detalle!tipo_cambio) 'adoorigen!tipo_caMBIO    'adoorigen.adosolicitud.Recordset!tipo_cambio
            rstpago_detalle!estado_aprobacion = "N"
            rstpago_detalle!fecha_pago = Format(Date, "DD/MM/YYYY")  ', "dd/mm/aaaa
            rstpago_detalle!fecha_registro = Format(Date, "DD/MM/YYYY")
            rstpago_detalle!usr_usuario = GlUsuario
            rstpago_detalle!hora_registro = Format(Time, "hh:mm:ss")
            rstpago_detalle.Update
            '======== fin graba pago_detalle
          Next
          Set rstdestino = New ADODB.Recordset
          If rstdestino.State = 1 Then rstdestino.Close
          rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
          If rstdestino.RecordCount > 0 Then
            rstdestino!estado_enviado = "S"
            rstdestino.Update
            rstao_solicitud_recibido.AddNew
            rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
            rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), CStr(0), CStr(adoorigen!codigo_solicitud))
            rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
            rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
            rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
            rstao_solicitud_recibido!fecha_solicitud = Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
            rstao_solicitud_recibido!swSubir = swSubir
            rstao_solicitud_recibido!usr_usuario = GlUsuario
            rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
            rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
            rstao_solicitud_recibido.Update
          End If
          If rstdestino.State = 1 Then rstdestino.Close
          rstao_solicitud_detalle.MoveNext
        Next
        db.CommitTrans
        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
      End If
    End If
  End If
  
  '---- ini formulario F06 ----
  If GlNombFor = "F06" Then
    
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from lo_pagos_conformidad where ges_gestion = '0' ", db, adOpenKeyset, adLockOptimistic
    rstdestino.AddNew
    rstdestino!ges_gestion = adoorigen!ges_gestion
    rstdestino!codigo_unidad = adoorigen!codigo_unidad
    rstdestino!codigo_grupo = adoorigen!codigo_solicitud_ant
    rstdestino!NUMERO_PAGO = adoorigen!Nro_pagos
    rstdestino!codigo_beneficiario = adoorigen!CI_aprueba
    
'    rstdestino!ges_gestion = adoorigen!ges_gestion
'    rstdestino!ges_gestion = adoorigen!ges_gestion
'    rstdestino!ges_gestion = adoorigen!ges_gestion
'    rstdestino!ges_gestion = adoorigen!ges_gestion
'    rstdestino!ges_gestion = adoorigen!ges_gestion
    
    Set rstorigen = New ADODB.Recordset
    If rstorigen.State = 1 Then rstorigen.Close
    rstorigen.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
    If rstorigen.RecordCount > 0 Then
      rstdestino!tipo_moneda = rstorigen!tipo_moneda
      rstdestino!monto_bs_ext = CDbl(rstorigen!monto_Bolivianos)
      rstdestino!monto_dol_ext = CDbl(rstorigen!monto_dolares)
      rstdestino!conformidad = "S"
      rstdestino!enviadoaudapre = "S"
      rstdestino!confo_procesada = "N"
    End If
    rstdestino!usr_usuario = GlUsuario
    rstdestino!fecha_registro = Format(Date, "dd/mm/yyyy")
    rstdestino!hora_registro = Format(Time, "hh:mm:ss")
    rstdestino.Update
    
    
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
    If rstdestino.RecordCount > 0 Then
      rstdestino!estado_enviado = "S"
      rstdestino.Update
      Set rstao_solicitud_recibido = New ADODB.Recordset
      If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
      rstao_solicitud_recibido.Open "select * from ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
      rstao_solicitud_recibido.AddNew
      rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
      rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), 0, adoorigen!codigo_solicitud)
      rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
      rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
      rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
      'rstao_solicitud_recibido!swSubir = swSubir
      rstao_solicitud_recibido!usr_usuario = GlUsuario
      rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
      rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
      rstao_solicitud_recibido.Update
    End If
    If rstdestino.State = 1 Then rstdestino.Close
    
'captu    '  0 descripcion_grupo varchar 50  0 0 0 ('')  0     0
'captu    '  0 concepto  varchar 250 0 0 0 ('')  0     0
    
    '  0 antecedente varchar 250 0 0 0 ('')  0     0
    '  0 nombre_proveedor  varchar 30  0 0 0 ('')  0     0
    '  0 idBeneficiario varchar 15  0 0 0 ('')  0     0
    
    '  0 fecha_envio datetime  8 0 0 0 (getdate()) 0     0
    '  0 NCite_conformidad char  15  0 0 0 ('')  0     0
    '  0 FCite_conformidad datetime  8 0 0 1 (getdate()) 0     0
    '  0 migrado char  1 0 0 0 ('N') 0     0
    '  0 Usr_Usuario varchar 15  0 0 0 ('')  0     0
    '  0 Fecha_Registro  datetime  8 0 0 0 (getdate()) 0     0
    '  0 Hora_Registro varchar 8 0 0 0 ('')  0     0
'captu    '  0 Emite_Factura char  1 0 0 0 ('N') 0     0
    '  0 Sesion  int 4 10  0 0 (0) 0     0
    '  0 porcentaje_pago numeric 9 18  2 0 (100) 0     0
  End If
  '---- fin formulario f06 ----
  

End Sub

Function fBuscaFte(orga_1)
  '======== BUSCA EL CODIGO DE FUENTE EN BASE AL ORGANISMO ========
  Dim rstfte As New ADODB.Recordset
  Set rstfte = New ADODB.Recordset
  If rstfte.State = 1 Then rstfte.Close
  rstfte.Open "SELECT * From fc_organismo_financiamiento WHERE Org_codigo = '" & orga_1 & "'", db, adOpenKeyset, adLockReadOnly
  If rstfte.RecordCount > 0 Then
    fBuscaFte = rstfte!FTE_codigo
  Else
    fBuscaFte = "Err Fte"
  End If
  If rstfte.State = 1 Then rstfte.Close
End Function

Function fBuscaOrgCorta(orga_1)
  '======== BUSCA LA DESCRIPCION CORTA DEL ORGANISMO EN BASE AL CODIGO ========
  Dim rstorg As New ADODB.Recordset
  Set rstorg = New ADODB.Recordset
  If rstorg.State = 1 Then rstorg.Close
  rstorg.Open "SELECT * From fc_organismo_financiamiento WHERE org_codigo = '" & orga_1 & "'", db, adOpenKeyset, adLockReadOnly
  If rstorg.RecordCount > 0 Then
    fBuscaOrgCorta = rstorg!org_sigla
  Else
    fBuscaOrgCorta = "XXX"
  End If
  If rstorg.State = 1 Then rstorg.Close
End Function

Function verporcen(par1, proy1, conv1, cat1, que)
' que = 1 pcen externo
' que = 2 pcen nal 1
' que = 3 pcen nal 2
  Dim rstfc_porcentaje_fte As New ADODB.Recordset
  Set rstfc_porcentaje_fte = New ADODB.Recordset
  If rstfc_porcentaje_fte.State = 1 Then rstfc_porcentaje_fte.Close
  rstfc_porcentaje_fte.Open "select * from so_porcentaje_convenio where Par_Codigo = '" & par1 & "' and codigo_convenio = '" & conv1 & "' and pro_proyecto = '" & proy1 & "' and codigo_categoria = '" & cat1 & "' ", db, adOpenKeyset, adLockReadOnly
  If rstfc_porcentaje_fte.RecordCount > 0 Then
    If par1 = "39600" Then
      verporcen = 100
    Else
      If que = 1 Then verporcen = rstfc_porcentaje_fte!prc_porcentaje
      If que = 2 Then verporcen = rstfc_porcentaje_fte!prc_porcentajeNal
      If que = 3 Then verporcen = (100 - rstfc_porcentaje_fte!prc_porcentaje) - rstfc_porcentaje_fte!prc_porcentajeNal
      'If verifica_porcen = 0 Then MsgBox "El porcentaje encontrado es igual a cero (0)", vbOKOnly + vbInformation, "Error en el Clasificador..."
'        verifica_porcen = rstfc_porcentaje_fte!por_fte_ext
    End If
  Else
    MsgBox "No se pudo determinar los porcentajes " & vbCrLf & _
    "para las fuentes de financiamiento." & vbCrLf & " ERROR en el CLASIFICADOR.", vbOKOnly + vbExclamation, "Error en el relacionador de porcentajes..."
  End If
  If rstfc_porcentaje_fte.State = 1 Then rstfc_porcentaje_fte.Close

End Function

Function ver_confor(adoorigen, GlNombFor)
  Dim rstdestino As New ADODB.Recordset
  Set rstdestino = New ADODB.Recordset
  If rstdestino.State = 1 Then rstdestino.Close
  rstdestino.Open "select * from ac_pagos_cronograma_detalle_1 where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_grupo = " & adoorigen!codigo_grupo & " and numero_pago = " & adoorigen!NUMERO_PAGO & " and idfuncionario = " & adoorigen!idfuncionario, db, adOpenKeyset, adLockReadOnly
  If rstdestino.RecordCount > 0 Then
'    swSubir = "Si correcto"
'    swconfor = 0
    ver_confor = 1
    'Exit Function
  Else
    MsgBox "NO EXISTE registro de este pago...", vbOKOnly, "ERROR ..."
'    swSubir = "NO Error No Existe"
'    swconfor = 0
    ver_confor = 0
    'Exit Function
  End If
End Function

Public Sub imprimereportedeactualizacion()
  Dim rstorigen As New ADODB.Recordset
  Dim sesion1 As Integer
  Dim rssesion As New ADODB.Recordset
  Dim rstao_solicitud_recibido As New ADODB.Recordset
  Set rssesion = New ADODB.Recordset
  If rssesion.State = 1 Then rssesion.Close
  rssesion.Open "select max (sesion) as sesion from ac_pagos_conformidad ", db, adOpenStatic, adLockReadOnly
  If IsNull(rssesion!sesion) Then
    sesion1 = 0
  Else
    sesion1 = rssesion!sesion
  End If
  If rssesion.State = 1 Then rssesion.Close

  Set rstorigen = New ADODB.Recordset
  If rstorigen.State = 1 Then rstorigen.Close
  rstorigen.Open "SELECT * FROM ac_pagos_conformidad where sesion = " & sesion1, db, adOpenKeyset, adLockOptimistic
  
  Set rstao_solicitud_recibido = New ADODB.Recordset
  If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
  rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
  While Not rstorigen.EOF
    rstao_solicitud_recibido.AddNew
    rstao_solicitud_recibido!ges_gestion = IIf(IsNull(rstorigen!ges_gestion), "", rstorigen!ges_gestion)
    rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(rstorigen!codigo_grupo), "0", CStr(rstorigen!codigo_grupo) & "/" & CStr(rstorigen!NUMERO_PAGO))
    rstao_solicitud_recibido!formulario = "F02"
    rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(rstorigen!codigo_unidad), "", rstorigen!codigo_unidad)
    rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(rstorigen!Concepto), "", rstorigen!Concepto)
    rstao_solicitud_recibido!fecha_solicitud = Format(rstorigen!fecha_ENVIO, "dd/mm/yyyy")
    rstao_solicitud_recibido!swSubir = rstorigen!confo_procesada
    rstao_solicitud_recibido!usr_usuario = GlUsuario
    rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
    rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
    rstao_solicitud_recibido.Update
    rstorigen.MoveNext
  Wend
  If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
  If rstorigen.State = 1 Then rstorigen.Close
End Sub


Function fbusCatConv(orga_1, que)
  '======== BUSCA EL CONVENIO Y CATEGORIA AL ORGANISMO ========
  '======== QUE: = 1 CONVENIO; = QUE = 2 CATEGORIA
  Dim rstcnv As New ADODB.Recordset
  Set rstcnv = New ADODB.Recordset
  If rstcnv.State = 1 Then rstcnv.Close
  rstcnv.Open "SELECT * From fc_categoria_financiador WHERE Org_codigo = '" & orga_1 & "'", db, adOpenKeyset, adLockReadOnly
  If rstcnv.RecordCount > 0 Then
    Select Case que
      Case 1
        fbusCatConv = rstcnv!codigo_convenio
      Case 2
        fbusCatConv = rstcnv!codigo_Categoria
    End Select
  Else
    If que = 1 Then fbusCatConv = "Err.Conv"
    If que = 2 Then fbusCatConv = "Err.Cate"
  End If
  If rstcnv.State = 1 Then rstcnv.Close
End Function


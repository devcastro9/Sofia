Attribute VB_Name = "verifica"
Option Explicit
    Dim Org_Codigo1 As String
    Public por_fte_ext1 As Double

Public Function verifica_ppto(adoorigen, GlNombFor)
  Dim rstdestino As New ADODB.Recordset
  Dim rstfc_relacionador_poa_ppto As New ADODB.Recordset
  Dim rstorigen As New ADODB.Recordset
  Dim rstpagos As New ADODB.Recordset
  Dim rstpago_detalle As New ADODB.Recordset
  Dim rscorrelativo As New ADODB.Recordset
  
  Dim Proyecto1 As String
  Dim Par_Codigo1 As String
  Dim Organismo1 As String
  Dim fte_codigo1 As String
  Dim pro_Programa1 As String
  Dim Pro_SubPrograma1 As String
  Dim Pro_Proyecto1 As String
  Dim Pro_Actividad1 As String
  Dim uni_codigo1 As String
  Dim codigo_categoria1 As String
  Dim codigo_convenio1 As String
  
  Dim Fte_contraparte1 As String
  Dim Org_Contraparte1 As String
  
  Dim por_fte_nal1 As Double
  Dim codigo_pago1 As Double
  Dim ges_gestion1 As String
  
  Dim swpresup As Integer
  Dim i As Integer
  Dim j As Integer
  Dim v_por_fte(2, 3)

  Dim rectot As Integer
  Dim rstao_solicitud_detalle As New ADODB.Recordset
  
  If GlNombFor <> "F01" Then
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
        v_EstPoa(i, 1) = rstao_solicitud_detalle!codigo_poa
'        v_EstPoa(i, 2) = rstfc_relacionador_poa_ppto!Proyecto 'Proyecto1
        v_EstPoa(i, 3) = rstfc_relacionador_poa_ppto!par_codigo 'Par_Codigo1
        v_EstPoa(i, 4) = fBuscaFte(rstfc_relacionador_poa_ppto!org_codigo) 'fte_codigo1
        v_EstPoa(i, 5) = rstfc_relacionador_poa_ppto!org_codigo 'Org_Codigo1
        v_EstPoa(i, 6) = rstfc_relacionador_poa_ppto!pro_programa 'pro_Programa1
'        v_EstPoa(i, 7) = rstfc_relacionador_poa_ppto!pro_subprograma 'Pro_SubPrograma1
        v_EstPoa(i, 8) = rstfc_relacionador_poa_ppto!pro_proyecto 'Pro_Proyecto1
        v_EstPoa(i, 9) = rstfc_relacionador_poa_ppto!pro_actividad 'Pro_Actividad1
        v_EstPoa(i, 10) = rstfc_relacionador_poa_ppto!uni_codigo 'uni_codigo1
        v_EstPoa(i, 11) = IIf(IsNull(rstfc_relacionador_poa_ppto!codigo_categoria), "xx", rstfc_relacionador_poa_ppto!codigo_categoria) 'codigo_categoria1
        v_EstPoa(i, 12) = rstfc_relacionador_poa_ppto!codigo_convenio 'codigo_convenio1
'aqui now consultar con tia la contraparte debe tener estructura.
        v_EstPoa(i, 13) = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
        v_EstPoa(i, 14) = rstao_solicitud_detalle!org_codigo_contra
        If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
        Dim rstfo_formulacion_gasto As New ADODB.Recordset
        Set rstfo_formulacion_gasto = New ADODB.Recordset
        If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
'        rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & pro_Programa1 & "' and pro_subprograma='" & Pro_SubPrograma1 & "' and pro_proyecto='" & Pro_Proyecto1 & "' and pro_actividad='" & Pro_Actividad1 & "' and par_codigo='" & Par_Codigo1 & "' and org_codigo= '" & Org_Codigo1 & "'", db, adOpenKeyset, adLockOptimistic

        rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & v_EstPoa(i, 6) & "' AND pro_proyecto = '" & v_EstPoa(i, 8) & "' and pro_actividad = '" & v_EstPoa(i, 9) & "' and par_codigo = '" & v_EstPoa(i, 3) & "' and org_codigo = '" & v_EstPoa(i, 5) & "' ", db, adOpenKeyset, adLockOptimistic
        'rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_proyecto='" & v_EstPoa(i, 8) & "' and par_codigo='" & v_EstPoa(i, 3) & "' and org_codigo= '" & v_EstPoa(i, 5) & "'", db, adOpenKeyset, adLockOptimistic
'        Print rstfo_formulacion_gasto.RecordCount
        If Not (rstfo_formulacion_gasto.EOF) Then
          If (rstfo_formulacion_gasto!FGS_VIGENTE - rstfo_formulacion_gasto!FGS_compromiso < rstao_solicitud_detalle!monto_Bolivianos) Then  'adoorigen         'adoorigen.adosolicitud.Recordset!monto_dolares ) Then
'            MsgBox "NO EXISTE Presupuesto para dar curso a la Solicitud ...", vbOKOnly, "ERROR"
'            verifica_ppto = 0 'swpresup = 0
'            Exit Function
            verifica_ppto = 1
          Else
            ' actualizar ahora??? consaultar con jorge
            'rstfo_formulacion_gasto!fgs_compromiso  = rstfo_formulacion_gasto!fgs_compromiso  + rstao_solicitud_m!monto_dolares
            'rstfo_formulacion_gasto.Update
            verifica_ppto = 1 'swpresup = 1
          End If
          If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
        Else
          MsgBox "NO EXISTE Estructura presupuestaria...", vbOKOnly, "ERROR ..."
          verifica_ppto = 0 'swpresup = 0
          Exit Function
        End If
      Else
        MsgBox "Noooo existe poa", vbOKOnly, "ERROR ..."
        verifica_ppto = 0 'swpresup = 0
        Exit Function
      End If
'          Else
'            verifica_ppto = 1 'swpresup = 1
'          End If
      rstao_solicitud_detalle.MoveNext
    Next
  End If
End Function

Public Function verifica_porcen(codigo_poa1, que)
'---- que determina que columna de porcentaje se debe utilizar ----
  Dim rstfc_relacionador_poa_ppto As New ADODB.Recordset
  Dim rstfc_porcentaje_fte As New ADODB.Recordset
  Dim Par_Codigo1 As String
  Dim fte_codigo1 As String
'  Dim Org_Codigo1 As String
  Dim codigo_convenio1 As String
  Dim Categoria1 As String
  Dim Pro_Proyecto1 As String
  Dim swsalir As Integer
  swsalir = 0
  Set rstfc_relacionador_poa_ppto = New ADODB.Recordset
  If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
  rstfc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & codigo_poa1 & "'", db, adOpenKeyset, adLockReadOnly
  If rstfc_relacionador_poa_ppto.RecordCount > 0 Then
    Par_Codigo1 = rstfc_relacionador_poa_ppto!par_codigo
    fte_codigo1 = rstfc_relacionador_poa_ppto!fte_codigo
    'fte_codigo1 = fBuscaFte(rstfc_relacionador_poa_ppto!org_codigo)
' aquiiiiiiiiiiiiiiiiiiiiii
    Org_Codigo1 = rstfc_relacionador_poa_ppto!org_codigo
    Pro_Proyecto1 = rstfc_relacionador_poa_ppto!pro_proyecto
    codigo_convenio1 = rstfc_relacionador_poa_ppto!codigo_convenio
    Categoria1 = IIf(IsNull(rstfc_relacionador_poa_ppto!codigo_categoria), "", rstfc_relacionador_poa_ppto!codigo_categoria)
    por_fte_ext1 = rstfc_relacionador_poa_ppto!por_ext
  Else
    MsgBox "No se pudo encontrar la estructura presupuestaria" & vbCrLf & _
    vbTab & "para este código POA.", vbOKOnly + vbExclamation, "Error en el relacionador Poa - Presupuesto..."
    swsalir = 1
  End If
  If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
  
'  If por_fte_ext1 = 100 Then
'    que = 1
'  Else
'    que = 2
'  End If
  
  If swsalir = 0 Then
    Set rstfc_porcentaje_fte = New ADODB.Recordset
    '''ALB AQUI
    If rstfc_porcentaje_fte.State = 1 Then rstfc_porcentaje_fte.Close
'    rstfc_porcentaje_fte.Open "select * from fc_porcentaje_fte where Par_Codigo = '" & Par_Codigo1 & "' and codigo_convenio = '" & codigo_convenio1 & "and pro_proyecto = '" & pro_proyecto1 & "' ", db, adOpenKeyset, adLockReadOnly
    rstfc_porcentaje_fte.Open "select * from so_porcentaje_convenio where Par_Codigo = '" & Par_Codigo1 & "' and codigo_convenio = '" & codigo_convenio1 & "' and pro_proyecto = '" & Pro_Proyecto1 & "' ", db, adOpenKeyset, adLockReadOnly
    If rstfc_porcentaje_fte.RecordCount > 0 Then
      If Par_Codigo1 = "39600" Then
        verifica_porcen = 396
      Else
        If que = 1 Then
          verifica_porcen = rstfc_porcentaje_fte!prc_porcentaje
        End If
        If que = 2 Then
          verifica_porcen = rstfc_porcentaje_fte!prc_porcentaje_aux
        End If
        'If verifica_porcen = 0 Then MsgBox "El porcentaje encontrado es igual a cero (0)", vbOKOnly + vbInformation, "Error en el Clasificador..."
'        verifica_porcen = rstfc_porcentaje_fte!por_fte_ext
      End If
    Else
      MsgBox "No se pudo determinar los porcentajes " & vbCrLf & _
      "para las fuentes de financiamiento." & vbCrLf & " ERROR en el CLASIFICADOR.", vbOKOnly + vbExclamation, "Error en el relacionador de porcentajes..."
    End If
    If rstfc_porcentaje_fte.State = 1 Then rstfc_porcentaje_fte.Close
  End If
End Function

Public Function fbusrelapoa_ppto(codigo_poa1, que)
  Dim rsfc_relacionador_poa_ppto As New ADODB.Recordset
  Set rsfc_relacionador_poa_ppto = New ADODB.Recordset
  If rsfc_relacionador_poa_ppto.State = 1 Then rsfc_relacionador_poa_ppto.Close
  rsfc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & codigo_poa1 & "'", db, adOpenKeyset, adLockReadOnly
  If rsfc_relacionador_poa_ppto.RecordCount > 0 Then
    Select Case que
      Case 1
        fbusrelapoa_ppto = rsfc_relacionador_poa_ppto!org_codigo
      Case 2
      
    End Select
  Else
    fbusrelapoa_ppto = "Err"
  End If
End Function


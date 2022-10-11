Attribute VB_Name = "generales"
Option Explicit

Public Sub prt_cmbteppto(ges, org, cod)
' OOOOJJJJJJOOOOOO
' ACTIVAR ESTE SUB PARA IMPRIMIR CMPBTE PRESUPUESTARIO DE GASTO

'    Dim Report As New CrtComprobante
Dim rsRepDet As New ADODB.Recordset
Dim rscorrelativo As New ADODB.Recordset
Dim rsRepCab As New ADODB.Recordset
Dim rsPagos As New ADODB.Recordset
Set rsPagos = New ADODB.Recordset
If rsPagos.State = 1 Then rsPagos.Close
Dim regularizacion1 As String

rsPagos.Open "select * from pagos where ges_gestion = '" & ges & "' and org_codigo = '" & org & "' and codigo_pago = " & cod, db, adOpenKeyset, adLockReadOnly
If rsPagos.RecordCount < 1 Then
  MsgBox "Error al generar la Impresión de Comprobante presupuestario", vbCritical + vbOKOnly, "Error al Imprimir"
  If rsPagos.State = 1 Then rsPagos.Close
  Exit Sub
End If
On Error GoTo error_GRABAR:
Screen.MousePointer = vbHourglass
        Set rscorrelativo = New ADODB.Recordset
        Set rsRepCab = New ADODB.Recordset

        db.Execute "DELETE from pagos_rep"
        If rsRepCab.State = 1 Then rsRepCab.Close
        rsRepCab.Open "select * from pagos_rep ", db, adOpenKeyset, adLockOptimistic
        rsRepCab.AddNew
        rsRepCab("Maquina") = GlMaquina
        rsRepCab("codigo_pago") = rsPagos!codigo_pago 'TxtComprobante.Text
        rsRepCab("ges_gestion") = rsPagos!Ges_gestion 'Year(Now)
        rsRepCab("org_codigo") = rsPagos!org_codigo 'DtCOrg.Text
        rsRepCab("codigo_unidad") = rsPagos!Codigo_unidad  'DtCUnidad.Text
        rsRepCab("nro_comprobante_anterior") = rsPagos!nro_comprobante_anterior 'TxtComprobanteAnterior.Text
        rsRepCab("codigo_documento") = IIf(IsNull(rsPagos!codigo_documento), "", rsPagos!codigo_documento) 'DtcDcu.Text
        rsRepCab("codigo_orden") = IIf(IsNull(rsPagos!Codigo_orden), "", rsPagos!Codigo_orden) 'TxtCodigoOrden.Text
        rsRepCab("codigo_solicitud") = rsPagos!codigo_solicitud 'Trim(txtnrosolicitud.Text)
        rsRepCab("tipo_formulario") = rsPagos!tipo_formulario 'DtcTipoCod.Text

        'rsRepCab("fecha_egreso") = CDate(rsPagos!fecha_egreso) 'CDate(dtpFecha)
        rsRepCab("fecha_egreso") = CDate(rsPagos!fecha_registro)

        rsRepCab("uni_codigo") = "CENTRAL"
        rsRepCab("fte_codigo") = rsPagos!fte_codigo 'DTcFte.Text
        rsRepCab("codigo_convenio") = rsPagos!codigo_convenio 'JQA 23/11/01
        rsRepCab("codigo_categoria") = rsPagos!codigo_Categoria 'DtcCat.Text
        rsRepCab("justificacion") = rsPagos!justificacion 'TxtJustificacion.Text
        rsRepCab("tipo_moneda") = "Bs." 'DtCTipoMoneda.Text
        rsRepCab("monto_bolivianos") = IIf(IsNull(rsPagos("monto_bolivianos")), 0, Round(rsPagos("monto_bolivianos"), 2))
        rsRepCab("monto_dolares") = IIf(IsNull(rsPagos("monto_dolares")), 0, Round(rsPagos("monto_dolares"), 2))
        rsRepCab("Deducciones") = IIf(IsNull(rsPagos("Deducciones")), 1, rsPagos("Deducciones"))
        rsRepCab("liquido_pagar") = rsPagos("liquido_pagar")
        LiteralCry = Str(Round(rsRepCab("monto_bolivianos"), 2))

        'JQA 23/11/01
        rsRepCab("estado_compromiso") = rsPagos("estado_compromiso")
        rsRepCab("estado_devengado") = rsPagos("estado_devengado")
        rsRepCab("estado_pagado") = rsPagos("estado_pagado")
        rsRepCab("estado_devolucion") = rsPagos("estado_devolucion")
        rsRepCab("estado_anulado") = rsPagos("estado_anulado")
        rsRepCab("estado_reversion_total") = rsPagos("estado_reversion_total")
        rsRepCab("estado_reversion_parcial") = rsPagos("estado_reversion_parcial")
        'JQA 23/11/01

       rsRepCab.Update
'       LblTitulo.Caption = "." ACUI

       Set rsRepDet = New ADODB.Recordset
       db.Execute "DELETE from pago_detalle_rep"

       If rsRepDet.State = 1 Then rsRepDet.Close
       rsRepDet.Open "select * from pago_detalle_rep ", db, adOpenKeyset, adLockOptimistic
       'rsRepDet.Open "select * from pago_detalle_rep ", db, adOpenKeyset, adLockOptimistic
'       While Not rsRepDet.EOF And rsRepDet.RecordCount > 0
'                     rsRepDet.Delete
'                     rsRepDet.MoveNext
'       Wend

       If Not IsNull(rsPagos("codigo_pago")) And Not IsNull(rsPagos("org_codigo")) Then
           Set rsdetalle = New ADODB.Recordset
           rsdetalle.Open "select * from pago_detalle where codigo_pago='" & rsPagos("codigo_pago") & "' and org_codigo='" & rsPagos("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
      'revvvvvvvvvv ---------
        If rsPagos("estado_devengado") = "S" Then
      '           rspagos("estado_aprobacion") = "N"
      '           rspagos.Update
              End If
        '''RepComprobante.Show vbModal

        If rsdetalle("monto_total") <> 0 Then
      '-----------
           'Set DtGDetalle.DataSource = rsDetalle
          If rsdetalle.RecordCount > 0 Then
          While Not rsdetalle.EOF
            rsRepDet.AddNew
            rsRepDet("Maquina") = GlMaquina
            rsRepDet("codigo_pago") = rsdetalle("codigo_pago")
            rsRepDet("ges_gestion") = rsdetalle("ges_gestion")
            rsRepDet("org_codigo") = rsdetalle("org_codigo")
            rsRepDet("codigo_pago_detalle") = rsdetalle("codigo_pago_detalle")
            rsRepDet("par_codigo") = rsdetalle("par_codigo")
            rsRepDet("pro_programa") = rsdetalle("pro_programa")
'            rsRepDet("Pro_subprograma") = rsdetalle("Pro_subprograma")
            rsRepDet("Pro_proyecto") = rsdetalle("Pro_proyecto")
            rsRepDet("Pro_actividad") = rsdetalle("Pro_actividad")
            If rsdetalle("codigo_beneficiario") <> "" Then
                rsRepDet("codigo_beneficiario") = rsdetalle("codigo_beneficiario")
              Else
                rsRepDet("codigo_beneficiario") = "-"
            End If
            rsRepDet("monto_total") = Round(rsdetalle("monto_total"), 2)
            If rsPagos("monto_bolivianos") = 0 Then
              LiteralCry = Str(Round(rsdetalle("monto_total"), 2))
            End If
            rsRepDet("monto_dolares") = Round(rsdetalle("monto_dolares"), 2)
            rsRepDet("tipo_cambio") = rsdetalle("tipo_cambio")
            rsRepDet("codigo_poa") = rsdetalle("codigo_poa")            'JQA 23/11/01
            rsRepDet("Deducciones") = rsdetalle("Deducciones")
            rsRepDet("saldo_bolivianos") = rsdetalle("saldo_bolivianos")
            rsRepDet("literal") = Literal(LiteralCry) + "  Bolivianos"
            rsRepDet.Update
            rsdetalle.MoveNext

          Wend
          End If
        End If

        'Report.lblFecha.SetText (Literal(LiteralCry) + "  Bolivianos")
'        RepComprobante.CRViewer1.ReportSource = Report
'        RepComprobante.CRViewer1.ViewReport

' JORGE
'    MsgBox "EL USUARIO NO TIENE ACCESO AL SERVIDOR SQL ..."
'    Exit Sub
        'RepComprobante.Show vbModal
        Dim iResult As Integer
        ff_egresos.Cry.Reset 'COMENTAR ?
        ff_egresos.Cry.WindowState = crptMaximized
        ff_egresos.Cry.WindowShowPrintSetupBtn = True
        ff_egresos.Cry.WindowShowPrintBtn = True
        ff_egresos.Cry.WindowShowRefreshBtn = True
        ff_egresos.Cry.ReportFileName = App.Path & "\FormsPresupuesto\Diseñadores\CrtComprobantePpto.rpt"
        ff_egresos.Cry.SelectionFormula = "{fv_comprobante2.Maquina} = '" & GlMaquina & "'" 'COMENTAR?
        iResult = ff_egresos.Cry.PrintReport
        If iResult <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox ff_egresos.Cry.LastErrorNumber & " : " & ff_egresos.Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
        End If
' JORGE
       Else
        MsgBox "No se registró el detalle del comprobante ..."
       End If
  Screen.MousePointer = vbDefault
  If rsPagos.State = 1 Then rsPagos.Close
Exit Sub
error_GRABAR:
  Screen.MousePointer = vbDefault
MsgBox Err.Number & " " & Err.Description

End Sub

Public Sub prt_cmbteIng(ges, org, cod)
' OOOOJJJJJJOOOOOO
' ACTIVAR ESTE SUB PARA IMPRIMIR CMPBTE PRESUPUESTARIO DE INGRESO

  Dim rsfo_ingresos As New ADODB.Recordset
  Dim rstfo_ingresos_rep As New ADODB.Recordset
  Set rstfo_ingresos_rep = New ADODB.Recordset
  Dim iResult As Integer

  Set rsfo_ingresos = New ADODB.Recordset
  If rsfo_ingresos.State = 1 Then rsfo_ingresos.Close
  rsfo_ingresos.Open "select * from fo_ingresos where ges_gestion = '" & ges & "' and org_codigo = '" & org & "' and Correlativo_ingreso = " & cod, db, adOpenKeyset, adLockReadOnly
  If rsfo_ingresos.RecordCount < 1 Then
    MsgBox "Error al generar la Impresión de Comprobante de Ingreso", vbCritical + vbOKOnly, "Error al Imprimir..."
    If rsfo_ingresos.State = 1 Then rsfo_ingresos.Close
    Exit Sub
  End If
  '  Cry.Reset
  FrmIngresosabm.Cry.ReportFileName = App.Path & "\Reportes\Ingresos\ComprobIngreso.rpt"
'  Cry.SelectionFormula = "{fv_comprobante2.Maquina} = '" & GlMaquina & "'"
  If rstfo_ingresos_rep.State = 1 Then rstfo_ingresos_rep.Close
  rstfo_ingresos_rep.Open "select * from fo_ingresos_rep where maquina = '" & GlMaquina & "'", db, adOpenKeyset, adLockOptimistic
  While Not (rstfo_ingresos_rep.EOF)
    rstfo_ingresos_rep.Delete
    rstfo_ingresos_rep.MoveNext
  Wend
  '====== ini cargado de la tabla aux para impresion ====
  rstfo_ingresos_rep.AddNew
  rstfo_ingresos_rep("Correlativo_ingreso") = rsfo_ingresos!correlativo_ingreso 'LblCorrelativo_ingreso.Caption
  rstfo_ingresos_rep("Correlativo_anterior") = rsfo_ingresos!correlativo_anterior
  rstfo_ingresos_rep("Ges_Gestion") = rsfo_ingresos!Ges_gestion 'Trim(lblges_gestion.Caption) ' TxtGes_Gestion.Text
  rstfo_ingresos_rep("Codigo_solicitud") = rsfo_ingresos!codigo_solicitud 'txtCodigo_solicitud.Text
  rstfo_ingresos_rep("Codigo_unidad") = rsfo_ingresos!Codigo_unidad 'Unidad Ejecutora
  rstfo_ingresos_rep("rbr_codigo") = rsfo_ingresos!rbr_codigo 'DtCrbr_codigo.Text
  rstfo_ingresos_rep("tipo_moneda") = rsfo_ingresos!tipo_moneda 'DtCdenominacion_moneda.BoundText
  rstfo_ingresos_rep("Codigo_tipo") = rsfo_ingresos!Codigo_tipo 'DtCDenominacion_tipo.BoundText
  rstfo_ingresos_rep("Codigo_tipo_solicitud") = rsfo_ingresos!Codigo_tipo_solicitud 'IIf(DtCDenominacion_tipo_solicitud.BoundText = "", 0, DtCDenominacion_tipo_solicitud.BoundText)
  rstfo_ingresos_rep("Codigo_documento") = rsfo_ingresos!codigo_documento 'DtCCodigo_documento.Text
  rstfo_ingresos_rep("Fecha_Ingreso") = CDate(rsfo_ingresos!Fecha_Ingreso) 'DTPFecha_Ingreso.Value
  rstfo_ingresos_rep("Tipo_Cambio") = rsfo_ingresos!tipo_cambio 'txtTipo_Cambio.Text
  rstfo_ingresos_rep("Concepto") = rsfo_ingresos!Concepto 'txtConcepto.Text
  rstfo_ingresos_rep("UNI_CODIGO") = rsfo_ingresos!uni_codigo 'Txtuni_codigo
  rstfo_ingresos_rep("fte_codigo") = rsfo_ingresos!fte_codigo 'DtCFte_codigo.Text
  rstfo_ingresos_rep("org_codigo") = rsfo_ingresos!org_codigo 'DtCOrg_codigo.Text
  rstfo_ingresos_rep("codigo_convenio") = rsfo_ingresos!codigo_convenio 'convenio
  rstfo_ingresos_rep("formulario") = rsfo_ingresos!formulario 'formulario
  rstfo_ingresos_rep("TIPO_Comp") = rsfo_ingresos!Tipo_Comp 'Tipo de Comprobante
  'If rsfo_ingresos("Codigo_tipo") = "DEV" Then
    rstfo_ingresos_rep("codigo_beneficiario") = rsfo_ingresos!codigo_beneficiario 'DtCcodigo_beneficiario.Text
  'End If

  If rsfo_ingresos("Codigo_tipo") = "DYR" Or rsfo_ingresos("Codigo_tipo") = "DVI" Or rsfo_ingresos("Codigo_tipo") = "REC" Or rsfo_ingresos("Codigo_tipo") = "ANI" Then
    rstfo_ingresos_rep("Cta_codigo") = rsfo_ingresos!Cta_codigo 'DtCCta_codigo.Text
  End If

'  rstfo_ingresos_rep("cta_codigo") = DtCCta_codigo.Text

  rstfo_ingresos_rep("numero_documento") = rsfo_ingresos!numero_documento 'txtNumero_documento.Text
  rstfo_ingresos_rep("monto_dolares") = Round(rsfo_ingresos!monto_dolares, 2) 'Round(Txtmonto_dolares.Text, 2)
  rstfo_ingresos_rep("monto_bolivianos") = Round(rsfo_ingresos!monto_Bolivianos, 2) 'Round(txtMonto_bolivianos.Text, 2)
  rstfo_ingresos_rep("usr_usuario") = GlUsuario
  rstfo_ingresos_rep("fecha_registro") = CDate(rsfo_ingresos!fecha_registro) 'Date
  rstfo_ingresos_rep("hora_registro") = Left(CStr(Time()), 8)
  rstfo_ingresos_rep("estado_recaudado") = IIf(rsfo_ingresos("estado_recaudado") = "S", "S", IIf(rsfo_ingresos("estado_recaudado") = "N", "S", IIf(IsNull(rsfo_ingresos!estado_recaudado), "", rsfo_ingresos!estado_recaudado)))
  rstfo_ingresos_rep("estado_devengado") = IIf(rsfo_ingresos("estado_devengado") = "S", "S", IIf(rsfo_ingresos("estado_devengado") = "N", "S", IIf(IsNull(rsfo_ingresos!estado_devengado), "", rsfo_ingresos!estado_devengado)))
  rstfo_ingresos_rep("estado_desafectado") = IIf(rsfo_ingresos("estado_desafectado") = "S", "S", IIf(rsfo_ingresos("estado_desafectado") = "N", "S", IIf(IsNull(rsfo_ingresos!estado_desafectado), "", rsfo_ingresos!estado_desafectado)))
  rstfo_ingresos_rep("estado_anulado") = IIf(rsfo_ingresos("estado_anulado") = "S", "S", IIf(rsfo_ingresos("estado_anulado") = "N", "S", IIf(IsNull(rsfo_ingresos!estado_anulado), "", rsfo_ingresos!estado_anulado)))
  rstfo_ingresos_rep("estado_aprobacion") = rsfo_ingresos("estado_aprobacion")
  LiteralCry = Str(Round(rstfo_ingresos_rep("monto_bolivianos"), 2))
  rstfo_ingresos_rep("literal") = Literal(LiteralCry) + "  Bolivianos"
  rstfo_ingresos_rep("maquina") = GlMaquina
  rstfo_ingresos_rep.Update
  If rstfo_ingresos_rep.State = 1 Then rstfo_ingresos_rep.Close
  '====== fin cargado de la tabla aux para impresion ====

  FrmIngresosabm.Cry.SelectionFormula = "{Vi_Fo_ingresos_rep.Maquina} = '" & GlMaquina & "'"
  FrmIngresosabm.Cry.WindowShowPrintBtn = True
  FrmIngresosabm.Cry.WindowShowExportBtn = True
  FrmIngresosabm.Cry.WindowShowPrintSetupBtn = True
  FrmIngresosabm.Cry.WindowState = crptMaximized
  iResult = FrmIngresosabm.Cry.PrintReport
  If iResult <> 0 Then
      MsgBox FrmIngresosabm.Cry.LastErrorNumber & " : " & FrmIngresosabm.Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

End Sub

Public Function VerPptoConvenio(Convenio, Categoria, org, cod)
'  swVerPptoConvenio = 1
  ' ==== INI CONTROL POR CONVENIO ====
  Dim rstacum As ADODB.Recordset
  Set rstacum = New ADODB.Recordset

  Dim rsfc_categoria_financiador As New ADODB.Recordset
  Set rsfc_categoria_financiador = New ADODB.Recordset
  If rsfc_categoria_financiador.State = 1 Then rsfc_categoria_financiador.Close
  rsfc_categoria_financiador.Open "select SUM(monto_vigente_us) AS acumconvig , SUM(monto_compromiso_us) AS acumconcom from fc_categoria_financiador where codigo_convenio = '" & Convenio & "' ", db, adOpenKeyset, adLockReadOnly
  If rsfc_categoria_financiador.RecordCount > 0 Then
    If rstacum.State = 1 Then rstacum.Close
    rstacum.Open "select sum (monto_dolares) as acumdl from pago_detalle where org_codigo = '" & org & "' and codigo_pago = " & cod, db, adOpenStatic, adLockReadOnly
    If (rsfc_categoria_financiador!acumconvig - rsfc_categoria_financiador!acumconcom) >= rstacum!acumDl Then
      'swVerPptoConvenio = 1
      VerPptoConvenio = 1
    Else
      'swVerPptoConvenio = 0
      VerPptoConvenio = 0
      MsgBox "¡¡ NO EXISTE PRESUPUESTO !!" & vbCrLf & vbCrLf & "Convenio : " & Convenio & vbCrLf & _
      vbCrLf & vbCrLf & " Monto Vigente        = " & rsfc_categoria_financiador!acumconvig & vbCrLf & "Total Comprometido = " & rsfc_categoria_financiador!acumconcom & vbCrLf & " Monto Solicitado     = " & rstacum!acumDl, vbCritical + vbOKOnly, "Error en montos"
    End If
    If rstacum.State = 1 Then rstacum.Close
  Else
    'swVerPptoConvenio = 0
    VerPptoConvenio = 0
    MsgBox "Error al buscar la categoriía para el convenio", vbCritical + vbOKOnly, "Error de datos"
  End If
  If rsfc_categoria_financiador.State = 1 Then rsfc_categoria_financiador.Close
  ' ==== FIN CONTROL POR CONVENIO ====


' ==== INI CONTROL POR CATEGORIA ====
'  Dim rstacum As ADODB.Recordset
'  Set rstacum = New ADODB.Recordset
'
'  Dim rsfc_categoria_financiador As New ADODB.Recordset
'  Set rsfc_categoria_financiador = New ADODB.Recordset
'  If rsfc_categoria_financiador.State = 1 Then rsfc_categoria_financiador.Close
'  rsfc_categoria_financiador.Open "select * from fc_categoria_financiador where codigo_convenio = '" & Convenio & "' and codigo_categoria = '" & Categoria & "' ", db, adOpenKeyset, adLockReadOnly
'  If rsfc_categoria_financiador.RecordCount > 0 Then
'    If rstacum.State = 1 Then rstacum.Close
'    rstacum.Open "select sum (monto_dolares) as acumdl from pago_detalle where org_codigo = '" & AdoRegularizacion.Recordset!org_codigo & "' and codigo_pago = " & AdoRegularizacion.Recordset!codigo_pago, db, adOpenStatic, adLockReadOnly
'    If (rsfc_categoria_financiador!monto_vigente_us - rsfc_categoria_financiador!monto_compromiso_us) >= rstacum!acumdl Then
'      swVerPptoConvenio = 1
'    Else
'      swVerPptoConvenio = 0
'      MsgBox "¡¡ NO EXISTE PRESUPUESTO !!" & vbCrLf & vbCrLf & "Convenio : " & AdoRegularizacion.Recordset!codigo_convenio & vbCrLf & "Categoria : " & AdoRegularizacion.Recordset!codigo_categoria & _
'      vbCrLf & vbCrLf & " Monto Vigente        = " & rsfc_categoria_financiador!monto_vigente_us & vbCrLf & "Total Comprometido = " & rsfc_categoria_financiador!monto_compromiso_us & vbCrLf & " Monto Solicitado     = " & rstacum!acumdl, vbCritical + vbOKOnly, "Error en montos"
'    End If
'    If rstacum.State = 1 Then rstacum.Close
'  Else
'    swVerPptoConvenio = 0
'    MsgBox "Error al buscar la categoriía para el convenio", vbCritical + vbOKOnly, "Error de datos"
'  End If
'  If rsfc_categoria_financiador.State = 1 Then rsfc_categoria_financiador.Close
' ==== FIN CONTROL POR CATEGORIA ====

End Function

Public Function verppto_nal(Poa)
''Verifica ppto
'  Dim pro As String
'  Dim pry As String
'  Dim act As String
'  Dim par As String
'  Dim org As String
'
'  verppto_nal = 0
'  Dim rsfc_relacionador_Poa_ppto As New ADODB.Recordset
'  Set rsfc_relacionador_Poa_ppto = New ADODB.Recordset
'  If rsfc_relacionador_Poa_ppto.State = 1 Then rsfc_relacionador_Poa_ppto.Close
'  rsfc_relacionador_Poa_ppto.Open "select * from fc_relacionador_Poa_ppto where codigo_poa = '" & Poa & "' ", db, adOpenKeyset, adLockOptimistic
'  If rsfc_relacionador_Poa_ppto.RecordCount > 0 Then
'    rsfc_relacionador_Poa_ppto
'  Else
'    MsgBox "El POA: " & Poa & " no se encuentra en el relacionador POA Presupuesto", vbCritical + vbOKOnly, "Error al aprobar..."
'    verppto_nal = 0
'  End If
'
'  Set RsDet = New ADODB.Recordset
'  If RsDet.State = 1 Then RsDet.Close
'  RsDet.Open "select * from pago_detalle where codigo_pago= " & AdoRegularizacion.Recordset!codigo_pago & " and org_codigo= '" & AdoRegularizacion.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'  '  Print rsDet.RecordCount
'  If RsDet.RecordCount > 0 Then
'    ppto2 = "0"
'    Set rsPpto = New ADODB.Recordset
'    If rsPpto.State = 1 Then rsPpto.Close
'    rsPpto.Open "select * from fo_formulacion_gasto where pro_programa='" & RsDet("pro_programa") & "' and pro_proyecto='" & RsDet("pro_proyecto") & "' and pro_actividad='" & RsDet("pro_actividad") & "' and par_codigo='" & RsDet("par_codigo") & "' and org_codigo= '" & RsDet("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'    If rsPpto.RecordCount > 0 Then
'       ppto2 = "1"
'       If AdoRegularizacion.Recordset("estado_compromiso") = "N" Then
'        If ((IIf(IsNull(rsPpto("FGS_vigente")), 0, rsPpto("FGS_vigente")) - IIf(IsNull(rsPpto("FGS_compromiso")), 0, rsPpto("FGS_compromiso")) + IIf(IsNull(rsPpto("FGS_acum_rev")), 0, rsPpto("FGS_acum_rev")) + IIf(IsNull(rsPpto("FGS_acum_dev")), 0, rsPpto("FGS_acum_dev"))) < RsDet("monto_total")) Then
'          If AdoRegularizacion.Recordset!fte_codigo = "41" Then
'            MsgBox "NO EXISTE PRESUPUESTO PARA COMPROMETER ...", vbOKOnly, "ERROR"
'            '----ini se desabilita el control solo por un tiempo SOLICITUD IMAÑA
'            Exit Function 'g-
'          Else
'            rsPpto("fgs_compromiso") = IIf(IsNull(rsPpto("fgs_compromiso")), 0, rsPpto("fgs_compromiso")) + RsDet("monto_total") 'g-
'            rsPpto.Update 'g-
'            '----fin se desabilita el control solo por un tiempo SOLICITUD IMAÑA
'          End If
'        Else
'          rsPpto("fgs_compromiso") = rsPpto("fgs_compromiso") + RsDet("monto_total")
'          rsPpto.Update
'        End If
'       End If
'       If AdoRegularizacion.Recordset("estado_devengado") = "N" Then
'        ' Para Validar lo Devengado
'        ' Modificado por Gerardo Rodriguez
'        Dim RsDevenga As ADODB.Recordset
'        Dim RsCompro As ADODB.Recordset
'        Dim GlSqlAux As String
'        Set RsDevenga = New ADODB.Recordset
'        Set RsCompro = New ADODB.Recordset
'        ' Para ACCESS
'        'GlSQLAux = "SELECT IIF(ISNULL(SUM(monto_Total)), 0, SUM(monto_Total)) AS TotalDevengado " & _
'        '           "FROM pagos, pago_Detalle " & _
'        '           "WHERE (pagos.codigo_pago = pago_detalle.codigo_pago) AND (pagos.Tipo_formulario = 'DEV') AND (pagos.estado_devengado = 'S') AND (pagos.Nro_Comprobante_Anterior = '" & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & "')"
'        ' Para SQL
'        '      GlSqlAux = "SELECT ISNULL(SUM(monto_Total), 0) AS TotalDevengado " & _
'        '                 "FROM pagos, pago_Detalle " & _
'        '                 "WHERE (pagos.codigo_pago = pago_detalle.codigo_pago) AND (pagos.Tipo_formulario = 'DEV') AND (pagos.estado_devengado = 'S') AND (pagos.Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & ") AND (pagos.org_codigo = '" & AdoRegularizacion.Recordset!Org_Codigo & "')"
'        'corregido por jorge . . .
'
'        GlSqlAux = "SELECT ISNULL(SUM(monto_bolivianos), 0) AS TotalDevengado " & _
'                   "FROM pagos " & _
'                   "WHERE (pagos.Tipo_formulario = 'DEV') AND (pagos.estado_devengado = 'S') AND (pagos.Nro_Comprobante_Anterior = " & AdoRegularizacion.Recordset!nro_comprobante_anterior & ") AND (pagos.org_codigo = '" & AdoRegularizacion.Recordset!org_codigo & "')"
'
'        RsDevenga.Open GlSqlAux, db, adOpenStatic
'
'        Dim rstcom As New ADODB.Recordset
'        Set rstcom = New ADODB.Recordset
'        If rstcom.State = 1 Then rstcom.Close
'        rstcom.Open "select * from pagos where org_codigo = '" & AdoRegularizacion.Recordset!org_codigo & "' and  nro_comprobante_anterior = " & AdoRegularizacion.Recordset!nro_comprobante_anterior & " and (tipo_formulario = 'COM' or tipo_formulario = 'COA')", db, adOpenKeyset, adLockReadOnly
'        While Not rstcom.EOF
'          GlSqlAux = "SELECT Sum(Monto_Total) AS MontoTotal FROM pago_detalle " & _
'                   "WHERE (pago_detalle.Codigo_Pago = " & rstcom!codigo_pago & ") AND (pago_detalle.org_codigo = '" & rstcom!org_codigo & "') "
'          If RsCompro.State = 1 Then RsCompro.Close
'          RsCompro.Open GlSqlAux, db, adOpenStatic
'          varcom = varcom + IIf(IsNull(RsCompro!MontoTotal), 0, RsCompro!MontoTotal)
'  '          varcom = varcom + rstcom!Monto_Total
'            rstcom.MoveNext
'        Wend
'  '      Print rstcom.RecordCount
'        If rstcom.State = 1 Then rstcom.Close
'
'        If AdoRegularizacion.Recordset!tipo_formulario = "CYD" Then
'          GlSqlAux = "SELECT Sum(Monto_Total) AS MontoTotal FROM pago_detalle " & _
'                     "WHERE (pago_detalle.Codigo_Pago = " & AdoRegularizacion.Recordset!nro_comprobante_anterior & ") AND (pago_detalle.org_codigo = '" & AdoRegularizacion.Recordset!org_codigo & "') "
'          If RsCompro.State = 1 Then RsCompro.Close
'          RsCompro.Open GlSqlAux, db, adOpenStatic
'          varcom = RsCompro!MontoTotal
'        End If
'
'        If (varcom < RsDevenga!TotalDevengado + RsDet("monto_total")) Then
'          MsgBox "La Suma de lo DEVENGADO excede el Monto del Compromiso del Comprobante '" & AdoRegularizacion.Recordset!nro_comprobante_anterior & "'.", vbExclamation + vbOKOnly, "ERROR" '"La estructura presupuestaria NO es válida o NO EXISTE PRESUPUESTO "
'          Exit Function
'        Else
'          rsPpto("fgs_devengado") = rsPpto("fgs_devengado") + RsDet("monto_total")
'          rsPpto.Update
'        End If
'
'  'ini antes
'  '      GlSqlAux = "SELECT Sum(Monto_Total) AS MontoTotal FROM pago_detalle " & _
'  '                 "WHERE (pago_detalle.Codigo_Pago = " & AdoRegularizacion.Recordset!nro_comprobante_anterior & ") AND (pago_detalle.org_codigo = '" & AdoRegularizacion.Recordset!org_codigo & "') "
'  '      RsCompro.Open GlSqlAux, db, adOpenStatic
'  '      If (RsCompro!MontoTotal < RsDevenga!TotalDevengado + RsDet("monto_total")) Then
'  '        MsgBox "La Suma de lo DEVENGADO excede el Monto del Compromiso del Comprobante '" & AdoRegularizacion.Recordset!nro_comprobante_anterior & "'.", vbExclamation + vbOKOnly, "ERROR" '"La estructura presupuestaria NO es válida o NO EXISTE PRESUPUESTO "
'  '        Exit Sub
'  '      Else
'  '        rsPpto("fgs_devengado") = rsPpto("fgs_devengado") + RsDet("monto_total")
'  '        rsPpto.Update
'  '      End If
'  'fin antes
'
'
'      '      If (rsPpto("FGS_compromiso") - rsPpto("FGS_devengado") < rsDet("monto_total")) Then
'      '        MsgBox "NO EXISTE PRESUPUESTO PARA DEVENGAR ", vbOKOnly, "ERROR"  '"La estructura presupuestaria NO es válida o NO EXISTE PRESUPUESTO "
'      '        Exit Sub
'      '      Else
'      '        rsPpto("fgs_devengado") = rsPpto("fgs_devengado") + rsDet("monto_total")
'      '        rsPpto.Update
'      '      End If
'       End If
'       'Verificar por que ...
'       If AdoRegularizacion.Recordset("estado_pagado") = "N" Then
'          If (rsPpto("FGS_compromiso") - rsPpto("FGS_pagado") < RsDet("monto_total")) Then
'             MsgBox "NO EXISTE PRESUPUESTO", vbOKOnly, "ERROR"  '"La estructura presupuestaria NO es válida o NO EXISTE PRESUPUESTO "
'             Exit Function
'          Else
'             rsPpto("fgs_pagado") = IIf(IsNull(rsPpto("fgs_pagado")), 0, rsPpto("fgs_pagado")) + RsDet("monto_total")
'             rsPpto.Update
'          End If
'       End If
'       'Verificar por que ... hasta aqui ...
'     Else
'      'g- MsgBox "La estructura presupuestaria NO es válida", vbOKOnly, "ERROR"
'      'g- Exit Sub
'    End If
'    If rsPpto.State = 1 Then rsPpto.Close
'    '************************

End Function

Public Function GeneraDPD(ges, org, cod)
  Dim rs
'  db.Execute " exec pdInsPagoDirecto_DPD " &
'  ges_gestion, org_codigo,
'  Tipo_Cambio,
'@Rbr_Codigo int,
'@TipoMoneda varchar(10),
'@Codigo_Beneficiario varchar(15),
'@FechaEnvio datetime,
'@FechaRecepcion datetime,
'@TipoDocumento varchar(9),
'@NroDocumento varchar(15),
'@Glosa varchar(350),
'@Autorizado money,
'@Retenciones money,
'@Multas money,
'@LiqPagable money,
'@Estado char(1),
'@usr_usuario varchar(15),
'@formulario varchar(15),
'@nro numeric (18) output
'
'as
'
'declare @CodPagoDirecto int
'declare @PDate varchar(10), @PTime varchar(8)
'exec edGetProcessDateTime @PDate out, @PTime out
'
'select @CodPagoDirecto = Valor
'From ac_correlativos
'where ID = 7  --7 ES DE PAGO DIRECTO
'Update ac_correlativos
'set valor = @CodPagoDirecto + 1
'Where Id = 7
'
'insert into pago_directo
'  ( ges_gestion,  CodPagoDirecto,    org_codigo,  Tipo_Cambio,  Rbr_Codigo,  TipoMoneda,  Codigo_Beneficiario,  FechaEnvio,  FechaRecepcion,  TipoDocumento,  NroDocumento,  Glosa,  Autorizado,  Retenciones,  Multas,  LiqPagable,  Estado,  usr_usuario, fecha_registro, hora_registro, formulario)
'values  (@ges_gestion, @CodPagoDirecto+1, @org_codigo, @Tipo_Cambio, @Rbr_Codigo, @TipoMoneda, @Codigo_Beneficiario, @FechaEnvio, @FechaRecepcion, @TipoDocumento, @NroDocumento, @Glosa, @Autorizado, @Retenciones, @Multas, @LiqPagable, @Estado, @usr_usuario, @PDate,         @PTime,    @formulario)
'
'set @nro = @CodPagoDirecto+1



End Function

Public Sub IGrptEjepptoConPar()
  Dim rsCom As New ADODB.Recordset
  Dim rsdev As New ADODB.Recordset
  Dim rspag As New ADODB.Recordset
  Dim rsIGrep_ppto_ejecConPar As New ADODB.Recordset
  
  Set rsCom = New ADODB.Recordset
  If rsCom.State = 1 Then rsCom.Close
  rsCom.Open "select * from v_Ppto_ComConCatPar order by codigo_convenio, codigo_categoria, par_codigo", db, adOpenKeyset, adLockReadOnly
  
  Set rsdev = New ADODB.Recordset
  If rsdev.State = 1 Then rsdev.Close
  rsdev.Open "select * from v_Ppto_DevConCatPar order by codigo_convenio, codigo_categoria, par_codigo", db, adOpenKeyset, adLockReadOnly
  
  Set rspag = New ADODB.Recordset
  If rspag.State = 1 Then rspag.Close
  rspag.Open "select * from v_Ppto_PagConCatPar order by codigo_convenio, codigo_categoria, par_codigo", db, adOpenKeyset, adLockReadOnly
  
  db.Execute "DELETE from IGrep_ppto_ejecConPar where maquina = '" & GlMaquina & "'"
  
  Set rsIGrep_ppto_ejecConPar = New ADODB.Recordset
  If rsIGrep_ppto_ejecConPar.State = 1 Then rsIGrep_ppto_ejecConPar.Close
  rsIGrep_ppto_ejecConPar.Open "select * from IGrep_ppto_ejecConPar where maquina = '" & GlMaquina & "' ", db, adOpenKeyset, adLockOptimistic
  
  db.BeginTrans
    While Not rsCom.EOF
      rsIGrep_ppto_ejecConPar.AddNew
'      rsIGrep_ppto_ejecConPar.CancelUpdate
      rsIGrep_ppto_ejecConPar!codigo_convenio = rsCom!codigo_convenio
      rsIGrep_ppto_ejecConPar!codigo_Categoria = rsCom!codigo_Categoria
      rsIGrep_ppto_ejecConPar!par_codigo = rsCom!par_codigo
      rsIGrep_ppto_ejecConPar!ComBs = rsCom!ComBs
      rsIGrep_ppto_ejecConPar!ComSus = rsCom!ComSus
      rsIGrep_ppto_ejecConPar!DevBs = 0
      rsIGrep_ppto_ejecConPar!DevSus = 0
      rsIGrep_ppto_ejecConPar!pagBs = 0
      rsIGrep_ppto_ejecConPar!pagSus = 0
      rsIGrep_ppto_ejecConPar!maquina = GlMaquina
      rsIGrep_ppto_ejecConPar.Update
      rsCom.MoveNext
    Wend

    If rsIGrep_ppto_ejecConPar.State = 1 Then rsIGrep_ppto_ejecConPar.Close
    While Not rsdev.EOF
      If rsIGrep_ppto_ejecConPar.State = 1 Then rsIGrep_ppto_ejecConPar.Close
      rsIGrep_ppto_ejecConPar.Open "select * from IGrep_ppto_ejecConPar where maquina = '" & GlMaquina & "' and codigo_convenio = '" & rsdev!codigo_convenio & "' and codigo_categoria = '" & rsdev!codigo_Categoria & "' and par_codigo = '" & rsdev!par_codigo & "' ", db, adOpenKeyset, adLockOptimistic
      If rsIGrep_ppto_ejecConPar.RecordCount < 1 Then
'        rsIGrep_ppto_ejecConPar.CancelUpdate
        rsIGrep_ppto_ejecConPar.AddNew
        rsIGrep_ppto_ejecConPar!codigo_convenio = rsdev!codigo_convenio
        rsIGrep_ppto_ejecConPar!codigo_Categoria = rsdev!codigo_Categoria
        rsIGrep_ppto_ejecConPar!par_codigo = rsdev!par_codigo
        rsIGrep_ppto_ejecConPar!ComBs = 0
        rsIGrep_ppto_ejecConPar!ComSus = 0
        rsIGrep_ppto_ejecConPar!maquina = GlMaquina
      Else
        rsIGrep_ppto_ejecConPar.MoveFirst
      End If
      rsIGrep_ppto_ejecConPar!DevBs = rsdev!DevBs
      rsIGrep_ppto_ejecConPar!DevSus = rsdev!DevSus
      rsIGrep_ppto_ejecConPar!pagBs = 0
      rsIGrep_ppto_ejecConPar!pagSus = 0
      rsIGrep_ppto_ejecConPar.Update
      rsdev.MoveNext
    Wend
    If rsIGrep_ppto_ejecConPar.State = 1 Then rsIGrep_ppto_ejecConPar.Close
    While Not rspag.EOF
      If rsIGrep_ppto_ejecConPar.State = 1 Then rsIGrep_ppto_ejecConPar.Close
      rsIGrep_ppto_ejecConPar.Open "select * from IGrep_ppto_ejecConPar where maquina = '" & GlMaquina & "' and codigo_convenio = '" & rspag!codigo_convenio & "' and codigo_categoria = '" & rspag!codigo_Categoria & "' and par_codigo = '" & rspag!par_codigo & "' ", db, adOpenKeyset, adLockOptimistic
      If rsIGrep_ppto_ejecConPar.RecordCount < 1 Then
'        rsIGrep_ppto_ejecConPar.CancelUpdate
        rsIGrep_ppto_ejecConPar.AddNew
        rsIGrep_ppto_ejecConPar!codigo_convenio = rspag!codigo_convenio
        rsIGrep_ppto_ejecConPar!codigo_Categoria = rspag!codigo_Categoria
        rsIGrep_ppto_ejecConPar!par_codigo = rspag!par_codigo
        rsIGrep_ppto_ejecConPar!ComBs = 0
        rsIGrep_ppto_ejecConPar!ComSus = 0
        rsIGrep_ppto_ejecConPar!DevBs = 0
        rsIGrep_ppto_ejecConPar!DevSus = 0
        rsIGrep_ppto_ejecConPar!maquina = GlMaquina
      Else
        rsIGrep_ppto_ejecConPar.MoveFirst
      End If
      rsIGrep_ppto_ejecConPar!pagBs = rspag!pagBs
      rsIGrep_ppto_ejecConPar!pagSus = rspag!pagSus
      rsIGrep_ppto_ejecConPar.Update
      rspag.MoveNext
    Wend
    If rsIGrep_ppto_ejecConPar.State = 1 Then rsIGrep_ppto_ejecConPar.Close
  db.CommitTrans
End Sub




'ALB







VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3765
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstdestino As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim VAR_SUB1 As String

Dim VAR_VTA As Integer

Private Sub Command1_Click()
'Actualiza Cuenta 2212
    VAR_SUB1 = "00"
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from CO_DIARIO WHERE Cod_Comp_Detalle = '4' ", db, adOpenKeyset, adLockBatchOptimistic
'
'    Set rs_aux4 = New ADODB.Recordset
'    If rs_aux4.State = 1 Then rs_aux4.Close
'    rs_aux4.Open "select * from CO_DIARIO WHERE Cod_Comp_Detalle = '3' ", db, adOpenKeyset, adLockBatchOptimistic
'    If rs_aux4.RecordCount > 0 Then
'       rs_aux4.MoveFirst
'       While Not rs_aux4.EOF
'        If rs_aux4!D_Cuenta = "2212" And rs_aux4!D_SubCta1 = "02" Then
'            VAR_SUB1 = "07"
'        End If
'        If rs_aux4!D_Cuenta = "2212" And rs_aux4!D_SubCta1 = "03" Then
'            VAR_SUB1 = "08"
'        End If
'        If rs_aux4!D_Cuenta = "2212" And rs_aux4!D_SubCta1 = "05" Then
'            VAR_SUB1 = "10"
'        End If
'        If rs_aux4!D_Cuenta = "2212" And rs_aux4!D_SubCta1 = "01" Then
'            VAR_SUB1 = "06"
'        End If
'        db.Execute "INSERT INTO CO_DIARIO (Cod_Comp, Cod_Comp_Detalle, D_Cuenta, D_Nombre, D_Subcta1, D_SubCta2, D_Aux1, D_Aux2, D_Aux3, D_Cta_Aux1, D_Des_Aux1, D_Cta_Aux2, D_Des_Aux2, D_Cta_Aux3, D_Des_Aux3, D_MontoBs, D_MontoDl, D_Cambio, H_Cuenta, H_Nombre, H_SubCta1, H_SubCta2, H_Aux1, H_Aux2, H_Aux3, H_Cta_Aux1, H_Des_Aux1, H_Cta_Aux2, H_Des_Aux2, H_Cta_Aux3, H_Des_Aux3, H_MontoBs, H_MontoDl, H_Cambio, NOMCTADEBE, NOMCTAHABER, Usr_codigo, Fecha_registro ) " & _
'        "VALUES (" & rs_aux4!Cod_Comp & ", '4', '2212', '" & rs_aux4!D_Nombre & "', '" & rs_aux4!D_SubCta1 & "', '" & rs_aux4!D_SubCta2 & "', '" & rs_aux4!d_Aux1 & "', '" & rs_aux4!d_Aux2 & "', '" & rs_aux4!d_Aux3 & "', '" & rs_aux4!D_Cta_Aux1 & "', '" & rs_aux4!D_Des_Aux1 & "', '" & rs_aux4!D_Cta_Aux2 & "', '" & rs_aux4!D_Des_Aux2 & "', '" & rs_aux4!D_Cta_Aux3 & "', '" & rs_aux4!D_Des_Aux3 & "', '" & rs_aux4!D_MontoBs & "', '" & rs_aux4!D_MontoDl & "', '" & rs_aux4!D_Cambio & "', " & _
'        "'1121', '" & VAR_SUB1 & "', '00', '00', '01', '03', '06', '" & rs_aux4!H_Cta_Aux1 & "', '" & rs_aux4!H_Des_Aux1 & "', '','','" & rs_aux4!H_Cta_Aux2 & "', '" & rs_aux4!H_Des_Aux2 & "', " & rs_aux4!H_MontoBs & ", " & rs_aux4!H_MontoDl & ", " & rs_aux4!H_Cambio & ", '" & rs_aux4!NOMCTADEBE & "', '" & rs_aux4!NOMCTAHABER & "', '" & rs_aux4!Usr_codigo & "', '" & rs_aux4!Fecha_registro & "')"
'
'            rs_aux4.MoveNext
'       Wend
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close

'Actualiza Auxiliares
'    Set rs_aux4 = New ADODB.Recordset
'    If rs_aux4.State = 1 Then rs_aux4.Close
'    rs_aux4.Open "select * from CO_DIARIO WHERE (Cod_Comp_Detalle = '3') AND (co_diario.D_Cuenta = '1121') AND (co_diario.D_Subcta1 = '02') ", db, adOpenKeyset, adLockBatchOptimistic
'    If rs_aux4.RecordCount > 0 Then
'       rs_aux4.MoveFirst
'       While Not rs_aux4.EOF
'        db.Execute "UPDATE CO_DIARIO SET D_Cta_Aux1 = '" & rs_aux4!D_Cta_Aux1 & "', D_Des_Aux1 = '" & IIf(IsNull(rs_aux4!D_Des_Aux1), "NO ASIGNADO", rs_aux4!D_Des_Aux1) & "' WHERE Cod_Comp= " & rs_aux4!Cod_Comp & " AND Cod_Comp_Detalle = '2' AND  (D_Cuenta = '1121') AND (D_Subcta1 = '02') "
'            rs_aux4.MoveNext
'       Wend
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close

'Actualiza Venta Detalle
    
    VAR_VTA = "0"

'    Set rs_aux4 = New ADODB.Recordset
'    If rs_aux4.State = 1 Then rs_aux4.Close
'    rs_aux4.Open "select * from ao_ventas_cabecera WHERE unidad_codigo = 'DNMAN'  ", db, adOpenKeyset, adLockBatchOptimistic
'    If rs_aux4.RecordCount > 0 Then
'       rs_aux4.MoveFirst
'       While Not rs_aux4.EOF
'        Set rs_aux1 = New ADODB.Recordset
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from ao_ventas_detalle WHERE venta_codigo = " & rs_aux4!venta_codigo & " ", db, adOpenKeyset, adLockBatchOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            VAR_VTA = rs_aux1.RecordCount + 1
'        End If

'        'TRAPO     --------------------------------------
'        Set rstdestino = New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_ventas_detalle WHERE venta_codigo = " & rs_aux4!venta_codigo & " and bien_codigo= '4211' ", db, adOpenKeyset, adLockBatchOptimistic
'        If rstdestino.RecordCount = 0 Then
'            db.Execute "INSERT INTO ao_ventas_detalle (ges_gestion, venta_codigo, bien_codigo, cotiza_codigo, venta_codigo_new, venta_codigo_det, venta_det_cantidad, venta_precio_unitario_bs, venta_descuento_bs, venta_precio_total_bs, venta_precio_unitario_dol, venta_descuento_dol, venta_precio_total_dol, concepto_venta, grupo_codigo, subgrupo_codigo, par_codigo, " & _
'            " bien_cantidad_por_empaque, tipo_descuento, almacen_codigo, modelo_codigo, modelo_codigo1, modelo_codigo_h, modelo_codigo_x, modelo_elegido, modelo_elegido_h , modelo_elegido_x, estado_almacen, estado_codigo, usr_codigo, fecha_registro, pais_codigo) " & _
'            "VALUES ('" & rs_aux4!ges_gestion & "', " & rs_aux4!venta_codigo & ", '4211',      '1',           '0',              " & VAR_VTA & ",  '1',                '0',                      '0',                '0',                   '0',                       '0',                 '0',                    'TRAPO',        '30000',      '33000',         '33100',  " & _
'            " '1',                       '0',            '0',            'S/M',         'S/M',          'S/M',           'S/M',           'S',            'N',               'N',              'REG',          'REG',        'ADMIN', '" & rs_aux4!fecha_registro & "', 'BOL' )"
'        End If

'        'GASOLINA  --------------------------------------
'        Set rstdestino = New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_ventas_detalle WHERE venta_codigo = " & rs_aux4!venta_codigo & " and bien_codigo= '479' ", db, adOpenKeyset, adLockBatchOptimistic
'        If rstdestino.RecordCount = 0 Then
'            db.Execute "INSERT INTO ao_ventas_detalle (ges_gestion, venta_codigo, bien_codigo, cotiza_codigo, venta_codigo_new, venta_codigo_det, venta_det_cantidad, venta_precio_unitario_bs, venta_descuento_bs, venta_precio_total_bs, venta_precio_unitario_dol, venta_descuento_dol, venta_precio_total_dol, concepto_venta, grupo_codigo, subgrupo_codigo, par_codigo, " & _
'            " bien_cantidad_por_empaque, tipo_descuento, almacen_codigo, modelo_codigo, modelo_codigo1, modelo_codigo_h, modelo_codigo_x, modelo_elegido, modelo_elegido_h , modelo_elegido_x, estado_almacen, estado_codigo, usr_codigo, fecha_registro, pais_codigo) " & _
'            "VALUES ('" & rs_aux4!ges_gestion & "', " & rs_aux4!venta_codigo & ", '479', '1', '0', " & VAR_VTA & ", '1', '0','0', '0', '0', '0', '0', 'GASOLINA', '30000', '34000', '34110',  " & _
'            " '1', '0', '0', 'S/M', 'S/M', 'S/M', 'S/M', 'S', 'N', 'N', 'REG', 'REG', 'ADMIN', '" & rs_aux4!fecha_registro & "', 'BOL' )"
'        End If

'        'ACEITE PREPARADO --------------------------------------
'        Set rstdestino = New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_ventas_detalle WHERE venta_codigo = " & rs_aux4!venta_codigo & " and bien_codigo= '500' ", db, adOpenKeyset, adLockBatchOptimistic
'        If rstdestino.RecordCount = 0 Then
'            db.Execute "INSERT INTO ao_ventas_detalle (ges_gestion, venta_codigo, bien_codigo, cotiza_codigo, venta_codigo_new, venta_codigo_det, venta_det_cantidad, venta_precio_unitario_bs, venta_descuento_bs, venta_precio_total_bs, venta_precio_unitario_dol, venta_descuento_dol, venta_precio_total_dol, concepto_venta, grupo_codigo, subgrupo_codigo, par_codigo, " & _
'            " bien_cantidad_por_empaque, tipo_descuento, almacen_codigo, modelo_codigo, modelo_codigo1, modelo_codigo_h, modelo_codigo_x, modelo_elegido, modelo_elegido_h , modelo_elegido_x, estado_almacen, estado_codigo, usr_codigo, fecha_registro, pais_codigo) " & _
'            "VALUES ('" & rs_aux4!ges_gestion & "', " & rs_aux4!venta_codigo & ", '500', '1', '0', " & VAR_VTA & ", '1', '0','0', '0', '0', '0', '0', 'ACEITE PREPARADO', '30000', '34000', '34110',  " & _
'            " '1', '0', '0', 'S/M', 'S/M', 'S/M', 'S/M', 'S', 'N', 'N', 'REG', 'REG', 'ADMIN', '" & rs_aux4!fecha_registro & "', 'BOL' )"
'        End If

'        'ACEITE DELGADO 20/50 --------------------------------------
'        Set rstdestino = New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_ventas_detalle WHERE venta_codigo = " & rs_aux4!venta_codigo & " and bien_codigo= '4529' ", db, adOpenKeyset, adLockBatchOptimistic
'        If rstdestino.RecordCount = 0 Then
'            db.Execute "INSERT INTO ao_ventas_detalle (ges_gestion, venta_codigo, bien_codigo, cotiza_codigo, venta_codigo_new, venta_codigo_det, venta_det_cantidad, venta_precio_unitario_bs, venta_descuento_bs, venta_precio_total_bs, venta_precio_unitario_dol, venta_descuento_dol, venta_precio_total_dol, concepto_venta, grupo_codigo, subgrupo_codigo, par_codigo, " & _
'            " bien_cantidad_por_empaque, tipo_descuento, almacen_codigo, modelo_codigo, modelo_codigo1, modelo_codigo_h, modelo_codigo_x, modelo_elegido, modelo_elegido_h , modelo_elegido_x, estado_almacen, estado_codigo, usr_codigo, fecha_registro, pais_codigo) " & _
'            "VALUES ('" & rs_aux4!ges_gestion & "', " & rs_aux4!venta_codigo & ", '4529', '1', '0', " & VAR_VTA & ", '1', '0','0', '0', '0', '0', '0', 'ACEITE DELGADO 20/50', '30000', '34000', '34110',  " & _
'            " '1', '0', '0', 'S/M', 'S/M', 'S/M', 'S/M', 'S', 'N', 'N', 'REG', 'REG', 'ADMIN', '" & rs_aux4!fecha_registro & "', 'BOL' )"
'        End If

'        rs_aux4.MoveNext
'       Wend
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close

End Sub

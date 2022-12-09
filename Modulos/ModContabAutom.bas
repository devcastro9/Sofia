Attribute VB_Name = "ModContabAutom"
Option Explicit

Public Function ADec(ByVal Nro As Double) As String
    ADec = Replace(CStr(Nro), ",", ".")
End Function

Public Function GetSelect(ByVal query_select As String) As ADODB.Recordset
    On Error GoTo Handler
    ' Si la conexion esta vacia
    If db Is Nothing Then
        Set GetSelect = Nothing
        Exit Function
    End If
    ' Si la conexion no es vacia, entonces cargar datos solo lectura
    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    With result
        .Open query_select, db, adOpenStatic, adLockReadOnly
        If .EOF And .BOF Then
            Set GetSelect = Nothing
            .Close
            Exit Function
        End If
        Set GetSelect = result
        .Close
    End With
    Exit Function
CleanExit:
    If Not result Is Nothing And result.State = adStateOpen Then result.Close
    Exit Function
Handler:
    If Err.Number > 0 Then
        MsgBox ("Select statement error: " & Err.Number & " : " & Err.Description)
    Else
        Err.Clear
    End If
    Resume CleanExit
End Function

Public Sub ExecProcedure(ByVal query_stored As String)
    On Error GoTo Handler
    ' Si la conexion esta vacia
    If db Is Nothing Then
        Exit Sub
    End If
    ' Si la conexion no es vacia, entonces ejecutar el procedimiento almacenado
    Dim stored As ADODB.Command
    Set stored = New ADODB.Command
    With stored
        .ActiveConnection = db
        .CommandText = query_stored
        .CommandType = adCmdText
        .Execute
    End With
Handler:
    MsgBox ("Stored Procedure error: " & Err.Number & " : " & Err.Description)
End Sub

Public Sub Contabiliza_Contratos(ByVal venta_codigo As Long)
    On Error GoTo Handler
    'Contabilizacion al momento de aprobacion
    'Vista relativa a contabilizacion
    Dim rs_data99 As New ADODB.Recordset
    'Declaracion de variables
    Dim VAR_CODTIPO As String
    Dim VAR_PARTIDA As String
    Dim VAR_EMPRESA As Integer
    Dim VAR_DPTO As Integer
    Dim VAR_TIPOCOMPID As Integer
    Dim VAR_FECHA As Date
    Dim VAR_MONEDAID As Integer
    Dim VAR_TIPOCAMBIO As Double
    Dim VAR_DEBEORG As Double
    Dim VAR_HABERORG As Double
    'Glosas superiores
    Dim VAR_EntregadoA As String
    Dim VAR_CONCEPTO As String
    'Otros valores
    Dim VAR_ConFac As Integer
    Dim VAR_SinFac As Integer
    Dim VAR_Automatico As Integer
    Dim VAR_GLOSA As String
    Dim VAR_TipoNotaId As Integer
    Dim VAR_NotaNro As Integer
    Dim VAR_EstadoId As Integer
    Dim VAR_iConcurrency_id As Integer
    Dim VAR_TipoAsientoId As Integer
    Dim VAR_CentroCostoId As Integer
    Dim VAR_TipoRetencionId As Integer
    Dim VAR_TipoId As Integer
    Dim VAR_CompDetIdOrg As Integer
    Dim VAR_AuxAna As String
    ' Reverse identification
    Dim cod1 As Long
    Dim idCAutom As Integer
    ' Variables adicionales
    Dim query_data As String
    Dim query_stored As String
    ' Data
    query_data = "SELECT [codTipo], [par], [empresaId], [dpto], [tipoCompId], [fecha], [tm], [tc], [bs], [dol], CONCAT('RESPONSABLE: ', [beneficiario_codigo], ' - ', [beneficiario_denominacion]) AS EntregadoA, CONCAT('REG. DEVENGADO ', [trans_descripcion2], ' ', [depto_descripcion], ' EDIFICIO ', [edif_codigo_corto], ' ', [edif_descripcion], ' VIGENCIA DEL ', FORMAT([venta_fecha_inicio], 'dd/MM/yyyy'), ' AL ', FORMAT([venta_fecha_fin], 'dd/MM/yyyy'), ' S/G ', [contratoOds], ' ', [unidad_codigo_ant]) AS PorConcepto, [solicitud_tipo], [nro], CONCAT('INGRESO POR: ', [venta_descripcion], '-  NRO. VENTA: ', [nro]) AS glosa, [CentroCostoId], [edif_codigo] " & _
        "FROM [dbo].[conta_contratos] " & _
        "WHERE [vtipo] <> 'A' AND [cod1] = " & venta_codigo
    Set rs_data99 = GetSelect(query_data)
    If rs_data99 Is Nothing Then
        If rs_data99.State = adStateOpen Then rs_data99.Close
        Exit Sub
    End If
    With rs_data99
        .Open
        ' Si no se encontro ningun registro
        If .RecordCount = 0 Then
            If .State = adStateOpen Then .Close
            Exit Sub
        End If
        ' Si existen registros se asignan a variables
        .MoveFirst
        VAR_CODTIPO = .Fields("codTipo")
        VAR_PARTIDA = .Fields("par")
        VAR_EMPRESA = .Fields("empresaId")
        VAR_DPTO = .Fields("dpto")
        VAR_TIPOCOMPID = .Fields("tipoCompId")
        VAR_FECHA = CDate(.Fields("Fecha"))
        'VAR_TIPOCAMBIO = IIf(IsNull(rs_data99!tc), GlTipoCambioOficial, rs_data99!tc)
        VAR_TIPOCAMBIO = GlTipoCambioOficial
        VAR_MONEDAID = 1
        VAR_DEBEORG = .Fields("bs")
        VAR_HABERORG = .Fields("bs")
'        If .Fields("tm") = "BOB" Then
'            VAR_MONEDAID = 1
'            VAR_DEBEORG = .Fields("bs")
'            VAR_HABERORG = .Fields("bs")
'        Else
'            VAR_MONEDAID = 2
'            VAR_DEBEORG = .Fields("dol")
'            VAR_HABERORG = .Fields("dol")
'        End If
        'Glosas superiores
        VAR_EntregadoA = .Fields("EntregadoA")
        VAR_CONCEPTO = .Fields("PorConcepto")
        ' Otros valores
        VAR_ConFac = 0
        VAR_SinFac = 1
        VAR_Automatico = 1 '0 Permite edicion, 1 no permite editar
        VAR_TipoNotaId = .Fields("solicitud_tipo")
        VAR_NotaNro = .Fields("Nro")
        ' Glosa general
        VAR_GLOSA = .Fields("glosa")
        VAR_EstadoId = 11 'Libro Mayor requiere que sean de EstadoId = 10 Cerrado OR EstadoId = 11 Abierto
        VAR_TipoAsientoId = 0 ' Operativo
        VAR_CentroCostoId = .Fields("CentroCostoId")
        VAR_TipoRetencionId = 0
        VAR_TipoId = 0
        VAR_CompDetIdOrg = 0
        VAR_AuxAna = .Fields("edif_codigo")
        'Reverse identification
        cod1 = venta_codigo
        idCAutom = 1 'Caso contratos
        query_stored = "EXECUTE [dbo].[conta_ingresos] '" & VAR_CODTIPO & "', '" & VAR_PARTIDA & "', " & VAR_EMPRESA & ", " & VAR_DPTO & ", " & VAR_TIPOCOMPID & ", '" & VAR_FECHA & "', " & VAR_MONEDAID & ", '" & ADec(VAR_TIPOCAMBIO) & "', '" & ADec(VAR_DEBEORG) & "', '" & ADec(VAR_HABERORG) & "', '" & VAR_EntregadoA & "', '" & VAR_CONCEPTO & "', " & VAR_ConFac & ", " & VAR_SinFac & ", " & VAR_Automatico & ", '" & VAR_GLOSA & "', " & VAR_TipoNotaId & ", " & VAR_NotaNro & ", " & VAR_EstadoId & ", '" & glusuario & "', " & VAR_TipoAsientoId & ", " & VAR_CentroCostoId & ", " & VAR_TipoRetencionId & ", " & VAR_TipoId & ", " & VAR_CompDetIdOrg & ", '" & VAR_AuxAna & "', " & venta_codigo & ", " & 0 & ", " & idCAutom
        'Debug.Print query_stored
        Call ExecProcedure(query_stored)
        .Close
    End With
Handler:
    If Err.Number > 0 Then
        MsgBox ("Contabilizacion de contratos: " & Err.Number & " : " & Err.Description)
    End If
End Sub

Public Sub Contabiliza_Facturacion(ByVal IdFactura As Long)
    On Error GoTo Handler
    'Contabilizacion al momento de aprobacion
    '=======================================================
    ' No usar: Ya se implementa en facturacion electronica
    '=======================================================
    Dim sql_exec As String
    sql_exec = "EXECUTE [dbo].[conta_fac_elec] " & IdFactura
    Call ExecProcedure(sql_exec)
Handler:
    If Err.Number > 0 Then
        MsgBox ("Contabilizacion de facturacion: " & Err.Number & " : " & Err.Description)
    End If
End Sub

Public Sub Contabiliza_Cobranzas(ByVal IdTraspasoBancos As Long)
    On Error GoTo Handler
    '====================================
    'Falta mejorar
    '====================================
'    Dim rs_data99 As New ADODB.Recordset
'    Dim VAR_CUENTA As String
'    Dim VAR_PARTIDA As String
'    Dim VAR_EMPRESA As Integer
'    Dim VAR_DPTO As Integer
'    Dim VAR_TIPOCOMPID As Integer
'    Dim VAR_FECHA As Date
'    Dim VAR_TIPOCAMBIO As Double
'    Dim VAR_BOB As Double
'    Dim VAR_USD As Double
'    'Glosas superiores
'    Dim VAR_EntregadoA As String
'    Dim VAR_CONCEPTO As String
'    'Glosa general
'    Dim VAR_GLOSA As String
'    'Otros valores
'    Dim VAR_ConFac As Integer
'    Dim VAR_SinFac As Integer
'    Dim VAR_Automatico As Integer
'    Dim VAR_TipoNotaId As Integer
'    Dim VAR_NotaNro As String
'    Dim VAR_EstadoId As Integer
'    Dim VAR_iConcurrency_id As Integer
'    Dim VAR_TipoAsientoId As Integer
'    Dim VAR_CentroCostoId As Integer
'    Dim VAR_TipoRetencionId As Integer
'    Dim VAR_TipoId As Integer
'    Dim VAR_CompDetIdOrg As Integer
'    Dim VAR_AuxAna As String
'    ' Reverse identification
'    Dim cod1 As Long
'    Dim idCAutom As Integer
'    ' Variables adicionales
'    Dim query_data As String
'    Dim query_stored As String
'    ' Data
'    query_data = "SELECT [cta], [par], [vtipo], [depto], [fecha], [bs2], [dol2], CONCAT([edcorto], ' -  EDIFICIO: ', [eddesc]) AS EntregadoA, CONCAT('REGISTRO FACTURA NRO. ', [nrofac], ' CON REGISTRO DE COBRANZA NRO. ', [doc_nro], ' ', [transdes], ' EDIFICIO ', [edcorto], ' ', [eddesc], ' S/G ', [contratoOds], ' ', [cods], ' ', [mes], ' TRASP.TES. ', [correl]) AS PorConcepto, CONCAT([glosa], ' Codigo de bancarizacion: ', [codBancarizacion]) AS glosa, [soltipo], [notaNro], [CentroCostoId], [edif_codigo] FROM [dbo].[conta_cobranzas] WHERE [cod3] = " & IdTraspasoBancos
'    Debug.Print query_data
'    If rs_data99.State = adStateOpen Then rs_data99.Close
'    Set rs_data99 = GetSelect(query_data)
'    If rs_data99.State = adStateClosed Then Exit Sub
'    If rs_data99.RecordCount > 0 Then
'        rs_data99.MoveFirst
'        VAR_TIPOCOMPID = 1
'        VAR_ConFac = 0 'Con factura
'        VAR_SinFac = 1 'Sin factura
'        VAR_Automatico = 1 '0 Permite edicion, 1 no permite editar
'        VAR_EstadoId = 11 'Libro Mayor requiere que sean de EstadoId = 10 Cerrado OR EstadoId = 11 Abierto
'        VAR_TipoAsientoId = 0 'Operativo
'        VAR_TipoRetencionId = 0
'        VAR_TipoId = 0
'        VAR_CompDetIdOrg = 0
'        'Reverse identification
'        cod1 = IdTraspasoBancos
'        idCAutom = 3 'Caso Cobranzas
'        Do While Not rs_data99.EOF
'            VAR_CUENTA = rs_data99!cta
'            VAR_PARTIDA = rs_data99!par
'            VAR_EMPRESA = IIf(rs_data99!vtipo = "G", 2, 1)
'            VAR_DPTO = rs_data99!depto
'            VAR_FECHA = rs_data99!Fecha
'            VAR_BOB = CDbl(rs_data99!bs2)
'            VAR_USD = CDbl(rs_data99!dol2)
'            'Glosas superiores
'            VAR_EntregadoA = rs_data99!EntregadoA
'            VAR_CONCEPTO = rs_data99!PorConcepto
''            VAR_EntregadoA = rs_data99!edcorto & " - EDIFICIO: " & rs_data99!eddesc
''            VAR_CONCEPTO = "REGISTRO FACTURA NRO. " & rs_data99!nrofac & " CON RECIBO DE COBRANZA NRO. " & rs_data99!doc_nro
''            VAR_CONCEPTO = VAR_CONCEPTO & " " & rs_data99!transdes & " EDIFICIO " & rs_data99!edcorto & " " & rs_data99!eddesc
''            VAR_CONCEPTO = VAR_CONCEPTO & " S/G " & rs_data99!contratoOds & " " & rs_data99!cods & " " & rs_data99!mes & " TRASP.TES. " & rs_data99!correl
'            'Glosa general
'            VAR_GLOSA = rs_data99!glosa
'            'Otros
'            VAR_TipoNotaId = rs_data99!soltipo
'            VAR_NotaNro = rs_data99!notaNro
'            VAR_CentroCostoId = rs_data99!CentroCostoId
'            VAR_AuxAna = rs_data99!edif_codigo
'            'query stored
'            query_stored = "EXECUTE [dbo].[conta_cobranzas] '" & VAR_CUENTA & "', '" & VAR_PARTIDA & "', " & VAR_EMPRESA & ", " & VAR_DPTO & ", " & VAR_TIPOCOMPID & ", '" & VAR_FECHA & "', '" & ADec(VAR_BOB) & "', '" & ADec(VAR_USD) & "', '" & VAR_EntregadoA & "', '" & VAR_CONCEPTO & "', " & VAR_ConFac & ", " & VAR_SinFac & ", " & VAR_Automatico & ", '" & VAR_GLOSA & "', " & VAR_TipoNotaId & ", " & VAR_NotaNro & ", " & VAR_EstadoId & ", '" & glusuario & "', " & VAR_TipoAsientoId & ", " & VAR_CentroCostoId & ", " & VAR_TipoRetencionId & ", " & VAR_TipoId & ", " & VAR_CompDetIdOrg & ", '" & VAR_AuxAna & "', " & cod1 & ", " & 0 & ", " & idCAutom
'            Debug.Print query_stored
'            ' EXEC stored procedured
'            Call ExecProcedure(query_stored)
'            ' Siguiente
'            rs_data99.MoveNext
'        Loop
'    End If
'    If rs_data99.State = adStateOpen Then rs_data99.Close
'    MsgBox "Cobranza contabilizada", vbInformation, "Hecho"
Handler:
    If Err.Number <> 0 Then
        MsgBox ("Cobranzas " & Err.Number & " : " & Err.Description)
    End If
End Sub

Public Sub Contabiliza_Pago(ByVal Id As Long, ByVal usuario_apr As String)
On Error GoTo Handler:
    Dim sql_exec As String
    sql_exec = "EXECUTE [dbo].[conta_tes_egreso] " & Id & ", '" & usuario_apr & "'"
    Call ExecProcedure(sql_exec)
Handler:
    If Err.Number > 0 Then
        MsgBox ("Cobranzas " & Err.Number & " : " & Err.Description)
    End If
End Sub

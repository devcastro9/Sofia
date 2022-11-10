Attribute VB_Name = "ModSeguridad"
Option Explicit

Public Sub SeguridadSet(ByRef frmCurrent As Form)
    ' ================================================
    '       Authorization
    ' ================================================
    On Error GoTo Handler
    Dim ctrIter As Control
    Dim rs_Right As ADODB.Recordset
    Dim sqlRolesAcceso As String
    ' Mapeado de controles y completado de Formularios
    Dim consultaDatos As String
'    If frmCurrent.Name <> "frmLogin" Then
'        consultaDatos = AccesoDatos_Roles()
'    End If
    Call Mapeado(frmCurrent)
    ' Acceso permitido al usuario por roles en el formulario
    sqlRolesAcceso = "SELECT [c].[NombreControl], CAST(MAX(CAST([d].[Visible] AS TINYINT)) AS BIT) AS [Visible], CAST(MAX(CAST([d].[Enabled] AS TINYINT)) AS BIT) AS [Enabled] " & _
                "FROM [dbo].[gc_usuarios_roles] AS [ur] " & _
                "INNER JOIN [dbo].[gc_roles] AS [r] " & _
                "ON [ur].[IdRole] = [r].[IdRole] " & _
                "INNER JOIN [dbo].[gc_right] AS [d] " & _
                "ON [r].[IdRole] = [d].[IdRole] " & _
                "INNER JOIN [dbo].[gc_controles] AS [c] " & _
                "ON [d].[CtrlId] = [c].[CtrlId] " & _
                "INNER JOIN [dbo].[gc_formularios] AS [f] " & _
                "ON [c].[FormId] = [f].[FormId] " & _
                "WHERE [f].[NombreForm] = '" & frmCurrent.Name & "' AND [ur].[usr_codigo] = '" & glusuario & "' " & _
                "GROUP BY [c].[NombreControl]"
    ' Right vs Forms.Controls (Un ciclo por Datos SQL y N-Ciclos por Controles de Formularios en RAM)
    Set rs_Right = New ADODB.Recordset
    With rs_Right
        .Open sqlRolesAcceso, db, adOpenForwardOnly, adLockReadOnly
        If .EOF And .BOF Then
            .Close ' No hay registros
            Exit Sub
        End If
        Do While Not .EOF ' Mientras haya registros
            For Each ctrIter In frmCurrent.Controls
                If ctrIter.Name = .Fields("NombreControl") Then
                    ctrIter.Visible = .Fields("Visible")
                    ctrIter.Enabled = .Fields("Enabled")
                    Exit For
                End If
            Next
            .MoveNext
        Loop
        .Close
    End With
Handler:
    If Err.Number > 0 Then
        MsgBox "Database error " & Err.Number & " : " & Err.Description
    Else
        Err.Clear
    End If
End Sub

Public Sub Mapeado(ByVal frmCurrent As Form)
    Dim ctrIter As Control
    Dim sqlMapeado As String
    sqlMapeado = ""
    For Each ctrIter In frmCurrent.Controls
        sqlMapeado = "EXECUTE [dbo].[mapear_controles] '" & ctrIter.Name & "', '" & TypeName(ctrIter) & "', '" & frmCurrent.Name & "', '" & frmCurrent.Caption & "'"
        Debug.Print sqlMapeado
        db.Execute sqlMapeado
    Next
End Sub

Public Function AccesoDatos_Roles() As String
    Dim departamentos(1 To 10) As Boolean
    Dim sqlAccesoDatos As String
    Dim rs_Acceso As ADODB.Recordset
    Dim contador As Integer
    Dim Index As Integer
    Dim deptos_sql As String
    contador = 0
    'Consulta SQL
    sqlAccesoDatos = "SELECT CAST(MAX(CAST([r].[CHQ] AS TINYINT)) AS BIT) AS [chq], CAST(MAX(CAST([r].[LPZ] AS TINYINT)) AS BIT) AS [lpz], CAST(MAX(CAST([r].[CBB] AS TINYINT)) AS BIT) AS [cbb], CAST(MAX(CAST([r].[ORU] AS TINYINT)) AS BIT) AS [oru], CAST(MAX(CAST([r].[PTS] AS TINYINT)) AS BIT) AS [pts], CAST(MAX(CAST([r].[TJA] AS TINYINT)) AS BIT) AS [tja], CAST(MAX(CAST([r].[SCZ] AS TINYINT)) AS BIT) AS [scz], CAST(MAX(CAST([r].[BEN] AS TINYINT)) AS BIT) AS [ben], CAST(MAX(CAST([r].[PDO] AS TINYINT)) AS BIT) AS [pdo], CAST(MAX(CAST([r].[EXT] AS TINYINT)) AS BIT) AS [ext] " & _
                "FROM [dbo].[gc_usuarios_roles] AS [u] " & _
                "INNER JOIN [dbo].[gc_roles] AS [r] " & _
                "ON [u].[IdRole] = [r].[IdRole] " & _
                "WHERE [u].[usr_codigo] = '" & glusuario & "' AND [r].[estado_codigo] = 'APR'"
    'RecordSet
    Set rs_Acceso = New ADODB.Recordset
    With rs_Acceso
        .Open sqlAccesoDatos, db, adOpenForwardOnly, adLockReadOnly
        If .EOF And .BOF Then
            .Close ' No hay registros
            AccesoDatos_Roles = ""
            Exit Function
        End If
        Do While Not .EOF
            departamentos(1) = .Fields("chq")
            departamentos(2) = .Fields("lpz")
            departamentos(3) = .Fields("cbb")
            departamentos(4) = .Fields("oru")
            departamentos(5) = .Fields("pts")
            departamentos(6) = .Fields("tja")
            departamentos(7) = .Fields("scz")
            departamentos(8) = .Fields("ben")
            departamentos(9) = .Fields("pdo")
            departamentos(10) = .Fields("ext")
            .MoveNext
        Loop
        .Close
    End With
    'Creacion del query
    deptos_sql = ""
    For Index = 1 To 10
        If departamentos(Index) Then
            deptos_sql = deptos_sql & Index & ", "
            contador = contador + 1
        End If
    Next Index
    deptos_sql = Left(deptos_sql, Len(deptos_sql) - 2)
    Select Case contador
        Case 1:
            AccesoDatos_Roles = " AND depto_codigo = " & deptos_sql
        Case 2 To 9:
            AccesoDatos_Roles = " AND depto_codigo IN (" & deptos_sql & ")"
        Case Else:
            AccesoDatos_Roles = ""
    End Select
    Exit Function
End Function

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
    Dim sqlMapeado As String
    Dim Mapear As Boolean
    ' Mapeado de controles y completado de Formularios
    Mapear = False
    If Mapear Then
        For Each ctrIter In frmCurrent.Controls
            sqlMapeado = "EXECUTE [dbo].[mapear_controles] '" & ctrIter.Name & "', '" & TypeName(ctrIter) & "', '" & frmCurrent.Name & "', '" & frmCurrent.Caption & "'"
            Debug.Print sqlMapeado
            db.Execute sqlMapeado
        Next
    End If
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

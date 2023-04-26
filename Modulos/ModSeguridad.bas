Attribute VB_Name = "ModSeguridad"
Option Explicit

Public Sub SeguridadSet(ByRef frmCurrent As Form)
'     ' ================================================
'     '       Authorization
'     ' ================================================
'     On Error GoTo Handler
'     Dim ctrUno As Control
'     Dim rs_Roles As ADODB.Recordset
'     Dim rs_Right As ADODB.Recordset
'     Dim SqlQuery As String
'     Dim sqlRoles As String
'     Dim sqlMapeado As String
'     Dim nombreForm As String
'     nombreForm = frmCurrent.Name
' '    For Each ctrUno In frmCurrent.Controls
' '        sqlMapeado = "EXECUTE [dbo].[mapear_controles] '" & ctrUno.Name & "', '" & TypeName(ctrUno) & "', '" & nombreForm & "'"
' '        Debug.Print sqlMapeado
' '        'db.Execute sqlMapeado
' '    Next
    
'     ' Roles
'     'sqlRoles = ""
' '    With rs_Roles
' '        .Open sqlRoles, db, adOpenForwardOnly, adLockReadOnly
' '
' '    End With
' '    sqlQuery = ""
' '    ' Right vs Forms.Controls
' '    With rs_Right
' '        .Open sqlQuery, db, adOpenForwardOnly, adLockReadOnly
' '        If .EOF And .BOF Then
' '            ' No hay registros
' '            .Close
' '            Exit Sub
' '        End If
' '        Do While Not .EOF
' '
' '            .MoveNext
' '        Loop
' '        .Close
' '    End With
' Handler:
'     If Err.Number > 0 Then
'         MsgBox "Database error " & Err.Number & " : " & Err.Description
'     End If
End Sub

Public Function DepartamentoPorRol() As String
    On Error GoTo Handler
    Dim sqlRolDepartamento As String
    Dim rs_DepartamentoPorRoles As ADODB.Recordset
    If Len(glusuario) < 2 Then
        DepartamentoPorRol = ""
        Exit Function
    End If
    sqlRolDepartamento = "SELECT [depto_codigo] FROM [dbo].[gv_rol_departamento] WHERE [usr_codigo] = '" & glusuario & "'"
    Set rs_DepartamentoPorRoles = New ADODB.Recordset
    With rs_DepartamentoPorRoles
        .Open sqlRolDepartamento, db, adOpenStatic, adLockReadOnly
        If .EOF And .BOF Then
            DepartamentoPorRol = ""
            .Close
            Exit Function
        End If
        .MoveFirst
        Select Case .RecordCount
            Case 1
                DepartamentoPorRol = " AND depto_codigo = " & .Fields("depto_codigo") & " "
            Case Is > 1:
                DepartamentoPorRol = " AND depto_codigo IN ("
                Do
                    DepartamentoPorRol = DepartamentoPorRol & "'" & .Fields("depto_codigo") & "' "
                    .MoveNext
                    If .EOF Then Exit Do
                    DepartamentoPorRol = DepartamentoPorRol & ", "
                Loop
                DepartamentoPorRol = DepartamentoPorRol & ") "
            Case Else:
                DepartamentoPorRol = ""
        End Select
        .Close
    End With
Handler:
    If Err.Number > 0 Then
        MsgBox ("Select 'Departamento' error: " & Err.Number & " : " & Err.Description)
    Else
        Err.Clear
    End If
End Function

Public Function UnidadPorRol() As String
    Dim sqlRolUnidad As String
    Dim rs_UnidadPorRoles As ADODB.Recordset
    glusuario = "ADMIN"
    If Len(glusuario) < 2 Then
        UnidadPorRol = ""
        Exit Function
    End If
    sqlRolUnidad = "SELECT [unidad_codigo] FROM [dbo].[gv_rol_unidad] WHERE [usr_codigo] = '" & glusuario & "'"
    Set rs_UnidadPorRoles = New ADODB.Recordset
    With rs_UnidadPorRoles
        .Open sqlRolUnidad, db, adOpenStatic, adLockReadOnly
        If .EOF And .BOF Then
            UnidadPorRol = ""
            .Close
            Exit Function
        End If
        .MoveFirst
        Select Case .RecordCount
            Case 1
                UnidadPorRol = " AND unidad_codigo = " & .Fields("unidad_codigo") & " "
            Case Is > 1:
                UnidadPorRol = " AND unidad_codigo IN ("
                Do
                    UnidadPorRol = UnidadPorRol & "'" & .Fields("unidad_codigo") & "' "
                    .MoveNext
                    If .EOF Then Exit Do
                    UnidadPorRol = UnidadPorRol & ", "
                Loop
                UnidadPorRol = UnidadPorRol & ") "
            Case Else:
                UnidadPorRol = ""
        End Select
        .Close
    End With
Handler:
    If Err.Number > 0 Then
        MsgBox ("Select 'Unidad' error: " & Err.Number & " : " & Err.Description)
    Else
        Err.Clear
    End If
End Function


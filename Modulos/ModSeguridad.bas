Attribute VB_Name = "ModSeguridad"
Option Explicit

Public Sub SeguridadSet(ByRef frmCurrent As Form)
    ' ================================================
    '       Authorization
    ' ================================================
    On Error GoTo Handler
    Dim ctrUno As Control
    Dim rs_Roles As ADODB.Recordset
    Dim rs_Right As ADODB.Recordset
    Dim SqlQuery As String
    Dim sqlRoles As String
    Dim sqlMapeado As String
    Dim Mapear As Boolean
    Mapear = True
    If Mapear Then
        For Each ctrUno In frmCurrent.Controls
            sqlMapeado = "EXECUTE [dbo].[mapear_controles] '" & ctrUno.Name & "', '" & TypeName(ctrUno) & "', '" & frmCurrent.Name & "', '" & frmCurrent.Caption & "'"
            Debug.Print sqlMapeado
            db.Execute sqlMapeado
        Next
    End If
    ' Roles
    'sqlRoles = ""
'    With rs_Roles
'        .Open sqlRoles, db, adOpenForwardOnly, adLockReadOnly
'
'    End With
'    sqlQuery = ""
'    ' Right vs Forms.Controls
'    With rs_Right
'        .Open sqlQuery, db, adOpenForwardOnly, adLockReadOnly
'        If .EOF And .BOF Then
'            ' No hay registros
'            .Close
'            Exit Sub
'        End If
'        Do While Not .EOF
'
'            .MoveNext
'        Loop
'        .Close
'    End With
Handler:
    If Err.Number > 0 Then
        MsgBox "Database error " & Err.Number & " : " & Err.Description
    End If
End Sub

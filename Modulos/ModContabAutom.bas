Attribute VB_Name = "ModContabAutom"
Option Explicit

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
        MsgBox ("Select statement error " & Err.Number & " : " & Err.Description)
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
    MsgBox ("Stored Procedure error " & Err.Number & " : " & Err.Description)
End Sub


Attribute VB_Name = "ModContabAutom"
Option Explicit

'Public Function GetSelect(ByVal query_select As String) As ADODB.Recordset
'    On Error GoTo Handler
'    Dim dbs As ADODB.Connection
'    Set dbs = GetDatabase(GlServidor, GlBaseDatos)
'    If Not dbs Is Nothing Then
'        Dim result As ADODB.Recordset
'        Set result = dbs.Execute(query_select)
'        Set GetSelect = result
'    End If
'    Exit Function
'CleanExit:
'    If Not result Is Nothing Then result.Close
'    If Not dbs Is Nothing And dbs.State = adStateOpen Then
'        dbs.Close
'    End If
'    Exit Function
'Handler:
'    MsgBox ("Select statement error " & Err.Number & " : " & Err.Description)
'    Resume CleanExit
'End Function

'Public Sub ExecProcedure(ByVal query_stored As String)
'    On Error GoTo Handler
'    Dim dbs As ADODB.Connection
'    Set dbs = GetDatabase("192.168.3.131", "CONDOBO")
'
'    If Not dbs Is Nothing Then
'        Dim stored As ADODB.Command
'        Set stored = New ADODB.Command
'        With stored
'            .ActiveConnection = dbs
'            .CommandText = query_stored
'            .CommandType = adCmdText
'            .Execute
'        End With
'    End If
'CleanExit:
'    If Not dbs Is Nothing And dbs.State = adStateOpen Then
'        dbs.Close
'    End If
'    Exit Sub
'Handler:
'    MsgBox ("Stored Procedure error " & Err.Number & " : " & Err.Description)
'    Resume CleanExit
'End Sub

Attribute VB_Name = "ContabilizaPCO"
Public Sub AsientoKFW_TGN(cod As Integer, org As String, gestion As String)
  Dim cmdasiento As ADODB.Command
  Set cmdasiento = New ADODB.Command
  db.BeginTrans
 ' Set cmdasiento = New ADODB.Command
  With cmdasiento
    .ActiveConnection = db
    .CommandType = adCmdStoredProc
    .CommandText = "AsientoKFW_TGN"
    .Parameters("@org") = org
    .Parameters("@cod") = cod
    .Parameters("@gestion") = gestion
    .Parameters("@USR") = GlUsuario
    .Parameters("@HORA") = Format(Time, "hh:mm:ss")
    .Execute
  End With
  db.CommitTrans
End Sub

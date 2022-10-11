Attribute VB_Name = "Module1"
Public fMainForm As frmMain
Public db As Connection

Sub Main()
   Dim fLogin As New frmLogin
   fLogin.Show vbModal
   If Not fLogin.OK Then
      'Fallo al iniciar la sesión, se sale de la aplicación
      MsgBox "Error de Login ..."
      End
   End If

   Unload fLogin

'   frmSplash.Show
'   frmSplash.Refresh
   Set fMainForm = New frmMain
   Load fMainForm
   'Unload frmSplash

''  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=\\sersis\Saf\pragma.mdb;"
  'FrmRegularizacion.Show
  fMainForm.Show
  
End Sub

Public Sub pErrorRst(prmErrores As ADODB.Errors)
   Dim e As ADODB.Error
   
   For Each e In prmErrores
      MsgBox "Error No. " & e.Number & " " & Trim(e.Description)
   Next
   
End Sub

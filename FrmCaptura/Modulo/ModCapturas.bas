Attribute VB_Name = "Module1"
Public fMainForm As frmMain
Public db As Connection
Public rsProg As New ADODB.Recordset
Public modalidad As String
Public VARPassword As String
Global GlUsuario As String
Public GlNombFor As String
' Datos del Tipo de Cambio
Public GlTipoCambioOficial As Currency
Public GlTipoCambioMercado As Currency


Sub Main()
   Dim fLogin As New frmLogin
   GlTipoCambioOficial = 6.2
   GlUsuario = "prueba"
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

' db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=c:\greco\captura\pragma.mdb;"
' db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=\\sersis\saf\udapre.mdb;"
' db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\pragma.mdb;"
'    db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=udapredb;Data Source=sersis"


' db.Open "Provider =SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000;Data Source=sersis"

  db.Open "Provider =SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000PRUEBA;Data Source=sersis"
  MsgBox "BASE DE PRUEBAS"
  
  'FrmRegularizacion.Show
  fMainForm.Show
  
End Sub

Public Sub pErrorRst(prmErrores As ADODB.Errors)
   Dim e As ADODB.Error
   
   For Each e In prmErrores
      MsgBox "Error No. " & e.Number & " " & Trim(e.Description)
   Next
   
End Sub

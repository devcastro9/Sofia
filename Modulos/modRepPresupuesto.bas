Attribute VB_Name = "Module1"
Option Explicit
Public db As New Connection
Public tFc_fuente_financiamiento As New ADODB.Recordset
Public tFc_organismo_financiamiento As New ADODB.Recordset
Public tFc_convenios As New ADODB.Recordset
Public tFc_estructura_programatica As New ADODB.Recordset
Public gl_usuario As String
Public gl_proceso As String

Sub Main()
Dim lineaComandos As String
  Set db = New Connection
  db.CursorLocation = adUseClient
'  db.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000;Data Source=sersis"
  'db.Open "DSN=ODBCSersis2;Description=saf2001;SERVER=sersis;UID=sa;PWD=;WSID=PS10;DATABASE=saf2001;Network=DBMSRPCN"
  db.Open "DSN=ODBCSersis;Description=queiros;SERVER=sersis;UID=sa;PWD=;WSID=PS10;DATABASE=queiros;Network=DBMSRPCN"
  lineaComandos = Command()
  'gl_usuario = GetArg(lineaComandos, 1)
  'gl_proceso = GetArg(lineaComandos, 2)
  Call frmRepPresupuesto.inicio(gl_usuario, gl_proceso)
End Sub
  


Attribute VB_Name = "modcnn"
Option Explicit
Public db As ADODB.Connection
Public glusuario As String
Public gl_nro_sol As Integer
Public tFc_bancos1 As New ADODB.Recordset
Public tFc_bancos2 As New ADODB.Recordset
Public tFc_categoria_financiador As New ADODB.Recordset
Public tpaises As New ADODB.Recordset
Public tFc_convenios As New ADODB.Recordset
Public tCtas As New ADODB.Recordset
Public gl_partida_consultores As String

Sub Main()
Dim lineaComandos As String
  'partida para Consultores
  gl_partida_consultores = "46200"
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2002;Data Source=SERSIS"
  'Obtiene los parametros del programa
  lineaComandos = Command()
  '1er parametro es el usuario
  glusuario = GetArg(lineaComandos, 1)
  frmSoesMain.frmSoesMain_procesar GetArg(lineaComandos, 2)
  'frmmenu.Show
  'frmSoesMain.frmSoesMain_procesar "ABM_SOES"
  'frmSoesMain.frmSoesMain_procesar "DEDUCCIONES"
End Sub


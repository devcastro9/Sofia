Attribute VB_Name = "ModAlmacen"
Option Explicit
'-- API
''**** Variables Públicas
Public Const FormatoFecha As String = "dd/mm/yyyy"
'--
Public rsPrm As ADODB.Recordset
Public Sub BotonesEditar(MiForm As Form)
On Error Resume Next
  MiForm.BtnSalir.Enabled = False
  MiForm.BtnCancelar.Enabled = True
  MiForm.BtnAñadir.Enabled = False
  MiForm.BtnModificar.Enabled = False
  MiForm.BtnEliminar.Enabled = False
  MiForm.BtnGrabar.Enabled = True
  MiForm.BtnBuscar.Enabled = False
  MiForm.BtnImprimir.Enabled = False
  MiForm.CmdRefrescar.Enabled = False
  MiForm.cmdVerificar.Enabled = False
End Sub

Public Sub BotonesNavegar(MiForm As Form)
On Error Resume Next
  MiForm.BtnSalir.Enabled = True
  MiForm.BtnCancelar.Enabled = False
  MiForm.BtnAñadir.Enabled = True
  If GlHayRegs Then
    MiForm.BtnModificar.Enabled = True
    MiForm.BtnEliminar.Enabled = True
    MiForm.BtnBuscar.Enabled = True
    MiForm.BtnImprimir.Enabled = True
    MiForm.cmdVerificar.Enabled = True
  Else
    MiForm.BtnModificar.Enabled = False
    MiForm.BtnEliminar.Enabled = False
    MiForm.BtnBuscar.Enabled = False
    MiForm.BtnImprimir.Enabled = False
    MiForm.cmdVerificar.Enabled = False
  End If
  MiForm.BtnGrabar.Enabled = False
  MiForm.CmdRefrescar.Enabled = True
End Sub
Public Function NombreTerminal() As String
    Dim nPC As String
    Dim buffer As String
    Dim estado As Long
    nPC = "ERROR"
    buffer = String$(255, " ")
    estado = GetComputerName(buffer, 255)
    If estado <> 0 Then
        nPC = Left(buffer, 255)
    End If
    nPC = RTrim(nPC)
    nPC = Mid(nPC, 1, Len(nPC) - 1)
    NombreTerminal = nPC
End Function
Public Function IngresoDeLicitacion(NoLicitacion As Long) As Long
Dim cmm As ADODB.Command
Dim prmNoLici As ADODB.Parameter
Dim prmId As ADODB.Parameter
    Set cmm = New ADODB.Command
    Set prmNoLici = New ADODB.Parameter
    Set prmId = New ADODB.Parameter
    With cmm
        Set prmNoLici = .CreateParameter("NroLicitacion", adInteger, adParamInput, , NoLicitacion)
        .Parameters.Append prmNoLici
        Set prmId = .CreateParameter("QIdIngreso", adInteger, adParamOutput)
        .Parameters.Append prmId
        .CommandType = adCmdStoredProc
        .CommandText = "ALIngresoDeLicitacion"
        .ActiveConnection = db
        .Execute
        IngresoDeLicitacion = prmId.Value
    End With
End Function
Public Function CantidadCajas(CodArt As String, CantEjm As Long) As Long
Dim rs As ADODB.Recordset
    CantidadCajas = 0
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT ISNULL(UnidadCaja,1) As UnidadCaja FROM ALCLDetalle WHERE CodGrupo + '-' + CodDetalle = '" & CodArt & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    If rs.RecordCount <= 0 Then Exit Function
    CantidadCajas = CantEjm / rs!UnidadCaja
End Function

Public Function CantidadEjm(CodArt As String, CantCaja As Long) As Long
Dim rs As ADODB.Recordset
    CantidadEjm = 0
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT ISNULL(UnidadCaja,1) As UnidadCaja FROM ALCLDetalle WHERE CodGrupo + '-' + CodDetalle = '" & CodArt & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    If rs.RecordCount <= 0 Then Exit Function
    CantidadEjm = CantCaja * rs!UnidadCaja
End Function

Attribute VB_Name = "Modmodppto"
'Ini. Variables Globales de Control de Accesos "Freddy
Public GlElegido As String
Public GlMaquina As String
Public GlTipoAcceso As String
Public GlNombreUsuario As String
Public GlUsuario As String

Public rsNivelAcceso As New ADODB.Recordset
Public rsAccesoSistema As New ADODB.Recordset
Public rsPrivAcceso As New ADODB.Recordset
'Ini. Variables Globales de Control de Accesos "Freddy

Public db As New ADODB.Connection
Public ErrLoop As Error
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public GlHayRegs As Boolean
Public Const SW_SHOWNORMAL = 1
'Public fMainForm As frmMain
Public usuario2 As String

' Datos del Tipo de Cambio
Public GlTipoCambioOficial As Currency
Public GlTipoCambioMercado As Currency

'Datos del buscador
Public Vquery As String
Public errCriterio As String
Public SwOrden As Boolean

'Datos Contabilidad
Public diarioFlag As Boolean
Public mayorflag As Boolean
Public conexion1 As ADODB.Connection

Public V_Porden, V_OrgF As Column
Public v_Estado As String

Public recsetAdicion As New ADODB.Recordset
Public ConexionA As New ADODB.Connection
Public ConexionRel As New ADODB.Connection
'Public ConexionComp As New ADODB.Connection
Public GlobErr As ADODB.Error

Public RegDato As Boolean

Public recSetAuxcomp1 As ADODB.Recordset
Public recSetAuxbenefi1 As ADODB.Recordset
Public recSetPartid1 As New ADODB.Recordset

Public recSetOrg As ADODB.Recordset
Public recSetGenera As ADODB.Recordset
Public recSetAuxRel As ADODB.Recordset
Public recsetaux As ADODB.Recordset
Public recSetAuxcomp As ADODB.Recordset
Public recSetPartida As ADODB.Recordset
Public recSetComp As ADODB.Recordset

Public recSetAuxActualizar As ADODB.Recordset
Public recSetAuxActualizar1 As ADODB.Recordset
Public recSetBusqueda As ADODB.Recordset
Public rsRegularizacion As ADODB.Recordset
Public rsDetalle As ADODB.Recordset

Public recSetAuxRe As ADODB.Recordset

Public Cod_Comp As Integer

Public Libroaux As Integer
Public ExistReg As Boolean
Public AUX As String

Public NumComp As Integer
Public NumCbte As String
Public LiteralCry  As String

'***contabilidad manual******
Public Flag_Actualizacion As String
Public d_Aux1 As String
Public Sw_Benefic As Boolean

Public d_Aux2 As String
Public d_Aux3 As String
Public h_Aux1 As String
Public h_Aux2 As String
Public h_Aux3 As String
Public Flag_Asiento
Public Cont_Comp As Long
Public swGrabaCopia As Integer

'CELIA
Public NrosChequeImprimir As String
Public NombreUsuario As String
Public moneda As String 'uno si es boliviano y 2 dolar

'CELIA 'GROVER
Public recSetPartida1 As New ADODB.Recordset

'CELIA 'IMPRESION
'Dim cryCmpte As New CryComprobante


Sub Main()
'   Dim fLogin As New frmLogin
'   fLogin.Show vbModal
'   If Not fLogin.OK Then
'      'Fallo al iniciar la sesión, se sale de la aplicación
'      MsgBox "Error de Login ..."
'      End
'   End If
'
'   Unload fLogin
'
'   Set fMainForm = New frmMain
'   Load fMainForm
   Set db = New Connection
   db.CursorLocation = adUseClient
   db.CommandTimeout = 30
   db.ConnectionTimeout = 15
   
  'saf2000
  db.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000;Data Source=sersis;" '
  
  'saf2000pruebas
  'db.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000prueba;Data Source=sersis;"
   SwOrden = True
 
   FrmModPresup.Show

  
End Sub

Public Sub BotonesHabilitar(Form1 As Form, TipoAcceso As String)
'Esta subrutina habilita o deshabilita los botones de comando
'segun el tipo de acceso que tenga asignado el usuario
On Error Resume Next
If rsPrivAcceso.State = 1 Then rsPrivAcceso.Close
rsPrivAcceso.Open "Select * From PrivilegioAcceso Where IdPrivAcceso='" & TipoAcceso & "'", db, adOpenStatic
If rsPrivAcceso.RecordCount = 1 Then
    Form1.cmdNuevo.Enabled = IIf(rsPrivAcceso!BtnAñadir, True, False)
    Form1.CmdAñadir.Enabled = IIf(rsPrivAcceso!BtnAñadir, True, False)
    Form1.CmdAdicionar.Enabled = IIf(rsPrivAcceso!BtnAñadir, True, False)
    
    Form1.cmdEditar.Enabled = IIf(rsPrivAcceso!BtnModificar, True, False)
    Form1.CmdModificar.Enabled = IIf(rsPrivAcceso!BtnModificar, True, False)
    
    Form1.cmdEliminar.Enabled = IIf(rsPrivAcceso!BtnEliminar, True, False)
    Form1.CmdBorrar.Enabled = IIf(rsPrivAcceso!BtnEliminar, True, False)
    
    Form1.CmdGrabar.Enabled = IIf(rsPrivAcceso!BtnGrabar, True, False)
    Form1.CmdCancelar.Enabled = IIf(rsPrivAcceso!BtnCancelar, True, False)
    
    Form1.CmdBuscar.Enabled = IIf(rsPrivAcceso!BtnBuscar, True, False)
    Form1.CmdBusqueda.Enabled = IIf(rsPrivAcceso!BtnBuscar, True, False)
    
    Form1.CmdImprimir.Enabled = IIf(rsPrivAcceso!BtnImprimir, True, False)
    Form1.cmdVer.Enabled = IIf(rsPrivAcceso!BtnVer, True, False)
    
    Form1.cmdDetalle.Enabled = IIf(rsPrivAcceso!BtnDetalle, True, False)
    Form1.cmdCopiarReg.Enabled = IIf(rsPrivAcceso!BtnCopiarReg, True, False)
    Form1.cmdAprobar.Enabled = IIf(rsPrivAcceso!BtnAprobar, True, False)
Else
    MsgBox "Los privilegios de acceso para este modulo no existen. Revise!", vbInformation + vbOKOnly, "Atención"
End If
End Sub

Public Sub BuscaTipoAcceso(opcMenu As String)
'Esta subrutina tiene el objetivo de encontrar el tipo de acceso
'asignado a la opcion de menu
Dim Encontrado As Boolean
Dim vPosPuntero As Variant
    Encontrado = False
    GlTipoAcceso = ""
    rsNivelAcceso.Requery
    If rsNivelAcceso.RecordCount > 0 Then
        'Guarda la posicion actual del puntero
        vPosPuntero = rsNivelAcceso.Bookmark
        rsNivelAcceso.MoveFirst
        While Not rsNivelAcceso.EOF And Not Encontrado
            If LCase(rsNivelAcceso!NombOpcMenu) = opcMenu Then
                Encontrado = True
                GlTipoAcceso = rsNivelAcceso!IdPrivAcceso
            Else
                rsNivelAcceso.MoveNext
            End If
        Wend
        'Reestablece la posicion del puntero
        rsNivelAcceso.Bookmark = vPosPuntero
    Else
        If GlUsuario = "ADMIN" Then GlTipoAcceso = "TOT" 'Solo deberia tener de ADMinistracion del sistema
    End If
End Sub

Public Sub pErrorRst(prmErrores As ADODB.Errors)
   Dim e As ADODB.Error
   
   For Each e In prmErrores
      MsgBox "Error No. " & e.Number & " " & Trim(e.Description)
   Next
   
End Sub

Public Function ValidaCriterio(v1, v2, v3)
Dim valor As Integer
    valor = 0
    If v1 <> "" Then
        valor = 1
    End If
    If v1 <> "" And v2 <> "" And "'" & v3 & "'" <> "" Then
        valor = 2
    End If
    ValidaCriterio = valor
End Function

Public Function Buscar(atrib1 As String, atrib2 As String, atrib3 As String, atrib4 As String, atrib5 As String, atrib6 As String) As Boolean
    Set recSetBusqueda = New ADODB.Recordset
    recSetBusqueda.CursorLocation = adUseClient
    If recSetBusqueda.State = 1 Then recSetBusqueda.Close
    recSetBusqueda.Open atrib1 & _
    " where   Cod_Trans='" & atrib2 & "' and Org_Codigo='" & atrib3 & "' " & _
    " and Ges_Gestion='" & atrib4 & "' and tipo_comp='" & atrib5 & "' and Cod_Trans_Detalle='" & atrib6 & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
    If recSetBusqueda.RecordCount > 0 Then
    Buscar = True
    Else
    Buscar = False
    End If
End Function

Public Function Buscar_G(Optional atrib1 As String, Optional atrib2 As String, Optional atrib3 As String, Optional atrib4 As String, Optional atrib5 As String, Optional atrib6 As String, Optional atrib7 As String) As Boolean
Set recSetBusqueda = New ADODB.Recordset
recSetBusqueda.CursorLocation = adUseClient
If recSetBusqueda.State = 1 Then recSetBusqueda.Close
recSetBusqueda.Open atrib1 & _
" where   Cuenta='" & atrib2 & "' and SubCta1='" & atrib3 & "' " & _
" and SubCta2='" & atrib4 & "' and Mov<>'" & atrib5 & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
'and Cod_Trans_Detalle='" & atrib6 & "'

If recSetBusqueda.RecordCount > 0 Then
Buscar_G = True
Else
Buscar_G = False
End If

End Function

'LITERAL DE CELIA TARQUINO
Public Function Literal(Cadena As String) As String
Dim sw As Integer
Dim sw1 As Integer
Dim swc As Integer
Dim VEC(20) As Long
sw = 0
      '*********PARTE DECIMAL*********
            Cadena = Round(Cadena, 2)
             X = Len(Cadena)
              For K = 1 To X
                  Z = Mid(Cadena, K, 1)
                  If (Z = ".") Or sw = 1 Then
                    D = D + Mid(Cadena, K, 1)
                    sw = 1
                  End If
              Next K
              
              D = Mid(D, 2, Len(D))
              
              'Para la parte decimal del monto
              If D = "00" Or D = "" Then
                 D = D & " 00/100"
              Else
                 If D >= 0 And D <= 9 And Len(D) = 1 Then
                    D = " " & D & "0" & "/100"
                 Else
                    D = " " & D & "/100 "
                 End If
              End If
      '*********PARTE ENTERA*********
 If Cadena <> "" Then
    Cadena = Int(Cadena)
 Else
    MsgBox "No existe monto"
 End If
   S = ""
   Z = ""
   c = 0
   K = 0
   sw1 = 0
   swc = 0
   
   
   X = Len(Cadena)
   For i = 1 To X
       A = Mid(Cadena, i, 1)
       VEC(i) = Mid(Cadena, i, 1)
   Next i
j = X
While j <> 0
K = K + 1
If K <> 8 Then
  If c <> 3 Then
       c = c + 1
      
       If c = 1 And (VEC(j - 1) <> 1 And VEC(j - 1) <> 2) Then
            Select Case VEC(j)
                Case 0: S = " " + S
                Case 1:
                   If sw1 <> 1 Then
                      S = "UNO " + Z + S
                   End If
                   If sw1 = 1 Then
                      S = "UN " + Z + S
                   End If
                   
                Case 2: S = "DOS " + Z + S
                Case 3: S = "TRES " + Z + S
                Case 4: S = "CUATRO " + Z + S
                Case 5: S = "CINCO " + Z + S
                Case 6: S = "SEIS " + Z + S
                Case 7: S = "SIETE " + Z + S
                Case 8: S = "OCHO " + Z + S
                Case 9: S = "NUEVE " + Z + S
          End Select
          
           'If J + 1 <> "" And sw1 <> 1 And VEC(J - 1) <> 0 And VEC(J) <> 0 Then
           If VEC(j - 1) <> 0 And VEC(j) <> 0 Then
                 S = "Y " + S
           End If
           
        End If
        
         If c = 2 And VEC(j) = 1 Then
               swc = 1
                Select Case VEC(j + 1)
                      Case 0: S = "DIEZ " + Z + S
                      Case 1: S = "ONCE " + Z + S
                      Case 2: S = "DOCE " + Z + S
                      Case 3: S = "TRECE " + Z + S
                      Case 4: S = "CATORCE " + Z + S
                      Case 5: S = "QUINCE " + Z + S
                      Case 6: S = "DIECISEIS " + Z + S
                      Case 7: S = "DIECISIETE " + Z + S
                      Case 8: S = "DIECIOCHO " + Z + S
                      Case 9: S = "DIECINUEVE " + Z + S
                End Select
          End If
          
          If c = 2 And VEC(j) = 2 Then
                Select Case VEC(j + 1)
                      Case 0: S = "VEINTE " + Z + S
                      Case 1: S = "VEINTIUNO " + Z + S
                      Case 2: S = "VEINTIDOS " + Z + S
                      Case 3: S = "VEINTITRES " + Z + S
                      Case 4: S = "VEINTICUATRO " + Z + S
                      Case 5: S = "VEINTICINCO " + Z + S
                      Case 6: S = "VEINTISEIS " + Z + S
                      Case 7: S = "VEINTISIETE " + Z + S
                      Case 8: S = "VEINTIOCHO " + Z + S
                      Case 9: S = "VEINTINUEVE " + Z + S
                End Select
          End If
   
        If c = 2 Then
            Select Case VEC(j)
                Case 3: S = "TREINTA " + Z + S
                Case 4: S = "CUARENTA " + Z + S
                Case 5: S = "CINCUENTA " + Z + S
                Case 6: S = "SESENTA " + Z + S
                Case 7: S = "SETENTA " + Z + S
                Case 8: S = "OCHENTA " + Z + S
                Case 9: S = "NOVENTA " + Z + S
            End Select
            
        End If
        
        If c = 3 Then
            Select Case VEC(j)
                Case 1:
                If j = 1 Then
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       S = "CIEN " + Z + S
                    Else
                       S = "CIENTO " + Z + S
                    End If
                Else
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       S = "CIEN " + Z + S
                    Else
                       S = "CIENTO " + Z + S
                    End If
                       'S = "CIENTO " + z + S
                End If
                Case 2: S = "DOSCIENTOS " + Z + S
                Case 3: S = "TRESCIENTOS " + Z + S
                Case 4: S = "CUATROCIENTOS " + Z + S
                Case 5: S = "QUINIENTOS " + Z + S
                Case 6: S = "SEISCIENTOS " + Z + S
                Case 7: S = "SETECIENTOS " + Z + S
                Case 8: S = "OCHOCIENTOS " + Z + S
                Case 9: S = "NOVECIENTOS " + Z + S
            End Select
        End If
   Else
     If j >= 3 Then
            If VEC(j) = 0 And VEC(j - 1) = 0 And VEC(j - 2) = 0 Then
            Else
              S = "MIL " + S
            End If
    Else
              S = "MIL " + S
    End If
        j = j + 1
        c = 0
        sw1 = 1
   End If
 Else
    If VEC(j) <> 1 Then
      S = "MILLONES " + S
    Else
'      If K > 7 Then
      If K <> 8 Then
        S = "MILLONES " + S
      Else
        S = "MILLON " + S
      End If
    End If
      j = j + 1
      c = 0
      sw1 = 1
 End If
   j = j - 1
   
Wend

Literal = S + D
End Function

Public Function Encriptar(Cadena As String) As String
Dim i, j As Byte

For i = 1 To Len(Cadena)
  j = Asc(Mid(Cadena, i, 1)) + 5
  Encriptar = Encriptar & Chr(j)
Next i
End Function

Public Function Desencriptar(Cadena As String) As String
Dim i, j As Byte

For i = 1 To Len(Cadena)
  j = Asc(Mid(Cadena, i, 1)) - 5
  Desencriptar = Desencriptar & Chr(j)
Next i
End Function

Public Function ControlErrores(Origen As String) As String
Dim AntError As Long
Dim Encontro As Boolean
  
  AntError = 0
  Encontro = False
  For Each ErrLoop In db.Errors
    If AntError <> ErrLoop.Number Then
      Encontro = True
      Select Case ErrLoop.Number
        Case -2147217900
              MsgBox "Error #" & ErrLoop.Number & vbCr & _
                 "   " & ErrLoop.Description & vbCr & _
                 "   (Source: " & ErrLoop.Source & ")" & vbCr & _
                 "   (SQL State: " & ErrLoop.SQLState & ")" & vbCr & _
                 "   (NativeError: " & ErrLoop.NativeError & ")", vbCritical + vbOKOnly, Origen
              ControlErrores = ""
        Case -2147217864
              MsgBox "Error #" & ErrLoop.Number & vbCr & _
                 "   " & ErrLoop.Description & vbCr & _
                 "   (Source: " & ErrLoop.Source & ")" & vbCr & _
                 "   (SQL State: " & ErrLoop.SQLState & ")" & vbCr & _
                 "   (NativeError: " & ErrLoop.NativeError & ")", vbCritical + vbOKOnly, Origen
              ControlErrores = ""
        Case -2147467259
            MsgBox "Error #" & ErrLoop.Number & vbCr & _
               "   " & ErrLoop.Description & vbCr & _
               "   (Source: " & ErrLoop.Source & ")" & vbCr & _
               "   (SQL State: " & ErrLoop.SQLState & ")" & vbCr & _
               "   (NativeError: " & ErrLoop.NativeError & ")", vbCritical + vbOKOnly, Origen
            ControlErrores = ""
        Case Else
          MsgBox "Error #" & ErrLoop.Number & vbCr & _
             "   " & ErrLoop.Description & vbCr & _
             "   (Source: " & ErrLoop.Source & ")" & vbCr & _
             "   (SQL State: " & ErrLoop.SQLState & ")" & vbCr & _
             "   (NativeError: " & ErrLoop.NativeError & ")", vbCritical + vbOKOnly, Origen
             ControlErrores = ""
      End Select
      AntError = ErrLoop.Number
    End If
  Next
  If (Not Encontro) And (Err.Number <> 0) Then
    MsgBox "Error: " & Err.Number & "; " & Err.Description, vbCritical + vbOKOnly, "Atención"
  End If
End Function



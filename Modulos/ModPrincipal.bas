Attribute VB_Name = "ModPrincipal"
'Ini. Variables Globales de Control de Accesos
Public GlElegido As String
Public GlMaquina As String          'Nombre del equipo con el que se trabaja
Public GlTipoAcceso As String
Public GlNombreUsuario As String    'Nombre del Logeado
Public glusuario As String          'Login
Public GlIdFuncionario As Integer   'Id del logeado
Public GlFechaProceso As Date       'Fecha de registro en el momento del logeo
Public GlParametro As String        'Parametro Sobre el que se registran Correlativos (NIT de la Entidad)
Public GlParametroDes As String     'Descripcion Parametro (Nombre o Razon Social de la Entidad)
Public GlServidor As String         'Nombre del servidor
Public GlBaseDatos As String        'Nombre de la Base de Datos
Public glGestion As String          'Gestión actual con la que se trabaja
Public glPassword As String         'Password del Usuario
Public GLCarpeta As String          'Carpeta donde se instala el Sistema
Public GLCarpeta2 As String         'Carpeta que Contienen Datos (.DBF y otros)
Public glProceso$

Public GlEdificio As String         'Codigo de Edificacion
Public GlUnidad As String           'Codigo de Unidad Ejecutora
Public GlSolicitud As Integer        'Codigo de Solicitud

Public GlHora1 As String
Public GlHora2 As String

Public rsNivelAcceso As New ADODB.Recordset
Public rsAccesoSistema As New ADODB.Recordset
Public rsPrivAcceso As New ADODB.Recordset
Public cnnString As String
Public iResult As Integer

Public GlSistema As String
Public gestion As String    'de n
Public nro_licitacion As Long   'de n
Public idBeneficiario As Integer 'de n
Public GlSqlAux As String                'Para los querys

'Ini. Variables Globales de Control de Accesos
Public db As New ADODB.Connection
Public cnn As New ADODB.Connection
Public db2 As New ADODB.Connection

Public ErrLoop As Error
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'*** Nombre de la computadora
Public Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Public GlHayRegs As Boolean
Public Const SW_SHOWNORMAL = 1
Public fMainForm As frmMain
Public usuario2 As String

' Datos del Tipo de Cambio
Public GlTipoCambioOficial As Currency  'Compra Dolar
Public GlTipoCambioMercado As Currency  'Venta Dolar
Public GlTipoCambioGestion As Currency  'Para cierre Dolar
Public GlTipoCambioEuro As Currency     'Euros España y otros
Public GlTipoCambioUfv As Currency      'UFV
Public GlTipoCambioRmb As Currency      'Reminbis China
Public GlTipoCambioBrl As Currency      'Reales Brasil

'Datos del buscador
Public Vquery As String
Public errCriterio As String
Public SwOrden As Boolean
Public swnuevo As Integer
Public queryinicial As String
Public queryinicial99 As String
Public buscados As Integer
Public GLREFRESH As Integer

'Buscador
Public PQConexion As New ADODB.Connection
Public rsTablaAux As New ADODB.Recordset
Public ElTDBGridAux As TrueOleDBGrid60.TDBGrid
Public ElGridAux As MSDataGridLib.DataGrid
Public ClBuscaGrid As New ClBuscaEnGridExterno

' datos de solicitudes (captura) g-
Public GlNombFor As String
Public Glaux As String
Public GlPuesto As String
Public GlExtension As String
Public VAR_DET As String

'Datos Contabilidad
Public diarioFlag As Boolean
Public mayorflag As Boolean
Public conexion1 As ADODB.Connection

'Public V_Porden, V_OrgF As Column
Public v_Estado As String
Public GlSW As String

Dim rs_PARAMETRO As New ADODB.Recordset
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
Public rsdetalle As ADODB.Recordset

Public recSetAuxRe As ADODB.Recordset

Public Cod_Comp As Long
Public NumComp As Long
Public Libroaux As Integer
Public GlCotiza As Integer

Public ExistReg As Boolean
Public Aux As String
Public parametro As String

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
'*** VARIABLES DE REPORTES ***
Public glRepPresup As String

'TESORERIA
Public NrosChequeImprimir As String
Public NombreUsuario As String
Public moneda As String 'uno si es boliviano y 2 dolar
Public recSetPartida1 As New ADODB.Recordset

'Compras
Public GldaCodigo, GldaDescrip As String
Public rstTemp As New ADODB.Recordset

Public glPersNew, GlArch As String
Public glBenef As String
Public VAR_PAISC As String
Public GlConti As String
Public VAR_TIPOC As String

'VENTAS
Public VAR_FLE, VAR_NAC, VAR_ALM, VAR_AGE, VAR_UTIL As Double

'Ppto
Public tFc_fuente_financiamiento As New ADODB.Recordset
Public tFc_organismo_financiamiento As New ADODB.Recordset
Public tFc_convenios As New ADODB.Recordset
Public tFc_estructura_programatica As New ADODB.Recordset
Public gl_usuario As String
Public gl_proceso As String

'Variables importantes para contabilizacion
Public nro_operacion As Integer
Public cod1 As Long
Public cod2 As Long

Public txtVersion As String

Public Function meses(nMes As Integer) As String
'funcion que devuelve el mes
'EMPLEADO EN CONSULTORIA
    Select Case nMes
        Case 1
            meses = "ENERO"
        Case 2
            meses = "FEBRERO"
        Case 3
            meses = "MARZO"
        Case 4
            meses = "ABRIL"
        Case 5
            meses = "MAYO"
        Case 6
            meses = "JUNIO"
        Case 7
            meses = "JULIO"
        Case 8
            meses = "AGOSTO"
        Case 9
            meses = "SEPTIEMBRE"
        Case 10
            meses = "OCTUBRE"
        Case 11
            meses = "NOVIEMBRE"
        Case 12
            meses = "DICIEMBRE"
        Case Else
            meses = "NO IDENTIFICADO"
    End Select
End Function

Public Sub Main()
    ' Conexion a Access
    Set cnn = New Connection
    With cnn
        .CursorLocation = adUseClient
        .CommandTimeout = 30
        .ConnectionTimeout = 15
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = App.Path & "\PARAMETROS.mdb"
        .Open
    End With
    ' Carga de parametros de sistema
    Call BUSCA_PARAMETRO
    ' Conexion a la base de datos Sql Server
    Set db = GetDatabase(GlServidor, GlBaseDatos)
    ' Variables
    GlFechaProceso = ObtenerFechaServidor()
    'GlFechaProceso = Date
    SwOrden = True
    ' Login
    frmLogin.Show
    ' ------------------------------------
    ' Codigo comentado
    ' ------------------------------------
   'INI Nueva conexion   - Initialize variables.
    'GlServidor = "SERVIDOR"
    'GlBaseDatos = "ADMIN_EMPRESA"
    
'    glPassword = "Servidor2020*"
'    Set db = New Connection
'    db.CursorLocation = adUseClient
'    db.CommandTimeout = 30
'    db.ConnectionTimeout = 15
       
       'db.Open "Provider=SQLOLEDB.1;Data Source=192.168.3.133;Initial Catalog=ADMIN_EMPRESA;User ID=sa;Password=Servidor2020*"
    'db.Open "  Provider=MSDASQL.1;Persist Security Info=False;User ID=sa;Data Source=Odbc_Admin_Empresa;Initial Catalog=ADMIN_EMPRESA;Password=Servidor2020*"
    'db.Open "Provider=SQLOLEDB.1;Password=Servidor2020*;Persist Security Info=True;User ID=sa;Initial Catalog=ADMIN_EMPRESA;Data Source=SSSOFIA"
    'db.Open "Provider=SQLOLEDB.1;Password=Servidor2020*;Persist Security Info=True;User ID=sa;Initial Catalog=ADMIN_EMPRESA;Data Source=192.168.3.141"
    'db.Open "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;Initial Catalog=ADMIN_EMPRESA;Data Source=SSSOFIA"
    
    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ADMIN_EMPRESA;Data Source=SSSOFIA
    'db.Provider = "SQLOLEDB.1"        ' Specifica proveedor: OLE DB .
'    ' Asignar propiedades de conexion SQLOLEDB .
    'db.Properties("Data Source").Value = GlServidor                             'OK
    'db.Properties("Initial Catalog").Value = GlBaseDatos                        'OK
    ' Decicion para tipo de autorizacion de logeo:  Windows NT  o  SQL Server .
    'If optWinNTAuth.Value = True Then
        '   Autentificacion por Windows
    '   db.Properties("Integrated Security").Value = "SSPI"
    'Else
        '   Autentificacion Mixta (Windows y SQL-Server)
        'db.Properties("User ID").Value = glusuario
        'db.Properties("User ID").Value = "sa"
        'db.Properties("Password").Value = "Servidor2020*"    'glPassword
    'End If
    'db.Open     ' Abre conexion de la BD.
'   'FIN Nueva conexion
'   'db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ADMIN_EMPRESA;Data Source=SERVIDOR"
    'CONDOBO
    'db2.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CONDOBO;Data Source=SSOFIA"

   ''db.Open "Provider=Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Password=Servidor2020*;Initial Catalog=ADMIN_EMPRESA;Data Source=192.168.3.133;Current Language=us_english"
   'SwOrden = True
   ' Cambia la Configuración regional del equipo
'   CambiarCR
   'Busca el Parametro para Correlativos JQA  02/02/2007
   'Call BUSCA_PARAMETRO
    'MsgBox "no comentar ALMACENES"
  '-- Parametros de Almacen  *** NO COLOCAR EN COMENTARIO
  'GlServidor = "SERVIDOR"
  'GlBaseDatos = "Prueba"
'  GlSqlAux = "SELECT * FROM ALParametros"
'  Set rsPrm = New ADODB.Recordset
'  rsPrm.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
'  If rsPrm.RecordCount <= 0 Then
'      MsgBox "OOPs!!! No se definieron los Parámetros del sistema de Almacenes.", vbCritical + vbOKOnly, "Cerrando el Sistema"
'      End
'  End If
   'frmLogin.Show
End Sub

Public Function GetDatabase(ByVal server As String, ByVal database As String) As ADODB.Connection
    On Error GoTo Handler
    Dim dbase As ADODB.Connection
    Set dbase = New ADODB.Connection
    With dbase
        .CursorLocation = adUseClient
        .CommandTimeout = 30
        .ConnectionTimeout = 5
        .Provider = "SQLOLEDB.1"
        .Properties("Data Source").Value = server
        .Properties("Initial Catalog").Value = database
        .Properties("Persist Security Info").Value = False
        .Properties("Integrated Security").Value = "SSPI"
        .Open
    End With
    Set GetDatabase = dbase
    Exit Function
CleanExit:
    If Not dbase Is Nothing And dbase.State = adStateOpen Then
        dbase.Close
    End If
    Exit Function
Handler:
    MsgBox "Database error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Function

' Version valida
Public Function EsVersionValida(ByVal Usuario As String) As Boolean
    On Error GoTo Handler:
    ' ADDODB
    Dim rsUsuario As New ADODB.Recordset
    Dim rsVersion As New ADODB.Recordset
    ' Queries
    Dim query_usuario As String
    Dim query_version As String
    ' Variables
    Dim EsValidar As Boolean
    Dim VersionMayor As Integer
    Dim VersionMenor As Integer
    Dim VersionRevision As Integer
    ' SQL
    query_usuario = "SELECT [u].[validar_version] FROM [dbo].[gc_usuarios] AS [u] WHERE [u].[usr_codigo] = '" & Usuario & "'"
    query_version = "SELECT [s].[version_principal], [s].[version_secundaria], [s].[version_revision] FROM [dbo].[gc_parametros_sistema] AS [s] WHERE [s].[param_codigo] = '1018533029' AND [s].[ges_gestion] = '" & glGestion & "'"
    ' Otros
    EsValidar = True
    ' Apertura de ADDODB
    rsUsuario.Open query_usuario, db, adOpenStatic, adLockReadOnly
    ' Ve si el usuario existe y se puede validar
    If rsUsuario.RecordCount = 0 Then
        EsVersionValida = False
        Exit Function
    End If
    EsValidar = rsUsuario!validar_version
    If EsValidar Then
        rsVersion.Open query_version, db, adOpenStatic, adLockReadOnly
        VersionMayor = rsVersion!version_principal
        VersionMenor = rsVersion!version_secundaria
        VersionRevision = rsVersion!version_revision
        EsVersionValida = App.Major >= VersionMayor And App.Minor >= VersionMenor And App.Revision >= VersionRevision
        ' Objetos ADDODB
        If rsVersion.Status = adStateOpen Then rsVersion.Close
    Else
        EsVersionValida = True
    End If
    ' Objetos ADDODB
    If rsUsuario.Status = adStateOpen Then rsUsuario.Close
    Exit Function
Handler:
    MsgBox "Error en Version: " & Err.Number & " : " & Err.Description
End Function

' Obtiene la fecha del Servidor mediante SQL SERVER
Public Function ObtenerFechaServidor() As Date
    Dim rsFecha As New ADODB.Recordset
    Dim Fecha As Date
    Dim query_fecha As String
    query_fecha = "SELECT CONVERT(DATE, GETDATE()) AS 'FechaServidor'"
    rsFecha.Open query_fecha, db, adOpenStatic, adLockReadOnly
    ObtenerFechaServidor = rsFecha!FechaServidor
    rsFecha.Close
End Function

Private Sub BUSCA_PARAMETRO()
    Set rs_PARAMETRO = New ADODB.Recordset
    rs_PARAMETRO.Open "select * from gc_parametros_sistema where estado_registro = 'S' ", cnn, adOpenDynamic, adLockReadOnly
    If rs_PARAMETRO.RecordCount > 0 Then
        rs_PARAMETRO.MoveFirst
        GlParametro = IIf(IsNull(rs_PARAMETRO!param_codigo), "1018533029", rs_PARAMETRO!param_codigo)
        GlParametroDes = IIf(IsNull(rs_PARAMETRO!param_descripcion), "EMPRESA", rs_PARAMETRO!param_descripcion)
        GlServidor = IIf(IsNull(rs_PARAMETRO!nombre_servidor), "SERVIDOR", rs_PARAMETRO!nombre_servidor)
        GlBaseDatos = IIf(IsNull(rs_PARAMETRO!nombre_BD), "ADMIN_EMPRESA", rs_PARAMETRO!nombre_BD)
        glGestion = IIf(IsNull(rs_PARAMETRO!idGestion), "2022", rs_PARAMETRO!idGestion)
        glDepto = IIf(IsNull(rs_PARAMETRO!CodDepto), "2", rs_PARAMETRO!CodDepto)
        glProvi = IIf(IsNull(rs_PARAMETRO!codprovi), "201", rs_PARAMETRO!codprovi)
        glMunic = IIf(IsNull(rs_PARAMETRO!codmunicip), "20101", rs_PARAMETRO!codmunicip)
        glcomuni = IIf(IsNull(rs_PARAMETRO!codcomunidad), "2010101", rs_PARAMETRO!codcomunidad)
        GLCarpeta2 = "PERSONAL"
    Else
        MsgBox "No exixte un servidor habilitado en los parametros...", vbCritical
        End
    End If
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
    Form1.BtnAñadir.Enabled = IIf(rsPrivAcceso!BtnAñadir, True, False)
    
    Form1.BtnModificar.Enabled = IIf(rsPrivAcceso!BtnModificar, True, False)
    Form1.CmdModificar.Enabled = IIf(rsPrivAcceso!BtnModificar, True, False)
    
    Form1.cmdEliminar.Enabled = IIf(rsPrivAcceso!BtnEliminar, True, False)
    Form1.BtnEliminar.Enabled = IIf(rsPrivAcceso!BtnEliminar, True, False)
    
    Form1.BtnGrabar.Enabled = IIf(rsPrivAcceso!BtnGrabar, True, False)
    Form1.BtnCancelar.Enabled = IIf(rsPrivAcceso!BtnCancelar, True, False)
    
    Form1.BtnBuscar.Enabled = IIf(rsPrivAcceso!BtnBuscar, True, False)
    Form1.BtnBuscarA.Enabled = IIf(rsPrivAcceso!BtnBuscar, True, False)
    
    Form1.BtnImprimir.Enabled = IIf(rsPrivAcceso!BtnImprimir, True, False)
    Form1.cmdVer.Enabled = IIf(rsPrivAcceso!BtnVer, True, False)
    
    Form1.BtnAceptar.Enabled = IIf(rsPrivAcceso!BtnVer, True, False)
    
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
        If glusuario = "ADMIN" Then GlTipoAcceso = "TOT" 'Solo deberia tener de ADMinistracion del sistema
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

'LITERAL DE CC -
Public Function Literal(Cadena As String) As String
Dim SW As Integer
Dim sw1 As Integer
Dim swc As Integer
Dim VEC(20) As Long
SW = 0
      '*********PARTE DECIMAL*********
            If Cadena < 0 Then Cadena = Cadena * (-1)
            Cadena = Round(Cadena, 2)
             x = Len(Cadena)
              For k = 1 To x
                  Z = Mid(Cadena, k, 1)
                  If (Z = ".") Or SW = 1 Then
                    d = d + Mid(Cadena, k, 1)
                    SW = 1
                  End If
              Next k
              
              d = Mid(d, 2, Len(d))
              
              'Para la parte decimal del monto
              If d = "00" Or d = "" Then
                 d = d & " 00/100"
              Else
                 If d >= 0 And d <= 9 And Len(d) = 1 Then
                    d = " " & d & "0" & "/100"
                 Else
                    d = " " & d & "/100 "
                 End If
              End If
      '*********PARTE ENTERA*********
 If Cadena <> "" Then
    Cadena = Int(Cadena)
 Else
    MsgBox "No existe monto"
 End If
   s = ""
   Z = ""
   c = 0
   k = 0
   sw1 = 0
   swc = 0
   
   
   x = Len(Cadena)
   For i = 1 To x
       a = Mid(Cadena, i, 1)
       VEC(i) = Mid(Cadena, i, 1)
   Next i
j = x
While j <> 0
k = k + 1
If k <> 8 Then
  If c <> 3 Then
       c = c + 1
      
       If c = 1 And (VEC(j - 1) <> 1 And VEC(j - 1) <> 2) Then
            Select Case VEC(j)
                Case 0: s = " " + s
                Case 1:
                   If sw1 <> 1 Then
                      s = "UNO " + Z + s
                   End If
                   If sw1 = 1 Then
                      s = "UN " + Z + s
                   End If
                   
                Case 2: s = "DOS " + Z + s
                Case 3: s = "TRES " + Z + s
                Case 4: s = "CUATRO " + Z + s
                Case 5: s = "CINCO " + Z + s
                Case 6: s = "SEIS " + Z + s
                Case 7: s = "SIETE " + Z + s
                Case 8: s = "OCHO " + Z + s
                Case 9: s = "NUEVE " + Z + s
          End Select
          
           'If J + 1 <> "" And sw1 <> 1 And VEC(J - 1) <> 0 And VEC(J) <> 0 Then
           If VEC(j - 1) <> 0 And VEC(j) <> 0 Then
                 s = "Y " + s
           End If
           
        End If
        
         If c = 2 And VEC(j) = 1 Then
               swc = 1
                Select Case VEC(j + 1)
                      Case 0: s = "DIEZ " + Z + s
                      Case 1: s = "ONCE " + Z + s
                      Case 2: s = "DOCE " + Z + s
                      Case 3: s = "TRECE " + Z + s
                      Case 4: s = "CATORCE " + Z + s
                      Case 5: s = "QUINCE " + Z + s
                      Case 6: s = "DIECISEIS " + Z + s
                      Case 7: s = "DIECISIETE " + Z + s
                      Case 8: s = "DIECIOCHO " + Z + s
                      Case 9: s = "DIECINUEVE " + Z + s
                End Select
          End If
          
          If c = 2 And VEC(j) = 2 Then
                Select Case VEC(j + 1)
                      Case 0: s = "VEINTE " + Z + s
                      Case 1: s = "VEINTIUNO " + Z + s
                      Case 2: s = "VEINTIDOS " + Z + s
                      Case 3: s = "VEINTITRES " + Z + s
                      Case 4: s = "VEINTICUATRO " + Z + s
                      Case 5: s = "VEINTICINCO " + Z + s
                      Case 6: s = "VEINTISEIS " + Z + s
                      Case 7: s = "VEINTISIETE " + Z + s
                      Case 8: s = "VEINTIOCHO " + Z + s
                      Case 9: s = "VEINTINUEVE " + Z + s
                End Select
          End If
   
        If c = 2 Then
            Select Case VEC(j)
                Case 3: s = "TREINTA " + Z + s
                Case 4: s = "CUARENTA " + Z + s
                Case 5: s = "CINCUENTA " + Z + s
                Case 6: s = "SESENTA " + Z + s
                Case 7: s = "SETENTA " + Z + s
                Case 8: s = "OCHENTA " + Z + s
                Case 9: s = "NOVENTA " + Z + s
            End Select
            
        End If
        
        If c = 3 Then
            Select Case VEC(j)
                Case 1:
                If j = 1 Then
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                Else
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                       'S = "CIENTO " + z + S
                End If
                Case 2: s = "DOSCIENTOS " + Z + s
                Case 3: s = "TRESCIENTOS " + Z + s
                Case 4: s = "CUATROCIENTOS " + Z + s
                Case 5: s = "QUINIENTOS " + Z + s
                Case 6: s = "SEISCIENTOS " + Z + s
                Case 7: s = "SETECIENTOS " + Z + s
                Case 8: s = "OCHOCIENTOS " + Z + s
                Case 9: s = "NOVECIENTOS " + Z + s
            End Select
        End If
   Else
     If j >= 3 Then
            If VEC(j) = 0 And VEC(j - 1) = 0 And VEC(j - 2) = 0 Then
            Else
              s = "MIL " + s
            End If
    Else
              s = "MIL " + s
    End If
        j = j + 1
        c = 0
        sw1 = 1
   End If
 Else
    If VEC(j) <> 1 Then
      s = "MILLONES " + s
    Else
'      If K > 7 Then
      If k <> 8 Then
        s = "MILLONES " + s
      Else
        s = "MILLON " + s
      End If
    End If
      j = j + 1
      c = 0
      sw1 = 1
 End If
   j = j - 1
   
Wend

Literal = s + d
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

' ==== función para consultoría  (ema)
Public Function fHoraValida(xHora As String) As Boolean
Dim h%, m%
h = Val(Mid(xHora, 1, 2))
m = Val(Mid(xHora, 4, 2))
If h >= 0 And h <= 24 Then
    If m >= 0 And m <= 60 Then
        fHoraValida = True
    Else
        fHoraValida = False
    End If
Else
    fHoraValida = False
End If
End Function


Public Function CalcTiempo(fecha1 As Date, fecha2 As Date) As String
'*******************************************************************
'**  Función que devuelve el literal de diferencia entre dos fechas
'**  HECHA POR : Dulfredo Rojas
'**  CORTESIA A: Freddy Quiroz (Tren Quiroz)
'**  FECHA     : 9 de Febrero del 2001
'*******************************************************************
On Error Resume Next
Dim año As Integer, mes As Byte, dia As Byte
Dim año1 As Integer, mes1 As Byte, dia1 As Byte
Dim año2 As Integer, mes2 As Byte, dia2 As Byte
Dim mesSiguiente As String
  'Valida
  If fecha1 > fecha2 Then
    MsgBox "La fecha inicial es mayor a la Final", vbCritical + vbOKOnly, "Error"
    CalcTiempo = ""
  Else
    año1 = Year(fecha1): año2 = Year(fecha2)
    mes1 = Month(fecha1): mes2 = Month(fecha2)
    dia1 = Day(fecha1):  dia2 = Day(fecha2)
    'Calcula los años
    año = año2 - año1
    'Calcula los meses
    If mes1 > mes2 Then
      año = año - 1
      mes = (12 - mes1) + mes2
    Else
      mes = mes2 - mes1
    End If
    'Calcula los dias
    If dia1 > dia2 Then
      If mes = 0 Then
        mes = 11
        año = año - 1
      Else
        mes = mes - 1
      End If
      mesSiguiente = (mes1 + mes + 1) Mod 12
      If mesSiguiente < mes1 Then
        mesSiguiente = "1/" & mesSiguiente & "/" & (año1 + año) + 1
      Else
        mesSiguiente = "1/" & mesSiguiente & "/" & (año1 + año)
      End If
      dia = ((Day(DateAdd("d", -1, mesSiguiente))) - dia1) + dia2
    Else
      dia = dia2 - dia1
    End If
    'Resultado
    If año > 0 Then
      CalcTiempo = año
      If año = 1 Then CalcTiempo = CalcTiempo & " año, " Else CalcTiempo = CalcTiempo & " años, "
    End If
    If mes > 0 Then
      CalcTiempo = CalcTiempo & mes
      If mes = 1 Then CalcTiempo = CalcTiempo & " mes, " Else CalcTiempo = CalcTiempo & " meses, "
    End If
    CalcTiempo = CalcTiempo & dia
    If dia = 1 Then CalcTiempo = CalcTiempo & " dia. " Else CalcTiempo = CalcTiempo & " dias."
  End If
End Function

Public Function pg_ReemplazaCarater(CadOrigen As String, CaracBuscar As String, CaracReempl As String) As String
    'AUTOR     : René Roque Mendoza.
    'PROPÓSITO : Remplazada en CadOrigen un caracterer por otro.
    
    Dim CadAux As String
    Dim i As Integer
    
    CadAux = ""
    For i = 1 To Len(CadOrigen)
        If Mid(CadOrigen, i, 1) = CaracBuscar Then ' si es el caracter a reemoplazar
            CadAux = CadAux & CaracReempl
          Else
            CadAux = CadAux & Mid(CadOrigen, i, 1)
        End If
    Next i
    pg_ReemplazaCarater = CadAux
End Function
'' procedimientos que se deben adicionar al modulo de administracion
'' nuevos procedimientos q pueden ser utulizados para optimizar los registros
''*****************/
Public Function pg_QuitaEspBlanco(cade As String) As String
    'AUTOR     : René Roque Mendoza.
    'PROPÓSITO : De cade suprime los espacios extras en blanco que podria tener y
               ' solo coloca un espacio en blanco entre palabra
    
    Dim CadenaAux As String, CadenaCopia As String
    
    CadenaAux = Trim(cade) & " " ' el ultimo caracter es de control
    CadenaCopia = ""
    If Len(CadenaAux) = 1 Then
        pg_QuitaEspBlanco = ""
        Exit Function
    End If
    
    Do While Len(CadenaAux) > 0
        CadenaCopia = LTrim(CadenaCopia & " " & Mid(CadenaAux, 1, InStr(CadenaAux, " ") - 1))
        CadenaAux = LTrim(Mid(CadenaAux, InStr(CadenaAux, " ")))
    Loop
    pg_QuitaEspBlanco = CadenaCopia

End Function

Public Function pg_BuscaTdbGrid(Grid As TDBGrid, rs_puntero As ADODB.Recordset, ColBuscar As String)
    'AUTOR          : René Roque Mendoza.
    'MODIFICADO POR : Maria Luisa Gonzales Mendoza.
    'MODIFICADO POR : Adett Grover Cruz M.
    'PROPÓSITO      : Realiza la busqueda de una especificación sobre la columna de la celda activa
                    ' Grid -> nombre del TDBGrid sobre la cual se ejecuta la operación de busqueda
                    ' rs_puntero -> recorset del grid
                    ' ColBuscar -> es la columna donde se hace la busqueda
    
    Dim CadBuscar As String
    Dim micriterio As String
    Dim CampoAct As Integer
    Dim a As Integer
    
    If rs_puntero.RecordCount = 0 Then Exit Function
    micriterio = "Digite " & LCase(Grid.Columns(Grid.Col).Caption) & " a buscar"
    CadBuscar = pg_QuitaEspBlanco(UCase(InputBox(micriterio, "Búsqueda")))
    ' verificamos que la cadena sea del tipo *a* donde a representa cualquier secuencia de caracteres
    Select Case Len(CadBuscar)
      Case Is >= 3
        If Left(CadBuscar, 1) = "*" And Right(CadBuscar, 1) <> "*" Then
            ' completamos la cadena al tipo *a*
            CadBuscar = CadBuscar & "*"
          Else
            ' es del tipo a* o a que son cadenas validas
            'CadBuscar = "*" & CadBuscar
        End If
      Case 2
        If Left(CadBuscar, 1) = "*" And Right(CadBuscar, 1) <> "*" Then
            CadBuscar = CadBuscar & "*"
          Else
            If Left(CadBuscar, 1) = "*" And Right(CadBuscar, 1) = "*" Then
                ' si ambos son *
                CadBuscar = ""
              Else
                ' es del tipo a* o a que son cadenas validas
                'CadBuscar = "*" & CadBuscar
            End If
        End If
      Case 1
        If CadBuscar = "*" Then CadBuscar = ""
    End Select

    
    On Error GoTo EtiqError
    If Len(CadBuscar) > 0 Then ' si introdujo una cadena a buscar
        CampoAct = Grid.Col
        
''        rs_puntero.Sort = ColBuscar ' se cancela ordenar para mantener el orden anterior
        
        Grid.Col = CampoAct
        micriterio = ColBuscar & " like " & Chr(39) & CadBuscar & Chr(39)
        ' verificamos si la longitud coincide con el tamaño del campo
        If Len(CadBuscar) <= rs_puntero.Fields(ColBuscar).DefinedSize Then
''            rs_puntero.MoveFirst ' se cancela ir al primer registro para poder buscar el siguiente
            rs_puntero.Find micriterio
            If Not rs_puntero.EOF Then
                Grid.MarqueeStyle = 4 'marca toda la fila
                Grid.SetFocus
              Else ' si no lo encontro
                Grid.MoveFirst
                MsgBox "No se encontró " & Grid.Columns(Grid.Col).Caption & " --> " & CadBuscar, vbInformation, "Información"
'                Grid.MoveFirst
            End If
          Else ' la longitud de la cadena a buscar es mayor a la longitud del campo
            MsgBox "No se encontró " & Grid.Columns(Grid.Col).Caption & " --> " & CadBuscar, vbInformation, "Información"
        End If
        'Grid.SetFocus
      Else ' solo tiene el foco
        Grid.MarqueeStyle = 6
        Grid.SetFocus
    End If
    On Error GoTo 0 ' desactiva el manejador de errores
    Exit Function
    
EtiqError:
    Select Case Err.Number
      Case -2147217881
        MsgBox "No se puede realizar la búsqueda, los tipos no coinciden" & Chr(13) & "debe buscar valores puntuales (sin el caracter *).", vbInformation, "Aviso"
      Case -2147217887
        MsgBox "No se encontró " & Grid.Columns(Grid.Col).Caption & " --> " & CadBuscar, vbInformation, "Información"
        Grid.MoveFirst
      Case Else ' si se produjo algun error
        MsgBox "Los cambios no se llevaron a cabo." & Chr(13) & "Anote el error y comuniquese con el soporte técnico." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description, vbCritical, "Error"
    End Select
    Exit Function
End Function

Public Function pg_Imprimir(Grid As TDBGrid, Titulo As String)
    
    'AUTOR     : Maria Luisa Gonzales Mendoza.
    'PROPÓSITO : Imprime lo que se tiene en el TDBgrid
               ' Grid   --> nombre del grid que sa va imprimir
               ' Titulo --> Módulo que invoca la impresión (Módulo de almacén,de producción o de distribución)
    'LLAMADA   : pg_Imprimir(Grid, "Titulo")
    'MODIFICADO: Adett Grover Cruz
    
    'nuestra el formulario que permite modificar la orientación de la página
    af_OrientacionPrint.Show vbModal
    If af_OrientacionPrint.orientacion = 0 Then Exit Function 'cancela la impresión
    
    Grid.PrintInfo.PreviewCaption = Titulo 'caption de la ventana de presentación preliminar
    Grid.PrintInfo.RepeatGridHeader = True
    Grid.PrintInfo.RepeatColumnHeaders = True ' el encabezado se imprime en cada pagina del reporte
    Grid.PrintInfo.PageHeaderFont.Bold = True
    Grid.PrintInfo.PageHeaderFont.Size = 11
    Grid.PrintInfo.PageHeaderFont.Italic = True
    Grid.PrintInfo.RepeatSplitHeaders = True 'encabezado de página
    Grid.PrintInfo.PageHeader = Grid.Splits(0).Caption & "\t\tFecha: " & G_Fecha_Ing & "   Hora: " & G_Hora_Ing 'encabezado de pagina
    Grid.PrintInfo.RepeatColumnFooters = True ' el pie de pagina se imprime en cada hoja
    Grid.PrintInfo.PageFooterFont.Italic = True
    Grid.PrintInfo.PageFooter = "Fuente: Gerencia de sistemas" & "\tServidor: " & G_Servidor & "\tPáginas del \p al \P" ' pie de pagina
    Grid.PrintInfo.SettingsMarginTop = 1000 'margen superior
    Grid.PrintInfo.SettingsMarginBottom = 1000 'margen inferior
    Grid.PrintInfo.SettingsMarginLeft = 1000 'margen izquierdo
    Grid.PrintInfo.SettingsMarginRight = 1000 'margen derecho
    Grid.PrintInfo.SettingsOrientation = af_OrientacionPrint.orientacion ' orientación (1 vertical, 2 horizontal)
    Grid.PrintInfo.SetMenuText 0, "Archivo"
    Grid.PrintInfo.SetMenuText 1, "Imprimir todo                        Ctrl+P"
    Grid.PrintInfo.SetMenuText 2, "Salir                                      Alt+F4"
    Grid.PrintInfo.SetMenuText 3, "Ver"
    Grid.PrintInfo.SetMenuText 4, "Aumentar Zoom            Ctrl +"
    Grid.PrintInfo.SetMenuText 5, "Disminuir Zoom           Ctrl -"
    Grid.PrintInfo.SetMenuText 6, "Ver página completa    Enter"
    Grid.PrintInfo.SetMenuText 7, "Primera página            Inicio"
    Grid.PrintInfo.SetMenuText 8, "Página anterior           Re Pág"
    Grid.PrintInfo.SetMenuText 9, "Página siguiente          Av Pág"
    Grid.PrintInfo.SetMenuText 10, "Última página              Fin"
    Grid.PrintInfo.SetMenuText 11, "Imprimir páginas...               Ctrl+S"
    Grid.PrintInfo.SetMenuText 12, "Imprimir página Actual        Ctrl+C"
    Grid.PrintInfo.SetMenuText 13, "Imprimir páginas"
    Grid.PrintInfo.SetMenuText 14, "Especifique páginas para imprimir:"
    Grid.PrintInfo.SetMenuText 15, "Aceptar"
    Grid.PrintInfo.SetMenuText 16, "Cancelar"
    Grid.PrintInfo.PrintPreview
    
End Function

Public Function pg_OrdenaTdbGrid(Grid As TDBGrid, rs_puntero As ADODB.Recordset, AscDes As Boolean)
    'AUTOR          : René Roque Mendoza.
    'MODIFICADO POR : Maria Luisa Gonzales Mendoza.
    'PROPÓSITO      : Ordena el Grid segun la columna de la celda activa
                    ' Grid -> nombre del TDBGrid sobre el cual se ejecuta la operación de ordenación
                    ' rs_puntero -> recorset del grid
                    ' AscDes -> True ordena ascendentemente y False descendentemente
    
    Dim CampoAct As Integer
    Dim micriterio As String
    
    If rs_puntero.RecordCount = 0 Then Exit Function
    Select Case AscDes
      Case True ' ordenar ascendentemente
        ' capturamos la columna activa
        CampoAct = Grid.Col
        micriterio = Grid.Columns(Grid.Col).DataField
        rs_puntero.Sort = micriterio
        Grid.Col = CampoAct
      Case False ' ordenar descendentemente
        ' capturamos la columna activa
        CampoAct = Grid.Col
        micriterio = Grid.Columns(Grid.Col).DataField & " desc"
        rs_puntero.Sort = micriterio
        Grid.Col = CampoAct
    End Select
End Function
' FIN ADETT

'crystaldesisions.net

Public Function Dias_Del_Mes(Optional ByVal Fecha As Variant) As Integer
    Dim mes As Integer, Y  As Integer
    If IsMissing(Fecha) Then Fecha = Date
     If IsDate(Fecha) Then
         Y = Year(Fecha)
        mes = Month(Fecha)
     ElseIf IsNumeric(Fecha) Then
          Y = Year(Date)
          mes = IIf(Fecha > 0 And Fecha < 13, CInt(Fecha), 0)
      ElseIf VarType(Fecha) = vbString Then
            Y = Year(Date)
         Select Case UCase(Left$(Fecha, 3))
             Case "FEB":                                             mes = 2
             Case "JAN", "MAR", "MAY", "JUL", "AUG", "OCT", "DEC":   mes = 1
             Case "APR", "JUN", "SEP", "NOV":                        mes = 4
         End Select
     End If
      Select Case mes
         Case 2:                     Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
         Case 1, 3, 5, 7, 8, 10, 12: Dias_Del_Mes = 31
         Case 4, 6, 9, 11:           Dias_Del_Mes = 30
     End Select
End Function

Public Function saltarYear(ByVal valor As Variant) As Boolean

     On Error GoTo LocalError

     Dim iYear As Integer

     If IsDate(valor) Then iYear = Year(valor) Else iYear = CInt(valor)

     If TypeName(iYear) = "Integer" Then
         saltarYear = Day(DateSerial(iYear, 3, 0)) = 29
     End If
 Exit Function

LocalError:
End Function


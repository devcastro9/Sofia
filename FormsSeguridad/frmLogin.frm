VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Gestion CGI"
   ClientHeight    =   3780
   ClientLeft      =   270
   ClientTop       =   1770
   ClientWidth     =   5160
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0A02
   ScaleHeight     =   3780
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Inicio de sesión"
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmLogin.frx":C986
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "Cancelar"
      Top             =   2280
      Width           =   1270
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmLogin.frx":D345
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Aceptar"
      Top             =   2280
      Width           =   1270
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   1845
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "QWERTYUOPLKJHGFZXCVB"
      Top             =   675
      Width           =   2460
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1845
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1305
      Width           =   2460
   End
   Begin Crystal.CrystalReport CR07 
      Left            =   0
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label fechaA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Tag             =   "LicenciaA"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label TC 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TDC -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Tag             =   "LicenciaA"
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   795
      TabIndex        =   5
      Tag             =   "Usuari&o:"
      Top             =   705
      Width           =   1035
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   4
      Tag             =   "&Contraseña:"
      Top             =   1305
      Width           =   1440
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Encontrado As Boolean
Dim rsPersonal As New ADODB.Recordset
Dim rsUsuarios As New ADODB.Recordset
Dim RsGlobal As ADODB.Recordset
Dim rs_aux12 As ADODB.Recordset
Dim rs_aux13 As ADODB.Recordset

Public OK As Boolean

Private Sub Form_Load()
'   Dim sBuffer As String
'   Dim lSize As Long
'   sBuffer = Space$(255)
'   lSize = Len(sBuffer)
'   Call GetUserName(sBuffer, lSize)
'   GlMaquina = "Desconocido"
'   If lSize > 0 Then GlMaquina = Left$(sBuffer, lSize)
'   'MsgBox GlMaquina
    Dim nPC As String
    Dim buffer As String
    Dim estado As Long
    frmLogin.Caption = "Sistema de Gestion CGI - v" & App.Major & "." & App.Minor & "." & App.Revision
    buffer = String$(255, " ")
    estado = GetComputerName(buffer, 255)
    If estado <> 0 Then
        nPC = Left(buffer, 255)
    End If
' aqui g
    GlMaquina = Left(nPC, Len(Trim(nPC)) - 1)
    txtUserName.Text = RTrim(GlMaquina)
    'ALB 24/04/2003
    TIPO_DE_CAMBIO
End Sub

Private Sub cmdCancel_Click()
'gerardo   rsPrm.Close
   db.Close 'Freddy
   If rsAccesoSistema.State = 1 Then rsAccesoSistema.Close
   Unload Me
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
    Dim version_valida As Boolean
    'CAMBIOS JQA 2014-ABR-16
   'Sentencia SQL que actualiza el estado de la sesion con el sistema
   db.Execute "Update gc_accesosistema Set estadosesion='I' Where salidasistema='A'"
   'Segmento de codigo que verifica la existencia de usuarios para el sistema
   If rsUsuarios.State = 1 Then rsUsuarios.Close
   rsUsuarios.Open "Select * " & _
                   "From GC_Usuarios " & _
                   "Where estado_codigo= 'APR' ", db, adOpenStatic   'usr_Activo=1
   If rsUsuarios.RecordCount = 0 Then
      MsgBox "El sistema no tiene registrado ningún usuario!!" & vbCr & _
             "Defina los niveles de acceso y usuarios necesarios para " & vbCr & _
             "operar el sistema normalmente.", vbCritical + vbInformation, "Atención"
      Unload frmLogin
      frmMain.Show
      Exit Sub
   End If
   'Segmento de codigo que verifica el login y password del usuario
   Encontrado = False
   If rsUsuarios.State = adStateOpen Then rsUsuarios.Close
   rsUsuarios.Open "Select * From GC_Usuarios Where usr_codigo='" & Trim(txtUserName) & "' And usr_clave = '" & Encriptar(Trim(txtPassword)) & "' And estado_codigo <> 'ANL' ", db, adOpenStatic
   If rsUsuarios.RecordCount = 1 Then
       If rsUsuarios!usr_clave = Encriptar(Trim(txtPassword)) Then  'DUL:Ya que no diferenciaba mayusculas de minusculas
         If rsAccesoSistema.State = 1 Then rsAccesoSistema.Close
         Encontrado = True
         glusuario = Trim(rsUsuarios!usr_codigo)
         GlNombreUsuario = IIf(IsNull(rsUsuarios!usr_nombres), "", rsUsuarios!usr_nombres) & " " & IIf(IsNull(rsUsuarios!usr_primer_apellido), "", rsUsuarios!usr_primer_apellido) & " " & IIf(IsNull(rsUsuarios!usr_segundo_apellido), "", rsUsuarios!usr_segundo_apellido)
         GlIdFuncionario = rsUsuarios!beneficiario_codigo       'idfuncionario
         'GlSistema = rsUsuarios!Sistema
         If Trim(GlNombreUsuario) = "" Then GlNombreUsuario = glusuario

         rsAccesoSistema.Open "Select * From gc_AccesoSistema Where usr_codigo='" & Trim(txtUserName) & "' And EstadoSesion='A'", db, adOpenKeyset, adLockOptimistic
         If rsAccesoSistema.RecordCount = 1 Then
           MsgBox "El usuario: " & glusuario & " se encuentra activo en la máquina: " & GlMaquina & "." & vbCrLf & _
                  "Cierre la sesión activa o ingrese un nuevo usuario", vbCritical + vbOKOnly, "Atención"
           Encontrado = False
           Exit Sub
         Else
            db.Execute "INSERT INTO gc_accesosistema (usr_codigo, maquina, fechalogin, estadosesion, salidasistema, version_principal, version_secundaria, version_revision) " & _
            " VALUES ('" & glusuario & "', '" & Trim(GlMaquina) & "', '" & Date & " " & Time & "', 'A', 'A', " & App.Major & ", " & App.Minor & ", " & App.Revision & ")"
         End If
      End If
   End If
   
   If Encontrado Then
        ' -----------------------------------------
        ' Validacion de version
        ' -----------------------------------------
        version_valida = EsVersionValida(glusuario)
        If version_valida Then
             frmMain.Show
         Else
             MsgBox "La version " & App.Major & "." & App.Minor & "." & App.Revision & " no es admitida. Actualice el sistema.", vbInformation, "Version no admitida"
             End
         End If
      ' Verificamos que exista el tipo de Cambio para la fecha de Maquina
     ' With FrmTipoCambio
        '.TcPrincipal Date, txtUserName.Text & ""
       ' Verificamos que exista el tipo de Cambio para la fecha de SQL Server
       With FrmTipoCambio
                .TcPrincipal GlFechaProceso, txtUserName.Text & ""
      ' Screen.MousePointer = vbHourglass
       If Not .TipoCambioHoy Then
           MsgBox "No se tiene registrado el Tipo de Cambio para la Fecha '" & Format(Date, "dd/mm/yyyy") & "'." & vbCrLf & _
                "Contactese con el Administrador.", vbInformation + vbOKOnly, "Atención"
       Else
'           frmMain.BalanceApertura.Enabled = False
           If UCase(glusuario) = "ADMIN" Then
'               frmMain.BalanceApertura.Enabled = True
           End If
           frmMain.NivelAcceso rsUsuarios!IdNivelAcceso
           Set RsGlobal = New ADODB.Recordset
'           RsGlobal.Open "Select * from Ac_Parametros", db, adOpenStatic, adLockOptimistic
'           GlGestion = RsGlobal!Gestion
'           GlTipoCambioGestion = RsGlobal!TipoCambioGestion
           RsGlobal.Open "Select * from gc_parametros_sistema WHERE estado_registro = 'APR' ", db, adOpenStatic, adLockOptimistic
           glGestion = RsGlobal!ges_gestion
           GlTipoCambioGestion = GlTipoCambioOficial    'RsGlobal!TipoCambioGestion
'           ' -----------------------------------------
'           ' Validacion de version
'           ' -----------------------------------------
'           version_valida = EsVersionValida(glusuario)
'           If version_valida Then
'                frmMain.Show
'            Else
'                MsgBox "La version " & App.Major & "." & App.Minor & "." & App.Revision & " no es admitida. Actualice el sistema.", vbInformation, "Version no admitida"
'                End
'            End If
       End If
      End With
      Unload frmLogin
      If glusuario = "RCUELA" Or glusuario = "VPAREDES" Or glusuario = "MPAREDES" Or glusuario = "APALACIOS" Or glusuario = "VBELLIDO" Or glusuario = "NPAREDES" Or glusuario = "JSAAVEDRA" Then 'Or glusuario = "ADMIN" Or glusuario = "CSALINAS"
        ' INI - ALERTAS VENTAS NUEVAS
        Set rs_aux12 = New ADODB.Recordset
        If rs_aux12.State = 1 Then rs_aux12.Close
        'AND (unidad_destino IS NOT NULL)
        rs_aux12.Open "Select * from ao_ventas_cabecera WHERE ((unidad_codigo = 'DVTA') OR (unidad_codigo = 'DCOMS') OR (unidad_codigo = 'DCOMB') OR (unidad_codigo = 'DCOMC')) AND (estado_codigo = 'APR') AND (unidad_destino IS NULL) ", db, adOpenStatic
        If rs_aux12.RecordCount > 0 Then
            rs_aux12.MoveFirst
             frmMain.ProgressBar1.Visible = True
             With frmMain.ProgressBar1
                .Max = rs_aux12.RecordCount
                .Min = 0
                .Value = 0
             End With
            While Not rs_aux12.EOF
                frmMain.ProgressBar1.Value = frmMain.ProgressBar1.Value + 1
                Set rs_aux13 = New ADODB.Recordset
                If rs_aux13.State = 1 Then rs_aux13.Close
                rs_aux13.Open "Select * from ao_ventas_cabecera WHERE ((unidad_codigo = 'DNMAN') OR (unidad_codigo = 'DMANS') OR (unidad_codigo = 'DMANB') OR (unidad_codigo = 'DMANC')) AND (estado_codigo = 'APR') AND (edif_codigo = '" & rs_aux12!EDIF_CODIGO & "') ", db, adOpenStatic
                If rs_aux13.RecordCount > 0 Then
                    'rs_aux12!unidad_destino = rs_aux13!unidad_CODIGO
                    db.Execute "UPDATE ao_ventas_cabecera SET unidad_destino = '" & rs_aux13!unidad_codigo & "' WHERE (venta_codigo = " & rs_aux12!venta_codigo & " ) "
                    db.Execute "UPDATE ao_ventas_alcance SET estado_mantenimiento = 'APR' WHERE (venta_codigo = " & rs_aux12!venta_codigo & " AND solicitud_tipo ='6') "
                End If
                rs_aux12.MoveNext
            Wend
            frmMain.ProgressBar1.Visible = False
        Else
            frmMain.ProgressBar1.Visible = False
        End If

        Dim iResult As Integer
        CR07.ReportFileName = App.Path & "\reportes\comercial\ar_lista_actas_entrega_definitiva_alerta.rpt"
        CR07.WindowShowPrintSetupBtn = True
        CR07.WindowShowRefreshBtn = True
        iResult = CR07.PrintReport
        If iResult <> 0 Then MsgBox CR07.LastErrorNumber & " : " & CR07.LastErrorString, vbCritical, "Error de impresión"
        CR07.WindowState = crptMaximized
        ' FIN - ALERTAS VENTAS NUEVAS
      End If
      frmMain.ProgressBar1.Visible = False
   Else
      MsgBox "Nombre de usuario o contraseña no válida; vuelva a intentarlo", vbExclamation + vbOKOnly, "Atención"
      txtPassword.SetFocus
      txtPassword.SelStart = 0
      txtPassword.SelLength = Len(txtPassword.Text)
   End If
'   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rsUsuarios.State = 1 Then rsUsuarios.Close
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TIPO_DE_CAMBIO()
'CAMBIOS JQA 2014-ABR-16
Dim TipoCambioOficial As Currency
Dim TipoCambioMercado As Currency
Dim LcSQLAux As String
Dim ExisteTCambio As Boolean
Dim rsAux As ADODB.Recordset
'    Date = "31/12/2006"              ' Para que cambie fecha al ingresar
    TipoCambioOficial = 0
    TipoCambioMercado = 0
    Set rsAux = New ADODB.Recordset
    LcSQLAux = "SELECT * FROM gc_tipo_cambio WHERE Fecha_Cambio = '" & Date & "' "
    rsAux.Open LcSQLAux, db, adOpenStatic
    ExisteTCambio = rsAux.RecordCount > 0
    If ExisteTCambio Then TipoCambioOficial = rsAux!cambio_oficial_compra: TipoCambioMercado = rsAux!cambio_mercado_venta
    fechaA.Caption = Date & "  - " & TipoCambioMercado
End Sub

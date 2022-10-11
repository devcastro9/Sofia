VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
   ClientHeight    =   6060
   ClientLeft      =   270
   ClientTop       =   1770
   ClientWidth     =   8340
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Inicio de sesión"
   Begin VB.Frame fraMainFrame 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6050
      Left            =   20
      TabIndex        =   2
      Top             =   0
      Width           =   8330
      Begin VB.CommandButton cmdOK 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   705
         Left            =   5580
         Picture         =   "frmLogin.frx":324A
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "Aceptar"
         Top             =   4380
         Width           =   900
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   705
         Left            =   6840
         Picture         =   "frmLogin.frx":3554
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Cancelar"
         Top             =   4380
         Width           =   900
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   5895
         MaxLength       =   15
         TabIndex        =   0
         Top             =   3360
         Width           =   2028
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5895
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3750
         Width           =   2028
      End
      Begin VB.Label lblLabels 
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
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   12
         Tag             =   "Usuari&o:"
         Top             =   3360
         Width           =   1305
      End
      Begin VB.Label lblLabels 
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
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   11
         Tag             =   "&Contraseña:"
         Top             =   3765
         Width           =   1365
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright @jqa -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Tag             =   "Copyright"
         Top             =   5700
         Width           =   1515
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Advertencia:  Esta Prohibida la copia parcial o total del producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Tag             =   "Advertencia"
         Top             =   5460
         Width           =   5595
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versión para:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4560
         TabIndex        =   8
         Tag             =   "Versión"
         Top             =   1740
         Width           =   1512
      End
      Begin VB.Label lblPlatform 
         BackStyle       =   0  'Transparent
         Caption         =   "Financiadores Externos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   4500
         TabIndex        =   7
         Tag             =   "Plataforma"
         Top             =   1980
         Width           =   3000
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GTZ - UDAPRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   3210
         TabIndex        =   6
         Tag             =   "ProductoOrganización"
         Top             =   180
         Width           =   2700
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAF-2000 / FE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   675
         Left            =   2490
         TabIndex        =   5
         Tag             =   "Producto"
         Top             =   600
         Width           =   3945
      End
      Begin VB.Label lblLicenseTo 
         BackStyle       =   0  'Transparent
         Caption         =   "La Paz - Bolivia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Tag             =   "LicenciaA"
         Top             =   5700
         Width           =   1590
      End
      Begin VB.Image Imglogo 
         Height          =   4380
         Left            =   -120
         Picture         =   "frmLogin.frx":375E
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   5940
         Left            =   30
         Picture         =   "frmLogin.frx":CAF0
         Stretch         =   -1  'True
         Top             =   60
         Width           =   8250
      End
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

'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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
    buffer = String$(255, " ")
    estado = GetComputerName(buffer, 255)
    If estado <> 0 Then
        nPC = Left(buffer, 255)
    End If
' aqui greco
    GlMaquina = Left(nPC, Len(Trim(nPC)) - 1)
End Sub

Private Sub cmdCancel_Click()
   db.Close 'Freddy
   If rsAccesoSistema.State = 1 Then rsAccesoSistema.Close
   Unload Me
End Sub

Private Sub cmdOK_Click()
   'Sentencia SQL que actualiza el estado de la sesion con el sistema
   Screen.MousePointer = vbHourglass
   db.Execute "Update AccesoSistema Set estadosesion='I' Where salidasistema='A'"
   'Segmento de codigo que verifica el login y password del usuario
   Encontrado = False
   If rsUsuarios.State = 1 Then rsUsuarios.Close
   If txtPassword = "" Then
      rsUsuarios.Open "Select * From Usuarios_Udapre Where usr_usuario='" & Trim(txtUserName) & "' And usr_clave = '" & Encriptar(Trim("x")) & "' And usr_Activo = 1", db, adOpenStatic
      If rsUsuarios.RecordCount = 1 Then
         Screen.MousePointer = vbDefault
         MsgBox "Bienvenido al Sistema SAF-2000! Registre su clave de acceso." & vbCrLf & _
                "Luego ingrese al sistema con su clave de acceso.", vbInformation + vbOKOnly, "Nuevo Usuario"
         GlUsuario = txtUserName
         rsPersonal.Open "Select * From rc_Personal Where IdFuncionario=" & rsUsuarios!IdFuncionario, db, adOpenStatic
         If rsPersonal.RecordCount = 1 Then
            GlNombreUsuario = Trim(rsPersonal!Nombres) & " " & Trim(rsPersonal!Paterno) & " " & Trim(rsPersonal!Materno)
         End If
         rsPersonal.Close
         FrmCambiarClave.Show vbModal
'         Unload FrmCambiarClave
         Unload Me
         Exit Sub
      End If
   End If
   If rsUsuarios.State = 1 Then rsUsuarios.Close
   rsUsuarios.Open "Select * " & _
                   "From Usuarios_Udapre " & _
                   "Where usr_usuario='" & Trim(txtUserName) & "' And usr_clave = '" & Encriptar(Trim(txtPassword)) & "' And usr_Activo = 1", db, adOpenStatic
   If rsUsuarios.RecordCount = 1 Then
       If rsUsuarios!usr_clave = Encriptar(Trim(txtPassword)) Then  'DUL:Ya que no diferenciaba mayusculas de minusculas
         If rsAccesoSistema.State = 1 Then rsAccesoSistema.Close
         Encontrado = True
         GlUsuario = Trim(rsUsuarios!usr_usuario)
         rsAccesoSistema.Open "Select * From AccesoSistema Where usr_usuario='" & Trim(txtUserName) & "' And EstadoSesion='A'", db, adOpenKeyset, adLockOptimistic
         If rsAccesoSistema.RecordCount = 1 Then
           MsgBox "El usuario: " & GlUsuario & " se encuentra activo en la máquina: " & GlMaquina & "." & vbCrLf & _
                  "Cierre la sesión activa o ingrese un nuevo usuario", vbCritical + vbOKOnly, "Atención"
           Encontrado = False
           Exit Sub
         Else
           rsAccesoSistema.AddNew
           rsAccesoSistema!usr_usuario = GlUsuario
           rsAccesoSistema!maquina = Trim(GlMaquina)
           rsAccesoSistema!FechaLogin = Date & " " & Time
           rsAccesoSistema!estadosesion = "A" 'Sesion con el sistema A=Activo T=terminado I=Interrumpido
           rsAccesoSistema!SalidaSistema = "A"  'Salida del sistema A=Anormal  N=Normal
           db.BeginTrans
           rsAccesoSistema.Update
           db.CommitTrans
         End If
         If GlUsuario = "ADMIN" Then GlNombreUsuario = GlUsuario
      End If
   End If
   'Segmento de codigo que busca el nombre del usuario
   If Encontrado Then
      rsPersonal.Open "Select * From rc_Personal Where IdFuncionario=" & rsUsuarios!IdFuncionario, db, adOpenStatic
      If rsPersonal.RecordCount = 1 Then
         GlNombreUsuario = Trim(rsPersonal!Nombres) & " " & Trim(rsPersonal!Paterno) & " " & Trim(rsPersonal!Materno)
      End If
      rsPersonal.Close
      Screen.MousePointer = vbDefault
      ' Verificamos que exista el tipo de Cambio para la fecha
      With FrmTipoCambio
       .TcPrincipal Date, txtUserName.Text & ""
       Screen.MousePointer = vbHourglass
       If Not .TipoCambioHoy Then
         MsgBox "No se tiene registrado el Tipo de Cambio para la Fecha '" & Format(Date, "dd/mm/yyyy") & "'." & vbCrLf & _
                "Contactese con el Administrador.", vbInformation + vbOKOnly, "Atención"
       Else
        frmMain.NivelAcceso rsUsuarios!idNivelAcceso
        frmMain.Show
       End If
      End With
      Unload frmLogin
   Else
      MsgBox "Nombre de usuario o contraseña no válida; vuelva a intentarlo", vbExclamation + vbOKOnly, "Atención"
      txtPassword.SetFocus
      txtPassword.SelStart = 0
      txtPassword.SelLength = Len(txtPassword.Text)
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rsUsuarios.State = 1 Then rsUsuarios.Close
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub picLogo_Click()

End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

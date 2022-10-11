VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3810
   ClientLeft      =   270
   ClientTop       =   1770
   ClientWidth     =   4830
   ForeColor       =   &H00C0FFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Inicio de sesión"
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00404000&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1620
         Width           =   2250
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C000&
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   825
         Left            =   840
         Picture         =   "frmLogin.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Aceptar"
         Top             =   2520
         Width           =   900
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C000&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   825
         Left            =   3000
         Picture         =   "frmLogin.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Cancelar"
         Top             =   2520
         Width           =   900
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00404000&
         ForeColor       =   &H80000005&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2040
         Width           =   2265
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "ADFIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1665
         Left            =   180
         TabIndex        =   8
         Tag             =   "Producto"
         Top             =   -120
         Width           =   4155
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADFIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1605
         Left            =   0
         TabIndex        =   10
         Tag             =   "Producto"
         Top             =   0
         Width           =   4800
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Todos los derechos reservados por SPC-Bolivia @ADFIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Tag             =   "Usuari&o:"
         Top             =   3480
         Width           =   4545
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRASEÑA :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Tag             =   "&Contraseña:"
         Top             =   2020
         Width           =   1875
      End
      Begin VB.Image Image1 
         Height          =   1290
         Left            =   120
         Picture         =   "frmLogin.frx":3DCE
         Top             =   0
         Width           =   4245
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Tag             =   "Usuari&o:"
         Top             =   1575
         Width           =   1785
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   2295
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Tag             =   "&Contraseña:"
         Top             =   1560
         Width           =   4845
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   1080
      Y1              =   0
      Y2              =   0
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
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long

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
    GlMaquina = Left(nPC, Len(Trim(nPC)) - 1)
End Sub
Private Sub cmdCancel_Click()
'   rsPrm.Close
   db.Close '
   If rsAccesoSistema.State = 1 Then rsAccesoSistema.Close
   Unload Me
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
   'Sentencia SQL que actualiza el estado de la sesion con el sistema
   Screen.MousePointer = vbHourglass
   db.Execute "Update AccesoSistema Set estadosesion='I' Where salidasistema='A'"
   
   'Segmento de codigo que verifica la existencia de usuarios para el sistema
   If rsUsuarios.State = 1 Then rsUsuarios.Close
   rsUsuarios.Open "Select * From Usuarios_queiros Where usr_Activo=1", db, adOpenStatic 'usr_Activo=1
   If rsUsuarios.RecordCount = 0 Then
      MsgBox "El sistema ADFIN-2002 no tiene registrado ningún usuario!!" & vbCr & _
             "Defina los niveles de acceso y usuarios necesarios para " & vbCr & _
             "operar el sistema normalmente.", vbCritical + vbInformation, "Atención"
      Unload frmLogin
      Screen.MousePointer = vbDefault
      frmMain.Show
      Exit Sub
   End If
   
   'Segmento de codigo que verifica el login y password del usuario
   Encontrado = False
   If rsUsuarios.State = 1 Then rsUsuarios.Close
   rsUsuarios.Open "Select * " & _
                   "From Usuarios_queiros Where usr_usuario='" & Trim(txtUserName) & "' And usr_clave = '" & Encriptar(Trim(txtPassword)) & "' And usr_Activo=1", db, adOpenStatic 'usr_Activo=1
   If rsUsuarios.RecordCount = 1 Then
       If rsUsuarios!usr_clave = Encriptar(Trim(txtPassword)) Then  'DUL:Ya que no diferenciaba mayusculas de minusculas
         If rsAccesoSistema.State = 1 Then rsAccesoSistema.Close
         Encontrado = True
         GlUsuario = Trim(rsUsuarios!usr_usuario)
         GlNombreUsuario = IIf(IsNull(rsUsuarios!NombreS), "", rsUsuarios!NombreS) & " " & IIf(IsNull(rsUsuarios!paterno), "", rsUsuarios!paterno) & " " & IIf(IsNull(rsUsuarios!materno), "", rsUsuarios!materno)
         GlIdFuncionario = rsUsuarios!idfuncionario
         If Trim(GlNombreUsuario) = "" Then GlNombreUsuario = GlUsuario

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
           rsAccesoSistema.Update
         End If
      End If
   End If
   
   If Encontrado Then
      ' Verificamos que exista el tipo de Cambio para la fecha
      With FrmTipoCambio
       .TcPrincipal Date, txtUserName.Text & ""
       Screen.MousePointer = vbHourglass
       If Not .TipoCambioHoy Then
         MsgBox "No se tiene registrado el Tipo de Cambio para la Fecha '" & Format(Date, "dd/mm/yyyy") & "'." & vbCrLf & _
                "Contactese con el Administrador.", vbInformation + vbOKOnly, "Atención"
       Else
        frmMain.mnuRepBalApertura.Enabled = False
        If UCase(GlUsuario) = "SIS" Or "SAF" Or UCase(GlUsuario) = "C_LOPEZ" Or UCase(GlUsuario) = "C_GARRON" Or UCase(GlUsuario) = "M_URQUIOLA" Or UCase(GlUsuario) = "I_IMAÑA" Or UCase(GlUsuario) = "M_YAÑEZ" Then
          frmMain.mnuRepBalApertura.Enabled = True
        End If
        frmMain.NivelAcceso rsUsuarios!IdNivelAcceso
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



Private Sub txtUserName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

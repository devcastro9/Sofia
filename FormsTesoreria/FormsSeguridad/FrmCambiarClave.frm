VERSION 5.00
Begin VB.Form FrmCambiarClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Clave de Acceso"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "FrmCambiarClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   15
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Frame FraNuevaClave 
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   4455
         Begin VB.TextBox txtNuevaClave 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtVerificaClave 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Clave"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   13
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Verifica Clave"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2760
            TabIndex        =   12
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox txtClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtLogin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   15
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   2400
         TabIndex        =   5
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Login:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   585
      End
   End
End
Attribute VB_Name = "FrmCambiarClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUsuarios As New ADODB.Recordset

Private Sub CmdGrabar_Click()
If ValidaDatos Then
    rsUsuarios!usr_Usuario = Trim(txtLogin)
    rsUsuarios!usr_clave = Encriptar(Trim(txtNuevaClave))
    rsUsuarios.Update
    MsgBox "Nueva clave fue grabado satisfactoriamente", vbInformation + vbOKOnly, "Atención"
    Unload Me
End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Se hace visible al principio el textbox para la clave
    txtClave.Visible = True
    rsUsuarios.Open "Select * From Usuarios_Udapre Where usr_usuario='" & GlUsuario & "' And Usr_Activo = 1", db, adOpenKeyset, adLockOptimistic
    If rsUsuarios.RecordCount = 0 Then
       MsgBox "El usuario no existe !...", vbExclamation + vbOKOnly, "Atención"
       Unload Me
    Else
       LblUsuario = GlNombreUsuario
       txtLogin = rsUsuarios!usr_Usuario
       If Desencriptar(rsUsuarios!usr_clave) = "x" Then
          'Se hace invisible el textbox para la clave si es un usuario nuevo
          FrmCambiarClave.Caption = "Nuevo Usuario: Ingrese su clave de acceso ..."
          Label4.Visible = False 'Label Clave: invisible
          txtClave.Visible = False
       End If
    End If
End Sub

Function ValidaDatos() As Boolean
    ValidaDatos = True
    If Trim(txtLogin) = "" Then
        MsgBox "Debe introducir su Login de usuario.", vbExclamation + vbOKOnly, "Atención"
        txtLogin.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Desencriptar(rsUsuarios!usr_clave) <> "x" Then
        If Trim(txtClave) = "" Then
            MsgBox "Debe introducir su clave de acceso.", vbExclamation + vbOKOnly, "Atención"
            txtClave.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtClave) <> Desencriptar(rsUsuarios!usr_clave) Then
            MsgBox "Clave de usuario incorrecta!", vbInformation + vbOKOnly, "Atención"
            txtClave.SetFocus
            txtClave.SelStart = 0
            txtClave.SelLength = Len(txtClave)
            ValidaDatos = False
            Exit Function
        End If
    End If
    If Trim(txtNuevaClave) = "" Then
        MsgBox "Debe introducir su nueva clave de usuario!", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(txtNuevaClave) < 6 Then
        MsgBox "Su nueva clave debe tener al menos seis caracteres!", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If BuscaCaracter(txtNuevaClave, " ") <> 0 Then
        MsgBox "Su nueva clave no debe contener espacios. Revise!!", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Trim(txtNuevaClave) <> Trim(txtVerificaClave) Then
        MsgBox "Verificacion de nueva clave incorrecta!. Intente nuevamente.", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave = ""
        txtVerificaClave = ""
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Trim(txtNuevaClave) = Trim(txtClave) Then
        MsgBox "Debe introducir una nueva clave de acceso!", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave = ""
        txtVerificaClave = ""
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    rsUsuarios.Close
End Sub

Private Function BuscaCaracter(Cadena As String, Caracter As String) As Byte
BuscaCaracter = 0
Dim i As Integer
  i = 1
  While Len(Cadena) > 0 And BuscaCaracter = 0
      If Mid(Cadena, 1, 1) = Caracter Then
         BuscaCaracter = i
      Else
         If Len(Cadena) > 1 Then
          Cadena = Mid(Cadena, 2, Len(Cadena) - 1)
          i = i + 1
         Else
          Cadena = ""
         End If
      End If
  Wend
End Function

VERSION 5.00
Begin VB.Form FrmCambiarClave 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Clave de Acceso"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "FrmCambiarClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   676
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5400
      TabIndex        =   12
      Top             =   0
      Width           =   5400
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2595
         Picture         =   "FrmCambiarClave.frx":0442
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   960
         Picture         =   "FrmCambiarClave.frx":0D2E
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   13
         Top             =   0
         Width           =   1280
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13095
         TabIndex        =   15
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      Begin VB.Frame FraNuevaClave 
         BackColor       =   &H00000000&
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   4455
         Begin VB.TextBox txtNuevaClave 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtVerificaClave 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2625
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Nueva Contraseña"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   195
            Left            =   195
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Repita Contraseña"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   195
            Left            =   2565
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtLogin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Contraseña Actual:"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3000
         TabIndex        =   5
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nombre del Usuario:"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   960
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

Private Sub BtnGrabar_Click()
    If ValidaDatos Then
        rsUsuarios!usr_codigo = Trim(txtLogin)
        rsUsuarios!usr_clave = Encriptar(Trim(txtNuevaClave))
        rsUsuarios.Update
        MsgBox "La nueva clave fue grabada satisfactoriamente ...", vbInformation + vbOKOnly, "Atención"
        Unload Me
    End If
End Sub

Private Sub BtnCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Se hace visible al principio el textbox para la clave
    txtClave.Visible = True
    rsUsuarios.Open "Select * From GC_Usuarios Where usr_codigo='" & glusuario & "' And estado_codigo <> 'ANL'", db, adOpenKeyset, adLockOptimistic
    If rsUsuarios.RecordCount = 0 Then
       MsgBox "El usuario no existe !...", vbExclamation + vbOKOnly, "Atención"
       Frame1.Enabled = False
       BtnGrabar.Enabled = False
    Else
       LblUsuario = GlNombreUsuario
       txtLogin = rsUsuarios!usr_codigo
       If Desencriptar(rsUsuarios!usr_clave) = "" Then
          'Se hace invisible el textbox para la clave si es un usuario nuevo
          FrmCambiarClave.Caption = "Nuevo Usuario: Ingrese su clave de acceso ..."
          Label4.Visible = False 'Label Clave: invisible
          txtClave.Visible = False
       End If
    End If
End Sub

Function ValidaDatos() As Boolean
    ValidaDatos = True
    If Desencriptar(rsUsuarios!usr_clave) <> "" Then
        If Trim(txtClave) = "" Then
            MsgBox "Debe introducir la clave de acceso.", vbExclamation + vbOKOnly, "Atención"
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
        MsgBox "Debe introducir la nueva clave de usuario!", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(txtNuevaClave) < 6 Then
        MsgBox "La nueva clave debe tener al menos seis caracteres!", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If BuscaCaracter(txtNuevaClave, " ") <> 0 Then
        MsgBox "La nueva clave no debe contener espacios. Revise por favor!", vbExclamation + vbOKOnly, "Atención"
        txtNuevaClave.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Trim(txtNuevaClave) <> Trim(txtVerificaClave) Then
        MsgBox "La nueva clave es incorrecta!. Intente nuevamente.", vbExclamation + vbOKOnly, "Atención"
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

Private Sub txtNuevaClave_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 And KeyAscii > 165) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtVerificaClave_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 And KeyAscii > 165) Then
        Beep
        KeyAscii = 0
    End If
End Sub

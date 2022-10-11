VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información de BD"
   ClientHeight    =   1545
   ClientLeft      =   9765
   ClientTop       =   8865
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Text            =   "DRojas"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      TabIndex        =   3
      Text            =   "Gtz_Udapre"
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Servidor"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Base de Datos"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Unload Me
End Sub
Private Sub cmdOK_Click()
    'comprobar si la contraseña es correcta
    If Trim(txtUserName.Text) = "" Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
        MsgBox "Ingrese el Nombre del Servidor al cual se conectara la Aplicación", , "Inicio de sesión"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim(txtPassword) = "" Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
        MsgBox "Ingrese la base de datos a la cual se conectara la Aplicación", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    LoginSucceeded = True
    GlServidor = txtUserName.Text
    GlNombreBD = txtPassword.Text
    GlBaseDatos = GlNombreBD
    Unload Me
End Sub

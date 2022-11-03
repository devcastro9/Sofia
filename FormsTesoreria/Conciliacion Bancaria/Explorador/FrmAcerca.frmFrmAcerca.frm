VERSION 5.00
Begin VB.Form FrmAcerca 
   Caption         =   "Acerca de Windows"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   405
      Left            =   3630
      TabIndex        =   0
      Top             =   2805
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   45
      Picture         =   "FrmAcerca.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   990
   End
   Begin VB.Label LblAcerca3 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   1185
      TabIndex        =   3
      Top             =   2265
      Width           =   4005
   End
   Begin VB.Label LblAcerca2 
      BackStyle       =   0  'Transparent
      Height          =   825
      Left            =   1185
      TabIndex        =   2
      Top             =   1245
      Width           =   3930
   End
   Begin VB.Label LblAcerca1 
      BackStyle       =   0  'Transparent
      Height          =   1035
      Left            =   1185
      TabIndex        =   1
      Top             =   180
      Width           =   3930
   End
End
Attribute VB_Name = "FrmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LblAcerca1.Caption = "Microsoft (R) Window" & _
    "Windows 95" & _
    "(C) 1981 - 1987 Microsoft Coporation"
    
    LblAcerca2.Caption = "Se autoriza el uso de este producto a" & _
    "." & _
    ".."
    
    LblAcerca3.Caption = "Memoria Física disponible para window:  32.260 KB" & _
    "Recursos del sistema:   56% disponible"
    
	Call SeguridadSet(Me)
End Sub

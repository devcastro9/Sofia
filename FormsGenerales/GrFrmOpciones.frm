VERSION 5.00
Begin VB.Form GrFrmOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   1725
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   2790
   Icon            =   "GrFrmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   2790
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   1500
      TabIndex        =   4
      Top             =   1320
      Width           =   1092
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   " Elija una Opción "
      Height          =   1152
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2652
      Begin VB.OptionButton OptOpciones 
         Height          =   312
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   2292
      End
      Begin VB.OptionButton OptOpciones 
         Height          =   312
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   2292
      End
   End
End
Attribute VB_Name = "GrFrmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public POpcionElegida As Byte

Private Sub BtnGrabar_Click()
  If OptOpciones(1).Value Then POpcionElegida = 1
  If OptOpciones(2).Value Then POpcionElegida = 2
  Unload Me
End Sub

Private Sub BtnCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  POpcionElegida = 0
	Call SeguridadSet(Me)
End Sub


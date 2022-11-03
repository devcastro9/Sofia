VERSION 5.00
Begin VB.Form GrFrmOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elija su opción"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   Icon            =   "GrFrmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraOp 
      Height          =   1635
      Left            =   53
      TabIndex        =   0
      Top             =   585
      Width           =   3000
      Begin VB.OptionButton Opcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   6
         Top             =   1305
         Width           =   2670
      End
      Begin VB.OptionButton Opcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   960
         Width           =   2670
      End
      Begin VB.OptionButton Opcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   600
         Width           =   2670
      End
      Begin VB.OptionButton Opcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2670
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Elegir"
      Height          =   570
      Left            =   825
      Picture         =   "GrFrmOpciones.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2295
      Width           =   1455
   End
   Begin VB.Label LblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Elija una de las opciones siguientes:"
      Height          =   495
      Left            =   53
      TabIndex        =   4
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "GrFrmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PSElegido As Byte

Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub CmdAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Not Opcion(1).Value Then Opcion(1).ForeColor = vbButtonText
  If Not Opcion(2).Value Then Opcion(2).ForeColor = vbButtonText
  If Not Opcion(3).Value Then Opcion(3).ForeColor = vbButtonText
  If Not Opcion(4).Value Then Opcion(4).ForeColor = vbButtonText
End Sub

Private Sub Form_Activate()
  If Opcion(4).Caption = "" Then
    Opcion(4).Visible = False
    FraOp.Height = 1275
    CmdAceptar.Top = 2040
    GrFrmOpciones.Height = 3270
  End If
  If Opcion(3).Caption = "" Then
    Opcion(3).Visible = False
    FraOp.Height = 915
    CmdAceptar.Top = 1680
    GrFrmOpciones.Height = 2910
  End If
  If Opcion(2).Caption = "" Then
    Opcion(2).Visible = False
    FraOp.Height = 550
    CmdAceptar.Top = 1320
    GrFrmOpciones.Height = 2565
  End If
  If Opcion(1).Caption = "" Then
    Opcion(1).Visible = False
  End If
End Sub

Private Sub Form_Load()
  PSElegido = 1
	Call SeguridadSet(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Not Opcion(1).Value Then Opcion(1).ForeColor = vbButtonText '&H80000012
  If Not Opcion(2).Value Then Opcion(2).ForeColor = vbButtonText
  If Not Opcion(3).Value Then Opcion(3).ForeColor = vbButtonText
  If Not Opcion(4).Value Then Opcion(4).ForeColor = vbButtonText
End Sub

Private Sub FraOp_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Not Opcion(1).Value Then Opcion(1).ForeColor = vbButtonText
  If Not Opcion(2).Value Then Opcion(2).ForeColor = vbButtonText
  If Not Opcion(3).Value Then Opcion(3).ForeColor = vbButtonText
  If Not Opcion(4).Value Then Opcion(4).ForeColor = vbButtonText
End Sub

Private Sub Opcion_Click(Index As Integer)
Dim I As Integer
  PSElegido = Index
  For I = 1 To 4
    Opcion(I).FontBold = False
    Opcion(I).ForeColor = vbButtonText
  Next I
  Opcion(Index).FontBold = True
  Opcion(Index).ForeColor = vbHighlight
End Sub

Private Sub Opcion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
  Opcion(Index).ForeColor = vbHighlight '&H8000000D
End Sub

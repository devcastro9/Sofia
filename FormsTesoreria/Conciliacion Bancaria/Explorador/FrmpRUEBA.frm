VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2055
      Left            =   2340
      TabIndex        =   2
      Top             =   720
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   390
      TabIndex        =   1
      Top             =   720
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   1950
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2115
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   720
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim Tam As Integer
        Tam = Frame1.Width + Frame2.Width
        Picture1.Left = Picture1.Left + X
        Frame1.Width = Frame1.Width + X
        Frame2.Left = Frame2.Left + X
        File1.Left = File1.Left + X
        Frame2.Width = Tam - Frame1.Width
        
End Sub

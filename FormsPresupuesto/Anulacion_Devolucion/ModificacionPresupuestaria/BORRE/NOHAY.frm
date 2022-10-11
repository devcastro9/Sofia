VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "INFORMACIÓN..."
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2745
      TabIndex        =   0
      Top             =   5310
      Width           =   1710
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   315
      Left            =   1545
      Shape           =   2  'Oval
      Top             =   4620
      Width           =   3075
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   390
      Left            =   1605
      Shape           =   2  'Oval
      Top             =   3855
      Width           =   885
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ejecución Presupuestaria    Reportes                  Modificación Presupuestaria "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   1710
      TabIndex        =   4
      Top             =   4230
      Width           =   2745
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clasificadores     Egresos     Tesorería   Contabilidad ......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   3900
      Width           =   6015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  SAF - 2000  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2505
      TabIndex        =   2
      Top             =   3105
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   2865
      Picture         =   "NOHAY.frx":0000
      Top             =   2115
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"NOHAY.frx":324A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   810
      TabIndex        =   5
      Top             =   120
      Width           =   4950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"NOHAY.frx":32D7
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1680
      Left            =   870
      TabIndex        =   1
      Top             =   165
      Width           =   4905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
  End
End Sub


VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8715
      Left            =   1350
      TabIndex        =   13
      Top             =   1050
      Width           =   8715
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   15180
      TabIndex        =   6
      Top             =   0
      Width           =   15240
      Begin VB.Label Label8 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Index           =   0
         Left            =   1245
         TabIndex        =   10
         Top             =   690
         Width           =   2460
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9210
         TabIndex        =   9
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   8
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Conciliacion Bancaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2160
         TabIndex        =   7
         Top             =   195
         Width           =   8265
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   8745
      Left            =   45
      TabIndex        =   0
      Top             =   1035
      Width           =   1230
      Begin VB.CommandButton Command1 
         Caption         =   "Ejemplo Ger"
         Height          =   795
         Left            =   150
         TabIndex        =   12
         Top             =   2715
         Width           =   945
      End
      Begin VB.CommandButton CmdConciliacionUDAPRE 
         Caption         =   "Conciliar pro fecha UDAPRE"
         Height          =   795
         Left            =   150
         MousePointer    =   4  'Icon
         Picture         =   "Form1111.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   285
         Width           =   945
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Limpiar"
         Height          =   720
         Left            =   135
         Picture         =   "Form1111.frx":09A2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   150
         Picture         =   "Form1111.frx":0DE4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5985
         Width           =   945
      End
      Begin VB.CommandButton CmdImprimirTotales 
         Caption         =   "Imprimir"
         Height          =   795
         Left            =   150
         Picture         =   "Form1111.frx":1226
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1095
         Width           =   945
      End
      Begin VB.CommandButton CmdUnionTablas 
         Caption         =   "Unión Tablas"
         Height          =   795
         Left            =   150
         MousePointer    =   4  'Icon
         Picture         =   "Form1111.frx":1890
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1905
         Width           =   945
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Ejemplo gerardo
On Error GoTo QError
    db.UNO 2, 3
    Exit Sub
QError:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Private Sub Form_Load()

	Call SeguridadSet(Me)
End Sub

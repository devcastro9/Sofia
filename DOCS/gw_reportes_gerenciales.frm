VERSION 5.00
Begin VB.Form gw_reportes_gerenciales 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reportes Gerenciales"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "gw_reportes_gerenciales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Label LGral 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0. GENERALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label LMan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5. MANTENIMIENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   9360
      Width           =   2295
   End
   Begin VB.Label LRep 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6. REPARCIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   13680
      TabIndex        =   6
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label LEme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7. EMERGENCIAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   14760
      TabIndex        =   5
      Top             =   5080
      Width           =   2055
   End
   Begin VB.Label LMod 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8. MODERNIZACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   15720
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label LIns 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4. INSTALACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   9360
      Width           =   2055
   End
   Begin VB.Label LCmx 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3. COMEX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label LCom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2. COMERCIAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   5080
      Width           =   2055
   End
   Begin VB.Label LRrhh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1. RR.HH."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Image RCmx 
      Height          =   1815
      Left            =   4440
      Picture         =   "gw_reportes_gerenciales.frx":0A02
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image RMan 
      Height          =   1815
      Left            =   10680
      Picture         =   "gw_reportes_gerenciales.frx":273E
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Image RRep 
      Height          =   1815
      Left            =   13560
      Picture         =   "gw_reportes_gerenciales.frx":5810
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image REme 
      Height          =   1815
      Left            =   14640
      Picture         =   "gw_reportes_gerenciales.frx":67E1
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Image RMod 
      Height          =   1815
      Left            =   15720
      Picture         =   "gw_reportes_gerenciales.frx":8BF2
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image RGral 
      Height          =   1815
      Left            =   9120
      Picture         =   "gw_reportes_gerenciales.frx":A06E
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Image RIns 
      Height          =   1815
      Left            =   7440
      Picture         =   "gw_reportes_gerenciales.frx":BD86
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Image RCom 
      Height          =   1815
      Left            =   3240
      Picture         =   "gw_reportes_gerenciales.frx":CFC0
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Image RRrhh 
      Height          =   1815
      Left            =   2040
      Picture         =   "gw_reportes_gerenciales.frx":E1C4
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   8040
      Picture         =   "gw_reportes_gerenciales.frx":F80C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3930
   End
End
Attribute VB_Name = "gw_reportes_gerenciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RGral_Click()
    gw_rep_generales.Show
End Sub

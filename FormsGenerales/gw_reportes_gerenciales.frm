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
   ScaleHeight     =   5130
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
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
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lbl_vta 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONTRATOS EN $"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "FACTURACION EN $"
      Height          =   255
      Left            =   9240
      TabIndex        =   10
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "COBRANZAS EN $"
      Height          =   255
      Left            =   12360
      TabIndex        =   9
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Image rep_vta 
      Height          =   1335
      Left            =   6000
      Picture         =   "gw_reportes_gerenciales.frx":0A02
      Stretch         =   -1  'True
      ToolTipText     =   "CONTRATOS EN $"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Image rep_fac 
      Height          =   1335
      Left            =   9240
      Picture         =   "gw_reportes_gerenciales.frx":1D2E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Image rep_cob 
      Height          =   1335
      Left            =   12360
      Picture         =   "gw_reportes_gerenciales.frx":319C
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label LGral 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0. MAS REPORTES"
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
      Top             =   6840
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
      Top             =   9000
      Visible         =   0   'False
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
      Top             =   6840
      Visible         =   0   'False
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
      Top             =   4725
      Visible         =   0   'False
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
      Top             =   2520
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
      Top             =   9000
      Visible         =   0   'False
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
      Top             =   6840
      Visible         =   0   'False
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
      Top             =   4725
      Visible         =   0   'False
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image RCmx 
      Height          =   1815
      Left            =   4440
      Picture         =   "gw_reportes_gerenciales.frx":40D4
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image RMan 
      Height          =   1815
      Left            =   10680
      Picture         =   "gw_reportes_gerenciales.frx":5E10
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image RRep 
      Height          =   1815
      Left            =   13560
      Picture         =   "gw_reportes_gerenciales.frx":8EE2
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image REme 
      Height          =   1815
      Left            =   14640
      Picture         =   "gw_reportes_gerenciales.frx":9EB3
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image RMod 
      Height          =   1815
      Left            =   15720
      Picture         =   "gw_reportes_gerenciales.frx":C2C4
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2295
   End
   Begin VB.Image RGral 
      Height          =   1815
      Left            =   9120
      Picture         =   "gw_reportes_gerenciales.frx":D740
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Image RIns 
      Height          =   1815
      Left            =   7440
      Picture         =   "gw_reportes_gerenciales.frx":F458
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image RCom 
      Height          =   1815
      Left            =   3240
      Picture         =   "gw_reportes_gerenciales.frx":10692
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image RRrhh 
      Height          =   1815
      Left            =   2040
      Picture         =   "gw_reportes_gerenciales.frx":11896
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   8040
      Picture         =   "gw_reportes_gerenciales.frx":12EDE
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

VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmPropiedades 
   Caption         =   "Propiedades de"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form2"
   ScaleHeight     =   5865
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      ForeColor       =   &H8000000A&
      Height          =   4920
      Left            =   135
      TabIndex        =   1
      Top             =   405
      Width           =   4695
      Begin VB.CheckBox Check5 
         Caption         =   "Sistema"
         Height          =   240
         Left            =   2685
         TabIndex        =   17
         Top             =   3585
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Oculto"
         Height          =   240
         Left            =   2670
         TabIndex        =   16
         Top             =   3285
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Modificado"
         Height          =   240
         Left            =   1230
         TabIndex        =   15
         Top             =   3540
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sólo  lectura"
         Height          =   240
         Left            =   1230
         TabIndex        =   14
         Top             =   3300
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Habilitar una vista de página en miniatura"
         Height          =   285
         Left            =   750
         TabIndex        =   13
         Top             =   4170
         Width           =   3270
      End
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   210
         TabIndex        =   12
         Top             =   3855
         Width           =   4260
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   285
         TabIndex        =   10
         Top             =   1065
         Width           =   4185
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   255
         TabIndex        =   9
         Top             =   2325
         Width           =   4260
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   240
         TabIndex        =   8
         Top             =   3015
         Width           =   4260
      End
      Begin VB.Label LblCreado 
         Caption         =   "Label8"
         Height          =   225
         Left            =   1965
         TabIndex        =   23
         Top             =   2805
         Width           =   1320
      End
      Begin VB.Label LblContiene 
         Caption         =   "Label8"
         Height          =   225
         Left            =   2010
         TabIndex        =   22
         Top             =   2115
         Width           =   1320
      End
      Begin VB.Label LblTamano 
         Caption         =   "Label8"
         Height          =   225
         Left            =   2025
         TabIndex        =   21
         Top             =   1845
         Width           =   1320
      End
      Begin VB.Label LblUbicacion 
         Caption         =   "Label8"
         Height          =   225
         Left            =   2025
         TabIndex        =   20
         Top             =   1515
         Width           =   1320
      End
      Begin VB.Label LblTipo 
         Caption         =   "Label8"
         Height          =   225
         Left            =   2025
         TabIndex        =   19
         Top             =   1215
         Width           =   2220
      End
      Begin VB.Label LblNOmbre 
         Caption         =   "Label8"
         Height          =   210
         Left            =   1905
         TabIndex        =   18
         Top             =   645
         Width           =   2715
      End
      Begin VB.Label Label7 
         Caption         =   "Atributos:"
         Height          =   225
         Left            =   420
         TabIndex        =   11
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Creado:"
         Height          =   225
         Left            =   405
         TabIndex        =   7
         Top             =   2805
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre MS-DOS:"
         Height          =   225
         Left            =   435
         TabIndex        =   6
         Top             =   2550
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Contiene:"
         Height          =   225
         Left            =   450
         TabIndex        =   5
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Tamaño:"
         Height          =   225
         Left            =   435
         TabIndex        =   4
         Top             =   1845
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Ubicación:"
         Height          =   225
         Left            =   435
         TabIndex        =   3
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
         Height          =   270
         Left            =   405
         TabIndex        =   2
         Top             =   1230
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   330
         Picture         =   "FrmPropiedades.frx":0000
         Stretch         =   -1  'True
         Top             =   465
         Width           =   930
      End
   End
   Begin ComctlLib.TabStrip TabPropiedades 
      Height          =   5505
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   9710
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPropiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

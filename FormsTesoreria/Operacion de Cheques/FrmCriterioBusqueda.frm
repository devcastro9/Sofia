VERSION 5.00
Begin VB.Form FrmCriterioBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Criterio de Busqueda"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1575
      TabIndex        =   8
      Top             =   1155
      Width           =   1425
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1155
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6540
      Begin VB.TextBox TxtValor 
         Height          =   315
         Left            =   3885
         TabIndex        =   3
         Top             =   630
         Width           =   2505
      End
      Begin VB.ComboBox CmbCampo 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   2475
      End
      Begin VB.ComboBox CmbOperador 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmCriterioBusqueda.frx":0000
         Left            =   2715
         List            =   "FrmCriterioBusqueda.frx":0019
         TabIndex        =   1
         Text            =   "="
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label LblCampo 
         Caption         =   "Campo"
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   345
         Width           =   615
      End
      Begin VB.Label LblOperador 
         Caption         =   "Operador"
         Height          =   255
         Left            =   2730
         TabIndex        =   5
         Top             =   345
         Width           =   885
      End
      Begin VB.Label LblValor 
         Caption         =   "Valor"
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   345
         Width           =   675
      End
   End
End
Attribute VB_Name = "FrmCriterioBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


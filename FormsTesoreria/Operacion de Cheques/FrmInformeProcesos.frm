VERSION 5.00
Begin VB.Form FrmInformeProcesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Procesamiento"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "FrmInformeProcesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   195
      TabIndex        =   1
      Top             =   3975
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   3945
      Left            =   165
      TabIndex        =   0
      Top             =   -45
      Width           =   7080
      Begin VB.ListBox LstProcesos 
         Height          =   2985
         Left            =   600
         TabIndex        =   2
         Top             =   660
         Width           =   6315
      End
      Begin VB.Label Label1 
         Caption         =   "Cheque/Transf.   Cod. Cuenta                 Obsevaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         TabIndex        =   3
         Top             =   345
         Width           =   6300
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   60
         Picture         =   "FrmInformeProcesos.frx":0ECA
         Top             =   285
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmInformeProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipo As String
Dim Proceso As String

Private Sub cmdCerrar_Click()
  Unload Me
End Sub

Public Sub Principal(pTipo As String, pEstado As String, pCadena As String)
  tipo = pTipo
  Select Case pEstado
         Case "E":
                 Proceso = "ENTREGADOS"
         Case "C":
                 Proceso = "COBRADOS"
         Case "D":
                 Proceso = "DEVUELTOS"
         Case "A":
                 Proceso = "ANULADOS"
  End Select
  LstProcesos.AddItem pCadena
End Sub

Private Sub Form_Load()
  If tipo = "C" Then
    FrmInformeProcesos.Caption = "Informe de cheques >> " & Proceso & " << "
    Label1 = "Cheques   Cuentas                       Obsevaciones"
  Else
    FrmInformeProcesos.Caption = "Informe de transferencias >> " & Proceso & " << "
    Label1 = "Transf.   Cuentas                       Obsevaciones"
  End If
	Call SeguridadSet(Me)
End Sub

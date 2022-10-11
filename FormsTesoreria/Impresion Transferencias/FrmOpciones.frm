VERSION 5.00
Begin VB.Form FrmOpciones 
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   390
      Left            =   825
      TabIndex        =   2
      Top             =   870
      Width           =   2070
   End
   Begin VB.OptionButton OptDol1 
      Caption         =   "Dolares"
      Height          =   525
      Left            =   2145
      TabIndex        =   1
      Top             =   300
      Width           =   1770
   End
   Begin VB.OptionButton OptBol1 
      Caption         =   "Bolivianos"
      Height          =   525
      Left            =   510
      TabIndex        =   0
      Top             =   330
      Width           =   1365
   End
End
Attribute VB_Name = "FrmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
If OptBol1.Value = True Then
    moneda = "1"
End If
If OptDol1.Value = True Then
    moneda = "2"
End If

    Unload Me
End Sub

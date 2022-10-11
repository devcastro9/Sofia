VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPF 
   Caption         =   "Reporte"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   Icon            =   "FrmPF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   180
      TabIndex        =   8
      Top             =   1065
      Width           =   4125
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   285
         Left            =   2610
         TabIndex        =   14
         Top             =   225
         Width           =   1035
      End
      Begin VB.OptionButton OptIngresos 
         Caption         =   "Ingresos"
         Height          =   285
         Left            =   1110
         TabIndex        =   10
         Top             =   225
         Width           =   1035
      End
      Begin VB.OptionButton OptGastos 
         Caption         =   "Gastos"
         Height          =   270
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   2865
      TabIndex        =   15
      Top             =   2355
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   4485
      TabIndex        =   2
      Top             =   0
      Width           =   4545
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Programación Financiera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -225
         TabIndex        =   7
         Top             =   225
         Width           =   4890
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   6
         Top             =   660
         Width           =   1305
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
         TabIndex        =   5
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Index           =   0
         Left            =   1245
         TabIndex        =   4
         Top             =   690
         Width           =   2460
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   675
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "FrmPF.frx":0ECA
         Top             =   0
         Width           =   11640
      End
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   915
      TabIndex        =   11
      Top             =   1650
      Width           =   2295
      Begin VB.OptionButton OptTri1 
         Caption         =   "1er Sem."
         Height          =   270
         Left            =   135
         TabIndex        =   13
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptTri2 
         Caption         =   "2do Sem."
         Height          =   285
         Left            =   1110
         TabIndex        =   12
         Top             =   210
         Width           =   1035
      End
   End
   Begin Crystal.CrystalReport CryPF 
      Left            =   3540
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   465
      Left            =   1455
      TabIndex        =   1
      Top             =   2355
      Width           =   1410
   End
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "Ejecutar"
      Height          =   465
      Left            =   150
      TabIndex        =   0
      Top             =   2355
      Width           =   1305
   End
End
Attribute VB_Name = "FrmPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEjecutar_Click()
Dim comPFinanciera As New ADODB.Command

If OptIngresos.Value = True Then
    op1 = "I"
End If
If OptGastos.Value = True Then
    op1 = "G"
End If

If opttodos.Value = True Then
    op1 = "T"
End If

If OptTri1.Value = True Then
    Tri = "1"
End If
If OptTri2.Value = True Then
    Tri = "2"
End If

    Set comPFinanciera = New ADODB.Command
    With comPFinanciera
        .CommandText = "Cel_Programacion_Financiera_FTe"
        .CommandType = adCmdStoredProc
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 1, op1)
        .Parameters.Append op1
        Set Tri = .CreateParameter("Tri", adVarChar, adParamInput, 1, Tri)
        .Parameters.Append Tri
        .ActiveConnection = db
        .Execute
    End With
    MsgBox "Fin Proceso"
End Sub

Private Sub CmdImprimir_Click()

 If OptIngresos.Value = True Then
        CryPF.Formulas(2) = "FTitulo=' INGRESOS POR FUENTE DE FINANCIAMIENTO '"
 End If
 If OptGastos.Value = True Then
        CryPF.Formulas(2) = "FTitulo=' GASTOS POR FUENTE DE FINANCIAMIENTO '"
 End If

 If opttodos.Value = True Then
        CryPF.Formulas(2) = "FTitulo=' INGRESOS - GASTOS POR FUENTE DE FINANCIAMIENTO '"
 End If
If OptTri1.Value = True Then
        CryPF.ReportFileName = App.Path & "\FormsTesoreria\Rpt_ProgramacionFinanciera1Semestrev1.rpt"
        IResult = CryPF.PrintReport
        If IResult <> 0 Then
           MsgBox CryPF.LastErrorNumber & " : " & CryPF.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
End If
If OptTri2.Value = True Then
      CryPF.ReportFileName = App.Path & "\FormsTesoreria\Rpt_ProgramacionFinanciera2SemestreV1.rpt"
        IResult = CryPF.PrintReport
        If IResult <> 0 Then
           MsgBox CryPF.LastErrorNumber & " : " & CryPF.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

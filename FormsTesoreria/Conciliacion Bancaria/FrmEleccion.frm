VERSION 5.00
Begin VB.Form FrmEleccion 
   Appearance      =   0  'Flat
   BackColor       =   &H00404000&
   Caption         =   "Conciliacion Bancaria"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   FillColor       =   &H80000006&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      ForeColor       =   &H00404000&
      Height          =   2820
      Left            =   -30
      TabIndex        =   0
      Top             =   -60
      Width           =   4170
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         CausesValidation=   0   'False
         Height          =   480
         Left            =   225
         TabIndex        =   5
         Top             =   2205
         Width           =   3630
      End
      Begin VB.CommandButton CmdAño 
         Caption         =   "Año"
         Height          =   480
         Left            =   225
         TabIndex        =   3
         Top             =   1680
         Width           =   3630
      End
      Begin VB.CommandButton CmdMes 
         Caption         =   "Mes"
         Height          =   480
         Left            =   225
         TabIndex        =   2
         Top             =   660
         Width           =   3630
      End
      Begin VB.CommandButton CmdFecha 
         Caption         =   "Fecha"
         Height          =   450
         Left            =   225
         TabIndex        =   1
         Top             =   1185
         Width           =   3630
      End
      Begin VB.Label LblTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   195
         Width           =   3810
      End
      Begin VB.Label LblTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   165
         Width           =   3810
      End
   End
End
Attribute VB_Name = "FrmEleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAño_Click()
   swFiltro = "GESTION"
   If swConciliacion = "CHEQUE" Then
    FrmConciliacion.LblTitulo.Caption = "Conciliacion Bancaria de Cheques por año"
   End If
   If swConciliacion = "TRANSFERENCIA" Then
    FrmConciliacion.LblTitulo.Caption = "Conciliacion Bancaria de Transferencias por año"
   End If
    FrmConciliacion.CmbMes.Enabled = False
    FrmConciliacion.DTPInicio.Enabled = False
    FrmConciliacion.DTPFin.Enabled = False
    FrmConciliacion.CmbAño.Enabled = True
    FrmConciliacion.Show
End Sub

Private Sub CmdFecha_Click()
   swFiltro = "FECHA"
   If swConciliacion = "CHEQUE" Then
    FrmConciliacion.LblTitulo.Caption = "Conciliacion Bancaria de Cheques por fecha"
   End If
   If swConciliacion = "TRANSFERENCIA" Then
    FrmConciliacion.LblTitulo.Caption = "Conciliacion Bancaria de Transferencias por fecha"
   End If
    FrmConciliacion.CmbMes.Enabled = False
    FrmConciliacion.DTPInicio.Enabled = True
    FrmConciliacion.DTPFin.Enabled = True
    FrmConciliacion.CmbAño.Enabled = False
    FrmConciliacion.Show
End Sub

Private Sub CmdMes_Click()
    swFiltro = "MES"
    If swConciliacion = "CHEQUE" Then
    FrmConciliacion.LblTitulo.Caption = "Conciliacion Bancaria de Cheques por mes"
    End If
    If swConciliacion = "TRANSFERENCIA" Then
    FrmConciliacion.LblTitulo.Caption = "Conciliacion Bancaria de Transferencias por mes"
    End If
    'Inhabilitando botones y cajas de diálogo
    FrmConciliacion.CmbMes.Enabled = True
    FrmConciliacion.DTPInicio.Enabled = False
    FrmConciliacion.DTPFin.Enabled = False
    FrmConciliacion.CmbAño.Enabled = False
    'FrmConciliacion.DtCCodigoBanco.Enabled = False
    'FrmConciliacion.DtCDescripcionBanco.Enabled = False
    'FrmConciliacion.DtCCuentaOrigen.Enabled = False
    'FrmConciliacion.DtcCtaTGN.Enabled = False
    'FrmConciliacion.DtCCuentaOrigenDes.Enabled = False
    FrmConciliacion.Show
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


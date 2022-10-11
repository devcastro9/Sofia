VERSION 5.00
Begin VB.Form frmBalanceGral 
   Caption         =   "Balance General"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox AVI 
      Height          =   1215
      Left            =   1080
      ScaleHeight     =   1155
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox CryBalGral 
      Height          =   480
      Left            =   960
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   10
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2160
      Left            =   1290
      TabIndex        =   4
      Top             =   0
      Width           =   2580
      Begin VB.PictureBox DTPfin 
         Height          =   405
         Left            =   945
         ScaleHeight     =   345
         ScaleWidth      =   1350
         TabIndex        =   5
         Top             =   1200
         Width           =   1410
      End
      Begin VB.PictureBox DTPinicio 
         Height          =   405
         Left            =   900
         ScaleHeight     =   345
         ScaleWidth      =   1350
         TabIndex        =   6
         Top             =   375
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Del :"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Al :"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   1290
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2130
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1185
      Begin VB.CommandButton CmdAceptar 
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
         Height          =   780
         Left            =   180
         Picture         =   "FRMBAL~1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   855
      End
      Begin VB.CommandButton Cmdsalir 
         Caption         =   "Salir"
         Height          =   780
         Left            =   195
         Picture         =   "FRMBAL~1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   2280
      Width           =   3735
   End
End
Attribute VB_Name = "frmBalanceGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
    MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
    Exit Sub
End If
'Se manda los parámetros necesarios  al store procedure
Dim IResult As Integer
'Me.ProgressBar1.Visible = True
'Me.ProgressBar1.Value = 0
    'AVI.Open "C:\Archivos de programa\Microsoft Visual Studio\Common\Graphics\Videos\filemove.avi"
    'AVI.Play
    'Usuario = "GABY"
    CryBalGral.Destination = crptToWindow
    CryBalGral.ReportFileName = App.Path & "\Reportes\Contabilidad\Bal_General\CryBalGeneral.rpt"
    CryBalGral.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
    CryBalGral.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
    CryBalGral.StoredProcParam(2) = Trim(GlMaquina)   'NOMBRE DE USUARIO
    CryBalGral.Formulas(0) = "Fecha_AInicio ='" & Me.DTPinicio.Value & "'"
    CryBalGral.Formulas(1) = "Fecha_Final ='" & Me.DTPfin.Value & "'"
    CryBalGral.SelectionFormula = "{BalGeneral;1.usr}='" & GlMaquina & "'"
    IResult = CryBalGral.PrintReport
    If IResult > 0 Then
        MsgBox CryBalGral.LastErrorNumber & " : " & CryBalGral.LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
    'AVI.Stop
    'AVI.Close
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    'frmprincipal.Show
End Sub
Private Sub DTPfin_Validate(Cancel As Boolean)
If DTPfin.Value < DTPinicio.Value Then
    MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
    DTPfin.SetFocus
End If
End Sub

Private Sub DTPinicio_Validate(Cancel As Boolean)
    If DTPinicio.Value > DTPfin.Value Then
        MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
        DTPfin.SetFocus
    End If
End Sub

Private Sub Form_Load()
  Me.DTPinicio.MinDate = CDate("01/01/2000")
  Me.DTPinicio.Value = CDate("01/01/2000")
  Me.DTPinicio.MinDate = CDate("01/01/2000")
  Me.DTPfin.MinDate = Me.DTPinicio.Value
  Me.DTPfin.Value = Date
  Me.DTPfin.MaxDate = Date
  Me.DTPinicio.MaxDate = Date
  Me.ProgressBar1.Min = 0
  Me.ProgressBar1.Visible = False
End Sub

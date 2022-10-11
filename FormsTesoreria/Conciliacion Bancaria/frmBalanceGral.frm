VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmBalanceGral 
   Caption         =   "Balance General"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryBalGral 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   1320
      TabIndex        =   10
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton opttodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton optctasmovim 
         Caption         =   "Cuentas con movimiento"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComCtl2.Animation AVI 
      Height          =   615
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   145
      FullHeight      =   41
   End
   Begin Crystal.CrystalReport CryBalGralSaldos 
      Left            =   480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   2505
      Begin MSComCtl2.DTPicker DTPfin 
         Height          =   405
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   714
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36633
      End
      Begin MSComCtl2.DTPicker DTPinicio 
         Height          =   405
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   714
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36633
      End
      Begin VB.Label Label1 
         Caption         =   "Del :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Al :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
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
         Picture         =   "frmBalanceGral.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   855
      End
      Begin VB.CommandButton Cmdsalir 
         Caption         =   "Salir"
         Height          =   780
         Left            =   195
         Picture         =   "frmBalanceGral.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmBalanceGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
    Dim IResult As Integer
    If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
        MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
        Exit Sub
    End If
    If Me.opttodas.Value = True Then
    'Se manda los parámetros necesarios  al store procedure
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
     End If
     If optctasmovim.Value = True Then
        CryBalGralSaldos.Destination = crptToWindow
        CryBalGralSaldos.ReportFileName = App.Path & "\Reportes\Contabilidad\Bal_General\CryBalGeneralSaldos.rpt"
        CryBalGralSaldos.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
        CryBalGralSaldos.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
        CryBalGralSaldos.StoredProcParam(2) = Trim(GlMaquina)   'NOMBRE DE USUARIO
        CryBalGralSaldos.Formulas(0) = "Fecha_AInicio ='" & Me.DTPinicio.Value & "'"
        CryBalGralSaldos.Formulas(1) = "Fecha_Final ='" & Me.DTPfin.Value & "'"
        CryBalGralSaldos.SelectionFormula = "{BalGeneralSaldos;1.Usr}='" & GlMaquina & "'"
        IResult = CryBalGralSaldos.PrintReport
        If IResult > 0 Then
            MsgBox CryBalGralSaldos.LastErrorNumber & " : " & CryBalGral.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
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


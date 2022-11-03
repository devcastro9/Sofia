VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBalanceGral 
   Caption         =   "Reportes Contabilidad"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   Icon            =   "frmBalanceGral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   0
      TabIndex        =   7
      Top             =   990
      Width           =   1035
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   675
         Left            =   120
         Picture         =   "frmBalanceGral.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   750
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   645
         Left            =   120
         Picture         =   "frmBalanceGral.frx":0794
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   765
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   5
      Top             =   60
      Width           =   4030
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE GENERAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   690
         TabIndex        =   6
         Top             =   285
         Width           =   2625
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   1030
      TabIndex        =   0
      Top             =   990
      Width           =   3000
      Begin Crystal.CrystalReport CryBalGral 
         Left            =   2040
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.OptionButton optctasmovim 
         Caption         =   "Con movimiento"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   420
         Width           =   1455
      End
      Begin VB.OptionButton opttodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   300
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPfin 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   2220
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83820545
         CurrentDate     =   43100
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTPinicio 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   1740
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83820545
         CurrentDate     =   42736
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo DtcDenom_Moneda 
         Bindings        =   "frmBalanceGral.frx":0E7E
         DataField       =   "denominacion_moneda"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1095
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_moneda"
         BoundColumn     =   "Tipo_moneda"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo DtCCod_Moneda 
         Bindings        =   "frmBalanceGral.frx":0E97
         DataField       =   "tipo_moneda"
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipo_moneda"
         BoundColumn     =   "denominacion_moneda"
         Text            =   "DataCombo19"
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   1680
         X2              =   2880
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir Cuentas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1500
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   1320
         X2              =   2760
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimientos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1510
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   1560
         X2              =   2880
         Y1              =   920
         Y2              =   920
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Moneda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   800
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   900
         TabIndex        =   4
         Top             =   1860
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   900
         TabIndex        =   3
         Top             =   2280
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmBalanceGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim CCRepmoneda As String
Private Sub Cmdimprimir_Click()
If DtcDenom_Moneda.Text = "" Then
  MsgBox "Seleccione una moneda", vbExclamation + vbDefaultButton1, "Atención!!!!!!!"
  Exit Sub
End If
    Select Case DtcDenom_Moneda.Text
      Case "BOLIVIANOS"
        CCRepmoneda = "Bs"
      Case "DOLARES AMERICANOS"
        CCRepmoneda = "Sus"
    End Select
    Dim IResult As Integer
    If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
        MsgBox "Seleccione un rango de fechas correcto", vbExclamation + vbDefaultButton1, "Atención!!!"
        Exit Sub
    End If
  '  If Me.opttodas.Value = True Then
    'Se manda los parámetros necesarios  al store procedure
    'Me.ProgressBar1.Visible = True
    'Me.ProgressBar1.Value = 0
    'AVI.Open "C:\Archivos de programa\Microsoft Visual Studio\Common\Graphics\Videos\filemove.avi"
    'AVI.Play
    'Usuario = "ADMIN"
        CryBalGral.Destination = crptToWindow
        CryBalGral.WindowShowPrintSetupBtn = True
        CryBalGral.WindowShowSearchBtn = True
        CryBalGral.WindowState = crptMaximized
        CryBalGral.ReportFileName = App.Path & "\Reportes\Contabilidad\CryBalGeneralnewnew.rpt"
        CryBalGral.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
        CryBalGral.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
        CryBalGral.StoredProcParam(2) = Trim(GlMaquina)   'NOMBRE DE USUARIO
        CryBalGral.StoredProcParam(3) = CCRepmoneda
        If opttodas.Value = True Then CryBalGral.StoredProcParam(4) = 0
        If optctasmovim.Value = True Then CryBalGral.StoredProcParam(4) = 1
        CryBalGral.Formulas(0) = "Fecha_AInicio ='" & Me.DTPinicio.Value & "'"
        CryBalGral.Formulas(1) = "Fecha_Final ='" & Me.DTPfin.Value & "'"
        CryBalGral.SelectionFormula = "{BalGeneral;1.usr}='" & GlMaquina & "'"
        IResult = CryBalGral.PrintReport
        If IResult > 0 Then
            MsgBox CryBalGral.LastErrorNumber & " : " & CryBalGral.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
     'End If
'     If optctasmovim.Value = True Then
'        CryBalGralSaldos.Destination = crptToWindow
'        CryBalGralSaldos.WindowShowPrintSetupBtn = True
'        CryBalGralSaldos.WindowShowSearchBtn = True
'        CryBalGralSaldos.WindowState = crptMaximized
'        CryBalGralSaldos.ReportFileName = App.Path & "\Reportes\Contabilidad\Bal_General\CryBalGeneralSaldos.rpt"
'        CryBalGralSaldos.StoredProcParam(0) = Format(Me.DTPInicio.Value, "dd/mm/yyyy")
'        CryBalGralSaldos.StoredProcParam(1) = Format(Me.DTPFin.Value, "dd/mm/yyyy")
'        CryBalGralSaldos.StoredProcParam(2) = Trim(GlMaquina)   'NOMBRE DE USUARIO
'        CryBalGralSaldos.Formulas(0) = "Fecha_AInicio ='" & Me.DTPInicio.Value & "'"
'        CryBalGralSaldos.Formulas(1) = "Fecha_Final ='" & Me.DTPFin.Value & "'"
'        CryBalGralSaldos.SelectionFormula = "{BalGeneralSaldos;1.Usr}='" & GlMaquina & "'"
'        iResult = CryBalGralSaldos.PrintReport
'        If iResult > 0 Then
'            MsgBox CryBalGralSaldos.LastErrorNumber & " : " & CryBalGral.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'     End If
'    'AVI.Stop
'    'AVI.Close
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    'frmprincipal.Show
End Sub

Private Sub DtCCod_Moneda_Click(Area As Integer)
  DtcDenom_Moneda.Text = DtCCod_Moneda.BoundText
End Sub
Private Sub DtcDenom_Moneda_Click(Area As Integer)
  DtCCod_Moneda.Text = DtcDenom_Moneda.BoundText
End Sub

Private Sub DTPfin_Validate(Cancel As Boolean)
If DTPfin.Value < DTPinicio.Value Then
    MsgBox "Seleccione un rango de fechas correcto", vbExclamation + vbDefaultButton1, "Atención!!!"
    DTPfin.SetFocus
End If
End Sub

Private Sub DTPinicio_Validate(Cancel As Boolean)
    If DTPinicio.Value > DTPfin.Value Then
        MsgBox "Seleccione un rango de fechas correcto", vbExclamation + vbDefaultButton1, "Atención!!!"
        DTPfin.SetFocus
    End If
End Sub

Private Sub Form_Load()
'-------------Límites de rangos de las fechas
 'DTPInicio.MinDate = RepFechInicio
 'DTPFin.MinDate = RepFechInicio
 'DTPInicio.MaxDate = RepFechFinal
 'DTPFin.MaxDate = RepFechFinal
 DTPinicio.Value = "01/01/2002"
 DTPfin.Value = "31/12/2002"
  '---------Tipo de Moneda
  sql1 = "select * from tipo_moneda"
  Set DtCCod_Moneda.RowSource = db.Execute(sql1, , cmdtext)
  Set DtcDenom_Moneda.RowSource = db.Execute(sql1, , cmdtext)
  opttodas.Value = True
	Call SeguridadSet(Me)
End Sub

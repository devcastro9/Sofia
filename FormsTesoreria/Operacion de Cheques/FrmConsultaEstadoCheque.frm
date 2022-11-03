VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmConsultaEstadoCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Rapida de Estado de ..."
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "FrmConsultaEstadoCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OptTransf 
      Caption         =   "Transferencia"
      Height          =   240
      Left            =   1335
      TabIndex        =   26
      Top             =   60
      Width           =   1395
   End
   Begin VB.OptionButton OptCheque 
      Caption         =   "Cheque"
      Height          =   300
      Left            =   105
      TabIndex        =   25
      Top             =   30
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.Frame Frame2 
      Caption         =   " Estados "
      Height          =   2040
      Left            =   75
      TabIndex        =   7
      Top             =   1890
      Width           =   8100
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   330
         Left            =   6540
         TabIndex        =   24
         Top             =   1530
         Width           =   1260
      End
      Begin VB.Line Line1 
         X1              =   435
         X2              =   5640
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Index           =   4
         Left            =   2745
         TabIndex        =   23
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Index           =   3
         Left            =   2745
         TabIndex        =   22
         Top             =   1419
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Index           =   2
         Left            =   2745
         TabIndex        =   21
         Top             =   1161
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Index           =   1
         Left            =   2745
         TabIndex        =   20
         Top             =   903
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Index           =   0
         Left            =   2745
         TabIndex        =   19
         Top             =   645
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   270
         Index           =   4
         Left            =   4305
         TabIndex        =   18
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   270
         Index           =   3
         Left            =   4305
         TabIndex        =   17
         Top             =   1425
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   270
         Index           =   2
         Left            =   4305
         TabIndex        =   16
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   270
         Index           =   1
         Left            =   4305
         TabIndex        =   15
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   270
         Index           =   0
         Left            =   4305
         TabIndex        =   14
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Anulado"
         Height          =   195
         Index           =   4
         Left            =   435
         TabIndex        =   13
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Devuelto"
         Height          =   195
         Index           =   3
         Left            =   435
         TabIndex        =   12
         Top             =   1425
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cobrado"
         Height          =   195
         Index           =   2
         Left            =   435
         TabIndex        =   11
         Top             =   1155
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entregado"
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   10
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Impresion"
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   9
         Top             =   645
         Width           =   675
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion Estado                     Estado                        Fecha Registro"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   390
         TabIndex        =   8
         Top             =   270
         Width           =   5220
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   8115
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   330
         Left            =   6540
         TabIndex        =   27
         Top             =   1005
         Width           =   1260
      End
      Begin VB.TextBox TxtCheques 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1035
         Width           =   1875
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "FrmConsultaEstadoCheque.frx":0ECA
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   165
         TabIndex        =   2
         Top             =   390
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "FrmConsultaEstadoCheque.frx":0EE2
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   3450
         TabIndex        =   3
         Top             =   390
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "FrmConsultaEstadoCheque.frx":0EFA
         DataField       =   "cta_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   2055
         TabIndex        =   4
         Top             =   390
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc AdoCuenta 
         Height          =   390
         Left            =   5175
         Top             =   180
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   688
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Cuenta Bancaria"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cheque"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   795
         Width           =   555
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Cuenta "
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   195
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmConsultaEstadoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscuenta As New ADODB.Recordset
Dim rscheques As New ADODB.Recordset

Private Sub cmdCerrar_Click()
  Unload Me
End Sub

Private Sub cmdProcesar_Click()
If TxtCheques <> "" Then
  If OptCheque.Value = True Then
    rscheques.Open "Select * From to_cheques_operaciones Where numero_cheque='" & TxtCheques & "' And cta_codigo='" & DtCCuentaOrigen.Text & "' And Cheq_Transf='C'", db, adOpenStatic
  Else
    rscheques.Open "Select * From to_cheques_operaciones Where numero_cheque='" & TxtCheques & "' And cta_codigo='" & DtCCuentaOrigen.Text & "' And Cheq_Transf='T'", db, adOpenStatic
  End If
  If rscheques.RecordCount > 0 Then
    label3(0) = rscheques!estado_impreso
    Label4(0) = IIf(IsNull(rscheques!fecha_impreso), "Null", rscheques!fecha_impreso)
    label3(1) = rscheques!estado_entregado
    Label4(1) = IIf(IsNull(rscheques!fecha_entregado), "Null", rscheques!fecha_entregado)
    label3(2) = rscheques!estado_cobrado
    Label4(2) = IIf(IsNull(rscheques!fecha_cobrado), "Null", rscheques!fecha_cobrado)
    label3(3) = rscheques!estado_devuelto
    Label4(3) = IIf(IsNull(rscheques!fecha_devuelto), "Null", rscheques!fecha_devuelto)
    label3(4) = rscheques!estado_anulado
    Label4(4) = IIf(IsNull(rscheques!fecha_anulado), "Null", rscheques!fecha_anulado)
  Else
    MsgBox "El Numero de cheque o transferencia no existe, verifique datos " & vbCr & "de Codigo de Cuenta y numero de Cheque o Tranferencia", vbCritical + vbOKOnly, "Atencion"
    LimpiaDatosEstado
  End If
  rscheques.Close
Else
  MsgBox "Ingrese el numero o identifacion del cheque o tranferencia", vbInformation + vbOKOnly, "Atencion"
  LimpiaDatosEstado
End If
End Sub

Private Sub Form_Load()
    'Abriendo cuenta bancaria
    Set rscuenta = New ADODB.Recordset
    rscuenta.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    Set AdoCuenta.Recordset = rscuenta
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
    'Limpia los datos de estado
    LimpiaDatosEstado
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If AdoCuenta.Recordset.EditMode <> 2 Then AdoCuenta.Recordset.CancelUpdate
  AdoCuenta.Recordset.Close
End Sub

Private Sub OptCheque_Click()
  Label10 = "Cheque"
End Sub

Private Sub OptTransf_Click()
  Label10 = "Transferencia"
End Sub

'Private Sub TxtCheques_KeyPress(KeyAscii As Integer)
'Dim bandera As Integer
'Dim i As Integer
'    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 45 Or KeyAscii = 47 Or KeyAscii = 8 Then
'    Else
'       KeyAscii = Asc(UCase(Chr(0)))
'    End If
'End Sub

Private Sub LimpiaDatosEstado()
Dim i As Byte
  For i = 0 To 4
    label3(i) = ""
    Label4(i) = ""
  Next i
End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
   DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
   DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
End Sub


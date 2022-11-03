VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmSaldosReales 
   Caption         =   "Saldos Reales"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrySaldoActual 
      Left            =   1365
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   5250
      TabIndex        =   1
      Top             =   0
      Width           =   5310
      Begin VB.Label Label2 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   75
         TabIndex        =   7
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1500
         TabIndex        =   6
         Top             =   615
         Width           =   2460
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7740
         TabIndex        =   5
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label Label7 
         Height          =   225
         Left            =   10485
         TabIndex        =   4
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "C O M P R O B  A N  T E -  C O N T A B L E -  M A N U A L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2955
         TabIndex        =   3
         Top             =   1125
         Width           =   8415
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "RESUMEN  SALDOS REALES DE CUENTAS BANCARIAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1545
         TabIndex        =   2
         Top             =   270
         Width           =   7245
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   6435
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   1260
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir "
         Height          =   705
         Left            =   165
         Picture         =   "FrmSaldosReales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   165
         Picture         =   "FrmSaldosReales.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1665
         Width           =   930
      End
      Begin VB.CommandButton CmdTesoreria 
         Caption         =   "Tesoreria Actualiza"
         Height          =   690
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid DtGCuentaBancaria 
      Height          =   6375
      Left            =   1335
      TabIndex        =   0
      Top             =   1200
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmSaldosReales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCB As New ADODB.Recordset
Dim rsCBria As New ADODB.Recordset
Dim rsPg As New ADODB.Recordset


Private Sub CmdImprimir_Click()
            CrySaldoActual.ReportFileName = "C:\SAF-2000\FormsTesoreria\Movimiento de Cuentas\Impresiones\Rpt_SaldoActual.rpt"
            IResult = CrySaldoActual.PrintReport
            If IResult <> 0 Then
                MsgBox CryMovi.LastErrorNumber & " : " & CryMovi.LastErrorString, vbCritical + vbOKOnly, "Error..."
            End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTesoreria_Click()
'Realiza la actualizacion de la Cta_Saldo actual de tesoreriaSet rsCuenta = New ADODB.Recordset
Dim suma As Variant
Dim sumartf As Variant

    MsgBox "Esperar mensaje de término"
    If rsCBria.State = 1 Then rsCBria.Close
    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic

    While Not rsCBria.EOF
                suma = 0
                If rsPg.State = 1 Then rsPg.Close
                rsPg.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & rsCBria("Cta_Codigo") & "'", db, adOpenKeyset, adLockOptimistic
                While Not rsPg.EOF
                      If Not IsNull(rsPg("monto_bolivianos")) Then
                          suma = suma + rsPg("monto_bolivianos")
                      End If
                      rsPg.MoveNext
                Wend
                rsCBria("Cta_Acumulado") = suma
                rsCBria.Update
                rsCBria.MoveNext
    Wend
'MsgBox "T E R M I N Ó  S I N  T R P"



''Caso traspasos
'    If rsCBria.State = 1 Then rsCBria.Close
'    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
'    While Not rsCBria.EOF
'                sumartf = 0
'                If rsPg.State = 1 Then rsPg.Close
'                rsPg.Open "SELECT pagos.org_codigo AS Expr1, pago_detalle.*, pagos.* " & _
'                "FROM pago_detalle INNER JOIN pagos ON pago_detalle.Ges_gestion = pagos.ges_gestion AND " & _
'                "pago_detalle.org_codigo = pagos.org_codigo AND pago_detalle.codigo_pago = pagos.codigo_pago WHERE cta_codigo_destino='" & rsCBria("Cta_Codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'                While Not rsPg.EOF
'                      If Not IsNull(rsPg("monto_Bolivianos")) And rsPg("tipo_comp") = "TRP" Then
'                          sumartf = sumartf + rsPg("monto_bolivianos")
'                      End If
'                      rsPg.MoveNext
'                Wend
'                If rsCBria("Cta_codigo") = "0922" Then
'                    MsgBox sumartf
'                End If
'                rsCBria("Cta_Acumulado") = rsCBria("Cta_Acumulado") + sumartf
'                rsCBria.Update
'                rsCBria.MoveNext
'
'    Wend

'DETERMINANDO SALDO ACTUAL
Dim X As Double
    If rsCBria.State = 1 Then rsCBria.Close
    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
    While Not rsCBria.EOF
    
                'x = rsCBria("Cta_saldo_inicial") - rsCBria("Cta_Acumulado") + rsCBria("Cta_Pco_Debe") - rsCBria("Cta_Pco_Haber") + rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe")
                'If rsCBria("Cta_codigo") = "0922" Then
                        'rsCBria("Cta_Saldo_Actual") = rsCBria("Cta_saldo_inicial") - rsCBria("Cta_Acumulado") + rsCBria("Cta_Pco_Debe") - rsCBria("Cta_Pco_Haber") + rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe")
                        rsCBria("Cta_Saldo_Actual") = (rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe") + rsCBria("Cta_saldo_inicial") + rsCBria("Cta_Pco_Debe")) - (rsCBria("Cta_Acumulado")) - rsCBria("Cta_Pco_Haber")
                        rsCBria.Update
                'End If

                    'MsgBox rsCBria("Cta_Saldo_Actual")
                
                rsCBria.MoveNext
                
    Wend
MsgBox "T E R M I N Ó  S A L D O   A C T U A L"
'db.ActualizaCtaBco2
'MsgBox "T E R M I N Ó  S A L D O   A C T U A L"

If rsCBria.State = 1 Then rsCBria.Close
rsCBria.Open "SELECT CTA_CODIGO, CTA_CODIGO_TGN, CTA_DESCRIPCION_LARGA, CTA_SALDO_INICIAL, CTA_SALDO_ACTUAL  FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
Set DtGCuentaBancaria.DataSource = rsCBria
    
End Sub

Private Sub Form_Load()
    If rsCB.State = 1 Then rsCB.Close
    rsCB.Open "SELECT CTA_CODIGO, CTA_CODIGO_TGN, CTA_DESCRIPCION_LARGA, CTA_SALDO_INICIAL, CTA_SALDO_ACTUAL FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
    If rsCB.RecordCount > 0 Then
        Set DtGCuentaBancaria.DataSource = rsCB
    Else
        Set DtGCuentaBancaria.DataSource = rsNada
        MsgBox "No existen registros", vbInformation + vbCritical, "Validaciòn de datos"
        Exit Sub
    End If

	Call SeguridadSet(Me)
End Sub

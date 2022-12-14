VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmSaldosBancarios 
   Caption         =   "Saldos Bancarios Actuales"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "FrmSaldosBancarios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FraOpciones 
      Height          =   6390
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1260
      Begin VB.CommandButton CmdDatosRecientes 
         Caption         =   "Actualizar Datos"
         Height          =   795
         Left            =   165
         TabIndex        =   12
         Top             =   1860
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton CmdTesoreria 
         Caption         =   "Actualizar Saldos"
         Height          =   795
         Left            =   165
         TabIndex        =   10
         Top             =   1065
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   165
         Picture         =   "FrmSaldosBancarios.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2640
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir "
         Height          =   795
         Left            =   165
         Picture         =   "FrmSaldosBancarios.frx":130C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   1545
         TabIndex        =   6
         Top             =   270
         Width           =   7245
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
         TabIndex        =   5
         Top             =   1125
         Width           =   8415
      End
      Begin VB.Label Label7 
         Height          =   225
         Left            =   10485
         TabIndex        =   4
         Top             =   660
         Width           =   1305
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7740
         TabIndex        =   3
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1500
         TabIndex        =   2
         Top             =   615
         Width           =   2460
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   75
         TabIndex        =   1
         Top             =   630
         Width           =   1110
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "FrmSaldosBancarios.frx":1976
         Top             =   0
         Width           =   11640
      End
   End
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
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSDataGridLib.DataGrid DtGCuentaBancaria 
      Height          =   6315
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   11139
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
Attribute VB_Name = "FrmSaldosBancarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCB As New ADODB.Recordset
Dim rsCBria As New ADODB.Recordset
Dim rsPg As New ADODB.Recordset


Private Sub CmdDatosRecientes_Click()
Dim rsGTZFiltro As New ADODB.Recordset
Dim rsCuen As New ADODB.Recordset
Dim rsGTZ As New ADODB.Recordset
Dim str1 As String
Dim rsMoviReal As New ADODB.Recordset

MsgBox "Esperar mensaje de terminado", vbCritical + vbDefaultButton1, "Mensaje Importante!!!"
    'Actualizando datos de GTZ
    db.Execute "delete from fc_datosGTZ"
    db.movimiento_Cuenta_Bancaria
    
    db.Execute "DELETE FROM To_MOvimientoRealTodos"

    
If rsCuen.State = 1 Then rsCuen.Close
rsCuen.Open "SELECT * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
If rsCuen.RecordCount > 0 Then
While Not rsCuen.EOF
    Set rsGTZFiltro = New ADODB.Recordset
    Set rsMoviRe6al = New ADODB.Recordset
    
        If rsMoviReal.State = 1 Then rsMoviReal.Close
        rsMoviReal.Open "select * from To_MOvimientoRealTodos order by fecha_pago ", db, adOpenKeyset, adLockOptimistic
        With rsGTZ
           If .State = adStateOpen Then
             .Close
           End If
           str1 = "select * from fc_datosGTZ  where cta_codigo= '" & rsCuen("cta_codigo") & "'  order by fecha_pago"
           .Open str1, db, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
             
             While Not .EOF
                         'Set DtGGTZ.DataSource = rsGTZ
                         rsMoviReal.AddNew
                         rsMoviReal("Nro_Cmpte") = rsGTZ("Nro_Cmpte")
                         rsMoviReal("Organismo") = rsGTZ("Organismo")
                         If Not IsNull(rsGTZ("Fecha_Pago")) Then rsMoviReal("Fecha_Pago") = Format(rsGTZ("Fecha_Pago"), "dd/mm/yyyy")
                         rsMoviReal("Monto") = rsGTZ("Monto")
                         rsMoviReal("MontoH") = rsGTZ("MontoH")
                         rsMoviReal("Cambio") = rsGTZ("Cambio")
                         rsMoviReal("Beneficiario") = rsGTZ("Beneficiario")
                         rsMoviReal("Nro_Doc") = rsGTZ("Nro_Doc")
                         rsMoviReal("Transf_Cheq") = rsGTZ("Transf_Cheq")
                         rsMoviReal("Cta_Codigo") = rsGTZ("Cta_Codigo")
                         rsMoviReal("Nombre_Cta") = rsGTZ("Nombre_Cta")
                         rsMoviReal("Bco_Codigo") = rsGTZ("Bco_Codigo")
                         rsMoviReal("justificacion") = rsGTZ("justificacion")
                         rsMoviReal("procedencia") = rsGTZ("procedencia")
                         rsMoviReal.Update
                         .MoveNext
             Wend
           End If
           
           If .State = 1 Then .Close
           str1 = "select * from fc_datosGTZ  where cta_codigo_destino= '" & rsCuen("cta_codigo") & "' and tipo_comp='TRP' order by fecha_pago"
           .Open str1, db, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
             While Not .EOF
                    
                        'Set DtGGTZ.DataSource = rsGTZ
                         rsMoviReal.AddNew
                         rsMoviReal("Nro_Cmpte") = rsGTZ("Nro_Cmpte")
                         rsMoviReal("Organismo") = rsGTZ("Organismo")
                         rsMoviReal("Fecha_Pago") = Format(rsGTZ("Fecha_Pago"), "dd/mm/yyyy")
                         rsMoviReal("Monto") = rsGTZ("Monto")
                         rsMoviReal("MontoH") = rsGTZ("MontoH")
                         rsMoviReal("Cambio") = rsGTZ("Cambio")
                         rsMoviReal("Beneficiario") = rsGTZ("Beneficiario")
                         rsMoviReal("Nro_Doc") = rsGTZ("Nro_Doc")
                         rsMoviReal("Transf_Cheq") = rsGTZ("Transf_Cheq")
                         rsMoviReal("Cta_Codigo") = rsGTZ("Cta_Codigo_destino")
                         rsMoviReal("Nombre_Cta") = rsGTZ("Nombre_Cta")
                         rsMoviReal("Bco_Codigo") = rsGTZ("Bco_Codigo")
                         rsMoviReal("justificacion") = rsGTZ("justificacion")
                         rsMoviReal("procedencia") = "4"
                         rsMoviReal.Update
                         .MoveNext
             Wend
           End If
          End With
  rsCuen.MoveNext
Wend

End If
MsgBox "Proceso terminado", vbCritical + vbDefaultButton1, "Mensaje Importante!!!"
End Sub


Private Sub cmdImprimir_Click()
            CrySaldoActual.ReportFileName = App.Path & "\FormsTesoreria\Movimiento de Cuentas\Impresiones\Rpt_SaldoActual.rpt"
            iResult = CrySaldoActual.PrintReport
            If iResult <> 0 Then
                MsgBox CryMovi.LastErrorNumber & " : " & CryMovi.LastErrorString, vbCritical + vbOKOnly, "Error..."
            End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTesoreria_Click()
            '''''''Realiza la actualizacion de la Cta_Saldo actual de tesoreriaSet rsCuenta = New ADODB.Recordset
            ''''''Dim suma As Variant
            ''''''Dim sumartf As Variant
            ''''''
            ''''''    MsgBox "Esperar mensaje de t?rmino"
            ''''''    db.Cel_Actualizar_DAR
            ''''''    If rsCBria.State = 1 Then rsCBria.Close
            ''''''    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
            ''''''
            ''''''    While Not rsCBria.EOF
            ''''''                suma = 0
            ''''''                If rsPg.State = 1 Then rsPg.Close
            ''''''                'rsPg.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & rsCBria("Cta_Codigo") & "'", db, adOpenKeyset, adLockOptimistic
            ''''''                rsPg.Open "SELECT  * FROM pagos INNER JOIN pago_detalle ON pagos.ges_gestion = pago_detalle.Ges_gestion AND pagos.org_codigo = pago_detalle.org_codigo AND pagos.codigo_pago = pago_detalle.codigo_pago WHERE pago_detalle.cta_codigo='" & rsCBria("Cta_Codigo") & "' and pagos.Estado_Pagado='S'", db, adOpenKeyset, adLockOptimistic
            ''''''                While Not rsPg.EOF
            ''''''                      If Not IsNull(rsPg("monto_bolivianos")) Then
            ''''''                          suma = suma + rsPg("monto_bolivianos")
            ''''''                      End If
            ''''''                      rsPg.MoveNext
            ''''''                Wend
            ''''''
            ''''''                rsCBria("Cta_Acumulado") = suma
            ''''''                rsCBria.Update
            ''''''                rsCBria.MoveNext
            ''''''    Wend
            '''''''MsgBox "T E R M I N ?  S I N  T R P"
            ''''''
            ''''''
            ''''''
            ''''''''Caso traspasos
            '''''''    If rsCBria.State = 1 Then rsCBria.Close
            '''''''    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
            '''''''    While Not rsCBria.EOF
            '''''''                sumartf = 0
            '''''''                If rsPg.State = 1 Then rsPg.Close
            '''''''                rsPg.Open "SELECT pagos.org_codigo AS Expr1, pago_detalle.*, pagos.* " & _
            '''''''                "FROM pago_detalle INNER JOIN pagos ON pago_detalle.Ges_gestion = pagos.ges_gestion AND " & _
            '''''''                "pago_detalle.org_codigo = pagos.org_codigo AND pago_detalle.codigo_pago = pagos.codigo_pago WHERE cta_codigo_destino='" & rsCBria("Cta_Codigo") & "' ", db, adOpenKeyset, adLockOptimistic
            '''''''                While Not rsPg.EOF
            '''''''                      If Not IsNull(rsPg("monto_Bolivianos")) And rsPg("tipo_comp") = "TRP" Then
            '''''''                          sumartf = sumartf + rsPg("monto_bolivianos")
            '''''''                      End If
            '''''''                      rsPg.MoveNext
            '''''''                Wend
            '''''''                If rsCBria("Cta_codigo") = "0922" Then
            '''''''                    MsgBox sumartf
            '''''''                End If
            '''''''                rsCBria("Cta_Acumulado") = rsCBria("Cta_Acumulado") + sumartf
            '''''''                rsCBria.Update
            '''''''                rsCBria.MoveNext
            '''''''
            '''''''    Wend
            ''''''
            '''''''DETERMINANDO SALDO ACTUAL
            ''''''Dim X As Double
            ''''''    If rsCBria.State = 1 Then rsCBria.Close
            ''''''    rsCBria.Open "SELECT * FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
            ''''''    While Not rsCBria.EOF
            ''''''
            ''''''                'x = rsCBria("Cta_saldo_inicial") - rsCBria("Cta_Acumulado") + rsCBria("Cta_Pco_Debe") - rsCBria("Cta_Pco_Haber") + rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe")
            ''''''                'If rsCBria("Cta_codigo") = "0922" Then
            ''''''                'rsCBria("Cta_Saldo_Actual") = rsCBria("Cta_saldo_inicial") - rsCBria("Cta_Acumulado") + rsCBria("Cta_Pco_Debe") - rsCBria("Cta_Pco_Haber") + rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe")
            ''''''                 rsCBria("Cta_Saldo_Actual") = (rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe") + rsCBria("Cta_saldo_inicial") + rsCBria("Cta_Pco_Debe")) - (rsCBria("Cta_Acumulado")) - rsCBria("Cta_Pco_Haber") + rsCBria("Cta_Acum_dev") + rsCBria("Cta_Acum_anl")
            ''''''                 '''rsCBria("Cta_Saldo_Actual") = (rsCBria("Cta_Ingresos") + rsCBria("Cta_Saldo_Debe") + rsCBria("Cta_saldo_inicial") + rsCBria("Cta_Pco_Debe")) - (rsCBria("Cta_Acumulado")) - rsCBria("Cta_Pco_Haber")
            ''''''                 rsCBria.Update
            ''''''                'End If
            ''''''
            ''''''                    'MsgBox rsCBria("Cta_Saldo_Actual")
            ''''''
            ''''''                rsCBria.MoveNext
            ''''''
            ''''''    Wend
            ''''''MsgBox "T E R M I N ?  S A L D O   A C T U A L"
            '''''''db.ActualizaCtaBco2
            '''''''MsgBox "T E R M I N ?  S A L D O   A C T U A L"
            ''''''
            ''''''If rsCBria.State = 1 Then rsCBria.Close
            ''''''rsCBria.Open "SELECT CTA_CODIGO, CTA_CODIGO_TGN, CTA_DESCRIPCION_LARGA, CTA_SALDO_INICIAL, CTA_SALDO_ACTUAL  FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
            ''''''Set DtGCuentaBancaria.DataSource = rsCBria
    
    
Dim MONTO As Double
Dim rsGTZ As New ADODB.Recordset
Dim rsctabancaria As New ADODB.Recordset

'db.Execute "DELETE FROM TO_MOVIMIENTOREALTODOS"
MsgBox "Espere fin de proceso"
If rsctabancaria.State = 1 Then rsctabancaria.Close
rsctabancaria.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
If rsctabancaria.RecordCount > 0 Then
   While Not rsctabancaria.EOF
        With rsGTZ
           If .State = adStateOpen Then
             .Close
           End If
           'str1 = "se
           ''str1 = "SELECT  * from to_movimientoRealTodos  where cta_codigo= '1-297916'  order by fecha_pago"
           str1 = "select * from to_movimientoRealTodos where cta_codigo= '" & rsctabancaria("Cta_codigo") & "' order by fecha_pago"
           .Open str1, db, adOpenKeyset, adLockOptimistic
           MONTO = 0
           If .RecordCount > 0 Then
               'MsgBox .RecordCount
             While Not .EOF
                If rsGTZ("procedencia") = "4" Then 'Traspasos
                    MONTO = MONTO + (rsGTZ("Monto") * (1))
                End If
                If rsGTZ("procedencia") = "1" Then 'Gastos
                  If Not IsNull(rsGTZ("Monto")) Then
                    MONTO = MONTO + (rsGTZ("Monto") * (-1))
                  End If
                End If
                If rsGTZ("procedencia") = "2" Then 'Ingresos
                    MONTO = MONTO + (rsGTZ("Monto") * (1))
                End If
                If rsGTZ("procedencia") = "3" Then 'PCO's
                   'MsgBox "PROCEDENCIA=3"
                   
                   If rsGTZ("MontoH") <> Null Or rsGTZ("MontoH") <> 0 Then
                    MONTO = MONTO + (rsGTZ("Montoh") * (-1))
                   Else
                    MONTO = MONTO + (rsGTZ("Monto") * (1))
                   End If
                End If
                .MoveNext
             Wend
           End If
       End With
        'MsgBox Monto
        'rsCtaBancaria("Cta_Saldo_Actual") = rsCtaBancaria("Cta_saldo_inicial") + monto
        rsctabancaria("Cta_Saldo_Actual") = rsctabancaria("Cta_saldo_inicial") + MONTO
        rsctabancaria.Update
        rsctabancaria.MoveNext
    Wend
 End If
 
MsgBox "Fin de proceso"
If rsctabancaria.State = 1 Then rsctabancaria.Close
rsctabancaria.Open "SELECT CTA_CODIGO, CTA_CODIGO_TGN, CTA_DESCRIPCION_LARGA, CTA_SALDO_INICIAL, CTA_SALDO_ACTUAL  from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
If rsctabancaria.RecordCount > 0 Then
    Set DtGCuentaBancaria.DataSource = rsctabancaria
End If
 
End Sub



Private Sub Form_Load()
Dim rsctabancaria As New ADODB.Recordset
'    db.Cel_Actualizar_DAR
'    If rsCB.State = 1 Then rsCB.Close
'    rsCB.Open "SELECT CTA_CODIGO, CTA_CODIGO_TGN, CTA_DESCRIPCION_LARGA, CTA_SALDO_INICIAL, CTA_SALDO_ACTUAL FROM fc_Cuenta_Bancaria", db, adOpenKeyset, adLockOptimistic
'    If rsCB.RecordCount > 0 Then
'        Set DtGCuentaBancaria.DataSource = rsCB
'    Else
'        Set DtGCuentaBancaria.DataSource = rsNada
'        MsgBox "No existen registros", vbInformation + vbCritical, "Validaci?n de datos"
'        Exit Sub
'    End If

If rsctabancaria.State = 1 Then rsctabancaria.Close
rsctabancaria.Open "SELECT CTA_CODIGO, CTA_CODIGO_TGN, CTA_DESCRIPCION_LARGA, CTA_SALDO_INICIAL, CTA_SALDO_ACTUAL from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
If rsctabancaria.RecordCount > 0 Then
    Set DtGCuentaBancaria.DataSource = rsctabancaria
End If

'te va hablar para que le digas si esta o no.



	Call SeguridadSet(Me)
End Sub



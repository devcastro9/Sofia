VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCuentas 
   Caption         =   "Movimiento de Cuentas Bancarias"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   11820
      TabIndex        =   6
      Top             =   0
      Width           =   11880
      Begin VB.Label Label8 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Index           =   0
         Left            =   1245
         TabIndex        =   10
         Top             =   690
         Width           =   2460
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
         TabIndex        =   9
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   8
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Movimiento de Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2160
         TabIndex        =   7
         Top             =   195
         Width           =   8265
      End
   End
   Begin VB.Frame Frame1 
      Height          =   10155
      Left            =   1320
      TabIndex        =   13
      Top             =   1035
      Width           =   8715
      Begin MSAdodcLib.Adodc AdoCuenta 
         Height          =   330
         Left            =   240
         Top             =   840
         Visible         =   0   'False
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
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
         Caption         =   "AdoCuenta"
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
      Begin MSDataGridLib.DataGrid DtGGTZ 
         Height          =   5610
         Left            =   210
         TabIndex        =   19
         Top             =   1320
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   9895
         _Version        =   393216
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSComCtl2.DTPicker DTPFFin 
         Height          =   300
         Left            =   1755
         TabIndex        =   16
         Top             =   450
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36705
      End
      Begin MSComCtl2.DTPicker DTPFInicio 
         Height          =   300
         Left            =   225
         TabIndex        =   15
         Top             =   435
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36705
      End
      Begin Crystal.CrystalReport CryMovi 
         Left            =   255
         Top             =   795
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "FrmMovimiento.frx":0000
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   4005
         TabIndex        =   21
         Top             =   405
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "FrmMovimiento.frx":0018
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   4005
         TabIndex        =   22
         Top             =   750
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "FrmMovimiento.frx":0030
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   6195
         TabIndex        =   23
         Top             =   405
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. "
         Height          =   195
         Left            =   3990
         TabIndex        =   24
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin"
         Height          =   210
         Left            =   1740
         TabIndex        =   18
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Top             =   210
         Width           =   1200
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   8745
      Left            =   60
      TabIndex        =   0
      Top             =   1035
      Width           =   1230
      Begin VB.CommandButton CmdPorCuenta 
         Caption         =   "Buscar por Cuenta"
         Height          =   705
         Left            =   165
         TabIndex        =   25
         Top             =   2475
         Width           =   930
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar por todas las Ctas."
         Height          =   705
         Left            =   150
         MousePointer    =   4  'Icon
         Picture         =   "FrmMovimiento.frx":0048
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1755
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Gastos"
         Height          =   705
         Left            =   150
         MousePointer    =   4  'Icon
         Picture         =   "FrmMovimiento.frx":09EA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2475
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ejemplo Ger"
         Height          =   705
         Left            =   150
         TabIndex        =   12
         Top             =   2460
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton CmdConciliacionUDAPRE 
         Caption         =   "Conciliar pro fecha UDAPRE"
         Height          =   705
         Left            =   165
         MousePointer    =   4  'Icon
         Picture         =   "FrmMovimiento.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2460
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   150
         Picture         =   "FrmMovimiento.frx":1D2E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5985
         Width           =   945
      End
      Begin VB.CommandButton CmdImprimirMovimiento 
         Caption         =   "Imprimir "
         Height          =   825
         Left            =   135
         Picture         =   "FrmMovimiento.frx":2170
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton CmdUnionTablas 
         Caption         =   "Unión Tablas"
         Height          =   690
         Left            =   150
         MousePointer    =   4  'Icon
         Picture         =   "FrmMovimiento.frx":27DA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1065
         Width           =   945
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Limpiar"
         Height          =   765
         Left            =   150
         Picture         =   "FrmMovimiento.frx":317C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaldoSBs As Double
Dim comGastos As ADODB.Command
Dim rsGTZ As ADODB.Recordset
Dim str1 As String

Private Sub CmdImprimirTotales_Click()
    
End Sub

Private Sub CmdBuscar_Click()
                    
Dim Monto As Double
    Set rsGTZFiltro = New ADODB.Recordset
    Set rsMoviReal = New ADODB.Recordset
    db.Execute "DELETE FROM to_MovimientoReal"
        If rsMoviReal.State = 1 Then rsMoviReal.Close
        rsMoviReal.Open "select * from to_movimientoReal order by fecha_pago ", db, adOpenKeyset, adLockOptimistic
        With rsGTZ
           If .State = adStateOpen Then
             .Close
           End If
           str1 = "select * from fc_datosGTZ  where fecha_pago >= '" & Str(DTPFInicio.Value) & "'  and fecha_pago <= '" & Str(DTPFFin.Value) & "' order by fecha_pago"
           .Open str1, db, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
             While Not .EOF
                    If rsGTZ("fecha_pago") >= DTPFInicio.Value And rsGTZ("fecha_pago") <= DTPFFin.Value Then
                     Set DtGGTZ.DataSource = rsGTZ
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
                     rsMoviReal("Cta_Codigo") = rsGTZ("Cta_Codigo")
                     rsMoviReal("Nombre_Cta") = rsGTZ("Nombre_Cta")
                     rsMoviReal("Bco_Codigo") = rsGTZ("Bco_Codigo")
                     rsMoviReal("justificacion") = rsGTZ("justificacion")
                     rsMoviReal("procedencia") = rsGTZ("procedencia")
                     
                     rsMoviReal.Update
                     
'                    db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto) " & _
'                                " values('" & rsGTZ("Nro_Cmpte") & "','" & rsGTZ("Organismo") & "', '" & rsGTZ("Fecha_Pago") & "', '" & Monto & " ') "
'                    db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Procedencia) " & _
'                                " values('" & rsGTZ("Nro_Cmpte") & "', '" & rsGTZ("Organismo") & "', '" & rsGTZ("Fecha_Pago") & "'," & rsGTZ("Monto") & ", '" & rsGTZ!Cambio & "','" & rsGTZ!Beneficiario & "','" & rsGTZ!Nro_Doc & "', '" & rsGTZ!Transf_Cheq & "', '" & rsGTZ!Cta_Codigo & "', '" & rsGTZ!Bco_Codigo & "', '" & rsGTZ!Procedencia & "') "
                            
'                    db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo) " & _
'                               " values('" & rsGTZ("Nro_Cmpte") & "', '" & rsGTZ("Organismo") & "', '" & rsGTZ("Fecha_Pago") & "'," & rsGTZ("Monto") & ", '" & rsGTZ!Cambio & "','" & rsGTZ!Beneficiario & "','" & rsGTZ!Nro_Doc & "', '" & rsGTZ!Transf_Cheq & "', '" & rsGTZ!Cta_Codigo & "', '" & rsGTZ!Bco_Codigo & "') "
                     
'                     db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo) " & _
'                                " values(1,2,'2/2/2000',4,5,6,7,8,9,10) "

                    End If
                    
                 .MoveNext
             Wend
           End If
       End With
End Sub

Private Sub CmdImprimirMovimiento_Click()
    'RepMovi.Show
    
            'CryMovi.Destination = crptToWindow
            'CryMovi.ReportFileName = App.Path & "\Celia\Proyectos\Conciliacion Bancaria\Movimiento de Cuentas\Impresiones\Rpt_MovimientoReal.RPT"
            CryMovi.ReportFileName = "c:\PRAGMA5\Proyectos\Conciliacion Bancaria\Movimiento de Cuentas\Impresiones\Rep_MovimientoReal.rpt"
            CryMovi.Formulas(0) = "Fecha_Ini = '" & DTPFInicio.Value & "'"
            CryMovi.Formulas(1) = "Fecha_Fin = '" & DTPFFin.Value & "'"
            IResult = CryMovi.PrintReport
            If IResult <> 0 Then
                MsgBox CryMovi.LastErrorNumber & " : " & CryMovi.LastErrorString, vbCritical + vbOKOnly, "Error..."
            End If
    
End Sub

Private Sub CmdPorCuenta_Click()
Dim Monto As Double
    If DtCCuentaOrigen.Text = "" Then
        MsgBox "Introduzca código de la cuenta !!", vbInformation + vbCritical
        Exit Sub
    End If
    Set rsGTZFiltro = New ADODB.Recordset
    Set rsMoviReal = New ADODB.Recordset
    db.Execute "DELETE FROM to_MovimientoReal"
        If rsMoviReal.State = 1 Then rsMoviReal.Close
        rsMoviReal.Open "select * from to_movimientoReal order by fecha_pago ", db, adOpenKeyset, adLockOptimistic
        With rsGTZ
           If .State = adStateOpen Then
             .Close
           End If
           str1 = "select * from fc_datosGTZ  where cta_codigo= '" & DtCCuentaOrigen.Text & "' and fecha_pago >= '" & Str(DTPFInicio.Value) & "'  and fecha_pago <= '" & Str(DTPFFin.Value) & "' order by fecha_pago"
           .Open str1, db, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
             While Not .EOF
           '         If rsGTZ("fecha_pago") >= DTPFInicio.Value And rsGTZ("fecha_pago") <= DTPFFin.Value Then
                     Set DtGGTZ.DataSource = rsGTZ
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
                     rsMoviReal("Cta_Codigo") = rsGTZ("Cta_Codigo")
                     rsMoviReal("Nombre_Cta") = rsGTZ("Nombre_Cta")
                     rsMoviReal("Bco_Codigo") = rsGTZ("Bco_Codigo")
                     rsMoviReal("justificacion") = rsGTZ("justificacion")
                     rsMoviReal("procedencia") = rsGTZ("procedencia")
                     
                     rsMoviReal.Update
                     
'                    db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto) " & _
'                                " values('" & rsGTZ("Nro_Cmpte") & "','" & rsGTZ("Organismo") & "', '" & rsGTZ("Fecha_Pago") & "', '" & Monto & " ') "
'                    db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Procedencia) " & _
'                                " values('" & rsGTZ("Nro_Cmpte") & "', '" & rsGTZ("Organismo") & "', '" & rsGTZ("Fecha_Pago") & "'," & rsGTZ("Monto") & ", '" & rsGTZ!Cambio & "','" & rsGTZ!Beneficiario & "','" & rsGTZ!Nro_Doc & "', '" & rsGTZ!Transf_Cheq & "', '" & rsGTZ!Cta_Codigo & "', '" & rsGTZ!Bco_Codigo & "', '" & rsGTZ!Procedencia & "') "
                            
'                    db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo) " & _
'                               " values('" & rsGTZ("Nro_Cmpte") & "', '" & rsGTZ("Organismo") & "', '" & rsGTZ("Fecha_Pago") & "'," & rsGTZ("Monto") & ", '" & rsGTZ!Cambio & "','" & rsGTZ!Beneficiario & "','" & rsGTZ!Nro_Doc & "', '" & rsGTZ!Transf_Cheq & "', '" & rsGTZ!Cta_Codigo & "', '" & rsGTZ!Bco_Codigo & "') "
                     
'                     db.Execute "INSERT INTO to_MovimientoReal(Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo) " & _
'                                " values(1,2,'2/2/2000',4,5,6,7,8,9,10) "

            '        End If
                    
                 .MoveNext
             Wend
           End If
       End With

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdUnionTablas_Click()

    'Se uniran las tablas Co_MovimientoPCO, pago_detalle, fo_ingresos
    db.Execute "delete from fc_datosGTZ"
    db.movimiento_Cuenta_Bancaria
    MsgBox "fin de proceso"
End Sub

Private Sub Command1_Click()

'Ejemplo gerardo
On Error GoTo QError
    db.uno 2, 3
    Exit Sub
QError:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Private Sub Command2_Click()
Dim saldo As Parameter
MsgBox "Empieza de proceso"
'Primera forma de llamar procedimientos almacenados
' SaldoIBs = db.Parameters("GastoBs")
' db.gastos Format(Date, "dd/mm/yyyy"), Format(Date, "dd/mm/yyyy")

'Ejemplo de ...
Dim TFechaAT As New ADODB.Parameter
Dim TFechaDT As New ADODB.Parameter
Dim TSaldo As New ADODB.Parameter
Set comGastos = New ADODB.Command
With comGastos
    .CommandText = "Gastos"
    .CommandType = adCmdStoredProc
    Set TFechaAT = .CreateParameter("FechaAT", adVarChar, adParamInput, 10, DTPFInicio.Value)
    .Parameters.Append TFechaAT
    Set TFechaDT = .CreateParameter("FechaDT", adVarChar, adParamInput, 10, DTPFFin.Value)
    .Parameters.Append TFechaDT
    Set TSaldo = .CreateParameter("GastoBs", adCurrency, adParamOutput)
    .Parameters.Append TSaldo
    .ActiveConnection = db
    .Execute
    MsgBox TSaldo.Value
End With
      
'With comGastos
'            .CommandType = adCmdStoredProc
'            .CommandText = "Gastos"
'            .Parameters.Append comGastos.CreateParameter(FechaAT, adVarChar, adParamInput, 10, Date)
'            .Parameters.Append comGastos.CreateParameter(FechaDT, adVarChar, adParamInput, 10, Date)
'            .Parameters.Append comGastos.CreateParameter("GastoBs", adDouble, adParamOutput)
'            '.Parameters("FechaAT") = DTPFInicio.Value 'Format(Date, "dd/mm/yyyy")
'            '.Parameters("FechaDT") = DTPFFin.Value 'Format(Date, "dd/mm/yyyy")
'            comGastos.ActiveConnection = db
'            comGastos.Execute
'            If Not IsNull(comGastos.Parameters("GastoBs")) Then
'                SaldoSBs = comGastos.Parameters("GastoBs")
'            End If
'End With
'MsgBox "Acumulado de gatos, TESORERIA  " & SaldoSBs

End Sub

Private Sub Command3_Click()
Set rsGTZ = New ADODB.Recordset
        With rsGTZ
           If .State = adStateOpen Then
             .Close
           End If
           .Open "select * from fc_DatosGTZ order by Nro_cmpte ", db, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Set DtGGTZ.DataSource = rsGTZ
           End If
       End With
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

Private Sub Form_Load()
       'Abrir la tabla fc_DatosGTZ
     
        Set rsGTZ = New ADODB.Recordset
        With rsGTZ
           If .State = adStateOpen Then
             .Close
           End If
           .Open "select * from fc_DatosGTZ order by Nro_cmpte ", db, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Set DtGGTZ.DataSource = rsGTZ
           End If
       End With
       
        'Determinar las cuentas
        Set rsCuenta = New ADODB.Recordset
        rsCuenta.Open "select * from fc_cuenta_bancaria order by Cta_codigo_tgn", db, adOpenKeyset, adLockOptimistic
        Set AdoCuenta.Recordset = rsCuenta
   
End Sub


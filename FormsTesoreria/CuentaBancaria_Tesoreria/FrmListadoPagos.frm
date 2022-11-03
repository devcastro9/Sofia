VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmListadoPagos 
   Caption         =   "Gastos Realizados en Tesorería"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   Icon            =   "FrmListadoPagos.frx":0000
   LinkTopic       =   "Listado de Pagos"
   ScaleHeight     =   4995
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   435
      Left            =   2535
      Top             =   5025
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   767
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
      Caption         =   ""
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
   Begin VB.Frame Frame1 
      Height          =   3885
      Left            =   60
      TabIndex        =   6
      Top             =   1050
      Width           =   7080
      Begin VB.Frame FraFecha 
         Caption         =   "Impresión Por "
         Height          =   855
         Left            =   2340
         TabIndex        =   18
         Top             =   150
         Width           =   2190
         Begin VB.OptionButton OptFechaImpresion 
            Caption         =   "Fecha de Impresión"
            Height          =   315
            Left            =   165
            TabIndex        =   21
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton OptFechaPago 
            Caption         =   "Fecha de Pago"
            Height          =   300
            Left            =   165
            TabIndex        =   20
            Top             =   195
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.CheckBox ChkDanida 
            Caption         =   "ORGANISMO 999"
            Height          =   405
            Left            =   2220
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   1830
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   990
         Left            =   255
         TabIndex        =   25
         Top             =   1410
         Width           =   1830
      End
      Begin VB.Frame FraOpcionesCuenta 
         Height          =   855
         Left            =   4560
         TabIndex        =   22
         Top             =   150
         Width           =   2385
         Begin VB.OptionButton OptUnaCuenta 
            Caption         =   "Por Cuenta"
            Height          =   330
            Left            =   105
            TabIndex        =   24
            Top             =   135
            Value           =   -1  'True
            Width           =   1830
         End
         Begin VB.OptionButton OptTodasCuentas 
            Caption         =   "X Todas las Cuentas"
            Height          =   315
            Left            =   105
            TabIndex        =   23
            Top             =   480
            Width           =   2010
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha"
         Height          =   1260
         Left            =   2325
         TabIndex        =   13
         Top             =   975
         Width           =   4620
         Begin MSComCtl2.DTPicker DTPFechaInicio 
            Height          =   375
            Left            =   210
            TabIndex        =   14
            Top             =   630
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   248643585
            CurrentDate     =   36413
         End
         Begin MSComCtl2.DTPicker DTPFechaFin 
            Height          =   375
            Left            =   2355
            TabIndex        =   15
            Top             =   645
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   248643585
            CurrentDate     =   36413
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio"
            Height          =   240
            Left            =   225
            TabIndex        =   17
            Top             =   420
            Width           =   1590
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin"
            Height          =   240
            Left            =   2370
            TabIndex        =   16
            Top             =   450
            Width           =   1590
         End
      End
      Begin VB.Frame FraCuenta 
         Caption         =   "Cuenta"
         Height          =   1320
         Left            =   2325
         TabIndex        =   9
         Top             =   2265
         Width           =   4635
         Begin MSDataListLib.DataCombo DtCCuentaOrigen 
            Bindings        =   "FrmListadoPagos.frx":0ECA
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   105
            TabIndex        =   10
            Top             =   225
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCDescripcion 
            Bindings        =   "FrmListadoPagos.frx":0EE2
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   105
            TabIndex        =   11
            Top             =   930
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Cta_descripcion_larga"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCTgn 
            Bindings        =   "FrmListadoPagos.frx":0EFA
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   105
            TabIndex        =   12
            Top             =   585
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Cta_codigo_tgn"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
      End
      Begin VB.CommandButton CmdEjecutar 
         Caption         =   "Ejecutar"
         Height          =   1080
         Left            =   255
         TabIndex        =   8
         Top             =   330
         Width           =   1830
      End
      Begin VB.CommandButton CmdTodasCtas 
         Caption         =   "Todas Cuentas"
         Height          =   975
         Left            =   255
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   1830
      End
      Begin Crystal.CrystalReport CryMov 
         Left            =   2265
         Top             =   3585
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   7140
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Pagos"
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
         Left            =   600
         TabIndex        =   5
         Top             =   105
         Width           =   8265
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Index           =   0
         Left            =   1245
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   675
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "FrmListadoPagos.frx":0F12
         Top             =   0
         Width           =   11640
      End
   End
End
Attribute VB_Name = "FrmListadoPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim comCuentasAcumuladas As ADODB.Command
Dim comPagos As ADODB.Command
Dim iResult  As Variant
Dim rscuenta As New ADODB.Recordset

Private Sub CmdEjecutar_Click()
   'Validación de fechas
    If DTPFechaInicio.Value > DTPFechaFin.Value Or DTPFechaFin.Value < DTPFechaInicio.Value Then
        MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
        Exit Sub
     End If
    If OptUnaCuenta.Value = True Then
        Reporte_UnaCta
    End If
    If OptTodasCuentas.Value = True Then
        Reporte_TodasCtas
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTodasCtas_Click()
    Dim fecha1 As New ADODB.Parameter
    Dim fecha2 As New ADODB.Parameter
    Dim op1 As New ADODB.Parameter
    op1 = "P"
    Set comPagos = New ADODB.Command
    'Set comPagos_Todos = New ADODB.Command
    With comPagos
        .CommandText = "Cel_Listado_Pagos_Todos"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, DTPFechaInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, DTPFechaFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 1, op1)
        .Parameters.Append op1
        
        .ActiveConnection = db
        .Execute
    End With
    
        CryMov.ReportFileName = App.Path & "\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\Rpt_CtaBancaria.rpt"
        iResult = CryMov.PrintReport
        If iResult <> 0 Then
           MsgBox CryMov.LastErrorNumber & " : " & CryMov.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If

End Sub

Public Sub Reporte_UnaCta()
Dim De As String
Dim A As String
Dim TC As String
Dim Cadena As String
If FrmListadoPagos.OptFechaPago.Value = True Then
     Cadena = "REPORTE POR FECHA DE PAGO"
Else
     Cadena = "REPORTE POR FECHA DE IMPRESION"
End If

If FrmListadoPagos.OptFechaPago.Value = True Then CryMov.Formulas(4) = "Fecha_Pago_Impresion='" & Cadena & "'"
If FrmListadoPagos.OptFechaImpresion.Value = True Then CryMov.Formulas(4) = "Fecha_Pago_Impresion='" & Cadena & "'"

If FrmListadoPagos.OptUnaCuenta.Value = True Then
    CryMov.Formulas(1) = "FCodigo_Cuenta='" & FrmListadoPagos.DtCCuentaOrigen.Text & "'"
    CryMov.Formulas(2) = "FDescripcion_Cuenta='" & FrmListadoPagos.DtCDescripcion.Text & "'"
End If
If FrmListadoPagos.OptTodasCuentas.Value = True Then
    TC = "Todas las cuentas"
    CryMov.Formulas(9) = "FTodasCuentas='" & TC & "'"
End If
    De = "De"
    A = "A"
    CryMov.Formulas(6) = "FFechaInicio='" & FrmListadoPagos.DTPFechaInicio.Value & "'"
    CryMov.Formulas(5) = "FFechaFin='" & FrmListadoPagos.DTPFechaFin.Value & "'"
    CryMov.Formulas(2) = "FDe='" & De & "' "
    CryMov.Formulas(0) = "Fa='" & A & "' "



    Dim fecha1 As New ADODB.Parameter
    Dim fecha2 As New ADODB.Parameter
    Dim Cta As New ADODB.Parameter
    Dim op1 As New ADODB.Parameter

    If OptFechaPago.Value = True Then
        op1 = "P"
    End If
    If OptFechaImpresion.Value = True Then
        op1 = "I"
    End If

    Set comPagos = New ADODB.Command
    With comPagos
        .CommandText = "Cel_Listado_Pagos_Cuenta"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, DTPFechaInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, DTPFechaFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 1, op1)
        .Parameters.Append op1
        Set Cta = .CreateParameter("Cuenta", adVarChar, adParamInput, 40, DtCCuentaOrigen.Text)
        .Parameters.Append Cta
        .ActiveConnection = db
        .Execute
    End With
    
        CryMov.ReportFileName = App.Path & "\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\Rpt_CtaBancaria.rpt"
        iResult = CryMov.PrintReport
        If iResult <> 0 Then
           MsgBox CryMov.LastErrorNumber & " : " & CryMov.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
    
End Sub

Public Sub Reporte_TodasCtas()
    
Dim De As String
Dim A As String
Dim TC As String
Dim Cadena As String
If FrmListadoPagos.OptFechaPago.Value = True Then
     Cadena = "REPORTE POR FECHA DE PAGO"
Else
     Cadena = "REPORTE POR FECHA DE IMPRESION"
End If

If FrmListadoPagos.OptFechaPago.Value = True Then CryMov.Formulas(4) = "Fecha_Pago_Impresion='" & Cadena & "'"
If FrmListadoPagos.OptFechaImpresion.Value = True Then CryMov.Formulas(4) = "Fecha_Pago_Impresion='" & Cadena & "'"

If FrmListadoPagos.OptUnaCuenta.Value = True Then
    CryMov.Formulas(1) = "FCodigo_Cuenta='" & FrmListadoPagos.DtCCuentaOrigen.Text & "'"
    CryMov.Formulas(2) = "FDescripcion_Cuenta='" & FrmListadoPagos.DtCDescripcion.Text & "'"
End If
If FrmListadoPagos.OptTodasCuentas.Value = True Then
    TC = "Todas las cuentas"
    CryMov.Formulas(9) = "FTodasCuentas='" & TC & "'"
End If

    De = "De"
    A = "A"
    CryMov.Formulas(6) = "FFechaInicio='" & FrmListadoPagos.DTPFechaInicio.Value & "'"
    CryMov.Formulas(5) = "FFechaFin='" & FrmListadoPagos.DTPFechaFin.Value & "'"
    CryMov.Formulas(2) = "FDe='" & De & "' "
    CryMov.Formulas(0) = "Fa='" & A & "' "

    
    Dim fecha1 As New ADODB.Parameter
    Dim fecha2 As New ADODB.Parameter
    Dim op1 As New ADODB.Parameter
    If OptFechaPago.Value = True Then
        op1 = "P"
    End If
    If OptFechaImpresion.Value = True Then
        op1 = "I"
    End If
    Set comPagos = New ADODB.Command
    With comPagos
        .CommandText = "Cel_Listado_Pagos_Todos"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, DTPFechaInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, DTPFechaFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 1, op1)
        .Parameters.Append op1
        
        .ActiveConnection = db
        .Execute
    End With
    
        CryMov.ReportFileName = App.Path & "\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\Rpt_CtaBancaria.rpt"
        iResult = CryMov.PrintReport
        If iResult <> 0 Then
           MsgBox CryMov.LastErrorNumber & " : " & CryMov.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If

End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCCuentaOrigen.BoundText
    DtCTgn.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
   DtCTgn.BoundText = DtCDescripcion.BoundText
   DtCCuentaOrigen.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub DtCTgn_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCTgn.BoundText
    DtCCuentaOrigen.BoundText = DtCTgn.BoundText
End Sub

Private Sub Form_Load()
    Set rscuenta = New ADODB.Recordset
    rscuenta.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    Set AdoCuenta.Recordset = rscuenta
    DTPFechaInicio.Value = Date
    DTPFechaFin.Value = Date
'    DTPFechaIn.Value = Date
'    DTPFechaFin.Value = Date
	Call SeguridadSet(Me)
End Sub

Private Sub OptTodasCuentas_Click()
    FraCuenta.Visible = False
End Sub

Private Sub OptUnaCuenta_Click()
    FraCuenta.Visible = True
End Sub

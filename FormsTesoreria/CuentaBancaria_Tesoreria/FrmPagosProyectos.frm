VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmPagosProyectos 
   Caption         =   "Listado de Pagos por  Proyecto"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmPagosProyectos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FraOpciones 
      Height          =   7845
      Left            =   45
      TabIndex        =   7
      Top             =   1035
      Width           =   1230
      Begin VB.CommandButton CmdImprimirMovimiento 
         Caption         =   "Imprimir "
         Height          =   750
         Left            =   150
         Picture         =   "FrmPagosProyectos.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   150
         Picture         =   "FrmPagosProyectos.frx":1534
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5985
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7860
      Left            =   1320
      TabIndex        =   6
      Top             =   1035
      Width           =   10695
      Begin MSDataGridLib.DataGrid DtGDatos 
         Height          =   4005
         Left            =   150
         TabIndex        =   33
         Top             =   255
         Width           =   10110
         _ExtentX        =   17833
         _ExtentY        =   7064
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
      Begin VB.Frame FraBusca 
         Height          =   3480
         Left            =   135
         TabIndex        =   10
         Top             =   4290
         Width           =   10080
         Begin VB.CommandButton CmdBeneficiario 
            Caption         =   "Por Beneficiario"
            Height          =   405
            Left            =   7725
            TabIndex        =   34
            Top             =   1575
            Width           =   2055
         End
         Begin VB.TextBox TxtMonto_Bolivianos 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2340
            TabIndex        =   16
            Top             =   1725
            Width           =   1335
         End
         Begin VB.TextBox TxtBeneficiario 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1455
            TabIndex        =   15
            Top             =   570
            Width           =   3930
         End
         Begin VB.TextBox TxtNroCheque 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   1725
            Width           =   2100
         End
         Begin VB.TextBox TxtCodigoSolicitud 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   75
            TabIndex        =   13
            Top             =   570
            Width           =   1350
         End
         Begin VB.CommandButton CmdBuscaSolicitud 
            Caption         =   "Buscar por Nro. Sol."
            Height          =   405
            Left            =   7710
            TabIndex        =   12
            Top             =   600
            Width           =   2145
         End
         Begin VB.CommandButton CmdBuscarNroCheque 
            Caption         =   "Buscar  Por Nro. Cheque"
            Height          =   420
            Left            =   7710
            TabIndex        =   11
            Top             =   180
            Width           =   2160
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   360
            Left            =   2310
            Top             =   1695
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   635
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
         Begin MSDataListLib.DataCombo DtCCuentaOrigen1 
            Bindings        =   "FrmPagosProyectos.frx":1976
            DataField       =   "cta_codigo"
            DataSource      =   "AdoCuenta"
            Height          =   315
            Left            =   135
            TabIndex        =   17
            Top             =   2385
            Visible         =   0   'False
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCCuentaOrigenDes1 
            Bindings        =   "FrmPagosProyectos.frx":198E
            DataField       =   "cta_codigo"
            DataSource      =   "AdoCuenta"
            Height          =   315
            Left            =   3885
            TabIndex        =   18
            Top             =   2400
            Visible         =   0   'False
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Cta_descripcion_larga"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcCtaTGN1 
            Bindings        =   "FrmPagosProyectos.frx":19A6
            DataField       =   "cta_codigo"
            DataSource      =   "AdoCuenta"
            Height          =   315
            Left            =   2265
            TabIndex        =   19
            Top             =   2400
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Cta_codigo_tgn"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DTPFFin 
            Height          =   300
            Left            =   1695
            TabIndex        =   25
            Top             =   3075
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   248643585
            CurrentDate     =   36705
         End
         Begin MSComCtl2.DTPicker DTPFInicio 
            Height          =   300
            Left            =   180
            TabIndex        =   26
            Top             =   3075
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   529
            _Version        =   393216
            Format          =   248643585
            CurrentDate     =   36705
         End
         Begin MSDataListLib.DataCombo DtCCuentaOrigen 
            Bindings        =   "FrmPagosProyectos.frx":19BE
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   1140
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
         Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
            Bindings        =   "FrmPagosProyectos.frx":19D6
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   3930
            TabIndex        =   30
            Top             =   1140
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
         Begin MSDataListLib.DataCombo DtcCtaTGN 
            Bindings        =   "FrmPagosProyectos.frx":19EE
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   2310
            TabIndex        =   31
            Top             =   1140
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
         Begin VB.Label LblCuenta 
            AutoSize        =   -1  'True
            Caption         =   "No. Cta. "
            Height          =   195
            Left            =   105
            TabIndex        =   32
            Top             =   930
            Width           =   630
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Fin"
            Height          =   210
            Left            =   1680
            TabIndex        =   28
            Top             =   2865
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   165
            TabIndex        =   27
            Top             =   2835
            Width           =   1200
         End
         Begin VB.Label Label13 
            Caption         =   "Monto Bolivianos"
            Height          =   165
            Left            =   2310
            TabIndex        =   24
            Top             =   1515
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Beneficiario"
            Height          =   195
            Left            =   1470
            TabIndex        =   23
            Top             =   360
            Width           =   2250
         End
         Begin VB.Label Label5 
            Caption         =   "Nro. Solicitud"
            Height          =   195
            Left            =   90
            TabIndex        =   22
            Top             =   330
            Width           =   1230
         End
         Begin VB.Label Label4 
            Caption         =   "Nro. de Cheque"
            Height          =   180
            Left            =   135
            TabIndex        =   21
            Top             =   1500
            Width           =   1290
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "No. Cta. "
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   2175
            Width           =   630
         End
      End
      Begin Crystal.CrystalReport CryMov 
         Left            =   180
         Top             =   3870
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
      Begin MSAdodcLib.Adodc AdoCuenta 
         Height          =   330
         Left            =   165
         Top             =   4005
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Pagos - Proyectos"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   195
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
         BackColor       =   &H8000000B&
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
         BackColor       =   &H8000000B&
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
         Picture         =   "FrmPagosProyectos.frx":1A06
         Top             =   0
         Width           =   11640
      End
   End
End
Attribute VB_Name = "FrmPagosProyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaldoSBs As Double
Dim comGastos As ADODB.Command
Dim rsGTZ As ADODB.Recordset
Dim str1 As String
Dim rsNada As ADODB.Recordset

Private Sub CmdBuscar_Click()
                   
Dim De As String
Dim A As String
Dim TC As String
Dim Cadena As String
Dim fecha1 As New ADODB.Parameter
Dim fecha2 As New ADODB.Parameter
Dim op1 As New ADODB.Parameter

'If FrmListadoPagos.OptUnaCuenta.Value = True Then
'    CryMov.Formulas(1) = "FCodigo_Cuenta='" & FrmListadoPagos.DtCCuentaOrigen.Text & "'"
'    CryMov.Formulas(2) = "FDescripcion_Cuenta='" & FrmListadoPagos.DtCDescripcion.Text & "'"
'End If
'If FrmListadoPagos.OptTodasCuentas.Value = True Then
'    TC = "Todas las cuentas"
'    'CryMov.Formulas(9) = "FTodasCuentas='" & TC & "'"
'End If
'
    De = "De"
    A = "A"
    'CryMov.Formulas(6) = "FFechaInicio='" & DTPFInicio.Value & "'"
    'CryMov.Formulas(5) = "FFechaFin='" & DTPFFin.Value & "'"
    'CryMov.Formulas(2) = "FDe='" & De & "' "
    'CryMov.Formulas(0) = "Fa='" & A & "' "

    
    
    'If OptFechaPago.Value = True Then
        op1 = "T"
    'End If
    'If OptFechaImpresion.Value = True Then
    '    op1 = "I"
    'End If
    Set comPagos = New ADODB.Command
    With comPagos
        .CommandText = "Cel_Tesoreria_Proyectos"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, DTPFInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, DTPFFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 1, op1)
        .Parameters.Append op1
        
        .ActiveConnection = db
        .Execute
    End With
    
        CryMov.ReportFileName = App.Path & "\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\RptPagosProyectos.rpt"
        iResult = CryMov.PrintReport
        If iResult <> 0 Then
           MsgBox CryMov.LastErrorNumber & " : " & CryMov.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
End Sub

Private Sub CmdBeneficiario_Click()
Dim rsTP As New ADODB.Recordset
 If txtbeneficiario.Text = "" Then
        MsgBox "No existe beneficiario", vbCritical + vbDefaultButton1
        Exit Sub
 End If
db.Execute "INSERT INTO to_Tesoreria_Proyectos(Nro_Cmpte, Organismo, Fecha_Pago, Monto, " & _
                             "Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Estado_Conciliacion, Procedencia, justificacion, nombre_cta, tipo_comp,cta_codigo_destino, proyecto, nro_solicitud, denominacion_beneficiario) " & _
                             "SELECT pago_detalle.codigo_pago, pago_detalle.org_codigo, " & _
                             "pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, " & _
                             "pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, " & _
                             "pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf, " & _
                             "pago_detalle.cta_codigo, fc_cuenta_bancaria.Bco_codigo, " & _
                             "pago_detalle.Estado_Conciliacion, 1, " & _
                             "pagos.justificacion, " & _
                             "fc_cuenta_bancaria .Cta_descripcion_larga, " & _
                             "tipo_comp, " & _
                             "pago_detalle.cta_codigo_destino, " & _
                             "fc_estructura_programatica.pro_descripcion_larga, pagos.codigo_solicitud, fc_beneficiario.denominacion_beneficiario " & _
                             "FROM pago_detalle INNER JOIN " & _
                             "fc_cuenta_bancaria ON " & _
                             "pago_detalle.Cta_codigo = fc_cuenta_bancaria.Cta_codigo Inner Join  fc_estructura_programatica ON  pago_detalle.Pro_proyecto = fc_estructura_programatica.Pro_proyecto Inner Join pagos ON         pago_detalle.Ges_gestion = pagos.ges_gestion AND         pago_detalle.org_codigo = pagos.org_codigo AND pago_detalle.codigo_pago = pagos.codigo_pago " & _
                             "INNER JOIN FC_BENEFICIARIO ON pago_detalle.codigo_beneficiario=fc_beneficiario.codigo_beneficiario " & _
                             "WHERE fc_beneficiario.denominacion_beneficiario like  '%" + txtbeneficiario.Text + "%' "
                             
               If rsTP.State = 1 Then rsTP.Close
               rsTP.Open "SELECT * FROM to_Tesoreria_Proyectos", db, adOpenKeyset, adLockBatchOptimistic
               If rsTP.RecordCount > 0 Then
                Set DtGDatos.DataSource = rsTP
               Else
                Set DtGDatos.DataSource = rsNada
               End If

End Sub

Private Sub CmdBuscarNroCheque_Click()
Dim codigo_solicitud As String

    If DtCCuentaOrigen.Text = "" Then
        MsgBox "No existe número de cuenta ", vbCritical + vbDefaultButton1
        Exit Sub
    End If

    If TxtNroCheque.Text = "" Then
        MsgBox "No existe número de cheque", vbCritical + vbDefaultButton1
        Exit Sub
    End If
    db.Execute "delete from to_Tesoreria_Proyectos"
    Set rsPAgoDetalle = New ADODB.Recordset
    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    rsPAgoDetalle.Open "SELECT * FROM pago_detalle WHERE cta_codigo='" & DtCCuentaOrigen.Text & "' and numero_cheque_trf= '" & TxtNroCheque.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsPAgoDetalle.RecordCount > 0 Then
                Set rspago = New ADODB.Recordset
                If rspago.State = 1 Then rspago.Close
                rspago.Open "SELECT * FROM pagos WHERE ges_gestion='" & rsPAgoDetalle("ges_gestion") & "' and org_codigo= '" & rsPAgoDetalle("org_codigo") & "' and codigo_Pago= '" & rsPAgoDetalle("codigo_pago") & "' ", db, adOpenKeyset, adLockOptimistic
                If rspago.RecordCount > 0 And Not IsNull(rspago("codigo_solicitud")) Then
                    codigo_solicitud = rspago("codigo_solicitud")
                End If
                
                
                db.Execute "INSERT INTO to_Tesoreria_Proyectos(Nro_Cmpte, Organismo, Fecha_Pago, Monto, " & _
                             "Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Estado_Conciliacion, Procedencia, justificacion, nombre_cta, tipo_comp,cta_codigo_destino, proyecto, nro_solicitud) " & _
                             "SELECT pago_detalle.codigo_pago, pago_detalle.org_codigo, " & _
                             "pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, " & _
                             "pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, " & _
                             "pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf, " & _
                             "pago_detalle.cta_codigo, fc_cuenta_bancaria.Bco_codigo, " & _
                             "pago_detalle.Estado_Conciliacion, 1, " & _
                             "pagos.justificacion, " & _
                             "fc_cuenta_bancaria .Cta_descripcion_larga, " & _
                             "tipo_comp, " & _
                             "pago_detalle.cta_codigo_destino, " & _
                             "fc_estructura_programatica.pro_descripcion_larga, pagos.codigo_solicitud" & _
                             "FROM pago_detalle INNER JOIN " & _
                             "fc_cuenta_bancaria ON " & _
                             "pago_detalle.Cta_codigo = fc_cuenta_bancaria.Cta_codigo Inner Join  fc_estructura_programatica ON         pago_detalle.Pro_proyecto = fc_estructura_programatica.Pro_proyecto Inner Join pagos ON         pago_detalle.Ges_gestion = pagos.ges_gestion AND         pago_detalle.org_codigo = pagos.org_codigo AND pago_detalle.codigo_pago = pagos.codigo_pago" & _
                             "WHERE pagos.codigo_solicitud='" & codigo_solicitud & "'"

                             

'                'Buscando todos con el mismo Nro. de solicitud
'                Set DtGTributosFiscales.DataSource = rspago
'                'While Not rsPago.EOF
'                    Set rsVista = New ADODB.Recordset
'
'
'
'                    'rsvista.Open " SELECT  pago_detalle.*, pagos.*, fc_beneficiario.denominacion_beneficiario " &
'                    rsVista.Open " SELECT fc_beneficiario.denominacion_beneficiario as BENEFICIARIO, pago_detalle.codigo_pago as CMPTE,Pago_detalle.org_codigo as ORGANISMO,pago_detalle.numero_cheque_trf AS CHEQTRANSF, pago_detalle.codigo_beneficiario as CODIGOBENEF, pago_detalle.monto_bolivianos as MONTO, pago_detalle.par_codigo as PARTIDA,pago_detalle.*, pagos.*  " & _
'                                 " FROM (pago_detalle INNER JOIN pagos ON (pago_detalle.Ges_gestion = pagos.ges_gestion) AND (pago_detalle.org_codigo = pagos.org_codigo) AND (pago_detalle.codigo_pago = pagos.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario WHERE pagos.codigo_solicitud='" & codigo_solicitud & "' ", db, adOpenKeyset, adLockOptimistic
'                    If rsVista.RecordCount > 0 Then
'                            Set DtGTributosFiscales.DataSource = rsVista
'                            rsVista.MoveFirst
'                            While Not rsVista.EOF
'                                    If rsVista("CODIGOBENEF") <> "035_SNII" Then
'                                        TxtBeneficiario.Text = rsVista("beneficiario")
'                                    End If
'                                    If rsVista("MONTO") > 0 Then
'                                        Monto = Monto + rsVista("MONTO")
'                                    End If
''                                   lstBeneficiario.AddItem rsVista("beneficiario")
'                                   If Not IsNull(rsVista("CheqTransf")) Then
'                                   Dim SQLVar As String
'                                   SQLVar = "insert into to_PagosTributos (T_beneficiario, codigo_pago, org_codigo, Nro_doc, monto_bolivianos) " & _
'                                            "values  ('" & rsVista!beneficiario & "','" & rsVista("cmpte") & "', '" & rsVista("organismo") & "','" & rsVista("CheqTransf") & "'," & CCur(rsVista("monto")) & " ) "
'                                   db.Execute SQLVar
'                                   End If
'                                   rsVista.MoveNext
'                            Wend
'                    Else
'                            MsgBox "No existen registros", vbInformation + vbCritical
'                            Set DtGTributosFiscales.DataSource = rsNada
'                    End If
'                    TxtMonto_Bolivianos = Monto
                
'                    Set rsPago = New ADODB.Recordset
'                    If rsPago.State = 1 Then rsPago.Close
'                    rsPago.Open "SELECT * FROM pagos WHERE codigo_solicitud= '" & codigo_solicitud & "'", db, adOpenKeyset, adLockOptimistic
'                    If rsPago.RecordCount > 0 Then
'                        Set DtGTributosFiscales.DataSource = rsPago
'                    Else
'                        MsgBox "No existen registros", vbInformation + vbCritical
'                         Set DtGTributosFiscales.DataSource = rsNada
'                    End If
        
    End If
End Sub

Private Sub CmdBuscaSolicitud_Click()
Dim codigo_solicitud As String
Dim rsTP As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset


db.Execute "delete from to_Tesoreria_Proyectos"
If TxtCodigoSolicitud.Text = "" Then
    MsgBox "No existe dato de solicitud", vbInformation + vbCritical
    Exit Sub
End If


                db.Execute "INSERT INTO to_Tesoreria_Proyectos(Nro_Cmpte, Organismo, Fecha_Pago, Monto, " & _
                             "Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Estado_Conciliacion, Procedencia, justificacion, nombre_cta, tipo_comp,cta_codigo_destino, proyecto, nro_solicitud) " & _
                             "SELECT pago_detalle.codigo_pago, pago_detalle.org_codigo, " & _
                             "pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, " & _
                             "pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, " & _
                             "pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf, " & _
                             "pago_detalle.cta_codigo, fc_cuenta_bancaria.Bco_codigo, " & _
                             "pago_detalle.Estado_Conciliacion, 1, " & _
                             "pagos.justificacion, " & _
                             "fc_cuenta_bancaria .Cta_descripcion_larga, " & _
                             "tipo_comp, " & _
                             "pago_detalle.cta_codigo_destino, " & _
                             "fc_estructura_programatica.pro_descripcion_larga, pagos.codigo_solicitud " & _
                             "FROM pago_detalle INNER JOIN " & _
                             "fc_cuenta_bancaria ON " & _
                             "pago_detalle.Cta_codigo = fc_cuenta_bancaria.Cta_codigo Inner Join  fc_estructura_programatica ON         pago_detalle.Pro_proyecto = fc_estructura_programatica.Pro_proyecto Inner Join pagos ON         pago_detalle.Ges_gestion = pagos.ges_gestion AND         pago_detalle.org_codigo = pagos.org_codigo AND pago_detalle.codigo_pago = pagos.codigo_pago" & _
                             " WHERE pagos.codigo_solicitud='" & TxtCodigoSolicitud.Text & "' "
                             
               If rsTP.State = 1 Then rsTP.Close
               rsTP.Open "SELECT * FROM to_Tesoreria_Proyectos", db, adOpenKeyset, adLockBatchOptimistic
               If rsTP.RecordCount > 0 Then
                Set DtGDatos.DataSource = rsTP
               Else
                Set DtGDatos.DataSource = rsNada
               End If


End Sub

Private Sub CmdImprimirMovimiento_Click()
        CryMov.ReportFileName = App.Path & "\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\RptPagosProyectos.rpt"
        iResult = CryMov.PrintReport
        If iResult <> 0 Then
           MsgBox CryMov.LastErrorNumber & " : " & CryMov.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
End Sub

Private Sub CmdPorCuenta_Click()
Dim MONTO As Double

'Validación de datos
If DTPFInicio.Value > DTPFFin.Value Or DTPFFin.Value < DTPFInicio.Value Then
     MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
     Exit Sub
End If
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
               Set DtGGTZ.DataSource = rsGTZ
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
           str1 = "select * from fc_datosGTZ  where cta_codigo_destino= '" & DtCCuentaOrigen.Text & "' and fecha_pago >= '" & Str(DTPFInicio.Value) & "'  and fecha_pago <= '" & Str(DTPFFin.Value) & "' and tipo_comp='TRP' order by fecha_pago"
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
       
  Set rsMoviReal = New ADODB.Recordset
  If rsMoviReal.State = 1 Then rsMoviReal.Close
        rsMoviReal.Open "select * from to_movimientoReal order by fecha_pago ", db, adOpenKeyset, adLockOptimistic
  If rsMoviReal.RecordCount > 0 Then
    Set DtGGTZ.DataSource = rsMoviReal
  Else
    Set DtGGTZ.DataSource = rsNada
  End If
        
       
       
'Activando la cta.
DtCCuentaOrigen.Visible = True
DtcCtaTGN.Visible = True
DtCCuentaOrigenDes.Visible = True
lblcuenta.Visible = True
       
       
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdUnionTablas_Click()
    'Se uniran las tablas Co_MovimientoPCO, pago_detalle, fo_ingresos
    db.Execute "delete from fc_datosGTZ"
    db.movimiento_Cuenta_Bancaria
    MsgBox "fin de proceso"
End Sub

Private Sub CmdSalirBuscar_Click()

End Sub

Private Sub Command1_Click()

'Ejemplo gerardo
On Error GoTo QError
    db.uno 2, 3
    Exit Sub
QError:
    MsgBox err.Number & " : " & err.Description
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
    .CommandText = "Cel_Tesoreria_Proyectos"
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

        db.Execute "DELETE FROM to_tesoreria_proyectos"
       
        'Determinar las cuentas
        Set rscuenta = New ADODB.Recordset
        rscuenta.Open "select * from fc_cuenta_bancaria order by Cta_codigo_tgn", db, adOpenKeyset, adLockOptimistic
        Set AdoCuenta.Recordset = rscuenta
   
End Sub




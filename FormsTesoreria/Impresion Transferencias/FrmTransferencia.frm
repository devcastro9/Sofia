VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmTransferencia 
   Caption         =   "Imprimir Transferencia"
   ClientHeight    =   8595
   ClientLeft      =   1215
   ClientTop       =   1365
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "IMPRESION DE CARTAS DE TRANSFERENCIAS"
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
         Left            =   2790
         TabIndex        =   5
         Top             =   225
         Width           =   7095
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   2
         Top             =   690
         Width           =   2460
      End
      Begin VB.Label Label2 
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
         Left            =   60
         TabIndex        =   1
         Top             =   675
         Width           =   1110
      End
   End
   Begin VB.Frame FraBusca 
      Height          =   2085
      Left            =   2640
      TabIndex        =   52
      Top             =   3810
      Visible         =   0   'False
      Width           =   2040
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   225
         TabIndex        =   67
         Top             =   1485
         Width           =   1515
      End
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   56
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2047
         TabIndex        =   55
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   225
         TabIndex        =   54
         Top             =   645
         Width           =   1515
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   225
         TabIndex        =   53
         Top             =   1095
         Width           =   1515
      End
      Begin VB.Label Label24 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   59
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label23 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   2310
         TabIndex        =   58
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label22 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   450
         TabIndex        =   57
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.ListBox LstGes 
      Height          =   4350
      Left            =   15120
      TabIndex        =   46
      Top             =   4590
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ListBox LstOrg 
      Height          =   4350
      Left            =   14745
      TabIndex        =   45
      Top             =   4605
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox LstBancoDestino 
      Height          =   4350
      Left            =   13980
      TabIndex        =   44
      Top             =   4605
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.ListBox LstObs 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   13545
      TabIndex        =   41
      Top             =   4605
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.ListBox LstDolares 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   12780
      TabIndex        =   38
      Top             =   4635
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.OptionButton OptDol 
      Caption         =   "Dólares"
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2490
      TabIndex        =   37
      Top             =   1080
      Width           =   1080
   End
   Begin VB.OptionButton OptBol 
      Caption         =   "Bolivianos"
      Height          =   270
      Left            =   1365
      TabIndex        =   36
      Top             =   1095
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.ListBox LstBanco 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   5490
      TabIndex        =   34
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   4635
      Width           =   1125
   End
   Begin VB.Frame FraOpciones 
      Height          =   9210
      Left            =   -15
      TabIndex        =   16
      Top             =   540
      Width           =   1245
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   585
         TabIndex        =   63
         Top             =   885
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   688
         Appearance      =   1
         _Version        =   327682
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cola Impr."
         Height          =   360
         Left            =   120
         TabIndex        =   66
         Top             =   5040
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cola Imp."
         Height          =   360
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton CmdReimpresion 
         Caption         =   "Reimprimir"
         Height          =   735
         Left            =   105
         Picture         =   "FrmTransferencia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   4305
         Width           =   930
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Busqueda"
         Height          =   735
         Left            =   105
         Picture         =   "FrmTransferencia.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3570
         Width           =   930
      End
      Begin VB.CommandButton CmdFiltro 
         Caption         =   "Filtro por Organismo"
         Height          =   735
         Left            =   105
         TabIndex        =   21
         Top             =   2820
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   105
         Picture         =   "FrmTransferencia.frx":076C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   735
         Left            =   105
         Picture         =   "FrmTransferencia.frx":0DD6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5400
         Width           =   930
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   735
         Left            =   105
         TabIndex        =   18
         Top             =   1335
         Width           =   930
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar Grid"
         Height          =   735
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2085
         Width           =   930
      End
   End
   Begin VB.ListBox LstLiteral 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   9120
      TabIndex        =   15
      Top             =   4635
      Width           =   330
   End
   Begin VB.ListBox LstJustificacion 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   9420
      TabIndex        =   14
      Top             =   4635
      Width           =   1140
   End
   Begin VB.ListBox LstDesCuenta 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   6600
      TabIndex        =   13
      Top             =   4635
      Width           =   930
   End
   Begin VB.ListBox LstFecha 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   3285
      TabIndex        =   12
      Top             =   4650
      Width           =   1005
   End
   Begin VB.ListBox LstDepto 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   11550
      TabIndex        =   11
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   4635
      Width           =   1185
   End
   Begin VB.ListBox LstCuentaDes 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   10545
      TabIndex        =   10
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   4635
      Width           =   990
   End
   Begin VB.ListBox LstMontoBol 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   7500
      TabIndex        =   9
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   4635
      Width           =   1635
   End
   Begin VB.ListBox LstTransf 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   1380
      TabIndex        =   8
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   4635
      Width           =   1110
   End
   Begin VB.ListBox LstComprobante 
      Height          =   4350
      Left            =   2490
      TabIndex        =   7
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   4635
      Width           =   810
   End
   Begin VB.ListBox LstCuentaOrigen 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   4290
      TabIndex        =   6
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   4635
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DtgTransferencias 
      Height          =   2565
      Left            =   1380
      TabIndex        =   22
      Top             =   1410
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   4524
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
            Type            =   1
            Format          =   "0.00"
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
   Begin VB.ListBox LstBDestino 
      Height          =   3765
      Left            =   14670
      TabIndex        =   62
      Top             =   510
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.ListBox LstRep 
      Height          =   3765
      Left            =   13620
      TabIndex        =   49
      Top             =   525
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ListBox LstCar 
      Height          =   3765
      Left            =   14070
      TabIndex        =   50
      Top             =   525
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.ListBox LstHono 
      Height          =   3765
      Left            =   14415
      TabIndex        =   61
      Top             =   510
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label25 
      Caption         =   "Lit."
      Height          =   195
      Left            =   9075
      TabIndex        =   60
      Top             =   4290
      Width           =   300
   End
   Begin VB.Label Label21 
      Caption         =   "ges"
      Height          =   405
      Left            =   15240
      TabIndex        =   48
      Top             =   4335
      Width           =   270
   End
   Begin VB.Label Label20 
      Caption         =   "Org"
      Height          =   240
      Left            =   14775
      TabIndex        =   47
      Top             =   4350
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label19 
      Caption         =   "Banco Des."
      Height          =   240
      Left            =   14175
      TabIndex        =   43
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblObs 
      Caption         =   "Observación"
      Height          =   210
      Left            =   13530
      TabIndex        =   42
      Top             =   4395
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label14 
      Caption         =   "Dólares"
      Height          =   195
      Left            =   12765
      TabIndex        =   40
      Top             =   4395
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label13 
      Caption         =   "Ciudad"
      Height          =   180
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   915
   End
   Begin VB.Label Label18 
      Caption         =   "Banco"
      Height          =   315
      Left            =   5445
      TabIndex        =   35
      Top             =   4275
      Width           =   960
   End
   Begin VB.Label Label12 
      Caption         =   "Ciudad"
      Height          =   180
      Left            =   11385
      TabIndex        =   33
      Top             =   4290
      Width           =   915
   End
   Begin VB.Label Label11 
      Caption         =   "Monto Bolivianos"
      Height          =   270
      Left            =   7470
      TabIndex        =   32
      Top             =   4275
      Width           =   1380
   End
   Begin VB.Label Label10 
      Caption         =   "Transferencia"
      Height          =   255
      Left            =   1365
      TabIndex        =   31
      Top             =   4275
      Width           =   1155
   End
   Begin VB.Label Label9 
      Caption         =   "Cuenta Origen"
      Height          =   315
      Left            =   4245
      TabIndex        =   30
      Top             =   4275
      Width           =   1245
   End
   Begin VB.Label Label8 
      Caption         =   " Cmpte"
      Height          =   300
      Left            =   2445
      TabIndex        =   29
      Top             =   4275
      Width           =   765
   End
   Begin VB.Label Label7 
      Caption         =   "Cuenta Destino"
      Height          =   360
      Left            =   10545
      TabIndex        =   28
      Top             =   4290
      Width           =   1110
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   75
      Left            =   10875
      TabIndex        =   27
      Top             =   1725
      Width           =   45
   End
   Begin VB.Label Label4 
      Caption         =   "COLA DE TRANSFERENCIAS  A IMPRESION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   1365
      TabIndex        =   26
      Top             =   3990
      Width           =   4200
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha"
      Height          =   270
      Left            =   3225
      TabIndex        =   25
      Top             =   4275
      Width           =   930
   End
   Begin VB.Label Label16 
      Caption         =   "Cuenta"
      Height          =   165
      Left            =   6570
      TabIndex        =   24
      Top             =   4290
      Width           =   600
   End
   Begin VB.Label Label17 
      Caption         =   "Justificacion"
      Height          =   210
      Left            =   9375
      TabIndex        =   23
      Top             =   4290
      Width           =   945
   End
End
Attribute VB_Name = "FrmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  SAF-2000
' Módulo:                   Control de Impresión de Transferencias
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmTransferencia.frm
' Descipción :              Comprobantes pagados
' Formularios relacionados: Main.frm (Padre)
'                           CryTransferencia
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creación         15/Ene/ 2000
' Fecha última modificación 01/May/ 2000
' Versión:                  2.0
'========================================================================================

Dim rsTransferencia As New ADODB.Recordset
Dim rsCorrel As New ADODB.Recordset
Dim rsTransfAux As New ADODB.Recordset
Dim punto As Integer
'Dim CryTrans As New CryTransferencia


Private Sub CmdBuscar_Click()
Dim condicion As String


                    If TxtCmpte.Text = "" Then
                        MsgBox "Necesita números de comprobante"
                        Exit Sub
                    Else
                        condicion = "pago_detalle.codigo_pago=" + "'" + TxtCmpte.Text + "'"
                    End If
                    
                    Set rsTransferencia = New ADODB.Recordset
                    'rsTransferencia.Open "SELECT pago_detalle.fecha_pago, Pagos.codigo_pago, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.monto_bolivianos, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, pago_detalle.org_codigo, pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo "
                     rsTransferencia.Open "SELECT pago_detalle.fecha_pago, Pagos.codigo_pago, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.monto_bolivianos, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, pago_detalle.org_codigo, pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
                    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' and " & condicion & " order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
                    If rsTransferencia.RecordCount > 0 Then
                        Set DtgTransferencias.DataSource = rsTransferencia
                    Else
                        MsgBox "Puede tratarse de cheque o no existe el registro porque ya fué aprobado", vbInformation
                    End If
                     FraBusca.Visible = False
End Sub

Private Sub CmdBusqueda_Click()
    FraBusca.Visible = True
End Sub

Private Sub CmdCancelar_Click()
    FraBusca.Visible = False
End Sub

Private Sub CmdFiltro_Click()
    Dim Resp As String
    Resp = InputBox("Introducir Organismo")
    If Resp <> "" Then
    Set rsTransferencia = New ADODB.Recordset
    If rsTransferencia.State = 1 Then rsTransferencia.Close
    rsTransferencia.Open "SELECT pago_detalle.fecha_pago, Pagos.codigo_pago, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.monto_bolivianos, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, pago_detalle.org_codigo, pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
                         "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where  pago_detalle.org_codigo='" & Resp & "' and pago_detalle.cheque_o_trf= 'T' order by pago_detalle.codigo_pago ", db, adOpenKeyset, adLockOptimistic
    If rsTransferencia.RecordCount > 0 Then
        Set DtgTransferencias.DataSource = rsTransferencia
    Else
        MsgBox "No existe la transferencia o los datos son incoherentes", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    End If
    
End Sub

Private Sub CmdImprimir_Click()
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim dia As String
    Dim mes As String
    Dim anio As String
    Dim FECHA As String
    
    
    Dim pname As String         'Stores the printer name
    Dim pport As String         'Stores the printer port information
    Dim pdriver As String       'Stores the printer driver information

'    pname = "HP LaserJet 4 Plus"
'    pport = "\\Jlv002\hp"
'    'pport = "\\Adb002\hp"
'    pdriver = "HP LaserJet 4 Plus"
'    Call CryTrans.SelectPrinter(pdriver, pname, pport)

    
    'pport = "\\Adb002\hp"
''''    pname = "HP LaserJet 4"
''''    pport = "\\Mrh002\hp"
''''    pdriver = "HP LaserJet 4"
''''    Call CryTrans.SelectPrinter(pdriver, pname, pport)
    
    If LstComprobante.ListCount = 0 Then
        MsgBox "No existen comprobantes", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
    'Limpiando la tabla auxiliar para cartas de transferencia
     Set rsTransfAux = New ADODB.Recordset
     If rsTransfAux.State = 1 Then rsTransferencia.Close
     rsTransfAux.Open "SELECT * FROM to_transferencia", db, adOpenKeyset, adLockOptimistic
     While Not rsTransfAux.EOF
         rsTransfAux.Delete
         rsTransfAux.MoveNext
     Wend
     
''''Comprobando si numeros de transferencia ya asignados
'''     sw = 0
'''     For j = 0 To LstComprobante.ListCount - 1
'''           LstTransf.ListIndex = 0
'''           If LstTransf.Text <> "" Then
'''                  MsgBox "El comprobante " + LstComprobante.Text + " ya tiene Nro. Transf. " + LstTransf.Text + "  En estos casos realizar REIMPRESION... "
'''                  CmdLimpiar_Click
'''                  Exit Sub
'''           End If
'''     Next j

     For j = LstTransf.ListCount - 1 To 0 Step -1
         LstTransf.RemoveItem j
     Next j
     'LstTransf.ListIndex = 1
     'Grabando los datos a la tabla auxiliar
     Set rsTransferencia = New ADODB.Recordset
     If rsTransferencia.State = 1 Then rsTransferencia.Close
     rsTransferencia.Open "SELECT * FROM to_transferencia", db, adOpenKeyset, adLockOptimistic
          For i = 0 To LstComprobante.ListCount - 1
              rsTransferencia.AddNew
              LstComprobante.ListIndex = i
              If LstComprobante.Text <> "" Then rsTransferencia("Nro_Cmpte") = Val(LstComprobante.Text)
              
              'If LstFecha.Text <> "" Then rsTransferencia("fecha_pago") = Date
              'Obtención de fecha
             ' LstFecha.ListIndex = I
             ' If LstFecha.Text <> "" Then
                  FECHA = Date
                  dia = Day(FECHA)
                  mes = Month(FECHA)
                  anio = Year(FECHA)
             ' Else
             '     MsgBox "no existe fecha en uno de los registros"
             '     Exit Sub
             ' End If
                Select Case mes
                           Case 1
                               mes = "ENERO"
                           Case 2
                               mes = "FEBRERO"
                           Case 3
                               mes = "MARZO"
                           Case 4
                               mes = "ABRIL"
                           Case 5
                               mes = "MAYO"
                           Case 6
                               mes = "JUNIO"
                           Case 7
                               mes = "JULIO"
                           Case 8
                               mes = "AGOSTO"
                           Case 9
                               mes = "SEPTIEMBRE"
                           Case 10
                               mes = "OCTUBRE"
                           Case 11
                               mes = "NOVIEMBRE"
                           Case 12
                               mes = "DICIEMBRE"
                        End Select
                        
                rsTransferencia("dia") = dia
                rsTransferencia("mes") = mes
                rsTransferencia("anio") = anio
     
              LstCuentaOrigen.ListIndex = i
              If LstCuentaOrigen.Text <> "" Then rsTransferencia("Cta_origen") = LstCuentaOrigen.Text
              LstBanco.ListIndex = i
              If LstBanco.Text <> "" Then rsTransferencia("banco") = LstBanco.Text
              LstDesCuenta.ListIndex = i
              If LstDesCuenta.Text <> "" Then rsTransferencia("Cta_origen_descripcion") = LstDesCuenta.Text
              If OptBol.Value = True Then
                LstMontoBol.ListIndex = i
                If LstMontoBol.Text <> "" Then rsTransferencia("monto") = LstMontoBol.Text: rsTransferencia("moneda") = "Bs.:"
              End If
              If OptDol.Value = True Then
                LstDolares.ListIndex = i
                If LstDolares.Text <> "" Then rsTransferencia("monto") = LstDolares.Text: rsTransferencia("moneda") = "$us.:"
              End If
              LstJustificacion.ListIndex = i
              If LstJustificacion.Text <> "" Then rsTransferencia("justificacion") = LstJustificacion.Text
              If LstCuentaDes.ListCount > 0 Then LstCuentaDes.ListIndex = i
              If LstCuentaDes.Text <> "" Then rsTransferencia("cta_destino") = LstCuentaDes.Text
              LstDepto.ListIndex = i
              If LstDepto.Text <> "" Then rsTransferencia("departamento") = LstDepto.Text
              If OptBol.Value = True Then
                LstLiteral.ListIndex = i
                If LstLiteral.Text <> "" Then rsTransferencia("literal") = Literal(LstMontoBol.Text) + " BOLIVIANOS" 'LstLiteral.Text
              Else
                LstLiteral.ListIndex = i
                If LstLiteral.Text <> "" Then rsTransferencia("literal") = Literal(LstDolares.Text) + " DOLARES" 'LstLiteral.Text
              End If
              If LstObs.ListCount > 0 Then
              LstObs.ListIndex = i
              If LstObs.Text <> "" Then rsTransferencia("obs") = LstObs.Text
              End If
              LstBancoDestino.ListIndex = i
              If LstBancoDestino.Text <> "" Then rsTransferencia("banco_destino") = LstBancoDestino.Text
         
                 'Buscando Nro. de correlativo de Transferencia
                  If rsCorrel.State = 1 Then rsCorrel.Close
                  Set rsCorrel = New ADODB.Recordset
                  rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= 'Transf' ", db, adOpenKeyset, adLockOptimistic
                  If rsCorrel.RecordCount > 0 Then
                     rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") + 1
                     rsCorrel.Update
                  Else
                     rsCorrel("numero_correlativo") = 0
                     rsCorrel.Update
                  End If
                  rsTransferencia("Nro_Transferencia") = rsCorrel("numero_correlativo")
                  LstGes.ListIndex = i
                  rsTransferencia("ges_gestion") = LstGes.Text
                  LstOrg.ListIndex = i
                  rsTransferencia("cod_org") = LstOrg.Text
                  LstRep.ListIndex = i
                  rsTransferencia("representante") = LstRep.Text
                  LstCar.ListIndex = i
                  rsTransferencia("cargo") = LstCar.Text
                  LstHono.ListIndex = i
                  If LstHono.Text = "H" Then
                     rsTransferencia("Honorarios") = "(Pago de Honorarios)"
                  End If
                  If LstHono.Text = "S" Then
                     rsTransferencia("Honorarios") = " "
                  End If
                   
                  If LstBDestino.ListCount > 0 Then
                     LstBDestino.ListIndex = i
                     rsTransferencia("beneficiario_destino") = LstBDestino.Text
                  End If
                    
                  rsTransferencia.Update
                  LstTransf.AddItem rsCorrel("numero_correlativo")
   Next i
     
   sino = MsgBox("Se imprimiran los comprobantes sin Nro. Transf ...!", vbYesNo, "Mensaje de Advertencia")
   If sino = vbYes Then
        Cmpte_NroTransferencia
        RepTransferencia.Show
        '''CryTrans.Database.Verify
        '''CryTrans.PrintOut
        
        Transferencia_Aprobados
   Else
        'restaurar_NroTransferencia
         If rsCorrel.State = 1 Then rsCorrel.Close
         rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= 'Transf' ", db, adOpenKeyset, adLockOptimistic
         If rsCorrel.RecordCount > 0 Then
            rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") - LstComprobante.ListCount
            rsCorrel.Update
         End If
        Exit Sub
   End If
   sw = 0
   Cola_Impresion
   'coloca_status_impresion_transferencia
   
End Sub
Public Sub Cola_Impresion()
    Dim SqlQuery As String
    'Mandando a la cola de impresión los cheques
    
     Set rsIT = New ADODB.Recordset
     If rsIT.State = 1 Then rsTransferencia.Close
     rsIT.Open "SELECT * FROM to_Transferencia", db, adOpenKeyset, adLockOptimistic
     If rsIT.RecordCount > 0 Then
     While Not rsIT.EOF
            Set rsComprobante = New ADODB.Recordset
            SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
                       "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.codigo_pago=" & rsIT("nro_cmpte") & " and pago_detalle.Ges_gestion= '" & rsIT("ges_gestion") & "' and pago_detalle.cta_codigo='" & rsIT("cta_origen") & "' order by Pago_detalle.codigo_pago"
            rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
            If rsComprobante.RecordCount > 0 Then
                 Set rsCmpteI = New ADODB.Recordset
                 If rsCmpteI.State = 1 Then rsCmpteI.Close
                 rsCmpteI.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
                 'If rsCmpteI.RecordCount > 0 Then
                        rsCmpteI.AddNew
                        rsCmpteI("Nro_Cmpte") = rsComprobante("codigo_pago")
                        rsCmpteI("Organismo") = rsComprobante("cta_descripcion_larga")
                        rsCmpteI("Fecha_Pago") = Format(rsComprobante("Fecha_pago"), "dd/mm/yyyy")
                        rsCmpteI("Monto") = rsComprobante("monto_bolivianos")
                        rsCmpteI("Cambio") = rsComprobante("tipo_cambio")
                        rsCmpteI("Beneficiario") = rsComprobante("denominacion_beneficiario")
                        rsCmpteI("Justificacion") = rsComprobante("Justificacion")
                        rsCmpteI("Nro_cheque") = rsComprobante("numero_cheque_trf")
                        rsCmpteI("banco") = rsComprobante("Bco_descripcion_larga")
                        rsCmpteI("Transf_cheq") = "TRANSFERENCIA"
                        rsCmpteI("Literal") = Literal(Str(rsComprobante("monto_bolivianoS")))
                    rsCmpteI.Update
                 ' End If
            End If
            rsIT.MoveNext
      Wend
     End If
End Sub

Private Sub CmdLimpiar_Click()
        LstComprobante.Clear
        LstFecha.Clear
        LstTransf.Clear
        LstCuentaOrigen.Clear
        LstDesCuenta.Clear
        LstMontoBol.Clear
        LstBanco.Clear
        LstJustificacion.Clear
        LstCuentaDes.Clear
        LstDepto.Clear
        LstLiteral.Clear
        LstDolares.Clear
        LstObs.Clear
        LstOrg.Clear
        LstGes.Clear
        LstRep.Clear
        LstCar.Clear
        LstHono.Clear
        LstBDestino.Clear
        LstBancoDestino.Clear
End Sub

Private Sub CmdReimpresion_Click()
    Dim i As Integer
    Dim K As Integer
    Dim dia As String
    Dim mes As String
    Dim anio As String
    Dim FECHA As String
    
    
    Dim pname As String         'Stores the printer name
    Dim pport As String         'Stores the printer port information
    Dim pdriver As String       'Stores the printer driver information

'    pname = "HP LaserJet 4 Plus"
'    pport = "\\Jlv002\hp"
'    'pport = "\\Adb002\hp"
'    pdriver = "HP LaserJet 4 Plus"
'    Call CryTrans.SelectPrinter(pdriver, pname, pport)
'
'    pname = "HP LaserJet 4"
'    pport = "\\lvq002\hp"
'    'pport = "\\Adb002\hp"
'    pdriver = "HP LaserJet 4"
'    Call CryTrans.SelectPrinter(pdriver, pname, pport)

    
    If LstComprobante.ListCount = 0 Then
        MsgBox "No existen comprobantes", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
    'Limpiando la tabla auxiliar para cartas de transferencia
     Set rsTransfAux = New ADODB.Recordset
     If rsTransfAux.State = 1 Then rsTransferencia.Close
     rsTransfAux.Open "SELECT * FROM to_transferencia", db, adOpenKeyset, adLockOptimistic
     While Not rsTransfAux.EOF
         rsTransfAux.Delete
         rsTransfAux.MoveNext
     Wend
     
     
    'Grabando los datos a la tabla auxiliar
     Set rsTransferencia = New ADODB.Recordset
     If rsTransferencia.State = 1 Then rsTransferencia.Close
     rsTransferencia.Open "SELECT * FROM to_transferencia", db, adOpenKeyset, adLockOptimistic
          For i = 0 To LstComprobante.ListCount - 1
              rsTransferencia.AddNew
              LstTransf.ListIndex = i
              If LstTransf.Text = "" Then
                MsgBox "No existen numeración de transferencias, debe tener asignado", vbInformation + vbCritical
                Exit Sub
              End If
              If LstTransf.Text <> "" Then rsTransferencia("Nro_Transferencia") = LstTransf.Text
              LstComprobante.ListIndex = i
              If LstComprobante.Text <> "" Then rsTransferencia("Nro_Cmpte") = LstComprobante.Text
                  FECHA = Date
                  dia = Day(FECHA)
                  mes = Month(FECHA)
                  anio = Year(FECHA)
                Select Case mes
                           Case 1
                               mes = "ENERO"
                           Case 2
                               mes = "FEBRERO"
                           Case 3
                               mes = "MARZO"
                           Case 4
                               mes = "ABRIL"
                           Case 5
                               mes = "MAYO"
                           Case 6
                               mes = "JUNIO"
                           Case 7
                               mes = "JULIO"
                           Case 8
                               mes = "AGOSTO"
                           Case 9
                               mes = "SEPTIEMBRE"
                           Case 10
                               mes = "OCTUBRE"
                           Case 11
                               mes = "NOVIEMBRE"
                           Case 12
                               mes = "DICIEMBRE"
                        End Select
                        
                rsTransferencia("dia") = dia
                rsTransferencia("mes") = mes
                rsTransferencia("anio") = anio
     
              LstCuentaOrigen.ListIndex = i
              If LstCuentaOrigen.Text <> "" Then rsTransferencia("Cta_origen") = LstCuentaOrigen.Text
              LstBanco.ListIndex = i
              If LstBanco.Text <> "" Then rsTransferencia("banco") = LstBanco.Text
              LstDesCuenta.ListIndex = i
              If LstDesCuenta.Text <> "" Then rsTransferencia("Cta_origen_descripcion") = LstDesCuenta.Text
              If OptBol.Value = True Then
                LstMontoBol.ListIndex = i
                If LstMontoBol.Text <> "" Then rsTransferencia("monto") = LstMontoBol.Text: rsTransferencia("moneda") = "Bs.:"
              End If
              If OptDol.Value = True Then
                LstDolares.ListIndex = i
                If LstDolares.Text <> "" Then rsTransferencia("monto") = LstDolares.Text: rsTransferencia("moneda") = "$us.:"
              End If
              LstJustificacion.ListIndex = i
              If LstJustificacion.Text <> "" Then rsTransferencia("justificacion") = LstJustificacion.Text
              If LstCuentaDes.ListCount > 0 Then LstCuentaDes.ListIndex = i
              If LstCuentaDes.Text <> "" Then rsTransferencia("cta_destino") = LstCuentaDes.Text
              LstDepto.ListIndex = i
              If LstDepto.Text <> "" Then rsTransferencia("departamento") = LstDepto.Text
              If OptBol.Value = True Then
                LstLiteral.ListIndex = i
                If LstLiteral.Text <> "" Then rsTransferencia("literal") = Literal(LstMontoBol.Text) + " BOLIVIANOS" 'LstLiteral.Text
              Else
                LstLiteral.ListIndex = i
                If LstLiteral.Text <> "" Then rsTransferencia("literal") = Literal(LstDolares.Text) + " DOLARES" 'LstLiteral.Text
              End If
              If LstObs.ListCount > 0 Then
              LstObs.ListIndex = i
              If LstObs.Text <> "" Then rsTransferencia("obs") = LstObs.Text
              End If
              LstBancoDestino.ListIndex = i
              If LstBancoDestino.Text <> "" Then rsTransferencia("banco_destino") = LstBancoDestino.Text
              
              LstGes.ListIndex = i
              rsTransferencia("ges_gestion") = LstGes.Text
              LstOrg.ListIndex = i
              rsTransferencia("cod_org") = LstOrg.Text
              LstRep.ListIndex = i
              rsTransferencia("representante") = LstRep.Text
              LstCar.ListIndex = i
              rsTransferencia("cargo") = LstCar.Text
              LstHono.ListIndex = i
              If LstHono.Text = "H" Then
                  rsTransferencia("Honorarios") = "(Pago de Honorarios)"
              End If
              If LstHono.Text = "S" Then
                  rsTransferencia("Honorarios") = " "
              End If
         
              If LstBDestino.ListCount > 0 Then
                    LstBDestino.ListIndex = i
                    rsTransferencia("beneficiario_destino") = LstBDestino.Text
              End If
              LstComprobante.ListIndex = i
              If LstComprobante.Text <> "" Then rsTransferencia("Nro_Cmpte") = LstComprobante.Text
              rsTransferencia.Update
         
   Next i
     
   sino = MsgBox("Se REIMPRIMIRAN los comprobantes ...!", vbYesNo, "Mensaje de Advertencia")
   If sino = vbYes Then
         RepTransferencia.Show
'        CryTrans.Database.Verify
'        CryTrans.PrintOut
   End If
   sw = 0
End Sub

Private Sub CmdRestaurar_Click()
    If rsTransferencia.State = 1 Then rsTransferencia.Close
    Set rsTransferencia = New ADODB.Recordset
'    rsTransferencia.Open "SELECT pago_detalle.fecha_pago, Pagos.codigo_pago, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.monto_bolivianos, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, pago_detalle.org_codigo, pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo "
    rsTransferencia.Open "SELECT pago_detalle.fecha_pago, Pagos.codigo_pago, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.monto_bolivianos, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, pago_detalle.org_codigo, pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T'", db, adOpenKeyset, adLockOptimistic
    Set DtgTransferencias.DataSource = rsTransferencia
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
    FrmColaImpresion.Show
End Sub

Private Sub DtgTransferencias_Click()
 Dim bandera As Integer
 Dim z As Integer
    bandera = 0
    z = 0
    For i = 0 To LstComprobante.ListCount - 1
         LstComprobante.ListIndex = i
         If LstComprobante.Text = DtgTransferencias.Columns(1) Then
              bandera = 1
         End If
    Next i
    If bandera = 0 Then
        LstComprobante.AddItem DtgTransferencias.Columns(1)
        LstFecha.AddItem DtgTransferencias.Columns(0)
        LstTransf.AddItem DtgTransferencias.Columns(2)
        LstCuentaOrigen.AddItem DtgTransferencias.Columns(3)
        LstBanco.AddItem DtgTransferencias.Columns(4)
        LstDesCuenta.AddItem DtgTransferencias.Columns(5)
        LstMontoBol.AddItem DtgTransferencias.Columns(6)
        LstLiteral.AddItem DtgTransferencias.Columns(7)
        LstDepto.AddItem DtgTransferencias.Columns(8)
        LstJustificacion.AddItem DtgTransferencias.Columns(9)
        LstCuentaDes.AddItem DtgTransferencias.Columns(10)
        LstDolares.AddItem DtgTransferencias.Columns(12)
        LstObs.AddItem DtgTransferencias.Columns(13)
        LstBancoDestino.AddItem DtgTransferencias.Columns(14)
        LstOrg.AddItem DtgTransferencias.Columns(15)
        LstGes.AddItem DtgTransferencias.Columns(16)
        LstRep.AddItem DtgTransferencias.Columns(17)
        LstCar.AddItem DtgTransferencias.Columns(18)
        LstHono.AddItem DtgTransferencias.Columns(19)
        LstBDestino.AddItem DtgTransferencias.Columns(20)
    End If
End Sub

Private Sub Form_Load()
    Set rsTransferencia = New ADODB.Recordset
    rsTransferencia.Open "SELECT pago_detalle.fecha_pago, Pagos.codigo_pago, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.monto_bolivianos, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, pago_detalle.org_codigo, pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
                         "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic

    Set DtgTransferencias.DataSource = rsTransferencia
	Call SeguridadSet(Me)
End Sub

Public Sub Cmpte_NroTransferencia()
'========================================================================================
' Módulo:                   Cmpte_NroTransferencia
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmTransferencia.frm
' Descipción :              Actualización de Nros. de transferencia en el registro de
'                           pago_detalle
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================
Dim NumeroTransferencia As String
If rsTransfAux.State = 1 Then rsTransfAux.Close
Set rsTransfAux = New ADODB.Recordset
rsTransfAux.Open "select * FROM to_transferencia", db, adOpenKeyset, adLockOptimistic
If rsTransfAux.RecordCount > 0 Then
        While Not rsTransfAux.EOF
            Set rsPagoDet = New ADODB.Recordset
            'rsPagoDet.Open "select * from pago_detalle where codigo_pago='" & rsTransfAux("Nro_Cmpte") & "'", db, adOpenKeyset, adLockOptimistic
             rsPagoDet.Open "select * from pago_detalle where codigo_pago='" & rsTransfAux("Nro_Cmpte") & "' and ges_gestion='" & rsTransfAux("ges_gestion") & "' and org_codigo='" & rsTransfAux("cod_org") & "'", db, adOpenKeyset, adLockOptimistic
                Select Case Len(rsTransfAux("Nro_Transferencia"))
                    Case 1
                        NumeroTransferencia = "0000" + rsTransfAux("Nro_Transferencia")
                    Case 2
                        NumeroTransferencia = "000" + rsTransfAux("Nro_Transferencia")
                    Case 3
                        NumeroTransferencia = "00" + rsTransfAux("Nro_Transferencia")
                    Case 4
                        NumeroTransferencia = "0" + rsTransfAux("Nro_Transferencia")
                    Case 5
                        NumeroTransferencia = rsTransfAux("Nro_Transferencia")
                End Select
                rsPagoDet("numero_cheque_trf") = NumeroTransferencia
                rsPagoDet("fecha_impresion_cheque") = Date
                rsPagoDet.Update
            rsTransfAux.MoveNext
        Wend
End If
End Sub

Private Sub LstBanco_DblClick()
    LstBanco.RemoveItem punto
    LstDesCuenta_DblClick
End Sub

Private Sub LstBancoDestino_DblClick()
    LstBancoDestino.RemoveItem punto
    LstOrg_DblClick
End Sub

Private Sub LstBDestino_DblClick()
    LstBDestino.RemoveItem punto
End Sub

Private Sub LstCar_DblClick()
    LstCar.RemoveItem punto
    LstHono_DblClick
End Sub

Private Sub LstComprobante_DblClick()
    punto = LstComprobante.ListIndex
    LstComprobante.RemoveItem punto
    LstFecha_DblClick
End Sub

Private Sub LstCuentaDes_DblClick()
    LstCuentaDes.RemoveItem punto
    LstDepto_DblClick
End Sub

Private Sub LstCuentaOrigen_DblClick()
    LstCuentaOrigen.RemoveItem punto
    LstBanco_DblClick
End Sub

Private Sub LstDepto_DblClick()
    LstDepto.RemoveItem punto
    LstDolares_DblClick
End Sub

Private Sub LstDesCuenta_DblClick()
    LstDesCuenta.RemoveItem punto
    LstMontoBol_DblClick
End Sub

Private Sub LstDolares_DblClick()
    LstDolares.RemoveItem punto
    LstObs_DblClick
End Sub

Private Sub LstFecha_DblClick()
    LstFecha.RemoveItem punto
    LstCuentaOrigen_DblClick
End Sub

Private Sub LstGes_DblClick()
    LstGes.RemoveItem punto
    LstRep_DblClick
End Sub

Private Sub LstHono_DblClick()
    LstHono.RemoveItem punto
    LstBDestino_DblClick

End Sub

Private Sub LstJustificacion_DblClick()
    LstJustificacion.RemoveItem punto
    LstCuentaDes_DblClick
End Sub

Private Sub LstLiteral_DblClick()
    LstLiteral.RemoveItem punto
    LstJustificacion_DblClick
End Sub

Private Sub LstMontoBol_DblClick()
    LstMontoBol.RemoveItem punto
    LstLiteral_DblClick
End Sub

Private Sub LstObs_DblClick()
    LstObs.RemoveItem punto
    LstBancoDestino_DblClick
End Sub

Private Sub LstOrg_DblClick()
    LstOrg.RemoveItem punto
    LstGes_DblClick
End Sub
Private Sub LstRep_DblClick()
    LstRep.RemoveItem punto
    LstCar_DblClick
End Sub

Public Sub Transferencia_Aprobados()
        'Determinando comprobante de pagos en detalle como APROBADOS CHEQUES
        For i = 0 To LstTransf.ListCount - 1
          LstTransf.ListIndex = i
          LstComprobante.ListIndex = i
          LstOrg.ListIndex = i
          LstGes.ListIndex = i
          'NroCheque = LstNroCheque.Text
            Set rspago = New ADODB.Recordset
            If rspago.State = 1 Then rspago.Close
            rspago.Open "SELECT * from pagos where codigo_pago= '" & LstComprobante.Text & "' and ges_gestion= '" & LstGes.Text & "' and org_codigo='" & LstOrg.Text & "'", db, adOpenKeyset, adLockOptimistic
            If rspago.RecordCount > 0 Then
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & LstComprobante.Text & "' and ges_gestion= '" & LstGes.Text & "' and org_codigo='" & LstOrg.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                     rsPAgoDetalle("estado_aprobacion") = "A"
                     rsPAgoDetalle.Update
                End If
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & LstComprobante.Text & "' and estado_aprobacion<>'A' and ges_gestion= '" & LstGes.Text & "' and org_codigo='" & LstOrg.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                    SumaMontosParciales = 0
                    While Not rsPAgoDetalle.EOF
                         SumaMontosParciales = SumaMontosParciales + rsPAgoDetalle("monto_bolivianos")
                         rsPAgoDetalle.MoveNext
                    Wend
                    If rspago("liquido_pagar") = SumaMontosParciales And SumaMontosParciales <> 0 Then
                     rspago("estado_aprobacion") = "A"
                     rspago("estado_pagado") = "S" 'Total
                     rspago.Update
                    Else
                     rspago("estado_pagado") = "P" 'Parcial
                     rspago.Update
                    End If
                End If
        
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
           End If
     Next i
End Sub


Public Sub coloca_status_impresion_transferencia()
''Colocando el status de S cuando imprime una transferencia
'
'Dim AUX, NUMERO As String
'Dim Car As String
'Dim i As Integer
'Dim LONGITUD As Integer
'
'NUMERO = ""
'AUX = TxtCheques.Text
'LONGITUD = Len(AUX)
'  While (LONGITUD + 1 > 0)
'      i = i + 1
'      Car = Mid(AUX, i, 1)
'      LONGITUD = LONGITUD - 1
'      If Car <> "," And Car <> "" Then
'         NUMERO = NUMERO + Car
'      Else
'                MsgBox NUMERO
'                T = CStr(NUMERO)
'                Select Case Len(T)
'                       Case 1
'                            S = "0000" + CStr(NUMERO)
'                       Case 2
'                            S = "000" + CStr(NUMERO)
'                       Case 3
'                            S = "00" + CStr(NUMERO)
'                       Case 4
'                            S = "0" + CStr(NUMERO)
'                       Case 5
'                            S = CStr(NUMERO)
'                End Select
'                Set rsComprobante = New ADODB.Recordset
'                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
'                                   "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.numero_cheque_trf='" & S & "' and pago_detalle.cheque_o_trf='C'", db, adOpenKeyset, adLockOptimistic
'                If rsComprobante.RecordCount > 0 Then
'                If rsCheques.State = 1 Then rsCheques.Close
'                rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & S & "'order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
'                If rsCheques.RecordCount > 0 Then
'                        rsCheques("estado_anulado") = "S"
'                Else
'                        rsCheques.AddNew
'                        rsCheques("numero_cheque") = S
'                        rsCheques("estado_anulado") = "S"
'                End If
'                rsCheques("usr_usuario") = LblUsuario.Caption
'                rsCheques("fecha_registro") = Date
'                rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
'
'                rsCheques.Update
'             End If
'            NUMERO = ""
'         End If
'  Wend
'
'

End Sub
Private Sub TxtCmpte_Validate(Cancel As Boolean)
'                    If Not IsNumeric(TxtCmpte.Text) Or Val(TxtCmpte.Text) > 0 Then
'                       KeepFocus = True
'                       MsgBox _
'                       "Escriba un número mayor a 0.", , "TxtCmpte.text"
'                    End If
End Sub

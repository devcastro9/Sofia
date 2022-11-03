VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmImprimirComprobante 
   Caption         =   "Imprimir Comprobantes"
   ClientHeight    =   8595
   ClientLeft      =   360
   ClientTop       =   1620
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FraBusca 
      Height          =   2085
      Left            =   2145
      TabIndex        =   46
      Top             =   4590
      Visible         =   0   'False
      Width           =   2040
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   50
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2047
         TabIndex        =   49
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   225
         TabIndex        =   48
         Top             =   780
         Width           =   1515
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   225
         TabIndex        =   47
         Top             =   1245
         Width           =   1515
      End
      Begin VB.Label Label20 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   53
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label19 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   2310
         TabIndex        =   52
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label21 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   450
         TabIndex        =   51
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.ListBox LstTransf_Cheq 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   11790
      TabIndex        =   44
      Top             =   5835
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprobantes de :"
      Height          =   1980
      Left            =   9795
      TabIndex        =   40
      Top             =   3285
      Width           =   2310
      Begin VB.OptionButton OptTransferencias 
         Caption         =   "Tranferencias"
         Height          =   555
         Left            =   180
         TabIndex        =   42
         Top             =   990
         Width           =   1815
      End
      Begin VB.OptionButton OptCheques 
         Caption         =   "Cheques"
         Height          =   450
         Left            =   195
         TabIndex        =   41
         Top             =   375
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   11820
      TabIndex        =   5
      Top             =   0
      Width           =   11880
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
         TabIndex        =   10
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   7
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "IMPRESION DE COMPROBANTES"
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
         Left            =   3645
         TabIndex        =   6
         Top             =   135
         Width           =   5055
      End
   End
   Begin VB.ListBox LstMonto 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   4332
      TabIndex        =   39
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   5835
      Width           =   1230
   End
   Begin VB.ListBox LstComprobante 
      Height          =   4155
      Left            =   1350
      TabIndex        =   38
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   5805
      Width           =   915
   End
   Begin VB.ListBox LstFecha 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   3240
      TabIndex        =   37
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   5835
      Width           =   1110
   End
   Begin VB.ListBox LstBeneficiario 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   6470
      TabIndex        =   36
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   5835
      Width           =   1635
   End
   Begin VB.ListBox LstNroCheque 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   9210
      TabIndex        =   35
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   5835
      Width           =   990
   End
   Begin VB.ListBox LstBanco 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   10185
      TabIndex        =   34
      ToolTipText     =   "Hacer Click para borrar"
      Top             =   5835
      Width           =   1185
   End
   Begin VB.ListBox LstOrganismo 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   2269
      TabIndex        =   33
      Top             =   5835
      Width           =   975
   End
   Begin VB.ListBox LstCambio 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   5551
      TabIndex        =   32
      Top             =   5835
      Width           =   930
   End
   Begin VB.ListBox LstJustificacion 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   8085
      TabIndex        =   31
      Top             =   5835
      Width           =   1140
   End
   Begin VB.ListBox LstLiteral 
      Enabled         =   0   'False
      Height          =   4155
      Left            =   11370
      TabIndex        =   30
      Top             =   5835
      Width           =   420
   End
   Begin VB.Frame FraOpciones 
      Height          =   9060
      Left            =   0
      TabIndex        =   11
      Top             =   1005
      Width           =   1245
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Busqueda"
         Height          =   795
         Left            =   105
         Picture         =   "FrmImprimirComprobantes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4350
         Width           =   930
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar Grid"
         Height          =   795
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2760
         Width           =   930
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   795
         Left            =   105
         TabIndex        =   16
         Top             =   1965
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   105
         Picture         =   "FrmImprimirComprobantes.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5145
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   795
         Left            =   120
         Picture         =   "FrmImprimirComprobantes.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   375
         Width           =   915
      End
      Begin VB.CommandButton CmdFiltro 
         Caption         =   "Filtro por Organismo"
         Height          =   795
         Left            =   105
         TabIndex        =   13
         Top             =   3555
         Width           =   930
      End
      Begin VB.CommandButton CmdImpresionRangos 
         Caption         =   "Imprime X Rangos"
         Height          =   795
         Left            =   120
         Picture         =   "FrmImprimirComprobantes.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1170
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rangos de Comprobantes"
      Height          =   2190
      Left            =   9780
      TabIndex        =   0
      Top             =   1095
      Width           =   2325
      Begin VB.TextBox TxtInicio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   495
         Width           =   1395
      End
      Begin VB.TextBox TxtFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   285
         TabIndex        =   1
         Top             =   1215
         Width           =   1395
      End
      Begin VB.Label Label13 
         Caption         =   "Nro Cmpte Inicial"
         Height          =   210
         Left            =   315
         TabIndex        =   4
         Top             =   855
         Width           =   1320
      End
      Begin VB.Label Label14 
         Caption         =   "Nro Cmpte Final"
         Height          =   210
         Left            =   330
         TabIndex        =   3
         Top             =   1635
         Width           =   1320
      End
   End
   Begin MSDataGridLib.DataGrid DtGComprobantes 
      Height          =   4200
      Left            =   1305
      TabIndex        =   18
      Top             =   1095
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   7408
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
   Begin VB.Label Label22 
      Caption         =   "T/C"
      Height          =   240
      Left            =   11835
      TabIndex        =   54
      Top             =   5580
      Width           =   345
   End
   Begin VB.Label Label18 
      Caption         =   "Literal"
      Height          =   195
      Left            =   11310
      TabIndex        =   43
      Top             =   5580
      Width           =   435
   End
   Begin VB.Label Label17 
      Caption         =   "Justificación"
      Height          =   210
      Left            =   8070
      TabIndex        =   29
      Top             =   5595
      Width           =   945
   End
   Begin VB.Label Label16 
      Caption         =   "Cambio"
      Height          =   165
      Left            =   5550
      TabIndex        =   28
      Top             =   5595
      Width           =   600
   End
   Begin VB.Label Label15 
      Caption         =   "Organismo"
      Height          =   270
      Left            =   2280
      TabIndex        =   27
      Top             =   5580
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "COLA DE COMPROBANTES  A IMPRESION"
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
      Left            =   1305
      TabIndex        =   26
      Top             =   5355
      Width           =   5745
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   75
      Left            =   10875
      TabIndex        =   25
      Top             =   1725
      Width           =   45
   End
   Begin VB.Label Label7 
      Caption         =   "Cheque"
      Height          =   270
      Left            =   9210
      TabIndex        =   24
      Top             =   5595
      Width           =   945
   End
   Begin VB.Label Label8 
      Caption         =   "Comprobante"
      Height          =   300
      Left            =   1365
      TabIndex        =   23
      Top             =   5580
      Width           =   960
   End
   Begin VB.Label Label9 
      Caption         =   "Monto"
      Height          =   315
      Left            =   4305
      TabIndex        =   22
      Top             =   5580
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3210
      TabIndex        =   21
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label Label11 
      Caption         =   "Beneficiario"
      Height          =   270
      Left            =   6495
      TabIndex        =   20
      Top             =   5580
      Width           =   1380
   End
   Begin VB.Label Label12 
      Caption         =   "Banco"
      Height          =   180
      Left            =   10215
      TabIndex        =   19
      Top             =   5595
      Width           =   915
   End
End
Attribute VB_Name = "FrmImprimirComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsComprobante As New ADODB.Recordset
Dim rsCheque As New ADODB.Recordset
Dim rsCorrel As New ADODB.Recordset
Dim punto As Variant
Dim NumeroCuenta As String
Public cryCmpte As New CryComprobante

Private Sub CmdBuscar_Click()
                    If TxtCmpte.Text = "" Then
                        MsgBox "Necesita números de comprobante"
                        Exit Sub
                    Else
                        CONDICION = "pago_detalle.codigo_pago=" + "'" + TxtCmpte.Text + "'"
                    End If
                    
                    SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE " & CONDICION & " order by Pago_detalle.codigo_pago "
                    If rsComprobante.State Then rsComprobante.Close
                    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
                    If rsComprobante.RecordCount > 0 Then
                        MsgBox rsComprobante("monto_Bolivianos")
                        Set DtGComprobantes.DataSource = rsComprobante
                    End If
                    FraBusca.Visible = False

End Sub

Private Sub CmdBusqueda_Click()
    FraBusca.Visible = True
End Sub

Private Sub CmdFiltro_Click()
Dim SqlQuery As String
Dim Resp As String

    Resp = InputBox("Introducir Organismo o Cuenta Bancaria")
    If Resp <> "" Then
      Set rsCheque = New ADODB.Recordset
      If rsCheque.State = 1 Then rsCheque.Close
'      rsCheque.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf,  fc_bancos.Bco_descripcion_larga " & _
'      "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.cta_codigo= '" & Resp & "' and pago_detalle.estado_aprobacion <> 'A' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic

      'SqlQuery = "SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal,pago_detalle.cta_codigo "
      SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
                 "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.cta_codigo= '" & Resp & "'"
      rsCheque.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
      
      If rsCheque.RecordCount > 0 Then
        Set DtGComprobantes.DataSource = rsCheque
        DtGComprobantes.Refresh
      Else
        MsgBox "No existen registros de la cuenta" + " " + Resp
      End If
    End If
    
End Sub
Private Sub CmdImpresionRangos_Click()
Dim SqlQuery As String
     CmdLimpiar_Click
     If TxtInicio = "" Then
        MsgBox "Introducir comprobante inicial", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
     End If
     If Val(TxtInicio.Text) > Val(TxtFin.Text) Then
        MsgBox "Comprobante inicial menor al comprobante final", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
     End If
     
     'Limpiando la tabla auxiliar para cheques
     Set rsCmpte = New ADODB.Recordset
     If rsCmpte.State = 1 Then rsCheques.Close
     rsCmpte.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
     While Not rsCmpte.EOF
         rsCmpte.Delete
         rsCmpte.MoveNext
     Wend
     
     MsgBox "Se imprimirán los comprobantes por rango"
     If TxtInicio.Text <> "" And TxtFin.Text <> "" Then
        Set rsComprobante = New ADODB.Recordset
                
        If rsComprobante.State = 1 Then rsComprobante.Close
        SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo" ',db, adOpenKeyset, adLockOptimistic
        rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
        If rsComprobante.RecordCount > 0 Then
           Set DtGComprobantes.DataSource = rsComprobante
           While Not rsComprobante.EOF
               If Val(rsComprobante("codigo_pago")) >= Val(TxtInicio.Text) And Val(rsComprobante("codigo_pago")) <= Val(TxtFin.Text) Then
                      rsCmpte.AddNew
                      rsCmpte("Nro_Cmpte") = rsComprobante("codigo_pago")
                      rsCmpte("fecha_pago") = rsComprobante("fecha_pago")
                      rsCmpte("organismo") = rsComprobante("cta_descripcion_larga")
                      rsCmpte("monto") = rsComprobante("monto_bolivianos")
                      rsCmpte("cambio") = rsComprobante("tipo_cambio")
                      rsCmpte("beneficiario") = rsComprobante("denominacion_beneficiario")
                      rsCmpte("justificacion") = rsComprobante("justificacion")
                      rsCmpte("nro_cheque") = rsComprobante("numero_cheque_trf")
                      rsCmpte("Banco") = Trim(rsComprobante("Bco_descripcion_larga"))
                      rsCmpte("literal") = rsComprobante("literal")
                End If
                rsComprobante.MoveNext
              Wend
        End If
    End If
        sino = MsgBox("Se imprimiran los comprobantes ...! ", vbYesNo, "Mensaje de Advertencia")    '  sino = MsgBox("Està seguro de eliminar este registro", vbYesNo + vbQuestion, "Atenciòn") then
        If sino = vbYes Then
                 FrmComprobante.Show
        Else
                 Exit Sub
        End If
   
End Sub

Private Sub CmdImprimir_Click()
   
    Dim i As Integer
    Dim dia As String
    Dim mes As String
    Dim anio As String
    
    Dim pname As String         'Stores the printer name
    Dim pport As String         'Stores the printer port information
    Dim pdriver As String       'Stores the printer driver information

    pname = "Epson LX-810"
    'pport = "\\Jcc003\epson"
    pport = "LPT1:  (Puerto de impresora ECP)"
    pdriver = "Epson LX-810"
    Call cryCmpte.SelectPrinter(pdriver, pname, pport)

    TxtInicio.Text = ""
    TxtFin.Text = ""
    
    If LstComprobante.ListCount = 0 Then
        MsgBox "No existen comprobantes", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
    'Limpiando la tabla auxiliar para cheques
     Set rsComprobante = New ADODB.Recordset
     If rsComprobante.State = 1 Then rsCheques.Close
     rsComprobante.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
     While Not rsComprobante.EOF
         rsComprobante.Delete
         rsComprobante.MoveNext
     Wend
     
    'Grabando los datos a la tabla auxiliar
     Set rsComprobante = New ADODB.Recordset
     If rsComprobante.State = 1 Then rsComprobante.Close
     rsComprobante.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
          For i = 0 To LstComprobante.ListCount - 1
              LstComprobante.ListIndex = i
              rsComprobante.AddNew
              If LstComprobante.Text <> "" Then rsComprobante("nro_cmpte") = LstComprobante.Text
              LstOrganismo.ListIndex = i
              If LstOrganismo.Text <> "" Then rsComprobante("Organismo") = Trim(LstOrganismo.Text)
              'LstFecha.ListIndex = I
              'If LstFecha.Text <> "" Then rsComprobante("fecha_pago") = LstFecha.Text
              rsComprobante("fecha_pago") = Date
              LstMonto.ListIndex = i
              If LstMonto.Text <> "" Then rsComprobante("monto") = LstMonto.Text
              LstCambio.ListIndex = i
              If LstCambio.Text <> "" Then rsComprobante("cambio") = CDbl(LstCambio.Text)
              LstBeneficiario.ListIndex = i
              If LstBeneficiario.Text <> "" Then rsComprobante("beneficiario") = LstBeneficiario.Text
              LstJustificacion.ListIndex = i
              If Not IsNull(LstJustificacion.Text) Then rsComprobante("Justificacion") = LstJustificacion.Text
              LstNroCheque.ListIndex = i
              If LstNroCheque.Text <> "" Then rsComprobante("Nro_Cheque") = LstNroCheque.Text
              LstBanco.ListIndex = i
              If LstBanco.Text <> "" Then rsComprobante("banco") = LstBanco.Text
              LstLiteral.ListIndex = i
              If LstLiteral.Text <> "" Then rsComprobante("literal") = LstLiteral.Text
              LiteralCry = ""
              LstMonto.ListIndex = i
              If LstMonto.Text <> "" Then
                 rsComprobante("literal") = Literal(CStr(LstMonto.Text)) + "  BOLIVIANOS"
              End If
             LstTransf_Cheq.ListIndex = i
             If LstTransf_Cheq.Text <> "" Then
                If LstTransf_Cheq.Text = "T" Then
                    rsComprobante("Transf_Cheq") = "TRANSFERENCIA"
                End If
                If LstTransf_Cheq.Text = "C" Then
                    rsComprobante("Transf_Cheq") = "CHEQUE"
                End If
             End If
         rsComprobante.Update
   Next i
   'sino = MsgBox("Se imprimiran los comprobantes ...!", vbYesNo, "Mensaje de Advertencia")
   'If sino = vbYes Then
        'If OptCheques.Value = True Then'
            'FrmComprobante.Show
            '**crycomprobante.PrintOut
            'cryCmpte.ReadRecords
            MsgBox "Esta es la impresion"
            cryCmpte.Database.Verify
            cryCmpte.PrintOut
            'cr.PrintOut 'Sends the Report to the Printer
        'End If
        'If OptTransferencias.Value = True Then
        '    FrmComprobanteTrans.Show
        'End If
   'Else
   '     Exit Sub
   'End If
   sw = 0
End Sub

Private Sub CmdLimpiar_Click()
        LstComprobante.Clear
        LstOrganismo.Clear
        LstFecha.Clear
        LstMonto.Clear
        LstCambio.Clear
        LstBeneficiario.Clear
        LstJustificacion.Clear
        LstNroCheque.Clear
        LstBanco.Clear
        LstLiteral.Clear
        LstTransf_Cheq.Clear
End Sub

Private Sub CmdRestaurar_Click()
    Dim SqlQuery As String
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
    'SqlQuery = " SELECT Pagos.codigo_pago, fc_organismo_financiamiento.Org_descripcion, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cheque_o_trf "
     SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
               "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo" ',db, adOpenKeyset, adLockOptimistic
    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGComprobantes.DataSource = rsComprobante
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    'MsgBox LstChequesCodigo.ListCount
    MsgBox LstChequesCodigo.ListIndex
    LstChequesDatos.RemoveItem punto
    'MsgBox LstChequesCodigo.Index(0)
End Sub
Private Sub dtgComprobantes_Click()
 Dim bandera As Integer
    bandera = 0
    For i = 0 To LstComprobante.ListCount - 1
         LstComprobante.ListIndex = i
         If LstComprobante.Text = DtGComprobantes.Columns(0) Then
              bandera = 1
         End If
    Next i
    If bandera = 0 Then
        LstComprobante.AddItem DtGComprobantes.Columns(0)
        LstOrganismo.AddItem DtGComprobantes.Columns(1)
        LstFecha.AddItem DtGComprobantes.Columns(2)
        LstMonto.AddItem DtGComprobantes.Columns(3)
        LstCambio.AddItem DtGComprobantes.Columns(4)
        LstBeneficiario.AddItem DtGComprobantes.Columns(5)
        LstJustificacion.AddItem DtGComprobantes.Columns(6)
        LstNroCheque.AddItem DtGComprobantes.Columns(7)
        LstBanco.AddItem DtGComprobantes.Columns(8)
        LstLiteral.AddItem DtGComprobantes.Columns(9)
        If Not IsNull(DtGComprobantes.Columns(11)) Then
            LstTransf_Cheq.AddItem DtGComprobantes.Columns(11)
        End If
        
    End If

End Sub

Private Sub dtgComprobantes_DblClick()
' Dim bandera As Integer
'    bandera = 0
'    For i = 0 To LstComprobante.ListCount - 1
'         LstComprobante.ListIndex = i
'         If LstComprobante.Text = dtgComprobantes.Columns(0) Then
'              bandera = 1
'         End If
'    Next i
'    If bandera = 0 Then
'        LstComprobante.AddItem dtgComprobantes.Columns(0)
'        LstMonto.AddItem dtgComprobantes.Columns(1)
'        LstFecha.AddItem dtgComprobantes.Columns(2)
'        LstBeneficiario.AddItem dtgComprobantes.Columns(3)
'
'        LstCuenta.AddItem dtgComprobantes.Columns(4)
'    End If


End Sub

Private Sub dtgComprobantes_HeadClick(ByVal ColIndex As Integer)
    Set rsComprobante = New ADODB.Recordset
    CmdLimpiar_Click
    If rsComprobante.State = 1 Then rsComprobante.Close
    Select Case ColIndex
        Case 0
'            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
'            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by  Pagos.codigo_pago", db, adOpenKeyset, adLockOptimistic

            'SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo "
             SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.codigo_pago"
            rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
        Case 1
            'rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by pago_detalle.monto_Bolivianos", db, adOpenKeyset, adLockOptimistic
        Case 2
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
        Case 3
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by fc_beneficiario.denominacion_beneficiario", db, adOpenKeyset, adLockOptimistic
        Case 4
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by pago_detalle.cta_codigo", db, adOpenKeyset, adLockOptimistic
    End Select
    Set DtGComprobantes.DataSource = rsComprobante
End Sub

Private Sub Form_Load()
Dim cryCmpte As New CryComprobante
    Dim SqlQuery As String
    Set rsComprobante = New ADODB.Recordset
    SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
               "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by Pago_detalle.codigo_pago"
    'SqlQuery = " SELECT Pagos.codigo_pago, fc_organismo_financiamiento.Org_descripcion, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal " & _
    '           "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo" ',db, adOpenKeyset, adLockOptimistic
    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGComprobantes.DataSource = rsComprobante
    End If
	Call SeguridadSet(Me)
End Sub
Private Sub LstCheques_Click()
    MsgBox LstChequesCodigo.ListCount
    LstCheques.RemoveItem LstCheques.ListIndex
End Sub

Private Sub LstChequesCodigo_Click()
    punto = LstChequesCodigo.ListIndex
    'MsgBox LstChequesCodigo.ListCount
    'LstChequesDatos.RemoveItem punto
    
    'LstChequesCodigo.RemoveItem LstChequesCodigo.ListIndex
    
    'LstChequesDatos_Click
End Sub

Private Sub LstChequesCodigo_DblClick()
    LstChequesCodigo.RemoveItem LstChequesCodigo.ListIndex
End Sub

Private Sub LstChequesDatos_Click()
    LstChequesDatos.RemoveItem LstChequesDatos.ListIndex
End Sub

Public Sub Cheques_Impresos_lista()
        'Determinando comprobante de pagos en detalle como APROBADOS CHEQUES
        For i = 0 To LstNroCheque.ListCount - 1
          LstNroCheque.ListIndex = i
          LstComprobante.ListIndex = i
          NroCheque = LstNroCheque.Text
          
            Set rspago = New ADODB.Recordset
            If rspago.State = 1 Then rspago.Close
            rspago.Open "SELECT * from pagos where codigo_pago= '" & LstComprobante.Text & "'", db, adOpenKeyset, adLockOptimistic
            If rspago.RecordCount > 0 Then
                Set rsPagoDetalle = New ADODB.Recordset
                If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                rsPagoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & LstComprobante.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                     rsPagoDetalle("estado_aprobacion") = "A"
                     rsPagoDetalle.Update
                End If
                Set rsPagoDetalle = New ADODB.Recordset
                If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                rsPagoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & LstComprobante.Text & "' and estado_aprobacion<>'A'", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                    SumaMontosParciales = 0
                    While Not rsPagoDetalle.EOF
                     SumaMontosParciales = SumaMontosParciales + rsPagoDetalle("monto_bolivianos")
                     rsPagoDetalle.MoveNext
                    Wend
                    If rspago("liquido_pagar") = SumaMontosParciales And SumaMontosParciales <> 0 Then
                     rspago("estado_aprobacion") = "A"
                     rspago.Update
                    End If
                End If
        
                If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
           End If
     Next i
End Sub


Private Sub LstBanco_DblClick()
    LstBanco.RemoveItem punto
    LstLiteral_DblClick
End Sub

Private Sub LstBeneficiario_DblClick()
    LstBeneficiario.RemoveItem punto
    LstJustificacion_DblClick
End Sub

Private Sub LstCambio_DblClick()
    LstCambio.RemoveItem punto
    LstBeneficiario_DblClick
End Sub

Private Sub LstComprobante_DblClick()
    punto = LstComprobante.ListIndex
    LstComprobante.RemoveItem punto
    LstOrganismo_DblClick
End Sub
Private Sub LstFecha_DblClick()
    LstFecha.RemoveItem punto
    LstMonto_DblClick
End Sub

Public Sub Cheques_Impresos_rango()
       Set rsCheque = New ADODB.Recordset
       If rsCheque.State = 1 Then rsCheque.Close
       rsCheque.Open "SELECT * from ts_cheque", db, adOpenKeyset, adLockOptimistic
       If rsCheque.RecordCount > 0 Then
       While Not rsCheque.EOF
            Set rspago = New ADODB.Recordset
            If rspago.State = 1 Then rspago.Close
            rspago.Open "SELECT * from pagos where codigo_pago='" & rsCheque("numero_comprobante") & "'", db, adOpenKeyset, adLockOptimistic
            If rspago.RecordCount > 0 Then
                Set rsPagoDetalle = New ADODB.Recordset
                If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                rsPagoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & rsCheque("numero_comprobante") & "' and estado_aprobacion<>'A'", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                     rsPagoDetalle("estado_aprobacion") = "A"
                     rsPagoDetalle.Update
                End If
                
                Set rsPagoDetalle = New ADODB.Recordset
                If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
                rsPagoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & rsCheque("numero_comprobante") & "'", db, adOpenKeyset, adLockOptimistic
                If rsPagoDetalle.RecordCount > 0 Then
                SumaMontosParciales = 0
                    While Not rsPagoDetalle.EOF
                     SumaMontosParciales = SumaMontosParciales + rsPagoDetalle("monto_bolivianos")
                     rsPagoDetalle.MoveNext
                    Wend
                    If rspago("liquido_pagar") = SumaMontosParciales And SumaMontosParciales <> 0 Then
                     rspago("estado_aprobacion") = "A"
                     rspago.Update
                    End If
                End If
                If rsPagoDetalle.State = 1 Then rsPagoDetalle.Close
           End If
           rsCheque.MoveNext
         Wend
        End If
End Sub

Public Sub Restaurar_numeracion_cheque()
   Set rsCheque = New ADODB.Recordset
   rsCheque.Open "SELECT * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
   If rsCheque.RecordCount > 0 Then
     While Not rsCheque.EOF
         Select Case rsCheque("cta_codigo")
            Case "4.41.1.1.1.402.208.11-2"
                  NumeroCuenta = "cta_1"
            Case "4.41.1.1.1.402.208.12-1"
                  NumeroCuenta = "cta_2"
            Case "4.41.1.1.1.402.208.14-0"
                  NumeroCuenta = "cta_3"
            Case "4.41.1.1.1.402.208.16-8"
                  NumeroCuenta = "cta_4"
            Case "4.41.1.1.1.402.208.18-6"
                  NumeroCuenta = "cta_5"
            Case "4.41.1.1.1.402.254.01-7"
                  NumeroCuenta = "cta_6"
            Case "4.41.1.1.1.402.254.02-6"
                  NumeroCuenta = "cta_7"
            Case "1-297792"
                  NumeroCuenta = "cta_8"
            Case "1-297809"
                  NumeroCuenta = "cta_9"
            Case "1-297841"
                  NumeroCuenta = "cta_10"
            Case "1-297867"
                  NumeroCuenta = "cta_11"
            Case "1-297875"
                  NumeroCuenta = "cta_12"
            Case "1-297883"
                  NumeroCuenta = "cta_13"
            Case "1-297891"
                  NumeroCuenta = "cta_14"
            Case "1-297916"
                  NumeroCuenta = "cta_15"
            Case "1-297924"
                  NumeroCuenta = "cta_16"
            Case "1-297932"
                  NumeroCuenta = "cta_17"
            Case "1-297940"
                  NumeroCuenta = "cta_18"
            Case "1-297958"
                  NumeroCuenta = "cta_19"
            Case "1-301973"
                  NumeroCuenta = "cta_20"
            Case "1-301999"
                  NumeroCuenta = "cta_21"
            Case "1-302731"
                  NumeroCuenta = "cta_22"
            Case "1-303515"
                  NumeroCuenta = "cta_23"
            Case "1-306379"
                  NumeroCuenta = "cta_24"
            Case "1-302731"
                  NumeroCuenta = "cta_25"
         End Select
         
         If rsCorrel.State = 1 Then rsCorrel.Close
         Set rsCorrel = New ADODB.Recordset
         rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= '" & NumeroCuenta & "' ", db, adOpenKeyset, adLockOptimistic
         If rsCorrel.RecordCount > 0 Then
            rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") - 1 'LstComprobante.ListCount
            rsCorrel.Update
         Else
            rsCorrel("numero_correlativo") = 0
            rsCorrel.Update
         End If
         rsCheque.MoveNext
     Wend
    End If
    'CmdLimpiar_Click
End Sub

Public Sub NrosCheque_Compte()
Dim NumeroCheque As String

If rsCheque.State = 1 Then rsCheque.Close
Set rsCheque = New ADODB.Recordset
rsCheque.Open "select * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
If rsCheque.RecordCount > 0 Then
        While Not rsCheque.EOF
            Set rsPagoDet = New ADODB.Recordset
            rsPagoDet.Open "select * from pago_detalle where codigo_pago='" & rsCheque("numero_comprobante") & "' and estado_aprobacion='N'", db, adOpenKeyset, adLockOptimistic
'            If rsPagoDet.RecordCount >= 0 Then
                'rsPagoDet("cta_codigo") = rsCheque("cta_codigo")
                'Determinar el numero con ceros
                
                Select Case Len(rsCheque("numero_cheque"))
                    Case 1
                        NumeroCheque = "0000" + rsCheque("numero_cheque")
                    Case 2
                        NumeroCheque = "000" + rsCheque("numero_cheque")
                    Case 3
                        NumeroCheque = "00" + rsCheque("numero_cheque")
                    Case 4
                        NumeroCheque = "0" + rsCheque("numero_cheque")
                    Case 5
                        NumeroCheque = rsCheque("numero_cheque")
                End Select
                
                rsPagoDet("numero_cheque_trf") = NumeroCheque
                rsPagoDet.Update
'            End If
            rsCheque.MoveNext
        Wend
End If
End Sub

Private Sub LstJustificacion_DblClick()
    LstJustificacion.RemoveItem punto
    LstNroCheque_DblClick
End Sub

Private Sub LstLiteral_DblClick()
    LstLiteral.RemoveItem punto
    LstTransf_Cheq_DblClick
End Sub

Private Sub LstMonto_DblClick()
    LstMonto.RemoveItem punto
    LstCambio_DblClick
End Sub


Private Sub LstNroCheque_DblClick()
    LstNroCheque.RemoveItem punto
    LstBanco_DblClick
End Sub

Private Sub LstOrganismo_DblClick()
    LstOrganismo.RemoveItem punto
    LstFecha_DblClick
End Sub

Private Sub LstTransf_Cheq_DblClick()
    LstTransf_Cheq.RemoveItem punto
End Sub

Private Sub OptCheques_Click()
    Dim SqlQuery As String
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
    'SqlQuery = " SELECT Pagos.codigo_pago,  fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal,pago_detalle.cheque_o_trf "
    SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
               "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where cheque_o_trf= 'C'"
    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGComprobantes.DataSource = rsComprobante
    End If
End Sub

Private Sub OptTransferencias_Click()
    Dim SqlQuery As String
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
    'SqlQuery = " SELECT Pagos.codigo_pago,  fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal,pago_detalle.cheque_o_trf "
    SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
               "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where cheque_o_trf= 'T'"
    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGComprobantes.DataSource = rsComprobante
    End If
    
    
'    Dim SqlQuery As String
'    Set rsComprobante = New adodb.Recordset
'    If rsComprobante.State = 1 Then rsComprobante.Close
'    SqlQuery = " SELECT Pagos.codigo_pago,  fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal " & _
'               "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where cheque_o_trf= 'T' "
'    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
'    If rsComprobante.RecordCount > 0 Then
'        Set DtGComprobantes.DataSource = rsComprobante
'    End If

End Sub


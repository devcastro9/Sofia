VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmImprimeComprobanteNuevo 
   Caption         =   "Impresión de Comprobantes"
   ClientHeight    =   8175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12615
   Icon            =   "FrmImprimeComprobanteNuevo.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   12615
   WindowState     =   2  'Maximized
   Begin VB.Frame FraOpciones 
      BackColor       =   &H00C0C0C0&
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
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   12390
      Begin VB.CommandButton CmdImprimeTrf 
         Caption         =   "Transfer."
         Height          =   720
         Left            =   9120
         Picture         =   "FrmImprimeComprobanteNuevo.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   480
         Picture         =   "FrmImprimeComprobanteNuevo.frx":7FD4
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Refresca"
         Height          =   720
         Left            =   3720
         Picture         =   "FrmImprimeComprobanteNuevo.frx":889E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   720
         Left            =   11280
         Picture         =   "FrmImprimeComprobanteNuevo.frx":8AA8
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Cmpbte."
         Height          =   720
         Left            =   10200
         Picture         =   "FrmImprimeComprobanteNuevo.frx":8CB2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   720
         Left            =   1560
         Picture         =   "FrmImprimeComprobanteNuevo.frx":A434
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdFiltro 
         Caption         =   "Filtrar"
         Height          =   720
         Left            =   2640
         Picture         =   "FrmImprimeComprobanteNuevo.frx":ACFE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   785
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      Picture         =   "FrmImprimeComprobanteNuevo.frx":B9C8
      ScaleHeight     =   675
      ScaleWidth      =   12555
      TabIndex        =   14
      Top             =   0
      Width           =   12615
      Begin VB.Label LblUni_descripcion_larga 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3480
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   5160
      End
      Begin VB.Label lblUni_codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1200
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTES DE PAGO POR COMPRAS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   405
         Left            =   5250
         TabIndex        =   16
         Top             =   120
         Width           =   6720
      End
      Begin VB.Label LblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "LblUsuario"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1200
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin Crystal.CrystalReport CryCompr 
      Left            =   6000
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSDataGridLib.DataGrid DtGCompr 
      Height          =   6480
      Left            =   6885
      TabIndex        =   12
      Top             =   1650
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   11430
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
      Caption         =   "COMPROBANTES A IMPRIMIR"
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
   Begin VB.CommandButton Seleccionar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccionar"
      Height          =   750
      Left            =   5750
      MaskColor       =   &H80000016&
      Picture         =   "FrmImprimeComprobanteNuevo.frx":D56E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2145
      Width           =   1020
   End
   Begin VB.CommandButton Retornar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar"
      Height          =   750
      Left            =   5750
      Picture         =   "FrmImprimeComprobanteNuevo.frx":D6F8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3315
      Width           =   1020
   End
   Begin VB.Frame FraBusca 
      Height          =   2100
      Left            =   1875
      TabIndex        =   0
      Top             =   4530
      Visible         =   0   'False
      Width           =   2040
      Begin VB.CommandButton CmdSalirBusca 
         Caption         =   "Salir"
         Height          =   375
         Left            =   225
         TabIndex        =   13
         Top             =   1635
         Width           =   1515
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   225
         TabIndex        =   4
         Top             =   1245
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   240
         TabIndex        =   3
         Top             =   780
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2047
         TabIndex        =   2
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   1
         Top             =   915
         Width           =   1515
      End
      Begin VB.Label Label21 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   450
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   2310
         TabIndex        =   6
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label20 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   5
         Top             =   645
         Width           =   795
      End
   End
   Begin MSDataGridLib.DataGrid DtGComprobantes 
      Height          =   6450
      Left            =   120
      TabIndex        =   8
      Top             =   1665
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   11377
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
      Caption         =   "COMPROBANTES A SELECCIONAR"
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
   Begin Crystal.CrystalReport CryCh 
      Left            =   4845
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryTrf 
      Left            =   6000
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   75
      Left            =   10875
      TabIndex        =   9
      Top             =   1725
      Width           =   45
   End
End
Attribute VB_Name = "FrmImprimeComprobanteNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  ADFIN-2002
' Módulo:                   Impresion de Comprobantes
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmImprimirComprobante.frm
' Descipción :              Luego de impreso los cheques y/o transferencias
'                           se requiere la impresión de sus comprobantes.
' Formularios relacionados: Main.frm (Padre)
'                           CryComprobante
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creación         10/Ene/ 2001
' Fecha última modificación 15/Mar/ 2001
' Versión:                  2.0
'========================================================================================

Public rsComprobante As New ADODB.Recordset
Dim rsCheque As New ADODB.Recordset
Dim rsCorrel As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim rsCom As New ADODB.Recordset

Dim punto As Variant
Dim NumeroCuenta As String

Private Sub CmdBuscar_Click()
 If TxtCmpte.Text = "" Then
    MsgBox "Necesita números de comprobante"
    Exit Sub
 Else
    condicion = "pago_detalle.codigo_pago=" + "'" + TxtCmpte.Text + "'"
 End If
 'SqlQuery = " SELECT DISTINCT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf,pago_detalle.org_codigo, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
 '"FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE (" & condicion & ") " & _
 '"order by Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf,pago_detalle.org_codigo "
                   
 SqlQuery = "SELECT DISTINCT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, " & _
    "pago_detalle.cheque_o_trf, pago_detalle.org_codigo, Pagos.codigo_unidad, Pagos.codigo_solicitud, pago_detalle.cta_codigo_destino, pago_detalle.numero_cheque_trf_destino, fc.cta_descripcion_larga as cta_descripcion_destino, fb.Bco_descripcion_larga as Bco_descripcion_destino" & _
    "FROM ((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario)" & _
    "INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo left outer JOIN fc_cuenta_bancaria Fc ON pago_detalle.cta_codigo_destino = fc.Cta_codigo left outer JOIN fc_bancos Fb ON fb.Bco_codigo = fc.Bco_codigo " & _
    "WHERE (" & condicion & ") order by Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.cta_codigo, pago_detalle.org_codigo"
                   
 If rsComprobante.State Then rsComprobante.Close
 rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
 If rsComprobante.RecordCount > 0 Then
    Set DtGComprobantes.DataSource = rsComprobante
 End If
 FraBusca.Visible = False
End Sub

Private Sub CmdBusqueda_Click()
    FraBusca.Visible = True
End Sub

Private Sub cmdFiltro_Click()
Dim SqlQuery As String
Dim Resp As String

    Resp = InputBox("Introducir Organismo o Cuenta Bancaria")
    If Resp <> "" Then
      Set rsCheque = New ADODB.Recordset
      If rsCheque.State = 1 Then rsCheque.Close
'      rsCheque.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf,  fc_bancos.Bco_descripcion_larga " & _
'      "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.cta_codigo= '" & Resp & "' and pago_detalle.estado_aprobacion <> 'A' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic

      'SqlQuery = "SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal,pago_detalle.cta_codigo "
      SqlQuery = " SELECT DISTINCT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf, Pagos.codigo_unidad, Pagos.codigo_solicitud  " & _
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
        SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, Pagos.codigo_unidad, Pagos.codigo_solicitud, pago_detalle.cta_codigo " & _
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
                      rsCmpte("cta_codigo") = rsComprobante("cta_codigo")
                      rsCmpte("codigo_unidad") = rsComprobante("codigo_unidad")
                      rsCmpte("codigo_solicitud") = rsComprobante("codigo_solicitud")
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

Private Sub CmdImprimeTrf_Click()
    Dim rsComp As New ADODB.Recordset
    If rsComp.State = 1 Then rsComp.Close
    rsComp.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
    If rsComp.RecordCount > 0 Then
           CryTrf.ReportFileName = App.Path & "\FormsTesoreria\Impresion Comprobantes de Pago\Rpt_Comprobantes_trf.rpt "
           IResult = CryTrf.PrintReport
           If IResult <> 0 Then
              MsgBox CryTrf.LastErrorNumber & " : " & CryTrf.LastErrorString, vbCritical + vbOKOnly, "Error..."
           End If
    Else
           MsgBox "No existen registros para imprimir", vbCritical + vbDefaultButton1, "Validación de Datos"
    End If
End Sub

Private Sub Cmdimprimir_Click()
    Dim rsComp As New ADODB.Recordset
    If rsComp.State = 1 Then rsComp.Close
    rsComp.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
    If rsComp.RecordCount > 0 Then
           CryCompr.ReportFileName = App.Path & "\FormsTesoreria\Impresion Comprobantes de Pago\Rpt_Comprobantes.rpt"
           IResult = CryCompr.PrintReport
           If IResult <> 0 Then
              MsgBox CryCompr.LastErrorNumber & " : " & CryCompr.LastErrorString, vbCritical + vbOKOnly, "Error..."
           End If
    Else
           MsgBox "No existen registros para imprimir", vbCritical + vbDefaultButton1, "Validación de Datos"
    End If
 End Sub

Private Sub CmdLimpiar_Click()
Dim rsComp As New ADODB.Recordset
    db.Execute "DELETE FROM to_comprobantes"
    If rsComp.State = 1 Then rsComp.Close
    rsComp.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
    If rsComp.RecordCount > 0 Then
        Set DtGCompr.DataSource = rsComp
    Else
        Set DtGCompr.DataSource = rsNada
    End If
    CmdRestaurar_Click
End Sub

Private Sub CmdRestaurar_Click()
    Dim SqlQuery As String
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
     SqlQuery = " SELECT DISTINCT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf,pago_detalle.org_codigo, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
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

    MsgBox LstChequesCodigo.ListIndex
    LstChequesDatos.RemoveItem punto
End Sub

Private Sub CmdSalirBusca_Click()
    FraBusca.Visible = False
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
             SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.codigo_pago"
            rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
        Case 1
            'rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by pago_detalle.monto_Bolivianos", db, adOpenKeyset, adLockOptimistic
        Case 2
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
        Case 3
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by fc_beneficiario.denominacion_beneficiario", db, adOpenKeyset, adLockOptimistic
        Case 4
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by pago_detalle.cta_codigo", db, adOpenKeyset, adLockOptimistic
    End Select
    Set DtGComprobantes.DataSource = rsComprobante
End Sub

Private Sub Form_Load()

    Dim SqlQuery As String
    Set rsComprobante = New ADODB.Recordset
    SqlQuery = " SELECT DISTINCT dbo.pagos.codigo_pago, dbo.fc_cuenta_bancaria.Cta_descripcion_larga, dbo.pago_detalle.fecha_pago, dbo.pago_detalle.monto_Bolivianos, dbo.pago_detalle.tipo_cambio, dbo.FC_BENEFICIARIO.denominacion_beneficiario, dbo.pagos.justificacion, dbo.pago_detalle.numero_cheque_trf," & _
                " dbo.fc_bancos.Bco_descripcion_larga, dbo.pago_detalle.literal, dbo.pago_detalle.cta_codigo, dbo.pago_detalle.cheque_o_trf, dbo.pago_detalle.org_codigo, dbo.pagos.codigo_unidad, dbo.pagos.codigo_solicitud, dbo.pago_detalle.cta_codigo_destino, dbo.pago_detalle.numero_cheque_trf_destino, Fc.Cta_descripcion_larga AS cta_descripcion_destino, Fb.Bco_descripcion_larga AS Bco_descripcion_destino " & _
                " FROM dbo.fc_bancos INNER JOIN dbo.pagos INNER JOIN dbo.pago_detalle ON dbo.pagos.ges_gestion = dbo.pago_detalle.ges_gestion AND dbo.pagos.org_codigo = dbo.pago_detalle.org_codigo AND dbo.pagos.codigo_pago = dbo.pago_detalle.codigo_pago INNER JOIN dbo.FC_BENEFICIARIO ON" & _
                " dbo.pago_detalle.codigo_beneficiario COLLATE Modern_Spanish_CI_AS = dbo.FC_BENEFICIARIO.codigo_beneficiario INNER JOIN dbo.fc_cuenta_bancaria ON dbo.pago_detalle.cta_codigo = dbo.fc_cuenta_bancaria.Cta_codigo ON dbo.fc_bancos.Bco_codigo = dbo.fc_cuenta_bancaria.Bco_codigo LEFT OUTER JOIN" & _
                " dbo.fc_cuenta_bancaria AS Fc ON dbo.pago_detalle.cta_codigo_destino = Fc.Cta_codigo LEFT OUTER JOIN dbo.fc_bancos AS Fb ON Fc.Bco_codigo = Fb.Bco_codigo "
    
'    SqlQuery = "SELECT DISTINCT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, " & _
'        "pago_detalle.cta_codigo, pago_detalle.cheque_o_trf, pago_detalle.org_codigo, Pagos.codigo_unidad, Pagos.codigo_solicitud, pago_detalle.cta_codigo_destino, pago_detalle.numero_cheque_trf_destino, fc.cta_descripcion_larga as cta_descripcion_destino, fb.Bco_descripcion_larga as Bco_descripcion_destino " & _
'        "FROM ((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) " & _
'        "INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo left outer JOIN fc_cuenta_bancaria Fc ON pago_detalle.cta_codigo_destino = fc.Cta_codigo left outer JOIN fc_bancos Fb ON fb.Bco_codigo = fc.Bco_codigo "

    'order by Pago_detalle.codigo_pago
    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        rsComprobante.Sort = "codigo_pago"
        Set DtGComprobantes.DataSource = rsComprobante
    End If
    'Lista los comprobantes para imprimir
    Refrescar
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
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & LstComprobante.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                     rsPAgoDetalle("estado_aprobacion") = "A"
                     rsPAgoDetalle.Update
                End If
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & LstComprobante.Text & "' and estado_aprobacion<>'A'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                    SumaMontosParciales = 0
                    While Not rsPAgoDetalle.EOF
                     SumaMontosParciales = SumaMontosParciales + rsPAgoDetalle("monto_bolivianos")
                     rsPAgoDetalle.MoveNext
                    Wend
                    If rspago("liquido_pagar") = SumaMontosParciales And SumaMontosParciales <> 0 Then
                     rspago("estado_aprobacion") = "A"
                     rspago.Update
                    End If
                End If
        
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
           End If
     Next i
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
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & rsCheque("numero_comprobante") & "' and estado_aprobacion<>'A'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                     rsPAgoDetalle("estado_aprobacion") = "A"
                     rsPAgoDetalle.Update
                End If
                
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & rsCheque("numero_comprobante") & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                SumaMontosParciales = 0
                    While Not rsPAgoDetalle.EOF
                     SumaMontosParciales = SumaMontosParciales + rsPAgoDetalle("monto_bolivianos")
                     rsPAgoDetalle.MoveNext
                    Wend
                    If rspago("liquido_pagar") = SumaMontosParciales And SumaMontosParciales <> 0 Then
                     rspago("estado_aprobacion") = "A"
                     rspago.Update
                    End If
                End If
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
           End If
           rsCheque.MoveNext
         Wend
        End If
End Sub

Private Sub OptCheques_Click()
    Dim SqlQuery As String
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
    'SqlQuery = " SELECT Pagos.codigo_pago,  fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal,pago_detalle.cheque_o_trf "
    SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
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
    SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf, Pagos.codigo_unidad, Pagos.codigo_solicitud " & _
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

Private Sub Retornar_Click()
   db.Execute "DELETE FROM to_Comprobantes where Nro_Cmpte='" & DtGCompr.Columns(0) & "' and organismo= '" & DtGCompr.Columns(1) & "' "
   Refrescar
End Sub

Private Sub Seleccionar_Click()
'Insertar nuevos Comprobantes
    Set rsCom = New ADODB.Recordset
    If rsCom.State = 1 Then rsCom.Close
    rsCom.Open "select * from to_comprobantes where Nro_Cmpte=" & rsComprobante("Codigo_Pago") & " and Organismo='" & rsComprobante("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
    
    If rsCom.RecordCount = 0 Then
       rsCom.AddNew
       rsCom("Nro_Cmpte") = rsComprobante("Codigo_Pago")
       rsCom("org_codigo") = rsComprobante("ORG_codigo")
       rsCom("Organismo") = rsComprobante("cta_descripcion_larga")
       rsCom("Fecha_Pago") = Format(rsComprobante("Fecha_Pago"), "dd/mm/yyyy")
       rsCom("Monto") = rsComprobante("Monto_Bolivianos")
       rsCom("Cambio") = rsComprobante("Tipo_Cambio")
       rsCom("Beneficiario") = rsComprobante("denominacion_Beneficiario")
       rsCom("Justificacion") = rsComprobante("Justificacion")
       rsCom("Nro_Cheque") = Val(rsComprobante("Numero_cheque_trf"))
       If rsComprobante("cheque_o_trf") = "D" Then
          rsCom("Transf_Cheq") = "DEPOSITO BANCO"
       End If
       If rsComprobante("cheque_o_trf") = "C" Then
          rsCom("Transf_Cheq") = "CHEQUE/EFECTIVO"
       End If
       If rsComprobante("cheque_o_trf") = "T" Then
          rsCom("Transf_Cheq") = "TRANSFERENCIA"
       End If
       rsCom("Banco") = rsComprobante("Bco_Descripcion_Larga")
       rsCom("Literal") = Literal(rsComprobante("Monto_Bolivianos")) & " " & "BOLIVIANOS"
       rsCom("cta_codigo") = rsComprobante("cta_codigo")
       rsCom("cta_codigo_destino") = rsComprobante("cta_codigo_destino")
       rsCom("numero_cheque_trf_destino") = rsComprobante("numero_cheque_trf_destino")
       rsCom("Cta_descripcion_destino") = rsComprobante("cta_descripcion_destino")
       rsCom("Banco_destino") = rsComprobante("Bco_descripcion_destino")
       
'pago_detalle.cta_codigo_destino, pago_detalle.numero_cheque_trf_destino, fc.cta_descripcion_larga as cta_descripcion_destino, fb.Bco_descripcion_larga as Bco_descripcion_destino

       rsCom("codigo_unidad") = rsComprobante("codigo_unidad")
       rsCom("codigo_solicitud") = rsComprobante("codigo_solicitud")
       rsCom.Update
'       db.Execute "insert into to_comprobantes (Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Justificacion, Nro_cheque, Banco, Transf_Cheq, Literal)" & _
'                  "values ('" & DtGComprobantes.Columns(0) & "','" & DtGComprobantes.Columns(1) & "','" & DtGComprobantes.Columns(2) & "', " & DtGComprobantes.Columns(3) & ",'" & DtGComprobantes.Columns(4) & "'," & DtGComprobantes.Columns(5) & ",'" & DtGComprobantes.Columns(6) & "','','" & DtGComprobantes.Columns(7) & "','" & DtGComprobantes.Columns(8) & "', '" & DtGComprobantes.Columns(9) & "' + 'BOLIVIANOS' ) "
    End If
    Refrescar
    Retornar.Enabled = True
End Sub

Public Sub Refrescar()
    Set rsCom = New ADODB.Recordset
    rsCom.Open "select * from to_comprobantes ", db, adOpenKeyset, adLockOptimistic
    If rsCom.RecordCount > 0 Then
        Set DtGCompr.DataSource = rsCom
    Else
        Set DtGCompr.DataSource = rsNada
        Retornar.Enabled = False
    End If
    
End Sub

Private Sub TxtCmpte_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

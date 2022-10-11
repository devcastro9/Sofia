VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_Imprime_Factura 
   BackColor       =   &H00404040&
   Caption         =   "Impresión de Comprobantes"
   ClientHeight    =   8175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   12615
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Retornar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Desmarcar"
      Height          =   720
      Left            =   6960
      Picture         =   "Frm_Imprime_Factura.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "Frm_Imprime_Factura.frx":020A
      ScaleHeight     =   960
      ScaleWidth      =   14940
      TabIndex        =   13
      Top             =   120
      Width           =   15000
      Begin VB.CommandButton CmdBusqueda 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   1320
         Picture         =   "Frm_Imprime_Factura.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdRestaurar 
         BackColor       =   &H00808000&
         Caption         =   "Refrescar"
         Height          =   720
         Left            =   2160
         Picture         =   "Frm_Imprime_Factura.frx":6C7F4
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdFiltro 
         BackColor       =   &H00808000&
         Caption         =   "Filtrar"
         Height          =   720
         Left            =   3000
         Picture         =   "Frm_Imprime_Factura.frx":6CE18
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdLimpiar 
         BackColor       =   &H00808000&
         Caption         =   "Limpiar"
         Height          =   720
         Left            =   3840
         Picture         =   "Frm_Imprime_Factura.frx":6D3F8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   13680
         Picture         =   "Frm_Imprime_Factura.frx":6E0C2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   7560
         Picture         =   "Frm_Imprime_Factura.frx":6E2CC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdImprimeTrf 
         BackColor       =   &H00808000&
         Caption         =   "Kardex"
         Height          =   720
         Left            =   9720
         Picture         =   "Frm_Imprime_Factura.frx":6E70E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprime Recibo / Kardex"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdFoto 
         BackColor       =   &H00808000&
         Caption         =   "&QR"
         Height          =   720
         Left            =   6840
         Picture         =   "Frm_Imprime_Factura.frx":6FE90
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Carga Imagen QR"
         Top             =   120
         Visible         =   0   'False
         Width           =   740
      End
      Begin VB.CommandButton CmdImprimir 
         BackColor       =   &H00C0C000&
         Caption         =   "Factura"
         Height          =   720
         Left            =   10800
         Picture         =   "Frm_Imprime_Factura.frx":70892
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprime Factura"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRESION DE FACTURAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   5220
         TabIndex        =   23
         Top             =   300
         Width           =   4065
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
      Left            =   8085
      TabIndex        =   10
      Top             =   1170
      Width           =   6870
      _ExtentX        =   12118
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
         TabIndex        =   11
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
      Top             =   1185
      Width           =   6705
      _ExtentX        =   11827
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
   Begin VB.CommandButton Seleccionar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Seleccionar"
      Height          =   720
      Left            =   6960
      Picture         =   "Frm_Imprime_Factura.frx":72014
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Aprueba Registro"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7200
      Picture         =   "Frm_Imprime_Factura.frx":7221E
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4560
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.Label lblUni_codigo 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2400
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "LblUsuario"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   240
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
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
Attribute VB_Name = "Frm_Imprime_Factura"
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
     rsCmpte.Open "SELECT * FROM fo_Comprobantes", db, adOpenKeyset, adLockOptimistic
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
    rsComp.Open "SELECT * FROM fo_Comprobantes", db, adOpenKeyset, adLockOptimistic
    If rsComp.RecordCount > 0 Then
           CryTrf.ReportFileName = App.Path & "\FormsTesoreria\Impresion Comprobantes de Pago\Rpt_Comprobantes_trf.rpt "
           iResult = CryTrf.PrintReport
           If iResult <> 0 Then
              MsgBox CryTrf.LastErrorNumber & " : " & CryTrf.LastErrorString, vbCritical + vbOKOnly, "Error..."
           End If
    Else
           MsgBox "No existen registros para imprimir", vbCritical + vbDefaultButton1, "Validación de Datos"
    End If
End Sub

Private Sub Cmdimprimir_Click()
    Dim rsComp As New ADODB.Recordset
    If rsComp.State = 1 Then rsComp.Close
    rsComp.Open "SELECT * FROM fo_Comprobantes", db, adOpenKeyset, adLockOptimistic
    If rsComp.RecordCount > 0 Then
           CryCompr.ReportFileName = App.Path & "\FormsTesoreria\Impresion Comprobantes de Pago\Rpt_Comprobantes.rpt"
           iResult = CryCompr.PrintReport
           If iResult <> 0 Then
              MsgBox CryCompr.LastErrorNumber & " : " & CryCompr.LastErrorString, vbCritical + vbOKOnly, "Error..."
           End If
    Else
           MsgBox "No existen registros para imprimir", vbCritical + vbDefaultButton1, "Validación de Datos"
    End If
 End Sub

Private Sub CmdLimpiar_Click()
Dim rsComp As New ADODB.Recordset
    db.Execute "DELETE FROM fo_Comprobantes"
    If rsComp.State = 1 Then rsComp.Close
    rsComp.Open "SELECT * FROM fo_Comprobantes", db, adOpenKeyset, adLockOptimistic
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

Private Sub cmdSalir_Click()
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
'    SqlQuery = " SELECT DISTINCT dbo.pagos.codigo_pago, dbo.fc_cuenta_bancaria.Cta_descripcion_larga, dbo.pago_detalle.fecha_pago, dbo.pago_detalle.monto_Bolivianos, dbo.pago_detalle.tipo_cambio, dbo.FC_BENEFICIARIO.denominacion_beneficiario, dbo.pagos.justificacion, dbo.pago_detalle.numero_cheque_trf," & _
'                " dbo.fc_bancos.Bco_descripcion_larga, dbo.pago_detalle.literal, dbo.pago_detalle.cta_codigo, dbo.pago_detalle.cheque_o_trf, dbo.pago_detalle.org_codigo, dbo.pagos.codigo_unidad, dbo.pagos.codigo_solicitud, dbo.pago_detalle.cta_codigo_destino, dbo.pago_detalle.numero_cheque_trf_destino, Fc.Cta_descripcion_larga AS cta_descripcion_destino, Fb.Bco_descripcion_larga AS Bco_descripcion_destino " & _
'                " FROM dbo.fc_bancos INNER JOIN dbo.pagos INNER JOIN dbo.pago_detalle ON dbo.pagos.ges_gestion = dbo.pago_detalle.ges_gestion AND dbo.pagos.org_codigo = dbo.pago_detalle.org_codigo AND dbo.pagos.codigo_pago = dbo.pago_detalle.codigo_pago INNER JOIN dbo.FC_BENEFICIARIO ON" & _
'                " dbo.pago_detalle.codigo_beneficiario COLLATE Modern_Spanish_CI_AS = dbo.FC_BENEFICIARIO.codigo_beneficiario INNER JOIN dbo.fc_cuenta_bancaria ON dbo.pago_detalle.cta_codigo = dbo.fc_cuenta_bancaria.Cta_codigo ON dbo.fc_bancos.Bco_codigo = dbo.fc_cuenta_bancaria.Bco_codigo LEFT OUTER JOIN" & _
'                " dbo.fc_cuenta_bancaria AS Fc ON dbo.pago_detalle.cta_codigo_destino = Fc.Cta_codigo LEFT OUTER JOIN dbo.fc_bancos AS Fb ON Fc.Bco_codigo = Fb.Bco_codigo "
    
    SqlQuery = "select * from av_rep_R101"
    'order by Pago_detalle.codigo_pago
    rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        rsComprobante.Sort = "edif_codigo"
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
   db.Execute "DELETE FROM fo_Comprobantes where Nro_Cmpte='" & DtGCompr.Columns(0) & "' and organismo= '" & DtGCompr.Columns(1) & "' "
   Refrescar
End Sub

Private Sub Seleccionar_Click()
'Insertar nuevos Comprobantes
    Set rsCom = New ADODB.Recordset
    If rsCom.State = 1 Then rsCom.Close
    rsCom.Open "select * from fo_comprobantes where campo =" & rsComprobante("cobranza_codigo") & " ", db, adOpenKeyset, adLockOptimistic
    If rsCom.RecordCount = 0 Then
       rsCom.AddNew
       rsCom("Nro_Cmpte") = rsComprobante("cobranza_codigo")
       rsCom("org_codigo") = rsComprobante("ges_gestion")
       rsCom("Organismo") = rsComprobante("venta_codigo")
       rsCom("Monto") = rsComprobante("unidad_codigo")
       rsCom("Cambio") = rsComprobante("solicitud_codigo")
       rsCom("Beneficiario") = rsComprobante("edif_codigo")
       rsCom("Justificacion") = rsComprobante("unidad_codigo_ant")
       rsCom("Fecha_Pago") = Format(rsComprobante("venta_fecha"), "dd/mm/yyyy")
       rsCom("Nro_Cheque") = rsComprobante("venta_tipo")
       
       rsCom("Banco") = rsComprobante("beneficiario_codigo")
       rsCom("Banco") = rsComprobante("beneficiario_codigo_resp")
       rsCom("Banco") = rsComprobante("venta_descripcion")
       rsCom("Banco") = rsComprobante("proceso_codigo")
       rsCom("Banco") = rsComprobante("subproceso_codigo")
       rsCom("Banco") = rsComprobante("etapa_codigo")
       rsCom("Banco") = rsComprobante("clasif_codigo")
       rsCom("Banco") = rsComprobante("doc_codigo")
       rsCom("Banco") = rsComprobante("doc_numero")
       
       rsCom("Banco") = rsComprobante("poa_codigo")
       rsCom("Banco") = rsComprobante("beneficiario_denominacion")
       rsCom("Banco") = rsComprobante("beneficiario_denominacion_resp")
       rsCom("Banco") = rsComprobante("venta_det_cantidad")
       rsCom("Banco") = rsComprobante("venta_precio_unitario_bs")
       rsCom("Banco") = rsComprobante("venta_descuento_bs")
       rsCom("Banco") = rsComprobante("venta_precio_total_bs")
       rsCom("Banco") = rsComprobante("venta_precio_unitario_dol")
       rsCom("Banco") = rsComprobante("venta_descuento_dol")
       rsCom("Banco") = rsComprobante("venta_precio_total_dol")
       rsCom("Banco") = rsComprobante("estado_cancelado")
       
       rsCom("Banco") = rsComprobante("estado_codigo")
       rsCom("Banco") = rsComprobante("beneficiario_denominacion_cobrador")
       rsCom("Banco") = rsComprobante("venta_monto_total_bs")
       rsCom("Banco") = rsComprobante("venta_monto_total_dol")
       rsCom("Banco") = rsComprobante("venta_monto_cobrado_bs")
       rsCom("Banco") = rsComprobante("venta_monto_cobrado_dol")
       
'       If rsComprobante("cheque_o_trf") = "D" Then
'          rsCom("Transf_Cheq") = "DEPOSITO BANCO"
'       End If
'       If rsComprobante("cheque_o_trf") = "C" Then
'          rsCom("Transf_Cheq") = "CHEQUE/EFECTIVO"
'       End If
'       If rsComprobante("cheque_o_trf") = "T" Then
'          rsCom("Transf_Cheq") = "TRANSFERENCIA"
'       End If
       'rsCom("Literal") = Literal(rsComprobante("Monto_Bolivianos")) & " " & "BOLIVIANOS"
       
       rsCom.Update
'wwwwwwwwwwwwwww
'                      dbo.av_ventas_detalle_rep.venta_saldo_p_cobrar_bs, dbo.av_ventas_detalle_rep.venta_saldo_p_cobrar_dol,
'                      dbo.av_ventas_detalle_rep.venta_plazo_dias_calendario, dbo.av_ventas_detalle_rep.beneficiario_telefono_fijo,
'                      dbo.av_ventas_detalle_rep.beneficiario_telefono_Cel, dbo.av_ventas_detalle_rep.beneficiario_email, dbo.av_ventas_detalle_rep.munic_descripcion,
'                      dbo.av_ventas_detalle_rep.zona_denominacion, dbo.av_ventas_detalle_rep.calle_denominacion, dbo.av_ventas_detalle_rep.calle_tipo,
'                      dbo.gc_proceso_nivel1.proceso_descripcion, dbo.gc_proceso_nivel2.subproceso_descripcion, dbo.gc_proceso_nivel3.etapa_descripcion,
'                      dbo.pc_poa_actividad.poa_descripcion, dbo.av_ventas_detalle_rep.beneficiario_edif_nro, dbo.av_ventas_detalle_rep.beneficiario_edif_piso_nro,
'                      dbo.av_ventas_detalle_rep.beneficiario_edif_depto_nro, dbo.ao_ventas_cobranza.cobranza_prog_codigo,
'                      dbo.ao_ventas_cobranza.beneficiario_codigo_resp AS beneficiario_codigo_resp_cbr, dbo.ao_ventas_cobranza.beneficiario_codigo AS beneficiario_codigo_cbr,
'                      dbo.ao_ventas_cobranza.beneficiario_codigo_fac, dbo.ao_ventas_cobranza.cobranza_programada_bs, dbo.ao_ventas_cobranza.cobranza_programada_dol,
'                      dbo.ao_ventas_cobranza.cobranza_total_bs, dbo.ao_ventas_cobranza.cobranza_total_dol, dbo.ao_ventas_cobranza.cobranza_fecha_prog,
'                      dbo.ao_ventas_cobranza.cobranza_fecha_cobro, dbo.ao_ventas_cobranza.cobranza_observaciones, dbo.ao_ventas_cobranza.literal,
'                      dbo.ao_ventas_cobranza.proceso_codigo AS proceso_codigo_cbr, dbo.ao_ventas_cobranza.subproceso_codigo AS subproceso_codigo_cbr,
'                      dbo.ao_ventas_cobranza.etapa_codigo AS etapa_codigo_cbr, dbo.ao_ventas_cobranza.clasif_codigo AS clasif_codigo_cbr,
'                      dbo.ao_ventas_cobranza.doc_codigo AS doc_codigo_cbr, dbo.ao_ventas_cobranza.doc_numero AS doc_numero_cbr,
'                      dbo.ao_ventas_cobranza.poa_codigo AS poa_codigo_cbr, dbo.ao_ventas_cobranza.estado_codigo AS estado_codigo_cbr, dbo.ao_ventas_cobranza.foto,
'                      dbo.ao_ventas_cobranza.doc_codigo_fac, dbo.ao_ventas_cobranza.cobranza_nro_factura, dbo.ao_ventas_cobranza.cobranza_nro_autorizacion,
'                      dbo.ao_ventas_cobranza.factura_impresa, dbo.ao_ventas_cobranza.cta_codigo, dbo.ao_ventas_cobranza.cobranza_codigo,
'                      dbo.fc_dosificacion_docs.dosifica_autorizacion, dbo.fc_dosificacion_docs.dosifica_fecha, dbo.fc_dosificacion_docs.correl, dbo.fc_dosificacion_docs.correl_ini,
'                      dbo.fc_dosificacion_docs.correl_fin, dbo.fc_dosificacion_docs.dosifica_fecha_ini, dbo.fc_dosificacion_docs.dosifica_fecha_fin,
'                      dbo.fc_dosificacion_docs.dosifica_fecha_limite , dbo.fc_dosificacion_docs.dosifica_codigo_control, dbo.ao_ventas_cobranza.cobranza_codigo_control



'       db.Execute "insert into fo_Comprobantes (Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Justificacion, Nro_cheque, Banco, Transf_Cheq, Literal)" & _
'                  "values ('" & DtGComprobantes.Columns(0) & "','" & DtGComprobantes.Columns(1) & "','" & DtGComprobantes.Columns(2) & "', " & DtGComprobantes.Columns(3) & ",'" & DtGComprobantes.Columns(4) & "'," & DtGComprobantes.Columns(5) & ",'" & DtGComprobantes.Columns(6) & "','','" & DtGComprobantes.Columns(7) & "','" & DtGComprobantes.Columns(8) & "', '" & DtGComprobantes.Columns(9) & "' + 'BOLIVIANOS' ) "
    
    
    End If
    Refrescar
    Retornar.Enabled = True
End Sub

Public Sub Refrescar()
    Set rsCom = New ADODB.Recordset
    rsCom.Open "select * from fo_comprobantes ", db, adOpenKeyset, adLockOptimistic
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

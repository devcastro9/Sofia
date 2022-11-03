VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmActivacionCheques 
   Caption         =   "Tesoreria  - Caja -  Operacion de Cheques "
   ClientHeight    =   8595
   ClientLeft      =   -3315
   ClientTop       =   -450
   ClientWidth     =   11400
   Icon            =   "FrmActivacionCheques.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelCuenta 
      Height          =   240
      Left            =   11565
      Picture         =   "FrmActivacionCheques.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Borra la Cuenta de la lista"
      Top             =   1575
      Width           =   255
   End
   Begin VB.CommandButton cmdDelChTr 
      Height          =   240
      Left            =   10365
      Picture         =   "FrmActivacionCheques.frx":1174
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Borra el Ch/Tr de la lista"
      Top             =   1575
      Width           =   255
   End
   Begin VB.ListBox LstCuenta 
      BackColor       =   &H00DEFEFA&
      Height          =   5130
      Left            =   10635
      TabIndex        =   14
      Top             =   1830
      Width           =   1170
   End
   Begin VB.OptionButton OptTransferencias 
      Caption         =   "Transferencias"
      Height          =   255
      Left            =   3300
      TabIndex        =   12
      Top             =   1230
      Width           =   1500
   End
   Begin VB.OptionButton OptCheques 
      Caption         =   "Cheques"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1230
      Width           =   1785
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   11340
      TabIndex        =   6
      Top             =   0
      Width           =   11400
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
         Left            =   60
         TabIndex        =   10
         Top             =   495
         Width           =   1110
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   9
         Top             =   525
         Width           =   2460
      End
      Begin VB.Label LblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   8
         Top             =   495
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPERACION CHEQUES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   4515
         TabIndex        =   7
         Top             =   135
         Width           =   3285
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "FrmActivacionCheques.frx":141E
         Top             =   0
         Width           =   11640
      End
   End
   Begin VB.ListBox LstCheques 
      BackColor       =   &H00DEFEFA&
      Height          =   5130
      Left            =   9450
      TabIndex        =   4
      Top             =   1830
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoPagos 
      Height          =   420
      Left            =   1740
      Top             =   8580
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   741
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
      Caption         =   "Cheques"
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
   Begin VB.Frame FraOpciones 
      Height          =   6060
      Left            =   15
      TabIndex        =   1
      Top             =   915
      Width           =   1245
      Begin VB.CommandButton CmdConsulta 
         Caption         =   "Consulta"
         Height          =   510
         Left            =   150
         TabIndex        =   25
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton CmdCobrado 
         Caption         =   "Cobrado"
         Height          =   510
         Left            =   150
         TabIndex        =   24
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton CmdDevuelto 
         Caption         =   "Devuelto"
         Height          =   510
         Left            =   150
         TabIndex        =   23
         Top             =   1530
         Width           =   975
      End
      Begin VB.CommandButton CmdAnulado 
         Caption         =   "Anulado"
         Height          =   510
         Left            =   150
         TabIndex        =   22
         Top             =   2550
         Width           =   975
      End
      Begin VB.CommandButton CmdEntregado 
         Caption         =   "Entregado"
         Height          =   510
         Left            =   150
         TabIndex        =   21
         Top             =   1020
         Width           =   975
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   510
         Left            =   150
         TabIndex        =   20
         Top             =   3060
         Width           =   975
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   510
         Left            =   150
         TabIndex        =   19
         Top             =   3570
         Width           =   975
      End
      Begin VB.CommandButton CmdActualizarDatos 
         Caption         =   "Actualizar Datos"
         Height          =   510
         Left            =   150
         TabIndex        =   18
         Top             =   4590
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   150
         Picture         =   "FrmActivacionCheques.frx":2848E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5145
         Width           =   975
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Impresión"
         Height          =   795
         Left            =   150
         Picture         =   "FrmActivacionCheques.frx":288D0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   195
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DtGPagos 
      Height          =   5415
      Left            =   1350
      TabIndex        =   0
      Top             =   1560
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   2058
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
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFechaRegistro 
      Height          =   330
      Left            =   7530
      TabIndex        =   15
      Top             =   1170
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   51314689
      CurrentDate     =   36413
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha de la Operación"
      Height          =   240
      Left            =   5790
      TabIndex        =   17
      Top             =   1245
      Width           =   1725
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha Inicio"
      Height          =   240
      Left            =   30
      TabIndex        =   16
      Top             =   0
      Width           =   1590
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "SELECCIONE EL CHEQUE A OPERARSE"
      Height          =   195
      Left            =   1350
      TabIndex        =   13
      Top             =   1005
      Width           =   3045
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     Cheques             Cuentas"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   9450
      TabIndex        =   5
      Top             =   1560
      Width           =   2370
   End
End
Attribute VB_Name = "FrmActivacionCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  ADFIN-2002
' Módulo:                   Operaciones sobre cheques y transferencias
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmActivaciónCheques
' Descipción :              Control de los status de Cheq/Trans de entrgado, pagado, anulado,cobrado
' Formularios relacionados: Main.frm (Padre)
'                           CryopCheques
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creación         14/Abril/ 2001
' Fecha última modificación 01/May/ 2001
' Versión:                  2.0
' Modificado por            Freddy Quiroz Nina
' Fecha de Inicio           28/Nov/2001
'========================================================================================

Dim rsComprobante As New ADODB.Recordset
Dim rscheques As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim NrosChequeImprimir As String
Dim queryinicial As String
Dim sino As Variant

Private Sub CmdActualizarDatos_Click()
    Screen.MousePointer = vbHourglass
    'MsgBox "Espere mensaje de finalización de actualizacion....", vbCritical + vbDefaultButton1, "Validación de Datos"
    'Copia_Registros_Cheques
    db.Execute "exec ActualizaChequesOperaciones"
    'MsgBox "Fin de proceso de actualizacion de cheques"
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdAnulado_Click()
  If LstCheques.ListCount > 0 Then
    sino = MsgBox("Está seguro de colocar este status de ANULADO?", vbYesNo + vbQuestion, "Atenciòn")
    If sino = vbYes Then RegistrarEstadoCheque ("A")
  Else
    MsgBox "No existen Cheques o Transferencias seleccionados!", vbInformation + vbOKOnly, "Atencion"
  End If
End Sub

Private Sub CmdBuscar_Click()
  Set FrmBuscaEnTodo.GridTrabajo = DtGPagos
  FrmBuscaEnTodo.QueryUtilizado = queryinicial
  Set FrmBuscaEnTodo.RecordsetTrabajo = rsComprobante
  FrmBuscaEnTodo.EnGridPropio = False
  FrmBuscaEnTodo.Show vbModal

'  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'  Dim ClBuscaSec As  ClBuscaSecuencialEnRS
'  PosibleApliqueFiltro = False
'  Dim rsNada As ADODB.Recordset
'  Dim GrSqlAux As String
'  Set ClBuscaGrid = New  ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.EsTdbGrid = False
'  Set ClBuscaGrid.GridTrabajo = DtGPagos
'  ClBuscaGrid.QueryUtilizado = queryinicial
'  Set ClBuscaGrid.RecordsetTrabajo = rsComprobante
'  ClBuscaGrid.CamposVisibles = "110"
'  ClBuscaGrid.Ejecutar
'  PosibleApliqueFiltro = True
End Sub

Private Sub CmdCobrado_Click()
  If LstCheques.ListCount > 0 Then
    sino = MsgBox("Está seguro de colocar este status de COBRADO?", vbYesNo + vbQuestion, "Atenciòn")
    If sino = vbYes Then RegistrarEstadoCheque ("C")
  Else
    MsgBox "No existen Cheques o Transferencias seleccionados!", vbInformation + vbOKOnly, "Atencion"
  End If
End Sub

Private Sub cmdDelChTr_Click()
  LstCheques_DblClick
End Sub

Private Sub cmdDelCuenta_Click()
  LstCuenta_DblClick
End Sub

Private Sub CmdDevuelto_Click()
  If LstCheques.ListCount > 0 Then
    sino = MsgBox("Está seguro de colocar este status de DEVUELTO?", vbYesNo + vbQuestion, "Atenciòn")
    If sino = vbYes Then RegistrarEstadoCheque ("D")
  Else
    MsgBox "No existen Cheques o Transferencias seleccionados!", vbInformation + vbOKOnly, "Atencion"
  End If
End Sub

Private Sub CmdEntregado_Click()
If LstCheques.ListCount > 0 Then
    sino = MsgBox("Está seguro de colocar este status de ENTREGADO?", vbYesNo + vbQuestion, "Atenciòn")
    If sino = vbYes Then RegistrarEstadoCheque ("E")
Else
  MsgBox "No existen Cheques o Transferencias seleccionados!", vbInformation + vbOKOnly, "Atencion"
End If
End Sub

Private Sub cmdImprimir_Click()
    If OptCheques.Value = True Then
        FrmDesplegado.LblTitulo = "Impresión Histórico de Cheques"
    End If
    If OptTransferencias.Value = True Then
        FrmDesplegado.LblTitulo = "Impresión Histórico de Transferencias"
    End If
    FrmDesplegado.Show
End Sub

Private Sub CmdLimpiar_Click()
    LstCheques.Clear
    LstCuenta.Clear
End Sub

Private Sub CmdConsulta_Click()
    FrmConsultaEstadoCheque.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
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

Private Sub DtGPagos_DblClick()
Dim bandera As Integer
Dim i As Integer
    If DtGPagos.Columns(0).Value = "" Then
        MsgBox "No existe(n) cheque(s) para procesar", vbInformation + vbCritical, "Atencion"
        Exit Sub
    End If
    
    bandera = 0
    For i = 0 To LstCheques.ListCount - 1
         LstCheques.ListIndex = i
         If LstCheques.Text = DtGPagos.Columns(0) Then
              bandera = 1
         End If
    Next i
    
    If bandera = 0 Then
        LstCheques.AddItem DtGPagos.Columns(0)
        LstCuenta.AddItem DtGPagos.Columns(9)
    End If
End Sub

Private Sub Form_Load()
    Lblusuario = GlNombreUsuario
    DTPFechaRegistro.Value = Date
    OptCheques.Value = True
    Set rsComprobante = New ADODB.Recordset
    DTPFechaRegistro.Value = Date
    Set rsComprobante = New ADODB.Recordset
'-------------------------
'    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf , fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, " & _
'    "pago_detalle.codigo_pago, to_cheques_operaciones.estado_impreso, to_cheques_operaciones.estado_entregado, to_cheques_operaciones.estado_cobrado, to_cheques_operaciones.estado_devuelto, to_cheques_operaciones.estado_anulado, fc_cuenta_bancaria.Cta_codigo,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, " & _
'    "pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
'    "FROM pago_detalle INNER JOIN " & _
'    "fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario Inner Join " & _
'    "fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.cta_codigo " & _
'    "LEFT JOIN to_cheques_operaciones ON pago_detalle.cta_codigo = to_cheques_operaciones.cta_codigo AND " & _
'    "pago_detalle.numero_cheque_trf = to_cheques_operaciones.numero_cheque WHERE (pago_detalle.cheque_o_trf = 'C') ", db, adOpenKeyset, adLockOptimistic
'-------------------------
    rsComprobante.Open queryinicial, db, adOpenKeyset, adLockOptimistic

    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set adoPagos.Recordset = rsComprobante
    End If
	Call SeguridadSet(Me)
End Sub

Public Sub Determina_Cheques()
Dim i As Integer
    NrosChequeImprimir = " "
    For i = 0 To LstCheques.ListCount - 2
        LstCheques.ListIndex = i
        NrosChequeImprimir = NrosChequeImprimir & "numero_cheque= " & "'" & LstCheques.Text & "'" & " Or "
    Next i
    LstCheques.ListIndex = i
    NrosChequeImprimir = NrosChequeImprimir + "numero_cheque = " & "'" & LstCheques.Text & "'"
End Sub

Private Sub LstCheques_Click()
  LstCuenta.ListIndex = LstCheques.ListIndex
End Sub

Private Sub LstCheques_DblClick()
  If LstCheques.ListCount > 0 Then
    LstCuenta.ListIndex = LstCheques.ListIndex
    If LstCheques.ListIndex > -1 Then
      LstCheques.RemoveItem (LstCheques.ListIndex)
      LstCuenta.RemoveItem (LstCuenta.ListIndex)
    End If
  End If
End Sub

Private Sub LstCuenta_Click()
  LstCheques.ListIndex = LstCuenta.ListIndex
End Sub

Private Sub LstCuenta_DblClick()
  If LstCheques.ListCount > 0 Then
    LstCheques.ListIndex = LstCuenta.ListIndex
    If LstCheques.ListIndex > -1 Then
      LstCheques.RemoveItem (LstCheques.ListIndex)
      LstCuenta.RemoveItem (LstCuenta.ListIndex)
    End If
  End If
End Sub

Private Sub OptCheques_Click()
   LblTitulo.Caption = "OPERACIONES CHEQUES"
   Label9 = "SELECCIONE LOS CHEQUES A PROCESARSE"
   Label1 = "     Cheques             Cuentas"
   CmdLimpiar_Click
'---------------------------
'    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf as NRO_DOC, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, " & _
'    "pago_detalle.codigo_pago, to_cheques_operaciones.estado_impreso, to_cheques_operaciones.estado_entregado, to_cheques_operaciones.estado_cobrado, to_cheques_operaciones.estado_devuelto, to_cheques_operaciones.estado_anulado, fc_cuenta_bancaria.Cta_codigo,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, " & _
'    "pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
'    "FROM pago_detalle INNER JOIN " & _
'    "fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario Inner Join " & _
'    "fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.cta_codigo " & _
'    "Inner Join  to_cheques_operaciones ON pago_detalle.cta_codigo = to_cheques_operaciones.cta_codigo AND " & _
'    "pago_detalle.numero_cheque_trf = to_cheques_operaciones.numero_cheque WHERE (pago_detalle.cheque_o_trf = 'C') ", db, adOpenKeyset, adLockOptimistic
'---------------------------
    queryinicial = "SELECT DISTINCT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, " & _
    "pago_detalle.codigo_pago, to_cheques_operaciones.estado_impreso, to_cheques_operaciones.estado_entregado, to_cheques_operaciones.estado_cobrado, to_cheques_operaciones.estado_devuelto, to_cheques_operaciones.estado_anulado, fc_cuenta_bancaria.Cta_codigo,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, " & _
    "pago_detalle.org_codigo , pago_detalle.fecha_pago " & _
    "FROM pagos INNER JOIN pago_detalle ON pagos.ges_gestion = pago_detalle.Ges_gestion AND " & _
    "pagos.org_codigo = pago_detalle.org_codigo AND pagos.codigo_pago = pago_detalle.codigo_pago INNER JOIN " & _
    "fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario INNER JOIN " & _
    "fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.cta_codigo " & _
    "INNER JOIN to_cheques_operaciones ON pago_detalle.cta_codigo = to_cheques_operaciones.cta_codigo AND " & _
    "pago_detalle.numero_cheque_trf = to_cheques_operaciones.numero_cheque WHERE (pagos.estado_pagado='S' AND pago_detalle.estado_aprobacion='A' AND pago_detalle.cheque_o_trf = 'C')"

    If rsComprobante.State = 1 Then rsComprobante.Close
    rsComprobante.Open queryinicial, db, adOpenKeyset, adLockOptimistic

    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set adoPagos.Recordset = rsComprobante
    End If
End Sub

Private Sub OptTransferencias_Click()
    LblTitulo.Caption = "OPERACIONES TRANSFERENCIAS"
    Label9 = "SELECCIONE LAS TRANSFERENCIAS A PROCESARSE"
    Label1 = "      Transf.             Cuentas"
    CmdLimpiar_Click
'------------------------------
'    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf as NRO_DOC, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, " & _
'    "pago_detalle.codigo_pago, to_cheques_operaciones.estado_impreso, to_cheques_operaciones.estado_entregado, to_cheques_operaciones.estado_cobrado, to_cheques_operaciones.estado_devuelto, to_cheques_operaciones.estado_anulado, fc_cuenta_bancaria.Cta_codigo,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, " & _
'    "pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
'    "FROM pago_detalle INNER JOIN " & _
'    "fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario Inner Join " & _
'    "fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.cta_codigo " & _
'    "Inner Join  to_cheques_operaciones ON pago_detalle.cta_codigo = to_cheques_operaciones.cta_codigo AND " & _
'    "pago_detalle.numero_cheque_trf = to_cheques_operaciones.numero_cheque WHERE (pago_detalle.cheque_o_trf = 'T') ", db, adOpenKeyset, adLockOptimistic
'------------------------------
    
    queryinicial = "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, " & _
    "pago_detalle.codigo_pago, to_cheques_operaciones.estado_impreso, to_cheques_operaciones.estado_entregado, to_cheques_operaciones.estado_cobrado, to_cheques_operaciones.estado_devuelto, to_cheques_operaciones.estado_anulado, fc_cuenta_bancaria.Cta_codigo,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga, " & _
    "pago_detalle.org_codigo , pago_detalle.fecha_pago " & _
    "FROM pagos INNER JOIN pago_detalle ON pagos.ges_gestion = pago_detalle.Ges_gestion AND " & _
    "pagos.org_codigo = pago_detalle.org_codigo AND pagos.codigo_pago = pago_detalle.codigo_pago INNER JOIN " & _
    "fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario INNER JOIN " & _
    "fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.cta_codigo " & _
    "INNER JOIN to_cheques_operaciones ON pago_detalle.cta_codigo = to_cheques_operaciones.cta_codigo AND " & _
    "pago_detalle.numero_cheque_trf = to_cheques_operaciones.numero_cheque WHERE (pagos.estado_pagado='S' AND pago_detalle.estado_aprobacion='A' AND pago_detalle.cheque_o_trf = 'T')"
    If rsComprobante.State = 1 Then rsComprobante.Close
    rsComprobante.Open queryinicial, db, adOpenKeyset, adLockOptimistic

    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set adoPagos.Recordset = rsComprobante
    End If
End Sub

Public Sub Copia_Registros_Cheques()
    'Abriendo operaciones cheques
    Set rsOpera = New ADODB.Recordset
    rsOpera.Open "select * from  to_cheques_operaciones", db, adOpenKeyset, adLockOptimistic

    'Abriendo cuenta bancaria
    Set rsPagoDet = New ADODB.Recordset
    rsPagoDet.Open "select * from pago_detalle", db, adOpenKeyset, adLockOptimistic
    If rsPagoDet.RecordCount > 0 Then
        While Not rsPagoDet.EOF
                If Not IsNull(rsPagoDet("Monto_Bolivianos")) Or rsPagoDet("Monto_Bolivianos") <> "" And rsPagoDet("numero_cheque_trf_destino") <> "" And Not IsNull(rsPagoDet("numero_cheque_trf_destino")) Then
                  Set rsCE = New ADODB.Recordset
                  rsCE.Open "select * from to_cheques_operaciones where numero_cheque='" & rsPagoDet("numero_cheque_trf") & "' and cta_codigo='" & rsPagoDet("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
                  If rsCE.RecordCount <= 0 Then
                    rsOpera.AddNew
                    rsOpera("numero_cheque") = rsPagoDet("numero_cheque_trf")
                    rsOpera("cta_codigo") = rsPagoDet("cta_codigo")
                    rsOpera("estado_impreso") = IIf(IsNull(rsPagoDet!fecha_impresion_cheque), "N", "S") '"S"
                    rsOpera("estado_entregado") = "N"
                    rsOpera("estado_cobrado") = "N"
                    rsOpera("estado_anulado") = "N"
                    rsOpera("estado_devuelto") = "N"
                    rsOpera("usr_usuario") = GlNombreUsuario '"General"
                    
                    rsOpera("fecha_registro") = CDate(Date)
                    rsOpera("hora_registro") = Time
                    rsOpera.Update
                  End If
                End If
                rsPagoDet.MoveNext
         Wend
    End If
End Sub

Public Sub Refrescar()
    If rsComprobante.State = 1 Then rsComprobante.Close
    rsComprobante.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set adoPagos.Recordset = rsComprobante
    End If
End Sub

Function Cheque_Transf_Valido(pImpreso As String, pEntregado As String, pCobrado As String, pDevuelto As String, pAnulado As String, pProceso As String) As String
Cheque_Transf_Valido = ""
Select Case pProceso
  Case "E":
          If pImpreso = "S" And pEntregado = "N" And pCobrado = "N" And pDevuelto = "N" And pAnulado = "N" Then
            Cheque_Transf_Valido = ""
          Else
            If pImpreso = "N" And pEntregado = "N" And pCobrado = "N" And pDevuelto = "N" And pAnulado = "N" Then
              Cheque_Transf_Valido = "El cheque o transferencia no ha sido Impreso"
            Else
              Cheque_Transf_Valido = "El cheque o transferencia ya ha sido Entregado, Cobrado, Devuelto o Anulado"
            End If
          End If
  Case "C":
          If pImpreso = "S" And pEntregado = "S" And pCobrado = "N" And pDevuelto = "N" And pAnulado = "N" Then
            Cheque_Transf_Valido = ""
          Else
            If pImpreso = "S" And pEntregado = "S" And pCobrado = "S" And pDevuelto = "N" And pAnulado = "N" Then
              Cheque_Transf_Valido = "El cheque o transferencia ya ha sido Cobrado"
            Else
              Cheque_Transf_Valido = "El cheque o transferencia ya ha sido Devuelto o Anulado"
            End If
          End If
  Case "D":
          If pImpreso = "S" And pEntregado = "S" And pCobrado = "N" And pDevuelto = "N" And pAnulado = "N" Then
            Cheque_Transf_Valido = ""
          Else
            If pImpreso = "S" And pEntregado = "S" And pCobrado = "N" And pDevuelto = "S" And pAnulado = "N" Then
              Cheque_Transf_Valido = "El cheque o transferencia ya ha sido Devuelto"
            Else
              Cheque_Transf_Valido = "El cheque o transferencia ya ha sido Cobrado o Anulado"
            End If
          End If
  Case "A":
          If pImpreso = "S" And pEntregado = "S" And pCobrado = "N" And pAnulado = "N" Then
            Cheque_Transf_Valido = ""
          Else
            If pImpreso = "S" And pEntregado = "S" And pCobrado = "N" And pAnulado = "S" Then
              Cheque_Transf_Valido = "El cheque o transferencia ya ha sido Anulado"
            Else
              Cheque_Transf_Valido = "El cheque o transferencia ya ha sido Cobrado"
            End If
          End If
End Select
End Function

Private Sub RegistrarEstadoCheque(pEstado As String)
Dim i As Integer
Dim HuboProblemas As Boolean
Dim CadProceso As String
Dim Proceso As String
Dim tipo As String
    If LstCheques.ListCount > 0 Then
         HuboProblemas = False
         If OptCheques.Value = True Then
            tipo = "C"
         Else
            tipo = "T"
         End If
         For i = 0 To LstCheques.ListCount - 1
              LstCheques.ListIndex = i
              LstCuenta.ListIndex = i
              rscheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "' and cta_codigo= '" & LstCuenta.Text & "'  order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
              With rscheques
                  If .RecordCount > 0 Then
                        CadProceso = Cheque_Transf_Valido(!estado_impreso, !estado_entregado, !estado_cobrado, !estado_devuelto, !estado_anulado, pEstado)
                        If CadProceso = "" Then
                             Select Case pEstado
                                    Case "E":
                                            !estado_entregado = "S"
                                            !fecha_entregado = DTPFechaRegistro.Value
                                            Proceso = "ENTREGADO"
                                    Case "C":
                                            !estado_cobrado = "S"
                                            !fecha_cobrado = DTPFechaRegistro.Value
                                            Proceso = "COBRADO"
                                    Case "D":
                                            !estado_devuelto = "S"
                                            !fecha_devuelto = DTPFechaRegistro.Value
                                            Proceso = "DEVUELTO"
                                    Case "A":
                                            !estado_anulado = "S"
                                            !fecha_anulado = DTPFechaRegistro.Value
                                            Proceso = "ANULADO"
                             End Select
                             !usr_usuario = GlUsuario 'Lblusuario.Caption
                             !fecha_registro = Format(Date, "dd/mm/yyyy")
                             !hora_registro = Format(Time, "hh:mm:ss")
                             .Update
                        Else
                             HuboProblemas = True
                             CadProceso = !numero_cheque & "    " & !Cta_Codigo & "     " & CadProceso
                             Call FrmInformeProcesos.Principal(tipo, pEstado, CadProceso)
                        End If
                  Else
                      CadProceso = !numero_cheque & "    " & !Cta_Codigo & "     " & "Cheque o Transferencia no existe"
                      Call FrmInformeProcesos.Principal(tipo, pEstado, CadProceso)
                  End If
                  .Close
              End With
         Next i
         If Not HuboProblemas Then
          CadProceso = "El cambio de estado << " & Proceso & " >> se realizo satisfactoriamente."
          Call FrmInformeProcesos.Principal(tipo, pEstado, CadProceso)
         End If

         FrmInformeProcesos.Show vbModal
         CmdLimpiar_Click
         Refrescar
    End If
End Sub

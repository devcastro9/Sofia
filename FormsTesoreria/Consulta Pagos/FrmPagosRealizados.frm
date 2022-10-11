VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPagosRealizados 
   Caption         =   "Pagos Pendientes de Pago"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "FrmPagosRealizados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryCmpte 
      Left            =   1530
      Top             =   1080
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   11775
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAGOS REALIZADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   4530
         TabIndex        =   3
         Top             =   105
         Width           =   3285
      End
      Begin VB.Label LblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10320
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   0
         Picture         =   "FrmPagosRealizados.frx":0ECA
         Top             =   0
         Width           =   15360
      End
   End
   Begin VB.Frame FraBusqueda 
      Height          =   1965
      Left            =   1365
      TabIndex        =   10
      Top             =   3180
      Visible         =   0   'False
      Width           =   6060
      Begin VB.Frame Frame1 
         Height          =   1065
         Left            =   105
         TabIndex        =   13
         Top             =   150
         Width           =   5820
         Begin VB.TextBox TxtValor 
            Height          =   285
            Left            =   3165
            TabIndex        =   16
            Top             =   645
            Width           =   2505
         End
         Begin VB.ComboBox CmbOperador 
            Height          =   315
            ItemData        =   "FrmPagosRealizados.frx":2A70
            Left            =   1965
            List            =   "FrmPagosRealizados.frx":2A83
            TabIndex        =   15
            Top             =   630
            Width           =   1065
         End
         Begin VB.ComboBox CmbCampo 
            Height          =   315
            Left            =   45
            TabIndex        =   14
            Top             =   630
            Width           =   1815
         End
         Begin VB.Label LblValor 
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3315
            TabIndex        =   19
            Top             =   255
            Width           =   675
         End
         Begin VB.Label LblOperador 
            Caption         =   "Operador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1965
            TabIndex        =   18
            Top             =   255
            Width           =   885
         End
         Begin VB.Label LblCampo 
            Caption         =   "Campo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   17
            Top             =   255
            Width           =   615
         End
      End
      Begin VB.CommandButton CmdCancelarBusqueda 
         Caption         =   "Cancelar"
         Height          =   390
         Left            =   3180
         TabIndex        =   12
         Top             =   1320
         Width           =   1140
      End
      Begin VB.CommandButton CmdEjecutarBusqueda 
         Caption         =   "Ejecutar"
         Height          =   390
         Left            =   1230
         TabIndex        =   11
         Top             =   1320
         Width           =   1140
      End
   End
   Begin VB.Frame FraOpciones 
      BackColor       =   &H00C0C0C0&
      Height          =   7155
      Left            =   10560
      TabIndex        =   4
      Top             =   1080
      Width           =   1245
      Begin VB.CommandButton CmdImprimeGrid 
         Caption         =   "Imprime grid"
         Height          =   735
         Left            =   135
         TabIndex        =   20
         Top             =   2685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   795
         Left            =   135
         Picture         =   "FrmPagosRealizados.frx":2A9A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1890
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   795
         Left            =   135
         Picture         =   "FrmPagosRealizados.frx":2CA4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   825
         Left            =   135
         Picture         =   "FrmPagosRealizados.frx":338E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4995
         Width           =   930
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar"
         Height          =   795
         Left            =   135
         Picture         =   "FrmPagosRealizados.frx":3598
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1095
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc AdoPagos 
      Height          =   420
      Left            =   1290
      Top             =   8220
      Visible         =   0   'False
      Width           =   10035
      _ExtentX        =   17701
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
   Begin MSDataGridLib.DataGrid DtGPagos 
      Height          =   7080
      Left            =   0
      TabIndex        =   8
      Top             =   1110
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   12488
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPagosRealizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  ADFIN-2002
' Módulo:                   Pagos Efectuados
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmPagosRealizados.frm
' Descipción :              Pagos realizados observados por fecha de impresión o
'                           por codigo de pago, beneficiario, CI, etc.
'                           Pagos que tienen Nro. de cheque o Nro. de transferencia
'                           y por tanto asignación de cuenta y otros
' Formularios relacionados: Main.frm (Padre)
'                           CryPagos
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creación         10/Mar/ 2001
' Fecha última modificación 1/May/ 2001
' Versión:                  2.0
'========================================================================================
'Dim ClBuscaGrid As  ClBuscaEnGridExterno
Dim queryinicial As String

Dim rsComprobante As New ADODB.Recordset
Dim rsCmpte As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Public NoRegistros As Long
'Para B{usqueda
Dim ClBuscaGrid As ClBuscaEnGridExterno

'    If DtGPagos.Columns(0) = "" Then
'        MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
'        Exit Sub
'    End If
'    MsgBox "Cheque " + AdoPagos.Recordset("numero_cheque_trf") + " Devuelto"
'    Set rsCheques = New ADODB.Recordset
'    If rsCheques.State = 1 Then rsCheque.Close
'    rsCheques.Open "SELECT * FROM to_cheques WHERE numero_cheque = '" & AdoPagos.Recordset("numero_cheque_trf") & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
'    If rsCheques.RecordCount > 0 Then
'        rsCheques("estado_anulado") = "S"
'    Else
'        rsCheques.AddNew
'        rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
'        rsCheques("estado_anulado") = "S"
'    End If
'        rsCheques.Update
'End Sub

Private Sub CmdBuscar_Click()
'        On Error GoTo Error:
'                For Each CAMPOS In rsComprobante.Fields
'                    CmbCampo.AddItem CAMPOS.Name
'                Next CAMPOS
'                FraBusqueda.Visible = True
'        Exit Sub
'Error:
'            MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"

'Freddy Q.
''''    queryinicial = "SELECT Pagos.ges_Gestion, Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario,pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga,fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden,pagos.estado_pagado,Pagos.estado_devolucion " & _
''''                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) " & _
''''                   "AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo"

    queryinicial = "SELECT Pagos.ges_Gestion, Pagos.codigo_pago,Pagos.nro_comprobante_anterior,Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario,pago_detalle.fecha_pago,pago_detalle.monto_total,pago_detalle.monto_Bolivianos,pago_detalle.monto_dolares,pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga,fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden,pagos.estado_pagado,Pagos.estado_devolucion " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) " & _
                   "AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo Where Pagos.estado_anulado Is NULL"


'Dulfredo Rojas
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = DtGPagos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = AdoPagos.Recordset
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar



'Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
''Set ClBuscaGrid = New  ClBuscaEnGridExterno
'PosibleApliqueFiltro = False
'
'Dim rsNada As ADODB.Recordset
'Dim ClBuscaSec As  ClBuscaSecuencialEnRS
'Dim GrSqlAux As String
'
'  GrFrmOpciones.OptOpciones(1).Caption = "Buscar en Cabecera"
''  GrFrmOpciones.OptOpciones(1).Enabled = False
'  GrFrmOpciones.OptOpciones(2).Caption = "Buscar en Cabecera y Detalle"
'  GrFrmOpciones.Show vbModal
'  If GrFrmOpciones.POpcionElegida <> 0 Then
'    If GrFrmOpciones.POpcionElegida = 1 Then
'      Set ClBuscaGrid = New  ClBuscaEnGridExterno
'      Set ClBuscaGrid.Conexión = db
'      ClBuscaGrid.EsTdbGrid = False
'      Set ClBuscaGrid.GridTrabajo = DtGPagos
'      ClBuscaGrid.QueryUtilizado = QueryInicial
'      Set ClBuscaGrid.RecordsetTrabajo = AdoPagos.Recordset
'      'ClBuscaGrid.CamposVisibles = "110"
'      ClBuscaGrid.Ejecutar
'      PosibleApliqueFiltro = True
'    Else
'      Set ClVBusca = New  ClBuscaEnGridPropio 'DUL: Instancio Componente de Busqueda
'      'Dll con Ventana de búsqueda
'      GrSqlAux = "SELECT Pagos.ges_Gestion, Pagos.org_codigo, Pagos.codigo_pago, fc_beneficiario.denominacion_beneficiario,pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga,fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
'                 "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo"
'        ''      GrSqlAux = "SELECT pg.ges_gestion, pg.codigo_pago, pg.org_codigo, " & _
'        ''                 "pg.Nro_Comprobante_Anterior, pg.estado_compromiso, " & _
'        ''                 "pg.estado_devengado, pg.estado_pagado, " & _
'        ''                 "pg.estado_reversion_total, pg.estado_devolucion, " & _
'        ''                 "pg.estado_anulado, Pd.par_codigo, " & _
'        ''                 "Pd.Pro_programa, Pd.Pro_subprograma, " & _
'        ''                 "Pd.Pro_proyecto, Pd.Pro_actividad, " & _
'        ''                 "Pd.cta_codigo, Pd.numero_cheque_trf, " & _
'        ''                 "Pd.cta_codigo_destino, " & _
'        ''                 "Pd.codigo_beneficiario, Pd.monto_total, " & _
'        ''                 "Pd.monto_Dolares " & _
'        ''                 "FROM Pagos as Pg INNER JOIN Pago_Detalle as Pd " & _
'        ''                 "ON (Pg.Ges_gestion = pd.ges_gestion)and(Pg.codigo_pago = Pd.codigo_pago)and(Pg.org_codigo = Pd.org_codigo) "
'                 'NO PUEDE IR : "ORDER BY pg.codigo_pago"
'      Set ClVBusca.Conexión = db
'      ClVBusca.QueryUtilizado = GrSqlAux
'      ClVBusca.FiltrosMultiples = True
'      ClVBusca.OcultarPrimero = True
''      ClVBusca.AliasColumnas = "ges_gestion-Cmbte-Org-Anterior-C-D-P-R-V-A-PARTIDA-" & _
''                               "PRO-SUB-PRY-ACT-CTA_ORIGEN-Cheque_Trf-CTA_DESTINO-" & _
''                               "BENEFICIARIO-MONTO_Bs-MONTO_US"
''      ClVBusca.TamañoCampos = "1-0.6-0.5-0.4-0.3-0.3-0.3-0.3-0.3-0.3-0.8-" & _
''                              "0.4-0.4-0.4-0.4-0.7-0.6-0.7-" & _
''                              "0.8-0.6-0.8"
'      ClVBusca.Ejecutar
'      'Si eligio alguno
'      If ClVBusca.ElegidoCol1 <> "" Then
'        'Desconecta para que no se disparen los eventos del ado
'        If PosibleApliqueFiltro Then
'          If rsComprobante.State = 1 Then rsComprobante.Close
'          rsComprobante.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'          rsComprobante.Sort = "codigo_pago"
'          Set AdoPagos.Recordset = rsComprobante
'          Set DtGPagos.DataSource = AdoPagos.Recordset
'        End If
'        Set DtGPagos.DataSource = rsNada
''        Set TxtJustificacion.DataSource = rsNada
'        EntrarAdo = False
'        'Buscar Registro
'        Set ClBuscaSec = New  ClBuscaSecuencialEnRS
'        Set ClBuscaSec.Recordset = AdoPagos.Recordset
'        ClBuscaSec.Campo1 = "ges_Gestion"
'        ClBuscaSec.ValorCampo1 = ClVBusca.ElegidoCol1
'        ClBuscaSec.Campo2 = "org_codigo"
'        ClBuscaSec.ValorCampo2 = ClVBusca.ElegidoCol2
'        ClBuscaSec.Campo3 = "codigo_pago"
'        ClBuscaSec.ValorCampo3 = ClVBusca.ElegidoCol3
'        If Not ClBuscaSec.Ejecutar Then
'          MsgBox "Error en búsqueda...", vbInformation + vbOKOnly, "Atención"
'        End If
'        Set ClBuscaSec = Nothing
'
'        Set DtGPagos.DataSource = AdoPagos.Recordset
''        Set TxtJustificacion.DataSource = AdoRegularizacion.Recordset
'        EntrarAdo = True
'        AdoPagos.Recordset.Move 0
'      End If
'    End If
'  End If
''
       
    
End Sub

Private Sub CmdCancelar_Click()
    FraBusca.Visible = False
End Sub

Private Sub CmdCobrado_Click()
Dim i As Integer
Dim Cheque_Inicial As Long
Dim Cheque_Final As Long
Dim s As String
Dim k As Long

LstCheques.ListIndex = 0
If LstCheques.Text <> "" Then
   If rsCheques.State = 1 Then rsCheques.Close
   For i = 0 To LstCheques.ListCount - 1
            LstCheques.ListIndex = i
            If DtGPagos.Columns(0) = "" Then
                MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
                Exit Sub
            End If
            If rsCheques.State = 1 Then rsCheques.Close
            rsCheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
            If rsCheques.RecordCount > 0 Then
                rsCheques("estado_cobrado") = "S"
            Else
                rsCheques.AddNew
                'rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
                rsCheques("numero_cheque") = LstCheques.Text
                rsCheques("estado_cobrado") = "S"
            End If
            rsCheques("usr_usuario") = LblUsuario.Caption
            rsCheques("fecha_registro") = Date
            rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
            rsCheques.Update
    Next i
End If

s = ""
If TxtInicial.Text <> "" Then
        Cheque_Inicial = Val(TxtInicial.Text)
        Cheque_Final = Val(TxtFinal.Text)
        For k = Cheque_Inicial To Cheque_Final Step 1
            s = ""
            T = CStr(k)
            Select Case Len(T)
                   Case 1
                        s = "0000" + CStr(k)
                   Case 2
                        s = "000" + CStr(k)
                   Case 3
                        s = "00" + CStr(k)
                   Case 4
                        s = "0" + CStr(k)
                   Case 5
                        s = CStr(k)
            End Select
            
             LstCheques.AddItem s
        Next k
        
        For i = 0 To LstCheques.ListCount - 1
            LstCheques.ListIndex = i
            If rsCheques.State = 1 Then rsCheques.Close
            rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & LstCheques.Text & "'order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
            If rsCheques.RecordCount > 0 Then
                    rsCheques("estado_cobrado") = "S"
            Else
                    rsCheques.AddNew
                    rsCheques("numero_cheque") = LstCheques.Text
                    rsCheques("estado_cobrado") = "S"
            End If
            rsCheques("usr_usuario") = LblUsuario.Caption
            rsCheques("fecha_registro") = Date
            rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
            rsCheques.Update
        Next i
End If

End Sub

Private Sub CmdDevuelto_Click()
Dim i As Long
Dim Cheque_Inicial As Long
Dim Cheque_Final As Long
Dim s As String
Dim k As Long

LstCheques.ListIndex = 0
If LstCheques.Text <> "" Then
   If rsCheques.State = 1 Then rsCheques.Close
   For i = 0 To LstCheques.ListCount - 1
            LstCheques.ListIndex = i
            If DtGPagos.Columns(0) = "" Then
                MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
                Exit Sub
            End If
            If rsCheques.State = 1 Then rsCheques.Close
            rsCheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
            If rsCheques.RecordCount > 0 Then
                rsCheques("estado_devuelto") = "S"
            Else
                rsCheques.AddNew
                'rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
                rsCheques("numero_cheque") = LstCheques.Text
                rsCheques("estado_devuelto") = "S"
            End If
            rsCheques("usr_usuario") = LblUsuario.Caption
            rsCheques("fecha_registro") = Date
            rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
            
            rsCheques.Update
    Next i
End If

s = ""
If TxtInicial.Text <> "" Then
        Cheque_Inicial = Val(TxtInicial.Text)
        Cheque_Final = Val(TxtFinal.Text)
        For k = Cheque_Inicial To Cheque_Final Step 1
            s = ""
            T = CStr(k)
            Select Case Len(T)
                   Case 1
                        s = "0000" + CStr(k)
                   Case 2
                        s = "000" + CStr(k)
                   Case 3
                        s = "00" + CStr(k)
                   Case 4
                        s = "0" + CStr(k)
                   Case 5
                        s = CStr(k)
            End Select
            
             LstCheques.AddItem s
        Next k
        
        For i = 0 To LstCheques.ListCount - 1
            LstCheques.ListIndex = i
            If rsCheques.State = 1 Then rsCheques.Close
            rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & LstCheques.Text & "'order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
            If rsCheques.RecordCount > 0 Then
                    rsCheques("estado_devuelto") = "S"
            Else
                    rsCheques.AddNew
                    rsCheques("numero_cheque") = LstCheques.Text
                    rsCheques("estado_devuelto") = "S"
            End If
            
            rsCheques("usr_usuario") = LblUsuario.Caption
            rsCheques("fecha_registro") = Date
            rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
            
            rsCheques.Update
        Next i
End If

'    If DtGPagos.Columns(0) = "" Then
'        MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
'        Exit Sub
'    End If
'    MsgBox "Cheque " + AdoPagos.Recordset("numero_cheque_trf") + " Devuelto"
'    Set rsCheques = New ADODB.Recordset
'    If rsCheques.State = 1 Then rsCheque.Close
'    rsCheques.Open "SELECT * FROM to_cheques WHERE numero_cheque = '" & AdoPagos.Recordset("numero_cheque_trf") & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
'    If rsCheques.RecordCount > 0 Then
'        rsCheques("estado_devuelto") = "S"
'    Else
'        rsCheques.AddNew
'        rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
'        rsCheques("estado_devuelto") = "S"
'    End If
'        rsCheques.Update
End Sub

Private Sub CmdCancelarBusqueda_Click()
    FraBusqueda.Visible = False
End Sub

Private Sub CmdEjecutar_Click()
    Set rsComprobante = New ADODB.Recordset
    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
End Sub

Private Sub CmdEntregado_Click()
Dim i As Integer
Dim Cheque_Inicial As Long
Dim Cheque_Final As Long
Dim s As String
Dim k As Long

s = ""
If TxtInicial.Text <> "" Then
        Cheque_Inicial = Val(TxtInicial.Text)
        Cheque_Final = Val(TxtFinal.Text)
        For k = Cheque_Inicial To Cheque_Final Step 1
            s = ""
            T = CStr(k)
            Select Case Len(T)
                   Case 1
                        s = "0000" + CStr(k)
                   Case 2
                        s = "000" + CStr(k)
                   Case 3
                        s = "00" + CStr(k)
                   Case 4
                        s = "0" + CStr(k)
                   Case 5
                        s = CStr(k)
            End Select
            
             LstCheques.AddItem s
        Next k
        
        For i = 0 To LstCheques.ListCount - 1
            LstCheques.ListIndex = i
            If rsCheques.State = 1 Then rsCheques.Close
            rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & LstCheques.Text & "'order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
            If rsCheques.RecordCount > 0 Then
                    rsCheques("estado_entregado") = "S"
            Else
                    rsCheques.AddNew
                    rsCheques("numero_cheque") = LstCheques.Text
                    rsCheques("estado_entregado") = "S"
            End If
            rsCheques("usr_usuario") = LblUsuario.Caption
            rsCheques("fecha_registro") = Date
            rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
            
            rsCheques.Update
        Next i
End If


LstCheques.ListIndex = 0
If LstCheques.Text <> "" Then
   If rsCheques.State = 1 Then rsCheques.Close
   For i = 0 To LstCheques.ListCount - 1
            LstCheques.ListIndex = i
            If DtGPagos.Columns(0) = "" Then
                MsgBox "No existe cheque ", vbInformation + vbCritical, "Validación de datos"
                Exit Sub
            End If
            If rsCheques.State = 1 Then rsCheques.Close
            rsCheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
            If rsCheques.RecordCount > 0 Then
                rsCheques("estado_entregado") = "S"
            Else
                rsCheques.AddNew
                'rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
                rsCheques("numero_cheque") = LstCheques.Text
                rsCheques("estado_entregado") = "S"
            End If
            rsCheques("usr_usuario") = LblUsuario.Caption
            rsCheques("fecha_registro") = Date
            rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
            rsCheques.Update
    Next i
End If

    
    'MsgBox "Cheque " + AdoPagos.Recordset("numero_cheque_trf") + " Entregado"
'    Set rsCheques = New ADODB.Recordset
'    If rsCheques.State = 1 Then rsCheques.Close
'    rsCheques.Open "SELECT * FROM to_cheques WHERE numero_cheque = '" & AdoPagos.Recordset("numero_cheque_trf") & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
'    If rsCheques.RecordCount > 0 Then
'        rsCheques("estado_entregado") = "S"
'    Else
'        rsCheques.AddNew
'        rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
'        rsCheques("estado_entregado") = "S"
'    End If
'        rsCheques.Update
        
End Sub
Private Sub CmdEjecutarBusqueda_Click()
Dim cadena_busqueda As String
   cadena_busqueda = ""
    If CmbCAMPO = "codigo_pago" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "org_codigo" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "denominacion_beneficiario" Then
        cadena_busqueda = "fc_beneficiario." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "fecha_pago" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " = " + "#" + TxtValor + "#"
    End If
    If CmbCAMPO = "monto_bolivianos" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "tipo_cambio" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "codigo_beneficiario" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "justificacion" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "cheque_o_trf" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "numero_cheque_trf" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "cta_codigo" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "Bco_descripcion_larga" Then
        cadena_busqueda = "fc_bancos." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "literal" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "cta_descripcion_larga" Then
        cadena_busqueda = "fc_cuenta_bancaria." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "Org_descripcion" Then
        cadena_busqueda = "fc_organismo_financiamiento." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    
    If CmbCAMPO = "codigo_solicitud" Then
        cadena_busqueda = "pagos." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "codigo_orden" Then
        cadena_busqueda = "pagos." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    
    'Realizar la busqueda dado un criterio
    Set rsComprobante = New ADODB.Recordset
    If cadena_busqueda <> "" Then
        rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo,fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden,pagos.estado_pagado,Pagos.estado_devolucion " & _
                            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE " & cadena_busqueda & " order by pago_detalle.numero_cheque_trf  ", db, adOpenKeyset, adLockOptimistic
        If rsComprobante.RecordCount > 0 Then
            NoRegistros = rsComprobante.RecordCount
            'AdoCuenta.Caption = NoRegistros
            Set DtGPagos.DataSource = rsComprobante
            Set AdoPagos.Recordset = rsComprobante
        Else
            Set DtGPagos.DataSource = rsNada
            'Set AdoPagos.Recordset = rsNada
        End If
    Else
        MsgBox "No existen datos para imprimir"
        Exit Sub
    End If
    FraBusqueda.Visible = False
End Sub
Private Sub CmdImprimeGrid_Click()
Dim i As Long
Dim AUXILIAR As String
On Error GoTo temporal:
    'Imprimir datos
'    Set rsCmpte = New ADODB.Recordset
'    If rsCmpte.State = 1 Then rsCmpte.Close
'    rsCmpte.Open "SELECT * FROM to_ListadoComprobantes", db, adOpenStatic, adLockOptimistic
'    If rsCmpte.RecordCount > 0 Then
'        While Not rsCmpte.EOF
'            rsCmpte.Delete
'            rsCmpte.MoveNext
'        Wend
'    End If
    db.Execute "DELETE FROM to_ListadoComprobantes"
    Set rsCmpte = New ADODB.Recordset
    If rsCmpte.State = 1 Then rscuentas.Close
    rsCmpte.Open "SELECT * FROM to_ListadoComprobantes", db, adOpenStatic, adLockOptimistic
    i = 0
    'While I <= NoRegistros - 1
    For i = 0 To NoRegistros - 1
        rsCmpte.AddNew
        DtGPagos.Row = i
        rsCmpte("codigo_pago") = DtGPagos.Columns(0)
        rsCmpte("nro_comprobante_anterior") = DtGPagos.Columns(0)
        
        rsCmpte("org_codigo") = DtGPagos.Columns(1)
        rsCmpte("denominacion_beneficiario") = DtGPagos.Columns(2)
        rsCmpte("fecha_pago") = Format(DtGPagos.Columns(3), "dd/mm/yyyy")
        rsCmpte("monto_dolares") = DtGPagos.Columns(4)
        
        rsCmpte("monto_bolivianos") = DtGPagos.Columns(4)
        rsCmpte("monto_dolares") = DtGPagos.Columns(4)
        
        rsCmpte("tipo_cambio") = DtGPagos.Columns(5)
        rsCmpte("codigo_beneficiario") = DtGPagos.Columns(6)
        rsCmpte("justificacion") = DtGPagos.Columns(7)
        rsCmpte("cheque_o_trf") = DtGPagos.Columns(8)
        rsCmpte("numero_cheque_trf") = DtGPagos.Columns(9)
        rsCmpte("cta_codigo") = DtGPagos.Columns(10)
        rsCmpte("bco_descripcion_larga") = DtGPagos.Columns(11)
        rsCmpte("literal") = DtGPagos.Columns(12)
        rsCmpte("cta_descripcion_larga") = DtGPagos.Columns(13)
        
        'Freddy
        rsCmpte("estado_pagado") = DtGPagos.Columns(14)
        rsCmpte("estado_devolucion") = DtGPagos.Columns(15)

        rsCmpte.Update
     'I = I + 1
     Next i
    'Wend
    'Cry.PaperOrientation = crLandscape
    RepPagos.Show
    
 Exit Sub
temporal:
    Set rsCmpte = New ADODB.Recordset
    If rsCmpte.State = 1 Then rsCmpte.Close
    rsCmpte.Open "SELECT * FROM to_ListadoComprobantes", db, adOpenDynamic, adLockOptimistic
    Resume
End Sub

Private Sub CmdLimpiar_Click()
    Set rsComprobante = New ADODB.Recordset
    rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_organismo_financiamiento.Org_descripcion, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga " & _
                        "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pagos.estado_pagado='S'", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
End Sub

Private Sub Cmdimprimir_Click()

        db.Execute "delete from to_listadocomprobantes"
        Set rsComprobante.DataSource = AdoPagos.Recordset
        If rsComprobante.RecordCount > 0 Then
             Set rsCmpte = New ADODB.Recordset
             If rsCmpte.State = 1 Then rsCmpte.Close
             rsCmpte.Open "SELECT * FROM to_ListadoComprobantes", db, adOpenKeyset, adLockOptimistic

        While Not rsComprobante.EOF
          'If rsComprobante("estado_pagado") = "S" Then
              rsCmpte.AddNew
                rsCmpte("codigo_pago") = rsComprobante("codigo_pago")
                rsCmpte("nro_comprobante_anterior") = rsComprobante("nro_comprobante_anterior")
                rsCmpte("org_codigo") = rsComprobante("org_codigo")
                If Not IsNull(rsComprobante("denominacion_beneficiario")) Then rsCmpte("denominacion_beneficiario") = rsComprobante("denominacion_beneficiario")
                If Not IsNull(rsComprobante("fecha_pago")) Then rsCmpte("fecha_pago") = Format(rsComprobante("fecha_pago"), "dd/mm/yyyy")
                rsCmpte("monto_total") = rsComprobante("monto_total")
                rsCmpte("monto_bolivianos") = rsComprobante("monto_bolivianos")
                rsCmpte("monto_dolares") = rsComprobante("monto_dolares")
                rsCmpte("tipo_cambio") = rsComprobante("tipo_cambio")
                rsCmpte("codigo_beneficiario") = rsComprobante("codigo_beneficiario")
                rsCmpte("justificacion") = rsComprobante("justificacion")
                rsCmpte("cheque_o_trf") = rsComprobante("cheque_o_trf")
                rsCmpte("numero_cheque_trf") = Val(rsComprobante("numero_cheque_trf"))
                rsCmpte("cta_codigo") = rsComprobante("cta_codigo")
                rsCmpte("bco_descripcion_larga") = rsComprobante("bco_descripcion_larga")
                rsCmpte("literal") = rsComprobante("literal")
                rsCmpte("cta_descripcion_larga") = rsComprobante("cta_descripcion_larga")
                'Freddy
                rsCmpte("estado_pagado") = rsComprobante("estado_pagado")
                rsCmpte("estado_devolucion") = rsComprobante("estado_devolucion")
                
              rsCmpte.Update
            'End If
              rsComprobante.MoveNext
          Wend
        End If


        Set rslsta = New ADODB.Recordset
        queryinicial = "SELECT * FROM to_ListadoComprobantes"
        rslsta.Open queryinicial, db, adOpenKeyset, adLockOptimistic
        If rslsta.RecordCount <= 0 Then
            MsgBox "No existen registros para imprimir", vbInformation + vbCritical, "Validación de datos"
            Exit Sub
        End If
       
       CryCmpte.ReportFileName = App.Path & "\FormsTesoreria\Consulta Pagos\Rep_ListaComprobantes.rpt"
       IResult = CryCmpte.PrintReport
       If IResult <> 0 Then
           MsgBox CryCmpte.LastErrorNumber & " : " & CryCmpte.LastErrorString, vbCritical + vbOKOnly, "Error..."
       End If


End Sub

Private Sub CmdImprimirBusqueda_Click()
Dim cadena_busqueda As String
   cadena_busqueda = ""
    If CmbCAMPO = "codigo_pago" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "org_codigo" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "denominacion_beneficiario" Then
        cadena_busqueda = "fc_beneficiario." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "fecha_pago" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " = " + "#" + TxtValor + "#"
    End If
    If CmbCAMPO = "monto_bolivianos" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "tipo_cambio" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "codigo_beneficiario" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "justificacion" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "cheque_o_trf" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "numero_cheque_trf" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "cta_codigo" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "Bco_descripcion_larga" Then
        cadena_busqueda = "fc_bancos." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "literal" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "cta_descripcion_larga" Then
        cadena_busqueda = "fc_cuenta_bancaria." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "Org_descripcion" Then
        cadena_busqueda = "fc_organismo_financiamiento." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    
    'Realizar la busqueda dado un criterio
    'Imprimir datos
    
    db.Execute "DELETE FROM to_ListadoComprobantes"
'    Set rsCmpte = New ADODB.Recordset
'    If rsCmpte.State = 1 Then rsCmpte.Close
'    rsCmpte.Open "SELECT * FROM to_ListadoComprobantes", db, adOpenStatic, adLockOptimistic
'    If rsCmpte.RecordCount > 0 Then
'        While Not rsCmpte.EOF
'            rsCmpte.Delete
'            rsCmpte.MoveNext
'        Wend
'    End If

    If cadena_busqueda <> "" Then
       Set rsCmpte = New ADODB.Recordset
       If rsCmpte.State = 1 Then rsCmpte.Close
            rsCmpte.Open "SELECT * FROM to_ListadoComprobantes", db, adOpenKeyset, adLockOptimistic
    
        Set rsComprobante = New ADODB.Recordset
        If rsComprobante.State = 1 Then rsComprobante.Close
        rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo,fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion " & _
                            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE " & cadena_busqueda & " order by pago_detalle.numero_cheque_trf ", db, adOpenKeyset, adLockOptimistic
        If rsComprobante.RecordCount > 0 Then
'        rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo,fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion " & _
'                            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE " & cadena_busqueda & "  ", db, adOpenKeyset, adLockOptimistic
'        If rsComprobante.RecordCount > 0 Then
'            Set rsCmpte = New ADODB.Recordset
          While Not rsComprobante.EOF
              rsCmpte.AddNew
                rsCmpte("codigo_pago") = rsComprobante("codigo_pago")
                rsCmpte("nro_comprobante_anterior") = rsComprobante("nro_comprobante_anterior")
                rsCmpte("org_codigo") = rsComprobante("org_codigo")
                If Not IsNull(rsComprobante("denominacion_beneficiario")) Then rsCmpte("denominacion_beneficiario") = rsComprobante("denominacion_beneficiario")
                If Not IsNull(rsComprobante("fecha_pago")) Then rsCmpte("fecha_pago") = Format(rsComprobante("fecha_pago"), "dd/mm/yyyy")
                rsCmpte("monto_total") = rsComprobante("monto_total")
                rsCmpte("monto_bolivianos") = rsComprobante("monto_bolivianos")
                rsCmpte("monto_dolares") = rsComprobante("monto_dolares")
                rsCmpte("tipo_cambio") = rsComprobante("tipo_cambio")
                rsCmpte("codigo_beneficiario") = rsComprobante("codigo_beneficiario")
                rsCmpte("justificacion") = rsComprobante("justificacion")
                rsCmpte("cheque_o_trf") = rsComprobante("cheque_o_trf")
                rsCmpte("numero_cheque_trf") = rsComprobante("numero_cheque_trf")
                rsCmpte("cta_codigo") = rsComprobante("cta_codigo")
                rsCmpte("bco_descripcion_larga") = rsComprobante("bco_descripcion_larga")
                rsCmpte("literal") = rsComprobante("literal")
                rsCmpte("cta_descripcion_larga") = rsComprobante("cta_descripcion_larga")
                'Freddy
                rsCmpte("estado_pagado") = rsComprobante("estado_pagado")
                rsCmpte("estado_devolucion") = rsComprobante("estado_devolucion")
                
              rsCmpte.Update
              rsComprobante.MoveNext
          Wend
        Else
            Set DtGPagos.DataSource = rsNada
            'Set AdoPagos.Recordset = rsNada
        End If
        
        Set rslsta = New ADODB.Recordset
        queryinicial = "SELECT * FROM to_ListadoComprobantes"
        rslsta.Open queryinicial, db, adOpenKeyset, adLockOptimistic
        If rslsta.RecordCount <= 0 Then
            MsgBox "No existen registros para imprimir", vbInformation + vbCritical, "Validación de datos"
            Exit Sub
        End If
    
        
        RepPagos.Show
    Else
        MsgBox "No existen registros para imprimir", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
    
End Sub

Private Sub CmdRestaurar_Click()
    Set rsComprobante = New ADODB.Recordset
    rsComprobante.Open "SELECT DISTINCT Pagos.codigo_pago,Pagos.nro_comprobante_anterior,Pagos.org_codigo,fc_beneficiario.denominacion_beneficiario,pago_detalle.fecha_pago, pago_detalle.monto_total,pago_detalle.monto_total,pago_detalle.monto_Bolivianos,pago_detalle.monto_dolares,pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga,fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden,pagos.estado_pagado,Pagos.estado_devolucion " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo Where pagos.estado_anulado is NULL", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub DtGPagos_HeadClick(ByVal ColIndex As Integer)
   
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
    Select Case ColIndex
        Case 0
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                       "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
        Case 1
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                       "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.org_codigo", db, adOpenKeyset, adLockOptimistic
        Case 2
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                       "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by fc_beneficiario.denominacion_beneficiario", db, adOpenKeyset, adLockOptimistic
        Case 3
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
        Case 4
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo,fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
        Case 5
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.monto_Bolivianos", db, adOpenKeyset, adLockOptimistic
        Case 6
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_organismo_financiamiento.Org_descripcion, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_beneficiario.denominacion_beneficiario,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.tipo_cambio", db, adOpenKeyset, adLockOptimistic
        Case 7
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.codigo_beneficiario", db, adOpenKeyset, adLockOptimistic
        Case 8
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by fc_beneficiario.denominacion_beneficiario", db, adOpenKeyset, adLockOptimistic
        Case 9
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by Pagos.justificacion", db, adOpenKeyset, adLockOptimistic
        Case 10
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.cheque_o_trf", db, adOpenKeyset, adLockOptimistic
        Case 11
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo,fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.numero_cheque_trf", db, adOpenKeyset, adLockOptimistic
        Case 12
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.cta_codigo", db, adOpenKeyset, adLockOptimistic
        Case 13
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by fc_bancos.Bco_descripcion_larga", db, adOpenKeyset, adLockOptimistic
        Case 14
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.Literal", db, adOpenKeyset, adLockOptimistic
        Case 15
                rsComprobante.Open "SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga, fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden " & _
                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by fc_organismo_financiamiento.Org_descripcion ", db, adOpenKeyset, adLockOptimistic
    End Select
    Set DtGPagos.DataSource = rsComprobante
End Sub

Private Sub Form_Load()

'    Set rslsta = New ADODB.Recordset
'    QueryInicial = "SELECT * FROM to_ListadoComprobantes"
'    rslsta.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'    If rslsta.RecordCount <= 0 Then
'        MsgBox "Busque registros para imprimir", vbInformation + vbCritical, "Validación de datos"
'    End If

    Set rsComprobante = New ADODB.Recordset
'    queryinicial = "SELECT DISTINCT Pagos.ges_Gestion, Pagos.codigo_pago,Pagos.nro_comprobante_anterior,Pagos.org_codigo, fc_beneficiario.denominacion_beneficiario,pago_detalle.fecha_pago,pago_detalle.monto_total,pago_detalle.monto_Bolivianos,pago_detalle.monto_dolares,pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga,fc_organismo_financiamiento.Org_descripcion,Pagos.codigo_solicitud,Pagos.codigo_orden,pagos.estado_pagado,Pagos.estado_devolucion " & _
'                   "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) " & _
'                   "AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo Where Pagos.estado_anulado Is NULL"
                   
    queryinicial = "SELECT DISTINCT dbo.pagos.ges_gestion, dbo.pagos.codigo_pago, dbo.pagos.Nro_Comprobante_Anterior, dbo.pagos.org_codigo, dbo.FC_BENEFICIARIO.denominacion_beneficiario, dbo.pago_detalle.fecha_pago, dbo.pago_detalle.monto_total, dbo.pago_detalle.monto_Bolivianos, dbo.pago_detalle.monto_Dolares, dbo.pago_detalle.tipo_cambio, dbo.pago_detalle.codigo_beneficiario, dbo.pagos.justificacion, dbo.pago_detalle.cheque_o_trf, dbo.pago_detalle.numero_cheque_trf, dbo.pago_detalle.cta_codigo, dbo.fc_bancos.Bco_descripcion_larga, dbo.pago_detalle.literal, dbo.fc_cuenta_bancaria.Cta_descripcion_larga, dbo.fc_organismo_financiamiento.Org_descripcion, " & _
                   "dbo.Pagos.codigo_solicitud , dbo.Pagos.codigo_orden, dbo.Pagos.estado_pagado, dbo.Pagos.estado_devolucion FROM dbo.fc_organismo_financiamiento INNER JOIN dbo.pagos INNER JOIN dbo.pago_detalle ON dbo.pagos.ges_gestion = dbo.pago_detalle.ges_gestion AND dbo.pagos.org_codigo = dbo.pago_detalle.org_codigo AND dbo.pagos.codigo_pago = dbo.pago_detalle.codigo_pago ON dbo.fc_organismo_financiamiento.Org_codigo = dbo.pagos.org_codigo INNER JOIN dbo.FC_BENEFICIARIO ON dbo.pago_detalle.codigo_beneficiario COLLATE Modern_Spanish_CI_AS = dbo.FC_BENEFICIARIO.codigo_beneficiario INNER JOIN dbo.fc_cuenta_bancaria ON dbo.pago_detalle.cta_codigo = dbo.fc_cuenta_bancaria.Cta_codigo INNER JOIN " & _
                   "dbo.fc_bancos ON dbo.fc_cuenta_bancaria.Bco_codigo = dbo.fc_bancos.Bco_codigo Where (dbo.Pagos.estado_anulado Is Null)"
                   
    rsComprobante.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
       Set AdoPagos.Recordset = rsComprobante
       Set DtGPagos.DataSource = AdoPagos.Recordset
    End If
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
    'LstComprobante.RemoveItem ListIndex
End Sub

Private Sub LstCheques_DblClick()
    'LstCheques.RemoveItem ListIndex
End Sub

Private Sub OptCheques_Click()
'    LblTitulo.Caption = "PAGOS REALIZADOS"
'    CmdEntregado.Enabled = True
'    CmdDevuelto.Enabled = True
'    CmdAnulado.Enabled = False
'    CmdCobrado.Enabled = True
'
'    '
'    If rsComprobante.State = 1 Then rsComprobante.Close
'    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
'                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='C'", db, adOpenKeyset, adLockOptimistic
'    If rsComprobante.RecordCount > 0 Then
'        Set DtGPagos.DataSource = rsComprobante
'        Set AdoPagos.Recordset = rsComprobante
'    End If
End Sub

Private Sub OptTransferencias_Click()
    lblTitulo.Caption = "PAGOS REALIZADOS"
    'CmdEntregado.Enabled = False
    'CmdDevuelto.Enabled = False
    'CmdAnulado.Enabled = True
    'CmdCobrado.Enabled = False
    '
    'If rsComprobante.State = 1 Then rsComprobante.Close
    'rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='T' ", db, adOpenKeyset, adLockOptimistic
'    If rsComprobante.RecordCount > 0 Then
'        Set DtGPagos.DataSource = rsComprobante
'        Set AdoPagos.Recordset = rsComprobante
'    End If
'
End Sub



Private Sub Form_Unload(Cancel As Integer)
  Set ClBuscaGrid = Nothing
End Sub


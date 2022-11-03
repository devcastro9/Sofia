VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmChequesNuevo 
   Caption         =   "Impresión de Cheques"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13815
   Icon            =   "FrmChequesNuevo.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryCh 
      Left            =   6165
      Top             =   3375
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintBtn=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Retornar 
      Caption         =   "<<"
      Height          =   750
      Left            =   6150
      TabIndex        =   30
      Top             =   2640
      Width           =   1020
   End
   Begin VB.CommandButton Seleccionar 
      Caption         =   ">>"
      Height          =   750
      Left            =   6150
      TabIndex        =   29
      Top             =   1845
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   13755
      TabIndex        =   9
      Top             =   0
      Width           =   13815
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   60
         TabIndex        =   14
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   13
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   12
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   11
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "IMPRESION  DE CHEQUES"
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
         Left            =   4140
         TabIndex        =   10
         Top             =   135
         Width           =   4065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   1320
      TabIndex        =   26
      Top             =   1050
      Width           =   12315
      Begin VB.Label Label7 
         Caption         =   "Cheques a imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5955
         TabIndex        =   28
         Top             =   240
         Width           =   2445
      End
      Begin VB.Label Label4 
         Caption         =   "Comprobantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   27
         Top             =   255
         Width           =   2445
      End
   End
   Begin MSDataGridLib.DataGrid DtGChequesSeleccionados 
      Height          =   6000
      Left            =   7230
      TabIndex        =   25
      Top             =   1845
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   10583
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
      Height          =   2115
      Left            =   1740
      TabIndex        =   0
      Top             =   3705
      Visible         =   0   'False
      Width           =   2040
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   5
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2047
         TabIndex        =   4
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   225
         TabIndex        =   3
         Top             =   780
         Width           =   1515
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   225
         TabIndex        =   2
         Top             =   1245
         Width           =   1515
      End
      Begin VB.CommandButton CmdSalirBusqueda 
         Caption         =   "Salir"
         Height          =   390
         Left            =   225
         TabIndex        =   1
         Top             =   1650
         Width           =   1515
      End
      Begin VB.Label Label20 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   8
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label19 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   2310
         TabIndex        =   7
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label21 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   450
         TabIndex        =   6
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   6855
      Left            =   15
      TabIndex        =   15
      Top             =   990
      Width           =   1245
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Busqueda"
         Height          =   855
         Left            =   180
         Picture         =   "FrmChequesNuevo.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3735
         Width           =   945
      End
      Begin VB.CommandButton CmdFiltro 
         Caption         =   "Filtro por Cta. Bancaria"
         Height          =   885
         Left            =   180
         TabIndex        =   21
         Top             =   2850
         Width           =   945
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Impresión"
         Height          =   885
         Left            =   180
         Picture         =   "FrmChequesNuevo.frx":0FCC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   195
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   180
         Picture         =   "FrmChequesNuevo.frx":1636
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4950
         Width           =   945
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   885
         Left            =   180
         TabIndex        =   18
         Top             =   1080
         Width           =   945
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar Grid"
         Height          =   885
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1965
         Width           =   945
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cola Imp."
         Height          =   360
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4590
         Width           =   945
      End
   End
   Begin MSDataGridLib.DataGrid DtGCheques 
      Height          =   6000
      Left            =   1290
      TabIndex        =   23
      Top             =   1845
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   10583
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
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   75
      Left            =   10875
      TabIndex        =   24
      Top             =   1725
      Width           =   45
   End
End
Attribute VB_Name = "FrmChequesNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  ADFIN
' Módulo:                   Impresión de cheques en una impresora matricial
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmChequesCuenta
' Descipción :              Impresión de los cheques con numeración automática
'                           de acuerdo a la cuenta bancaria
' Formularios relacionados: Main.frm (Padre)
'                           CryCheque
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creación         20/Ene/ 2001
' Fecha última modificación 20/Mar/ 2001
' Versión:                  2.0
'========================================================================================

Public rsComprobante As New ADODB.Recordset
Dim rsCheque As New ADODB.Recordset
Dim rsCorrel As New ADODB.Recordset

Dim punto As Variant
Dim NumeroCuenta As String

'Dim CryCheq As New CryCheque
Dim rsNada As New ADODB.Recordset


Private Sub CmdBuscar_Click()
Dim condicion As String
                    If TxtCmpte.Text = "" Then
                        MsgBox "Necesita números de comprobante"
                        Exit Sub
                    Else
                        condicion = "pago_detalle.codigo_pago=" + "'" + TxtCmpte.Text + "'"
                    End If
                    If rsComprobante.State Then rsComprobante.Close
                    Set rsComprobante = New ADODB.Recordset
                    rsComprobante.Open "SELECT DISTINCT Pagos.codigo_pago,pagos.org_codigo,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
                    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where  (pago_detalle.estado_aprobacion<>'A' or pago_detalle.estado_aprobacion IS NULL)  and " & condicion & " and pago_detalle.cheque_o_trf='C' " & _
                    "order by Pagos.codigo_pago,pagos.org_codigo,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga", db, adOpenKeyset, adLockOptimistic
'
'                    rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
'                    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A' and " & condicion & " and pago_detalle.cheque_o_trf='C'order by pago_detalle.codigo_pago  ", db, adOpenKeyset, adLockOptimistic
                    '"FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where  " & CONDICION & " order by pago_detalle.codigo_pago  ", db, adOpenKeyset, adLockOptimistic
                    If rsComprobante.RecordCount > 0 Then
                       Set DtGCheques.DataSource = rsComprobante
                       Seleccionar.Enabled = True
                       DtGCheques.Enabled = True
                    Else
                        MsgBox "Puede tratarse de transferencia o no existe el registro porque ya fué aprobado", vbInformation
                        DtGCheques.Enabled = False
                        Seleccionar.Enabled = False
                    End If
                        FraBusca.Visible = False
                        

End Sub

Private Sub CmdBusqueda_Click()
    FraBusca.Visible = True
End Sub

Private Sub CmdFiltro_Click()
'========================================================================================
' Módulo:                   CmdFiltro
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmChequesCuenta
' Descipción :              Se listan todos los registros con un tipo de organismo
'                           financiador
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================
    Dim Resp As String
    Resp = InputBox("Introducir Organismo o Cuenta Bancaria")
    If Resp <> "" Then
'      Set rsCheque = New ADODB.Recordset
'      If rsCheque.State = 1 Then rsCheque.Close
'      'rsCheque.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal,  pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf,  fc_bancos.Bco_descripcion_larga "
'      rsCheque.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
'      "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.cta_codigo= '" & Resp & "' and pago_detalle.estado_aprobacion <> 'A' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
      
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsCheque.Close
    rsComprobante.Open "SELECT Pagos.codigo_pago,pagos.org_codigo,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where  pago_detalle.cheque_o_trf='C'and pago_detalle.cta_codigo= '" & Resp & "' and pago_detalle.estado_aprobacion <> 'A' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
      If rsComprobante.RecordCount > 0 Then
        Set DtGCheques.DataSource = rsComprobante
        DtGCheques.Enabled = True
        DtGCheques.Refresh
      Else
        MsgBox "No existen registros de la cuenta" + " " + Resp
      End If
    End If
    
End Sub
Private Sub CmdImpresionRangos_Click()
'     CmdLimpiar_Click
     If TxtInicio = "" Then
        MsgBox "Introducir comprobante inicial", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
     End If
     If Val(TxtInicio.Text) > Val(TxtFin.Text) Then
        MsgBox "Comprobante inicial menor al comprobante final", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
     End If
     
     'Limpiando la tabla auxiliar para cheques
     Set rscheques = New ADODB.Recordset
     If rscheques.State = 1 Then rscheques.Close
     rscheques.Open "SELECT * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
     While Not rscheques.EOF
         rscheques.Delete
         rscheques.MoveNext
     Wend
     MsgBox "Se imprimirán todos los comprobantes de los que no se emitieron cheques"
     If TxtInicio.Text <> "" And TxtFin.Text <> "" Then
        Set rsComprobante = New ADODB.Recordset
        If rsComprobante.State = 1 Then rsComprobante.Close
        
        '********
        '   rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _

          rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pagos.ges_gestion,pago_detalle.literal,  pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf,  fc_bancos.Bco_descripcion_larga " & _
           "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by pago_detalle.cta_codigo", db, adOpenKeyset, adLockOptimistic
             'Grabando los datos a la tabla auxiliar de cheques
             Set rsCheque = New ADODB.Recordset
             If rsCheque.State = 1 Then rscheques.Close
             rsCheque.Open "SELECT * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
             While Not rsComprobante.EOF
                   If Val(rsComprobante("codigo_pago")) >= Val(TxtInicio.Text) And Val(rsComprobante("codigo_pago")) <= Val(TxtFin.Text) Then
                      rsCheque.AddNew
                      rsCheque("numero_comprobante") = rsComprobante("codigo_pago")
                      rsCheque("monto_bolivianos") = rsComprobante("monto_bolivianos")
                      rsCheque("denominacion_beneficiario") = rsComprobante("denominacion_beneficiario")
                      rsCheque("cta_codigo") = rsComprobante("cta_codigo")
                      If rsComprobante("fecha_pago") <> "" Then
                          dia = Day(rsComprobante("fecha_pago"))
                          mes = Month(rsComprobante("fecha_pago"))
                          anio = Year(rsComprobante("fecha_pago"))
                      Else
                          MsgBox "no existe fecha en uno de los registros"
                          Exit Sub
                      End If
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
                         
                 rsCheque("dia") = dia
                 rsCheque("mes") = mes
                 rsCheque("anio") = anio
                 
                 'rsCheque("literal") = Literal(CStr(rsComprobante("Literal"))) + " BOLIVIANOS"
                 rsCheque("literal") = CStr(rsComprobante("Literal"))
                 Set DtGCheques.DataSource = rsComprobante
                 
                 Select Case rsComprobante("cta_codigo")
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
                 
                 'Abriendo correlativo para hallar el numero de cheque
                 If rsCorrel.State = 1 Then rsCorrel.Close
                 Set rsCorrel = New ADODB.Recordset
                 rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= '" & NumeroCuenta & "' ", db, adOpenKeyset, adLockOptimistic
                 If rsCorrel.RecordCount > 0 Then
                    rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") + 1
                    rsCorrel.Update
                 Else
                    rsCorrel("numero_correlativo") = 0
                    rsCorrel.Update
                 End If
                 'MsgBox "Se imprimirá el Nro. de cheque ....   " & rsCorrel("numero_correlativo"), vbInformation, "Información"
                 LstNroCheque.AddItem rsCorrel("numero_correlativo")
                 rsCheque("numero_cheque") = rsCorrel("numero_correlativo")
                 rsCheque("cod_org") = rsComprobante("org_codigo")
                 rsCheque("ges_gestion") = rsComprobante("ges_gestion")
                      
                 rsCheque.Update
           End If
          
          'MsgBox rsComprobante("codigo_orden")
          rsComprobante.MoveNext
          
        Wend
      End If
           sino = MsgBox("Esta seguro de la asignación de numeros a los cheques, verifique los datos", vbYesNo, "Mensaje de Advertencia")    '  sino = MsgBox("Està seguro de eliminar este registro", vbYesNo + vbQuestion, "Atenciòn") then
           If sino = vbYes Then
                 RepCheque.Show
                 'Ocultando los cheques impresos
                 Cheques_Impresos_rango
                 'Restaurando grid
                 CmdRestaurar_Click
           Else
'                 If rsCorrel.State = 1 Then rsCorrel.Close
'                 Set rsCorrel = New ADODB.Recordset
'                 rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= '" & NumeroCuenta & "' ", db, adOpenKeyset, adLockOptimistic
'                 If rsCorrel.RecordCount > 0 Then
'                    rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") - LstComprobante.ListCount
'                    rsCorrel.Update
'                 Else
'                    rsCorrel("numero_correlativo") = 0
'                    rsCorrel.Update
'                 End If
           Restaurar_numeracion_cheque
           End If
   ' Cheques_Impresos_rango
End Sub

Private Sub cmdImprimir_Click()
    Dim i As Integer
    Dim dia As String
    Dim mes As String
    Dim anio As String
    Dim Fecha As String
       
    'Abriendo tabla de cheques
    Set rsc = New ADODB.Recordset
    rsc.Open "select * from ts_cheque ", db, adOpenKeyset, adLockOptimistic 'where numero_comprobante='" & DtGCheques.Columns(0) & "' and cod_org= '" & DtGCheques.Columns(5) & "' and  ges_gestion= '" & DtGCheques.Columns(6) & "'
    If rsc.RecordCount > 0 Then
       While Not rsc.EOF
         Select Case rsc("cta_codigo")
            Case "70032233"
                  NumeroCuenta = "cta_1"
            Case "11918512"
                  NumeroCuenta = "cta_2"
            Case "70038533"
                  NumeroCuenta = "cta_3"
            Case "74016636"
                  NumeroCuenta = "cta_4"
         End Select
         
         'Abriendo correlativo para hallar el numero de cheque
         If rsCorrel.State = 1 Then rsCorrel.Close
         Set rsCorrel = New ADODB.Recordset
         rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= '" & NumeroCuenta & "' ", db, adOpenKeyset, adLockOptimistic
         If rsCorrel.RecordCount > 0 Then
            rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") + 1
            NUM = rsCorrel("numero_correlativo")
            rsCorrel.Update
         Else
            rsCorrel("numero_correlativo") = 0
            rsCorrel.Update
         End If
         
        'Refrescar
        Set rsCH = New ADODB.Recordset
        If rsCH.State = 1 Then rsCH.Close
        rsCH.Open "SELECT * FROM ts_cheque where numero_comprobante ='" & rsc("numero_comprobante") & "' and cod_org= '" & rsc("Cod_Org") & "' and  ges_gestion= '" & rsc("Ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
        If rsCH.RecordCount > 0 Then
            rsCH("Numero_Cheque") = NUM
            rsCH.Update
        End If
        
        Refrescar
        
      rsc.MoveNext
     Wend
    Else
       MsgBox "No existe Nro. de Cuenta"
       Exit Sub
    End If
   sino = MsgBox("Esta seguro de la asignación de numeros a los cheques, verifique los datos", vbYesNo, "Mensaje de Advertencia")
   If sino = vbYes Then
         NrosCheque_Compte
         'RepCheque.Show
         CryCh.ReportFileName = App.Path & "\FormsTesoreria\Impresion de Cheques\Rpt_Cheques_1.rpt"  ' g- "C:\ADFIN-2001\FormsTesoreria\Impresion de Cheques\Rpt_Cheques_1.rpt"
         iResult = CryCh.PrintReport
         If iResult <> 0 Then
            MsgBox CryCh.LastErrorNumber & " : " & CryCh.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
         Coloca_Status_Impreso
         CmdRestaurar_Click
   Else
        Restaurar_numeracion_cheque
   End If
   SW = 0
   Cola_Impresion
  
End Sub

Private Sub CmdLimpiar_Click()
   db.Execute "DELETE FROM ts_cheque"
   Refrescar
End Sub

Private Sub CmdRestaurar_Click()
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
'    rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal,  pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf,  fc_bancos.Bco_descripcion_larga " & _
'    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'C' and pago_detalle.estado_aprobacion <> 'A'", db, adOpenKeyset, adLockOptimistic
    
    rsComprobante.Open "SELECT DISTINCT Pagos.codigo_pago,pagos.org_codigo,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A' and pago_detalle.cheque_o_trf='C' " & _
    "order by Pagos.codigo_pago,pagos.org_codigo,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga", db, adOpenKeyset, adLockOptimistic
    
    If rsComprobante.RecordCount > 0 Then
        Set DtGCheques.DataSource = rsComprobante
        DtGCheques.Enabled = True
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    'MsgBox LstChequesCodigo.ListCount
    MsgBox LstChequesCodigo.ListIndex
    LstChequesDatos.RemoveItem punto
    'MsgBox LstChequesCodigo.Index(0)
End Sub

Private Sub CmdSalirBusqueda_Click()
    FraBusca.Visible = False
End Sub

Private Sub Command1_Click()
    FrmColaImpresion.Show vbModal
End Sub


Private Sub DtGCheques_HeadClick(ByVal ColIndex As Integer)
    Set rsComprobante = New ADODB.Recordset
    
    If rsComprobante.State = 1 Then rsComprobante.Close
    Select Case ColIndex
        Case 0
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
            "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A'order by  Pagos.codigo_pago", db, adOpenKeyset, adLockOptimistic
        Case 1
            rsComprobante.Open "SELECT Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
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
    Set DtGCheques.DataSource = rsComprobante
End Sub

Private Sub Form_Load()
    
    'Limpiando la tabla de cheques por imprimir
    db.Execute "DELETE FROM ts_cheque"
    
    'Abriendo tabla de comprobantes
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
    'g-- COLOCAMOS EL ESTADO PAGADO
    rsComprobante.Open "SELECT DISTINCT Pagos.codigo_pago,pagos.org_codigo,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
     "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where (pago_detalle.estado_aprobacion <> 'A' or pago_detalle.estado_aprobacion IS NULL) and pago_detalle.cheque_o_trf='C' AND  pagos.estado_pagado = 'S' " & _
     "ORDER BY Pagos.codigo_pago,pagos.org_codigo,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.ges_gestion,pago_detalle.literal, fc_bancos.Bco_descripcion_larga", db, adOpenKeyset, adLockOptimistic
     '"order by pago_detalle.codigo_pago"
     
                       '"FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.estado_aprobacion <> 'A' and pago_detalle.cheque_o_trf='C' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
'                       FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where (pago_detalle.estado_aprobacion <> 'A' or pago_detalle.estado_aprobacion IS NULL) and pago_detalle.cheque_o_trf='C' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGCheques.DataSource = rsComprobante
        DtGCheques.Enabled = True
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
          LstOrg.ListIndex = i
          LstGes.ListIndex = i
          
          NroCheque = LstNroCheque.Text
          
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
                     rspago("estado_aprobacion") = "A"
                     rspago("estado_pagado") = "P" 'Parcial
                     rspago.Update
                    End If
                End If
        
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
           End If
     Next i
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    FrmCP.AdoPago.Refresh
'End Sub




Private Sub LstBeneficiario_DblClick()
    LstBeneficiario.RemoveItem punto
    LstCuenta_DblClick
End Sub



Private Sub LstComprobante_DblClick()
    punto = LstComprobante.ListIndex
    LstComprobante.RemoveItem punto 'ListIndex
    LstMonto_DblClick
End Sub

Private Sub LstCuenta_DblClick()
'   LstCuenta.RemoveItem punto
    LstCuenta.RemoveItem punto
    LstOrg_DblClick
End Sub

Private Sub LstFecha_Click()
'    LstFecha.RemoveItem punto
'    LstBeneficiario_Click
End Sub

Private Sub LstFecha_DblClick()
    LstFecha.RemoveItem punto
    LstBeneficiario_DblClick
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

'========================================================================================
' Módulo:                   NrosCheque
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmChequesCuenta
' Descipción :              Se colan el formato de ##### de los cheques(con 5 digitos)
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================
Dim NumeroCheque As String


Set rsCheque = New ADODB.Recordset
If rsCheque.State = 1 Then rsCheque.Close
rsCheque.Open "select * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
If rsCheque.RecordCount > 0 Then
        While Not rsCheque.EOF
            Set rsPagoDet = New ADODB.Recordset
            If rsPagoDet.State = 1 Then rsPagoDet.Close
            rsPagoDet.Open "select * from pago_detalle where codigo_pago='" & rsCheque("numero_comprobante") & "' and ges_gestion='" & rsCheque("ges_gestion") & "' and org_codigo='" & rsCheque("cod_org") & "' ", db, adOpenKeyset, adLockOptimistic
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
                If NumeroCheque <> "" Then
                    rsPagoDet("numero_cheque_trf") = NumeroCheque
                    rsPagoDet("estado_aprobacion") = "A"
                    rsPagoDet("fecha_impresion_cheque") = Date
                    rsPagoDet("cheque_o_trf") = "C"
                    rsPagoDet.Update
                End If
            rsCheque.MoveNext
        Wend
End If
End Sub

Private Sub LstGes_DblClick()
  LstGes.RemoveItem punto
End Sub

Private Sub LstMonto_DblClick()
    LstMonto.RemoveItem punto
    LstFecha_DblClick
End Sub

Private Sub LstOrg_DblClick()
    LstOrg.RemoveItem punto
    LstGes_DblClick
End Sub


Public Sub Cheques_Impreso()
'========================================================================================
' Módulo:                   Cobrado_Lista
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmActivacionCheques.frm
' Descipción :              Se coloca el status de cobrado
'                           de acuerdo a una lista y en el caso de cheques
'                           de acuerdo a la cuenta bancaria
'                           si se trata de cheques
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================

Dim i As Integer
  If LstNroCheque.Text <> "" Then
   Set rscheques = New ADODB.Recordset
   If rscheques.State = 1 Then rscheques.Close
   For i = 0 To LstNroCheque.ListCount - 1
            LstNroCheque.ListIndex = i
            rscheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstNroCheque.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
            If rscheques.RecordCount > 0 Then
                   rscheques("estado_impreso") = "S"
            Else
                rscheques.AddNew
                rscheques("numero_cheque") = LstNroCheque.Text
                rscheques("estado_impreso") = "S"
            End If
            rscheques("usr_usuario") = Lblusuario.Caption
            rscheques("fecha_registro") = Date
            rscheques("hora_registro") = Format(Time, "hh:mm:ss")
            rscheques.Update
    Next i
  End If
End Sub
Public Sub Cola_Impresion()
'========================================================================================
' Módulo:                   Cola_Impresión
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmChequesCuenta.frm
' Descipción :              Se recuperan los datos de los cheques y las
'                           transferencias que se imprimieron
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================
    
    Dim SqlQuery As String
    'Mandando a la cola de impresión los cheques
    
     Set rsIC = New ADODB.Recordset
     If rsIC.State = 1 Then rsTransferencia.Close
     rsIC.Open "SELECT * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
     If rsIC.RecordCount > 0 Then
     While Not rsIC.EOF
            Set rsComprobante = New ADODB.Recordset
            SqlQuery = " SELECT Pagos.codigo_pago, fc_cuenta_bancaria.cta_descripcion_larga, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion,  pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, pago_detalle.cta_codigo, pago_detalle.cheque_o_trf " & _
                       "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE pago_detalle.codigo_pago='" & rsIC("numero_comprobante") & "' and pago_detalle.Ges_gestion= '" & rsIC("ges_gestion") & "' and pago_detalle.cta_codigo='" & rsIC("cta_codigo") & "' order by Pago_detalle.codigo_pago"
            rsComprobante.Open SqlQuery, db, adOpenKeyset, adLockOptimistic
            If rsComprobante.RecordCount > 0 Then
                 Set rsCmpteI = New ADODB.Recordset
                 If rsCmpteI.State = 1 Then rsCmpteI.Close
                 rsCmpteI.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
                 
                        rsCmpteI.AddNew
                        If Not IsNull(rsComprobante("codigo_pago")) Then rsCmpteI("Nro_Cmpte") = rsComprobante("codigo_pago")
                        If Not IsNull(rsComprobante("cta_descripcion_larga")) Then rsCmpteI("Organismo") = rsComprobante("cta_descripcion_larga")
                        If Not IsNull(rsComprobante("Fecha_pago")) Then rsCmpteI("Fecha_Pago") = Format(rsComprobante("Fecha_pago"), "dd/mm/yyyy")
                        If Not IsNull(rsComprobante("monto_bolivianos")) Then rsCmpteI("Monto") = rsComprobante("monto_bolivianos")
                        If Not IsNull(rsComprobante("tipo_cambio")) Then rsCmpteI("Cambio") = rsComprobante("tipo_cambio")
                        If Not IsNull(rsComprobante("denominacion_beneficiario")) Then rsCmpteI("Beneficiario") = rsComprobante("denominacion_beneficiario")
                        If Not IsNull(rsComprobante("Justificacion")) Then rsCmpteI("Justificacion") = rsComprobante("Justificacion")
                        If Not IsNull(rsComprobante("numero_cheque_trf")) Then rsCmpteI("Nro_cheque") = rsComprobante("numero_cheque_trf")
                        If Not IsNull(rsComprobante("Bco_descripcion_larga")) Then rsCmpteI("banco") = rsComprobante("Bco_descripcion_larga")
                        rsCmpteI("Transf_cheq") = "CHEQUE"
                        rsCmpteI("Literal") = Literal(rsComprobante("monto_bolivianos"))
                    rsCmpteI.Update

            End If
            rsIC.MoveNext
      Wend
      End If
End Sub


Private Sub Retornar_Click()
'    Set rsCheq = New ADODB.Recordset
'    rsCheq.Open "select * from ts_cheque", db, adOpenKeyset, adLockOptimistic
'    If rsCheq.RecordCount > 0 Then
'       Set DtGChequesSeleccionados.DataSource = rsCheq
'    Else
'       Set DtGChequesSeleccionados.DataSource = rsNada
'       Exit Sub
'    End If
    
    db.Execute "DELETE FROM ts_cheque where numero_comprobante='" & DtGChequesSeleccionados.Columns(1) & "' and cod_org= '" & DtGChequesSeleccionados.Columns(9) & "' and  ges_gestion= '" & DtGChequesSeleccionados.Columns(10) & "'"
    Set rsCheq = New ADODB.Recordset
    rsCheq.Open "select * from ts_cheque", db, adOpenKeyset, adLockOptimistic
    If rsCheq.RecordCount > 0 Then
       Set DtGChequesSeleccionados.DataSource = rsCheq
    Else
       Set DtGChequesSeleccionados.DataSource = rsNada
       DtGChequesSeleccionados.Enabled = False
       Retornar.Enabled = False
    End If
End Sub

Private Sub Seleccionar_Click()
 Dim bandera As Integer
               Fecha = Date
                  dia = Day(Fecha)
                  mes = Month(Fecha)
                  anio = Year(Fecha)
    
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
    
 
    'Ingresando datos a ts_cheque
    If DtGCheques.Columns(0) = "" Then
        MsgBox "No existe Nro de comprobante", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
    If DtGCheques.Columns(1) = "" Then
        MsgBox "No existe Monto Bolivianos", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
   If DtGCheques.Columns(5) = "" Then
        MsgBox "No existe Organismo", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    Set rsc = New ADODB.Recordset
    rsc.Open "select * from ts_cheque where numero_comprobante='" & DtGCheques.Columns(0) & "' and cod_org= '" & DtGCheques.Columns(1) & "' and  ges_gestion= '" & DtGCheques.Columns(6) & "'", db, adOpenKeyset, adLockOptimistic
    If rsc.RecordCount = 0 Then
        db.Execute "insert into ts_Cheque (numero_comprobante, cod_org, monto_bolivianos, dia, mes, anio, denominacion_beneficiario, numero_cheque, cta_codigo, ges_gestion, literal)" & _
                  "values ('" & DtGCheques.Columns(0) & "','" & DtGCheques.Columns(1) & "'," & DtGCheques.Columns(2) & ", " & dia & ",'" & mes & "'," & anio & ",'" & DtGCheques.Columns(4) & "','','" & DtGCheques.Columns(5) & "','" & DtGCheques.Columns(6) & "', '" & Literal(DtGCheques.Columns(2)) & "'  ) "
    
    End If
    Set rsCheq = New ADODB.Recordset
    rsCheq.Open "select * from ts_cheque", db, adOpenKeyset, adLockOptimistic
    If rsCheq.RecordCount > 0 Then
       Set DtGChequesSeleccionados.DataSource = rsCheq
       DtGChequesSeleccionados.Enabled = True
       Retornar.Enabled = True
    End If

End Sub

Public Sub Coloca_Status_Impreso()
    Set rsCheque = New ADODB.Recordset
    rsCheque.Open "select * from ts_cheque", db, adOpenKeyset, adLockOptimistic
    If rsCheque.RecordCount > 0 Then
       While Not rsCheque.EOF
            Set rsOP = New ADODB.Recordset
            rsOP.Open "select * from to_cheques_Operaciones WHERE numero_cheque='" & rsCheque("numero_cheque") & "' and cta_codigo='" & rsCheque("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
            If rsOP.RecordCount > 0 Then
                rsOP("estado_impreso") = "S"
                rsOP("Fecha_impreso") = Date
            Else
                rsOP.AddNew
                rsOP("numero_cheque") = Mid(CStr(100000 + Val(rsCheque("numero_cheque"))), 2, 5)
                rsOP("cta_codigo") = rsCheque("cta_codigo")
                rsOP("estado_impreso") = "S"
                rsOP("estado_entregado") = "N"
                rsOP("estado_anulado") = "N"
                rsOP("estado_cobrado") = "N"
                rsOP("estado_devuelto") = "N"
                rsOP("fecha_registro") = Date
                rsOP("Fecha_impreso") = Date
                rsOP("Cheq_Transf") = "C"
                
                'rsOP("hora_registro") = Str(Time)
                rsOP.Update
            End If
            'db.Execute "UPDATE to_cheques_Operaciones WHERE numero_cheque='" & rsCheque("numero_cheque") & "' and cta_codigo='" & rsCheque("cta_codigo") & "'"
            rsCheque.MoveNext
        Wend
            
    End If

End Sub

Public Sub Refrescar()
       'Refrescar
        Set rsCH = New ADODB.Recordset
        rsCH.Open "SELECT * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
        If rsCH.RecordCount > 0 Then
            Set DtGChequesSeleccionados.DataSource = rsCH
        Else
            Set DtGChequesSeleccionados.DataSource = rsNada
        End If
End Sub

Private Sub TxtCmpte_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

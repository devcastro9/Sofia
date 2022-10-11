VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmTransferenciasNuevo 
   Caption         =   "Impresión de Transferencias"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   Icon            =   "FrmTransferenciasNuevo.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8205
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryTr 
      Left            =   6330
      Top             =   4005
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   1305
      TabIndex        =   32
      Top             =   1080
      Width           =   11160
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
         Left            =   150
         TabIndex        =   34
         Top             =   255
         Width           =   2445
      End
      Begin VB.Label Label7 
         Caption         =   "Transferencias a imprimir"
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
         TabIndex        =   33
         Top             =   240
         Width           =   3150
      End
   End
   Begin VB.CommandButton Seleccionar 
      Caption         =   ">>"
      Height          =   855
      Left            =   6195
      TabIndex        =   31
      Top             =   1860
      Width           =   1020
   End
   Begin VB.CommandButton Retornar 
      Caption         =   "<<"
      Height          =   855
      Left            =   6195
      TabIndex        =   30
      Top             =   2715
      Width           =   1020
   End
   Begin MSDataGridLib.DataGrid DtGTransferenciasImprimir 
      Height          =   6570
      Left            =   7305
      TabIndex        =   29
      Top             =   1800
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   11589
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
   Begin VB.Frame FraOpciones 
      Height          =   7275
      Left            =   0
      TabIndex        =   17
      Top             =   1065
      Width           =   1245
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   585
         TabIndex        =   18
         Top             =   885
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   688
         Appearance      =   1
         _Version        =   327682
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar Grid"
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1335
         Width           =   915
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   735
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2070
         Width           =   930
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   735
         Left            =   105
         Picture         =   "FrmTransferenciasNuevo.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5385
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   120
         Picture         =   "FrmTransferenciasNuevo.frx":130C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   915
      End
      Begin VB.CommandButton CmdFiltro 
         Caption         =   "Filtro por Organismo"
         Height          =   735
         Left            =   105
         TabIndex        =   22
         Top             =   2805
         Width           =   930
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Busqueda"
         Height          =   735
         Left            =   105
         Picture         =   "FrmTransferenciasNuevo.frx":1976
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3555
         Width           =   930
      End
      Begin VB.CommandButton CmdReimpresion 
         Caption         =   "Reimprimir"
         Height          =   735
         Left            =   105
         Picture         =   "FrmTransferenciasNuevo.frx":1A78
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4290
         Width           =   930
      End
      Begin VB.CommandButton CmdColaImpresion 
         Caption         =   "Cola Impr."
         Height          =   360
         Left            =   120
         TabIndex        =   19
         Top             =   5025
         Width           =   915
      End
   End
   Begin VB.OptionButton OptBol 
      Caption         =   "Bolivianos"
      Height          =   270
      Left            =   1380
      TabIndex        =   16
      Top             =   1095
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.OptionButton OptDol 
      Caption         =   "Dólares"
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2505
      TabIndex        =   15
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Frame FraBusca 
      Height          =   2085
      Left            =   2295
      TabIndex        =   6
      Top             =   3660
      Visible         =   0   'False
      Width           =   2040
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   225
         TabIndex        =   11
         Top             =   1095
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   225
         TabIndex        =   10
         Top             =   645
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2047
         TabIndex        =   9
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   8
         Top             =   915
         Width           =   1515
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   225
         TabIndex        =   7
         Top             =   1485
         Width           =   1515
      End
      Begin VB.Label Label22 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   450
         TabIndex        =   14
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   2310
         TabIndex        =   13
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label24 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   12
         Top             =   645
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   10260
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      Begin VB.Label Label2 
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
         TabIndex        =   5
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   2
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   2790
         TabIndex        =   1
         Top             =   225
         Width           =   7095
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   0
         Picture         =   "FrmTransferenciasNuevo.frx":20E2
         Top             =   0
         Width           =   15360
      End
   End
   Begin MSDataGridLib.DataGrid DtgTransferencias 
      Height          =   6570
      Left            =   1320
      TabIndex        =   26
      Top             =   1785
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   11589
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
   Begin Crystal.CrystalReport CryCh 
      Left            =   4875
      Top             =   4065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label13 
      Caption         =   "Ciudad"
      Height          =   180
      Left            =   15
      TabIndex        =   28
      Top             =   0
      Width           =   915
   End
   Begin VB.Label Label21 
      Caption         =   "ges"
      Height          =   405
      Left            =   15255
      TabIndex        =   27
      Top             =   4335
      Width           =   270
   End
End
Attribute VB_Name = "FrmTransferenciasNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'========================================================================================
' Sistema:                  Atencion 2002
' Módulo:                   Control de Impresión de Transferencias
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmTransferencia.frm
' Descipción :              Comprobantes pagados
' Formularios relacionados: Main.frm (Padre)
'                           CryTransferencia
' Autor:
' Fecha de creación
' Fecha última modificación
' Versión:                  2.0
'========================================================================================

Dim rsTransferencia As New ADODB.Recordset
Dim rsCorrel As New ADODB.Recordset
Dim rsTransfAux As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim rsCta As New ADODB.Recordset
Dim punto As Integer
Public Sub Restaurar_Numeracion_Transferencia()
   Set rsTransferencia = New ADODB.Recordset
   If rsTransferencia.State Then rsTransferencia.Close
   rsTransferencia.Open "SELECT * FROM to_transferencia", db, adOpenKeyset, adLockOptimistic
   If rsTransferencia.RecordCount > 0 Then
      While Not rsTransferencia.EOF
         Set rsCorrel = New ADODB.Recordset
         If rsCorrel.State = 1 Then rsCorrel.Close
         rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= 'transf ' ", db, adOpenKeyset, adLockOptimistic
         If rsCorrel.RecordCount > 0 Then
            rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") - 1
            rsCorrel.Update
         Else
            rsCorrel("numero_correlativo") = 0
            rsCorrel.Update
         End If
         rsTransferencia.MoveNext
     Wend
    End If
    Refrescar
    
    'Actualizando la vista de selecciones
    Set rsTransferencia = New ADODB.Recordset
    rsTransferencia.Open "SELECT Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
                         "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
    If rsTransferencia.RecordCount > 0 Then
        Set DtgTransferencias.DataSource = rsTransferencia
    Else
        MsgBox "No existen registros", vbInformation + vbCritical, "Validación de datos"
        Set DtgTransferencias.DataSource = rsTransferencia
        Exit Sub
    End If
End Sub



Private Sub CmdBuscar_Click()
Dim condicion As String
                    If TxtCmpte.Text = "" Then
                        MsgBox "Necesita números de comprobante"
                        Exit Sub
                    Else
                        condicion = "pago_detalle.codigo_pago=" + "'" + TxtCmpte.Text + "'"
                    End If
                    Set rsTransferencia = New ADODB.Recordset
'                     rsTransferencia.Open "SELECT pago_detalle.fecha_pago, Pagos.codigo_pago, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.monto_bolivianos, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, pago_detalle.org_codigo, pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
'                    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' and " & condicion & " order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
                    rsTransferencia.Open "SELECT DISTINCT Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
                         "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' and (" & condicion & ") " & _
                         "order by Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino", db, adOpenKeyset, adLockOptimistic

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

Private Sub CmdColaImpresion_Click()
    FrmColaImpresion.Show
End Sub

Private Sub cmdFiltro_Click()
    Dim Resp As String
    Resp = InputBox("Introducir Organismo")
    If Resp <> "" Then
    Set rsTransferencia = New ADODB.Recordset
    If rsTransferencia.State = 1 Then rsTransferencia.Close
'    rsTransferencia.Open "SELECT Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
'                         "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.org_codigo='" & Resp & "' and pago_detalle.cheque_o_trf= 'T' order by pago_detalle.codigo_pago ", db, adOpenKeyset, adLockOptimistic
                         
    rsTransferencia.Open "SELECT Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
                         "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.org_codigo='" & Resp & "'  and pago_detalle.cheque_o_trf= 'T' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
    If rsTransferencia.RecordCount > 0 Then
        Set DtgTransferencias.DataSource = rsTransferencia
    Else
        MsgBox "No existen transferencias o los datos son incoherentes", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    End If
    
End Sub

Private Sub Cmdimprimir_Click()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim dia As String
    Dim mes As String
    Dim anio As String
    Dim Fecha As String
    
    
    'Verificar si se trata de impresión o reimpresión
    Set rsT = New ADODB.Recordset
    rsT.Open "select * from to_transferencia ", db, adOpenKeyset, adLockOptimistic
    If rsT.RecordCount > 0 Then
        While Not rsT.EOF
          If rsT("Nro_Transferencia") <> "" Then
            MsgBox "Reimprima las transferencias que ya tienen numeración", vbCritical + vbDefaultButton1
            db.Execute "delete from to_transferencia"
            Refrescar
            Exit Sub
          End If
          rsT.MoveNext
        Wend
    End If
               
    'Abriendo tabla de transferencias
    Set rsT = New ADODB.Recordset
    rsT.Open "select * from to_transferencia", db, adOpenKeyset, adLockOptimistic
    If rsT.RecordCount > 0 Then
    NUM = rsT.RecordCount
    While Not rsT.EOF
            'Buscando Nro. de correlativo de Transferencia
            If rsCorrel.State = 1 Then rsCorrel.Close
            Set rsCorrel = New ADODB.Recordset
            rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= 'Transf' ", db, adOpenKeyset, adLockOptimistic
            If rsCorrel.RecordCount > 0 Then
                   rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") + 1
                   NUM = rsCorrel("numero_correlativo")
                   rsCorrel.Update
            Else
                   rsCorrel("numero_correlativo") = 0
                   rsCorrel.Update
            End If
            
            'Refrescando los datos de to_transferencia
            Set rsTr = New ADODB.Recordset
            rsTr.Open "select * from to_transferencia where Nro_Cmpte='" & rsT("Nro_Cmpte") & "' and Cod_Org='" & rsT("Cod_org") & "' and ges_gestion='" & rsT("ges_gestion") & "'", db, adOpenKeyset, adLockOptimistic
            If rsTr.RecordCount > 0 Then
                rsTr("Nro_Transferencia") = NUM
                rsTr.Update
            End If
            Refrescar
            rsT.MoveNext
    Wend
    Else
        MsgBox "Elija registros para imprimir", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    sino = MsgBox("Se imprimiran con Nro.(s) Transf ...!", vbYesNo, "Mensaje de Advertencia")
    If sino = vbYes Then
         Cmpte_NroTransferencia
         
         ' Adiciona cta_codigo_tgn cuando son traspasos TRP ... (Jorge)
         'MsgBox rst("cta_destino")
         Dim Cta_tgn_destino As String
         Dim rsctabco As New ADODB.Recordset
         rsctabco.CursorLocation = adUseClient
         'Set rsctabco = New ADODB.Recordset
         If rsT.State = 1 Then rsT.Close
         rsT.Open "select * from to_transferencia", db, adOpenKeyset, adLockOptimistic
         If rsT.RecordCount > 0 Then
            rsT.MoveFirst
            Do While Not rsT.EOF
              Set rsctabco = New ADODB.Recordset
              
              rsctabco.Open "select * from fc_cuenta_bancaria where cta_codigo='" & rsT("cta_destino") & "' ", db, adOpenKeyset, adLockReadOnly
              If rsctabco.RecordCount > 0 Then
                rsT!Cta_tgn_destino = IIf(IsNull(rsctabco!cta_codigo_tgn), "", rsctabco!cta_codigo_tgn)
              Else
                rsT!Cta_tgn_destino = ""
              End If
              rsT.Update
              rsT.MoveNext
            Loop
          End If
         'g--
'         If rsctabco.RecordCount > 0 Then
'            CryTr.Formulas(0) = "Cta_tgn_destino = '" & rsctabco("cta_codigo_tgn") & "' "
'         Else
'            CryTr.Formulas(0) = "Cta_tgn_destino = '" & "." & "' "
'         End If
'         ' Fin Adiciona cta_codigo_tgn cuando son traspasos TRP ...
         
         CryTr.ReportFileName = App.Path & "\FormsTesoreria\Impresion Transferencias\Rpt_transferencia.rpt"
         IResult = CryTr.PrintReport
         If IResult <> 0 Then
            MsgBox CryTr.LastErrorNumber & " : " & CryTr.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
         Coloca_Status_Impreso
         CmdRestaurar_Click
   Else
         Restaurar_Numeracion_Transferencia
         Exit Sub
   End If
   SW = 0
   Cola_Impresion
   
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
    db.Execute "DELETE FROM to_transferencia"
    Refrescar
End Sub
Private Sub CmdReimpresion_Click()

         'Verificar si se trata de impresión o reimpresión
          Set rsT = New ADODB.Recordset
          rsT.Open "select * from to_transferencia ", db, adOpenKeyset, adLockOptimistic
          If rsT.RecordCount > 0 Then
               While Not rsT.EOF
                 If rsT("Nro_Transferencia") = "" Or IsNull(rsT("Nro_Transferencia")) Then
                   'MsgBox "Imprima las transferencias que no tienen numeración", vbCritical + vbDefaultButton1
                   MsgBox "Elegir botón de imprimir para los comprobantes que no tienen numeración ", vbCritical + vbDefaultButton1
                   db.Execute "delete from to_transferencia"
                   Refrescar
                   Exit Sub
                 End If
                 rsT.MoveNext
                Wend
           End If

         'Abriendo tabla de transferencias
         Set rsT = New ADODB.Recordset
         If rsT.State = 1 Then rsT.Close
         rsT.Open "select * from to_transferencia", db, adOpenKeyset, adLockOptimistic
         If rsT.RecordCount <= 0 Then
                MsgBox "Elija registros para imprimir", vbInformation + vbCritical, "Validación de datos"
                Exit Sub
         End If
         
         ' Adiciona cta_codigo_tgn cuando son traspasos TRP ... (Jorge)
         'MsgBox rst("cta_destino")
         Dim Cta_tgn_destino As String
         Dim rsctabco As New ADODB.Recordset
         rsT.MoveFirst
         If rsT.RecordCount > 0 Then
            Do While Not rsT.EOF
              Set rsctabco = New ADODB.Recordset
              rsctabco.CursorLocation = adUseClient
              rsctabco.Open "select * from fc_cuenta_bancaria where cta_codigo='" & rsT("cta_destino") & "' ", db, adOpenKeyset, adLockReadOnly
              If rsctabco.RecordCount > 0 Then
                rsT!Cta_tgn_destino = IIf(IsNull(rsctabco!cta_codigo_tgn), "", rsctabco!cta_codigo_tgn)
              Else
                rsT!Cta_tgn_destino = ""
              End If
              
              rsT.Update
              rsT.MoveNext
            Loop
          End If
          'g--
'          If rsctabco.RecordCount > 0 Then
'            CryTr.Formulas(0) = "Cta_tgn_destino = '" & rsctabco("cta_codigo_tgn") & "' "
'          Else
'            CryTr.Formulas(0) = "Cta_tgn_destino = '" & "." & "' "
'          End If
         ' Fin Adiciona cta_codigo_tgn cuando son traspasos TRP ...
         
         'CryTr.Formulas(0) = "For_Copia='Copia'"
         CryTr.ReportFileName = App.Path & "\FormsTesoreria\Impresion Transferencias\Rpt_transferencia.rpt"
         IResult = CryTr.PrintReport
         If IResult <> 0 Then
            MsgBox CryTr.LastErrorNumber & " : " & CryTr.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
         
End Sub

Private Sub CmdRestaurar_Click()
    If rsTransferencia.State = 1 Then rsTransferencia.Close
    Set rsTransferencia = New ADODB.Recordset
    rsTransferencia.Open "SELECT DISTINCT Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino,  pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
    "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' " & _
    "order by Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino,  pago_detalle.ges_gestion, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino", db, adOpenKeyset, adLockOptimistic
    Set DtgTransferencias.DataSource = rsTransferencia
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub



Private Sub DtgTransferencias_Click()
' Dim bandera As Integer
' Dim z As Integer
'    bandera = 0
'    z = 0
'    For i = 0 To LstComprobante.ListCount - 1
'         LstComprobante.ListIndex = i
'         If LstComprobante.Text = DtgTransferencias.Columns(1) Then
'              bandera = 1
'         End If
'    Next i
'    If bandera = 0 Then
'        LstComprobante.AddItem DtgTransferencias.Columns(1)
'        LstFecha.AddItem DtgTransferencias.Columns(0)
'        LstTransf.AddItem DtgTransferencias.Columns(2)
'        LstCuentaOrigen.AddItem DtgTransferencias.Columns(3)
'        LstBanco.AddItem DtgTransferencias.Columns(4)
'        LstDesCuenta.AddItem DtgTransferencias.Columns(5)
'        LstMontoBol.AddItem DtgTransferencias.Columns(6)
'        LstLiteral.AddItem DtgTransferencias.Columns(7)
'        LstDepto.AddItem DtgTransferencias.Columns(8)
'        LstJustificacion.AddItem DtgTransferencias.Columns(9)
'        LstCuentaDes.AddItem DtgTransferencias.Columns(10)
'        LstDolares.AddItem DtgTransferencias.Columns(12)
'        LstObs.AddItem DtgTransferencias.Columns(13)
'        LstBancoDestino.AddItem DtgTransferencias.Columns(14)
'        LstOrg.AddItem DtgTransferencias.Columns(15)
'        LstGes.AddItem DtgTransferencias.Columns(16)
'        LstRep.AddItem DtgTransferencias.Columns(17)
'        LstCar.AddItem DtgTransferencias.Columns(18)
'        LstHono.AddItem DtgTransferencias.Columns(19)
'        LstBDestino.AddItem DtgTransferencias.Columns(20)
'    End If
End Sub

Private Sub Form_Load()
    'Borrando la tabla auxiliar de transferencias
    db.Execute "DELETE  FROM to_transferencia"
    
    Set rsTransferencia = New ADODB.Recordset
    rsTransferencia.Open "SELECT DISTINCT Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
                         "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' " & _
                         "order by Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino", db, adOpenKeyset, adLockOptimistic
    If rsTransferencia.RecordCount > 0 Then
        Set DtgTransferencias.DataSource = rsTransferencia
    Else
        MsgBox "No existen registros", vbInformation + vbCritical, "Validación de datos"
        Set DtgTransferencias.DataSource = rsTransferencia
        Exit Sub
    End If
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
                If NumeroTransferencia <> "" Then
                        rsPagoDet("numero_cheque_trf") = NumeroTransferencia
                        rsPagoDet("cheque_o_trf") = "T"
                        rsPagoDet("estado_aprobacion") = "A"
                        rsPagoDet("fecha_impresion_cheque") = Date
                        rsPagoDet.Update
                End If
            rsTransfAux.MoveNext
        Wend
End If
End Sub

Private Sub Retornar_Click()

    Set rsT = New ADODB.Recordset
    rsT.Open "select * from to_transferencia", db, adOpenKeyset, adLockOptimistic
    If rsT.RecordCount > 0 Then
    db.Execute "DELETE FROM to_transferencia where Nro_Cmpte='" & DtGTransferenciasImprimir.Columns(1) & "' and cod_org= '" & DtGTransferenciasImprimir.Columns(2) & "' and  ges_gestion= '" & DtGTransferenciasImprimir.Columns(15) & "'"
    End If
    Refrescar
    
End Sub

Private Sub Seleccionar_Click()
    Dim bandera As Integer
    Dim rsb As New ADODB.Recordset
    FrmOpciones.Show vbModal

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
    If DtgTransferencias.Columns(0) = "" Then
        MsgBox "No existe Nro de comprobante", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
    If DtgTransferencias.Columns(1) = "" Then
        MsgBox "No existe Organismo", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    
   If DtgTransferencias.Columns(2) = "" Then
        MsgBox "No existe Monto", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    Set rsT = New ADODB.Recordset
    rsT.Open "select * from to_transferencia where Nro_Cmpte='" & DtgTransferencias.Columns(0) & "' and cod_org= '" & DtgTransferencias.Columns(1) & "' and  ges_gestion= '" & DtgTransferencias.Columns(4) & "'", db, adOpenKeyset, adLockOptimistic
    If rsT.RecordCount = 0 Then
        
        rsT.AddNew
'       db.Execute "insert into to_transferencia(Nro_Transferencia, Nro_Cmpte, Fecha_Pago, cta_origen, Banco, Cta_origen_descripcion, Monto)" & _
'                  "values (0,'" & rsTransferencia("codigo_pago") & "','" & rsTransferencia("fecha_pago") & "', '" & rsTransferencia("cta_codigo") & "','" & Trim(rsTransferencia("Bco_descripcion_larga")) & "','" & rsTransferencia("Cta_descripcion_larga") & "','" & rsTransferencia("Cta_descripcion_larga") & "', " & rsTransferencia("monto_bolivianos") & ""
       rsT!Nro_Transferencia = rsTransferencia("numero_cheque_trf")
       rsT!Nro_Cmpte = rsTransferencia("Codigo_Pago")
       rsT!fecha_pago = CDate(rsTransferencia("fecha_pago"))
       rsT!cta_origen = rsTransferencia("cta_codigo")
       rsT!Cta_origen_descripcion = rsTransferencia("Cta_descripcion_larga")
       rsT!Monto = rsTransferencia("monto_bolivianos")
       rsT!cta_destino = rsTransferencia("cta_codigo_destino")
       rsT!departamento = rsTransferencia("departamento")
       rsT!justificacion = rsTransferencia("justificacion")
       rsT!Banco = rsTransferencia("Bco_descripcion_larga")
       rsT!dia = dia
       rsT!mes = mes
       rsT!anio = anio
       If moneda = "1" Then
            rsT!Monto = rsTransferencia("monto_bolivianos")
            rsT!Literal = Literal(rsTransferencia("monto_bolivianos")) + d + " BOLIVIANOS"
            rsT!moneda = "Bs."
       End If
       
       If moneda = "2" Then
            rsT!Monto = rsTransferencia("monto_dolares")
            rsT!Literal = Literal(rsTransferencia("monto_dolares")) + d + " DOLARES"
            rsT!moneda = "$us"
            
       End If
       rsT!banco_destino = rsTransferencia("banco_destino")
       rsT!Obs = rsTransferencia("observacion")
       rsT!cod_org = rsTransferencia("org_codigo")
       rsT!ges_gestion = rsTransferencia("ges_gestion")
       
       If rsTransferencia("honorarios") = "H" Then
            rsT!honorarios = "Honorarios"
       End If
       If rsTransferencia("honorarios") = "S" Then
            rsT!honorarios = ""
       End If
       
       
       'Cta tgn
       Set rsCta = New ADODB.Recordset
       If rsCta.State = 1 Then rsCta.Close
       rsCta.Open "select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsTransferencia("Cta_codigo")) & "'", db, adOpenKeyset, adLockOptimistic
       If rsCta.RecordCount > 0 Then
         rsT!Cta_tgn = rsCta("Cta_codigo_tgn")
           Set rsb = New ADODB.Recordset
           If rsb.State = 1 Then rsb.Close
           rsb.Open "select * from fc_bancos where Bco_Codigo= '" & Trim(rsCta("Bco_codigo")) & "'", db, adOpenKeyset, adLockOptimistic
           If rsb.RecordCount > 0 Then
                rsT!Beneficiario_destino = rsTransferencia("beneficiario_destino")
                rsT!Representante = rsb("representante")
                rsT!Cargo = rsb("cargo")
           End If
       End If
      'cuenta codigo rsT!cta_origen = rsTransferencia("cta_codigo")
   
       
            '      rsTransferencia.Open "SELET Pagos.codigo_pago,pago_detalle.org_codigo,pago_detalle.monto_bolivianos,pago_detalle.cta_codigo,pago_detalle.ges_gestion,pago_detalle.fecha_pago, pago_detalle.numero_cheque_trf, fc_bancos.Bco_descripcion_larga, fc_cuenta_bancaria.Cta_descripcion_larga, pago_detalle.literal, pago_detalle.departamento, Pagos.justificacion, pago_detalle.cta_codigo_destino, pago_detalle.cheque_o_trf, pago_detalle.monto_dolares, pago_detalle.observacion, pago_detalle.banco_destino, fc_bancos.representante, fc_bancos.cargo, pago_detalle.honorarios, pago_detalle.beneficiario_destino " & _
            '      "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf= 'T' order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
       rsT.Update
    End If
    Set rsT = New ADODB.Recordset
    rsT.Open "select Nro_Transferencia,* from to_transferencia", db, adOpenKeyset, adLockOptimistic
    If rsT.RecordCount > 0 Then
       Set DtGTransferenciasImprimir.DataSource = rsT
       DtGTransferenciasImprimir.Enabled = True
       Retornar.Enabled = True
    End If
   
End Sub
Public Sub Refrescar()
    Set rsT = New ADODB.Recordset
    rsT.Open "select Nro_Transferencia,* from to_transferencia", db, adOpenKeyset, adLockOptimistic
    If rsT.RecordCount > 0 Then
       Set DtGTransferenciasImprimir.DataSource = rsT
    Else
       Set DtGTransferenciasImprimir.DataSource = rsNada
       DtGTransferenciasImprimir.Enabled = False
       Retornar.Enabled = False
    End If
End Sub

Public Sub Coloca_Status_Impreso()
    Set rsT = New ADODB.Recordset
    rsT.Open "select * from to_Transferencia", db, adOpenKeyset, adLockOptimistic
    If rsT.RecordCount > 0 Then
       While Not rsT.EOF
            Set rsOP = New ADODB.Recordset
            rsOP.Open "select * from to_cheques_Operaciones WHERE numero_cheque='" & rsT("Nro_Transferencia") & "' and cta_codigo='" & rsT("cta_origen") & "'", db, adOpenKeyset, adLockOptimistic
            If rsOP.RecordCount > 0 Then
                rsOP("estado_impreso") = "S"
                rsOP("Fecha_impreso") = Date
            Else
                rsOP.AddNew
                rsOP("numero_cheque") = Mid(CStr(100000 + Val(rsT("Nro_Transferencia"))), 2, 5)
                rsOP("cta_codigo") = rsT("cta_origen")
                rsOP("estado_impreso") = "S"
                rsOP("estado_entregado") = "N"
                rsOP("estado_anulado") = "N"
                rsOP("estado_cobrado") = "N"
                rsOP("estado_devuelto") = "N"
                rsOP("fecha_registro") = Date
                rsOP("Fecha_impreso") = Date
                rsOP("Cheq_Transf") = "T"
                'rsOP("hora_registro") = Str(Time)
                rsOP.Update
            End If
            'db.Execute "UPDATE to_cheques_Operaciones WHERE numero_cheque='" & rsCheque("numero_cheque") & "' and cta_codigo='" & rsCheque("cta_codigo") & "'"
            rsT.MoveNext
        Wend
            
    End If

End Sub

Private Sub TxtCmpte_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDesplegado 
   Caption         =   "Desplegando Cheques"
   ClientHeight    =   8535
   ClientLeft      =   180
   ClientTop       =   1815
   ClientWidth     =   13305
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoAuxiliar 
      Height          =   465
      Left            =   7470
      Top             =   7470
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   820
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
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame3 
      Height          =   7530
      Left            =   345
      TabIndex        =   15
      Top             =   1020
      Width           =   1320
      Begin VB.CommandButton CmdFiltCriterio 
         Caption         =   "Filtra CRITERIO"
         Height          =   975
         Left            =   150
         Picture         =   "FrmDesplegado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1245
         Width           =   1095
      End
      Begin VB.CommandButton CmdAnulaFCriterio 
         Caption         =   "Anula Filtro CRITERIO"
         Height          =   975
         Left            =   150
         Picture         =   "FrmDesplegado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2235
         Width           =   1095
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprime Cheque"
         Height          =   975
         Left            =   165
         Picture         =   "FrmDesplegado.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   255
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   975
         Left            =   135
         Picture         =   "FrmDesplegado.frx":0EEE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6315
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CRITERIO DE BUSQUEDA"
      Height          =   1110
      Left            =   1900
      TabIndex        =   8
      Top             =   1050
      Width           =   11280
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   645
         Width           =   1575
      End
      Begin VB.ComboBox CmbOperador 
         Height          =   315
         ItemData        =   "FrmDesplegado.frx":1330
         Left            =   2025
         List            =   "FrmDesplegado.frx":1343
         TabIndex        =   13
         Top             =   630
         Width           =   1065
      End
      Begin VB.ComboBox CmbCampo 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   630
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   375
         Left            =   5325
         TabIndex        =   20
         Top             =   540
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24444929
         CurrentDate     =   36413
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   375
         Left            =   7275
         TabIndex        =   21
         Top             =   555
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24444929
         CurrentDate     =   36413
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio"
         Height          =   240
         Left            =   5355
         TabIndex        =   23
         Top             =   330
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin"
         Height          =   240
         Left            =   7320
         TabIndex        =   22
         Top             =   345
         Width           =   1590
      End
      Begin VB.Label LblValor 
         Caption         =   "Valor"
         Height          =   285
         Left            =   3270
         TabIndex        =   11
         Top             =   255
         Width           =   675
      End
      Begin VB.Label LblOperador 
         Caption         =   "Operador"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   315
         Width           =   885
      End
      Begin VB.Label LblCampo 
         Caption         =   "Campo"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   690
      Left            =   365
      ScaleHeight     =   630
      ScaleWidth      =   12750
      TabIndex        =   1
      Top             =   210
      Width           =   12810
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRESION HISTORICO DE CHEQUES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   135
         Width           =   7320
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1875
      Top             =   8130
      Width           =   11340
      _ExtentX        =   20003
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\MVB5\Labs\Neptuno.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\MVB5\Labs\Neptuno.mdb"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5820
      Left            =   1875
      TabIndex        =   0
      Top             =   2235
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10266
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
   Begin VB.Frame Frame1 
      Caption         =   "BUSQUEDA"
      Height          =   870
      Left            =   1950
      TabIndex        =   3
      Top             =   7695
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton CmdUltimo 
         Caption         =   "Ultimo"
         Height          =   510
         Left            =   3515
         Picture         =   "FrmDesplegado.frx":135A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton CmdAnterior 
         Caption         =   "Anterior"
         Height          =   510
         Left            =   1225
         Picture         =   "FrmDesplegado.frx":145C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton CmdSiguiente 
         Caption         =   "Siguiente"
         Height          =   510
         Left            =   2370
         Picture         =   "FrmDesplegado.frx":1556
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton CmdPrimero 
         Caption         =   "Primero"
         Height          =   510
         Left            =   80
         Picture         =   "FrmDesplegado.frx":1658
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmDesplegado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsbusca As New ADODB.Recordset
Dim rsauxiliar As New ADODB.Recordset
Dim rsCheque As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim rsEntregados As New ADODB.Recordset

Dim CAMPOS As ADODB.Field
Dim Vquery As String
Public NoRegistros As Long
Dim Cry As CryOpCheques

Private Sub CmdAnulaF_Click()
    'Adodc1.Refresh
    rsbusca.Filter = adFilterNone
End Sub
'Private Sub CmdAnterior_Click()
'    If ValidaCriterio(CmbCampo.Text, CmbOperador.Text, TxtValor.Text) = 2 Then
'        If (Not rsbusca.BOF) Then
'            rsbusca.Find CmbCampo.Text & " " & CmbOperador.Text & " '" & TxtValor.Text & "'", , adSearchForward
'            rsbusca.MovePrevious
'            If (rsbusca.BOF) Then
'                rsbusca.MoveFirst
'                CmdAnterior.Enabled = False
'                CmdPrimero.Enabled = False
'                CmdSiguiente.Enabled = True
'                CmdUltimo.Enabled = True
'            Else
'                CmdAnterior.Enabled = True
'                CmdPrimero.Enabled = True
'                CmdSiguiente.Enabled = True
'                CmdUltimo.Enabled = True
'            End If
'        End If
'
'    Else
'        MsgBox errCriterio, vbExclamation, "ERROR"
'    End If
'End Sub

Private Sub CmdAnulaFCriterio_Click()
   db.Execute "DELETE FROM to_operaciones_cheques"
    rsbusca.Close
    If rsbusca.State = 1 Then rsbusca.Close
    rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
    If rsbusca.RecordCount > 0 Then
        NoRegistros = rsbusca.RecordCount
        'AdoCuenta.Caption = NoRegistros
        'Vaciando los datos al temporal to_operaciones_cheques
        While Not rsbusca.EOF
            If rsbusca.RecordCount > 0 Then
                db.Execute "insert into to_operaciones_cheques(numero_cheque,fecha_registro,monto_bolivianos,denominacion_beneficiario,cta_descripcion_larga,cta_codigo, codigo_pago,estado_impreso,estado_entregado,estado_cobrado,estado_anulado) values ('" & rsbusca!numero_cheque & "','" & rsbusca!fecha_registro & "','" & rsbusca!monto_bolivianos & "','" & rsbusca!denominacion_beneficiario & "','" & rsbusca!cta_descripcion_larga & "','" & rsbusca!cta_codigo & "', '" & rsbusca!codigo_pago & "','" & rsbusca!estado_impreso & "','" & rsbusca!estado_entregado & "','" & rsbusca!estado_cobrado & "','" & rsbusca!estado_anulado & "') "
                rsbusca.MoveNext
            End If
        Wend
    Else
        NoRegistros = 0
    End If
    Set Adodc1.Recordset = rsbusca
    Set DataGrid1.DataSource = Adodc1
End Sub
Private Sub CmdFiltCriterio_Click()
Dim errCriterio As Variant
Dim sw As Integer
db.Execute "DELETE FROM to_operaciones_cheques"
    sw = 0
    If rsbusca.State = 1 Then rsbusca.Close
    'esta bien
    Set DataGrid1.DataSource = rsauxiliar
    If CmbCampo.Text <> "" And CmbOperador.Text <> "" And "'" & TxtValor.Text & "'" <> "" Then
        rsbusca.Open Vquery & " Where " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'", db, adOpenStatic, adLockReadOnly
                        ''                        If rsbusca.RecordCount > 0 Then
                        ''                           NoRegistros = rsbusca.RecordCount
                        ''                           Set Adodc1.Recordset = rsbusca
                        ''                           Set DataGrid1.DataSource = rsbusca
                        ''                        Else
                        ''                            Set Adodc1.Recordset = rsNada
                        ''                            Set DataGrid1.DataSource = rsNada
                        ''                            NoRegistros = 0
                        ''                        End If
                        'Vaciando los datos al temporal to_operaciones_cheques
                        While Not rsbusca.EOF
                            If rsbusca.RecordCount > 0 Then
                                If rsbusca("Fecha_Registro") >= DTPFechaInicio.Value And rsbusca("Fecha_Registro") <= DTPFechaFin.Value Then
                                    db.Execute "insert into to_operaciones_cheques(numero_cheque,fecha_registro,monto_bolivianos,denominacion_beneficiario,cta_descripcion_larga,cta_codigo, codigo_pago,estado_impreso,estado_entregado,estado_cobrado,estado_anulado,Cheq_Transf) values ('" & rsbusca!numero_cheque & "','" & rsbusca!fecha_registro & "','" & rsbusca!monto_bolivianos & "','" & rsbusca!denominacion_beneficiario & "','" & rsbusca!cta_descripcion_larga & "','" & rsbusca!cta_codigo & "', '" & rsbusca!codigo_pago & "','" & rsbusca!estado_impreso & "','" & rsbusca!estado_entregado & "','" & rsbusca!estado_cobrado & "','" & rsbusca!estado_anulado & "', '" & rsbusca!cheque_o_trf & "')"
                                    sw = 1
                                End If
                                 rsbusca.MoveNext
                            End If
                        Wend
                                  If sw <> 1 And rsbusca.EOF Then
                                         'Set Adodc1.Recordset = rsNada
                                         Set DataGrid1.DataSource = rsNada
                                         NoRegistros = 0
                                         sw = 0
                                  Else
                                        If rsEntregados.State = 1 Then rsEntregados.Close
                                        rsEntregados.Open "SELECT * from to_operaciones_cheques", db, adOpenKeyset, adLockOptimistic
                                        If rsEntregados.RecordCount > 0 Then
                                               Set Adodc1.Recordset = rsEntregados
                                               Set DataGrid1.DataSource = rsEntregados
                                               NoRegistros = rsbusca.RecordCount
                                        End If
                                  End If
  
                        
                    Else
                        MsgBox errCriterio, vbExclamation, "COLOQUE DATOS"
                        rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
                        If rsbusca.RecordCount > 0 Then
                           NoRegistros = rsbusca.RecordCount
                        Else
                           NoRegistros = 0
                        End If
                        Set Adodc1.Recordset = rsbusca
                        Set DataGrid1.DataSource = rsbusca
           End If
           
'           If SW <> 1 Then
'                 Set Adodc1.Recordset = rsNada
'                 Set DataGrid1.DataSource = rsNada
'                 NoRegistros = 0
'                 SW = 0
'            Else
'                 NoRegistros = rsbusca.RecordCount
'            End If
           
End Sub
Private Sub CmdImprimir_Click()
'    Dim errCriterio As Variant
'    Dim VVista As String
'            'Vaciamos los datos a un auxiliar
'             Set rsCheque = New ADODB.Recordset
'             rsCheque.CursorLocation = adUseClient
'             If rsCheque.State = 1 Then rsCheque.Close
'             rsCheque.Open "select * from ts_cheque_operacion", db, adOpenKeyset, adLockOptimistic
'             If rsCheque.RecordCount > 0 Then
'               While Not rsCheque.EOF
'                     rsCheque.Delete
'                     rsCheque.MoveNext
'               Wend
'            End If
'
'    Set rsbusca = New ADODB.Recordset
'    If rsbusca.State = 1 Then rsbusca.Close
'    VVista = "SELECT fc_organismo_financiamiento.Org_descripcion, pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, fc_beneficiario.denominacion_beneficiario, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,to_cheques.estado_impreso " & _
'             "FROM fc_organismo_financiamiento, ((pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo) INNER JOIN to_cheques ON pago_detalle.numero_cheque_trf = to_cheques.numero_cheque " & _
'             "WHERE pago_detalle.cheque_o_trf= 'C'; "
'   ' If CmbCampo.Text <> "" And CmbOperador.Text <> "" And "'" & TxtValor.Text & "'" <> "" Then
'       'rsbusca.Open VVista & " Where " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'", db, adOpenStatic, adLockReadOnly
'       'MsgBox VVista
'       'rsbusca.Open VVista & " AND " & to_cheques.CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'", db, adOpenStatic, adLockReadOnly
'       rsbusca.Open VVista, db, adOpenStatic, adLockReadOnly
'
'       If rsbusca.RecordCount > 0 Then
'       While Not rsbusca.EOF
'
'            Dim X As String
'            'If rsCheque.State = 1 Then rsCheque.Close
'            Set rsCheque = New ADODB.Recordset
'            rsCheque.Open "select * from ts_cheque_operacion", db, adOpenStatic, adLockOptimistic
'                    rsCheque.AddNew
'                    rsCheque("fecha") = rsbusca("fecha_pago")
'                    rsCheque("numero_cheque") = rsbusca("numero_cheque_trf")
'                    rsCheque("denominacion_beneficiario") = rsbusca("denominacion_beneficiario")
'                    rsCheque("monto_bolivianos") = rsbusca("monto_bolivianos")
'                    rsCheque("cta_codigo") = rsbusca("cta_codigo")
'                    rsCheque.Update
'            rsbusca.MoveNext
'        Wend
'      End If
'    '  End If


                ''''Dim I As Integer
                ''''Dim AUXILIAR As String
                ''''
                ''''On Error GoTo temporal:
                ''''    Set rsCheque = New ADODB.Recordset
                ''''    If rsCheque.State = 1 Then rsCheque.Close
                ''''    rsCheque.Open "SELECT * FROM to_operaciones_cheques", db, adOpenDynamic, adLockOptimistic
                ''''    If rsCheque.RecordCount > 0 Then
                ''''      While Not rsCheque.EOF And Not rsCheque.BOF
                ''''            rsCheque.Delete
                ''''            rsCheque.Update
                ''''            rsCheque.MoveNext
                ''''      Wend
                ''''    End If
                ''''
                ''''    Set rsCheque = New ADODB.Recordset
                ''''    If rsCheque.State = 1 Then rsCheque.Close
                ''''    rsCheque.Open "SELECT * FROM to_operaciones_cheques", db, adOpenStatic, adLockOptimistic
                ''''    For I = 0 To NoRegistros - 1
                ''''        rsCheque.AddNew
                ''''        DataGrid1.Row = I
                ''''        rsCheque("numero_cheque") = DataGrid1.Columns(0)
                ''''        rsCheque("fecha_registro") = DataGrid1.Columns(1)
                ''''        rsCheque("monto_bolivianos") = DataGrid1.Columns(2)
                ''''        rsCheque("denominacion_beneficiario") = DataGrid1.Columns(3)
                ''''        rsCheque("cta_descripcion_larga") = DataGrid1.Columns(4)
                ''''        rsCheque("cta_codigo") = DataGrid1.Columns(5)
                ''''        rsCheque("estado_impreso") = DataGrid1.Columns(6)
                ''''        rsCheque("estado_entregado") = DataGrid1.Columns(7)
                ''''        rsCheque("estado_cobrado") = DataGrid1.Columns(8)
                ''''        rsCheque("estado_anulado") = DataGrid1.Columns(9)
                ''''        rsCheque("codigo_pago") = DataGrid1.Columns(10)
                ''''        rsCheque.Update
                ''''    Next I
                ''''    'Cry.PaperOrientation = crLandscape
                ''''    RepOperacionesCheques.Show
                ''''
                '''' Exit Sub
                ''''temporal:
                ''''    Set rsCheque = New ADODB.Recordset
                ''''    If rsCheque.State = 1 Then rsCheque.Close
                ''''    rsCheque.Open "SELECT * FROM to_operaciones_cheques", db, adOpenDynamic, adLockOptimistic
                ''''    Resume
 RepOperacionesCheques.Show
End Sub
'Private Sub CmdOrdenar_Click()
'Dim I As Integer
'
'    I = InStr(1, UCase(FrmBuscador.Txtquery.Text), "ORDER BY ")
'    If I > 0 Then
'        FrmBuscador.Txtquery = Mid(FrmBuscador.Txtquery, 1, I - 1)
'    End If
'    rsbusca.Close
'    If rsbusca.State = 1 Then rsbusca.Close
'    If CmbCampo.Text <> "" Then
'        rsbusca.Open FrmBuscador.Txtquery.Text & " order by " & CmbCampo.Text, db, adOpenStatic, adLockReadOnly
'        Set Adodc1.Recordset = rsbusca
'        Set DataGrid1.DataSource = rsbusca
'    Else
'        MsgBox "Es necesario establecer el nombre en la casilla de 'Campo'", vbOKOnly, "ERROR"
'        rsbusca.Open FrmBuscador.Txtquery.Text, db, adOpenStatic, adLockReadOnly
'        Set Adodc1.Recordset = rsbusca
'        Set DataGrid1.DataSource = rsbusca
'    End If
'End Sub
'
'Private Sub CmdPrimero_Click()
'    If ValidaCriterio(CmbCampo.Text, CmbOperador.Text, TxtValor.Text) = 2 Then
'        If (Not rsbusca.BOF) Then
'            'rsbusca.MoveFirst
'            rsbusca.Find CmbCampo.Text & " " & CmbOperador.Text & " '" & TxtValor.Text & "'", , adSearchForward
'            rsbusca.MoveFirst
'            CmdPrimero.Enabled = False
'            CmdAnterior.Enabled = False
'            CmdSiguiente.Enabled = True
'            CmdUltimo.Enabled = True
'        End If
'    Else
'        MsgBox errCriterio, vbExclamation, "ERROR"
'    End If
'End Sub
'
Private Sub CmdSalir_Click()
    'If CmbCampo <> "" Then
        'Vquery = rsbusca("Ges_gestion")
        'vquery = rsbusca("NombreCompañía")
    'End If
    Unload Me
End Sub

Private Sub CmdSalirPagos_Click()
    Unload Me
End Sub


'Private Sub CmdSiguiente_Click()
'    If ValidaCriterio(CmbCampo.Text, CmbOperador.Text, TxtValor.Text) = 2 Then
'        'If (Not rsbusca.EOF) And (Not rsbusca.BOF) Then
'        'Ojjjjjjjjjjjooooooooo
'        If (Not rsbusca.EOF) Or (Not rsbusca.BOF) Then
'            'rsbusca.MoveNext
'            rsbusca.Find CmbCampo.Text & " " & CmbOperador.Text & " '" & TxtValor.Text & "'", , adSearchForward
'            rsbusca.MoveNext
'            If (rsbusca.EOF) Then
'                rsbusca.MoveLast
'                CmdAnterior.Enabled = True
'                CmdPrimero.Enabled = True
'                CmdSiguiente.Enabled = False
'                CmdUltimo.Enabled = False
'            End If
'        End If
'    Else
'        MsgBox errCriterio, vbExclamation, "ERROR"
'    End If
'End Sub
'
'Private Sub CmdUltimo_Click()
'    If ValidaCriterio(CmbCampo.Text, CmbOperador.Text, TxtValor.Text) = 2 Then
'        If (Not rsbusca.EOF) Then
'            rsbusca.Find CmbCampo.Text & " " & CmbOperador.Text & " '" & TxtValor.Text & "'", , adSearchForward
'            rsbusca.MoveLast
'            CmdAnterior.Enabled = True
'            CmdPrimero.Enabled = True
'            CmdSiguiente.Enabled = False
'            CmdUltimo.Enabled = False
'        End If
'    Else
'        MsgBox errCriterio, vbExclamation, "ERROR"
'    End If
'End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim i As Integer
Dim SwOrden As Variant
    
    Print ColIndex
    'DataGrid1.Columns(ColIndex).Caption
    'Combo1.Text = Adodc1.
    If rsbusca.State = 1 Then rsbusca.Close
    If SwOrden Then
        rsbusca.Open Vquery & " order by " & CmbCampo.List(ColIndex) & " asc ", db, adOpenStatic, adLockReadOnly
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = Adodc1
        SwOrden = False
    Else
        rsbusca.Open Vquery & " order by " & CmbCampo.List(ColIndex) & " desc ", db, adOpenStatic, adLockReadOnly
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = Adodc1
        SwOrden = True
    End If
    '  i = InStr(1, UCase(Form1.Text1.Text), "ORDER BY")
    '  If i > 0 Then
    '    Form1.Text1 = Mid(Form1.Text1, 1, i - 1)
    '  End If
    '  If rsbusca.State = 1 Then rsbusca.Close
    '  rsbusca.Open Form1.Text1.Text & " order by " & Combo1.Text, db, adOpenStatic, adLockReadOnly
    '  Set Adodc1.Recordset = rsbusca
    '  Set DataGrid1.DataSource = Adodc1
  
End Sub

Private Sub Form_Load()
Dim errCriterio As Variant
'Vquery = "select * from to_cheques_operaciones"
Vquery = "select * from cns_cheques"

    errCriterio = "Existe Error en el Criterio"
    If rsbusca.State = 1 Then rsbusca.Close
    On Error GoTo Error:
        rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
        If rsbusca.RecordCount > 0 Then
            NoRegistros = rsbusca.RecordCount
            'AdoCuenta.Caption = NoRegistros
        Else
            NoRegistros = 0
        End If
        
        'Set DataGrid1.DataSource = rsbusca
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = Adodc1
        For Each CAMPOS In rsbusca.Fields
            CmbCampo.AddItem CAMPOS.Name
        Next CAMPOS
        
        DTPFechaInicio.Value = Date
        DTPFechaFin.Value = Date
    Exit Sub
Error:
    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"
	Call SeguridadSet(Me)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


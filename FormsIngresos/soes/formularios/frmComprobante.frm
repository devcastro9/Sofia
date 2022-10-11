VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComprobante 
   Caption         =   "Lista de Comprobantes"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmComprobante.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   120
      TabIndex        =   13
      Top             =   4035
      Width           =   7875
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   6780
         MousePointer    =   4  'Icon
         Picture         =   "frmComprobante.frx":324A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   5805
         MousePointer    =   4  'Icon
         Picture         =   "frmComprobante.frx":3554
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   1005
      End
      Begin MSAdodcLib.Adodc adoComprobantes 
         Height          =   330
         Left            =   75
         Top             =   165
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "Adodc1"
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
   Begin VB.Frame Frame1 
      Caption         =   "Comprobantes"
      Height          =   915
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   7860
      Begin VB.CommandButton cmd_selecionar 
         Caption         =   ">"
         Height          =   300
         Left            =   2565
         TabIndex        =   9
         Top             =   300
         Width           =   345
      End
      Begin VB.TextBox txt_total_ele 
         Alignment       =   1  'Right Justify
         DataField       =   "soc_nro_sol"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4020
         TabIndex        =   6
         Top             =   300
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.TextBox txt_al 
         DataField       =   "soc_nro_sol"
         Height          =   285
         Left            =   1605
         TabIndex        =   3
         Top             =   300
         Width           =   800
      End
      Begin VB.TextBox txt_Del 
         DataField       =   "soc_nro_ref"
         Height          =   285
         Left            =   435
         TabIndex        =   2
         Top             =   315
         Width           =   800
      End
      Begin MSMask.MaskEdBox txt_total_monto 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   6015
         TabIndex        =   11
         Top             =   165
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_total_monto_bs 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   6015
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Monto Total Bs"
         Height          =   195
         Index           =   4
         Left            =   4860
         TabIndex        =   10
         Top             =   540
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Monto Total Us"
         Height          =   195
         Index           =   2
         Left            =   4845
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Elegidos "
         Height          =   195
         Index           =   1
         Left            =   2955
         TabIndex        =   7
         Top             =   330
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "al  "
         Height          =   195
         Index           =   0
         Left            =   1305
         TabIndex        =   5
         Top             =   360
         Width           =   210
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "del "
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
   End
   Begin MSDataGridLib.DataGrid dgComprobantes 
      Bindings        =   "frmComprobante.frx":385E
      Height          =   3060
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5398
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "codigo_pago"
         Caption         =   "Nro. Compr."
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
         DataField       =   "codigo_unidad"
         Caption         =   "Unidad"
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
      BeginProperty Column02 
         DataField       =   "monto_Bolivianos"
         Caption         =   "Monto Bs."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "monto_Dolares"
         Caption         =   "Monto Us."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "codigo_categoria"
         Caption         =   "Categoria"
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
      BeginProperty Column05 
         DataField       =   "tiqueado"
         Caption         =   "Elegir"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
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
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   360
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim accion As String
Public frmComprobante_ret As String

Public Sub frmComprobante_procesar(proceso As String)
  accion = proceso
  frmComprobante_ret = ""
  If proceso = "SELECT_DET_SOES" Then
    Caption = ""
  End If
  Show vbModal
End Sub

Private Sub adoComprobantes_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If Not (adoComprobantes.Recordset.EOF Or adoComprobantes.Recordset.BOF) Then
    adoComprobantes.Caption = CStr(adoComprobantes.Recordset.Bookmark) & " de " & CStr(adoComprobantes.Recordset.RecordCount)
  Else
    adoComprobantes.Caption = "0 de 0"
  End If
End Sub

Private Sub cmd_selecionar_Click()
  MarcaFilas Val(txt_Del.Text), Val(txt_al.Text)
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim varBmk As Variant, cant As Integer
  cant = 0
  For Each varBmk In dgComprobantes.SelBookmarks
    adoComprobantes.Recordset.Bookmark = varBmk
    pre_insert_Detalle_soa
    frmDetalleSoes.Detalle_soa_refresca False, frmDetalleSoes.txtsoe_cod_convenio, Val(frmDetalleSoes.txtsoc_nro_sol)
    cant = cant + 1
    frmDetalleSoes.txtsoe_cant_comp.Text = Str(Val(frmDetalleSoes.txtsoe_cant_comp.Text) + cant)
  Next
  frmDetalleSoes.LlenaTotalMonto
  Unload Me
End Sub

Private Sub LlenaTotalMontoComp()
Dim total, total_bs, cant As Double
  total = 0
  total_bs = 0
  cant = 0
  If Me.adoComprobantes.Recordset.RecordCount > 0 Then
    Me.adoComprobantes.Recordset.MoveFirst
    While Not Me.adoComprobantes.Recordset.EOF
      total = total + IIf(IsNull(Me.adoComprobantes.Recordset!monto_dolares), 0, Me.adoComprobantes.Recordset!monto_dolares)
      total_bs = total_bs + IIf(IsNull(Me.adoComprobantes.Recordset!monto_bolivianos), 0, Me.adoComprobantes.Recordset!monto_bolivianos)
      cant = cant + 1
      Me.adoComprobantes.Recordset.MoveNext
    Wend
  End If
  txt_total_ele.Text = Str(cant)
  txt_total_monto.Text = Str(total)
  txt_total_monto_bs.Text = Str(total_bs)
End Sub

Function GetNro_veces_enviado(ges_gestion, org_codigo As String, codigo_pago As Integer)
  GetNro_veces_enviado = 0
End Function

Private Sub MarcaFilas(ini, fin As Integer)
  If Me.adoComprobantes.Recordset.RecordCount > 0 Then
    Me.adoComprobantes.Recordset.MoveFirst
    While Not Me.adoComprobantes.Recordset.EOF
      If Me.adoComprobantes.Recordset!codigo_pago >= ini _
        And Me.adoComprobantes.Recordset!codigo_pago <= fin Then
        dgComprobantes.SelBookmarks.Add adoComprobantes.Recordset.Bookmark
      End If
      Me.adoComprobantes.Recordset.MoveNext
    Wend
  End If
End Sub

Function get_porciento_bid2(codigo_convenio, pro_proyecto, ByVal par_cod As String, dePersonal As Boolean)
Dim porciento_bid As Double
'  Datos.rsdbo_apGeneralSearching.Close
  If dePersonal Then
    Datos.dbo_apGeneralSearching _
      "SELECT isnull(prc_porcentaje_aux, 0) as porcentaje " _
      & "FROM so_porcentaje_convenio " _
      & "WHERE codigo_convenio = '" & codigo_convenio & "'" _
      & " and pro_proyecto = '" & pro_proyecto & "'" _
      & " and par_codigo = '" & par_cod & "'"
  Else
    Datos.dbo_apGeneralSearching _
      "SELECT isnull(prc_porcentaje, 0) as porcentaje " _
      & "FROM so_porcentaje_convenio " _
      & "WHERE codigo_convenio = '" & codigo_convenio & "'" _
      & " and pro_proyecto = '" & pro_proyecto & "'" _
      & " and par_codigo = '" & par_cod & "'"
  End If
  On Error GoTo ControlError
    With Datos.rsdbo_apGeneralSearching
      porciento_bid = Datos.rsdbo_apGeneralSearching!porcentaje
     .Close
    End With
    porciento_bid = IIf(IsNull(porciento_bid), 0, porciento_bid)
    get_porciento_bid2 = porciento_bid
    Exit Function
ControlError:
  Datos.rsdbo_apGeneralSearching.Close
  MsgBox "Atención. No se encuentra registrado ningun porcentaje.. Verifique esta Información"
  porciento_bid = IIf(IsNull(porciento_bid), 0, porciento_bid)
  get_porciento_bid2 = porciento_bid
End Function

Private Sub pre_insert_Detalle_soa()
Dim tc As Double, cod_beneficiario, par_codigo, pro_proyecto As String
Dim porciento_bid, monto_pago As Double
Dim monto_tot_bs, monto_equi, monto_bid, monto_otr As Double
Dim retInput As String, proceso_consultor As Boolean
Dim ok As Boolean
  ok = True
  porciento_bid = 0
  monto_pago = IIf(IsNull(frmComprobante.adoComprobantes.Recordset!monto_bolivianos), 0, frmComprobante.adoComprobantes.Recordset!monto_bolivianos)
  'GetTc2 obtiene ... cod_beneficiario, par_codigo, pro_proyecto
  frmSoesMain.GetTc2 frmComprobante.adoComprobantes.Recordset!org_codigo, frmComprobante.adoComprobantes.Recordset!codigo_pago, _
           frmSoesMain.dcmCtas.BoundText, tc, cod_beneficiario, par_codigo, pro_proyecto
  
  If tc = 0 Then
    MsgBox "tipo de cambio cero para el comprobante " & frmComprobante.adoComprobantes.Recordset!codigo_pago & " con organismo " & frmComprobante.adoComprobantes.Recordset!org_codigo _
      & ". No se registrará el comprobante"
  Else
   
    If par_codigo <> gl_partida_consultores Then  'si la partida no es de planilla
      porciento_bid = get_porciento_bid2(frmComprobante.adoComprobantes.Recordset!codigo_convenio, _
         pro_proyecto, par_codigo, False)
      While porciento_bid <= 0 Or porciento_bid > 100
        porciento_bid = getNumber("Ingrese Porcentaje para el comprobante " _
          & adoComprobantes.Recordset!codigo_pago _
          & " con Fuente de financiamiento " & adoComprobantes.Recordset!org_codigo _
          & " Proyecto " & pro_proyecto _
          & " Partida " & par_codigo _
          & ":", "Porcentaje correspondiente al BID")
        If porciento_bid <= 0 Or porciento_bid > 100 Then
          MsgBox "Valor no corresponde a un porcentaje"
        End If
      Wend
      Call GetValMontos(frmSoesMain.txtsoc_tipo_mone_sol.Text, tc, porciento_bid, monto_pago, monto_tot_bs, monto_equi, monto_bid, monto_otr)
      Call insert_Detalle_soa(cod_beneficiario, monto_pago, tc, monto_equi, monto_bid, porciento_bid, monto_otr, monto_otr_bs, 0, monto_tot_bs, 0, "")
    Else
      'Se obtiene el 2do porcentaje correspondiente a personal permanente
       porciento_bid = get_porciento_bid2(frmComprobante.adoComprobantes.Recordset!codigo_convenio, _
         pro_proyecto, par_codigo, True)
      If porciento_bid = 0 Then
        MsgBox "Porcentaje es cero. El comprobante no se procesará. " _
          & "Verifique porcentaje para el comprobante " & adoComprobantes.Recordset!codigo_pago _
          & " Proyecto " & pro_proyecto _
          & " Partida " & par_codigo
      Else
        proceso_consultor = False
        'Datos.dbo_so_comprobantes "LISTA_CONSULTORES", 0, 0, frmComprobante.adoComprobantes.Recordset!codigo_orden, "", "", ""
        Datos.dbo_so_comprobantes "LISTA_CONSULTORES", 0, 0, Trim(Str(frmComprobante.adoComprobantes.Recordset!codigo_pago)), frmComprobante.adoComprobantes.Recordset!org_codigo, "", ""
        With Datos.rsdbo_so_comprobantes
          Do While Not .EOF
            Call GetValMontos(frmSoesMain.txtsoc_tipo_mone_sol.Text, tc, porciento_bid, !monto_bs_ext, monto_tot_bs, monto_equi, monto_bid, monto_otr)
            Call insert_Detalle_soa(cod_beneficiario, monto_pago, tc, monto_equi, monto_bid, porciento_bid, monto_otr, monto_otr_bs, !idFuncionario, monto_tot_bs, 0, IIf(IsNull(!codigo_prisma), " ", !codigo_prisma))
            proceso_consultor = True
            .MoveNext
          Loop
          .Close
        End With
        If Not proceso_consultor Then
          MsgBox "Atención. No se encontró informacion detallada del Personal al que se pago con este comprobante. Verifique con el Supervisor esta Información"
          Call GetValMontos(frmSoesMain.txtsoc_tipo_mone_sol.Text, tc, porciento_bid, monto_pago, monto_tot_bs, monto_equi, monto_bid, monto_otr)
          Call insert_Detalle_soa(cod_beneficiario, monto_pago, tc, monto_equi, monto_bid, porciento_bid, monto_otr, monto_otr_bs, 0, monto_tot_bs, 0, "")
        End If
      End If
    End If
  End If
End Sub

Private Sub insert_Detalle_soa(ByVal cod_beneficiario As String, _
ByVal monto_pago, ByVal tc, ByVal monto_equi, ByVal monto_bid, ByVal porciento_bid, ByVal monto_otr, ByVal monto_otr_bs As Double, nro_veces As Integer, ByVal monto_tot_bs As Double, ByVal idFuncionario As Double, ByVal codigo_prisma As String)
    Datos.dbo_so_detalle_soes "INSERT" _
      , Val(frmDetalleSoes.txtsoc_nro_sol.Text) _
      , frmSoesMain.cb_codigo_convenio.Text _
      , Val(frmDetalleSoes.txtsoe_nro_sec.Text) _
      , frmComprobante.adoComprobantes.Recordset!ges_gestion _
      , frmComprobante.adoComprobantes.Recordset!org_codigo _
      , frmComprobante.adoComprobantes.Recordset!codigo_pago _
      , nro_veces _
      , cod_beneficiario _
      , codigo_prisma _
      , IIf(IsNull(frmComprobante.adoComprobantes.Recordset!fecha_registro), CDate("01/01/1900"), frmComprobante.adoComprobantes.Recordset!fecha_registro) _
      , monto_pago _
      , tc _
      , IIf(IsNull(monto_equi), 0, monto_equi) _
      , IIf(IsNull(monto_bid), 0, monto_bid) _
      , IIf(IsNull(porciento_bid), 0, porciento_bid) _
      , IIf(IsNull(monto_otr), 0, monto_otr) _
      , IIf(IsNull(monto_tot_bs), 0, monto_tot_bs) _
      , "No"
End Sub


Private Sub GetValMontos(moneda As String, tc, porc_bid, monto_pago, monto_tot_bs, monto_equi, monto_bid, monto_otr As Double)
  'monto_otr_bs = (monto_pago / porc_fin) * (100 - porc_fin)
  monto_tot_bs = monto_pago + ((monto_pago / porc_bid) * (100 - porc_bid))
  If moneda = "USD" Then
    monto_equi = monto_tot_bs / tc
  Else
    monto_equi = monto_tot_bs
  End If
  monto_bid = monto_equi * porc_bid / 100
  monto_otr = IIf(IsNull(monto_equi - monto_bid), 0, monto_equi - monto_bid)
End Sub


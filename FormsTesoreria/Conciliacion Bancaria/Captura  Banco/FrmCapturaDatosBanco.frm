VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form FrmCapturaDatosBanco 
   Caption         =   "Capturando datos de banco"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   Icon            =   "FrmCapturaDatosBanco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryBan 
      Left            =   6810
      Top             =   1815
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
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   645
      Left            =   -195
      Top             =   8235
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1138
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
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   1290
      TabIndex        =   1
      Top             =   8070
      Width           =   11280
      Begin VB.TextBox TxtCodigoBanco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   975
         TabIndex        =   25
         Top             =   960
         Width           =   2160
      End
      Begin VB.TextBox TxtOrganismo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Top             =   1335
         Width           =   2130
      End
      Begin VB.TextBox TxtBanco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   21
         Top             =   960
         Width           =   5370
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "FrmCapturaDatosBanco.frx":0ECA
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   990
         TabIndex        =   2
         Top             =   225
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "FrmCapturaDatosBanco.frx":0EE2
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   3180
         TabIndex        =   3
         Top             =   225
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "FrmCapturaDatosBanco.frx":0EFA
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   990
         TabIndex        =   20
         Top             =   600
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Label Label8 
         Caption         =   "Organismo"
         Height          =   210
         Left            =   150
         TabIndex        =   24
         Top             =   1395
         Width           =   810
      End
      Begin VB.Label Label5 
         Caption         =   "Banco"
         Height          =   180
         Left            =   150
         TabIndex        =   23
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. "
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   630
      End
   End
   Begin MSDataGridLib.DataGrid DtGDatosBanco 
      Height          =   6255
      Left            =   7410
      TabIndex        =   19
      Top             =   1785
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   11033
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
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   1320
      TabIndex        =   15
      Top             =   1050
      Width           =   11310
      Begin VB.Label Label4 
         Caption         =   "Datos EXCEL"
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
         TabIndex        =   17
         Top             =   255
         Width           =   2445
      End
      Begin VB.Label Label7 
         Caption         =   "Datos  Base"
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
         Left            =   5970
         TabIndex        =   16
         Top             =   240
         Width           =   2445
      End
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "FrmCapturaDatosBanco.frx":0F12
      Height          =   6255
      Left            =   1350
      OleObjectBlob   =   "FrmCapturaDatosBanco.frx":0F26
      TabIndex        =   0
      Top             =   1785
      Width           =   5280
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   3255
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   5595
      Width           =   2475
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   8670
      TabIndex        =   5
      Top             =   0
      Width           =   8730
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "CAPTURA DE DATOS DEL BANCO "
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
         Height          =   315
         Left            =   3540
         TabIndex        =   10
         Top             =   135
         Width           =   5235
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   9
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   8
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   7
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   60
         TabIndex        =   6
         Top             =   675
         Width           =   1110
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   7035
      Left            =   15
      TabIndex        =   11
      Top             =   990
      Width           =   1245
      Begin VB.CommandButton CmdLimpiarDatosBase 
         Caption         =   "Limpiar Datos Base"
         Height          =   735
         Left            =   135
         TabIndex        =   27
         Top             =   2865
         Width           =   945
      End
      Begin VB.CommandButton CmdConciliar 
         Caption         =   "Conciliar"
         Height          =   720
         Left            =   135
         TabIndex        =   26
         Top             =   3600
         Width           =   945
      End
      Begin VB.CommandButton CmdTransferencia 
         Caption         =   "Tranferir Datos de Excel "
         Height          =   735
         Left            =   135
         TabIndex        =   18
         Top             =   2130
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   135
         Picture         =   "FrmCapturaDatosBanco.frx":3628
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5895
         Width           =   930
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Buscar Documento Execel"
         Height          =   915
         Left            =   135
         Picture         =   "FrmCapturaDatosBanco.frx":3A6A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1215
         Width           =   945
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Impresión Datos Banco"
         Height          =   975
         Left            =   135
         Picture         =   "FrmCapturaDatosBanco.frx":3B6C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmCapturaDatosBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBANCO As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset

Private Sub CmdBusqueda_Click()
    FrmExplorador.Show
End Sub

Private Sub CmdConciliar_Click()
    FrmConciliacion.Show
End Sub

Private Sub cmdImprimir_Click()
        
        CryBan.ReportFileName = App.Path & "\FormsTesoreria\Conciliacion Bancaria\REPORTES\RptDatosBanco.rpt"
        iResult = CryBan.PrintReport
        If iResult <> 0 Then
           MsgBox CryBan.LastErrorNumber & " : " & CryBan.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
End Sub

Private Sub CmdLimpiarDatosBase_Click()
        Set rsBANCO = New ADODB.Recordset
        rsBANCO.Open "select * from fc_DatosBanco order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
        If rsBANCO.RecordCount > 0 Then
           db.Execute "delete from fc_DatosBanco"
           Set DtGDatosBanco.DataSource = rsNada
        End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub CmdTransferencia_Click()
Dim NumeroCheque As String
Dim rsDatosBanco As New ADODB.Recordset
    'Limpiando fc_DatosBanco
    'db.Execute "delete from fc_DatosBanco"
    'Validacion de datos
    If DtCCuentaOrigen.Text = "" Then
        MsgBox "Elija una cuenta para realizar la transferencia", vbInformation + vbDefaultButton1, "Validación de datos"
        Exit Sub
    End If
    Data1.Recordset.MoveFirst
    While Not (Data1.Recordset.EOF)
      If Not IsNull(Data1.Recordset("fecha_pago")) Then
        If Not IsNull(Data1.Recordset("Nro_Doc")) Then
                Select Case Len(Data1.Recordset("Nro_Doc"))
                    Case 1
                        NumeroCheque = "0000" & Data1.Recordset("Nro_Doc")
                    Case 2
                        NumeroCheque = "000" & Data1.Recordset("Nro_Doc")
                    Case 3
                        NumeroCheque = "00" & Data1.Recordset("Nro_Doc")
                    Case 4
                        NumeroCheque = "0" & Data1.Recordset("Nro_Doc")
                    Case 5
                        NumeroCheque = Data1.Recordset("Nro_Doc")
                   Case Else
                        NumeroCheque = Data1.Recordset("Nro_Doc")
                End Select
        If rsDatosBanco.State = 1 Then rsDatosBanco.Close
        rsDatosBanco.Open "SELECT * FROM fc_datosBanco WHERE Cta_Codigo='" & DtCCuentaOrigen.Text & "' and Nro_doc ='" & NumeroCheque & "' and Fecha_Pago= '" & Data1.Recordset("Fecha_Pago") & "'", db, adOpenKeyset, adLockOptimistic
        If rsDatosBanco.RecordCount <= 0 Then
                db.Execute "insert into fc_datosBanco (nro_cmpte, Organismo, Fecha_Pago, Monto,Nro_Doc, Cta_Codigo,  Estado_Conciliacion, Justificacion, Bco_Codigo) " & _
                            "values ('1',' " & txtorganismo.Text & " ', ' " & Format(Data1.Recordset("Fecha_Pago")) & " ', " & Data1.Recordset("Monto") & ", '" & NumeroCheque & "',  '" & DtCCuentaOrigen.Text & "', 'N','" & Data1.Recordset("Justificacion") & "','1')"
        Else
           MsgBox "Registro existente", vbCritical + vbDefaultButton2, "Validación de Dtaos"
        End If
      End If
      Else
        
      End If
      Data1.Recordset.MoveNext
    Wend
  
    'Abriendo Tabla  de registros del Banco
    Set rsBANCO = New ADODB.Recordset
    rsBANCO.Open "select * from fc_DatosBanco order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
    If rsBANCO.RecordCount > 0 Then
        Set DtGDatosBanco.DataSource = rsBANCO
    End If
    
'Error_Trasferencia:
'  If Err.Number = 13 Then
'  End If
'  If Err.Number = 3265 Then
'     'MsgBox Err.Number & Err.Description
'     MsgBox "Anotar los nombres de campos correctos y necesarios en EXCEL ", vbCritical + vbDefaultButton1
'     Exit Sub
'  End If
  
 
End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
End Sub

Private Sub DtCCuentaOrigen_Change()
        'Determinar los bancos
        Set rsCta = New ADODB.Recordset
        If rsCta.State = 1 Then rsCta.Close
        rsCta.Open "select * from fc_Cuenta_Bancaria where cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
        If rsCta.RecordCount > 0 Then
            txtorganismo.Text = rsCta("org_codigo")
            Set rsBco = New ADODB.Recordset
            If rsBco.State = 1 Then rsCta.Close
            rsBco.Open "select * from fc_Bancos where Bco_codigo='" & rsCta("Bco_codigo") & "'", db, adOpenKeyset, adLockOptimistic
            If rsCta.RecordCount > 0 Then
                TxtBanco.Text = rsBco("Bco_descripcion_larga")
                TxtCodigoBanco.Text = rsBco("Bco_codigo")
            End If
        End If

End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
   DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
   DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
End Sub

Private Sub Form_Load()
Dim xxx As Recordset

  'Determinar las cuentas
  Set rscuenta = New ADODB.Recordset
  rscuenta.Open "select * from fc_cuenta_bancaria order by Cta_codigo_tgn", db, adOpenKeyset, adLockOptimistic
  Set AdoCuenta.Recordset = rscuenta

  Data1.Connect = "Excel 8.0"
  Data1.DatabaseName = FrmExplorador.TxtRutaNombreArchivo.Text
  Data1.RecordSource = "Hoja1$"
'  Data1.Database.Connection.Execute
  TDBGrid1.DataSource = Data1
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  
'  db.Execute "delete from fc_DatosBanco"
  'Abriendo Tabla  de registros del Banco
  Set rsBANCO = New ADODB.Recordset
  rsBANCO.Open "select * from fc_DatosBanco order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
  If rsBANCO.RecordCount > 0 Then
     Set DtGDatosBanco.DataSource = rsBANCO
  End If

  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Data1.Connect = ""
  Data1.DatabaseName = ""
  Data1.RecordSource = ""
  Print Data1.DatabaseName
  Print Data1.RecordSource
  
End Sub


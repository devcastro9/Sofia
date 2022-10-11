VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form FrmCapturaDatosBanco 
   Caption         =   "Capturando datos de banco"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   1275
      TabIndex        =   1
      Top             =   8100
      Width           =   9390
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   300
         Left            =   345
         TabIndex        =   2
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36670
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   300
         Left            =   2625
         TabIndex        =   3
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36670
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "Form1.frx":0000
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   315
         TabIndex        =   4
         Top             =   1005
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
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "Form1.frx":0018
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   315
         TabIndex        =   5
         Top             =   1350
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
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "Form1.frx":0030
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   2505
         TabIndex        =   6
         Top             =   1005
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
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. "
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha Inicio"
         Height          =   225
         Left            =   330
         TabIndex        =   8
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Fin"
         Height          =   225
         Left            =   2640
         TabIndex        =   7
         Top             =   225
         Width           =   1410
      End
   End
   Begin MSDataGridLib.DataGrid DtGDatosBanco 
      Height          =   6285
      Left            =   7275
      TabIndex        =   24
      Top             =   1785
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   11086
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
      TabIndex        =   20
      Top             =   1050
      Width           =   11130
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   2445
      End
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "Form1.frx":0048
      Height          =   6225
      Left            =   1320
      OleObjectBlob   =   "Form1.frx":005C
      TabIndex        =   0
      Top             =   1815
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
      ScaleWidth      =   15180
      TabIndex        =   10
      Top             =   0
      Width           =   15240
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3540
         TabIndex        =   15
         Top             =   135
         Width           =   5235
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   14
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9210
         TabIndex        =   13
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   12
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Top             =   675
         Width           =   1110
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   7035
      Left            =   15
      TabIndex        =   16
      Top             =   990
      Width           =   1245
      Begin VB.CommandButton CmdTransferencia 
         Caption         =   "Tranferir Datos Excel "
         Height          =   735
         Left            =   180
         TabIndex        =   23
         Top             =   1935
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   195
         Picture         =   "Form1.frx":274A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2670
         Width           =   930
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Busqueda"
         Height          =   855
         Left            =   180
         Picture         =   "Form1.frx":2B8C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Width           =   945
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Impresión"
         Height          =   885
         Left            =   180
         Picture         =   "Form1.frx":2C8E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   195
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmCapturaDatosBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBanco As New ADODB.Recordset

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub CmdTransferencia_Click()
    Data1.Recordset.MoveFirst
    While Not (Data1.Recordset.EOF)
      If Not IsNull(Data1.Recordset("fecha_pago")) Then
        'db.Execute "insert into to_DatosBanco(Nro_Cmpte) values 12 )"
        db.Execute "insert into to_datosBanco (nro_cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo,  Estado_Conciliacion, Justificacion) " & _
                   "values ('" & Int(Data1.Recordset("Nro_Cmpte")) & "',' " & Data1.Recordset("Organismo") & " ', ' " & Format(Data1.Recordset("Fecha_Pago")) & " ', '" & Data1.Recordset("Monto") & "', '" & Data1.Recordset("Cambio") & "', '" & Data1.Recordset("Beneficiario") & "', '" & Data1.Recordset("Nro_Doc") & "', '" & Data1.Recordset("Transf_Cheq") & "', '" & Data1.Recordset("Cta_Codigo") & "', '" & Data1.Recordset("Estado_Conciliacion") & " ','" & Data1.Recordset("Justificacion") & "')"
        
      Else
        
      End If
      Data1.Recordset.MoveNext
    Wend
  
  
    'Abriendo Tabla  de registros del Banco
    Set rsBanco = New ADODB.Recordset
    rsBanco.Open "select * from to_DatosBanco order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
    If rsBanco.RecordCount > 0 Then
        Set DtGDatosBanco.DataSource = rsBanco
    End If

End Sub

Private Sub Form_Load()
Dim xxx As Recordset

  'Determinar las cuentas
  Set rsCuenta = New ADODB.Recordset
  rsCuenta.Open "select * from fc_cuenta_bancaria order by Cta_codigo_tgn", db, adOpenKeyset, adLockOptimistic
  Set AdoCuenta.Recordset = rsCuenta

  Data1.Connect = "Excel 8.0"
  Data1.DatabaseName = "c:\mis documentos\fc_DatosBanco.xls"    ' "c:\mis documentos\grecocon.xls"
  Data1.RecordSource = "Hoja1$"
'  Data1.Database.Connection.Execute
  TDBGrid1.DataSource = Data1
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  
  db.Execute "delete from to_DatosBanco"
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Data1.Connect = ""
  Data1.DatabaseName = ""
  Data1.RecordSource = ""
  Print Data1.DatabaseName
  Print Data1.RecordSource
  
End Sub


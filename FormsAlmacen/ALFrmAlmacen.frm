VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ALFrmAlmacen 
   Caption         =   "Estado del Almacen"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "ALFrmAlmacen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoGrupos 
      Height          =   495
      Left            =   10800
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
   Begin MSDataGridLib.DataGrid tdbgAlmacen 
      Height          =   5175
      Left            =   0
      TabIndex        =   16
      Top             =   1560
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   9128
      _Version        =   393216
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "CodDestino"
         Caption         =   "CodAlmacen"
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
         DataField       =   "nro_licitacion"
         Caption         =   "Nro_Compra"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "CodDetalle"
         Caption         =   "Cod_Producto"
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
      BeginProperty Column03 
         DataField       =   "DescDetalle"
         Caption         =   "Descripcion del Produrcto"
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
      BeginProperty Column04 
         DataField       =   "Nro_Lote"
         Caption         =   "Nro_Lote"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "fechaVenc"
         Caption         =   "FechaVencimiento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "StockActual"
         Caption         =   "Stock Actual"
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
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4289.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1049.953
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      Picture         =   "ALFrmAlmacen.frx":6852
      ScaleHeight     =   810
      ScaleWidth      =   11340
      TabIndex        =   10
      Top             =   555
      Width           =   11400
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO DEL ALMACEN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Index           =   1
         Left            =   7275
         TabIndex        =   13
         Top             =   255
         Width           =   3615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   1605
         TabIndex        =   12
         Top             =   330
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "."
         ForeColor       =   &H0000C000&
         Height          =   180
         Left            =   4815
         TabIndex        =   11
         Top             =   675
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   0
         Picture         =   "ALFrmAlmacen.frx":9AEC
         Top             =   0
         Width           =   15360
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11400
      TabIndex        =   6
      Top             =   6735
      Width           =   11400
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   1215
         TabIndex        =   7
         Top             =   255
         Width           =   6345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Almacen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   7620
         TabIndex        =   8
         Top             =   90
         Width           =   3060
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Almacen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   7635
         TabIndex        =   9
         Top             =   105
         Width           =   3060
      End
   End
   Begin VB.PictureBox picBoton 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5310
      Left            =   0
      ScaleHeight     =   5310
      ScaleWidth      =   1020
      TabIndex        =   4
      Top             =   1425
      Width           =   1020
      Begin VB.Frame FraOpcionesDetalle 
         BorderStyle     =   0  'None
         Height          =   5730
         Left            =   15
         TabIndex        =   5
         Top             =   90
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11400
      TabIndex        =   3
      Top             =   0
      Width           =   11400
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   580
         Left            =   10645
         Picture         =   "ALFrmAlmacen.frx":B692
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
      Begin VB.Frame TDBFrame3D1 
         Height          =   495
         Left            =   1560
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   5175
         Begin MSDataListLib.DataCombo tdbcGrupos 
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   450
         Left            =   7545
         Picture         =   "ALFrmAlmacen.frx":B89C
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "Grupo"
         Height          =   450
         Left            =   900
         TabIndex        =   0
         Top             =   60
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   450
         Left            =   6750
         Picture         =   "ALFrmAlmacen.frx":BCDE
         TabIndex        =   1
         Top             =   60
         Visible         =   0   'False
         Width           =   795
      End
      Begin Crystal.CrystalReport Cry 
         Left            =   0
         Top             =   0
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
   End
End
Attribute VB_Name = "ALFrmAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAlmacen As ADODB.Recordset
Dim RsGrupos As ADODB.Recordset
Dim CodGrupo As String
'--
'JQA
'Dim ClBuscaGrid As  ClBuscaEnGridPropio

Private Sub CmdBuscar_Click()
Dim BookMark As Variant
    If RsAlmacen.RecordCount <= 0 Then MsgBox "No Existen Items en Almacen para realizar la Busqueda.", vbInformation + vbOKOnly, "Atención": Exit Sub
    If Trim(CodGrupo) <> "" Then
        BookMark = RsAlmacen.BookMark
        RsAlmacen.Find "CodDestino = '" & CodGrupo & "'"
        If RsAlmacen.EOF Then
            MsgBox "Item '" & tdbcGrupos.Text & "' no registrado en Almacen.", vbInformation + vbOKOnly, "Atención"
            RsAlmacen.BookMark = BookMark
        Else
            MsgBox "Item '" & tdbcGrupos.Text & "' encontrado.", vbInformation + vbOKOnly, "Atención"
        End If
    End If
End Sub

Private Sub cmdFiltrar_Click()
    If tdbcGrupos.Text = "" Then CodGrupo = ""
    If CodGrupo = "" Then
        RsAlmacen.Filter = adFilterNone
    Else
        RsAlmacen.Filter = "CodDestino = '" & CodGrupo & "'"
    End If
End Sub

Private Sub cmdItem_Click()
'JQA
'  Set ClBuscaGrid = New  ClBuscaEnGridPropio
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.FiltrosMultiples = True
'  ClBuscaGrid.QueryUtilizado = "SELECT CodGrupo +'-'+ CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle"
'  ClBuscaGrid.Título = "Elija una Item"
'  ClBuscaGrid.OcultarPrimero = True
'  ClBuscaGrid.Ejecutar
'  If ClBuscaGrid.ElegidoCol1 <> "" Then
'    CodGrupo = ClBuscaGrid.ElegidoCol1
'    tdbcGrupos.Text = ClBuscaGrid.ElegidoCol2
'  End If
'  Set ClBuscaGrid = Nothing
'JQA
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    With Me
        .Top = 0
        .Left = 0
    End With
    '-- JQA 05-2008
    'GlSqlAux = "SELECT CodGrupo +'-'+ CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle ORDER BY CodGrupo"
    GlSqlAux = "SELECT CodDestino+'-'+DescDestino FROM ALCLDestinos "
    Set RsGrupos = New ADODB.Recordset
    RsGrupos.Open GlSqlAux, db, adOpenStatic
    Set tdbcGrupos.RowSource = RsGrupos
    Set AdoGrupos.Recordset = RsGrupos
    '--
    'GlSqlAux = "SELECT * FROM AlClDestino_Det"
    GlSqlAux = "SELECT * FROM av_stock_almacenes"
    Set RsAlmacen = New ADODB.Recordset
    RsAlmacen.Open GlSqlAux, db, adOpenStatic
    Set tdbgAlmacen.DataSource = RsAlmacen
End Sub
Private Sub Form_Resize()
On Error Resume Next
    tdbgAlmacen.Width = Me.ScaleWidth
End Sub

Private Sub tdbcGrupos_ItemChange()
'    CodGrupo = tdbcGrupos.Columns("CodGrupo").Value
End Sub

Private Sub tdbcGrupos_NotInList(NewEntry As String, Retry As Integer)
    CodGrupo = ""
    tdbcGrupos.Text = ""
End Sub


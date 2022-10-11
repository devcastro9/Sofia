VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_almacen_inventario 
   Caption         =   "Inventario de Almacenes"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "aw_almacen_inventario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   1935
      Left            =   2160
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   6135
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   18
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox BtnImprimir1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1440
            Picture         =   "aw_almacen_inventario.frx":6852
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   20
            ToolTipText     =   "Kardex por Bien Elegido"
            Top             =   0
            Width           =   1455
         End
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3240
            Picture         =   "aw_almacen_inventario.frx":711F
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   19
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   14175
            TabIndex        =   21
            Top             =   195
            Width           =   1005
         End
      End
      Begin MSComCtl2.DTPicker DTP_Finicio 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90177537
         CurrentDate     =   42880
      End
      Begin MSComCtl2.DTPicker DTP_Ffin 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   3600
         TabIndex        =   25
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90177537
         CurrentDate     =   42880
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE INICIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   840
         TabIndex        =   24
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE FIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3600
         TabIndex        =   22
         Top             =   1080
         Width           =   1485
      End
   End
   Begin VB.PictureBox fraOpciones 
      Align           =   1  'Align Top
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11400
      TabIndex        =   8
      Top             =   0
      Width           =   11400
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17400
         Picture         =   "aw_almacen_inventario.frx":78B9
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   12
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Caption         =   "Digitaliza"
         Height          =   710
         Left            =   9000
         Picture         =   "aw_almacen_inventario.frx":807B
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2400
         Picture         =   "aw_almacen_inventario.frx":84BD
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   10
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4080
         Picture         =   "aw_almacen_inventario.frx":8C72
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   9
         ToolTipText     =   "Saldos por Almacen"
         Top             =   0
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13305
         TabIndex        =   13
         Top             =   195
         Width           =   885
      End
   End
   Begin VB.PictureBox Fra_Elegir 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11400
      TabIndex        =   7
      Top             =   660
      Width           =   11400
      Begin VB.CommandButton CmdFiltrar 
         Height          =   450
         Left            =   6390
         Picture         =   "aw_almacen_inventario.frx":953F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1275
      End
      Begin VB.CommandButton cmdItem 
         Appearance      =   0  'Flat
         Caption         =   "Elija el Almacen --->"
         Height          =   570
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "aw_almacen_inventario.frx":9E43
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   120
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "almacen_descripcion"
         BoundColumn     =   "almacen_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "aw_almacen_inventario.frx":9E5C
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   -120
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "almacen_codigo"
         BoundColumn     =   "almacen_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc_2 
         Bindings        =   "aw_almacen_inventario.frx":9E75
         DataField       =   "Item"
         Height          =   315
         Left            =   9120
         TabIndex        =   26
         Top             =   120
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Titulo"
         BoundColumn     =   "Item"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_cod_2 
         Bindings        =   "aw_almacen_inventario.frx":9E92
         DataField       =   "Item"
         Height          =   315
         Left            =   7800
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Item"
         BoundColumn     =   "Item"
         Text            =   ""
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11400
      TabIndex        =   3
      Top             =   7920
      Width           =   11400
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   1215
         TabIndex        =   4
         Top             =   255
         Width           =   6945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Inventario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   8340
         TabIndex        =   5
         Top             =   90
         Width           =   3360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Inventario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   8355
         TabIndex        =   6
         Top             =   105
         Width           =   3360
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgInventario 
      Align           =   3  'Align Left
      Bindings        =   "aw_almacen_inventario.frx":9EAF
      Height          =   6705
      Left            =   0
      OleObjectBlob   =   "aw_almacen_inventario.frx":9EC7
      TabIndex        =   2
      Top             =   1215
      Width           =   12345
   End
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   12360
      Top             =   6360
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      BackColor       =   16777215
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   12360
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Ado_datos1"
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
   Begin Crystal.CrystalReport Cry 
      Left            =   12480
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   13080
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc ado_datos_busq 
      Height          =   330
      Left            =   12360
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "ado_datos_busq"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Digite ""DOBLE CLICK"", para ver KARDEX de cada Item"
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   12480
      TabIndex        =   23
      Top             =   5040
      Width           =   1695
   End
End
Attribute VB_Name = "aw_almacen_inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim cnn As ADODB.Connection
Dim RsInventario As ADODB.Recordset
Dim rs_datos1 As ADODB.Recordset
Dim RsGrupos As ADODB.Recordset

Dim CodGrupo As String
Dim cmm As ADODB.Command
Private Sub buscar()
 If (tdbgInventario.SelBookmarks.Count <> 0) Then
            tdbgInventario.SelBookmarks.Remove 0
   End If
   
   If RsInventario.RecordCount > 0 Then
   
   RsInventario.Find "Item = '" & dtc_cod_2.Text & "'", , , 1
   
   tdbgInventario.SelBookmarks.Add (RsInventario.Bookmark)
 
 Else
 'sino = MsgBox("No se encontro ningun bien con ese nombre", vbInformation, "Aviso")
 'Call Carga_Beneficiario(1)
 'dtc_buscar_desc.Text = ""
 End If
End Sub
Private Sub BtnCancelar3_Click()
    Fra_reporte.Visible = False
    tdbgInventario.Enabled = True
    Fra_Elegir.Enabled = True
End Sub

Private Sub BtnImprimir_Click()
  If dtc_codigo1.Text <> "" Then
    'If Ado_datos.Recordset.RecordCount > 0 Then
      Dim iResult As Integer
      Screen.MousePointer = vbHourglass
      Cry.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex_tot_alm.rpt"
      Cry.StoredProcParam(0) = dtc_codigo1.Text         'Ado_datos.Recordset!almacen_codigo
      'If Trim(CodGrupo) <> "" Then
      '    Cry.SelectionFormula = "{ALInventarioFisico;1.Item} = '" & CodGrupo & "'"
      'End If
      iResult = Cry.PrintReport
      Screen.MousePointer = vbDefault
      If iResult <> 0 Then
          MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Atención"
      End If
      
'      Dim IResult As Integer
'        'Dim co As New ADODB.Command
'        CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
'        CryV01.WindowShowPrintSetupBtn = True
'        CryV01.WindowShowRefreshBtn = True
'        'CryV01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
'        CryV01.StoredProcParam(0) = Ado_datos.Recordset!Item
'        CryV01.StoredProcParam(1) = Format(DTPicker3.Value, "dd/mm/yyyy")
'        CryV01.StoredProcParam(2) = Ado_datos.Recordset!almacen_codigo            'dtc_codigo1.Text
'        DTPicker3.Value = Date
''        CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
''        VAR_TITULO = "MODULO ALMACENES"
''        CryV01.Formulas(0) = "titulo = '" & VAR_TITULO & "' "
'        CryV01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
'        CryV01.Formulas(2) = "FechaAl = '" & DTPicker3.Value & "' "
'
'        IResult = CryV01.PrintReport
'        If IResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'        CryV01.WindowState = crptMaximized
    'Else
    '      MsgBox "No se puede Imprimir. Debe elegir el Almacen y vuelva a intentar ...", , "Atención"
    'End If
  Else
        MsgBox "No se puede Imprimir. Debe elegir el Almacen y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir1_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        'CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
        CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado.rpt" '
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        'CryV01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
        CryV01.StoredProcParam(0) = Ado_datos.Recordset!Item
        CryV01.StoredProcParam(1) = Trim(Str(Ado_datos.Recordset!almacen_codigo))            'dtc_codigo1.Text
        CryV01.StoredProcParam(2) = Format(DTP_Finicio.Value, "dd/mm/yyyy")
        CryV01.StoredProcParam(3) = Format(DTP_Ffin.Value, "dd/mm/yyyy")
        
'        DTPicker3.Value = Date
''        CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
''        VAR_TITULO = "MODULO ALMACENES"
''        CryV01.Formulas(0) = "titulo = '" & VAR_TITULO & "' "
'        CryV01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
'        CryV01.Formulas(2) = "FechaAl = '" & DTPicker3.Value & "' "

        CryV01.Formulas(0) = "almace = '" & dtc_desc1.Text & "' "
        'CryV01.Formulas(2) = "DEL_AL = '' "
        'CryV01.Formulas(3) = "fechafin = '" & DTP_Ffin.Value & "' "
        
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
        Fra_reporte.Visible = False
        tdbgInventario.Enabled = True
        Fra_Elegir.Enabled = True
    Else
        MsgBox "No se puede Imprimir. Verifique si existen datos y vuelva a intentar ...", , "Atención"
    End If
    'Fra_reporte.Visible = True
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub cmdFiltrar_Click()
'    If tdbcGrupos.Text = "" Then CodGrupo = ""
'    If CodGrupo = "" Then
'        RsInventario.Filter = adFilterNone
'    Else
'        RsInventario.Filter = "Item = '" & CodGrupo & "'"
'    End If
'    Totales
    If dtc_codigo1.Text = "" Then
        MsgBox "Debe Elegir un Almacen, vuelva a intentar ...", vbInformation + vbOKOnly, "Atención"
    Else
        Set RsInventario = New ADODB.Recordset
        If RsInventario.State = 1 Then RsInventario.Close
        'RsInventario.Open "select * from AV_inventario_saldos  ", db, adOpenKeyset, adLockReadOnly
        RsInventario.Open "select * from av_almacen_inventario where almacen_codigo = " & dtc_codigo1.Text & " order by Titulo ", db, adOpenKeyset, adLockReadOnly
        Set Ado_datos.Recordset = RsInventario
        Set ado_datos_busq.Recordset = RsInventario
        dtc_cod_2.BoundText = dtc_desc_2.BoundText
'    Totales
    End If
    
    
    
End Sub

Private Sub cmdItem_Click()
'JQA
'  Set ClBuscaGrid = New ClBuscaEnGridExterno    ' ClBuscaEnGridPropio
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.FiltrosMultiples = True
'  'ClBuscaGrid.QueryUtilizado = "SELECT CodGrupo +'-'+ CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle"
'  ClBuscaGrid.QueryUtilizado = "SELECT almacen_codigo +'-'+ almacen_descripcion As CodGrupo, almacen_descripcion as DescDetalle FROM ac_almacenes where almacen_codigo <> '0' AND almacen_codigo <> '1'  "
'  ClBuscaGrid.Título = "Elija un Almacen"
'  ClBuscaGrid.OcultarPrimero = True
'  ClBuscaGrid.Ejecutar
'  If ClBuscaGrid.ElegidoCol1 <> "" Then
'    CodGrupo = ClBuscaGrid.ElegidoCol1
'    tdbcGrupos.Text = ClBuscaGrid.ElegidoCol2
'  End If
'  Set ClBuscaGrid = Nothing
'JQA
End Sub

Private Sub dtc_cod_2_Click(Area As Integer)
dtc_desc_2.BoundText = dtc_cod_2.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub



Private Sub dtc_desc_2_Change()
dtc_cod_2.BoundText = dtc_desc_2.BoundText
 If dtc_cod_2.SelectedItem <> "" Then
 'busq = busq + 1
 'If busq = 2 Then
 Call buscar
 'busq = 0
 'End If
 End If
End Sub

Private Sub dtc_desc_2_Click(Area As Integer)
dtc_cod_2.BoundText = dtc_desc_2.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    Me.Top = 0
    Me.Left = 0
    '--
    'ac_almacenes ' Origen
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select * from ac_almacenes where almacen_codigo <> '0' AND almacen_codigo <> '1' AND almacen_tipo = '" & Aux & "' ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
   
    'VAR_SW = rs_datos1.RecordCount
    '-- JQA 05-2017
'    'GlSqlAux = "SELECT CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle WHERE ESTADO = 1 "
'    GlSqlAux = "SELECT almacen_codigo As CodGrupo, almacen_descripcion as DescDetalle FROM ac_almacenes where almacen_codigo <> '0' AND almacen_codigo <> '1'  "
'    Set RsGrupos = New ADODB.Recordset
'    RsGrupos.Open GlSqlAux, db, adOpenStatic
'    Set Ado_datos1.Recordset = RsGrupos
'    'Set tdbcGrupos.RowSource = RsGrupos
    '--
'    'JQA 04/2017
'    Set RsInventario = New ADODB.Recordset
'    If RsInventario.State = 1 Then RsInventario.Close
'    'RsInventario.Open "select * from AV_inventario_saldos  ", db, adOpenKeyset, adLockReadOnly
'    RsInventario.Open "select * from av_almacen_inventario  ", db, adOpenKeyset, adLockReadOnly
''    Set adopuestosol.Recordset = RsInventario
''    adopuestosol.Refresh
'    Set tdbgInventario.DataSource = RsInventario
'    Totales

'    GlSqlAux = "ALInventarioFisico"
'    Set cmm = New ADODB.Command
'    cmm.CommandType = adCmdStoredProc
'    cmm.CommandText = GlSqlAux
'    cmm.ActiveConnection = db
'    Set RsInventario = New ADODB.Recordset
'    Set RsInventario = cmm.Execute
'    Set cmm = Nothing
    

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
On Error Resume Next
'    tdbgInventario.Width = Me.ScaleWidth - picBoton.Width
End Sub

Public Sub Totales()
Dim rs As ADODB.Recordset
Dim ValorSus As Currency
Dim PrecIng As Currency
Dim EjmIng As Long
Dim PrecSal As Long
Dim EjmEnt As Long
Dim valor As Long
Dim Ejmtot As Long
    Set rs = New ADODB.Recordset
    Set rs = RsInventario
    PrecIng = 0
    EjmIng = 0
    PrecSal = 0
    EjmEnt = 0
    valor = 0
    Ejmtot = 0
    'ValorSus = 0
    While Not rs.EOF
        'JQA 04/2008
        PrecIng = PrecIng + IIf(IsNull(rs!PrecIng), 0, rs!PrecIng)
        'CajaIng = CajaIng + 1
        EjmIng = EjmIng + IIf(IsNull(rs!EjmIng), 0, rs!EjmIng)
        PrecSal = PrecSal + IIf(IsNull(rs!PrecSal), 0, rs!PrecSal)
        'CajaEnt = CajaEnt + 1
        EjmEnt = EjmEnt + IIf(IsNull(rs!EjmEnt), 0, rs!EjmEnt)
        valor = valor + IIf(IsNull(rs!valor), 0, rs!valor)
        'CajaSal = CajaSal + 1
        'EjmSal = EjmSal + IIf(IsNull(rs!EjmSal), 0, rs!EjmSal)
        Ejmtot = EjmIng - EjmEnt
        valor = PrecIng - PrecSal
        'ValorSus = ValorSus + IIf(IsNull(rs!valor), 0, rs!valor)
        rs.MoveNext
    Wend
    tdbgInventario.Columns("Titulo").FooterText = "TOTALES"
    tdbgInventario.Columns("PrecIng").FooterText = Format(PrecIng, "###,###,##0") & ""
    tdbgInventario.Columns("EjmIng").FooterText = Format(EjmIng, "###,###,##0") & ""
    tdbgInventario.Columns("PrecSal").FooterText = Format(PrecSal, "###,###,##0") & ""
    tdbgInventario.Columns("EjmEnt").FooterText = Format(EjmEnt, "###,###,##0") & ""
    tdbgInventario.Columns("valor").FooterText = Format(valor, "###,###,##0") & ""
    tdbgInventario.Columns("Ejmtot").FooterText = Format(Ejmtot, "###,###,##0") & ""
    'tdbgInventario.Columns("Valor").FooterText = Format(ValorSus, "###,###,##0.00") & " $us"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  db.Close
'  Set db = Nothing
End Sub

Private Sub tdbcGrupos_ItemChange()
'    CodGrupo = tdbcGrupos.Columns("CodGrupo").Value
End Sub
Private Sub tdbcGrupos_NotInList(NewEntry As String, Retry As Integer)
'    CodGrupo = ""
'    tdbcGrupos.Text = ""
End Sub

Private Sub tdbgInventario_DblClick()
    Fra_reporte.Visible = True
    tdbgInventario.Enabled = False
    Fra_Elegir.Enabled = False
End Sub

VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AlmFrmInventario 
   Caption         =   "Inventario Físico Valorado"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "AlmFrmInventario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11400
      TabIndex        =   16
      Top             =   990
      Width           =   11400
      Begin VB.CommandButton CmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   450
         Left            =   6270
         Picture         =   "AlmFrmInventario.frx":6852
         TabIndex        =   4
         Top             =   60
         Width           =   795
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "Item"
         Height          =   450
         Left            =   540
         TabIndex        =   2
         Top             =   60
         Width           =   495
      End
      Begin TrueOleDBList60.TDBCombo tdbcGrupos 
         Height          =   330
         Left            =   1140
         OleObjectBlob   =   "AlmFrmInventario.frx":6C94
         TabIndex        =   3
         Top             =   60
         Width           =   5100
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11400
      TabIndex        =   10
      Top             =   7920
      Width           =   11400
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   1215
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   105
         Width           =   3360
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      Picture         =   "AlmFrmInventario.frx":8BA3
      ScaleHeight     =   930
      ScaleWidth      =   11340
      TabIndex        =   7
      Top             =   0
      Width           =   11400
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTARIO FISICO VALORADO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   345
         Index           =   0
         Left            =   7995
         TabIndex        =   15
         Top             =   240
         Width           =   4755
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   1005
         TabIndex        =   14
         Top             =   210
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
         TabIndex        =   8
         Top             =   675
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   0
         Picture         =   "AlmFrmInventario.frx":BE3D
         Top             =   0
         Width           =   15360
      End
   End
   Begin VB.PictureBox picBoton 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   1020
      TabIndex        =   6
      Top             =   1545
      Width           =   1020
      Begin VB.Frame FraOpcionesDetalle 
         BackColor       =   &H00C0E0FF&
         Height          =   5730
         Left            =   10
         TabIndex        =   9
         Top             =   120
         Width           =   990
         Begin VB.CommandButton CmdSalir 
            Caption         =   "Salir"
            Height          =   855
            Left            =   60
            Picture         =   "AlmFrmInventario.frx":D9E3
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   1035
            Width           =   855
         End
         Begin Crystal.CrystalReport Cry 
            Left            =   375
            Top             =   2085
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
         Begin VB.CommandButton CmdImprimir 
            Caption         =   "Imprimir"
            Height          =   855
            Left            =   60
            Picture         =   "AlmFrmInventario.frx":DBED
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   195
            Width           =   855
         End
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgInventario 
      Align           =   3  'Align Left
      Bindings        =   "AlmFrmInventario.frx":E2D7
      Height          =   6375
      Left            =   1020
      OleObjectBlob   =   "AlmFrmInventario.frx":E2F0
      TabIndex        =   5
      Top             =   1545
      Width           =   12345
   End
End
Attribute VB_Name = "AlmFrmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim cnn As ADODB.Connection
Dim RsInventario As ADODB.Recordset
Dim RsGrupos As ADODB.Recordset
Dim CodGrupo As String
Dim cmm As ADODB.Command
'--
'JQA
'Dim ClBuscaGrid As  ClBuscaEnGridPropio
'JQA
Private Sub cmdFiltrar_Click()
    If tdbcGrupos.Text = "" Then CodGrupo = ""
    If CodGrupo = "" Then
        RsInventario.Filter = adFilterNone
    Else
        RsInventario.Filter = "Item = '" & CodGrupo & "'"
    End If
    Totales
End Sub

Private Sub Cmdimprimir_Click()
Dim IResult As Integer
    Screen.MousePointer = vbHourglass
    Cry.ReportFileName = App.Path & "\Reportes\Almacen\ALInventarioGral.rpt"
    If Trim(CodGrupo) <> "" Then
        Cry.SelectionFormula = "{ALInventarioFisico;1.Item} = '" & CodGrupo & "'"
    End If
    IResult = Cry.PrintReport
    Screen.MousePointer = vbDefault
    If IResult <> 0 Then
        MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Atención"
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

Private Sub CmdSalirDetalle_Click()
    Unload Me
End Sub
Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    Me.Top = 0
    Me.Left = 0
    '--
    'Set db = New ADODB.Connection
    'db.Open "PROVIDER=MSDataShape;Data PROVIDER=MSDASQL;driver={SQL Server};server=" & GlServidor & ";uid=sa;pwd=;database=" & GlBaseDatos & ";"
      
    '-- JQA 05-2008
    'GlSqlAux = "SELECT CodGrupo +'-'+ CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle WHERE ESTADO = 1 "
    GlSqlAux = "SELECT CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle WHERE ESTADO = 1 "
    Set RsGrupos = New ADODB.Recordset
    RsGrupos.Open GlSqlAux, db, adOpenStatic
    Set tdbcGrupos.RowSource = RsGrupos
    '--
    'JQA 04/2008
    Set RsInventario = New ADODB.Recordset
    If RsInventario.State = 1 Then RsInventario.Close
    RsInventario.Open "select * from AV_inventario_saldos  ", db, adOpenKeyset, adLockReadOnly
'    Set adopuestosol.Recordset = RsInventario
'    adopuestosol.Refresh

'    GlSqlAux = "ALInventarioFisico"
'    Set cmm = New ADODB.Command
'    cmm.CommandType = adCmdStoredProc
'    cmm.CommandText = GlSqlAux
'    cmm.ActiveConnection = db
'    Set RsInventario = New ADODB.Recordset
'    Set RsInventario = cmm.Execute
'    Set cmm = Nothing
    Set tdbgInventario.DataSource = RsInventario
    Totales
    Screen.MousePointer = vbDefault
	Call SeguridadSet(Me)
End Sub
Private Sub Form_Resize()
On Error Resume Next
    tdbgInventario.Width = Me.ScaleWidth - picBoton.Width
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
    CodGrupo = tdbcGrupos.Columns("CodGrupo").Value
End Sub
Private Sub tdbcGrupos_NotInList(NewEntry As String, Retry As Integer)
    CodGrupo = ""
    tdbcGrupos.Text = ""
End Sub


VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AlFrmExistencia_Almacen 
   Caption         =   "Estado Almacen"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   Icon            =   "AlFrmExistencia_Almacen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   10455
      TabIndex        =   14
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton CmdBusCabeza 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   6885
         Picture         =   "AlFrmExistencia_Almacen.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   75
         Width           =   660
      End
      Begin VB.CommandButton CmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   450
         Left            =   6000
         Picture         =   "AlFrmExistencia_Almacen.frx":10D4
         TabIndex        =   17
         Top             =   60
         Width           =   795
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "Item"
         Height          =   450
         Left            =   120
         TabIndex        =   16
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   450
         Left            =   9600
         Picture         =   "AlFrmExistencia_Almacen.frx":1516
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   795
      End
      Begin TrueOleDBList60.TDBCombo tdbcGrupos 
         Height          =   330
         Left            =   750
         OleObjectBlob   =   "AlFrmExistencia_Almacen.frx":1958
         TabIndex        =   19
         Top             =   120
         Width           =   5100
      End
   End
   Begin VB.PictureBox picBoton 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6825
      Left            =   0
      ScaleHeight     =   6825
      ScaleWidth      =   1020
      TabIndex        =   11
      Top             =   1785
      Width           =   1020
      Begin VB.Frame FraOpcionesDetalle 
         BorderStyle     =   0  'None
         Height          =   5730
         Left            =   0
         TabIndex        =   12
         Top             =   90
         Width           =   990
         Begin VB.CommandButton CmdSalir 
            Caption         =   "Salir"
            Height          =   855
            Left            =   75
            Picture         =   "AlFrmExistencia_Almacen.frx":3867
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   60
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
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10455
      TabIndex        =   7
      Top             =   8610
      Width           =   10455
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   1440
         TabIndex        =   8
         Top             =   255
         Width           =   6945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen"
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
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      Picture         =   "AlFrmExistencia_Almacen.frx":3A71
      ScaleHeight     =   930
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   795
      Width           =   10455
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   90
         TabIndex        =   5
         Top             =   570
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   195
         Left            =   945
         TabIndex        =   4
         Top             =   600
         Width           =   2310
      End
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
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   1
         Left            =   5115
         TabIndex        =   1
         Top             =   255
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "AlFrmExistencia_Almacen.frx":6D0B
         Top             =   0
         Width           =   11640
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   1005
         TabIndex        =   2
         Top             =   450
         Width           =   540
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         Caption         =   "."
         ForeColor       =   &H0000C000&
         Height          =   180
         Left            =   4815
         TabIndex        =   6
         Top             =   675
         Width           =   2655
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgAlmacen 
      Align           =   3  'Align Left
      Height          =   6825
      Left            =   1020
      OleObjectBlob   =   "AlFrmExistencia_Almacen.frx":2DD7B
      TabIndex        =   18
      Top             =   1785
      Width           =   8025
   End
End
Attribute VB_Name = "AlFrmExistencia_Almacen"
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

Private Sub CmdBusCabeza_Click()
Dim BookMark As Variant
    If RsAlmacen.RecordCount <= 0 Then MsgBox "No Existen Items en Almacen para realizar la Busqueda.", vbInformation + vbOKOnly, "Atención": Exit Sub
    If Trim(CodGrupo) <> "" Then
        BookMark = RsAlmacen.BookMark
        RsAlmacen.Find "CodArt = '" & CodGrupo & "'"
        If RsAlmacen.EOF Then
            MsgBox "Item '" & tdbcGrupos.Text & "' no registrado en Almacen.", vbInformation + vbOKOnly, "Atención"
            RsAlmacen.BookMark = BookMark
        Else
            MsgBox "Item '" & tdbcGrupos.Text & "' encontrado.", vbInformation + vbOKOnly, "Atención"
        End If
    End If
End Sub

Private Sub CmdBuscar_Click()
Dim BookMark As Variant
    If RsAlmacen.RecordCount <= 0 Then MsgBox "No Existen Items en Almacen para realizar la Busqueda.", vbInformation + vbOKOnly, "Atención": Exit Sub
    If Trim(CodGrupo) <> "" Then
        BookMark = RsAlmacen.BookMark
        RsAlmacen.Find "CodArt = '" & CodGrupo & "'"
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
        RsAlmacen.Filter = "CodArt = '" & CodGrupo & "'"
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
    '--
    GlSqlAux = "SELECT CodGrupo +'-'+ CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle ORDER BY CodGrupo"
    
    Set RsGrupos = New ADODB.Recordset
    RsGrupos.Open GlSqlAux, DB, adOpenStatic
    Set tdbcGrupos.RowSource = RsGrupos
    '--
    GlSqlAux = "SELECT * FROM ALMaterial"
    Set RsAlmacen = New ADODB.Recordset
    RsAlmacen.Open GlSqlAux, DB, adOpenStatic
    Set tdbgAlmacen.DataSource = RsAlmacen
	Call SeguridadSet(Me)
End Sub
Private Sub Form_Resize()
On Error Resume Next
    tdbgAlmacen.Width = Me.ScaleWidth
End Sub

Private Sub tdbcGrupos_ItemChange()
    CodGrupo = tdbcGrupos.Columns("CodGrupo").Value
End Sub

Private Sub tdbcGrupos_NotInList(NewEntry As String, Retry As Integer)
    CodGrupo = ""
    tdbcGrupos.Text = ""
End Sub



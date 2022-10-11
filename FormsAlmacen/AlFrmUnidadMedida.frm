VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AlFrmUnidadMedida 
   Caption         =   "Clasificadores - Almacenes  - Unidades de Medida"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   Icon            =   "AlFrmUnidadMedida.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10395
   WindowState     =   2  'Maximized
   Begin VB.Frame frmabm 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   10335
      Begin VB.OptionButton OptActivos 
         BackColor       =   &H00808000&
         Caption         =   "APROBADOS"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00808000&
         Caption         =   "TODOS"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton CmdAddCabeza 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   3480
         Picture         =   "AlFrmUnidadMedida.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdImpCabeza 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   7320
         Picture         =   "AlFrmUnidadMedida.frx":D340
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdSalCabeza 
         Caption         =   "Salir"
         Height          =   720
         Left            =   8880
         Picture         =   "AlFrmUnidadMedida.frx":DA2A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdDelCabeza 
         Caption         =   "Borrar"
         Height          =   720
         Left            =   5400
         Picture         =   "AlFrmUnidadMedida.frx":DC34
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdBusCabeza 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   6360
         Picture         =   "AlFrmUnidadMedida.frx":E31E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdModCabeza 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   4440
         Picture         =   "AlFrmUnidadMedida.frx":E528
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   765
      End
      Begin Crystal.CrystalReport CryF01 
         Left            =   5280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame FrmDatos 
      BackColor       =   &H00C0C0C0&
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   10275
      Begin VB.TextBox TextCOD_Umedida 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Textdescri 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   840
         Width           =   6255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "APROBADO"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SIN APROBAR"
         Height          =   255
         Left            =   4215
         TabIndex        =   3
         Top             =   1365
         Width           =   1560
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   270
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   270
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1665
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   270
         Index           =   4
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      Picture         =   "AlFrmUnidadMedida.frx":E732
      ScaleHeight     =   930
      ScaleWidth      =   10335
      TabIndex        =   0
      Top             =   0
      Width           =   10395
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDADES DE MEDIDA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   345
         Index           =   0
         Left            =   6645
         TabIndex        =   1
         Top             =   240
         Width           =   3315
      End
   End
   Begin MSAdodcLib.Adodc AdodcTabla 
      Height          =   375
      Left            =   0
      Top             =   7680
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
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
      Caption         =   "Unidades de Medida"
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
   Begin MSDataGridLib.DataGrid DtgMain 
      Bindings        =   "AlFrmUnidadMedida.frx":102D8
      Height          =   5655
      Left            =   0
      TabIndex        =   19
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9975
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Unidad"
         Caption         =   "Codigo"
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
      BeginProperty Column01 
         DataField       =   "Descripcion"
         Caption         =   "Descripción (Unidad de Medida)"
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
         DataField       =   "estatus"
         Caption         =   "Estado"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5235.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   675.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmgrabcabeza 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   10275
      Begin VB.CommandButton CmdGraCabeza 
         Caption         =   "Grabar"
         Height          =   675
         Left            =   4440
         Picture         =   "AlFrmUnidadMedida.frx":102F1
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdCanCabeza 
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   5640
         Picture         =   "AlFrmUnidadMedida.frx":104FB
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   765
      End
   End
End
Attribute VB_Name = "AlFrmUnidadMedida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTabla As New ADODB.Recordset
Dim queryinicial As String
Dim swgraba As Integer
Dim Marca As BookmarkEnum
Dim PosibleApliqueFiltro As Boolean
Dim ClBuscaGrid As ClBuscaEnGridExterno


Private Sub CmdAddCabeza_Click()
'adicion
Dim cod_medida As String
DtgMain.Visible = False
FrmDatos.Visible = True
frmabm.Visible = False
frmgrabcabeza.Visible = True
Option1 = True
'saca  correlativo
'DE.dbo_AL_MAXCOD_Umedida cod_medida
'TextCOD_Umedida = cod_medida + 1
swgraba = 1
Textdescri = ""
TextCOD_Umedida.SetFocus
If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
    Marca = AdodcTabla.Recordset.BookMark
End If
End Sub

Private Sub CmdBusCabeza_Click()
'BUSQUEDA
' Dim ClBuscaSec As ClBuscaSecuencialEnRS
 
  PosibleApliqueFiltro = False
  Dim rsNada As ADODB.Recordset
  Dim GrSqlAux As String
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = DtgMain
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = AdodcTabla.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True
End Sub

Private Sub CmdCanCabeza_Click()
DtgMain.Visible = True
frmabm.Visible = True
FrmDatos.Visible = False
frmgrabcabeza.Visible = False
If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
    AdodcTabla.Recordset.Move Marca - 1
End If
swgraba = 0
End Sub

Private Sub CmdDelCabeza_Click()
  On Error GoTo UpdateErr
   If AdodcTabla.Recordset.RecordCount > 0 Then
      If ExisteReg(AdodcTabla.Recordset!Unidad) Then MsgBox "No se puede eliminar la UNIDAD DE MEDIDA que ya fue utilizada en un BIEN o SERVICIO.", vbInformation + vbOKOnly, "Atención": Exit Sub
      sino = MsgBox("Está Seguro de ELIMINAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         AdodcTabla.Recordset.Delete
         AdodcTabla.Recordset.Requery
         'AdodcTabla.Recordset!estado = "E"
         'AdodcTabla.Recordset.Update
      End If
   Else
        MsgBox "No existen registros.", vbExclamation, "Atención"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdGraCabeza_Click()
Dim estatus2 As String
DtgMain.Visible = True
frmabm.Visible = True
FrmDatos.Visible = False
frmgrabcabeza.Visible = False
' grabar
If swgraba = 1 Then
    If TextCOD_Umedida.Text = "" Then
        MsgBox "Ingrese la Unidad de Medida..."
        Exit Sub
    Else
    'TextCOD_Umedida = AdodcTabla.Recordset.RecordCount + 1
    ' cod_medida + 1
        db.Execute "insert  into Al_UnidadMedida (Unidad ,descripcion, Estatus ) Values ('" & TextCOD_Umedida & "','" & Textdescri & "','S')"
    'DE.dbo_al_inserta_Umedida Textdescri
        Option1 = True
    End If
    AdodcTabla.Refresh
    AdodcTabla.Recordset.MoveLast
End If
'modificar
If swgraba = 2 Then
    If Option1 = True Then
        estatus2 = "S"
    End If
    If Option2 = True Then
        estatus2 = "N"
    End If
    'PROC ALM Modifica Marcas
    db.Execute "UPDATE Al_UnidadMedida SET descripcion='" & Textdescri & "', estatus ='" & estatus2 & "' WHERE Unidad ='" & TextCOD_Umedida & "'"
    'DE.dbo_al_Modi_Umedida AdodcTabla.Recordset!Unidad, Textdescri, estatus
    AdodcTabla.Refresh
    If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
        AdodcTabla.Recordset.Move Marca - 1
    End If
End If
End Sub

Private Sub CmdModCabeza_Click()
'MODIFICAR
 If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
    Dim cod_medida As String
    DtgMain.Visible = False
    FrmDatos.Visible = True
    frmabm.Visible = False
    frmgrabcabeza.Visible = True
    'muestra datos
    TextCOD_Umedida = AdodcTabla.Recordset!Unidad
    Textdescri = AdodcTabla.Recordset!descripcion
    If AdodcTabla.Recordset!estatus = "S" Then
        Option1 = True
    Else
        Option2 = True
    End If
    'Bandera para modificar
    swgraba = 2
    Textdescri.SetFocus
    Marca = AdodcTabla.Recordset.BookMark
Else
MsgBox "No existen registros", vbCritical, "Atencion"
End If
End Sub

Private Sub CmdSalCabeza_Click()
Unload Me
End Sub

Private Sub Form_Load()
    OptActivos = True
    swgraba = 0
    DtgMain.Visible = True
    FrmDatos.Visible = False
    frmabm.Visible = True
    frmgrabcabeza.Visible = False
    Set rsTabla = New ADODB.Recordset
    If rsTabla.State = 1 Then rsTabla.Close
    queryinicial = "select * from Al_UnidadMedida "
    'rsTabla.Open queryinicial & " order by CAST(Unidad AS INT) ", db, adOpenKeyset, adLockOptimistic
    rsTabla.Open queryinicial & " order by Unidad  ", db, adOpenKeyset, adLockOptimistic
    
    Set AdodcTabla.Recordset = rsTabla
End Sub

Private Sub OptActivos_Click()
Set rsTabla = New ADODB.Recordset
    If rsTabla.State = 1 Then rsTabla.Close
    queryinicial = "select * from Al_UnidadMedida where estatus ='S'"
    rsTabla.Open queryinicial & " order by Unidad ", db, adOpenKeyset, adLockOptimistic
    Set AdodcTabla.Recordset = rsTabla
End Sub

Private Sub Option3_Click()
Set rsTabla = New ADODB.Recordset
    If rsTabla.State = 1 Then rsTabla.Close
    queryinicial = "select * from Al_UnidadMedida "
    rsTabla.Open queryinicial & " order by Unidad ", db, adOpenKeyset, adLockOptimistic
    Set AdodcTabla.Recordset = rsTabla
End Sub

Private Function ExisteReg(Unidad As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ALCLDetalle WHERE Unidad = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function


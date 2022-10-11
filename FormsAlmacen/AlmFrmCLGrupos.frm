VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AlmFrmCLGrupos 
   Caption         =   "Mesa de Entrada - Clasificadores - Almacenes - Grupos"
   ClientHeight    =   8865
   ClientLeft      =   9015
   ClientTop       =   9210
   ClientWidth     =   12300
   Icon            =   "AlmFrmCLGrupos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12300
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdodcTabla 
      Height          =   375
      Left            =   0
      Top             =   7560
      Width           =   11175
      _ExtentX        =   19711
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
      Caption         =   "Grupos de Bienes"
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
   Begin VB.Frame frmabm 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
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
      TabIndex        =   23
      Top             =   720
      Width           =   11175
      Begin VB.CommandButton cmdAprueba 
         BackColor       =   &H0080C0FF&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   6480
         Picture         =   "AlmFrmCLGrupos.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Aprueba Registro"
         Top             =   160
         Width           =   770
      End
      Begin VB.CommandButton CmdBusCabeza 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3960
         Picture         =   "AlmFrmCLGrupos.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   160
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdDelCabeza 
         Caption         =   "Anular"
         Height          =   720
         Left            =   2880
         Picture         =   "AlmFrmCLGrupos.frx":0E16
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   160
         Width           =   765
      End
      Begin VB.CommandButton CmdSalCabeza 
         Caption         =   "Salir"
         Height          =   720
         Left            =   8160
         Picture         =   "AlmFrmCLGrupos.frx":1500
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   160
         Width           =   765
      End
      Begin VB.CommandButton CmdImpCabeza 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4920
         Picture         =   "AlmFrmCLGrupos.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   160
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdAddCabeza 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   960
         Picture         =   "AlmFrmCLGrupos.frx":1DF4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Nuevo Registro"
         Top             =   160
         Width           =   765
      End
      Begin VB.CommandButton CmdModCabeza 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   1920
         Picture         =   "AlmFrmCLGrupos.frx":88E2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   160
         Width           =   765
      End
      Begin Crystal.CrystalReport CryF01 
         Left            =   5520
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
   End
   Begin VB.Frame frmgrabcabeza 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   11235
      Begin VB.CommandButton CmdCanCabeza 
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   5520
         Picture         =   "AlmFrmCLGrupos.frx":91AC
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdGraCabeza 
         Caption         =   "Grabar"
         Height          =   675
         Left            =   4080
         Picture         =   "AlmFrmCLGrupos.frx":93B6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame FrmDatos 
      BackColor       =   &H00C0C0C0&
      Height          =   3375
      Left            =   480
      TabIndex        =   10
      Top             =   2880
      Width           =   10215
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   31
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Aprobado"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   2805
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobado"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   2805
         Width           =   1215
      End
      Begin VB.TextBox Textdescri 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   840
         Width           =   7455
      End
      Begin VB.TextBox TxtGrupo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtDescripAnt 
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   7455
      End
      Begin MSDataListLib.DataCombo DtcUnidad 
         Bindings        =   "AlmFrmCLGrupos.frx":95C0
         DataField       =   "codigo_unidad"
         DataSource      =   "AdodcTabla"
         Height          =   315
         Left            =   2400
         TabIndex        =   34
         Top             =   2280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_unidad"
         BoundColumn     =   "codigo_unidad"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcUnidadDes 
         Bindings        =   "AlmFrmCLGrupos.frx":95D8
         DataField       =   "codigo_unidad"
         DataSource      =   "AdodcTabla"
         Height          =   315
         Left            =   4920
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Uni_descripcion_larga"
         BoundColumn     =   "codigo_unidad"
         Text            =   ""
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Relacionada:"
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
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   33
         Top             =   2400
         Width           =   1710
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIGLA DEL GRUPO:"
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
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO DE REGISTRO:"
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
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   19
         Top             =   2805
         Width           =   1860
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION GRUPO:"
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
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   18
         Top             =   885
         Width           =   1800
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO DEL GRUPO:"
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
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   405
         Width           =   1680
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Anterior:"
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
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   16
         Top             =   1365
         Visible         =   0   'False
         Width           =   1770
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12300
      TabIndex        =   5
      Top             =   8370
      Width           =   12300
      Begin VB.Frame Frame3 
         Height          =   60
         Left            =   15
         TabIndex        =   6
         Top             =   255
         Width           =   6930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
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
         Index           =   1
         Left            =   7020
         TabIndex        =   7
         Top             =   75
         Width           =   1845
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
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
         Index           =   0
         Left            =   7035
         TabIndex        =   8
         Top             =   90
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      Picture         =   "AlmFrmCLGrupos.frx":95F0
      ScaleHeight     =   795
      ScaleWidth      =   12240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   12300
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO DE BIENES"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   405
         Index           =   0
         Left            =   7230
         TabIndex        =   9
         Top             =   120
         Width           =   2985
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIMOSUD"
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   1185
         TabIndex        =   4
         Top             =   435
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   1185
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "."
         ForeColor       =   &H0000C000&
         Height          =   180
         Left            =   2595
         TabIndex        =   2
         Top             =   675
         Width           =   2655
      End
   End
   Begin TrueOleDBGrid60.TDBGrid DtgMain 
      Height          =   5760
      Left            =   0
      OleObjectBlob   =   "AlmFrmCLGrupos.frx":B196
      TabIndex        =   0
      Top             =   1785
      Width           =   11195
   End
   Begin MSAdodcLib.Adodc AdoGRUPO 
      Height          =   375
      Left            =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "AdoGRUPO"
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
   Begin MSAdodcLib.Adodc AdoUnidad 
      Height          =   330
      Left            =   2160
      Top             =   7920
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
      Caption         =   "AdoUnidad"
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
Attribute VB_Name = "AlmFrmCLGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupos As ADODB.Recordset
Dim rstbeneaux As ADODB.Recordset
Dim rs_GRUPO As ADODB.Recordset
Dim rs_unidad_ejecutora As New ADODB.Recordset
Dim rscorrelativo As New ADODB.Recordset

Dim cod_grupo, SQL_FOR As String
Dim swgraba As Integer
Dim SW As Boolean
Dim Marca As BookmarkEnum
Dim sino As String
Dim num_comprobante As Integer ' variable donde se almacena correlativo de comprobante
'Buscador
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim queryinicial As String

Private Sub CmdAddCabeza_Click()
    'adicion
    DtgMain.Visible = False
    FrmDatos.Visible = True
    frmabm.Visible = False
    frmgrabcabeza.Visible = True
    Option1 = True
    'saca  correlativo
    'DE.dbo_AL_MAXCOD_Montador cod_MONTADOR
    'TextCOD_MONTADOR = cod_MONTADOR + 1
    swgraba = 1
    Textdescri = ""
    Textdescri.SetFocus
'    If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
'        Marca = AdodcTabla.Recordset.BookMark
'    End If
End Sub

Private Sub cmdAprueba_Click()
   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If AdodcTabla.Recordset("activo") = "N" Then
      If sino = vbYes Then
         Dim RUTA1, RUTA2 As String
         RUTA1 = "BIENES" + "\" + Trim(AdodcTabla.Recordset("CodGrupo"))
         MsgBox RUTA1
         MkDir RUTA1
'        MkDir RUTA1 + "\CONTRATOS"
'        MkDir RUTA1 + "\FINIQUITO"
'        MkDir RUTA1 + "\MEMORANDUMS"
'        MkDir RUTA1 + "\RESPALDOS"
'        MkDir RUTA1 + "\HOJA_VIDA"
'        MkDir RUTA1 + "\OTROS"
'        MkDir RUTA1 + "\EVALUACIONES"
'        MkDir RUTA1 + "\LICENCIAS"
'        MkDir RUTA1 + "\VACACIONES"
        AdodcTabla.Recordset("activo") = "S"
        AdodcTabla.Recordset("EsALMACEN") = 1
'        adoLista.Recordset("usr_aprueba") = GlUsuario
        AdodcTabla.Recordset.Update
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub CmdBusCabeza_Click()
  PosibleApliqueFiltro = False
  'Dim GrSqlAux As String
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = DtgMain
  ClBuscaGrid.QueryUtilizado = queryinicial
  'Set ClBuscaGrid.RecordsetTrabajo = AdodcTabla.Recordset
  Set ClBuscaGrid.RecordsetTrabajo = RsGrupos.DataSource
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
    If RsGrupos.RecordCount > 0 Then
        If RsGrupos.BOF Or RsGrupos.EOF Then Exit Sub
        If ExisteGrupo(RsGrupos!CodGrupo) Then MsgBox "No se puede eliminar al Grupo que ya tiene registrado un SUB-GRUPO.", vbInformation + vbOKOnly, "Atención": Exit Sub
        If MsgBox("Esta seguro de eliminar el Grupo seleccionado.", vbQuestion + vbYesNo, "Atención") = vbYes Then
            'RsGrupos.Delete
            RsGrupos!ACTIVO = "E"
            RsGrupos.Update
            RsGrupos.Requery
        End If
    Else
        MsgBox "No existen registros.", vbExclamation, "Atención"
    End If

End Sub

Public Sub genera_codigo()
  Set rscorrelativo = New ADODB.Recordset
  rscorrelativo.CursorLocation = adUseClient
  If rscorrelativo.State = 1 Then rscorrelativo.Close
  rscorrelativo.Open "SELECT numero_correlativo, tipo_tramite FROM fc_correl WHERE (tipo_tramite = 'GRUPO')", db, adOpenKeyset, adLockOptimistic
  If rscorrelativo.RecordCount <> 0 Then
    rscorrelativo.MoveFirst
    num_comprobante = rscorrelativo!numero_correlativo + 1
    rscorrelativo!numero_correlativo = rscorrelativo!numero_correlativo + 1
    rscorrelativo.Update
  Else
    num_comprobante = 1
    rscorrelativo!numero_correlativo = 1
    rscorrelativo.Update
  End If
End Sub

Private Sub CmdGraCabeza_Click()
Dim estatus2 As String
DtgMain.Visible = True
frmabm.Visible = True
FrmDatos.Visible = False
frmgrabcabeza.Visible = False
' grabar
If swgraba = 1 Then
    Call genera_codigo
'    DE.dbo_AL_MAXCOD_Montador cod_MONTADOR
'    TextCOD_MONTADOR = cod_MONTADOR + 1
    'TxtGrupo = AdodcTabla.Recordset.RecordCount + 1
    If num_comprobante < 10 Then
        TxtGrupo = "0" + Trim(Str(num_comprobante))
    Else
        TxtGrupo = num_comprobante
    End If
    Set rstbeneaux = New ADODB.Recordset
    SQL_FOR = "select * from ALCLGrupo where CodGrupo = '" & TxtGrupo.Text & "'  "
    rstbeneaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
    If rstbeneaux.RecordCount > 0 Then
        SW = True
        MsgBox " CODIGO DUPLICADO"
        Exit Sub
    End If

    db.Execute "INSERT INTO ALCLGrupo (CodGrupo, Grp_Codigo, DescGrupo, codigo_unidad, EsALMACEN, activo) VALUES ('" & TxtGrupo & "', '" & TxtGrupo & "', '" & Textdescri & "', '" & DtcUnidad & "', '1',  'S')  "
    
    'DE.dbo_al_inserta_montador Textdescri
    
    Option1 = True
    Call ABRIR_TABLA
    'AdodcTabla.Refresh
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
    
    db.Execute "UPDATE ALCLGrupo SET DescGrupo='" & Textdescri & "', activo='" & estatus2 & "', codigo_unidad='" & DtcUnidad & "' WHERE CodGrupo='" & TxtGrupo & "'"
    Call ABRIR_TABLA
'    AdodcTabla.Refresh
    If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
        AdodcTabla.Recordset.Move Marca - 1
    End If
End If

End Sub

Private Sub CmdModCabeza_Click()
'    DtgMain.Enabled = True
    'MODIFICAR
 If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
    DtgMain.Visible = False
    FrmDatos.Visible = True
    frmabm.Visible = False
    frmgrabcabeza.Visible = True
'    DtcGrupoD.Enabled = True
    'muestra datos
'    TxtGrupo = AdodcTabla.Recordset!CodGrupo
    TxtGrupo = AdodcTabla.Recordset!CodGrupo
    Textdescri = AdodcTabla.Recordset!descgrupo
    If AdodcTabla.Recordset!ACTIVO = "S" Then
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
On Error GoTo QError
    If Not (RsGrupos.BOF Or RsGrupos.EOF) Then
        If RsGrupos.EditMode <> adEditNone Then RsGrupos.Update
    End If
    Unload Me
    Exit Sub
QError:
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub dtcUnidad_Click(Area As Integer)
    DtcUnidadDes.BoundText = DtcUnidad.BoundText
End Sub

Private Sub DtcUnidadDes_Click(Area As Integer)
    DtcUnidad.BoundText = DtcUnidadDes.BoundText
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 3315
    Me.Width = 7365
    '
    FrmDatos.Visible = False
    Call ABRIR_TABLA
    
    Set rs_GRUPO = New ADODB.Recordset
    rs_GRUPO.Open "SELECT * FROM fc_partida_gasto WHERE par_activo='1' OR par_activo='S' order by par_descripcion_larga", db, adOpenStatic
    Set AdoGRUPO.Recordset = rs_GRUPO
    
   Set rs_unidad_ejecutora = New ADODB.Recordset
   If rs_unidad_ejecutora.State = 1 Then rs_unidad_ejecutora.Close
   rs_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora WHERE UNI_ACTIVO='S' ", db, adOpenKeyset, adLockReadOnly
   Set AdoUnidad.Recordset = rs_unidad_ejecutora
   AdoUnidad.Refresh
End Sub

Private Sub ABRIR_TABLA()
    'queryinicial = "SELECT * FROM ALCLGrupo ORDER BY CodGrupo"
    Set RsGrupos = New ADODB.Recordset
    If RsGrupos.State = 1 Then RsGrupos.Close
    queryinicial = "SELECT * FROM ALCLGrupo WHERE ACTIVO <> 'E' "
    RsGrupos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
    RsGrupos.Sort = "CodGrupo"
    DtgMain.DataSource = RsGrupos
    Set AdodcTabla.Recordset = RsGrupos
    
'   Set rstbeneficiario = New ADODB.Recordset
'   If rstbeneficiario.State = 1 Then rstbeneficiario.Close
'   queryinicial = "select * from fc_Beneficiario WHERE codigo_beneficiario <> ' ' and codigo_beneficiario <> '-' and tipo_beneficiario < 20 "
'   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
'   rstbeneficiario.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
'   rstbeneficiario.Sort = "denominacion_beneficiario"
'   Set adoLista.Recordset = rstbeneficiario
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
'    DtgMain.Width = Me.ScaleWidth - picBoton.Width
End Sub

Private Function ExisteGrupo(CodGrupo As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'queryinicial = "SELECT Count(*) AS Cuantos FROM ALCLDetalle WHERE CodGrupo = '" & CodGrupo & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM AL_Montador WHERE CodGrupo = '" & CodGrupo & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteGrupo = rs!Cuantos > 0
End Function

Private Sub Form_Unload(Cancel As Integer)
    DtgMain.Enabled = False
End Sub


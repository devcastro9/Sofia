VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ALFrm_montado 
   Caption         =   "Mesa de Entrada - Clasificadores - Almacenes  - Sub-Grupo"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   Icon            =   "ALFrm_montado.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   10560
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
      TabIndex        =   23
      Top             =   840
      Width           =   10455
      Begin VB.CommandButton CmdModCabeza 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   4680
         Picture         =   "ALFrm_montado.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   160
         Width           =   765
      End
      Begin VB.CommandButton CmdAddCabeza 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   3720
         Picture         =   "ALFrm_montado.frx":711C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo Registro"
         Top             =   160
         Width           =   765
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00808000&
         Caption         =   "TODOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptActivos 
         BackColor       =   &H00808000&
         Caption         =   "PENDIENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton CmdImpCabeza 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   7560
         Picture         =   "ALFrm_montado.frx":DC0A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   160
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdSalCabeza 
         Caption         =   "Salir"
         Height          =   720
         Left            =   9240
         Picture         =   "ALFrm_montado.frx":E2F4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   160
         Width           =   765
      End
      Begin VB.CommandButton CmdDelCabeza 
         Caption         =   "Anular"
         Height          =   720
         Left            =   5640
         Picture         =   "ALFrm_montado.frx":E4FE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   160
         Width           =   765
      End
      Begin VB.CommandButton CmdBusCabeza 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   6600
         Picture         =   "ALFrm_montado.frx":EBE8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   160
         Width           =   765
      End
      Begin Crystal.CrystalReport CryLista 
         Left            =   8520
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO DE REGISTROS"
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
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   860
      Left            =   0
      Picture         =   "ALFrm_montado.frx":EDF2
      ScaleHeight     =   795
      ScaleWidth      =   10500
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   10560
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-GRUPO DE BIENES"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   405
         Index           =   0
         Left            =   6450
         TabIndex        =   22
         Top             =   120
         Width           =   3765
      End
   End
   Begin VB.Frame FrmDatos 
      BackColor       =   &H00C0C0C0&
      Height          =   3135
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   10335
      Begin MSDataListLib.DataCombo DtcGrupoD 
         Bindings        =   "ALFrm_montado.frx":10998
         DataField       =   "CodGrupo"
         DataSource      =   "AdodcTabla"
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777152
         ListField       =   "DescGrupo"
         BoundColumn     =   "CodGrupo"
         Text            =   "-"
      End
      Begin VB.TextBox TxtDescripAnt 
         DataField       =   "CD_OCA"
         DataSource      =   "AdodcTabla"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   1680
         Width           =   7455
      End
      Begin VB.TextBox TextCOD_MONTADOR 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "Cod_montador"
         DataSource      =   "AdodcTabla"
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Textdescri 
         DataField       =   "DESCRIPCION"
         DataSource      =   "AdodcTabla"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1200
         Width           =   7455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobado"
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Aprobado"
         Height          =   255
         Left            =   3855
         TabIndex        =   14
         Top             =   2685
         Width           =   1440
      End
      Begin MSDataListLib.DataCombo DtcGrupo 
         Bindings        =   "ALFrm_montado.frx":109AF
         DataField       =   "CodGrupo"
         DataSource      =   "AdodcTabla"
         Height          =   315
         Left            =   2400
         TabIndex        =   29
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         BackColor       =   12632256
         ListField       =   "CodGrupo"
         BoundColumn     =   "CodGrupo"
         Text            =   "-"
      End
      Begin MSDataListLib.DataCombo DtcParD 
         Bindings        =   "ALFrm_montado.frx":109C6
         DataField       =   "par_codigo"
         DataSource      =   "AdodcTabla"
         Height          =   315
         Left            =   3840
         TabIndex        =   12
         Top             =   2160
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777152
         ListField       =   "par_descripcion_larga"
         BoundColumn     =   "par_codigo"
         Text            =   "-"
      End
      Begin MSDataListLib.DataCombo DtcPar 
         Bindings        =   "ALFrm_montado.frx":109DF
         DataField       =   "par_codigo"
         DataSource      =   "AdodcTabla"
         Height          =   315
         Left            =   2400
         TabIndex        =   30
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         BackColor       =   12632256
         ListField       =   "par_codigo"
         BoundColumn     =   "par_codigo"
         Text            =   "-"
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partida Relacionada:"
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
         Left            =   600
         TabIndex        =   31
         Top             =   2205
         Width           =   1740
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion Anterior:"
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
         Left            =   600
         TabIndex        =   28
         Top             =   1725
         Width           =   1770
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO DE BIENES:"
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
         Left            =   600
         TabIndex        =   27
         Top             =   520
         Width           =   1500
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código SubGrupo:"
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
         Left            =   600
         TabIndex        =   20
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción del Sub-Grupo"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   960
         Width           =   2220
      End
      Begin VB.Label LblCabecera 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Registro:"
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
         Left            =   600
         TabIndex        =   18
         Top             =   2680
         Width           =   1680
      End
   End
   Begin MSAdodcLib.Adodc AdodcTabla 
      Height          =   375
      Left            =   120
      Top             =   7800
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
      Caption         =   "Grupo al cual pertenece un Sub-Grupo"
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
      Bindings        =   "ALFrm_montado.frx":109F8
      Height          =   5895
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10398
      _Version        =   393216
      BackColor       =   12640511
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CodGrupo"
         Caption         =   "Cod.Grupo"
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
         DataField       =   "Cod_Montador"
         Caption         =   "Cod.Sub-Grupo"
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
         DataField       =   "DESCRIPCION"
         Caption         =   "Descripción Subgrupo"
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
      BeginProperty Column03 
         DataField       =   "par_codigo"
         Caption         =   "Partida"
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
      BeginProperty Column04 
         DataField       =   "estado"
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
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   6554.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   615.118
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
      Left            =   0
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   10515
      Begin VB.CommandButton CmdGraCabeza 
         Caption         =   "Grabar"
         Height          =   675
         Left            =   4920
         Picture         =   "ALFrm_montado.frx":10A11
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton CmdCanCabeza 
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   6360
         Picture         =   "ALFrm_montado.frx":10C1B
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc AdoGRUPO 
      Height          =   375
      Left            =   120
      Top             =   8280
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
   Begin MSAdodcLib.Adodc AdoPartida 
      Height          =   375
      Left            =   2400
      Top             =   8280
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
      Caption         =   "AdoPartida"
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
      AutoSize        =   -1  'True
      Caption         =   "UNIDAD"
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
      Height          =   240
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "ALFrm_montado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTabla As New ADODB.Recordset
Dim rs_GRUPO As New ADODB.Recordset
Dim rsAuxTabla As New ADODB.Recordset
Dim rs_partida As New ADODB.Recordset
Dim rscorrelativo As New ADODB.Recordset

Dim num_comprobante As Integer ' variable donde se almacena correlativo de comprobante

Dim queryinicial As String
Dim SQL_FOR, sino As String
Dim swgraba As Integer
Dim Marca As BookmarkEnum
Dim PosibleApliqueFiltro As Boolean
Dim ClBuscaGrid As ClBuscaEnGridExterno

Private Sub CmdAddCabeza_Click()
    'adicion
    Dim cod_MONTADOR As String
    DtgMain.Visible = False
    FrmDatos.Visible = True
    frmabm.Visible = False
    frmgrabcabeza.Visible = True
    DtcGrupoD.Enabled = True
    DtcGrupoD.Text = ""
    DtcGrupo.Text = ""
    Textdescri.Text = ""
    DtcParD.Text = "SIN PARTIDA"
    DtcPar.Text = "99999"
    Option1 = True
    'saca  correlativo
    'DE.dbo_AL_MAXCOD_Montador cod_MONTADOR
    'TextCOD_MONTADOR = cod_MONTADOR + 1
    swgraba = 1
    DtcGrupoD.SetFocus
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
    AdodcTabla.Recordset.CancelUpdate
    Call OptActivos_Click
    If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
        AdodcTabla.Recordset.Move Marca - 1
    End If
    swgraba = 0
End Sub

Private Sub CmdDelCabeza_Click()
  On Error GoTo UpdateErr
   If AdodcTabla.Recordset.RecordCount > 0 Then
      If ExisteGrupo(AdodcTabla.Recordset!cod_MONTADOR) Then MsgBox "No se puede eliminar al SubGrupo que ya tiene registrado un BIEN o SERVICIO.", vbInformation + vbOKOnly, "Atención": Exit Sub
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         'AdodcTabla.Recordset.Delete
         
         AdodcTabla.Recordset!estado = "E"
         AdodcTabla.Recordset.Update
         AdodcTabla.Recordset.Requery
      End If
   Else
        MsgBox "No existen registros.", vbExclamation, "Atención"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Public Sub genera_codigo()
  Set rscorrelativo = New ADODB.Recordset
  rscorrelativo.CursorLocation = adUseClient
  If rscorrelativo.State = 1 Then rscorrelativo.Close
  rscorrelativo.Open "SELECT CodGrupo, Correl_Sub FROM AlClGrupo WHERE (CodGrupo = '" & DtcGrupo.Text & "')", db, adOpenKeyset, adLockOptimistic
  If rscorrelativo.RecordCount > 0 Then
    rscorrelativo.MoveFirst
    num_comprobante = rscorrelativo!Correl_Sub + 1
    rscorrelativo!Correl_Sub = rscorrelativo!Correl_Sub + 1
    rscorrelativo.Update
  Else
    num_comprobante = 1
    rscorrelativo!Correl_Sub = 1
    rscorrelativo.Update
  End If
End Sub

Private Sub CmdGraCabeza_Click()
Dim estatus2 As String
If DtcGrupo.Text = "" Then
    MsgBox "Error, Debe elegir el GRUPO que se requiere, vuelva a intentar..."
    Exit Sub
End If
DtgMain.Visible = True
frmabm.Visible = True
FrmDatos.Visible = False
frmgrabcabeza.Visible = False
' grabar
If swgraba = 1 Then
'    DE.dbo_AL_MAXCOD_Montador cod_MONTADOR
'    TextCOD_MONTADOR = cod_MONTADOR + 1
    Call genera_codigo
    If num_comprobante < 10 Then
        TextCOD_MONTADOR = DtcGrupo.Text + "000" + Trim(Str(num_comprobante))
    End If
    If num_comprobante > 9 And num_comprobante < 100 Then
        TextCOD_MONTADOR = DtcGrupo.Text + "00" + Trim(Str(num_comprobante))
    End If
    If num_comprobante > 99 And num_comprobante < 1000 Then
        TextCOD_MONTADOR = DtcGrupo.Text + "0" + Trim(Str(num_comprobante))
    End If
    If num_comprobante > 999 Then
        TextCOD_MONTADOR = DtcGrupo.Text + Trim(Str(num_comprobante))
    End If
  
'    Set rsAuxTabla = New ADODB.Recordset
'    rsAuxTabla.Open "select * from AL_Montador where CodGRUPO = '" & DtcGrupo.Text & "'  ", db, adOpenKeyset, adLockOptimistic
'    'If rsAuxTabla.RecordCount > 0 Then
'        TextCOD_MONTADOR = DtcGrupo.Text + "0000" + rsAuxTabla.RecordCount + 1
'    'End If
    Set rstbeneaux = New ADODB.Recordset
    SQL_FOR = "select * from AL_Montador where Cod_montador = '" & TextCOD_MONTADOR.Text & "'  "
    rstbeneaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
    If rstbeneaux.RecordCount > 0 Then
        'TextCOD_MONTADOR = TextCOD_MONTADOR + "A"
'        SW = True
        MsgBox " El CODIGO ya existe, verifique el registro y vuelva a intentar ..."
        Exit Sub
    End If
    db.Execute "INSERT INTO AL_Montador (CodGrupo, Cod_montador, DESCRIPCION, CD_OCA, par_codigo, ESTADO, usr_usuario, fecha_registro, hora_registro) VALUES ('" & DtcGrupo & "', '" & TextCOD_MONTADOR & "', '" & Textdescri & "', '" & Trim(TxtDescripAnt) & "',  '" & DtcPar & "', 'N', '" & GlUsuario & "',  '01/08/2011', '12:00')  "
    
    'DE.dbo_al_inserta_montador Textdescri
    Option1 = True
    Call OptActivos_Click
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
    'PROC ALM Modifica Marcas
    'DB.Execute "UPDATE AL_Montador SET (CodGrupo, Cod_montador, DESCRIPCION, CD_OCA, ESTADO, usr_usuario, fecha_registro, hora_registro) VALUES ('" & TxtGrupo & "', '" & TextCOD_MONTADOR & "', '" & Textdescri & "', '" & Trim(TxtDescripAnt) & "',  'N', '" & GlUsuario & "',  '01/08/2011', '12:00')  "
    db.Execute "UPDATE AL_Montador SET DESCRIPCION='" & Textdescri & "', CD_OCA='" & Trim(TxtDescripAnt) & "', par_codigo='" & DtcPar & "', ESTADO='" & estatus2 & "' WHERE Cod_montador='" & TextCOD_MONTADOR & "'"
    'DE.dbo_al_Modi_Montador AdodcTabla.Recordset!cod_MONTADOR, Textdescri, estatus
    Call OptActivos_Click
'    AdodcTabla.Refresh
    If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
        AdodcTabla.Recordset.Move Marca - 1
    End If
End If
DtcGrupoD.Enabled = False
End Sub

Private Sub CmdImpCabeza_Click()
  Dim IResult As Integer
  With CryLista
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True

        .ReportFileName = App.Path & "\Reportes\Almacen\Productos_SubGrupo.rpt"
    IResult = .PrintReport
    If IResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With
End Sub

Private Sub CmdModCabeza_Click()
'MODIFICAR
 If Not (AdodcTabla.Recordset.EOF) Or Not (AdodcTabla.Recordset.BOF) Then
    Dim cod_MONTADOR As String
    DtgMain.Visible = False
    FrmDatos.Visible = True
    frmabm.Visible = False
    frmgrabcabeza.Visible = True
    DtcGrupoD.Enabled = False
    'muestra datos
'    DtcGrupo = AdodcTabla.Recordset!CodGrupo
    TextCOD_MONTADOR = AdodcTabla.Recordset!cod_MONTADOR
    Textdescri = AdodcTabla.Recordset!descripcion
    If AdodcTabla.Recordset!estado = "S" Then
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

Private Sub DtcGrupo_Click(Area As Integer)
    DtcGrupoD.BoundText = DtcGrupo.BoundText
End Sub

Private Sub DtcGrupoD_Click(Area As Integer)
    DtcGrupo.BoundText = DtcGrupoD.BoundText
End Sub

Private Sub DtcGrupoD_LostFocus()
'    TxtGrupo.Text = DtcGrupo.Text
End Sub

Private Sub DtcPar_Click(Area As Integer)
    DtcParD.BoundText = DtcPar.BoundText
End Sub

Private Sub DtcParD_Click(Area As Integer)
    DtcPar.BoundText = DtcParD.BoundText
End Sub

Private Sub Form_Load()

Option3 = True
swgraba = 0
DtgMain.Visible = True
FrmDatos.Visible = False
frmabm.Visible = True
frmgrabcabeza.Visible = False
    Set rsTabla = New ADODB.Recordset
    If rsTabla.State = 1 Then rsTabla.Close
    queryinicial = "select * from Al_Montador where estado<>'E' "
    'queryinicial = "select * from Al_Montador "
    rsTabla.Open queryinicial & " order by codGrupo, cod_MONTADOR ", db, adOpenKeyset, adLockOptimistic
    Set AdodcTabla.Recordset = rsTabla
    
    Set rs_GRUPO = New ADODB.Recordset
    rs_GRUPO.Open "SELECT * FROM ALCLGrupo WHERE activo='S' ", db, adOpenStatic
    Set AdoGRUPO.Recordset = rs_GRUPO
        
    Set rs_partida = New ADODB.Recordset
    rs_partida.Open "SELECT * FROM fc_partida_gasto WHERE par_activo='1' OR par_activo='S' order by par_descripcion_larga", db, adOpenStatic
    Set AdoPartida.Recordset = rs_partida

End Sub

Private Sub OptActivos_Click()
Set rsTabla = New ADODB.Recordset
    If rsTabla.State = 1 Then rsTabla.Close
    queryinicial = "SELECT * From Al_montador where estado ='N'"
    'rsTabla.Open queryinicial & " order by CAST(cod_MONTADOR AS INT)", db, adOpenKeyset, adLockOptimistic  'JQA JUL-2008
    rsTabla.Open queryinicial & " order by codGrupo, cod_MONTADOR ", db, adOpenKeyset, adLockOptimistic
    Set AdodcTabla.Recordset = rsTabla
End Sub

Private Sub Option3_Click()
Set rsTabla = New ADODB.Recordset
    If rsTabla.State = 1 Then rsTabla.Close
    queryinicial = "SELECT * From Al_Montador"
    rsTabla.Open queryinicial & " order by CodGrupo, cod_MONTADOR ", db, adOpenKeyset, adLockOptimistic
    Set AdodcTabla.Recordset = rsTabla
End Sub

Private Function ExisteGrupo(cod_MONTADOR As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ALCLDetalle WHERE cod_MONTADOR = '" & cod_MONTADOR & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteGrupo = rs!Cuantos > 0
End Function


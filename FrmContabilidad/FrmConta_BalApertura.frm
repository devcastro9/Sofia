VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmConta_BalApertura 
   BackColor       =   &H00000000&
   Caption         =   "Contabilidad - Balance de Apertura"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   Icon            =   "FrmConta_BalApertura.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fra_BusquedaC 
      BackColor       =   &H80000017&
      Caption         =   "Búsqueda (Cuenta)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   5400
      TabIndex        =   55
      Top             =   4560
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Height          =   1500
         Left            =   45
         Negotiate       =   -1  'True
         Picture         =   "FrmConta_BalApertura.frx":0A02
         ScaleHeight     =   1440
         ScaleWidth      =   15360
         TabIndex        =   56
         Top             =   240
         Width           =   15420
         Begin VB.ComboBox CboCampoC 
            Height          =   315
            ItemData        =   "FrmConta_BalApertura.frx":54F2
            Left            =   240
            List            =   "FrmConta_BalApertura.frx":54FC
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   360
            Width           =   1590
         End
         Begin VB.ComboBox CboOperadorC 
            Height          =   315
            ItemData        =   "FrmConta_BalApertura.frx":5516
            Left            =   1890
            List            =   "FrmConta_BalApertura.frx":5520
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox TxtValorC 
            Height          =   336
            Left            =   2880
            MultiLine       =   -1  'True
            TabIndex        =   57
            Top             =   360
            Width           =   2700
         End
         Begin VB.PictureBox TbrAvanzadas 
            Height          =   420
            Left            =   720
            Negotiate       =   -1  'True
            ScaleHeight     =   360
            ScaleWidth      =   4110
            TabIndex        =   64
            Top             =   960
            Width           =   4170
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   3000
            TabIndex        =   62
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operador:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   1920
            TabIndex        =   61
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Columna:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   345
            TabIndex        =   60
            Top             =   120
            Width           =   1305
         End
      End
   End
   Begin VB.Frame Fra_BuscaGral 
      BackColor       =   &H80000017&
      Caption         =   "Búsqueda por:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   1200
      TabIndex        =   45
      Top             =   1200
      Visible         =   0   'False
      Width           =   6495
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Height          =   1500
         Left            =   40
         Negotiate       =   -1  'True
         Picture         =   "FrmConta_BalApertura.frx":552D
         ScaleHeight     =   1440
         ScaleWidth      =   15360
         TabIndex        =   46
         Top             =   240
         Width           =   15420
         Begin VB.CommandButton CmBusq 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Filtrar"
            Height          =   360
            Left            =   120
            Picture         =   "FrmConta_BalApertura.frx":A01D
            TabIndex        =   54
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.ComboBox CboStatus 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cboUnidad 
            Height          =   315
            ItemData        =   "FrmConta_BalApertura.frx":A167
            Left            =   240
            List            =   "FrmConta_BalApertura.frx":A169
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtCtaNom 
            Height          =   336
            Left            =   3000
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   360
            Width           =   3060
         End
         Begin VB.CommandButton CmdSaleB 
            Caption         =   "Salir"
            Height          =   360
            Left            =   4800
            TabIndex        =   47
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox Toolbar1 
            Height          =   420
            Left            =   1080
            Negotiate       =   -1  'True
            ScaleHeight     =   360
            ScaleWidth      =   4110
            TabIndex        =   81
            Top             =   840
            Width           =   4170
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre de Cuenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   3675
            TabIndex        =   51
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   2130
            TabIndex        =   50
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Código Cuenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   49
            Top             =   120
            Width           =   1665
         End
      End
   End
   Begin MSAdodcLib.Adodc adosolicitud1 
      Height          =   330
      Left            =   1080
      Top             =   3360
      Width           =   14085
      _ExtentX        =   24844
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
      BackColor       =   12640511
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
      Caption         =   " <-- Inicio                                          Balance de Apertura                                               Fin -->"
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
   Begin VB.Frame Fra_Busqueda 
      BackColor       =   &H80000017&
      Caption         =   "Búsqueda (Auxiliares)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   9000
      TabIndex        =   23
      Top             =   6840
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox PicCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Height          =   1500
         Left            =   40
         Negotiate       =   -1  'True
         Picture         =   "FrmConta_BalApertura.frx":A16B
         ScaleHeight     =   1440
         ScaleWidth      =   15360
         TabIndex        =   38
         Top             =   240
         Width           =   15420
         Begin VB.TextBox TxtValor 
            Height          =   336
            Left            =   2880
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   360
            Width           =   2700
         End
         Begin VB.ComboBox CboOperador 
            Height          =   315
            ItemData        =   "FrmConta_BalApertura.frx":EC5B
            Left            =   1890
            List            =   "FrmConta_BalApertura.frx":EC65
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   360
            Width           =   915
         End
         Begin VB.ComboBox CboCampo 
            Height          =   315
            ItemData        =   "FrmConta_BalApertura.frx":EC72
            Left            =   240
            List            =   "FrmConta_BalApertura.frx":EC7C
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   360
            Width           =   1590
         End
         Begin VB.PictureBox ToolbarAux 
            Height          =   420
            Left            =   840
            Negotiate       =   -1  'True
            ScaleHeight     =   360
            ScaleWidth      =   4110
            TabIndex        =   65
            Top             =   960
            Width           =   4170
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Columna:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   345
            TabIndex        =   41
            Top             =   120
            Width           =   1425
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operador:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   2040
            TabIndex        =   40
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Left            =   3120
            TabIndex        =   39
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin TrueOleDBGrid60.TDBDropDown TDBPlan 
      Bindings        =   "FrmConta_BalApertura.frx":EC96
      Height          =   2655
      Left            =   4800
      OleObjectBlob   =   "FrmConta_BalApertura.frx":ECAC
      TabIndex        =   30
      Top             =   720
      Width           =   10215
   End
   Begin VB.Frame frmabm 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8370
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   1035
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H8000000A&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":16D54
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   7200
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H8000000A&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":16F5E
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Imprime Balance de Apertura"
         Top             =   5640
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H8000000A&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":1751B
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Busca un Registro"
         Top             =   4920
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H8000000D&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":17AD3
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3840
         Width           =   765
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H8000000D&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":17CDD
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Aprueba Registro"
         Top             =   3120
         Width           =   770
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H8000000A&
         Caption         =   "Anular"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":17EE7
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Anula Registro Activo"
         Top             =   2040
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H8000000A&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":18BB1
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   1320
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H8000000A&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":19191
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Nuevo Registro"
         Top             =   600
         Width           =   765
      End
      Begin Crystal.CrystalReport CryF01 
         Left            =   240
         Top             =   6240
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
   End
   Begin MSAdodcLib.Adodc Adoconvenio 
      Height          =   330
      Left            =   2160
      Top             =   8520
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
      Caption         =   "adoconvenio"
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
   Begin Crystal.CrystalReport CRyAux12 
      Left            =   1440
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryOrg 
      Left            =   240
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryLMayorCtaBancaria 
      Left            =   0
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryLMayorBenef 
      Left            =   480
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   1080
      TabIndex        =   11
      Top             =   3600
      Width           =   14085
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         Caption         =   "Basurero"
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   5640
         TabIndex        =   127
         Top             =   120
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtbusca1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   137
            Top             =   1320
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.ComboBox cboCtaBancaria 
            Height          =   315
            Left            =   120
            TabIndex        =   136
            Text            =   "Combo1"
            Top             =   960
            Visible         =   0   'False
            Width           =   795
         End
         Begin MSDataListLib.DataCombo DtCIdConvenio 
            Bindings        =   "FrmConta_BalApertura.frx":197B5
            Height          =   315
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "codigo_convenio"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtCDesConvenio 
            Bindings        =   "FrmConta_BalApertura.frx":197CF
            DataField       =   "aux1"
            Height          =   315
            Left            =   960
            TabIndex        =   129
            Top             =   240
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripcion"
            BoundColumn     =   "aux"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcGrBien 
            Bindings        =   "FrmConta_BalApertura.frx":197E9
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   2520
            TabIndex        =   130
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "cod_montador"
            BoundColumn     =   "cod_montador"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcGrBienDes 
            Bindings        =   "FrmConta_BalApertura.frx":19801
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   3720
            TabIndex        =   131
            Top             =   240
            Visible         =   0   'False
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "descripcion"
            BoundColumn     =   "cod_montador"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcProy 
            Bindings        =   "FrmConta_BalApertura.frx":19819
            DataField       =   "denominacion_aux3"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   132
            Top             =   600
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Pro_proyecto"
            BoundColumn     =   "Pro_proyecto"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcProyDes 
            Bindings        =   "FrmConta_BalApertura.frx":1982F
            DataField       =   "denominacion_aux3"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   960
            TabIndex        =   133
            Top             =   600
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Pro_descripcion_larga"
            BoundColumn     =   "Pro_proyecto"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtCOrg 
            Height          =   315
            Left            =   2520
            TabIndex        =   134
            Top             =   600
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DTCNomOrg 
            Height          =   315
            Left            =   3480
            TabIndex        =   135
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcDenomAux2 
            Height          =   315
            Left            =   960
            TabIndex        =   138
            Top             =   960
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "denominacion_convenio"
            BoundColumn     =   "codigo_convenio"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcDenomAux3 
            Height          =   315
            Left            =   3480
            TabIndex        =   139
            Top             =   960
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "denominacion_convenio"
            BoundColumn     =   "codigo_convenio"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Fra_Org 
         BackColor       =   &H00404040&
         Caption         =   "09-FINANCIADOR"
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   8520
         TabIndex        =   122
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton Command6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo"
            Height          =   700
            Left            =   4080
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmConta_BalApertura.frx":19845
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   840
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.CommandButton BtnAprobar09 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aceptar"
            Height          =   700
            Left            =   2400
            Picture         =   "FrmConta_BalApertura.frx":19F60
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Aprueba Registro"
            Top             =   840
            Width           =   770
         End
         Begin MSDataListLib.DataCombo Dtc_Org 
            Bindings        =   "FrmConta_BalApertura.frx":1A16A
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   125
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "org_codigo"
            BoundColumn     =   "org_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo Dtc_OrgD 
            Bindings        =   "FrmConta_BalApertura.frx":1A187
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   126
            Top             =   360
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "org_descripcion"
            BoundColumn     =   "org_codigo"
            Text            =   "Todos"
         End
      End
      Begin VB.Frame Fra_Depto 
         BackColor       =   &H00404040&
         Caption         =   "06-DEPARTAMENTOS DEL PAIS"
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   8520
         TabIndex        =   117
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton BtnAprobar06 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aceptar"
            Height          =   700
            Left            =   2400
            Picture         =   "FrmConta_BalApertura.frx":1A1A4
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Aprueba Registro"
            Top             =   840
            Width           =   770
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo"
            Height          =   700
            Left            =   4080
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmConta_BalApertura.frx":1A3AE
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   840
            Visible         =   0   'False
            Width           =   780
         End
         Begin MSDataListLib.DataCombo Dtc_Depto 
            Bindings        =   "FrmConta_BalApertura.frx":1AAC9
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   120
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo Dtc_DeptoD 
            Bindings        =   "FrmConta_BalApertura.frx":1AAE8
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   121
            Top             =   360
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
      End
      Begin VB.Frame Fra_UEjec 
         BackColor       =   &H00404040&
         Caption         =   "04-UNIDAD_EJECUTORA"
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   8520
         TabIndex        =   112
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo"
            Height          =   700
            Left            =   4080
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmConta_BalApertura.frx":1AB07
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   840
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.CommandButton BtnAprobar04 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aceptar"
            Height          =   700
            Left            =   2400
            Picture         =   "FrmConta_BalApertura.frx":1B222
            Style           =   1  'Graphical
            TabIndex        =   113
            ToolTipText     =   "Aprueba Registro"
            Top             =   840
            Width           =   770
         End
         Begin MSDataListLib.DataCombo Dtc_Uejec 
            Bindings        =   "FrmConta_BalApertura.frx":1B42C
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   115
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "unidad_codigo"
            BoundColumn     =   "unidad_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo Dtc_UejecD 
            Bindings        =   "FrmConta_BalApertura.frx":1B449
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   116
            Top             =   360
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   "Todos"
         End
      End
      Begin VB.Frame Fra_Proy 
         BackColor       =   &H00404040&
         Caption         =   "03-PROYECTOS"
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   8520
         TabIndex        =   107
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton BtnAprobar03 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aceptar"
            Height          =   700
            Left            =   2400
            Picture         =   "FrmConta_BalApertura.frx":1B466
            Style           =   1  'Graphical
            TabIndex        =   109
            ToolTipText     =   "Aprueba Registro"
            Top             =   840
            Width           =   770
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo"
            Height          =   700
            Left            =   4080
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmConta_BalApertura.frx":1B670
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   840
            Visible         =   0   'False
            Width           =   780
         End
         Begin MSDataListLib.DataCombo Dtc_Proy 
            Bindings        =   "FrmConta_BalApertura.frx":1BD8B
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   110
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "edif_codigo"
            BoundColumn     =   "edif_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo Dtc_ProyD 
            Bindings        =   "FrmConta_BalApertura.frx":1BDA6
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   111
            Top             =   360
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "edif_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
      End
      Begin VB.Frame Fra_CtaBco 
         BackColor       =   &H00404040&
         Caption         =   "02-CUENTAS BANCARIAS"
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   8520
         TabIndex        =   102
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo"
            Height          =   700
            Left            =   4080
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmConta_BalApertura.frx":1BDC1
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   840
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.CommandButton BtnAprobar02 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aceptar"
            Height          =   700
            Left            =   2400
            Picture         =   "FrmConta_BalApertura.frx":1C4DC
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Aprueba Registro"
            Top             =   840
            Width           =   770
         End
         Begin MSDataListLib.DataCombo Dtc_CtaBco 
            Bindings        =   "FrmConta_BalApertura.frx":1C6E6
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   105
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo Dtc_CtaBcoD 
            Bindings        =   "FrmConta_BalApertura.frx":1C701
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   106
            Top             =   360
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
      End
      Begin VB.Frame Fra_Benef 
         BackColor       =   &H00404040&
         Caption         =   "01-BENEFICIARIOS"
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   8520
         TabIndex        =   95
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton BtnAprobar01 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aceptar"
            Height          =   700
            Left            =   2400
            Picture         =   "FrmConta_BalApertura.frx":1C71C
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Aprueba Registro"
            Top             =   840
            Width           =   770
         End
         Begin VB.CommandButton BtnAux2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nuevo"
            Height          =   700
            Left            =   4080
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmConta_BalApertura.frx":1C926
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   840
            Visible         =   0   'False
            Width           =   780
         End
         Begin MSDataListLib.DataCombo Dtc_benef 
            Bindings        =   "FrmConta_BalApertura.frx":1D041
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   100
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo Dtc_benefD 
            Bindings        =   "FrmConta_BalApertura.frx":1D059
            DataField       =   "denominacion_aux1"
            DataSource      =   "adosolicitud1"
            Height          =   315
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
      End
      Begin VB.TextBox TxtCodAnt 
         Height          =   330
         Left            =   11040
         TabIndex        =   78
         Top             =   2760
         Width           =   1800
      End
      Begin VB.TextBox TxtHaber 
         Height          =   330
         Left            =   6240
         TabIndex        =   76
         Top             =   2760
         Width           =   1800
      End
      Begin VB.TextBox TxtDebe 
         Height          =   330
         Left            =   1980
         TabIndex        =   75
         Top             =   2760
         Width           =   1800
      End
      Begin VB.ComboBox cbosubcta1 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   540
         Width           =   1140
      End
      Begin VB.ComboBox cbosubcta2 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   900
         Width           =   1140
      End
      Begin VB.ComboBox cbocta 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   180
         Width           =   1140
      End
      Begin VB.CheckBox Chkaux1 
         BackColor       =   &H00000000&
         Caption         =   "Auxiliar 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1420
         Width           =   975
      End
      Begin VB.CheckBox Chkaux2 
         BackColor       =   &H00000000&
         Caption         =   "Auxiliar 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1900
         Width           =   1005
      End
      Begin VB.CheckBox Chkaux3 
         BackColor       =   &H00000000&
         Caption         =   "Auxiliar 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   2380
         Width           =   1080
      End
      Begin VB.TextBox txtax1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1380
         Width           =   465
      End
      Begin VB.TextBox Txtax2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1860
         Width           =   465
      End
      Begin VB.TextBox txtax3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2340
         Width           =   465
      End
      Begin MSComCtl2.DTPicker DTPfin 
         Height          =   360
         Left            =   3480
         TabIndex        =   8
         Top             =   3225
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         _Version        =   393216
         Format          =   109117441
         CurrentDate     =   36614
      End
      Begin MSComCtl2.DTPicker DTPinicio 
         Height          =   345
         Left            =   1320
         TabIndex        =   7
         Top             =   3240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         _Version        =   393216
         Format          =   109117441
         CurrentDate     =   37257
         MaxDate         =   2958101
         MinDate         =   36892
      End
      Begin MSDataListLib.DataCombo cbocta1 
         Bindings        =   "FrmConta_BalApertura.frx":1D071
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   10440
         TabIndex        =   66
         Top             =   900
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "cod_cta"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtccta 
         Bindings        =   "FrmConta_BalApertura.frx":1D088
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   11040
         TabIndex        =   67
         Top             =   660
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "cuenta"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcsub1 
         Bindings        =   "FrmConta_BalApertura.frx":1D09F
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   11880
         TabIndex        =   68
         Top             =   660
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "SubCta1"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcsub2 
         Bindings        =   "FrmConta_BalApertura.frx":1D0B6
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   12720
         TabIndex        =   69
         Top             =   660
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "SubCta2"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcCtaNom 
         Bindings        =   "FrmConta_BalApertura.frx":1D0CD
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   2940
         TabIndex        =   70
         Top             =   900
         Visible         =   0   'False
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NombreCta"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcAux1 
         Bindings        =   "FrmConta_BalApertura.frx":1D0E4
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   11040
         TabIndex        =   71
         Top             =   360
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Aux1"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcAux2 
         Bindings        =   "FrmConta_BalApertura.frx":1D0FB
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   11880
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Aux2"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcAux3 
         Bindings        =   "FrmConta_BalApertura.frx":1D112
         DataField       =   "cod_cta"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   12720
         TabIndex        =   73
         Top             =   360
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Aux3"
         BoundColumn     =   "cod_cta"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcTAux1Des 
         Bindings        =   "FrmConta_BalApertura.frx":1D129
         DataField       =   "aux1"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   2040
         TabIndex        =   98
         Top             =   1380
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "descripcion"
         BoundColumn     =   "aux"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcTAux1 
         Bindings        =   "FrmConta_BalApertura.frx":1D148
         DataField       =   "aux1"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   1320
         TabIndex        =   99
         Top             =   1380
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "aux"
         BoundColumn     =   "aux"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo DtcTAux2 
         Bindings        =   "FrmConta_BalApertura.frx":1D167
         DataField       =   "aux2"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   1320
         TabIndex        =   140
         Top             =   1860
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "aux"
         BoundColumn     =   "aux"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo DtcTAux2Des 
         Bindings        =   "FrmConta_BalApertura.frx":1D186
         DataField       =   "aux2"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   2040
         TabIndex        =   141
         Top             =   1860
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "descripcion"
         BoundColumn     =   "aux"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcTAux3 
         Bindings        =   "FrmConta_BalApertura.frx":1D1A5
         DataField       =   "aux3"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   1320
         TabIndex        =   142
         Top             =   2340
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "aux"
         BoundColumn     =   "aux"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo DtcTAux3Des 
         Bindings        =   "FrmConta_BalApertura.frx":1D1C4
         DataField       =   "aux3"
         DataSource      =   "adosolicitud1"
         Height          =   315
         Left            =   2040
         TabIndex        =   143
         Top             =   2340
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "descripcion"
         BoundColumn     =   "aux"
         Text            =   "Todos"
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "de"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   345
         Left            =   13080
         TabIndex        =   80
         Top             =   120
         Width           =   225
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código Anterior CGI:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9480
         TabIndex        =   79
         Top             =   2805
         Width           =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Saldo Haber Bs."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5040
         TabIndex        =   77
         Top             =   2805
         Width           =   1155
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Saldo Debe Bs."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   74
         Top             =   2805
         Width           =   1215
      End
      Begin VB.Label Lbl_Aux3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   36
         Top             =   2340
         Width           =   1860
      End
      Begin VB.Label Lbl_Aux2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   35
         Top             =   1860
         Width           =   1860
      End
      Begin VB.Label Lbl_Aux1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   34
         Top             =   1380
         Width           =   1860
      End
      Begin VB.Label LblNom_Aux3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6240
         TabIndex        =   33
         Top             =   2340
         Width           =   4590
      End
      Begin VB.Label LblNom_Aux2 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6240
         TabIndex        =   32
         Top             =   1860
         Width           =   4590
      End
      Begin VB.Label LblNom_Aux1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6240
         TabIndex        =   31
         Top             =   1380
         Width           =   4590
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   345
         Index           =   3
         Left            =   13440
         TabIndex        =   28
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Nro. Registros:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   345
         Index           =   1
         Left            =   10920
         TabIndex        =   29
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label marca11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cuenta:"
         DataField       =   "correl"
         DataSource      =   "adosolicitud1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   12480
         TabIndex        =   25
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Lblsub1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2760
         TabIndex        =   21
         Top             =   540
         Width           =   60
      End
      Begin VB.Label lblcuenta 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2760
         TabIndex        =   20
         Top             =   180
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Subcuenta 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   940
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Subcuenta 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   580
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   220
         Width           =   555
      End
      Begin VB.Label lbsub2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Top             =   900
         Width           =   60
      End
      Begin VB.Label Label4 
         Caption         =   "Desde:"
         Height          =   240
         Left            =   300
         TabIndex        =   15
         Top             =   3315
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta:"
         Height          =   240
         Left            =   2880
         TabIndex        =   14
         Top             =   3330
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   1800
      Left            =   1080
      TabIndex        =   0
      Top             =   6720
      Width           =   14100
      Begin VB.CommandButton BtnBuscarB 
         BackColor       =   &H8000000D&
         Caption         =   "Busca Cta -->"
         Height          =   840
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":1D1E3
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Busca un Registro"
         Top             =   480
         Width           =   765
      End
      Begin MSDataGridLib.DataGrid DtgPlanCtas 
         Height          =   1650
         Left            =   960
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   2910
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "PLAN DE CUENTAS"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Cuenta"
            Caption         =   "Cuenta"
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
            DataField       =   "SubCta1"
            Caption         =   "SubCta1"
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
            DataField       =   "SubCta2"
            Caption         =   "SubCta2"
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
            DataField       =   "Aux1"
            Caption         =   "Aux1"
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
            DataField       =   "Aux2"
            Caption         =   "Aux2"
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
            DataField       =   "Aux3"
            Caption         =   "Aux3"
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
         BeginProperty Column06 
            DataField       =   "NombreCta"
            Caption         =   "NombreCta"
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   6194.835
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DtGbenef 
         Bindings        =   "FrmConta_BalApertura.frx":1D79B
         Height          =   1650
         Left            =   960
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   2910
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
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
         Caption         =   "BENEFICIARIOS"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "codigo_beneficiario"
            Caption         =   "Código Beneficiario"
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
            DataField       =   "denominacion_beneficiario"
            Caption         =   "Denominación"
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
               ColumnWidth     =   6089.953
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DTGBanco 
         Height          =   1575
         Left            =   960
         TabIndex        =   24
         Top             =   120
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   2778
         _Version        =   393216
         BackColor       =   16777152
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
         Caption         =   "CUENTAS BANCARIAS"
         ColumnCount     =   3
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
         BeginProperty Column02 
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
               ColumnWidth     =   4350.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5265.071
            EndProperty
         EndProperty
      End
      Begin Crystal.CrystalReport CryBenefConvenios 
         Left            =   7200
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin Crystal.CrystalReport CryLMayor 
      Left            =   960
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryConv_Conv 
      Left            =   1920
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TrueOleDBGrid60.TDBGrid DataGrid3 
      Bindings        =   "FrmConta_BalApertura.frx":1D7B3
      Height          =   3375
      Left            =   1080
      OleObjectBlob   =   "FrmConta_BalApertura.frx":1D7CF
      TabIndex        =   27
      Top             =   0
      Width           =   14055
   End
   Begin MSAdodcLib.Adodc AdodcOrganismo 
      Height          =   330
      Left            =   4320
      Top             =   8520
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "AdodcOrganismo"
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
   Begin VB.Frame FrmGraba 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8355
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   1035
      Begin VB.CommandButton BtnBuscarA 
         BackColor       =   &H8000000D&
         Caption         =   "Busca Nom.Cta."
         Height          =   840
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":27571
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Busca un Registro"
         Top             =   4560
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimirA 
         BackColor       =   &H8000000A&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":27B29
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Imprime Balance"
         Top             =   1320
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H8000000A&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":280E6
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   5760
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H8000000A&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":282F0
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   3000
         Width           =   765
      End
      Begin VB.CommandButton BtnEnviar 
         BackColor       =   &H8000000A&
         Caption         =   "Grabar"
         Height          =   720
         Left            =   120
         Picture         =   "FrmConta_BalApertura.frx":284FA
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2040
         Width           =   770
      End
   End
   Begin MSAdodcLib.Adodc AdoPlan 
      Height          =   330
      Left            =   0
      Top             =   8520
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
      Caption         =   "AdoPlan"
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
   Begin VB.PictureBox ImlImagenesAv 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   5640
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   144
      Top             =   8400
      Width           =   1200
   End
   Begin VB.PictureBox ImlImagenesA 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   6240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   145
      Top             =   8400
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc AdoPlan1 
      Height          =   330
      Left            =   6480
      Top             =   8520
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "AdoPlan1"
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
   Begin MSAdodcLib.Adodc AdoPlan2 
      Height          =   330
      Left            =   8760
      Top             =   8520
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "AdoPlan2"
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
   Begin MSAdodcLib.Adodc AdoPlan3 
      Height          =   330
      Left            =   11040
      Top             =   8520
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "AdoPlan3"
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
   Begin MSAdodcLib.Adodc AdoProy 
      Height          =   330
      Left            =   0
      Top             =   8880
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "AdoProy"
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
   Begin MSAdodcLib.Adodc AdoGrBien 
      Height          =   330
      Left            =   2160
      Top             =   8880
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "AdoGrBien"
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
   Begin MSAdodcLib.Adodc Ado_TipoAuxiliar 
      Height          =   330
      Left            =   4320
      Top             =   8880
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Ado_TipoAuxiliar"
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
   Begin MSAdodcLib.Adodc Ado_Benef 
      Height          =   330
      Left            =   6480
      Top             =   8880
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Ado_Benef"
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
   Begin MSAdodcLib.Adodc Ado_CtaBanco 
      Height          =   330
      Left            =   8760
      Top             =   8880
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Ado_CtaBanco"
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
   Begin MSAdodcLib.Adodc Ado_Proyecto 
      Height          =   330
      Left            =   13320
      Top             =   8520
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Ado_Proyecto"
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
   Begin MSAdodcLib.Adodc Ado_UEjecutora 
      Height          =   330
      Left            =   11040
      Top             =   8880
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Ado_UEjecutora"
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
   Begin MSAdodcLib.Adodc Ado_Departamento 
      Height          =   330
      Left            =   13320
      Top             =   8880
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Ado_Departamento"
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
Attribute VB_Name = "FrmConta_BalApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/****** REFORMULADO EL 20 de junio
'/************  RECORDSETS
Dim sql1 As String
Dim sql2 As String
Dim lcta As String
Dim nombenef As String
Dim combenef As ADODB.Command
Dim comAux12 As ADODB.Command
Dim comORG As ADODB.Command
'---
Dim rsOrganismo As ADODB.Recordset
Dim comctabancaria As ADODB.Command
Dim rsplanctas, rsPlanBusq As New ADODB.Recordset
Dim rsPlanCta1, rsPlanCta2, rsPlanCta3 As New ADODB.Recordset
Dim rscuentas As ADODB.Recordset
Dim rsnombresub1 As ADODB.Recordset
Dim rssubcuenta As ADODB.Recordset
Dim rscta_bancaria As ADODB.Recordset
Dim rsbeneficiario As ADODB.Recordset
Dim rssaldos As ADODB.Recordset
Dim rsctabancaria As ADODB.Recordset
Dim rsConvenio As ADODB.Recordset
Dim rstAo_solicitud1 As ADODB.Recordset
Dim rs_bien As ADODB.Recordset
Dim rsProy As ADODB.Recordset
Dim rsGrupoBien As ADODB.Recordset
'---
Dim rs_tipo_auxiliar As New ADODB.Recordset
Dim rs_beneficiario As New ADODB.Recordset
Dim rs_cuentabancaria As New ADODB.Recordset
Dim rs_proyecto As New ADODB.Recordset
Dim rs_UnidadEjecutora As New ADODB.Recordset
Dim rs_Organismo As New ADODB.Recordset
'--- Variables
Dim SaldoIBs As Double
Dim SaldoISus As Double
Dim benef As String
Dim ctabancaria As String
Dim nombanco As String
Dim nomctabancaria As String
'/**********
Dim existereporte As New ADODB.Recordset
Dim reporte As New ADODB.Recordset
Dim BUSCA, swgraba3 As Integer
Dim OrdenarAsc As Boolean
Dim ListaCampos() As String
Dim parametro As String
Dim denominacion As String
Public aux1 As String
Public AUX2 As String
Public aux3 As String
'Dim consul As New ADODB.Recordset
Dim saldobs As Double
Dim saldosus As Double
Dim saldobs1 As Double
Dim saldosus1 As Double
Dim auxsaldobs As Double
Dim auxsaldosus As Double

Dim VARC, VARS1, VARS2, VARA1, VARA2, VARA3 As String
Dim VARAA1, VARAA2, VARAA3, VARNC, VARCA, VARES As String
Dim VarNom1, VarNom2, VarNom3, VarCta As String
Dim VARDB, VARHB, VARPT As Currency

''Private Sub cboaux_LostFocus()
''If Me.cboaux = "01" Then
''Me.Frr01.Visible = True
''End If
''End Sub

Private Sub adosolicitud1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not adosolicitud1.Recordset.BOF) And (Not adosolicitud1.Recordset.EOF) Then
    If adosolicitud1.Recordset("STATUS") = "S" Then
       BtnAprobar.Visible = False
       BtnDesAprobar.Visible = True
    Else
      BtnAprobar.Visible = True
      BtnDesAprobar.Visible = False
    End If
    cbocta.Text = adosolicitud1.Recordset("Cuenta")
    cbosubcta1.Text = adosolicitud1.Recordset("SubCta1")
    cbosubcta2.Text = adosolicitud1.Recordset("SubCta2")
    txtax1.Text = adosolicitud1.Recordset("Aux1")
    Txtax2.Text = adosolicitud1.Recordset("Aux2")
    txtax3.Text = adosolicitud1.Recordset("Aux3")
    lblcuenta.Caption = "-"
    Lblsub1.Caption = "-"
    lbsub2.Caption = IIf(IsNull(adosolicitud1.Recordset("NombreCta")), "-", adosolicitud1.Recordset("NombreCta"))
    Lbl_Aux1 = adosolicitud1.Recordset("denominacion_aux1")
    Lbl_Aux2 = adosolicitud1.Recordset("denominacion_aux2")
    Lbl_Aux3 = adosolicitud1.Recordset("denominacion_aux3")
    LblNom_Aux1 = IIf(IsNull(adosolicitud1.Recordset("Nom_Aux1")), "-", adosolicitud1.Recordset("Nom_Aux1"))
    LblNom_Aux2 = IIf(IsNull(adosolicitud1.Recordset("Nom_Aux2")), "-", adosolicitud1.Recordset("Nom_Aux2"))
    LblNom_Aux3 = IIf(IsNull(adosolicitud1.Recordset("Nom_Aux3")), "-", adosolicitud1.Recordset("Nom_Aux3"))
    TxtDebe.Text = IIf(IsNull(adosolicitud1.Recordset("DebeSaldoIBs")), "0", adosolicitud1.Recordset("DebeSaldoIBs"))
    TxtHaber.Text = IIf(IsNull(adosolicitud1.Recordset("HaberSaldoIBs")), "0", adosolicitud1.Recordset("HaberSaldoIBs"))
    TxtCodAnt.Text = IIf(IsNull(adosolicitud1.Recordset("Cod_Anterior")), "0", adosolicitud1.Recordset("Cod_Anterior"))
  End If
End Sub


Private Sub BtnAprobar01_Click()
    If DtcTAux1.Text = "01" Then
        Lbl_Aux1.Caption = Dtc_benef.Text
        LblNom_Aux1.Caption = Dtc_benefD.Text
    End If
    If DtcTAux2.Text = "01" Then
        Lbl_Aux2.Caption = Dtc_benef.Text
        LblNom_Aux2.Caption = Dtc_benefD.Text
    End If
    If DtcTAux3.Text = "01" Then
        Lbl_Aux3.Caption = Dtc_benef.Text
        LblNom_Aux3.Caption = Dtc_benefD.Text
    End If
    Fra_Benef.Visible = False
End Sub

Private Sub BtnAprobar02_Click()
    If DtcTAux1.Text = "01" Then
        Lbl_Aux1.Caption = Dtc_CtaBco.Text
        LblNom_Aux1.Caption = Dtc_CtaBcoD.Text
    End If
    If DtcTAux2.Text = "01" Then
        Lbl_Aux2.Caption = Dtc_CtaBco.Text
        LblNom_Aux2.Caption = Dtc_CtaBcoD.Text
    End If
    If DtcTAux3.Text = "01" Then
        Lbl_Aux3.Caption = Dtc_CtaBco.Text
        LblNom_Aux3.Caption = Dtc_CtaBcoD.Text
    End If
    Fra_CtaBco.Visible = False
End Sub

Private Sub BtnAprobar03_Click()
    If DtcTAux1.Text = "01" Then
        Lbl_Aux1.Caption = Dtc_Proy.Text
        LblNom_Aux1.Caption = Dtc_ProyD.Text
    End If
    If DtcTAux2.Text = "01" Then
        Lbl_Aux2.Caption = Dtc_Proy.Text
        LblNom_Aux2.Caption = Dtc_ProyD.Text
    End If
    If DtcTAux3.Text = "01" Then
        Lbl_Aux3.Caption = Dtc_Proy.Text
        LblNom_Aux3.Caption = Dtc_ProyD.Text
    End If
    Fra_Proy.Visible = False
End Sub

Private Sub BtnAprobar04_Click()
    If DtcTAux1.Text = "01" Then
        Lbl_Aux1.Caption = Dtc_Uejec.Text
        LblNom_Aux1.Caption = Dtc_UejecD.Text
    End If
    If DtcTAux2.Text = "01" Then
        Lbl_Aux2.Caption = Dtc_Uejec.Text
        LblNom_Aux2.Caption = Dtc_UejecD.Text
    End If
    If DtcTAux3.Text = "01" Then
        Lbl_Aux3.Caption = Dtc_Uejec.Text
        LblNom_Aux3.Caption = Dtc_UejecD.Text
    End If
    Fra_UEjec.Visible = False
End Sub

Private Sub BtnAprobar06_Click()
    If DtcTAux1.Text = "01" Then
        Lbl_Aux1.Caption = Dtc_Dpto.Text
        LblNom_Aux1.Caption = Dtc_DptoD.Text
    End If
    If DtcTAux2.Text = "01" Then
        Lbl_Aux2.Caption = Dtc_Dpto.Text
        LblNom_Aux2.Caption = Dtc_DptoD.Text
    End If
    If DtcTAux3.Text = "01" Then
        Lbl_Aux3.Caption = Dtc_Dpto.Text
        LblNom_Aux3.Caption = Dtc_DptoD.Text
    End If
    Fra_Depto.Visible = False
End Sub

Private Sub BtnAprobar09_Click()
    If DtcTAux1.Text = "01" Then
        Lbl_Aux1.Caption = Dtc_Org.Text
        LblNom_Aux1.Caption = Dtc_OrgD.Text
    End If
    If DtcTAux2.Text = "01" Then
        Lbl_Aux2.Caption = Dtc_Org.Text
        LblNom_Aux2.Caption = Dtc_OrgD.Text
    End If
    If DtcTAux3.Text = "01" Then
        Lbl_Aux3.Caption = Dtc_Org.Text
        LblNom_Aux3.Caption = Dtc_OrgD.Text
    End If
    
    Fra_Org.Visible = False
End Sub

Private Sub cbocta_Click()
  Me.cbosubcta1.Clear
  Me.cbosubcta2.Clear
  'cbosubcta1Nom.Clear
  'cbosubcta2Nom.Clear

  rsplanctas.MoveFirst
  rsplanctas.Find "cuenta=" & "'" & Trim(cbocta.Text) & "'"
  Me.lblcuenta = rsplanctas!NombreCta
  If rscuentas.State = adStateOpen Then rscuentas.Close
  
  rscuentas.Open "SELECT Cuenta, SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "')", db, adOpenKeyset, adLockReadOnly
  Do While Not rscuentas.EOF
    Me.cbosubcta1.AddItem rscuentas!subcta1
    rscuentas.MoveNext
  Loop
  If rscuentas.RecordCount = 0 Then
  Me.cbosubcta1.AddItem "00"
  End If

End Sub

Private Sub cbocta1_Click(Area As Integer)
    DtcCtaNom.BoundText = cbocta1.BoundText
    dtccta.BoundText = cbocta1.BoundText
    dtcsub1.BoundText = cbocta1.BoundText
    dtcsub2.BoundText = cbocta1.BoundText
    DtcAux1.BoundText = cbocta1.BoundText
    DtcAux2.BoundText = cbocta1.BoundText
    DtcAux3.BoundText = cbocta1.BoundText
End Sub

Private Sub cboctaNom_Change()
'  cbosubcta1.Clear
'  cbosubcta2.Clear
''  cbosubcta1Nom.Clear
''  cbosubcta2Nom.Clear
'
'  rsplanctas.MoveFirst
'  rsplanctas.Find "cuenta=" & "'" & Trim(cbocta.Text) & "'"
'  Me.lblcuenta = rsplanctas!NombreCta
'  If rscuentas.State = adStateOpen Then rscuentas.Close
'
'  rscuentas.Open "SELECT Cuenta, SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "')", db, adOpenKeyset, adLockReadOnly
'  Do While Not rscuentas.EOF
'    Me.cbosubcta1.AddItem rscuentas!subcta1
'    rscuentas.MoveNext
'  Loop
'  If rscuentas.RecordCount = 0 Then
'  Me.cbosubcta1.AddItem "00"
'  End If
End Sub

Private Sub cbosubcta1_Click()
On Error GoTo Laberror1
Me.cbosubcta2.Clear

  If rsnombresub1.State = adStateOpen Then rsnombresub1.Close
  rsnombresub1.Open "SELECT NombreCta FROM CC_Plan_Cuentas WHERE   (SubCta2 = '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & (Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
  Me.Lblsub1 = rsnombresub1!NombreCta
  If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
  rssubcuenta.Open "SELECT Cuenta, SubCta1, SubCta2, NombreCta, Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & Trim(Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
  If rssubcuenta.RecordCount = 0 Then
    Me.cbosubcta2 = "00"
    Else
      rssubcuenta.MoveFirst
      Do While Not rssubcuenta.EOF
        Me.cbosubcta2.AddItem rssubcuenta!subcta2
        rssubcuenta.MoveNext
      Loop
    End If

Exit Sub
Laberror1:
If Err.Number = 3021 Then
 MsgBox "Elija una cuenta", vbCritical + vbDefaultButton1
 Me.cbocta.SetFocus
End If
End Sub

Private Sub cbosubcta2_Click()
On Error GoTo labelerr2
   dtccta.Text = cbocta.Text
   dtcsub1.Text = cbosubcta1.Text
   dtcsub2.Text = cbosubcta2.Text
   DtcAux1.Text = txtax1.Text
   DtcAux2.Text = Txtax2.Text
   DtcAux3.Text = txtax3.Text
  Call carga_ctas
   DtcCtaNom.Text = lbsub2.Caption
   cbocta1.Text = Trim(dtccta.Text) + Trim(dtcsub1.Text) + Trim(dtcsub2.Text)
'  With rssubcuenta
'    .MoveFirst
'    .Find "subcta2=" & "'" & Trim(Me.cbosubcta2) & "'"
'    Me.lbsub2 = !NombreCta
'    Me.txtax1 = !aux1
'    Me.Txtax2 = !AUX2
'    Me.txtax3 = !aux3
'    Chkaux1.Enabled = True
'    Chkaux2.Enabled = True
'    Chkaux3.Enabled = True
'    Chkaux1.Value = 1
'    Chkaux2.Value = 1
'    Chkaux3.Value = 1
'    'BtnBuscarA.Enabled = False
'    '--------
'    Call Limpia_combos
'
'    BtnEnviar.Visible = True
'    BtnGrabar.Visible = False
'
'    Select Case !aux1
'      Case "00"
'        'SSTabCuenta.TabEnabled(0) = False
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'        Chkaux1.Enabled = False
'        Chkaux1.Value = 0
'      Case "01"
'        txtbusca1.Visible = True
'        txtbusca1.Top = 1260
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'        'BtnBuscarA.Enabled = True
'      Case "02"
'        cboCtaBancaria.Visible = True
'        cboCtaBancaria.Top = 1260
'        'txtbusca1.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'      Case "03"
'        txtbusca1.Visible = True
'        txtbusca1.Top = 1260
'      Case "08"
'        DTCNomOrg.Visible = True
'        DTCNomOrg.Top = 1260
'        DtCOrg.Visible = True
'        DtCOrg.Top = 1260
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'      Case "09"
'        DtCIdConvenio.Visible = True
'        DtCIdConvenio.Top = 1260
'        DtCDesConvenio.Visible = True
'        DtCDesConvenio.Top = 1260
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'    End Select
'    Select Case !AUX2
'      Case "00"
'        'SSTabCuenta.TabEnabled(0) = False
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'        Chkaux2.Enabled = False
'        Chkaux2.Value = 0
'      Case "01"
'        txtbusca1.Visible = True
'        txtbusca1.Top = 1620
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'        'Me.BtnBuscarA.Enabled = True
'      Case "02"
'        cboCtaBancaria.Visible = True
'        cboCtaBancaria.Top = 1620
'        'txtbusca1.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'      Case "08"
'        DTCNomOrg.Visible = True
'        DTCNomOrg.Top = 1620
'        DtCOrg.Visible = True
'        DtCOrg.Top = 1620
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'      Case "09"
'        DtCIdConvenio.Visible = True
'        DtCIdConvenio.Top = 1620
'        DtCDesConvenio.Visible = True
'        DtCDesConvenio.Top = 1620
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'    End Select
'    Select Case !aux3
'      Case "00"
'        'SSTabCuenta.TabEnabled(0) = False
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'        Chkaux3.Enabled = False
'        Chkaux3.Value = 0
'      Case "01"
'        txtbusca1.Visible = True
'        txtbusca1.Top = 1980
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'        'Me.BtnBuscar.Enabled = True
'      Case "02"
'        cboCtaBancaria.Visible = True
'        cboCtaBancaria.Top = 1980
'        'txtbusca1.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'      Case "08"
'        DTCNomOrg.Visible = True
'        DTCNomOrg.Top = 1980
'        DtCOrg.Visible = True
'        DtCOrg.Top = 1980
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCDesConvenio.Visible = False
'        'DtCIdConvenio.Visible = False
'      Case "09"
'        DtCIdConvenio.Visible = True
'        DtCIdConvenio.Top = 1980
'        DtCDesConvenio.Visible = True
'        DtCDesConvenio.Top = 1980
'        'txtbusca1.Visible = False
'        'cboCtaBancaria.Visible = False
'        'DtCOrg.Visible = False
'        'DTCNomOrg.Visible = False
'    End Select
'  End With
'  'SSTabCuenta_Click (0)
''*******Se filtra si la cuenta es de bancos....
'If Me.cbocta = "1111" And Me.cbosubcta1 = "02" Then
'    Select Case Me.cbosubcta2
'        Case "01"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
'        Case "02"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
'        Case "03"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
'     End Select
'    Me.cboCtaBancaria.Clear
'    If rscta_bancaria.State = 1 Then rscta_bancaria.Close
'    rscta_bancaria.Open sql1, db, adOpenKeyset, adLockReadOnly
'    If rscta_bancaria.RecordCount <> 0 Then
'        rscta_bancaria.MoveFirst
'    End If
'        Do While Not rscta_bancaria.EOF
'          cboCtaBancaria.AddItem rscta_bancaria!Cta_Codigo
'          rscta_bancaria.MoveNext
'        Loop
'    Me.cboCtaBancaria.Visible = True
'    Me.cboCtaBancaria.Text = Me.cboCtaBancaria.List(0)
'    Me.txtbusca1.Visible = False
'    Me.DTGBanco.Visible = True
'    Me.DtGbenef.Visible = False
'    DtgPlanCtas.Visible = False
'    Set Me.DTGBanco.DataSource = rscta_bancaria
'End If
'
''************Se habilita la tabla de beneficiarios
'    If Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01" Then
'        If rsbeneficiario.State = 1 Then rsbeneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario WHERE tipoben_codigo < '20' order by beneficiario_denominacion"
'        rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rsbeneficiario
'        Me.DtGbenef.Visible = True
'        Me.DTGBanco.Visible = False
'        Me.txtbusca1.Visible = True
'        Me.BtnBuscar.Enabled = True
'        Me.cboCtaBancaria.Visible = False
'        DtgPlanCtas.Visible = False
'    End If
''****habilitamos boton de búsqueda
'
'    If Me.txtax1 = "00" Or Me.txtax1 = "02" Then
'        Me.BtnBuscar.Enabled = False
'    Else
'        Me.BtnBuscar.Enabled = True
'    End If
'    If Me.txtax1 = "03" Then
'      Me.txtbusca1.Visible = True
'    End If
'    '-------- habilito datacombos para organismo financiadores
'  If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
'    txtbusca1.Visible = False
'    DtCIdConvenio.Visible = False
'    DtCDesConvenio.Visible = False
'    Dtc_benef.Visible = True
'    DtcCodAux2.Visible = True
'    DtcCodAux3.Visible = True
'    Dtc_benefD.Visible = True
'    DtcDenomAux2.Visible = True
'    DtcDenomAux3.Visible = True
'  Else
'    'txtbusca1.Visible = True
'    'DtCIdConvenio.Visible = True
'    'DtCDesConvenio.Visible = True
'    Dtc_benef.Visible = False
'    DtcCodAux2.Visible = False
'    DtcCodAux3.Visible = False
'    Dtc_benefD.Visible = False
'    DtcDenomAux2.Visible = False
'    DtcDenomAux3.Visible = False
'  End If

    
    Exit Sub
labelerr2:
    If Err.Number = 3021 Then
      MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
      Me.cbosubcta2.SetFocus
    End If

''-------- habilito datacombos para organismo financiadores
'  If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
'    txtbusca1.Visible = False
'    DtCIdConvenio.Visible = False
'    DtCDesConvenio.Visible = False
'    Dtc_benef.Visible = True
'    DtcCodAux2.Visible = True
'    DtcCodAux3.Visible = True
'    Dtc_benefD.Visible = True
'    DtcDenomAux2.Visible = True
'    DtcDenomAux3.Visible = True
'  Else
'    'txtbusca1.Visible = True
'    'DtCIdConvenio.Visible = True
'    'DtCDesConvenio.Visible = True
'    Dtc_benef.Visible = False
'    DtcCodAux2.Visible = False
'    DtcCodAux3.Visible = False
'    Dtc_benefD.Visible = False
'    DtcDenomAux2.Visible = False
'    DtcDenomAux3.Visible = False
'  End If
'--------------

'    Me.Chkaux2.Value = False
'    Me.Chkaux3.Value = False
'    DtCIdConvenio.Visible = False
'    DtCDesConvenio.Visible = False
'    Me.DTCNomOrg.Visible = False
'    Me.DtcOrg.Visible = False
'    Me.Txtbusca2.Visible = True
'    Me.BtnBuscar.Enabled = True
'    Me.Chkaux1.Enabled = True
'    Me.Chkaux2.Enabled = True
'    Me.Chkaux3.Enabled = True
'    Me.txtax1.Enabled = True
'    Me.Txtax2.Enabled = True
'    Me.txtax3.Enabled = True
'    Me.txtbusca1.Enabled = True
'    Me.Txtbusca2.Enabled = True
'    Me.Txtbusca3.Enabled = True
'    With rssubcuenta
'      .MoveFirst
'      .Find "subcta2=" & "'" & Trim(Me.cbosubcta2) & "'"
'      Me.lbsub2 = !NombreCta
'      Me.txtax1 = !aux1
'      Me.Txtax2 = !aux2
'      Me.txtax3 = !aux3
'      If !aux1 = "00" Then
'        Me.Chkaux1.Enabled = False
'        Me.txtax1.Enabled = False
'        Me.txtbusca1.Enabled = False
'      End If
'      If !aux2 = "00" Then
'        Me.Chkaux2.Enabled = False
'        Me.Txtax2.Enabled = False
'        Me.Txtbusca2.Enabled = False
'      End If
'      If !aux3 = "00" Then
'        Me.Chkaux3.Enabled = False
'        Me.txtax3.Enabled = False
'        Me.Txtbusca3.Enabled = False
'      End If
'
'      If Me.Chkaux1.Enabled = True And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = False Then
'        Me.Chkaux1.Value = 1
'      End If
'      If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = True And Me.Chkaux3.Enabled = False Then
'        Me.Chkaux2.Value = 1
'      End If
'      If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = True Then
'        Me.Chkaux3.Value = 1
'      End If
'    End With
'
'    If (Me.txtax1 <> "00" And Me.txtax1 <> "01" And Me.txtax1 <> "02" And txtax1 <> "09" And txtax1 <> "08") Then
'      f = 1
'      Me.Chkaux1.Enabled = False
'      Me.txtax1.Enabled = False
'      Me.txtbusca1.Enabled = False
'      Me.DTCNomOrg.Visible = False
'      Me.DtcOrg.Visible = False
'    End If
'    If (Me.Txtax2 <> "00" And Me.Txtax2 <> "01" And Me.Txtax2 <> "02" And Me.Txtax2 <> "09") Then
'      f = 2
'      Me.Chkaux2.Enabled = False
'      Me.Txtax2.Enabled = False
'      Me.Txtbusca2.Enabled = False
'      Me.DTCNomOrg.Visible = False
'      Me.DtcOrg.Visible = False
'    End If
'    'g--
'    If (Me.Txtax2 = "08") Then
'      f = 8
'      Me.Chkaux2.Enabled = True
'      Me.Txtax2.Enabled = True
'      Me.Txtbusca2.Enabled = True
'      Me.Txtbusca2.Visible = False
'      Me.DTCNomOrg.Visible = True
'      Me.DtcOrg.Visible = True
'    End If
'
'    If (Me.txtax1 = "09") Then
'      f = 9
'      Me.Chkaux2.Enabled = True
'      Me.Txtax2.Enabled = True
'      Me.txtbusca1.Enabled = True
'      Me.txtbusca1.Visible = False
'      'Me.DTCNomOrg.Visible = True
'      'Me.DtcOrg.Visible = True
'      DtCDesConvenio.Visible = True
'      DtCIdConvenio.Visible = True
'    End If
'
'
'
'    'g--
'    If (Me.txtax3 <> "00" And Me.txtax3 <> "01" And Me.txtax3 <> "02") Then
'      f = 3
'      Me.Chkaux3.Enabled = False
'      Me.txtax3.Enabled = False
'      Me.Txtbusca3.Enabled = False
'      Me.DTCNomOrg.Visible = False
'      Me.DtcOrg.Visible = False
'    End If
'    'If f = 1 Or f = 2 Or f = 3 Then
'      '  MsgBox "Por el momento solo se trabaja con Auxiliares de Beneficiarios y Ctas. Corrientes", vbInformation + vbDefaultButton1, "Atencion"
'        Me.cbocta.SetFocus
'    'End If
'    If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = False Then
'    Me.BtnBuscar.Enabled = False
'    Else
''    Me.BtnGrabar.Enabled = False
'    End If
'If (Me.cbosubcta1.Text) = "00" And Me.cbosubcta2.Text = "00" Then
'    'Me.BtnGrabar.Enabled = True
'End If
''*******Se filtra si la cuenta es de bancos....
'If Me.cbocta = "1111" And Me.cbosubcta1 = "02" Then
'    Select Case Me.cbosubcta2
'        Case "01"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
'        Case "02"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
'        Case "03"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
'     End Select
'    Me.cboCtaBancaria.Clear
'    If rscta_bancaria.State = 1 Then rscta_bancaria.Close
'    rscta_bancaria.Open sql1, db, adOpenKeyset, adLockReadOnly
'    If rscta_bancaria.RecordCount <> 0 Then
'        rscta_bancaria.MoveFirst
'    End If
'        Do While Not rscta_bancaria.EOF
'          cboCtaBancaria.AddItem rscta_bancaria!cta_codigo
'          rscta_bancaria.MoveNext
'        Loop
'    Me.cboCtaBancaria.Visible = True
'    Me.cboCtaBancaria.Text = Me.cboCtaBancaria.List(0)
'    Me.txtbusca1.Visible = False
'    Me.DTGBanco.Visible = True
'    Me.DtGbenef.Visible = False
'    Set Me.DTGBanco.DataSource = rscta_bancaria
'End If
'
''************Se habilita la tabla de beneficiarios
'    If Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01" Then
'        If rsBeneficiario.State = 1 Then rsBeneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario order by beneficiario_denominacion"
'        rsBeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rsBeneficiario
'        Me.DtGbenef.Visible = True
'        Me.DTGBanco.Visible = False
'        Me.txtbusca1.Visible = True
'        Me.BtnBuscar.Enabled = True
'        Me.cboCtaBancaria.Visible = False
'    End If
''****habilitamos boton de búsqueda
'    If Me.txtax1 = "00" Or Me.txtax1 = "02" Then
'        Me.BtnBuscar.Enabled = False
'    Else
'        Me.BtnBuscar.Enabled = True
'    End If
'
'    Exit Sub
'labelerr2:
'    If err.Number = 3021 Then
'      MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
'      Me.cbosubcta2.SetFocus
'    End If
End Sub

Private Sub Limpia_combos()
    txtbusca1.Visible = False
    Dtc_benef.Visible = False
    Dtc_benefD.Visible = False
    DtCDesConvenio.Visible = False
    DtCIdConvenio.Visible = False
'    DtcCodAux2.Visible = False
'    cboCtaBancaria.Visible = False
    DtcDenomAux2.Visible = False
'    DtcCodAux3.Visible = False
    DtcDenomAux3.Visible = False
    DtCOrg.Visible = False
    DTCNomOrg.Visible = False
End Sub

Private Sub cbosubcta2_LostFocus()
    If (Me.txtax1 = "01" And Me.Txtax2 = "00" And Me.txtax3 = "00") Then
      Me.DtGbenef.Visible = True
      Me.DTGBanco.Visible = False
      DtgPlanCtas.Visible = False
    End If
    If (Me.txtax1 = "00" And Me.Txtax2 = "00" And Me.txtax3 = "00") Then
    End If
End Sub

Private Sub Chkaux1_Click()
'habilita el grid de beneficiarios
    If Me.Chkaux1.Value = 1 And (Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01") Then
        Me.DtGbenef.Visible = True
        Me.DTGBanco.Visible = False
        DtgPlanCtas.Visible = False
        DtcTAux1.Visible = True
    End If
    'habilita el grid de cuentas corrientes
    If Me.Chkaux1.Value = 1 And (Me.txtax1 = "02" Or Me.Txtax2 = "02" Or Me.txtax3 = "02") Then
        Me.DTGBanco.Visible = True
        Me.DtGbenef.Visible = False
        DtgPlanCtas.Visible = False
    End If
End Sub
Private Sub Chkaux2_Click()
'habilita el grid de beneficiarios
    If Me.Chkaux2.Value = 1 And (Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01") Then
        Me.DtGbenef.Visible = True
        Me.DTGBanco.Visible = False
        DtgPlanCtas.Visible = False
    End If
    'habilita el grid de cuentas corrientes
    If Me.Chkaux2.Value = 1 And (Me.txtax1 = "02" Or Me.Txtax2 = "02" Or Me.txtax3 = "02") Then
        Me.DTGBanco.Visible = True
        Me.DtGbenef.Visible = False
        DtgPlanCtas.Visible = False
    End If
End Sub
Private Sub Chkaux3_Click()
'habilita el grid de beneficiarios
    If Me.Chkaux3.Value = 1 And (Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01") Then
        Me.DtGbenef.Visible = True
        Me.DTGBanco.Visible = False
        DtgPlanCtas.Visible = False
    End If
    'habilita el grid de cuentas corrientes
    If Me.Chkaux3.Value = 1 And (Me.txtax1 = "02" Or Me.Txtax2 = "02" Or Me.txtax3 = "02") Then
        Me.DTGBanco.Visible = True
        Me.DtGbenef.Visible = False
        DtgPlanCtas.Visible = False
    End If
End Sub

'Private Sub BtnGrabar_Click()
'Call existecta(Trim(Me.cbocta), Trim(Me.cbosubcta1), Trim(Me.cbosubcta2))
'If lcta = "S" Then
'    If Me.cbocta.Text = "" Then
'        MsgBox "Elija una cuenta", vbCritical + vbDefaultButton1
'        Me.cbocta.SetFocus
'        Exit Sub
'    End If
'    If Me.cbosubcta1.Text = "" Then
'        MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
'        Me.cbosubcta1.SetFocus
'        Exit Sub
'    End If
'    If Me.cbosubcta2.Text = "" Then
'        MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
'        Me.cbosubcta2.SetFocus
'        Exit Sub
'    End If
'    If Me.txtax1 = "02" Then
'        If Me.cboCtaBancaria.Text = "" Then
'            MsgBox "Elija una cuenta bancaria", vbCritical + vbDefaultButton1
'            Me.cboCtaBancaria.SetFocus
'            Exit Sub
'        End If
'    End If
'    If Me.txtax1 = "01" And Me.Chkaux1.Value = 1 Then
'        If Me.txtbusca1.Text = "" Then
'            MsgBox "Escriba un beneficiario", vbCritical + vbDefaultButton1
'            Me.txtbusca1.SetFocus
'            Exit Sub
'        End If
'    End If
'    If Me.txtax1 = "02" Or Txtax2 = "02" Or txtax3 = "02" Then
'        If Me.cboCtaBancaria = "" Then
'            MsgBox "Seleccione una Cuenta Bancaria", vbCritical + vbDefaultButton1
'            Exit Sub
'        End If
'    End If
''    If Me.txtax1 = "01" Or Txtax2 = "01" Or txtax3 = "01" Then
''        If txtbusca1.Text = "" Then
''            MsgBox "Introduzca un Beneficiario", vbCritical + vbDefaultButton1
''            Exit Sub
''        End If
''    End If
'    If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
'        MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
'        Exit Sub
'    End If
'    If Me.txtax1 = "00" And Me.Txtax2 = "00" And Me.txtax3 = "00" Then
'    '****si la cuenta no tiene auxiliares
'        Call Mayor000
'    Else
'        '****llamada al store procedure de Saldos para beneficiarios "SaldoBenef
'        If Chkaux1.Value = 0 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
'          MsgBox "Seleccione una opción", vbExclamation + vbDefaultButton1, "REPORTES"
'          Exit Sub
'        End If
'        '***** si el aux es 1
'           If Chkaux1.Value = 1 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
'              Select Case Trim(txtax1.Text)
'                Case "01"
'
'                  reporteBeneficiario  'procedimiento para reporte con beneficiarios
'                Case "02"
'                  ReporteCtaBancaria
'                Case "08"
'                  'ReporteOrg   ' procedimiento para reporte con organismos
'                Case "09"
'                  reporteconvenio
'              End Select
'
'
'           End If
'
'           If Chkaux1.Value = 0 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
'              Select Case Trim(Txtax2.Text)
'                Case "01"
'                  reporteBeneficiario  'procedimiento para reporte con beneficiarios
'                Case "02"
'                  ReporteCtaBancaria
'                Case "08"
'                '  ReporteOrg   ' procedimiento para reporte con organismos
'              End Select
'           End If
'
'           If Chkaux1.Value = 0 And Chkaux2.Value = 0 And Chkaux3.Value = 1 Then
'              Select Case Trim(txtax3.Text)
'                Case "01"
'                  reporteBeneficiario  'procedimiento para reporte con beneficiarios
'                Case "02"
'                  ReporteCtaBancaria
'                Case "08"
'                  ReporteOrg   ' procedimiento para reporte con organismos
'              End Select
'           End If
'
'           If Chkaux1.Value = 1 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
'              If Trim(txtax1.Text) = "01" And Trim(Txtax2.Text) = "08" Then
'                  If rsBeneficiario.State = 1 Then rsBeneficiario.Close
'                  rsBeneficiario.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
'                  If rsBeneficiario.RecordCount <> 0 Then
'                    nombenef = rsBeneficiario!beneficiario_denominacion
'                  Else
'                    nombenef = ""
'                  End If
'                  ReporteAux1_2 Trim(txtbusca1.Text), Trim(DtCOrg.Text), Trim(txtax1.Text), Trim(Txtax2.Text), Trim(txtax3.Text), nombenef, Trim(DTCNomOrg.Text)
'              End If
'           End If
'     End If
'    End If
'
'    '---si el auxiliar es 2
''       If (Me.txtax1 = "02") Or (Me.Txtax2 = "02") Or (Me.txtax3 = "02") Then
''            ReporteCtaBancaria
''        End If
''    '---si el auxiliar es 8
''        If Me.Chkaux2.Value = 1 And ((Me.Txtax2 = "08")) Then 'Or (Me.Txtax2 = "02") Or (Me.txtax3 = "02"))
''          ReporteOrg
''        End If
''        If Me.Chkaux1.Value = 1 And Chkaux2.Value = 1 Then
''          ReporteAux1_2
''        End If
''    '---auxiliar 1 y  2
'    'End If
''End If
'End Sub

Private Sub cmdBusca_Click()
    Me.Fra_Busqueda.Visible = True
End Sub

Private Sub BtnAprobar_Click()
' INI JQA JUL-2005
If adosolicitud1.Recordset("STATUS") = "S" Or adosolicitud1.Recordset("STATUS") = "E" Then
    MsgBox "El registro ya esta Aprobado, o fue Anulado... Verifique los Datos ....!!!", vbCritical, "ATENCION !!"
Else
    sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
    If sino = vbYes Then
        adosolicitud1.Recordset("STATUS") = "S"
        adosolicitud1.Recordset.Update
        adosolicitud1.Recordset.Requery
        'adosolicitud.Refresh
        'adosolicitud.Recordset.Move marca1 - 1
    End If
End If
'   Actualiza Totales por Partida (Insumo del POA)
'db.Execute "UPDATE ao_Solicitud_detalle SET ao_Solicitud_detalle.monto_bolivianos = (SELECT SUM(Total_venta) FROM ao_solicitud_bien WHERE ao_solicitud_bien.CODIGO_UNIDAD = ao_Solicitud_detalle.CODIGO_UNIDAD AND ao_solicitud_bien.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND ao_solicitud_bien.cod_univ = ao_Solicitud_detalle.par_codigo) FROM ao_Solicitud_detalle, ao_solicitud_bien"
End Sub

Private Sub BtnBuscarB_Click()
    Me.Fra_BusquedaC.Visible = True
    Me.CboCampoC.Text = Me.CboCampoC.List(0)
    'Me.CboOperador.Text = Me.CboCampo.List(0)
    CboOperadorC.Text = "como"
    DtgPlanCtas.Visible = True
    DtGbenef.Visible = False
    DTGBanco.Visible = False
End Sub

Private Sub BtnBuscar_Click()
    TxtCtaNom = ""
    'Fra_BuscaGral.Visible = True
    frmListaPlanCtas.Show vbModal
End Sub

Private Sub BtnBuscarA_Click()
    'Busca Cuenta Contable
    Me.Fra_Busqueda.Visible = True
    Me.CboCampo.Text = Me.CboCampo.List(0)
    'Me.CboOperador.Text = Me.CboCampo.List(0)
    CboOperador.Text = "como"
    DtgPlanCtas.Visible = False
'    DtGbenef.Visible = True
'    DTGBanco.Visible = True
End Sub

'Private Sub BtnCancelar_Click()
'  'Me.txtbusca1 = ""
' ' Me.Txtbusca2 = ""
''  Me.Txtbusca3 = ""
''  Me.cbocta.SetFocus
'  'Me.txtaux = ""
'  'Me.cboaux.Text = Me.cboaux.List(0)
'End Sub

Private Sub BtnDesAprobar_Click()
  If adosolicitud1.Recordset!Status = "S" And adosolicitud1.Recordset!Verificado = "N" Then
    sino = MsgBox("Esta seguro(a) de Desverificar(Ok) el registro?", vbYesNo + vbExclamation, "Confirmación")
    If sino = vbYes Then
        adosolicitud1.Recordset!Status = "N"
        adosolicitud1.Recordset.Update
        BtnDesAprobar.Visible = False
        BtnAprobar.Visible = True
        BtnAprobar.Enabled = True
        BtnModificar.Enabled = True
        BtnEliminar.Enabled = True
        'BtnEnviar.Enabled = False
'        CmdDetallePoa.Enabled = True
    End If
  Else
    MsgBox "No se puede DESAPROBAR, si el registro esta en uso y VERIFICADO !!! ...."
  End If
End Sub

Private Sub BtnEnviar_Click()

  If adosolicitud1.Recordset!Status <> "S" And swgraba3 <> 2 Then
' If swgraba3 <> 0 Then
    DataGrid3.Columns("Cuenta").Value = cbocta.Text
    DataGrid3.Columns("SubCta1").Value = cbosubcta1.Text
    DataGrid3.Columns("SubCta2").Value = cbosubcta2.Text
    DataGrid3.Columns("Aux1").Value = txtax1.Text
    DataGrid3.Columns("Aux2").Value = Txtax2.Text
    DataGrid3.Columns("Aux3").Value = txtax3.Text
    DataGrid3.Columns("NombreCta").Value = lbsub2.Caption
    DataGrid3.Columns("DebeSaldoIBs").Value = IIf(TxtDebe.Text = "", "0", TxtDebe.Text)
    DataGrid3.Columns("HaberSaldoIBs").Value = IIf(TxtHaber.Text = "", "0", TxtHaber.Text)
    DataGrid3.Columns("Cod_Anterior").Value = IIf(TxtCodAnt.Text = "", "-", TxtCodAnt.Text)
    DataGrid3.Columns("cod_cta").Value = cbocta1.Text
    
    DataGrid3.Columns("denominacion_aux1").Value = Lbl_Aux1.Caption
    DataGrid3.Columns("Nom_Aux1").Value = LblNom_Aux1.Caption
        
    DataGrid3.Columns("denominacion_aux2").Value = Lbl_Aux2.Caption
    DataGrid3.Columns("Nom_Aux2") = LblNom_Aux2.Caption
        
    DataGrid3.Columns("denominacion_aux3").Value = Lbl_Aux3.Caption
    DataGrid3.Columns("Nom_Aux3") = LblNom_Aux3.Caption

'    If txtax1.Text = "01" Then
'        DataGrid3.Columns("denominacion_aux1").Value = txtbusca1.Text
'        DataGrid3.Columns("Nom_Aux1").Value = LblNom_Aux1.Caption
'    End If
'    If txtax1.Text = "02" Then
'        DataGrid3.Columns("denominacion_aux1").Value = cboCtaBancaria.Text
'        DataGrid3.Columns("Nom_Aux1").Value = LblNom_Aux1.Caption
'    End If
'    If txtax1.Text = "03" Then
'        DataGrid3.Columns("denominacion_aux1").Value = DtcProy.Text
'        DataGrid3.Columns("Nom_Aux1").Value = DtcProyDes.Text
'    End If
'    If txtax1.Text = "05" Then
'        DataGrid3.Columns("denominacion_aux1").Value = DtcGrBien.Text
'        DataGrid3.Columns("Nom_Aux1").Value = DtcGrBienDes.Text
'    End If
'    If txtax1.Text = "09" Then  'financiador
'        DataGrid3.Columns("denominacion_aux1").Value = DtCIdConvenio.Text
'        DataGrid3.Columns("Nom_Aux1").Value = DtCDesConvenio.Text
'    End If
'    If txtax1.Text = "00" Then
'        DataGrid3.Columns("denominacion_aux1").Value = ""
'        DataGrid3.Columns("Nom_Aux1").Value = ""
'    End If
    
'    If Txtax2.Text = "01" Then
'        DataGrid3.Columns("denominacion_aux2").Value = txtbusca1.Text
'        DataGrid3.Columns("Nom_Aux2") = LblNom_Aux2.Caption
'    End If
'    If Txtax2.Text = "02" Then
'        DataGrid3.Columns("denominacion_aux2").Value = cboCtaBancaria.Text
'        DataGrid3.Columns("Nom_Aux2") = LblNom_Aux2.Caption
'    End If
'    If Txtax2.Text = "03" Then
'        DataGrid3.Columns("denominacion_aux2").Value = DtcProy.Text
'        DataGrid3.Columns("Nom_Aux2").Value = DtcProyDes.Text
'    End If
'    If Txtax2.Text = "05" Then
'        DataGrid3.Columns("denominacion_aux2").Value = DtcGrBien.Text
'        DataGrid3.Columns("Nom_Aux2") = DtcGrBienDes.Text
'    End If
'    If Txtax2.Text = "09" Then
'        DataGrid3.Columns("denominacion_aux2").Value = DtCIdConvenio.Text
'        DataGrid3.Columns("Nom_Aux2") = DtCDesConvenio.Text
'    End If
'    If Txtax2.Text = "00" Then
'        DataGrid3.Columns("denominacion_aux2").Value = ""
'        DataGrid3.Columns("Nom_Aux2") = ""
'    End If
'
'    If txtax3.Text = "01" Then
'        DataGrid3.Columns("denominacion_aux3").Value = txtbusca1.Text
'        DataGrid3.Columns("Nom_Aux3") = LblNom_Aux3.Caption
'    End If
'    If txtax3.Text = "02" Then
'        DataGrid3.Columns("denominacion_aux3").Value = cboCtaBancaria.Text
'        DataGrid3.Columns("Nom_Aux3") = LblNom_Aux3.Caption
'    End If
'    If txtax3.Text = "03" Then
'        DataGrid3.Columns("denominacion_aux3").Value = DtcProy.Text
'        DataGrid3.Columns("Nom_Aux3").Value = DtcProyDes.Text
'    End If
'    If txtax3.Text = "05" Then
'        DataGrid3.Columns("denominacion_aux3").Value = DtcGrBien.Text
'        DataGrid3.Columns("Nom_Aux3") = DtcGrBienDes.Text
'    End If
'    If txtax3.Text = "09" Then
'        DataGrid3.Columns("denominacion_aux3").Value = DtCIdConvenio.Text
'        DataGrid3.Columns("Nom_Aux3") = DtCDesConvenio.Text
'    End If
'    If txtax3.Text = "00" Then
'        DataGrid3.Columns("denominacion_aux3").Value = ""
'        DataGrid3.Columns("Nom_Aux3") = ""
'    End If

   'If adosolicitud1.Recordset.RecordCount > 0 And Not IsNull(DataGrid3.Columns("NombreCta").Value) And (DataGrid3.Columns("NombreCta").Value) <> "" Then
    If Not IsNull(DataGrid3.Columns("NombreCta").Value) And (DataGrid3.Columns("NombreCta").Value) <> "" Then
      'If adosolicitud1.Recordset!Status = "N" Then
        VARC = DataGrid3.Columns("Cuenta").Value
        VARS1 = DataGrid3.Columns("SubCta1").Value
        VARS2 = DataGrid3.Columns("SubCta2").Value
        VARA1 = DataGrid3.Columns("Aux1").Value
        VARA2 = DataGrid3.Columns("Aux2").Value
        VARA3 = DataGrid3.Columns("Aux3").Value
        VARAA1 = DataGrid3.Columns("denominacion_aux1").Value
        VARAA2 = DataGrid3.Columns("denominacion_aux2").Value
        VARAA3 = DataGrid3.Columns("denominacion_aux3").Value
        Dim rs_ao_bien As New ADODB.Recordset
        Set rs_ao_bien = New ADODB.Recordset
        If rs_ao_bien.State = 1 Then rs_ao_bien.Close
        
        tot_form = 0
        rs_ao_bien.Open "select COUNT(*) AS tot_form from co_balanceApertura where Cuenta = '" & VARC & "' and SubCta1 = '" & VARS1 & "' and SubCta2 = '" & VARS2 & "' and Aux1 = '" & VARA1 & "' and Aux2 = '" & VARA2 & "' and Aux3 = '" & VARA3 & "' and denominacion_aux1 = '" & VARAA1 & "' and denominacion_aux2 = '" & VARAA2 & "' and denominacion_aux3 = '" & VARAA3 & "'  ", db, adOpenDynamic, adLockOptimistic
        If rs_ao_bien!tot_form > swgraba3 Then
        'If rs_ao_bien!tot_form > 0 Then
            MsgBox "No se puede Guardar un registro ya EXISTENTE, verifique por favor !!...", vbInformation, "Formulario"
            cbosubcta2.SetFocus
            'DataGrid3.SetFocus
            Exit Sub
        Else
            If rs_ao_bien.State = 1 Then rs_ao_bien.Close
            
            'marca1 = adosolicitud1.Recordset.Bookmark
            VARNC = DataGrid3.Columns("NombreCta").Value
            VARDB = DataGrid3.Columns("DebeSaldoIBs").Value
            VARHB = DataGrid3.Columns("HaberSaldoIBs").Value
            VARCA = DataGrid3.Columns("Cod_Anterior").Value
            VARES = DataGrid3.Columns("status").Value
            VCorrel = DataGrid3.Columns("correl").Value
            VarNom1 = DataGrid3.Columns("Nom_Aux1").Value
            VarNom2 = DataGrid3.Columns("Nom_Aux2").Value
            VarNom3 = DataGrid3.Columns("Nom_Aux3").Value
            VarCta = DataGrid3.Columns("cod_cta").Value

            db.Execute "UPDATE co_balanceApertura SET cuenta= '" & VARC & "', subcta1='" & VARS1 & "', subcta2 = '" & VARS2 & "', aux1 = '" & VARA1 & "', aux2 = '" & VARA2 & "', aux3 = '" & VARA3 & "', denominacion_aux1 = '" & VARAA1 & "', denominacion_aux2 = '" & VARAA2 & "', denominacion_aux3 = '" & VARAA3 & "', NombreCta = '" & RTrim(VARNC) & "', DebeSaldoIBs = " & VARDB & ", HaberSaldoIBs = " & VARHB & ", Cod_Anterior = '" & VARCA & "', Status = '" & VARES & "', Nom_Aux1 = '" & VarNom1 & "', Nom_Aux2 = '" & VarNom2 & "', Nom_Aux3 = '" & VarNom3 & "', cod_cta = '" & VarCta & "'  WHERE correl = '" & VCorrel & "' "
            marca1 = rstAo_solicitud1.Bookmark
            Call Abre_Balance
            If swgraba3 = 0 Then
                rstAo_solicitud1.MoveLast
            Else
                rstAo_solicitud1.Bookmark = marca1
            End If
            swgraba3 = 2
            Frame1.Enabled = False
            Frame4.Visible = False
            frmabm.Visible = True
            FrmGraba.Visible = False
            BtnEnviar.Visible = False
            BtnGrabar.Visible = False
            
            Call Limpia_combos
                    
            DataGrid3.AllowAddNew = False
            DataGrid3.AllowDelete = False
            DataGrid3.AllowUpdate = False
            DataGrid3.Enabled = True
        End If
      'Else
      '   MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Formulario 04"
      'End If
    Else
         MsgBox "Verifique los datos para continuar ... ", vbInformation, "Formulario"
    End If
 Else
    MsgBox "ERROR, NO se puede modificar un registro aprobado..."
 End If
End Sub

Private Sub BtnGrabar_Click()
On Error GoTo Error
  'If adosolicitud1.Recordset!Status = "N" Or IsNull(adosolicitud1.Recordset!Status) Then
  If adosolicitud1.Recordset!Status <> "S" And swgraba3 <> 2 Then
' If swgraba3 <> 0 Then
   'If adosolicitud1.Recordset.RecordCount > 0 And Not IsNull(DataGrid3.Columns("NombreCta").Value) And (DataGrid3.Columns("NombreCta").Value) <> "" Then
    If Not IsNull(DataGrid3.Columns("NombreCta").Value) And (DataGrid3.Columns("NombreCta").Value) <> "" Then
      'If adosolicitud1.Recordset!Status = "N" Then
        VARC = DataGrid3.Columns("Cuenta").Value
        VARS1 = DataGrid3.Columns("SubCta1").Value
        VARS2 = DataGrid3.Columns("SubCta2").Value
        VARA1 = DataGrid3.Columns("Aux1").Value
        VARA2 = DataGrid3.Columns("Aux2").Value
        VARA3 = DataGrid3.Columns("Aux3").Value
        VARAA1 = DataGrid3.Columns("denominacion_aux1").Value
        VARAA2 = DataGrid3.Columns("denominacion_aux2").Value
        VARAA3 = DataGrid3.Columns("denominacion_aux3").Value
        Dim rs_ao_bien As New ADODB.Recordset
        Set rs_ao_bien = New ADODB.Recordset
        If rs_ao_bien.State = 1 Then rs_ao_bien.Close
        
        tot_form = 0
        rs_ao_bien.Open "select COUNT(*) AS tot_form from co_balanceApertura where Cuenta = '" & VARC & "' and SubCta1 = '" & VARS1 & "' and SubCta2 = '" & VARS2 & "' and Aux1 = '" & VARA1 & "' and Aux2 = '" & VARA2 & "' and Aux3 = '" & VARA3 & "' and denominacion_aux1 = '" & VARAA1 & "' and denominacion_aux2 = '" & VARAA2 & "' and denominacion_aux3 = '" & VARAA3 & "'  ", db, adOpenDynamic, adLockOptimistic
        If rs_ao_bien!tot_form > swgraba3 Then
            MsgBox "No se puede Guardar un registro ya EXISTENTE, verifique por favor !!...", vbInformation, "Formulario 04"
            DataGrid3.SetFocus
            Exit Sub
        Else
            If rs_ao_bien.State = 1 Then rs_ao_bien.Close
            
            'marca1 = adosolicitud1.Recordset.Bookmark
            VARNC = DataGrid3.Columns("NombreCta").Value
            VARDB = DataGrid3.Columns("DebeSaldoIBs").Value
            VARHB = DataGrid3.Columns("HaberSaldoIBs").Value
            VARCA = DataGrid3.Columns("Cod_Anterior").Value
            VARES = DataGrid3.Columns("status").Value
            VCorrel = DataGrid3.Columns("correl").Value
            VarNom1 = DataGrid3.Columns("Nom_Aux1").Value
            VarNom2 = DataGrid3.Columns("Nom_Aux2").Value
            VarNom3 = DataGrid3.Columns("Nom_Aux3").Value
            marca1 = rstAo_solicitud1.Bookmark
'            Call Abre_Balance
    '        'MarcaB = rstAo_solicitud1.Bookmark
            'rstAo_solicitud1.Bookmark = marca1
            'db.Execute "UPDATE co_balanceApertura SET cuenta= '" & VARC & "', subcta1='" & VARS1 & "', subcta2 = '" & VARS2 & "', aux1 = '" & VARA1 & "', aux2 = '" & VARA2 & "', aux3 = '" & VARA3 & "', denominacion_aux1 = '" & VARAA1 & "', denominacion_aux2 = '" & VARAA2 & "', denominacion_aux3 = '" & VARAA3 & "', NombreCta = '" & VARNC & "', DebeSaldoIBs = " & VARDB & ", HaberSaldoIBs = " & VARHB & ", Cod_Anterior = '" & VARCA & "', Status = '" & VARES & "' WHERE correl = '" & VCorrel & "' "
            db.Execute "UPDATE co_balanceApertura SET cuenta= '" & VARC & "', subcta1='" & VARS1 & "', subcta2 = '" & VARS2 & "', aux1 = '" & VARA1 & "', aux2 = '" & VARA2 & "', aux3 = '" & VARA3 & "', denominacion_aux1 = '" & VARAA1 & "', denominacion_aux2 = '" & VARAA2 & "', denominacion_aux3 = '" & VARAA3 & "', NombreCta = '" & VARNC & "', DebeSaldoIBs = " & VARDB & ", HaberSaldoIBs = " & VARHB & ", Cod_Anterior = '" & VARCA & "', Status = '" & VARES & "', Nom_Aux1 = '" & VarNom1 & "', Nom_Aux2 = '" & VarNom2 & "', Nom_Aux3 = '" & VarNom3 & "'  WHERE correl = '" & VCorrel & "' "
            swgraba3 = 2
            Call Abre_Balance
            rstAo_solicitud1.Bookmark = marca1
            'rstAo_solicitud1.MoveLast
            Frame1.Enabled = False
            Frame4.Visible = False
            frmabm.Visible = True
            FrmGraba.Visible = False
            BtnEnviar.Visible = False
            BtnGrabar.Visible = False
            
            Call Limpia_combos
                    
            DataGrid3.AllowAddNew = False
            DataGrid3.AllowDelete = False
            DataGrid3.AllowUpdate = False
        End If
      'Else
      '   MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Formulario 04"
      'End If
    Else
         MsgBox "Verifique los datos para continuar ... ", vbInformation, "Formulario 04"
    End If
 Else
    MsgBox "ERROR, NO se puede modificar un registro aprobado..."
 End If
 'End If
Error:
ErrorB = Err.Number
    If Err.Number = -2147467259 Then
        MsgBox "ERROR B.A.: El registro actual ya fue GUARDADO anteriormente, verifique por favor !!...", vbInformation, "Formulario 04"
    End If
    If Err.Number = -2147217887 Then
        MsgBox "Se producjo un error desconocido!...", vbCritical + vbOKOnly, "Error..."
    End If

End Sub

Private Sub BtnImprimirA_Click()
    cc_balapertura.Show
End Sub

Private Sub BtnImprimir_Click()
'oooooooooooooo
Call existecta(Trim(Me.cbocta), Trim(Me.cbosubcta1), Trim(Me.cbosubcta2))
If lcta = "S" Then
    If Me.cbocta.Text = "" Then
        MsgBox "Elija una cuenta", vbExclamation + vbDefaultButton1
        Me.cbocta.SetFocus
        Exit Sub
    End If
    If Me.cbosubcta1.Text = "" Then
        MsgBox "Elija una subcuenta", vbExclamation + vbDefaultButton1
        Me.cbosubcta1.SetFocus
        Exit Sub
    End If
    If Me.cbosubcta2.Text = "" Then
        MsgBox "Elija una subcuenta", vbExclamation + vbDefaultButton1
        Me.cbosubcta2.SetFocus
        Exit Sub
    End If
    If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
        MsgBox "Seleccione un rango de fechas correcto", vbExclamation + vbDefaultButton1
        Exit Sub
    End If
    '----preguntar si los tres chek estan en 1
    If Chkaux1.Value = 1 And Chkaux2.Value = 1 And Chkaux3.Value = 1 Then
      If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
      '---reporte de 2 organismos
      With CryConv_Conv
         .Destination = crptToWindow
         .WindowState = crptMaximized
         .WindowShowPrintSetupBtn = True
         .WindowShowSearchBtn = True
         .ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAux1_2_3.rpt"
         .StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
         .StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
         .StoredProcParam(2) = Trim(Me.cbocta.Text)
         .StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
         .StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
         .StoredProcParam(5) = Trim(Dtc_benef.Text) 'Trim(Me.txtbusca1)
'         .StoredProcParam(6) = Trim(DtcCodAux2.Text) 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2)
'         .StoredProcParam(7) = Trim(DtcCodAux3.Text)
         .StoredProcParam(8) = Trim(Me.txtax1)
         .StoredProcParam(9) = Trim(Me.Txtax2)
         .StoredProcParam(10) = Trim(Me.txtax3)
         .Formulas(2) = "nomaux1 = '" & Trim(Dtc_benefD.Text) & "'"    'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
         .Formulas(3) = "nomaux2 = '" & Trim(DtcDenomAux2.Text) & "'"   'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
         .Formulas(4) = "nomaux3 = '" & Trim(DtcDenomAux3.Text) & "'"    'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
         .Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
         .Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
         .Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
         '.Formulas(12) = "SIBs = " & SaldoIBs
         '.Formulas(13) = "SISus = " & SaldoISus
         '.Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
         '.Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
         iResult = .PrintReport
        Exit Sub
        End With
        End If
    End If
    If Chkaux1.Value = 1 Then
    '--- reportes financiadores en gral
      If Chkaux1.Value = 1 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
        If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
         With CryConv_Conv
            .Destination = crptToWindow
            .WindowState = crptMaximized
             .WindowShowPrintSetupBtn = True
             .WindowShowSearchBtn = True
             .WindowShowGroupTree = True
             .ReportFileName = App.Path & "\REPORTES\Contabilidad\Libro_Mayor_Aux\CryLibroConv_Conv.rpt"
             .StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
             .StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
             .StoredProcParam(2) = Trim(Me.cbocta.Text)
             .StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
             .StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
             .StoredProcParam(5) = Trim(Me.txtax1)
             .StoredProcParam(6) = Trim(Me.Txtax2)
             .StoredProcParam(7) = Trim(Me.txtax3)
             .StoredProcParam(8) = Trim(Dtc_benef.Text) '(Me.txtbusca1)
             .Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
             .Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
             .Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
             .Formulas(3) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
             .Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
             .Formulas(6) = "nomsubcta2 = '" & Trim(Me.lbsub2) & "'"
             .Formulas(7) = "organismo = '" & Trim(Dtc_benef.Text) & "'" '& Trim(Me.txtbusca1) & "'"
             .Formulas(11) = "SIBs = " & Val(SaldoIBs)
             .Formulas(12) = "SISus= " & Val(SaldoISus)
             .Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
             .Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
            iresesult = .PrintReport
            Exit Sub
         End With
      End If
     End If
      Select Case Trim(txtax1)
        Case "01"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Escriba un beneficiario", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "02"
'            If Me.cboCtaBancaria = "" Then
'              MsgBox "Seleccione una Cuenta Bancaria", vbExclamation + vbDefaultButton1
'              Exit Sub
'            End If
       Case "03"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Seleccione un Proyecto", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "05"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Seleccione un Bien o Servicio", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "08"
            If DtCOrg.Text = "" Then
              MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
        Case "09"
            If DtCDesConvenio.Text = "" Then
              MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
      End Select
    End If
    '**************
    If Chkaux1.Value = 2 Then
      Select Case Trim(txtax1)
        Case "01"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Escriba un beneficiario", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "02"
'            If Me.cboCtaBancaria = "" Then
'              MsgBox "Seleccione una Cuenta Bancaria", vbExclamation + vbDefaultButton1
'              Exit Sub
'            End If
        Case "03"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Seleccione un Proyecto", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "05"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Seleccione un Bien o Servicio", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "08"
            If DtCOrg.Text = "" Then
              MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
        Case "09"
            If DtCDesConvenio.Text = "" Then
              MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
      End Select
    End If
    '*********
    If Chkaux3.Value = 1 Then
      Select Case Trim(txtax1)
        Case "01"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Escriba un beneficiario", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "02"
'            If Me.cboCtaBancaria = "" Then
'              MsgBox "Seleccione una Cuenta Bancaria", vbExclamation + vbDefaultButton1
'              Exit Sub
'            End If
        Case "03"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Seleccione un Proyecto", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "05"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Seleccione un Bien o Servicio", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "08"
            If DtCOrg.Text = "" Then
              MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
        Case "09"
            If DtCDesConvenio.Text = "" Then
              MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
      End Select
    End If
'***************
  If Chkaux1.Value = 1 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
    Select Case Trim(txtax1.Text)
        Case "01"
             If Chkaux2.Value = 0 And Trim(Txtax2) = "09" Then
                reporteBeneficiario_COnvenios
             Else
                reporteBeneficiario  'procedimiento para reporte con beneficiarios
             End If
        Case "03"
             If Chkaux2.Value = 0 And Trim(Txtax2) = "09" Then
                reporteBeneficiario_COnvenios
             Else
                reporteBeneficiario  'procedimiento para reporte con beneficiarios
             End If
        Case "02"
             ReporteCtaBancaria
        Case "05"
             If Chkaux2.Value = 0 And Trim(Txtax2) = "09" Then
                reporteBeneficiario_COnvenios
             Else
                reporteBeneficiario  'procedimiento para reporte con beneficiarios
             End If
        Case "08"
             ' procedimiento para reporte con organismos
             txtbusca1 = DtCOrg.Text
             reporteBeneficiario
        Case "09"
             reporteconvenio
    End Select
  End If
  If Chkaux1.Value = 0 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
    Select Case Trim(Txtax2)
      Case "01"
        ReporteOrg Trim(txtax1), nombenef
      Case "02"
'        ReporteOrg Trim(cboCtaBancaria), nomctabancaria
      Case "03"
        ReporteOrg Trim(txtax1), nombenef
      Case "05"
        ReporteOrg Trim(txtax1), nombenef
      Case "08"
        ReporteOrg Trim(DtCOrg.Text), Trim(DTCNomOrg.Text)
      Case "09"
         If (cbocta = "1121" And cbosubcta1 = "02") Or (cbocta = "2116" And cbosubcta1 = "04" And cbosubcta2 <> "03") Then
'          ReporteOrg Trim(DtcCodAux2.Text), Trim(DtcDenomAux2.Text)
         Else
          ReporteOrg Trim(DtCIdConvenio.Text), Trim(DtCDesConvenio.Text)
         End If
    End Select
  End If
  If Chkaux1.Value = 0 And Chkaux2.Value = 0 And Chkaux3.Value = 1 Then
  
  End If
  If Chkaux1.Value = 1 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
      'reporte benficiario con organismos
      If Trim(txtax1.Text) = "01" And Trim(Txtax2.Text) = "08" Then
        If rsbeneficiario.State = 1 Then rsbeneficiario.Close
          rsbeneficiario.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
          If rsbeneficiario.RecordCount <> 0 Then
            nombenef = rsbeneficiario!beneficiario_denominacion
          Else
            nombenef = ""
          End If
          ReporteAux1_2 Trim(txtbusca1.Text), Trim(DtCOrg.Text), Trim(txtax1.Text), Trim(Txtax2.Text), Trim(txtax3.Text), nombenef, Trim(DTCNomOrg.Text)
    End If
    'reporte benficiario con convenios
    If Trim(txtax1.Text) = "01" And Trim(Txtax2.Text) = "09" Then
        If rsbeneficiario.State = 1 Then rsbeneficiario.Close
        rsbeneficiario.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
        If rsbeneficiario.RecordCount <> 0 Then
           nombenef = rsbeneficiario!beneficiario_denominacion
        Else
           nombenef = ""
        End If
        ReporteAux1_2 Trim(txtbusca1.Text), Trim(DtCIdConvenio.Text), Trim(txtax1.Text), Trim(Txtax2.Text), Trim(txtax3.Text), nombenef, Trim(DtCDesConvenio.Text)
    End If
  End If
End If
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnCancelar_Click()
    'Dtereportes.Connection1.Close
    'Unload Me
    Frame1.Enabled = False
    Frame4.Visible = False
    frmabm.Visible = True
    FrmGraba.Visible = False
    BtnEnviar.Visible = False
    BtnGrabar.Visible = False
    
    Call Limpia_combos
            
    DataGrid3.AllowAddNew = False
    DataGrid3.AllowDelete = False
    DataGrid3.AllowUpdate = False
    DataGrid3.Enabled = True
End Sub

Private Sub DataGrid1_LostFocus()
    parametro = DtEreportes.rsbenef!beneficiario_codigo
End Sub

Private Sub Dtc_benef_Click(Area As Integer)
    Dtc_benefD.BoundText = Dtc_benef.BoundText
End Sub

Private Sub Dtc_benefD_Click(Area As Integer)
    Dtc_benef.BoundText = Dtc_benefD.BoundText
End Sub

Private Sub Dtc_CtaBco_Click(Area As Integer)
    Dtc_CtaBcoD.BoundText = Dtc_CtaBco.BoundText
End Sub

Private Sub Dtc_CtaBcoD_Click(Area As Integer)
    Dtc_CtaBco.BoundText = Dtc_CtaBcoD.BoundText
End Sub

Private Sub Dtc_Org_Click(Area As Integer)
    Dtc_OrgD.BoundText = Dtc_Org.BoundText
End Sub

Private Sub Dtc_OrgD_Click(Area As Integer)
    Dtc_Org.BoundText = Dtc_OrgD.BoundText
End Sub

Private Sub Dtc_Proy_Click(Area As Integer)
    Dtc_ProyD.BoundText = Dtc_Proy.BoundText
End Sub

Private Sub Dtc_ProyD_Click(Area As Integer)
    Dtc_Proy.BoundText = Dtc_ProyD.BoundText
End Sub

Private Sub Dtc_Uejec_Click(Area As Integer)
    Dtc_UejecD.BoundText = Dtc_Uejec.BoundText
End Sub

Private Sub Dtc_UejecD_Click(Area As Integer)
    Dtc_Uejec.BoundText = Dtc_UejecD.BoundText
End Sub

Private Sub DtcAux1_Click(Area As Integer)
    DtcCtaNom.BoundText = DtcAux1.BoundText
    cbocta1.BoundText = DtcAux1.BoundText
    dtccta.BoundText = DtcAux1.BoundText
    dtcsub1.BoundText = DtcAux1.BoundText
    dtcsub2.BoundText = DtcAux1.BoundText
    DtcAux2.BoundText = DtcAux1.BoundText
    DtcAux3.BoundText = DtcAux1.BoundText
End Sub

Private Sub DtcAux2_Click(Area As Integer)
    DtcCtaNom.BoundText = DtcAux2.BoundText
    cbocta1.BoundText = DtcAux2.BoundText
    dtccta.BoundText = DtcAux2.BoundText
    dtcsub1.BoundText = DtcAux2.BoundText
    dtcsub2.BoundText = DtcAux2.BoundText
    DtcAux1.BoundText = DtcAux2.BoundText
    DtcAux3.BoundText = DtcAux2.BoundText
End Sub

Private Sub DtcAux3_Click(Area As Integer)
    DtcCtaNom.BoundText = DtcAux3.BoundText
    cbocta1.BoundText = DtcAux3.BoundText
    dtccta.BoundText = DtcAux3.BoundText
    dtcsub1.BoundText = DtcAux3.BoundText
    dtcsub2.BoundText = DtcAux3.BoundText
    DtcAux1.BoundText = DtcAux3.BoundText
    DtcAux2.BoundText = DtcAux3.BoundText
End Sub

Private Sub DtcCta_Click(Area As Integer)
    DtcCtaNom.BoundText = dtccta.BoundText
    cbocta1.BoundText = dtccta.BoundText
    dtcsub1.BoundText = dtccta.BoundText
    dtcsub2.BoundText = dtccta.BoundText
    DtcAux1.BoundText = dtccta.BoundText
    DtcAux2.BoundText = dtccta.BoundText
    DtcAux3.BoundText = dtccta.BoundText
End Sub

Private Sub DtcCtaNom_Click(Area As Integer)
    cbocta1.BoundText = DtcCtaNom.BoundText
    dtccta.BoundText = DtcCtaNom.BoundText
    dtcsub1.BoundText = DtcCtaNom.BoundText
    dtcsub2.BoundText = DtcCtaNom.BoundText
    DtcAux1.BoundText = DtcCtaNom.BoundText
    DtcAux2.BoundText = DtcCtaNom.BoundText
    DtcAux3.BoundText = DtcCtaNom.BoundText
End Sub

Private Sub DataGrid3_Click()
'    MsgBox "---"
End Sub

Private Sub DtcCtaNom_LostFocus()
On Error GoTo err1
    cbocta.Text = dtccta.Text
    cbosubcta1.Text = dtcsub1.Text
    cbosubcta2.Text = dtcsub2.Text
    txtax1.Text = DtcAux1.Text
    Txtax2.Text = DtcAux2.Text
    txtax3.Text = DtcAux3.Text
    lbsub2.Caption = DtcCtaNom.Text
    
    '-- Cuenta
    rsplanctas.MoveFirst
    rsplanctas.Find "cuenta=" & "'" & Trim(cbocta.Text) & "'"
    Me.lblcuenta = rsplanctas!NombreCta
    If rscuentas.State = adStateOpen Then rscuentas.Close
    '-- SubCuenta1
    If rsnombresub1.State = adStateOpen Then rsnombresub1.Close
    rsnombresub1.Open "SELECT NombreCta FROM CC_Plan_Cuentas WHERE   (SubCta2 = '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & (Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
    Me.Lblsub1 = rsnombresub1!NombreCta
    '-- SubCuenta2
    If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
    rssubcuenta.Open "SELECT Cuenta, SubCta1, SubCta2, NombreCta, Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & Trim(Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
    
    Call carga_ctas

err1:
    If Err.Number = 7005 Then
'        DtgPlanCtas.Refresh
    End If
End Sub

Private Sub DtCDesConvenio_Change()
  DtCIdConvenio.BoundText = DtCDesConvenio.BoundText
End Sub


Private Sub DtcGrBien_Click(Area As Integer)
    DtcGrBienDes.BoundText = DtcGrBien.BoundText
End Sub

Private Sub DtcGrBienDes_Click(Area As Integer)
    DtcGrBien.BoundText = DtcGrBienDes.BoundText
End Sub

Private Sub DtCIdConvenio_Change()
  DtCDesConvenio.BoundText = DtCIdConvenio.BoundText
End Sub

Private Sub DTCNomOrg_Click(Area As Integer)
  DtCOrg.BoundText = DTCNomOrg.BoundText
End Sub

Private Sub DtcOrg_Click(Area As Integer)
  DTCNomOrg.BoundText = DtCOrg.BoundText
End Sub

Private Sub DtcProy_Click(Area As Integer)
    DtcProyDes.BoundText = DtcProy.BoundText
End Sub

Private Sub DtcProyDes_Click(Area As Integer)
    DtcProy.BoundText = DtcProyDes.BoundText
End Sub

Private Sub dtcsub1_Click(Area As Integer)
    DtcCtaNom.BoundText = dtcsub1.BoundText
    cbocta1.BoundText = dtcsub1.BoundText
    dtccta.BoundText = dtcsub1.BoundText
    dtcsub2.BoundText = dtcsub1.BoundText
    DtcAux1.BoundText = dtcsub1.BoundText
    DtcAux2.BoundText = dtcsub1.BoundText
    DtcAux3.BoundText = dtcsub1.BoundText
End Sub

Private Sub dtcsub2_Click(Area As Integer)
    DtcCtaNom.BoundText = dtcsub2.BoundText
    cbocta1.BoundText = dtcsub2.BoundText
    dtccta.BoundText = dtcsub2.BoundText
    dtcsub1.BoundText = dtcsub2.BoundText
    DtcAux1.BoundText = dtcsub2.BoundText
    DtcAux2.BoundText = dtcsub2.BoundText
    DtcAux3.BoundText = dtcsub2.BoundText
End Sub

Private Sub DtcTAux1_Click(Area As Integer)
    DtcTAux1Des.BoundText = DtcTAux1.BoundText
End Sub

Private Sub DtcTAux1_LostFocus()
  Select Case DtcTAux1.Text
    Case "01"
        'Beneficiarios
        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rs_beneficiario
        Set Ado_Benef.Recordset = rs_beneficiario
    
        Dtc_benefD.Visible = True
        Dtc_benef.Visible = True
        Fra_Benef.Visible = True
    Case "02"
        'Cuentas Bancarias
        If rs_cuentabancaria.State = 1 Then rs_cuentabancaria.Close
        sql2 = "SELECT cta_codigo, cta_descripcion From fc_cuenta_bancaria where estado_codigo= 'APR' order by cta_descripcion"
        rs_cuentabancaria.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DTGBanco.DataSource = rs_cuentabancaria
        Set Ado_CtaBanco.Recordset = rs_cuentabancaria
    
        Dtc_CtaBcoD.Visible = True
        Dtc_CtaBco.Visible = True
        Fra_CtaBco.Visible = True
    Case "03"
      'Proyectos (CAMBIAR TABLA)
        If rs_proyecto.State = 1 Then rs_proyecto.Close
        sql2 = "SELECT * From gc_edificaciones where estado_codigo= 'APR' order by edif_descripcion"
        rs_proyecto.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_proyecto
        Set Ado_Proyecto.Recordset = rs_proyecto
    
        Dtc_ProyD.Visible = True
        Dtc_Proy.Visible = True
        Fra_Proy.Visible = True
    Case "04"
      'Unidad Ejecutora
        If rs_UnidadEjecutora.State = 1 Then rs_UnidadEjecutora.Close
        sql2 = "SELECT * From gc_unidad_ejecutora where estado_codigo= 'APR' order by unidad_descripcion"
        rs_UnidadEjecutora.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_UnidadEjecutora
        Set Ado_Departamento.Recordset = rs_UnidadEjecutora
    
        Dtc_UejecD.Visible = True
        Dtc_Uejec.Visible = True
        Fra_UEjec.Visible = True
    Case "05"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT * From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
    Case "06"
      'Departamentos del Pais
        If rs_departamento.State = 1 Then rs_departamento.Close
        sql2 = "SELECT * From gc_departamento where estado_codigo= 'APR' order by depto_descripcion"
        rs_departamento.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_Departamento
        Set Ado_Departamento.Recordset = rs_departamento
    
        Dtc_DeptoD.Visible = True
        Dtc_Depto.Visible = True
        Fra_Depto.Visible = True
    Case "07"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
    Case "08"
      'Beneficiarios
        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rs_beneficiario
        Set Ado_Benef.Recordset = rs_beneficiario
    
        Dtc_benefD.Visible = True
        Dtc_benef.Visible = True
        Fra_Benef.Visible = True
    Case "09"
      'Beneficiarios
        If rs_Organismo.State = 1 Then rs_Organismo.Close
        sql2 = "SELECT * From fc_organismo_financiamiento where estado_codigo= 'APR' order by org_descripcion"
        rs_Organismo.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_Organismo
        Set AdodcOrganismo.Recordset = rs_Organismo
    
        Dtc_OrgD.Visible = True
        Dtc_Org.Visible = True
        Fra_Org.Visible = True
    Case "10"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
  End Select

End Sub

Private Sub DtcTAux1Des_Click(Area As Integer)
    DtcTAux1.BoundText = DtcTAux1Des.BoundText
End Sub

Private Sub DtcTAux1Des_LostFocus()
  Select Case DtcTAux1.Text
    Case "01"
        'Beneficiarios
        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rs_beneficiario
        Set Ado_Benef.Recordset = rs_beneficiario
    
        Dtc_benefD.Visible = True
        Dtc_benef.Visible = True
        Fra_Benef.Visible = True
    Case "02"
        'Cuentas Bancarias
        If rs_cuentabancaria.State = 1 Then rs_cuentabancaria.Close
        sql2 = "SELECT cta_codigo, cta_descripcion From fc_cuenta_bancaria where estado_codigo= 'APR' order by cta_descripcion"
        rs_cuentabancaria.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DTGBanco.DataSource = rs_cuentabancaria
        Set Ado_CtaBanco.Recordset = rs_cuentabancaria
    
        Dtc_CtaBcoD.Visible = True
        Dtc_CtaBco.Visible = True
        Fra_CtaBco.Visible = True
    Case "03"
      'Proyectos (CAMBIAR TABLA)
        If rs_proyecto.State = 1 Then rs_proyecto.Close
        sql2 = "SELECT * From gc_edificaciones where estado_codigo= 'APR' order by edif_descripcion"
        rs_proyecto.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_proyecto
        Set Ado_Proyecto.Recordset = rs_proyecto
    
        Dtc_ProyD.Visible = True
        Dtc_Proy.Visible = True
        Fra_Proy.Visible = True
    Case "04"
      'Unidad Ejecutora
        If rs_UnidadEjecutora.State = 1 Then rs_UnidadEjecutora.Close
        sql2 = "SELECT * From gc_unidad_ejecutora where estado_codigo= 'APR' order by unidad_descripcion"
        rs_UnidadEjecutora.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_UnidadEjecutora
        Set Ado_Departamento.Recordset = rs_UnidadEjecutora
    
        Dtc_UejecD.Visible = True
        Dtc_Uejec.Visible = True
        Fra_UEjec.Visible = True
    Case "05"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT * From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
    Case "06"
      'Departamentos del Pais
        If rs_departamento.State = 1 Then rs_departamento.Close
        sql2 = "SELECT * From gc_departamento where estado_codigo= 'APR' order by depto_descripcion"
        rs_departamento.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_Departamento
        Set Ado_Departamento.Recordset = rs_departamento
    
        Dtc_DeptoD.Visible = True
        Dtc_Depto.Visible = True
        Fra_Depto.Visible = True
    Case "07"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
    Case "08"
      'Beneficiarios
        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rs_beneficiario
        Set Ado_Benef.Recordset = rs_beneficiario
    
        Dtc_benefD.Visible = True
        Dtc_benef.Visible = True
        Fra_Benef.Visible = True
    Case "09"
      'Beneficiarios
        If rs_Organismo.State = 1 Then rs_Organismo.Close
        sql2 = "SELECT * From fc_organismo_financiamiento where estado_codigo= 'APR' order by org_descripcion"
        rs_Organismo.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_Organismo
        Set AdodcOrganismo.Recordset = rs_Organismo
    
        Dtc_OrgD.Visible = True
        Dtc_Org.Visible = True
        Fra_Org.Visible = True
    Case "10"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
  End Select

End Sub

Private Sub DtcTAux2_Click(Area As Integer)
  Select Case DtcTAux2.Text
    Case "01"
        'Beneficiarios
        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rs_beneficiario
        Set Ado_Benef.Recordset = rs_beneficiario
    
        Dtc_benefD.Visible = True
        Dtc_benef.Visible = True
        Fra_Benef.Visible = True
    Case "02"
        'Cuentas Bancarias
        If rs_cuentabancaria.State = 1 Then rs_cuentabancaria.Close
        sql2 = "SELECT cta_codigo, cta_descripcion From fc_cuenta_bancaria where estado_codigo= 'APR' order by cta_descripcion"
        rs_cuentabancaria.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DTGBanco.DataSource = rs_cuentabancaria
        Set Ado_CtaBanco.Recordset = rs_cuentabancaria
    
        Dtc_CtaBcoD.Visible = True
        Dtc_CtaBco.Visible = True
        Fra_CtaBco.Visible = True
    Case "03"
      'Proyectos (CAMBIAR TABLA)
        If rs_proyecto.State = 1 Then rs_proyecto.Close
        sql2 = "SELECT * From gc_edificaciones where estado_codigo= 'APR' order by edif_descripcion"
        rs_proyecto.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_proyecto
        Set Ado_Proyecto.Recordset = rs_proyecto
    
        Dtc_ProyD.Visible = True
        Dtc_Proy.Visible = True
        Fra_Proy.Visible = True
    Case "04"
      'Unidad Ejecutora
        If rs_UnidadEjecutora.State = 1 Then rs_UnidadEjecutora.Close
        sql2 = "SELECT * From gc_unidad_ejecutora where estado_codigo= 'APR' order by unidad_descripcion"
        rs_UnidadEjecutora.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_UnidadEjecutora
        Set Ado_Departamento.Recordset = rs_UnidadEjecutora
    
        Dtc_UejecD.Visible = True
        Dtc_Uejec.Visible = True
        Fra_UEjec.Visible = True
    Case "05"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT * From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
    Case "06"
      'Departamentos del Pais
        If rs_departamento.State = 1 Then rs_departamento.Close
        sql2 = "SELECT * From gc_departamento where estado_codigo= 'APR' order by depto_descripcion"
        rs_departamento.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_Departamento
        Set Ado_Departamento.Recordset = rs_departamento
    
        Dtc_DeptoD.Visible = True
        Dtc_Depto.Visible = True
        Fra_Depto.Visible = True
    Case "07"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
    Case "08"
      'Beneficiarios
        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rs_beneficiario
        Set Ado_Benef.Recordset = rs_beneficiario
    
        Dtc_benefD.Visible = True
        Dtc_benef.Visible = True
        Fra_Benef.Visible = True
    Case "09"
      'Beneficiarios
        If rs_Organismo.State = 1 Then rs_Organismo.Close
        sql2 = "SELECT * From fc_organismo_financiamiento where estado_codigo= 'APR' order by org_descripcion"
        rs_Organismo.Open sql2, db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rs_Organismo
        Set AdodcOrganismo.Recordset = rs_Organismo
    
        Dtc_OrgD.Visible = True
        Dtc_Org.Visible = True
        Fra_Org.Visible = True
    Case "10"
      'Beneficiarios
'        If rs_beneficiario.State = 1 Then rs_beneficiario.Close
'        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
'        rs_beneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rs_beneficiario
'        Set Ado_Benef.Recordset = rs_beneficiario
'
'        Dtc_benefD.Visible = True
'        Dtc_benef.Visible = True
'        Fra_Benef.Visible = True
  End Select

End Sub

Private Sub DtcTAux2Des_Click(Area As Integer)
    DtcTAux2.BoundText = DtcTAux2Des.BoundText
End Sub

Private Sub DTGBanco_Click()
'Me.txtbusca1.Text = Me.DTGBanco.Columns(0).Value
   On Error GoTo error3
'    Me.cboCtaBancaria.Text = Me.DTGBanco.Columns(0).Value
error3:
    If Err.Number = 7005 Then
        MsgBox "No existen datos", vbCritical + vbDefaultButton1
        Exit Sub
    End If
    
End Sub

Private Sub DtGbenef_Click()
    On Error GoTo err1
        Me.txtbusca1.Text = Me.DtGbenef.Columns(0)
        If txtax1 = "01" Then
            LblNom_Aux1 = DtGbenef.Columns(1)
        End If
        If txtax1 = "02" Then
            LblNom_Aux2 = DtGbenef.Columns(1)
        End If
'        If txtax1 = "03" Then
'            LblNom_Aux3 = DtGbenef.Columns(1)
'        End If
'        If txtax1 = "05" Then
'            LblNom_Aux1 = DtGbenef.Columns(1)
'        End If
err1:
    If Err.Number = 7005 Then
    DtGbenef.Refresh
    End If

End Sub

Private Sub DtgPlanCtas_Click()
On Error GoTo err1
    cbocta.Text = DtgPlanCtas.Columns(0)
    cbosubcta1.Text = DtgPlanCtas.Columns(1)
    cbosubcta2.Text = DtgPlanCtas.Columns(2)
    txtax1.Text = DtgPlanCtas.Columns(3)
    Txtax2.Text = DtgPlanCtas.Columns(4)
    txtax3.Text = DtgPlanCtas.Columns(5)
    lbsub2.Caption = DtgPlanCtas.Columns(6)
    
    '-- Cuenta
    rsplanctas.MoveFirst
    rsplanctas.Find "cuenta=" & "'" & Trim(cbocta.Text) & "'"
    Me.lblcuenta = rsplanctas!NombreCta
    If rscuentas.State = adStateOpen Then rscuentas.Close
    '-- SubCuenta1
    If rsnombresub1.State = adStateOpen Then rsnombresub1.Close
    rsnombresub1.Open "SELECT NombreCta FROM CC_Plan_Cuentas WHERE   (SubCta2 = '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & (Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
    Me.Lblsub1 = rsnombresub1!NombreCta
    '-- SubCuenta2
    If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
    rssubcuenta.Open "SELECT Cuenta, SubCta1, SubCta2, NombreCta, Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & Trim(Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
    
    Call carga_ctas

err1:
    If Err.Number = 7005 Then
    DtgPlanCtas.Refresh
    End If
End Sub

Private Sub DTPfin_Validate(Cancel As Boolean)
If DTPfin.Value < DTPinicio.Value Then
    MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
    DTPfin.SetFocus
End If
End Sub

Private Sub DTPinicio_LostFocus()
Me.DTPfin.MinDate = Me.DTPinicio.Value
End Sub
Private Sub DTPinicio_Validate(Cancel As Boolean)
If DTPinicio.Value > DTPfin.Value Then
    MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
    DTPfin.SetFocus
End If
End Sub

Private Sub Form_Load()
'Me.BtnGrabar.Enabled = True
On Error GoTo error_conec
    Set rsOrganismo = New ADODB.Recordset
    Set rsplanctas = New ADODB.Recordset
    Set rscuentas = New ADODB.Recordset
    Set rsnombresub1 = New ADODB.Recordset
    Set rssubcuenta = New ADODB.Recordset
    Set rscta_bancaria = New ADODB.Recordset
    Set rsbeneficiario = New ADODB.Recordset
    Set rsConvenio = New ADODB.Recordset
    Set rsGrupoBien = New ADODB.Recordset
    Set rsProy = New ADODB.Recordset
    '-----------
'    With rsConvenio
'        If .State = 1 Then .Close
'        .CursorLocation = adUseClient
'        sql1 = "SELECT Codigo_Convenio, Denominacion_Convenio," & _
'            " org_codigo From fc_convenios"
'        .Open sql1, db, adOpenKeyset, adLockReadOnly
'        Set AdoConvenio.Recordset = rsConvenio
'    End With
    '-----------
    If rsplanctas.State = 1 Then rsplanctas.Close
    rsplanctas.Open "SELECT Cuenta, NombreCta FROM CC_Plan_Cuentas WHERE SubCta1 = '00' AND SubCta2 = '00' order by Cuenta", db, adOpenKeyset, adLockReadOnly
    rsplanctas.MoveFirst
    Do While Not rsplanctas.EOF
        Me.cbocta.AddItem rsplanctas!cuenta
        rsplanctas.MoveNext
    Loop
    'Beneficiarios
    If rsbeneficiario.State = 1 Then rsbeneficiario.Close
    sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where estado_codigo= 'APR' order by beneficiario_denominacion"
    rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsbeneficiario
    Set Ado_Benef.Recordset = rsbeneficiario
    
    Me.cbocta.Text = Me.cbocta.List(0)
   ' Me.DTPfin.MaxDate = CDate(Date)
   ' Me.DTPinicio.MaxDate = CDate(Date)
    Me.DTPfin.Value = Date
    Me.DTPinicio.Value = CDate("01/01/2016")
    'Me.DTPinicio.MinDate = CDate("01/01/2001")
    Me.DTPfin.MinDate = Date    'CDate(Me.DTPinicio.Value)
'    Me.PRB.Visible = False
    '----------
    If rsOrganismo.State = 1 Then rsOrganismo.Close
    rsOrganismo.CursorLocation = adUseClient
    rsOrganismo.Open "SELECT Org_codigo, Org_descripcion" & _
                      " FROM fc_organismo_financiamiento order by org_Codigo", db, adOpenKeyset, adLockReadOnly
    'MsgBox rsorganismo.RecordCount
    Set AdodcOrganismo.Recordset = rsOrganismo
    'Print AdodcOrganismo.Recordset.RecordCount
    AdodcOrganismo.Refresh '
    'PARTIDAS
    If rsGrupoBien.State = 1 Then rsGrupoBien.Close
    rsGrupoBien.Open "SELECT grupo_codigo, par_descripcion AS DESCRIPCION, PAR_CODIGO  From fc_partida_gasto WHERE estado_codigo = 'APR' order by par_descripcion", db, adOpenKeyset, adLockReadOnly
    Set AdoGrBien.Recordset = rsGrupoBien
    AdoGrBien.Refresh
    'TIPO AUXILIARES
    If rs_tipo_auxiliar.State = 1 Then rs_tipo_auxiliar.Close
    rs_tipo_auxiliar.Open "SELECT * From cc_tipo_auxiliar WHERE estado_codigo = 'APR' order by descripcion", db, adOpenKeyset, adLockReadOnly
    Set Ado_TipoAuxiliar.Recordset = rs_tipo_auxiliar
    Ado_TipoAuxiliar.Refresh
    '
    If rsProy.State = 1 Then rsProy.Close
    rsProy.Open "SELECT pro_codigo, Pro_programa, Pro_proyecto, Pro_actividad, pro_descripcion From fc_estructura_programatica WHERE estado_codigo = 'APR' AND pro_nivel > 1 order by pro_descripcion", db, adOpenKeyset, adLockReadOnly
    Set AdoProy.Recordset = rsProy
    AdoProy.Refresh
    
    Set DtCOrg.RowSource = AdodcOrganismo.Recordset
    DtCOrg.ListField = "org_codigo"
    DtCOrg.BoundColumn = "org_codigo" 'AdodcOrganismo.Recordset!org_codigo
'DtCOrg.ReFill
    'Set DTCNomOrg.DataSource = AdodcOrganismo.Recordset
    Set DTCNomOrg.RowSource = AdodcOrganismo.Recordset
    DTCNomOrg.ListField = "Org_descripcion"
    DTCNomOrg.BoundColumn = "org_codigo"
    Me.DTCNomOrg.Visible = False
    Me.DtCOrg.Visible = False
    If Not rsOrganismo.EOF And Not rsOrganismo.BOF Then
      rsOrganismo.MoveFirst
      DtCOrg.Text = rsOrganismo!org_codigo
      DtcOrg_Click (0)
    End If
'    If Not rsConvenio.EOF And Not rsConvenio.BOF Then
'      rsConvenio.MoveFirst
'      DtCIdConvenio.Text = rsConvenio!codigo_convenio
'      DtCIdConvenio_Change
'    End If
    '----------1121 y 2116
    'Dim sql1 As String
    sql31 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario " & _
          "WHERE (tipoben_codigo > '20) and estado_codigo= 'APR' ORDER BY beneficiario_denominacion"
    'Set Dtc_benef.RowSource = db.Execute(sql31, , commantext)
'    Set Dtc_benefD.RowSource = db.Execute(sql31, , commantext)
'    sql32 = "SELECT Codigo_Convenio, Denominacion_Convenio " & _
'            "From fc_convenios ORDER BY Denominacion_Convenio"
'    Set DtcCodAux2.RowSource = db.Execute(sql32, , commantext)
'    Set DtcDenomAux2.RowSource = db.Execute(sql32, , commantext)
'    Set DtcCodAux3.RowSource = db.Execute(sql32, , commantext)
'    Set DtcDenomAux3.RowSource = db.Execute(sql32, , commantext)
'    Dtc_benefD
 '   DtcDenomAux2
'-----------------
    cboUnidad.AddItem "Todas"
'    deCD.dbo_edGeneralSearching "SELECT cuenta FROM Co_BalanceApertura group by cuenta ORDER BY cuenta"
'    While Not deCD.rsdbo_edGeneralSearching.EOF
'        cboUnidad.AddItem deCD.rsdbo_edGeneralSearching!cuenta
'        deCD.rsdbo_edGeneralSearching.MoveNext
'    Wend
'    deCD.rsdbo_edGeneralSearching.Close
    
    strUnidad = "%"
    Set rstAo_solicitud1 = New ADODB.Recordset
    If rstAo_solicitud1.State = 1 Then rstAo_solicitud1.Close
    queryinicial2 = "SELECT * FROM Co_BalanceApertura"
    rstAo_solicitud1.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rstAo_solicitud1.Sort = "correl"
    Set adosolicitud1.Recordset = rstAo_solicitud1
    Label6(3) = adosolicitud1.Recordset.RecordCount
    CboStatus.AddItem "Todas"
'    deCD.dbo_edGeneralSearching "SELECT VERIFICADO FROM Co_BalanceApertura group by VERIFICADO ORDER BY VERIFICADO"
'    While Not deCD.rsdbo_edGeneralSearching.EOF
'        CboStatus.AddItem deCD.rsdbo_edGeneralSearching!Verificado
'        deCD.rsdbo_edGeneralSearching.MoveNext
'    Wend
'    deCD.rsdbo_edGeneralSearching.Close
    strUnidad = "%"
    
    Set rs_bien = New ADODB.Recordset
    If rs_bien.State = 1 Then rs_bien.Close
    'SqlBienes = "select * from CC_Plan_Cuentas WHERE PAR_CODIGO= '" & varpar & "' "
    SqlBienes = "select * from CC_Plan_Cuentas "
    rs_bien.Open SqlBienes, db, adOpenStatic, adLockReadOnly
    rs_bien.Sort = "Cuenta, SubCta1, SubCta2"
    TDBPlan.DataSource = rs_bien
    Set AdoPlan.Recordset = rs_bien
    
    Set rsPlanBusq = New ADODB.Recordset
    If rsPlanBusq.State = 1 Then rsPlanBusq.Close
    sql2 = "SELECT * From CC_Plan_Cuentas where mov = 'D' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3"
    rsPlanBusq.Open sql2, db, adOpenKeyset, adLockReadOnly
'    rsPlanBusq.Sort = "correl"
    Set DtgPlanCtas.DataSource = rsPlanBusq
    
    Set rsPlanCta1 = New ADODB.Recordset
    If rsPlanCta1.State = 1 Then rsPlanCta1.Close
    rsPlanCta1.Open "SELECT * From CC_Plan_Cuentas where nivel = '1' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3", db, adOpenKeyset, adLockReadOnly
'    While Not rsPlanCta1.EOF
'        cbocta1.AddItem rsPlanCta1!cuenta
'        cboctaNom.AddItem rsPlanCta1!NombreCta
'        rsPlanCta1.MoveNext
'    Wend
    Set AdoPlan1.Recordset = rsPlanCta1
    
    Set rsPlanCta2 = New ADODB.Recordset
    If rsPlanCta2.State = 1 Then rsPlanCta2.Close
    rsPlanCta2.Open "SELECT * From CC_Plan_Cuentas where nivel = '2' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3", db, adOpenKeyset, adLockReadOnly
'    While Not rsPlanCta2.EOF
'        cbocta2.AddItem rsPlanCta2!cuenta
'        cboctaNom2.AddItem rsPlanCta2!NombreCta
'        rsPlanCta2.MoveNext
'    Wend
    Set AdoPlan2.Recordset = rsPlanCta2
    
    Set rsPlanCta3 = New ADODB.Recordset
    If rsPlanCta3.State = 1 Then rsPlanCta3.Close
    rsPlanCta3.Open "SELECT * From CC_Plan_Cuentas where nivel = '3' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3", db, adOpenKeyset, adLockReadOnly
'    While Not rsPlanCta3.EOF
'        cbocta3.AddItem rsPlanCta3!cuenta
'        cboctaNom3.AddItem rsPlanCta3!NombreCta
'        rsPlanCta3.MoveNext
'    Wend
    Set AdoPlan3.Recordset = rsPlanCta3
    
    Fra_BuscaGral.Visible = False
    swgraba3 = 2
    
    Frame1.Enabled = False
    Frame4.Visible = False
    frmabm.Visible = True
    FrmGraba.Visible = False
    
    Call Limpia_combos
    
    DataGrid3.AllowAddNew = False
    DataGrid3.AllowDelete = False
    DataGrid3.AllowUpdate = False
    DataGrid3.Enabled = True
'-----------------
  '  Exit Sub
error_conec:
    If Err.Number = -2147220992 Then
      MsgBox "ERROR EN LA CONECCION, Revise su conección a la red", vbCritical + vbDefaultButton1, "Atencion"
      End
    End If

	Call SeguridadSet(Me)
End Sub

Public Sub Mayor000()
  Dim iResult As Integer
    Set commayor = New ADODB.Command ' para obtener los saldos
    With commayor
        .CommandType = adCmdStoredProc
        .CommandText = "SaldoLMayor"
        .Parameters.Append commayor.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
        .Parameters.Append commayor.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
        .Parameters.Append commayor.CreateParameter("cuenta", adVarChar, adParamInput, 5)
        .Parameters.Append commayor.CreateParameter("subcta1", adVarChar, adParamInput, 3)
        .Parameters.Append commayor.CreateParameter("subcta2", adVarChar, adParamInput, 3)
        .Parameters.Append commayor.CreateParameter("SIBs", adDouble, adParamOutput)
        .Parameters.Append commayor.CreateParameter("SISus", adDouble, adParamOutput)
        .Parameters("FFInicio") = Me.DTPinicio.Value
        .Parameters("FFFinal") = Me.DTPfin.Value
        .Parameters("cuenta") = Trim(Me.cbocta.Text)
        .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
        .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
        .ActiveConnection = db
        .Execute
        SaldoIBs = .Parameters("SIBs")
        SaldoISus = .Parameters("SISus")
    End With
        CryLMayor.Destination = crptToWindow
        CryLMayor.WindowState = crptMaximized
        CryLMayor.WindowShowPrintSetupBtn = True
        CryLMayor.WindowShowSearchBtn = True
        CryLMayor.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor\CryLMayor.rpt"
        CryLMayor.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
        CryLMayor.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
        CryLMayor.StoredProcParam(2) = Trim(Me.cbocta.Text)
        CryLMayor.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
        CryLMayor.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
        
        CryLMayor.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
        CryLMayor.Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
        CryLMayor.Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
        CryLMayor.Formulas(4) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
        CryLMayor.Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
        CryLMayor.Formulas(6) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
        CryLMayor.Formulas(9) = "SIBs = " & SaldoIBs
        CryLMayor.Formulas(10) = "SISus = " & SaldoISus
        CryLMayor.Formulas(11) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
        CryLMayor.Formulas(12) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = CryLMayor.PrintReport
        If iResult <> 0 Then
            MsgBox CryLMayor.LastErrorNumber & " : " & CryLMayor.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
End Sub

Private Sub TbrAvanzadas_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "Refrescar"
      OpcionRefrescar
    Case "Filtrar"
      OpcionFiltrar
    Case "Ordenar"
      If (CmbCampo.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      OpcionOrdenar OrdenarAsc
    Case "Buscar"
      If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      Select Case Button.Caption
'        Case "&Primero"
'          OpcionBuscarPrimero
'        Case "&Anterior"
'          OpcionBuscarAnterior
'        Case "&Siguiente"
'          OpcionBuscarSiguiente
      End Select
    Case "Salir"
      OpcionSalir
  End Select
End Sub

Private Sub OpcionOrdenar(Ascendente As Boolean)
Dim AuxCampo As String
On Error GoTo Que_Error
  If (CmbCampo.Text = "") Then
    MsgBox "Debe elegir el campo por el que quiere ordenar", vbInformation + vbOKOnly, "Error de procedimiento"
  Else
    Screen.MousePointer = vbHourglass
    'check for the use of the ctrl key for descending sort
    AuxCampo = NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))   ' NombreCampo(CmbCampo.Text)
    If Not Ascendente Then
      rsTablaAux.Sort = AuxCampo & " DESC"
    Else
      rsTablaAux.Sort = AuxCampo & " ASC"
    End If
    Screen.MousePointer = vbDefault
  End If
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Function NombreCampo(CampoBusca As String) As String
  If InStr(1, CampoBusca, " ") Then
    NombreCampo = "[" & CampoBusca & "]"
'    NombreCampo = rsTablaAux.Fields(DtgElige.Col).Name
  Else
    NombreCampo = CampoBusca
  End If
End Function

Private Sub OpcionRefrescar()
  On Error GoTo RefErr
    If rsPlanBusq.State = 1 Then rsPlanBusq.Close
    sql2 = "SELECT * From CC_Plan_Cuentas where mov = 'D' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3"
    rsPlanBusq.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set DtgPlanCtas.DataSource = rsPlanBusq
  Exit Sub
RefErr:
  MsgBox "Error:" & Err & " " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionFiltrar()
On Error GoTo Que_Error
 Select Case Me.CboCampoC
    Case "Codigo"
        Select Case Me.CboOperadorC
            Case "="
                 sql2 = "SELECT * From CC_Plan_Cuentas where  Cuenta ='" & Trim(TxtValorC) & "' and mov = 'D' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3"
            Case "como"
                 sql2 = "SELECT * From CC_Plan_Cuentas where  Cuenta like '" & Trim(Me.TxtValorC) & "'+'%' and mov = 'D'  order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3"
        End Select
    Case "Denominacion"
        Select Case Me.CboOperadorC
            Case "="
                sql2 = "SELECT * From CC_Plan_Cuentas where  NombreCta ='" & Trim(TxtValorC) & "' and mov = 'D' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3"
        Case "como"
                sql2 = "SELECT * From CC_Plan_Cuentas where  NombreCta like '%' + '" & Trim(Me.TxtValorC) & "'+'%' and mov = 'D' order by Cuenta, SubCta1, SubCta2, Aux1, Aux2, Aux3"
    End Select
 End Select
    If rsPlanBusq.State = 1 Then rsPlanBusq.Close
    rsPlanBusq.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set DtgPlanCtas.DataSource = rsPlanBusq
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionSalir()
    Fra_BusquedaC.Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Refrescar"
      OpcionRefrescar3
    Case "Filtrar"
      OpcionFiltrar3
    Case "Ordenar"
      If (CmbCampo.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      OpcionOrdenar OrdenarAsc
    Case "Buscar"
      If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      Select Case Button.Caption
'        Case "&Primero"
'          OpcionBuscarPrimero
'        Case "&Anterior"
'          OpcionBuscarAnterior
'        Case "&Siguiente"
'          OpcionBuscarSiguiente
      End Select
    Case "Salir"
      Fra_BuscaGral.Visible = False
  End Select
End Sub

Private Sub OpcionRefrescar3()
'REFRESCAR
  On Error GoTo RefErr
    If rstAo_solicitud1.State = 1 Then rstAo_solicitud1.Close
    sql2 = "SELECT * From Co_BalanceApertura where CUENTA <> '-' order by correl "
    rstAo_solicitud1.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set DataGrid3.DataSource = rstAo_solicitud1
    Set adosolicitud1.Recordset = rstAo_solicitud1
    Label6(3) = adosolicitud1.Recordset.RecordCount
  Exit Sub
RefErr:
  MsgBox "Error:" & Err & " " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionFiltrar3()
'FILTRO
    Set rstAo_solicitud1 = New ADODB.Recordset
    If rstAo_solicitud1.State = 1 Then rstAo_solicitud1.Close
    If (cboUnidad.Text = "Todas" And CboStatus = "Todas" And TxtCtaNom = "") Or (cboUnidad.Text = "" And CboStatus = "" And TxtCtaNom = "") Then
        queryinicial2 = "SELECT * FROM Co_BalanceApertura order by correl"
    Else
        If cboUnidad.Text = "Todas" And CboStatus = "" And TxtCtaNom = "" Then
           queryinicial2 = "SELECT * FROM Co_BalanceApertura WHERE status = '" & CboStatus & "'  order by correl"
        Else
           If CboStatus = "Todas" And cboUnidad.Text = "" And TxtCtaNom = "" Then
            queryinicial2 = "SELECT * FROM Co_BalanceApertura WHERE CUENTA = '" & cboUnidad & "'  order by correl"
           Else
            If cboUnidad.Text = "" And CboStatus = "" And TxtCtaNom <> "" Then
                queryinicial2 = "SELECT * FROM Co_BalanceApertura WHERE NombreCta like '%' + '" & Trim(TxtCtaNom) & "' + '%'   order by correl"
            Else
                If cboUnidad.Text <> "" And CboStatus = "" And TxtCtaNom = "" Then
                    queryinicial2 = "select * from Co_BalanceApertura WHERE CUENTA = '" & cboUnidad.Text & "'  order by correl"
                Else
                    queryinicial2 = "select * from Co_BalanceApertura WHERE CUENTA = '" & cboUnidad.Text & "' AND status = '" & CboStatus & "' and ('%' + NombreCta like '" & TxtCtaNom & "' + '%' )   order by correl"
                End If
            End If
           End If
        End If
    End If
    rstAo_solicitud1.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    Set adosolicitud1.Recordset = rstAo_solicitud1
    DataGrid3.DataSource = adosolicitud1.Recordset
    rstAo_solicitud1.Requery
    Label6(3) = adosolicitud1.Recordset.RecordCount
End Sub

Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "Refrescar"
      OpcionRefrescar2
    Case "Filtrar"
      OpcionFiltrar2
    Case "Ordenar"
      If (CmbCampo.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      OpcionOrdenar OrdenarAsc
    Case "Buscar"
      If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      Select Case Button.Caption
'        Case "&Primero"
'          OpcionBuscarPrimero
'        Case "&Anterior"
'          OpcionBuscarAnterior
'        Case "&Siguiente"
'          OpcionBuscarSiguiente
      End Select
    Case "Salir"
      OpcionSalir2
  End Select
End Sub

Private Sub OpcionRefrescar2()
  On Error GoTo RefErr
    If rsbeneficiario.State = 1 Then rsbeneficiario.Close
    sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario WHERE estado_codigo = 'APR' order by beneficiario_denominacion"
    rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsbeneficiario
    Set Ado_Benef.Recordset = rsbeneficiario
  Exit Sub
RefErr:
  MsgBox "Error:" & Err & " " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionFiltrar2()
On Error GoTo Que_Error
 Select Case Me.CboCampo
    Case "Codigo"
        Select Case Me.CboOperador
            Case "="
                 sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where beneficiario_codigo ='" & Trim(Me.TxtValor) & "' AND tipoben_codigo < '20' order by beneficiario_codigo"
            Case "como"
                 sql2 = " select beneficiario_codigo, beneficiario_denominacion from  gc_beneficiario WHERE beneficiario_codigo like '" & Trim(Me.TxtValor) & "'+'%' AND tipoben_codigo < '20' order by beneficiario_codigo"
        End Select
    Case "Denominacion"
        Select Case Me.CboOperador
            Case "="
                sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario where  beneficiario_denominacion ='" & Trim(Me.TxtValor) & "' AND tipoben_codigo < '20' order by beneficiario_denominacion"
        Case "como"
                sql2 = " select beneficiario_codigo, beneficiario_denominacion from  gc_beneficiario WHERE beneficiario_denominacion like '" & Trim(Me.TxtValor) & "'+'%'  AND tipoben_codigo < '20' order by beneficiario_denominacion"
    End Select
 End Select
    If rsbeneficiario.State = 1 Then rsbeneficiario.Close
    rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsbeneficiario
    Set Ado_Benef.Recordset = rsbeneficiario
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionSalir2()
    Fra_Busqueda.Visible = False
End Sub

Private Sub txtbusca1_LostFocus()
    Me.BtnImprimir.Enabled = True
End Sub

Public Sub existecta(cuenta As String, subcta1 As String, subcta2 As String)
    Dim rsexiste As ADODB.Recordset
    Set rsexiste = New ADODB.Recordset
    If rsexiste.State = 1 Then rsexiste.Close
    rsexiste.CursorLocation = adUseClient
    rsexiste.Open "SELECT * From CC_Plan_Cuentas WHERE (Cuenta='" & Trim(cuenta) & "') AND (SubCta1='" & Trim(subcta1) & "') AND (SubCta2='" & Trim(subcta2) & "')", db, adOpenKeyset, adLockReadOnly
    If rsexiste.RecordCount <> 0 Then
            If rsexiste!mov = "T" Then
                MsgBox "La cuenta es de título"
                lcta = "N"
            Else
                lcta = "S"
            End If
    Else
        MsgBox "La cuenta no existe"
        cbocta.SetFocus
        lcta = "N"
    End If
End Sub
Public Sub ReporteOrg(AUX2 As String, NOMAUX2 As String)
  Dim iResult As Integer
    Set comORG = New ADODB.Command ' para obtener los saldos
    With comORG
        .CommandType = adCmdStoredProc
        .CommandText = "SaldoOrganismo"
        .Parameters.Append comORG.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("cuenta", adVarChar, adParamInput, 5)
        .Parameters.Append comORG.CreateParameter("subcta1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("subcta2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("organismo", adVarChar, adParamInput, 85)
        .Parameters.Append comORG.CreateParameter("aux1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux3", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("SIBs", adDouble, adParamOutput)
        .Parameters.Append comORG.CreateParameter("SISus", adDouble, adParamOutput)
        .Parameters("FFInicio") = Me.DTPinicio.Value
        .Parameters("FFFinal") = Me.DTPfin.Value
        .Parameters("cuenta") = Trim(Me.cbocta.Text)
        .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
        .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
        .Parameters("organismo") = AUX2  'Trim(DtcOrg.Text) 'Trim(Me.cbosubcta2.Text)
        .Parameters("aux1") = Trim(Me.txtax1.Text)
        .Parameters("aux2") = Trim(Me.Txtax2.Text)
        .Parameters("aux3") = Trim(Me.txtax3.Text)
        .ActiveConnection = db
        .Execute
        SaldoIBs = .Parameters("SIBs")
        SaldoISus = .Parameters("SISus")
    End With
       CryOrg.Destination = crptToWindow
       CryOrg.WindowState = crptMaximized
       CryOrg.WindowShowPrintSetupBtn = True
       CryOrg.WindowShowSearchBtn = True
       CryOrg.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxORG.rpt"
       CryOrg.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(2) = Trim(Me.cbocta.Text)
       CryOrg.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
       CryOrg.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
       CryOrg.StoredProcParam(5) = AUX2 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2.Text)
       CryOrg.StoredProcParam(6) = Trim(Me.txtax1.Text)
       CryOrg.StoredProcParam(7) = Trim(Me.Txtax2.Text)
       CryOrg.StoredProcParam(8) = Trim(Me.txtax3.Text)
      
       CryOrg.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
       CryOrg.Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
       CryOrg.Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
       CryOrg.Formulas(3) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
       CryOrg.Formulas(4) = "nomorg = '" & NOMAUX2 & "'" 'Trim(DTCNomOrg.Text) & "'"
       CryOrg.Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
       CryOrg.Formulas(6) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
       CryOrg.Formulas(7) = "organismo ='" & AUX2 & "'" '& Trim(DtCOrg.Text) & "'"
       CryOrg.Formulas(10) = "SIBs = " & SaldoIBs
       CryOrg.Formulas(11) = "SISus = " & SaldoISus
       CryOrg.Formulas(12) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
       CryOrg.Formulas(13) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = CryOrg.PrintReport
        If iResult <> 0 Then
            MsgBox CryOrg.LastErrorNumber & " : " & CryOrg.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

End Sub
Public Sub ReporteAux1_2(busca1 As String, busca2 As String, ax1 As String, ax2 As String, ax3 As String, nombusca1 As String, nombusca2 As String)
   If rsbeneficiario.State = 1 Then rsbeneficiario.Close
   rsbeneficiario.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
   If rsbeneficiario.RecordCount <> 0 Then
      nombenef = rsbeneficiario!beneficiario_denominacion
   Else
      nombenef = ""
   End If
 Dim iResult As Integer
   Set comAux12 = New ADODB.Command
   With comAux12
       .CommandType = adCmdStoredProc
       .CommandText = "Saldos_Aux1_2"
       .Parameters.Append comAux12.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
       .Parameters.Append comAux12.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
       .Parameters.Append comAux12.CreateParameter("cuenta", adVarChar, adParamInput, 5)
       .Parameters.Append comAux12.CreateParameter("subcta1", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("subcta2", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("busca1", adVarChar, adParamInput, 15)
       .Parameters.Append comAux12.CreateParameter("busca2", adVarChar, adParamInput, 15)
       .Parameters.Append comAux12.CreateParameter("aux1", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("aux2", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("aux3", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("SIBs", adDouble, adParamOutput)
       .Parameters.Append comAux12.CreateParameter("SISus", adDouble, adParamOutput)
       .Parameters("FFInicio") = Me.DTPinicio.Value
       .Parameters("FFFinal") = Me.DTPfin.Value
       .Parameters("cuenta") = Trim(Me.cbocta.Text)
       .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
       .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
       .Parameters("busca1") = busca1 'Trim(Me.txtbusca1.Text)
       .Parameters("busca2") = busca2 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2.Text)
       .Parameters("aux1") = ax1 'Trim(Me.txtax1)
       .Parameters("aux2") = ax2 'Trim(Me.txtax1)"00"
       .Parameters("aux3") = ax3 'Trim(Me.txtax1)"00"
       .ActiveConnection = db
       .Execute
       SaldoIBs = .Parameters("SIBs")
       SaldoISus = .Parameters("SISus")
   End With
   
   'Me.ProgressBar1.Visible = True
   'Me.ProgressBar1.Value = 0
'   CRyAux12
       CRyAux12.Destination = crptToWindow
       CRyAux12.WindowState = crptMaximized
       CRyAux12.WindowShowPrintSetupBtn = True
       CRyAux12.WindowShowSearchBtn = True
       CRyAux12.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAux1_2.rpt"
       CRyAux12.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
       CRyAux12.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
       CRyAux12.StoredProcParam(2) = Trim(Me.cbocta.Text)
       CRyAux12.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
       CRyAux12.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
       CRyAux12.StoredProcParam(5) = busca1 'Trim(Me.txtbusca1)
       CRyAux12.StoredProcParam(6) = busca2 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2)
       CRyAux12.StoredProcParam(7) = ax1 'Trim(Me.txtax1)
       CRyAux12.StoredProcParam(8) = ax2 'Trim(Me.Txtax2)
       CRyAux12.StoredProcParam(9) = ax3 'Trim(Me.txtax3)
       
       CRyAux12.Formulas(0) = "aux2 = '" & busca2 & "'"    'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
       CRyAux12.Formulas(1) = "benef = '" & busca1 & "'" '& Trim(Me.txtbusca1) & "'"
       CRyAux12.Formulas(2) = "cta = '" & Trim(Me.cbocta.Text) & "'"
       CRyAux12.Formulas(3) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
       CRyAux12.Formulas(4) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
       CRyAux12.Formulas(5) = "nomaux2 = '" & nombusca2 & "'" 'Trim(DTCNomOrg.Text) & "'"
       CRyAux12.Formulas(6) = "nombenef = '" & nombusca1 & "'" 'nombenef & "'"
       CRyAux12.Formulas(7) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
       CRyAux12.Formulas(8) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
       CRyAux12.Formulas(9) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
       CRyAux12.Formulas(12) = "SIBs = " & SaldoIBs
       CRyAux12.Formulas(13) = "SISus = " & SaldoISus
       CRyAux12.Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
       CRyAux12.Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
       iResult = CRyAux12.PrintReport
'*****fin aux1
'Exit Sub
End Sub
Public Sub reporteBeneficiario()
If rsbeneficiario.State = 1 Then rsbeneficiario.Close
rsbeneficiario.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
If rsbeneficiario.RecordCount <> 0 Then
  nombenef = rsbeneficiario!beneficiario_denominacion
Else
  nombenef = ""
End If
            Dim iResult As Integer
            Set combenef = New ADODB.Command
            With combenef
                .CommandType = adCmdStoredProc
                .CommandText = "SaldoBenef"
                .Parameters.Append combenef.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
                .Parameters.Append combenef.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
                .Parameters.Append combenef.CreateParameter("cuenta", adVarChar, adParamInput, 5)
                .Parameters.Append combenef.CreateParameter("subcta1", adVarChar, adParamInput, 3)
                .Parameters.Append combenef.CreateParameter("subcta2", adVarChar, adParamInput, 3)
                .Parameters.Append combenef.CreateParameter("beneficiario", adVarChar, adParamInput, 15)
                .Parameters.Append combenef.CreateParameter("aux1", adVarChar, adParamInput, 3)
                .Parameters.Append combenef.CreateParameter("aux2", adVarChar, adParamInput, 3)
                .Parameters.Append combenef.CreateParameter("aux3", adVarChar, adParamInput, 3)
                .Parameters.Append combenef.CreateParameter("SIBs", adDouble, adParamOutput)
                .Parameters.Append combenef.CreateParameter("SISus", adDouble, adParamOutput)
                .Parameters("FFInicio") = Me.DTPinicio.Value
                .Parameters("FFFinal") = Me.DTPfin.Value
                .Parameters("cuenta") = Trim(Me.cbocta.Text)
                .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
                .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
                .Parameters("beneficiario") = Trim(Me.txtbusca1.Text)
                .Parameters("aux1") = Trim(Me.txtax1)
                .Parameters("aux2") = "00"
                .Parameters("aux3") = "00"
                .ActiveConnection = db
                .Execute
                SaldoIBs = .Parameters("SIBs")
                SaldoISus = .Parameters("SISus")
            End With
            
            'Me.ProgressBar1.Visible = True
            'Me.ProgressBar1.Value = 0
                CryLMayorBenef.Destination = crptToWindow
                CryLMayorBenef.WindowState = crptMaximized
                CryLMayorBenef.WindowShowPrintSetupBtn = True
                CryLMayorBenef.WindowShowSearchBtn = True
                CryLMayorBenef.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxBenef.rpt"
                CryLMayorBenef.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
                CryLMayorBenef.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
                CryLMayorBenef.StoredProcParam(2) = Trim(Me.cbocta.Text)
                CryLMayorBenef.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
                CryLMayorBenef.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
                CryLMayorBenef.StoredProcParam(5) = Trim(Me.txtbusca1)
                CryLMayorBenef.StoredProcParam(6) = Trim(Me.txtax1)
                CryLMayorBenef.StoredProcParam(7) = "00"
                CryLMayorBenef.StoredProcParam(8) = "00"
                CryLMayorBenef.Formulas(0) = "benef = '" & Trim(Me.txtbusca1) & "'"
                CryLMayorBenef.Formulas(1) = "cta = '" & Trim(Me.cbocta.Text) & "'"
                CryLMayorBenef.Formulas(2) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
                CryLMayorBenef.Formulas(3) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
                If txtax1 = "03" Then
                  NombreCaja Trim(txtbusca1)
                End If
                CryLMayorBenef.Formulas(4) = "nombenef = '" & nombenef & "'"
                CryLMayorBenef.Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
                CryLMayorBenef.Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
                CryLMayorBenef.Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
                CryLMayorBenef.Formulas(10) = "SIBs = " & SaldoIBs
                CryLMayorBenef.Formulas(11) = "SISus = " & SaldoISus
                CryLMayorBenef.Formulas(12) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
                CryLMayorBenef.Formulas(13) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
                iResult = CryLMayorBenef.PrintReport
End Sub
Public Sub ReporteCtaBancaria()
'  Set rsctabancaria = New ADODB.Recordset
'            If rsctabancaria.State = 1 Then rsctabancaria.Close
'            Dim SQLVar As String
'            SQLVar = "SELECT fc_bancos.Bco_descripcion_larga,fc_cuenta_bancaria.Cta_codigo," & _
'                     " fc_cuenta_bancaria.Cta_descripcion_larga FROM fc_bancos INNER JOIN " & _
'                     " fc_cuenta_bancaria ON  fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo " & _
'                     "WHERE fc_cuenta_bancaria.Cta_codigo='" & Trim(Me.cboCtaBancaria) & "'"
'            rsctabancaria.Open SQLVar, db, adOpenKeyset, adLockReadOnly
'            ctabancaria = Trim(rsctabancaria!Cta_Codigo)
'            nombanco = Trim(rsctabancaria!Bco_descripcion_larga)
'            nomctabancaria = Trim(rsctabancaria!Cta_descripcion_larga)
'            Set comctabancaria = New ADODB.Command
'            With comctabancaria
'                .CommandType = adCmdStoredProc
'                .CommandText = "SaldoCtaBancaria"
'                .Parameters.Append comctabancaria.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
'                .Parameters.Append comctabancaria.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
'                .Parameters.Append comctabancaria.CreateParameter("cuenta", adVarChar, adParamInput, 5)
'                .Parameters.Append comctabancaria.CreateParameter("subcta1", adVarChar, adParamInput, 3)
'                .Parameters.Append comctabancaria.CreateParameter("subcta2", adVarChar, adParamInput, 3)
'                .Parameters.Append comctabancaria.CreateParameter("ctabancaria", adVarChar, adParamInput, 40)
'                .Parameters.Append comctabancaria.CreateParameter("aux1", adVarChar, adParamInput, 3)
'                .Parameters.Append comctabancaria.CreateParameter("aux2", adVarChar, adParamInput, 3)
'                .Parameters.Append comctabancaria.CreateParameter("aux3", adVarChar, adParamInput, 3)
'                .Parameters.Append comctabancaria.CreateParameter("SIBs", adDouble, adParamOutput)
'                .Parameters.Append comctabancaria.CreateParameter("SISus", adDouble, adParamOutput)
'                .Parameters("FFInicio") = Me.DTPinicio.Value
'                .Parameters("FFFinal") = Me.DTPfin.Value
'                .Parameters("cuenta") = Trim(Me.cbocta.Text)
'                .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
'                .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
'                .Parameters("ctabancaria") = Trim(Me.cboCtaBancaria.Text)
'                .Parameters("aux1") = Trim(Me.txtax1)
'                .Parameters("aux2") = "00"
'                .Parameters("aux3") = "00"
'                .ActiveConnection = db
'                .Execute
'                SaldoIBs = .Parameters("SIBs")
'                SaldoISus = .Parameters("SISus")
'            End With
'                CryLMayorCtaBancaria.Destination = crptToWindow
'                CryLMayorCtaBancaria.WindowState = crptMaximized
'                CryLMayorCtaBancaria.WindowShowPrintSetupBtn = True
'                CryLMayorCtaBancaria.WindowShowSearchBtn = True
'                CryLMayorCtaBancaria.ReportFileName = App.Path & "\REPORTES\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxCta.rpt"
'                CryLMayorCtaBancaria.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
'                CryLMayorCtaBancaria.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
'                CryLMayorCtaBancaria.StoredProcParam(2) = Trim(Me.cbocta.Text)
'                CryLMayorCtaBancaria.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
'                CryLMayorCtaBancaria.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
'                CryLMayorCtaBancaria.StoredProcParam(5) = Trim(Me.cboCtaBancaria)
'                CryLMayorCtaBancaria.StoredProcParam(6) = Trim(Me.txtax1)
'                CryLMayorCtaBancaria.StoredProcParam(7) = "00"
'                CryLMayorCtaBancaria.StoredProcParam(8) = "00"
'
'                CryLMayorCtaBancaria.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
'                CryLMayorCtaBancaria.Formulas(1) = "ctabanco = '" & Trim(Me.cboCtaBancaria) & "'"
'                CryLMayorCtaBancaria.Formulas(2) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
'                CryLMayorCtaBancaria.Formulas(3) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
'                CryLMayorCtaBancaria.Formulas(4) = "nombanco = '" & nombanco & "'"
'                CryLMayorCtaBancaria.Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
'                CryLMayorCtaBancaria.Formulas(6) = "nomctaBancaria = '" & nomctabancaria & "'"
'                CryLMayorCtaBancaria.Formulas(7) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
'                CryLMayorCtaBancaria.Formulas(8) = "nomsubcta2 = '" & Trim(Me.lbsub2) & "'"
'                CryLMayorCtaBancaria.Formulas(11) = "SIBs = " & Val(SaldoIBs)
'                CryLMayorCtaBancaria.Formulas(12) = "SISus= " & Val(SaldoISus)
'                CryLMayorCtaBancaria.Formulas(13) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
'                CryLMayorCtaBancaria.Formulas(14) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
'                iResult = CryLMayorCtaBancaria.PrintReport
'        'End If
'        If iResult <> 0 Then
'           MsgBox CryLMayorBenef.LastErrorNumber & " : " & CryLMayorBenef.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
End Sub

Public Sub reporteconvenio()
'funciona para todos los otros auxiliares
Dim iResult As Integer
    Set comORG = New ADODB.Command ' para obtener los saldos
    With comORG
        .CommandType = adCmdStoredProc
        .CommandText = "SaldoConvenio"
        .Parameters.Append comORG.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("cuenta", adVarChar, adParamInput, 5)
        .Parameters.Append comORG.CreateParameter("subcta1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("subcta2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("organismo", adVarChar, adParamInput, 15)
        .Parameters.Append comORG.CreateParameter("aux1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux3", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("SIBs", adDouble, adParamOutput)
        .Parameters.Append comORG.CreateParameter("SISus", adDouble, adParamOutput)
        .Parameters("FFInicio") = Me.DTPinicio.Value
        .Parameters("FFFinal") = Me.DTPfin.Value
        .Parameters("cuenta") = Trim(Me.cbocta.Text)
        .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
        .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
        .Parameters("organismo") = Trim(DtCIdConvenio)  'Trim(Me.cbosubcta2.Text)
        .Parameters("aux1") = Trim(Me.txtax1.Text)
        .Parameters("aux2") = Trim(Me.Txtax2.Text)
        .Parameters("aux3") = Trim(Me.txtax3.Text)
        .ActiveConnection = db
        .Execute
        SaldoIBs = .Parameters("SIBs")
        SaldoISus = .Parameters("SISus")
    End With
       CryOrg.Destination = crptToWindow
       CryOrg.WindowState = crptMaximized
       CryOrg.WindowShowPrintSetupBtn = True
       CryOrg.WindowShowSearchBtn = True
       CryOrg.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMConvenio.rpt"
       ''"\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxORG.rpt"
       CryOrg.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(2) = Trim(Me.cbocta.Text)
       CryOrg.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
       CryOrg.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
       CryOrg.StoredProcParam(5) = Trim(DtCIdConvenio)   'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2.Text)
       CryOrg.StoredProcParam(6) = Trim(Me.txtax1.Text)
       CryOrg.StoredProcParam(7) = Trim(Me.Txtax2.Text)
       CryOrg.StoredProcParam(8) = Trim(Me.txtax3.Text)
      
       CryOrg.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
       CryOrg.Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
       CryOrg.Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
       CryOrg.Formulas(3) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
       CryOrg.Formulas(4) = "nomorg = '" & Trim(DtCDesConvenio) & "'" ' Trim(DTCNomOrg.Text) & "'"
       CryOrg.Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
       CryOrg.Formulas(6) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
       CryOrg.Formulas(7) = "organismo ='" & Trim(DtCIdConvenio) & "'" 'Trim(DtcOrg.Text) & "'"
       CryOrg.Formulas(10) = "SIBs = " & SaldoIBs
       CryOrg.Formulas(11) = "SISus = " & SaldoISus
       CryOrg.Formulas(12) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
       CryOrg.Formulas(13) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = CryOrg.PrintReport
        If iResult <> 0 Then
            MsgBox CryOrg.LastErrorNumber & " : " & CryOrg.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

End Sub

Public Sub reporteBeneficiario_COnvenios()
  If rsbeneficiario.State = 1 Then rsbeneficiario.Close
  rsbeneficiario.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
  If rsbeneficiario.RecordCount <> 0 Then
    nombenef = rsbeneficiario!beneficiario_denominacion
  Else
    nombenef = ""
  End If
  With CryBenefConvenios
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        '.ReportFileName = App.path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLMBenef_Convenios.rpt"
        .ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLMBenef_Convenios.rpt"
       ''"\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxORG.rpt"
        .StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
        .StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
        .StoredProcParam(2) = Trim(Me.cbocta.Text)
        .StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
        .StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
        .StoredProcParam(5) = Trim(Me.txtax1.Text)
        .StoredProcParam(6) = Trim(Me.Txtax2.Text)
        .StoredProcParam(7) = Trim(Me.txtax3.Text)
        .StoredProcParam(8) = Trim(txtbusca1.Text)
        .Formulas(0) = "benef = '" & Trim(txtbusca1) & "'"
        .Formulas(1) = "cta = '" & Trim(Me.cbocta.Text) & "'"
        .Formulas(2) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
        .Formulas(3) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
        .Formulas(4) = "nombenef = '" & Trim(nombenef) & "'"
        .Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
        .Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
        .Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
        .Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
        .Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = .PrintReport
        If iResult <> 0 Then
            MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
End With
End Sub
Public Sub NombreCaja(cajas As String)
Dim rsbuscaja As ADODB.Recordset
Set rsbuscaja = New ADODB.Recordset
rsbuscaja.Open "select denominacion_caja from cc_cajas where codigo_caja='" & cajas & "'", db, adOpenKeyset, adLockReadOnly
If rsbuscaja.RecordCount <> 0 Then
  nombenef = Trim(rsbuscaja!denominacion_caja)
End If
End Sub

Private Sub CmBusq_Click()
''FILTRO
'    Set rstAo_solicitud1 = New ADODB.Recordset
'    If rstAo_solicitud1.State = 1 Then rstAo_solicitud1.Close
'    If cboUnidad.Text = "Todas" And CboStatus = "Todas" And TxtCtaNom = "" Then
'        queryinicial2 = "SELECT * FROM Co_BalanceApertura order by correl"
'    Else
'        If cboUnidad.Text = "Todas" And CboStatus = "" And TxtCtaNom = "" Then
'           queryinicial2 = "SELECT * FROM Co_BalanceApertura WHERE status = '" & CboStatus & "'  order by correl"
'        Else
'           If CboStatus = "Todas" And cboUnidad.Text = "" And TxtCtaNom = "" Then
'            queryinicial2 = "SELECT * FROM Co_BalanceApertura WHERE CUENTA = '" & cboUnidad & "'  order by correl"
'           Else
'            If cboUnidad.Text = "" And CboStatus = "" And TxtCtaNom <> "" Then
'                queryinicial2 = "SELECT * FROM Co_BalanceApertura WHERE NombreCta like '%' + '" & Trim(TxtCtaNom) & "' + '%'   order by correl"
'            Else
'                queryinicial2 = "select * from Co_BalanceApertura WHERE CUENTA = '" & cboUnidad.Text & "' AND status = '" & CboStatus & "' and ('%' + NombreCta like '" & TxtCtaNom & "' + '%' )   order by correl"
'            End If
'           End If
'        End If
'    End If
'    rstAo_solicitud1.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'    Set adosolicitud1.Recordset = rstAo_solicitud1
'    Label6(3) = adosolicitud1.Recordset.RecordCount
End Sub

Private Sub BtnAñadir_Click()
'ADICIONA BIEN
   Call Abre_Balance
   'If adosolicitud1.Recordset.RecordCount > 0 Then
      'If adosolicitud1.Recordset!Status = "N" Then
        swgraba3 = 0
        Frame1.Enabled = True
        Frame4.Visible = True
        frmabm.Visible = False
        FrmGraba.Visible = True
        BtnEnviar.Visible = True
        BtnGrabar.Visible = False
        
        Call Limpia_combos
                
        DataGrid3.AllowAddNew = False
        DataGrid3.AllowDelete = False
        DataGrid3.AllowUpdate = False
        DataGrid3.Enabled = False
      
        'marca1 = adosolicitud1.Recordset.Bookmark
        Dim rs_ao_bien As New ADODB.Recordset
        Set rs_ao_bien = New ADODB.Recordset
        If rs_ao_bien.State = 1 Then rs_ao_bien.Close
        'rs_ao_bien.Open "select * from ao_solicitud_bien where ges_gestion = '" & adosolicitud1.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud1.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud1.Recordset!codigo_solicitud & " and codDetalle = '0' ", db, adOpenDynamic, adLockOptimistic
        rs_ao_bien.Open "select * from co_balanceApertura where Cuenta = '0' and  SubCta1 = '0' and  SubCta2 = '0' ", db, adOpenDynamic, adLockOptimistic
        If rs_ao_bien.RecordCount > 0 Then
            MsgBox "No se puede Adicionar un NUEVO registro, mientras exista otro PENDIENTE !!...", vbInformation, "Formulario "
            Exit Sub
        Else
            If rs_ao_bien.State = 1 Then rs_ao_bien.Close
            'db.Execute "UPDATE co_balanceApertura SET cuenta= '" & VARC & "', subcta1='" & VARS1 & "', subcta2 = '" & VARS2 & "', aux1 = '" & VARA1 & "', aux2 = '" & VARA2 & "', aux3 = '" & VARA3 & "', denominacion_aux1 = '" & VARAA1 & "', denominacion_aux2 = '" & VARAA2 & "', denominacion_aux3 = '" & VARAA3 & "', NombreCta = '" & VARNC & "', DebeSaldoIBs = " & VARDB & ", HaberSaldoIBs = " & VARHB & ", Cod_Anterior = '" & VARCA & "', Status = '" & VARES & "' WHERE correl = '" & VCorrel & "' "
            db.Execute "INSERT INTO co_balanceApertura (cuenta,subcta1,subcta2, aux1, AUX2, aux3, denominacion_aux1, denominacion_aux2, denominacion_aux3, NombreCta, DebeSaldoIBs, DebeSaldoISus, HaberSaldoIBs, HaberSaldoISus, Cod_Anterior, Status, Verificado, Nom_Aux1, Nom_Aux2, Nom_Aux3) VALUES ('0', '0', '0', '0', '0', '0','-', '-', '-', '', 0, 0, 0, 0, '0', 'N', 'N', '', '', '') "
'            rstAo_solicitud1.AddNew
'            rstAo_solicitud1!cuenta = "0"
'            rstAo_solicitud1!subcta1 = "0"
'            rstAo_solicitud1!subcta2 = "0"
'            rstAo_solicitud1!aux1 = "0"
'            rstAo_solicitud1!AUX2 = "0"
'            rstAo_solicitud1!aux3 = "0"
'            rstAo_solicitud1!denominacion_aux1 = ""
'            rstAo_solicitud1!denominacion_aux2 = ""
'            rstAo_solicitud1!denominacion_aux3 = ""
'            rstAo_solicitud1!DebeSaldoIBs = 0
'            rstAo_solicitud1!DebeSaldoISus = 0
'            rstAo_solicitud1!HaberSaldoIBs = 0
'            rstAo_solicitud1!HaberSaldoISus = 0
'            rstAo_solicitud1!Status = "N"
'            rstAo_solicitud1.Update
            'swgraba3 = 0
            Call Abre_Balance
            rstAo_solicitud1.MoveLast
            'adosolicitud1.Recordset.Bookmark = marca1
            'adosolicitud1.Refresh
        End If
'        DataGrid3.Enabled = False
      'Else
      '   MsgBox "No se puede modificar un registro APROBADO o ANULADO ", vbInformation, "Formulario 1"
      'End If
   'Else
   '       MsgBox "No Existen Registros habilitados ", vbInformation, "Formulario 1"
   'End If
End Sub

Private Sub BtnModificar_Click()
'MODIFICA BIEN
   If adosolicitud1.Recordset.RecordCount > 0 Then
      If adosolicitud1.Recordset!Status = "N" Then
        swgraba3 = 1
        Frame1.Enabled = True
        Frame4.Visible = True
        frmabm.Visible = False
        FrmGraba.Visible = True
        BtnEnviar.Visible = True
        BtnGrabar.Visible = False
        
        Call Limpia_combos
                
        DataGrid3.AllowAddNew = False
        DataGrid3.AllowDelete = False
        DataGrid3.AllowUpdate = False
        DataGrid3.Enabled = False
      Else
         MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Formulario"
      End If
   Else
          MsgBox "No Existen Registros ", vbInformation, "Formulario"
   End If
End Sub

Private Sub BtnEliminar_Click()
'ELIMINA BIEN
   If adosolicitud1.Recordset.RecordCount > 0 Then
      If adosolicitud1.Recordset!Status = "N" Then
        sino = MsgBox("Está seguro(a) de ELIMINAR este registro ? ", vbYesNo + vbExclamation, "Atención")
        If sino = vbYes Then
            DataGrid3.AllowAddNew = False
            DataGrid3.AllowDelete = False
            DataGrid3.AllowUpdate = False
            VARC = DataGrid3.Columns("Cuenta").Value
            VARS1 = DataGrid3.Columns("SubCta1").Value
            VARS2 = DataGrid3.Columns("SubCta2").Value
            VARA1 = DataGrid3.Columns("Aux1").Value
            VARA2 = DataGrid3.Columns("Aux2").Value
            VARA3 = DataGrid3.Columns("Aux3").Value
            VARAA1 = DataGrid3.Columns("denominacion_aux1").Value
            VARAA2 = DataGrid3.Columns("denominacion_aux2").Value
            VARAA3 = DataGrid3.Columns("denominacion_aux3").Value
            db.Execute "DELETE co_balanceApertura where Cuenta = '" & VARC & "' and SubCta1 = '" & VARS1 & "' and SubCta2 = '" & VARS2 & "' and Aux1 = '" & VARA1 & "' and Aux2 = '" & VARA2 & "' and Aux3 = '" & VARA3 & "' and denominacion_aux1 = '" & VARAA1 & "' and denominacion_aux2 = '" & VARAA2 & "' and denominacion_aux3 = '" & VARAA3 & "'  "
            swgraba3 = 2
            Call Abre_Balance
        End If

      Else
         MsgBox "No se puede ELIMINAR un registro APROBADO ", vbInformation, "Formulario "
      End If
   Else
          MsgBox "No Existen Registros ", vbInformation, "Formulario "
   End If
End Sub

Private Sub DataGrid3_LostFocus()
'On Error GoTo Error
'
'' If swgraba3 <> 0 Then
'   'If adosolicitud1.Recordset.RecordCount > 0 And Not IsNull(DataGrid3.Columns("NombreCta").Value) And (DataGrid3.Columns("NombreCta").Value) <> "" Then
'   If Not IsNull(DataGrid3.Columns("NombreCta").Value) And (DataGrid3.Columns("NombreCta").Value) <> "" Then
'      'If adosolicitud1.Recordset!Status = "N" Then
'        VARC = DataGrid3.Columns("Cuenta").Value
'        VARS1 = DataGrid3.Columns("SubCta1").Value
'        VARS2 = DataGrid3.Columns("SubCta2").Value
'        VARA1 = DataGrid3.Columns("Aux1").Value
'        VARA2 = DataGrid3.Columns("Aux2").Value
'        VARA3 = DataGrid3.Columns("Aux3").Value
'        VARAA1 = DataGrid3.Columns("denominacion_aux1").Value
'        VARAA2 = DataGrid3.Columns("denominacion_aux2").Value
'        VARAA3 = DataGrid3.Columns("denominacion_aux3").Value
'        Dim rs_ao_bien As New ADODB.Recordset
'        Set rs_ao_bien = New ADODB.Recordset
'        If rs_ao_bien.State = 1 Then rs_ao_bien.Close
'
'        tot_form = 0
'        rs_ao_bien.Open "select COUNT(*) AS tot_form from co_balanceApertura where Cuenta = '" & VARC & "' and SubCta1 = '" & VARS1 & "' and SubCta2 = '" & VARS2 & "' and Aux1 = '" & VARA1 & "' and Aux2 = '" & VARA2 & "' and Aux3 = '" & VARA3 & "' and denominacion_aux1 = '" & VARAA1 & "' and denominacion_aux2 = '" & VARAA2 & "' and denominacion_aux3 = '" & VARAA3 & "'  ", db, adOpenDynamic, adLockOptimistic
'        If rs_ao_bien!tot_form > 1 Then
'            MsgBox "No se puede Guardar un registro ya EXISTENTE, verifique por favor !!...", vbInformation, "Formulario 04"
'            DataGrid3.SetFocus
'            Exit Sub
'        Else
'            If rs_ao_bien.State = 1 Then rs_ao_bien.Close
'
''            marca1 = adosolicitud1.Recordset.Bookmark
'            'Dim VARPU, VARCAN, VARPT As Double
''            VARAA1 = DataGrid3.Columns("denominacion_aux1").Value
''            VARAA2 = DataGrid3.Columns("denominacion_aux2").Value
''            VARAA3 = DataGrid3.Columns("denominacion_aux3").Value
'            VARNC = DataGrid3.Columns("NombreCta").Value
'            VARDB = DataGrid3.Columns("DebeSaldoIBs").Value
'            VARHB = DataGrid3.Columns("HaberSaldoIBs").Value
'            'VARPT = DataGrid3.Columns("precio_venta").Value * DataGrid3.Columns("cantidad").Value
'            VARCA = DataGrid3.Columns("Cod_Anterior").Value
'            VARES = DataGrid3.Columns("status").Value
'    '        MarcaB = adosolicitud11.Recordset.Bookmark
'    '        Call Abre_Balance
'    '        'MarcaB = rstAo_solicitud1.Bookmark
'    '        adosolicitud11.Recordset.Bookmark = MarcaB
'            rstAo_solicitud1!cuenta = VARC
'            rstAo_solicitud1!subcta1 = VARS1
'            rstAo_solicitud1!subcta2 = VARS2
'            rstAo_solicitud1!aux1 = VARA1
'            rstAo_solicitud1!AUX2 = VARA2
'            rstAo_solicitud1!aux3 = VARA3
'            rstAo_solicitud1!denominacion_aux1 = VARAA1
'            rstAo_solicitud1!denominacion_aux2 = VARAA2
'            rstAo_solicitud1!denominacion_aux3 = VARAA3
'            rstAo_solicitud1!NombreCta = VARNC
'            rstAo_solicitud1!DebeSaldoIBs = VARDB
'            rstAo_solicitud1!HaberSaldoIBs = VARHB
'            rstAo_solicitud1!Cod_Anterior = VARCA
'            rstAo_solicitud1!Status = VARES
'            rstAo_solicitud1.Update
''            swgraba3 = 0
'            Call Abre_Balance
'            rstAo_solicitud1.MoveLast
'            BtnAñadir.Enabled = True
'            BtnModificar.Enabled = True
'            BtnEliminar.Enabled = True
'            DataGrid3.AllowAddNew = False
'            DataGrid3.AllowDelete = False
'            DataGrid3.AllowUpdate = False
'        End If
'      'Else
'      '   MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Formulario 04"
'      'End If
'   Else
'         MsgBox "Verifique los datos para continuar ... ", vbInformation, "Formulario 04"
'   End If
' 'End If
'Error:
'ErrorB = Err.Number
'    If Err.Number = -2147467259 Then
'        MsgBox "ERROR B.A.: El registro actual ya fue GUARDADO anteriormente, verifique por favor !!...", vbInformation, "Formulario 04"
'    End If
'    If Err.Number = -2147217887 Then
'        MsgBox "Se producjo un error desconocido!...", vbCritical + vbOKOnly, "Error..."
'    End If
End Sub

Private Sub DataGrid3_UnboundColumnFetch(Bookmark As Variant, ByVal Col As Integer, Value As Variant)
    RSCloneB.Bookmark = Bookmark
    Value = RSCloneB("NombreCta")
End Sub

Private Sub TDBPlan_DropDownClose()
On Error GoTo err1
    DataGrid3.Columns("Cuenta").Value = TDBPlan.Columns("Cuenta").Value
    DataGrid3.Columns("SubCta1").Value = TDBPlan.Columns("SubCta1").Value
    DataGrid3.Columns("SubCta2").Value = TDBPlan.Columns("SubCta2").Value
    DataGrid3.Columns("Aux1").Value = TDBPlan.Columns("Aux1").Value
    DataGrid3.Columns("Aux2").Value = TDBPlan.Columns("Aux2").Value
    DataGrid3.Columns("Aux3").Value = TDBPlan.Columns("Aux3").Value
    DataGrid3.Columns("NombreCta").Value = TDBPlan.Columns("NombreCta").Value
    
    cbocta.Text = DataGrid3.Columns(0)
    cbosubcta1.Text = DataGrid3.Columns(1)
    cbosubcta2.Text = DataGrid3.Columns(2)
    txtax1.Text = DataGrid3.Columns(3)
    Txtax2.Text = DataGrid3.Columns(4)
    txtax3.Text = DataGrid3.Columns(5)
    
    '-- Cuenta
    rsplanctas.MoveFirst
    rsplanctas.Find "cuenta=" & "'" & Trim(cbocta.Text) & "'"
    Me.lblcuenta = rsplanctas!NombreCta
    If rscuentas.State = adStateOpen Then rscuentas.Close
    '-- SubCuenta1
    If rsnombresub1.State = adStateOpen Then rsnombresub1.Close
    rsnombresub1.Open "SELECT NombreCta FROM CC_Plan_Cuentas WHERE   (SubCta2 = '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & (Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
    Me.Lblsub1 = rsnombresub1!NombreCta
    '-- SubCuenta2
    If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
    rssubcuenta.Open "SELECT Cuenta, SubCta1, SubCta2, NombreCta, Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & Trim(Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly

    Call carga_ctas

err1:
    If Err.Number = 7005 Then
    DtgPlanCtas.Refresh
    End If
End Sub
    
Private Sub carga_ctas()
  With rssubcuenta
    .MoveFirst
    .Find "subcta2=" & "'" & Trim(Me.cbosubcta2) & "'"
    Me.lbsub2 = !NombreCta
    Me.txtax1 = !aux1
    Me.Txtax2 = !AUX2
    Me.txtax3 = !aux3
    Chkaux1.Enabled = True
    Chkaux2.Enabled = True
    Chkaux3.Enabled = True
    Chkaux1.Value = 1
    Chkaux2.Value = 1
    Chkaux3.Value = 1
    'BtnBuscarA.Enabled = False
    '--------
    Call Limpia_combos
    
    BtnEnviar.Visible = True
    BtnGrabar.Visible = False
    
    Select Case !aux1
      Case "00"
        Chkaux1.Enabled = False
        Chkaux1.Value = 0
      Case "01"
        txtbusca1.Visible = True
        txtbusca1.Top = 1260
      Case "02"
'        cboCtaBancaria.Visible = True
'        cboCtaBancaria.Top = 1260
      Case "03"
        DtcProy.Visible = True
        DtcProy.Top = 1260
        DtcProyDes.Visible = True
        DtcProyDes.Top = 1260
      Case "05"
        DtcGrBien.Visible = True
        DtcGrBien.Top = 1260
        DtcGrBienDes.Visible = True
        DtcGrBienDes.Top = 1260
      Case "08"
        DTCNomOrg.Visible = True
        DTCNomOrg.Top = 1260
        DtCOrg.Visible = True
        DtCOrg.Top = 1260
      Case "09"
        DtCIdConvenio.Visible = True
        DtCIdConvenio.Top = 1260
        DtCDesConvenio.Visible = True
        DtCDesConvenio.Top = 1260
    End Select
    Select Case !AUX2
      Case "00"
        Chkaux2.Enabled = False
        Chkaux2.Value = 0
      Case "01"
        txtbusca1.Visible = True
        txtbusca1.Top = 1620
      Case "02"
'        cboCtaBancaria.Visible = True
''        cboCtaBancaria.Top = 1620
      Case "03"
        DtcProy.Visible = True
        DtcProy.Top = 1620
        DtcProyDes.Visible = True
        DtcProyDes.Top = 1620
      Case "05"
        DtcGrBien.Visible = True
        DtcGrBien.Top = 1620
        DtcGrBienDes.Visible = True
        DtcGrBienDes.Top = 1620
      Case "08"
        DTCNomOrg.Visible = True
        DTCNomOrg.Top = 1620
        DtCOrg.Visible = True
        DtCOrg.Top = 1620
      Case "09"
        DtCIdConvenio.Visible = True
        DtCIdConvenio.Top = 1620
        DtCDesConvenio.Visible = True
        DtCDesConvenio.Top = 1620
    End Select
    Select Case !aux3
      Case "00"
        Chkaux3.Enabled = False
        Chkaux3.Value = 0
      Case "01"
        txtbusca1.Visible = True
        txtbusca1.Top = 1980
      Case "02"
'        cboCtaBancaria.Visible = True
'        cboCtaBancaria.Top = 1980
      Case "03"
        DtcProy.Visible = True
        DtcProy.Top = 1980
        DtcProyDes.Visible = True
        DtcProyDes.Top = 1980
      Case "05"
        DtcGrBien.Visible = True
        DtcGrBien.Top = 1980
        DtcGrBienDes.Visible = True
        DtcGrBienDes.Top = 1980
      Case "08"
        DTCNomOrg.Visible = True
        DTCNomOrg.Top = 1980
        DtCOrg.Visible = True
        DtCOrg.Top = 1980
      Case "09"
        DtCIdConvenio.Visible = True
        DtCIdConvenio.Top = 1980
        DtCDesConvenio.Visible = True
        DtCDesConvenio.Top = 1980
    End Select
  End With
'*******Se filtra si la cuenta es de bancos....
If Me.cbocta = "1111" And Me.cbosubcta1 = "01" Then
    Select Case Me.cbosubcta2
        Case "00"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '10' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
        Case "01"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
        Case "02"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '43' or fc_cuenta_bancaria.Fte_codigo = '70' order by fc_cuenta_bancaria.Cta_codigo"
        Case "03"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
     End Select
'    Me.cboCtaBancaria.Clear
    If rscta_bancaria.State = 1 Then rscta_bancaria.Close
    rscta_bancaria.Open sql1, db, adOpenKeyset, adLockReadOnly
    If rscta_bancaria.RecordCount <> 0 Then
        rscta_bancaria.MoveFirst
    End If
        Do While Not rscta_bancaria.EOF
'          cboCtaBancaria.AddItem rscta_bancaria!Cta_Codigo
          rscta_bancaria.MoveNext
        Loop
'    Me.cboCtaBancaria.Visible = True
'    Me.cboCtaBancaria.Text = Me.cboCtaBancaria.List(0)
    Me.txtbusca1.Visible = False
    Me.DTGBanco.Visible = True
    Me.DtGbenef.Visible = False
    DtgPlanCtas.Visible = False
    Set Me.DTGBanco.DataSource = rscta_bancaria
End If
If Me.cbocta = "1111" And Me.cbosubcta1 = "02" Then
    Select Case Me.cbosubcta2
        Case "00"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '30' or fc_cuenta_bancaria.Fte_codigo = '50' order by fc_cuenta_bancaria.Cta_codigo"
        Case "01"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
        Case "02"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '43' or fc_cuenta_bancaria.Fte_codigo = '70' order by fc_cuenta_bancaria.Cta_codigo"
        Case "03"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.cta_descripcion,  fc_bancos.bco_descripcion FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
     End Select
'    Me.cboCtaBancaria.Clear
    If rscta_bancaria.State = 1 Then rscta_bancaria.Close
    rscta_bancaria.Open sql1, db, adOpenKeyset, adLockReadOnly
    If rscta_bancaria.RecordCount <> 0 Then
        rscta_bancaria.MoveFirst
    End If
        Do While Not rscta_bancaria.EOF
'          cboCtaBancaria.AddItem rscta_bancaria!Cta_Codigo
          rscta_bancaria.MoveNext
        Loop
'    Me.cboCtaBancaria.Visible = True
'    Me.cboCtaBancaria.Text = Me.cboCtaBancaria.List(0)
    Me.txtbusca1.Visible = False
    Me.DTGBanco.Visible = True
    Me.DtGbenef.Visible = False
    DtgPlanCtas.Visible = False
    Set Me.DTGBanco.DataSource = rscta_bancaria
End If

'************Se habilita la tabla de beneficiarios
    If Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01" Then
        If rsbeneficiario.State = 1 Then rsbeneficiario.Close
        sql2 = "SELECT beneficiario_codigo, beneficiario_denominacion From gc_beneficiario WHERE estado_codigo = 'APR' order by beneficiario_denominacion"
        rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rsbeneficiario
        Set Ado_Benef.Recordset = rsbeneficiario
        Me.DtGbenef.Visible = True
        Me.DTGBanco.Visible = False
        Me.txtbusca1.Visible = True
        Me.BtnBuscarA.Enabled = True
'        Me.cboCtaBancaria.Visible = False
        DtgPlanCtas.Visible = False
    End If
    If Me.txtax1 = "03" Or Me.Txtax2 = "03" Or Me.txtax3 = "03" Then
        If rsProy.State = 1 Then rsProy.Close
        rsProy.Open "SELECT pro_codigo, Pro_programa, Pro_proyecto, Pro_actividad, pro_descripcion From fc_estructura_programatica WHERE estado_codigo = 'APR' AND pro_nivel > 1 order by pro_descripcion", db, adOpenKeyset, adLockReadOnly
        'rsProy.Open "SELECT Pro_programa, Pro_proyecto, Pro_actividad, Pro_descripcion_larga From fc_estructura_programatica WHERE Pro_activo = '1' order by Pro_descripcion_larga", db, adOpenKeyset, adLockReadOnly
        'Set Me.DtGbenef.DataSource = rsproyecto
        Set AdoProy.Recordset = rsProy
        AdoProy.Refresh
        'Me.DtGbenef.Visible = True
        'Me.DTGBanco.Visible = False
        'Me.txtbusca1.Visible = True
        'Me.BtnBuscarA.Enabled = True
        'Me.cboCtaBancaria.Visible = False
        'DtgPlanCtas.Visible = False
    End If
    If Me.txtax1 = "05" Or Me.Txtax2 = "05" Or Me.txtax3 = "05" Then
        If rsGrupoBien.State = 1 Then rsGrupoBien.Close
        rsGrupoBien.Open "SELECT Cod_montador, DESCRIPCION, PAR_CODIGO  From Al_Montador WHERE estado = 'S' order by DESCRIPCION", db, adOpenKeyset, adLockReadOnly
        Set AdoGrBien.Recordset = rsGrupoBien
        AdoGrBien.Refresh
    End If

'****habilitamos boton de búsqueda
    If Me.txtax1 = "00" Or Me.txtax1 = "02" Then
        Me.BtnBuscarA.Enabled = False
    Else
        Me.BtnBuscarA.Enabled = True
    End If
'    If Me.txtax1 = "03" Then
'      Me.txtbusca1.Visible = True
'    End If
'    If Me.txtax1 = "05" Then
'      Me.txtbusca1.Visible = True
'    End If
    '-------- habilito datacombos para organismo financiadores
  If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
    txtbusca1.Visible = False
    DtCIdConvenio.Visible = False
    DtCDesConvenio.Visible = False
    Dtc_benef.Visible = True
    DtcCodAux2.Visible = True
'    DtcCodAux3.Visible = True
    Dtc_benefD.Visible = True
    DtcDenomAux2.Visible = True
    DtcDenomAux3.Visible = True
  Else
    Dtc_benef.Visible = False
    DtcCodAux2.Visible = False
'    DtcCodAux3.Visible = False
    Dtc_benefD.Visible = False
    DtcDenomAux2.Visible = False
    DtcDenomAux3.Visible = False
  End If
'--------------------
End Sub

Private Sub Abre_Balance()
    'Call Abre_Sol_Det
'    Dim varpar As String
'    varpar = rs_ao_solicitud_detalle!par_codigo
'    varpar = adoao_solicitud_detalle.Recordset!par_codigo
    Set rs_bien = New ADODB.Recordset
    If rs_bien.State = 1 Then rs_bien.Close
    'SqlBienes = "select * from CC_Plan_Cuentas WHERE PAR_CODIGO= '" & varpar & "' "
    SqlBienes = "select * from CC_Plan_Cuentas "
    rs_bien.Open SqlBienes, db, adOpenStatic, adLockReadOnly
    rs_bien.Sort = "Cuenta, SubCta1, SubCta2"
    TDBPlan.DataSource = rs_bien
    Set AdoPlan.Recordset = rs_bien

    Set rstAo_solicitud1 = New ADODB.Recordset
    If rstAo_solicitud1.State = 1 Then rstAo_solicitud1.Close
    'SqlBien = "select * from co_balanceApertura where ges_gestion = '" & adosolicitud1.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud1.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud1.Recordset!codigo_solicitud
    SqlBien = "select * from co_balanceApertura "
    rstAo_solicitud1.Open SqlBien, db, adOpenKeyset, adLockOptimistic
    rstAo_solicitud1.Sort = "correl"
'    If rstAo_solicitud1.RecordCount < 1 Then
'        TDBSolicitud_Bien.DataSource = RSNADA
'    Else
        DataGrid3.DataSource = rstAo_solicitud1
'    End If
    Set adosolicitud1.Recordset = rstAo_solicitud1
    'adosolicitud11.Refresh
End Sub


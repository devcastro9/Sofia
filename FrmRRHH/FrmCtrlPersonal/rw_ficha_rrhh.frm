VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form rw_ficha_rrhh 
   BackColor       =   &H00E0E0E0&
   Caption         =   "RRHH - Procesos - Ficha Personal"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "rw_ficha_rrhh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9480
      Left            =   6720
      TabIndex        =   11
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   16722
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PERSONALES"
      TabPicture(0)   =   "rw_ficha_rrhh.frx":0A02
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraGrabarCancelar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraOpciones"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDatos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "CONTROL ASISTENCIA"
      TabPicture(1)   =   "rw_ficha_rrhh.frx":0A1E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame18"
      Tab(1).Control(1)=   "Label44"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "PERMISOS-VACACIONES"
      TabPicture(2)   =   "rw_ficha_rrhh.frx":0A3A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame14"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(2)=   "Label45"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "MOVILIDAD PERSONAL"
      TabPicture(3)   =   "rw_ficha_rrhh.frx":0A56
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label46"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame17"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame16"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "CURRICULARES"
      TabPicture(4)   =   "rw_ficha_rrhh.frx":0A72
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame15"
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(2)=   "Label5"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "CONTRATOS Y LIQUIDACIONES"
      TabPicture(5)   =   "rw_ficha_rrhh.frx":0A8E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame10"
      Tab(5).Control(1)=   "Frame20"
      Tab(5).Control(2)=   "Label16"
      Tab(5).ControlCount=   3
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PROGRAMACION DE VACACIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4335
         Left            =   -74880
         TabIndex        =   158
         Top             =   1140
         Width           =   12615
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   390
            Left            =   120
            TabIndex        =   186
            Top             =   3840
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   688
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.CommandButton CmdApr2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":0AAA
            Style           =   1  'Graphical
            TabIndex        =   165
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdElim2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Anular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":1034
            Style           =   1  'Graphical
            TabIndex        =   164
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdAdd2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":1A36
            Style           =   1  'Graphical
            TabIndex        =   163
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdMod2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Modif."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":1FC0
            Style           =   1  'Graphical
            TabIndex        =   162
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Height          =   650
            Left            =   11595
            TabIndex        =   161
            Top             =   140
            Width           =   615
            Begin VB.Image Img_CV 
               Height          =   540
               Left            =   15
               Picture         =   "rw_ficha_rrhh.frx":254A
               Top             =   60
               Width           =   555
            End
         End
         Begin VB.CommandButton btnimprimir1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3480
            Picture         =   "rw_ficha_rrhh.frx":28D2
            Style           =   1  'Graphical
            TabIndex        =   160
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Vacaciones para todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   159
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   1455
         End
         Begin MSAdodcLib.Adodc Ado_VacacionesProg 
            Height          =   330
            Left            =   120
            Top             =   3495
            Width           =   12375
            _ExtentX        =   21828
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
            BackColor       =   16761024
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
            Caption         =   " <--- Programación de Vacaciones --->"
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
         Begin MSDataGridLib.DataGrid DtgVacacionesProg 
            Bindings        =   "rw_ficha_rrhh.frx":2E5C
            Height          =   2655
            Left            =   120
            TabIndex        =   166
            Top             =   840
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BackColor       =   16761024
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
            Caption         =   "PROGRAMACION DE VACACIONES"
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "Correl"
               Caption         =   "Nro."
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
               DataField       =   "ges_Gestion"
               Caption         =   "Gestion"
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
               DataField       =   "Mes_control"
               Caption         =   "Mes Control"
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
               DataField       =   "fecha_ini_Prog"
               Caption         =   "Fecha Inicio"
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
               DataField       =   "fecha_fin_Prog"
               Caption         =   "Fecha Fin"
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
               DataField       =   "dias_Programados"
               Caption         =   "Dia.Prog"
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
               DataField       =   "horas_Programadas"
               Caption         =   "Hrs.Prog"
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
            BeginProperty Column07 
               DataField       =   "minutos_programados"
               Caption         =   "Min.Prog"
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
            BeginProperty Column08 
               DataField       =   "dias_utilizados"
               Caption         =   "Dia.Usado"
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
            BeginProperty Column09 
               DataField       =   "horas_utilizadas"
               Caption         =   "Hr.Usada"
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
            BeginProperty Column10 
               DataField       =   "Dias_Pendientes"
               Caption         =   "Saldo Días"
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
            BeginProperty Column11 
               DataField       =   "estado_codigo"
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
            BeginProperty Column12 
               DataField       =   "observaciones"
               Caption         =   "Observaciones"
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
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1620.284
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   989.858
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   720
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   840.189
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   3195.213
               EndProperty
            EndProperty
         End
         Begin VB.Label LblCV 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Prog. Vacacion -->"
            DataField       =   "ARCHIVO_VAC"
            DataSource      =   "adoLista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   9780
            TabIndex        =   168
            Top             =   555
            Width           =   1605
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Ver Prog.Vacacion -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9480
            TabIndex        =   167
            Top             =   240
            Width           =   1890
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "MEMORANDAS PARA SANCIONES Y AMONESTACIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   148
         Top             =   1140
         Width           =   12615
         Begin MSDataGridLib.DataGrid DtG_Memo 
            Bindings        =   "rw_ficha_rrhh.frx":2E7C
            Height          =   2745
            Left            =   120
            TabIndex        =   155
            Top             =   840
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   4842
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BackColor       =   16777152
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
            Caption         =   "MEMORANDAS PARA SANCIONES Y AMONESTACIONES"
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "numero"
               Caption         =   "Correlativo"
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
               DataField       =   "correl"
               Caption         =   "No.Memo"
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
               DataField       =   "fecha_aprobacion"
               Caption         =   "Fecha Ejecucion"
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
               DataField       =   "fecha_memo"
               Caption         =   "Fecha_Memo"
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
               DataField       =   "usr_codigo"
               Caption         =   "Emitido Por:"
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
               DataField       =   "tipo_memo"
               Caption         =   "Tipo Memo"
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
            BeginProperty Column06 
               DataField       =   "Observaciones"
               Caption         =   "Aclaracion"
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
            BeginProperty Column07 
               DataField       =   "monto"
               Caption         =   "Monto Sanción"
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
            BeginProperty Column08 
               DataField       =   "Dias"
               Caption         =   "Días Sanción"
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
            BeginProperty Column09 
               DataField       =   "descuento_pla"
               Caption         =   "Dcto. Planilla"
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
            BeginProperty Column10 
               DataField       =   "estado_codigo"
               Caption         =   "Estado"
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
                  ColumnWidth     =   510.236
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   989.858
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1035.213
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1425.26
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column09 
                  Alignment       =   2
                  ColumnWidth     =   989.858
               EndProperty
               BeginProperty Column10 
                  Alignment       =   2
                  ColumnWidth     =   540.284
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton CmdElim4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "AnuLar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":2E97
            Style           =   1  'Graphical
            TabIndex        =   154
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdApr4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":3899
            Style           =   1  'Graphical
            TabIndex        =   153
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdMod4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Modif."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":3E23
            Style           =   1  'Graphical
            TabIndex        =   152
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdAdd4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":43AD
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   690
            Left            =   11355
            TabIndex        =   150
            Top             =   100
            Width           =   615
            Begin VB.Image Img_CTO 
               Height          =   540
               Left            =   20
               Picture         =   "rw_ficha_rrhh.frx":4937
               Top             =   100
               Width           =   555
            End
         End
         Begin VB.CommandButton btnimprimir3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            Picture         =   "rw_ficha_rrhh.frx":4CBF
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin MSAdodcLib.Adodc Ado_Memo 
            Height          =   330
            Left            =   120
            Top             =   3600
            Width           =   12375
            _ExtentX        =   21828
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
            BackColor       =   16777152
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
            Caption         =   " <--- Memorandas para Sanciones y Amonestaciones --->"
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
         Begin VB.Label LblCto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Cargar Archivo -->"
            DataField       =   "ARCHIVO"
            DataSource      =   "Ado_Memo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   9240
            TabIndex        =   157
            Top             =   540
            Width           =   1560
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Ver Memo-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9705
            TabIndex        =   156
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ASCENSOS, PROMOCIONES Y CAMBIOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   138
         Top             =   5220
         Width           =   12615
         Begin VB.CommandButton CmdElim5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "AnuLar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":5249
            Style           =   1  'Graphical
            TabIndex        =   144
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdApr5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":5C4B
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Aprueba Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdMod5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Modif."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":61D5
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdAdd5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":675F
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00E0E0E0&
            Height          =   690
            Left            =   11355
            TabIndex        =   140
            Top             =   100
            Width           =   615
            Begin VB.Image ImgFiniquito 
               Height          =   540
               Left            =   20
               Picture         =   "rw_ficha_rrhh.frx":6CE9
               Top             =   100
               Width           =   555
            End
         End
         Begin VB.CommandButton btnimprimir2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            Picture         =   "rw_ficha_rrhh.frx":7071
            Style           =   1  'Graphical
            TabIndex        =   139
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DtgMovilidad 
            Bindings        =   "rw_ficha_rrhh.frx":75FB
            Height          =   2775
            Left            =   120
            TabIndex        =   145
            Top             =   840
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   4895
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
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
            Caption         =   "ASCENSOS, PROMOCIONES Y CAMBIOS"
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "numero_cambio"
               Caption         =   "No.Memo"
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
               DataField       =   "fecha_inicio_contrato"
               Caption         =   "Fecha. Memo"
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
               DataField       =   "unidad_antigua"
               Caption         =   "Unidad Origen"
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
               DataField       =   "puesto_antiguo"
               Caption         =   "Puesto Origen"
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
               DataField       =   "unidad_descripcion"
               Caption         =   "Unidad Destino"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "puesto_descripcion"
               Caption         =   "Puesto_Destino"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "item"
               Caption         =   "Autorizado Por:"
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
            BeginProperty Column07 
               DataField       =   "estado_codigo"
               Caption         =   "Estado"
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
            BeginProperty Column08 
               DataField       =   "tipo_mov"
               Caption         =   "Tipo Movilidad"
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
            BeginProperty Column09 
               DataField       =   "beneficiario_denominacion"
               Caption         =   "Cambio de puesto con:"
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
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1065.26
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2564.788
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2294.929
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2174.74
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   2415.118
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1544.882
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   555.024
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   2204.788
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoMovilidad 
            Height          =   330
            Left            =   120
            Top             =   3630
            Width           =   12375
            _ExtentX        =   21828
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
            Caption         =   " <--- Ascensos, Promociones y Cambios --->"
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
         Begin VB.Label LblLiq 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Cargar Archivo -->"
            DataField       =   "ARCHIVO"
            DataSource      =   "AdoMovilidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   9420
            TabIndex        =   147
            Top             =   540
            Width           =   1560
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Ver Registro -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9615
            TabIndex        =   146
            Top             =   240
            Width           =   1350
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CONTROL DE ASISTENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   8175
         Left            =   -74880
         TabIndex        =   128
         Top             =   1140
         Width           =   12615
         Begin VB.TextBox txt_mes 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFF00&
            Height          =   285
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   136
            Text            =   "0"
            Top             =   480
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.CommandButton CmdApr1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":7618
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdElim1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "AnuLar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":7BA2
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdMod1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Modifica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":85A4
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdAdd1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":8B2E
            Style           =   1  'Graphical
            TabIndex        =   132
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cbo_gestion 
            Height          =   315
            ItemData        =   "rw_ficha_rrhh.frx":90B8
            Left            =   10920
            List            =   "rw_ficha_rrhh.frx":90DD
            TabIndex        =   131
            Text            =   "GESTION"
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cbo_mes 
            Height          =   315
            ItemData        =   "rw_ficha_rrhh.frx":9123
            Left            =   9120
            List            =   "rw_ficha_rrhh.frx":914E
            TabIndex        =   130
            Text            =   "MES"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3480
            Picture         =   "rw_ficha_rrhh.frx":91BD
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin MSAdodcLib.Adodc AdoAsistencia 
            Height          =   330
            Left            =   120
            Top             =   7800
            Width           =   12375
            _ExtentX        =   21828
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
            BackColor       =   12648384
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
            Caption         =   " <--- Control de Asistencia --->"
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
         Begin MSDataGridLib.DataGrid DtgAsistencia 
            Bindings        =   "rw_ficha_rrhh.frx":9747
            Height          =   6945
            Left            =   120
            TabIndex        =   177
            Top             =   840
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   12250
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
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
            Caption         =   "CONTROL DE ASISTENCIA"
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "Fecha_control"
               Caption         =   "Fecha Control"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "HoraTres"
               Caption         =   "Marca.Ingreso"
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
               DataField       =   "HoraCuatro"
               Caption         =   "Marca.Salida"
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
               DataField       =   "Tardanza"
               Caption         =   "Tardanza"
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
               DataField       =   "TiemAsist"
               Caption         =   "Tiempo.Trabajo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "hh:mm:ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "MES"
               Caption         =   "Mes"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "hh:mm:ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "GESTION"
               Caption         =   "Gestion"
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
            BeginProperty Column07 
               DataField       =   "HoraUno"
               Caption         =   "Hora.Ingreso"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "hh:mm:ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "HoraDos"
               Caption         =   "Hora.Salida"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "hh:mm:ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   4
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   659.906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   975.118
               EndProperty
            EndProperty
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filtrar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   8475
            TabIndex        =   137
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "REGISTRO DE PERMISOS (Licencias - Bajas)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3855
         Left            =   -74880
         TabIndex        =   119
         Top             =   5460
         Width           =   12615
         Begin VB.Frame Frame19 
            BackColor       =   &H00E0E0E0&
            Height          =   650
            Left            =   11520
            TabIndex        =   124
            Top             =   140
            Width           =   615
            Begin VB.Image Img_03 
               Height          =   540
               Left            =   0
               Picture         =   "rw_ficha_rrhh.frx":9763
               Top             =   80
               Width           =   555
            End
         End
         Begin VB.CommandButton CmdApr3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":9AEB
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdElim3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Anular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":A075
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdMod3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Modifica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":AA77
            Style           =   1  'Graphical
            TabIndex        =   121
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdAdd3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":B001
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DtgPermiso 
            Bindings        =   "rw_ficha_rrhh.frx":B58B
            Height          =   2505
            Left            =   120
            TabIndex        =   125
            Top             =   840
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   4419
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BackColor       =   12632319
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
            Caption         =   "REGISTRO DE PERMISOS (Licencias - Bajas)"
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "Mes_control"
               Caption         =   "Mes Control"
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
               DataField       =   "Fecha_control"
               Caption         =   "Fecha Solicitud"
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
               DataField       =   "FechaDesde"
               Caption         =   "Fecha Desde"
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
               DataField       =   "FechaHasta"
               Caption         =   "Fecha Hasta"
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
               DataField       =   "horadesde"
               Caption         =   "Hora Desde"
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
               DataField       =   "horahasta"
               Caption         =   "Hora Hasta"
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
               DataField       =   "dias_permiso"
               Caption         =   "Dias P."
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
            BeginProperty Column07 
               DataField       =   "horas_permiso"
               Caption         =   "Horas P."
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
            BeginProperty Column08 
               DataField       =   "minutos_permiso"
               Caption         =   "Minutos P."
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
            BeginProperty Column09 
               DataField       =   "estado_codigo"
               Caption         =   "Estado"
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
            BeginProperty Column10 
               DataField       =   "TipoPermiso"
               Caption         =   "Tipo.Lic."
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
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1035.213
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   929.764
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   870.236
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   824.882
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   870.236
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoPermiso 
            Height          =   330
            Left            =   120
            Top             =   3360
            Width           =   12375
            _ExtentX        =   21828
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
            BackColor       =   12632319
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
            Caption         =   " <--- Registro de Permisos --->"
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
         Begin VB.Label Lbl06 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Permisos -->"
            DataField       =   "ARCHIVO"
            DataSource      =   "AdoPermiso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   9975
            TabIndex        =   127
            Top             =   555
            Width           =   1050
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Ver Permiso -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9720
            TabIndex        =   126
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.PictureBox fraDatos 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Height          =   8055
         Left            =   120
         ScaleHeight     =   7995
         ScaleWidth      =   12555
         TabIndex        =   52
         Top             =   1320
         Width           =   12615
         Begin VB.TextBox txt_obs 
            BackColor       =   &H00FFFFFF&
            DataField       =   "observaciones"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   600
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   179
            Top             =   3600
            Width           =   11565
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   5265
            TabIndex        =   94
            Top             =   1935
            Width           =   255
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Empresa"
            ForeColor       =   &H00800000&
            Height          =   600
            Left            =   9420
            TabIndex        =   92
            Top             =   200
            Width           =   980
            Begin VB.Label lblActivo 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               DataField       =   "sigla_emprea"
               DataMember      =   "sigla_emprea"
               DataSource      =   "Ado_datos"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   120
               TabIndex        =   93
               Top             =   200
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000040&
            Height          =   1680
            Left            =   120
            TabIndex        =   80
            Top             =   0
            Width           =   12390
            Begin VB.TextBox dtc_codigo3 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               DataField       =   "puesto_codigo"
               DataSource      =   "Ado_datos"
               Enabled         =   0   'False
               Height          =   315
               Left            =   6960
               TabIndex        =   185
               Top             =   720
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox dtc_cargo 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               DataField       =   "cargo_codigo"
               DataSource      =   "Ado_datos"
               Enabled         =   0   'False
               Height          =   315
               Left            =   -120
               TabIndex        =   184
               Top             =   1680
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox dtc_desc3 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               DataField       =   "puesto_descripcion"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   4545
               Locked          =   -1  'True
               TabIndex        =   183
               Top             =   1215
               Width           =   4815
            End
            Begin VB.PictureBox Img_Foto 
               AutoRedraw      =   -1  'True
               Height          =   1395
               Left            =   10335
               ScaleHeight     =   1335
               ScaleWidth      =   1815
               TabIndex        =   82
               Top             =   120
               Width           =   1875
               Begin VB.Image Image2 
                  Height          =   1335
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   195
               Left            =   8625
               TabIndex        =   81
               Top             =   1320
               Width           =   255
            End
            Begin MSDataListLib.DataCombo TxtProfesion 
               Bindings        =   "rw_ficha_rrhh.frx":B5A4
               DataField       =   "ocup_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   120
               TabIndex        =   83
               Top             =   1215
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   14737632
               ForeColor       =   0
               ListField       =   "ocup_descripcion"
               BoundColumn     =   "ocup_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_Ocup 
               Bindings        =   "rw_ficha_rrhh.frx":B5C0
               DataField       =   "ocup_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3720
               TabIndex        =   84
               Top             =   1200
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BackColor       =   -2147483637
               ListField       =   "ocup_codigo"
               BoundColumn     =   "ocup_codigo"
               Text            =   ""
            End
            Begin VB.Label Label18 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Label18"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   720
               Left            =   120
               TabIndex        =   182
               Top             =   120
               Visible         =   0   'False
               Width           =   9135
            End
            Begin VB.Label txtDenominacion 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "beneficiario_denominacion"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   178
               Top             =   480
               Width           =   5175
            End
            Begin VB.Label TxtNIT 
               BackColor       =   &H00404040&
               Caption         =   "-"
               DataField       =   "beneficiario_nit"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   9285
               TabIndex        =   91
               Top             =   480
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label DtcDepto3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "depto_sigla"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8355
               TabIndex        =   90
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Dtc_doc_id 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "tipodoc_codigo"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7425
               TabIndex        =   89
               Top             =   480
               Width           =   615
            End
            Begin VB.Label txtCodigo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "beneficiario_codigo"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5520
               TabIndex        =   88
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label LblInicial 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Label34"
               DataField       =   "ARCHIVO_FOTO"
               DataSource      =   "Ado_datos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   8325
               TabIndex        =   87
               Top             =   1305
               Width           =   1815
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   $"rw_ficha_rrhh.frx":B5DC
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
               Height          =   360
               Left            =   120
               TabIndex        =   86
               Top             =   210
               Width           =   9120
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Profesion u Ocupacion Principal                                  Puesto Actual del Funcionario"
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
               Left            =   120
               TabIndex        =   85
               Top             =   960
               Width           =   7035
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Lugar del Nacimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   975
            Left            =   120
            TabIndex        =   69
            Top             =   4000
            Width           =   12435
            Begin MSDataListLib.DataCombo Dtc_prov 
               Bindings        =   "rw_ficha_rrhh.frx":B66C
               DataField       =   "prov_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4980
               TabIndex        =   70
               Top             =   480
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "prov_descripcion"
               BoundColumn     =   "prov_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_munic 
               Bindings        =   "rw_ficha_rrhh.frx":B683
               DataField       =   "munic_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   8640
               TabIndex        =   71
               Top             =   480
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "munic_descripcion"
               BoundColumn     =   "munic_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_depto 
               Bindings        =   "rw_ficha_rrhh.frx":B69A
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2400
               TabIndex        =   72
               Top             =   480
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "depto_descripcion"
               BoundColumn     =   "depto_codigo"
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
            Begin MSDataListLib.DataCombo TxtNacionalidad 
               Bindings        =   "rw_ficha_rrhh.frx":B6B2
               DataField       =   "pais_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   75
               TabIndex        =   73
               Top             =   480
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "pais_descripcion"
               BoundColumn     =   "pais_codigo"
               Text            =   "DataCombo5"
            End
            Begin MSDataListLib.DataCombo DtcPaisCod 
               Bindings        =   "rw_ficha_rrhh.frx":B6C8
               DataField       =   "pais_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3000
               TabIndex        =   74
               Top             =   360
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BackColor       =   -2147483629
               ForeColor       =   16777215
               ListField       =   "pais_codigo"
               BoundColumn     =   "pais_codigo"
               Text            =   "DataCombo5"
            End
            Begin MSDataListLib.DataCombo DtcPaisSigla 
               Bindings        =   "rw_ficha_rrhh.frx":B6DE
               DataField       =   "pais_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2040
               TabIndex        =   75
               Top             =   360
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   -2147483629
               ForeColor       =   16777215
               ListField       =   "pais_cod_telefonico"
               BoundColumn     =   "pais_codigo"
               Text            =   "DataCombo5"
            End
            Begin MSDataListLib.DataCombo Dtc_depto_cod 
               Bindings        =   "rw_ficha_rrhh.frx":B6F4
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3840
               TabIndex        =   76
               Top             =   360
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483624
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
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
            Begin MSDataListLib.DataCombo Dtc_munic_cod 
               Bindings        =   "rw_ficha_rrhh.frx":B70C
               DataField       =   "munic_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   10800
               TabIndex        =   77
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483624
               ListField       =   "munic_codigo"
               BoundColumn     =   "munic_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_prov_cod 
               Bindings        =   "rw_ficha_rrhh.frx":B723
               DataField       =   "prov_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   7320
               TabIndex        =   78
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483624
               ListField       =   "prov_codigo"
               BoundColumn     =   "prov_codigo"
               Text            =   ""
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   $"rw_ficha_rrhh.frx":B73A
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
               Index           =   3
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   9360
            End
         End
         Begin VB.TextBox txt_sueldo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "beneficiario_haber_mensual"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            MaxLength       =   20
            TabIndex        =   68
            Top             =   2550
            Width           =   1440
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "bono_antiguedad"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3240
            MaxLength       =   20
            TabIndex        =   67
            Text            =   "0"
            Top             =   2550
            Width           =   1320
         End
         Begin VB.TextBox txt_otro 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "beneficiario_otro_mensual"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   66
            Top             =   2550
            Width           =   1440
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "DATOS DEL FONDO DE PENSIONES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   915
            Left            =   120
            TabIndex        =   60
            Top             =   6000
            Width           =   12420
            Begin VB.TextBox txt_afp 
               BackColor       =   &H00FFFFFF&
               DataField       =   "asegurado_codigo_afp"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   240
               MaxLength       =   15
               TabIndex        =   61
               Top             =   480
               Width           =   2205
            End
            Begin MSDataListLib.DataCombo dtc_afp_des 
               Bindings        =   "rw_ficha_rrhh.frx":B7E4
               DataField       =   "beneficiario_codigo_afp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2640
               TabIndex        =   62
               Top             =   480
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_afp 
               Bindings        =   "rw_ficha_rrhh.frx":B7F9
               DataField       =   "beneficiario_codigo_afp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   7560
               TabIndex        =   63
               Top             =   240
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   ""
            End
            Begin MSComCtl2.DTPicker DTP_FechaAfp 
               DataField       =   "fecha_asegurado_afp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   9000
               TabIndex        =   64
               Top             =   480
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   115736577
               CurrentDate     =   40179
               MinDate         =   2
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   $"rw_ficha_rrhh.frx":B80E
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
               Left            =   240
               TabIndex        =   65
               Top             =   240
               Width           =   10230
            End
         End
         Begin VB.Frame FraSS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "DATOS DEL SEGURO SOCIAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   915
            Left            =   120
            TabIndex        =   54
            Top             =   5040
            Width           =   12420
            Begin VB.TextBox txt_ss 
               BackColor       =   &H00FFFFFF&
               DataField       =   "asegurado_codigo_caja"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   240
               MaxLength       =   15
               TabIndex        =   55
               Top             =   465
               Width           =   2205
            End
            Begin MSDataListLib.DataCombo DtcSSEnt 
               Bindings        =   "rw_ficha_rrhh.frx":B8B3
               DataField       =   "beneficiario_codigo_seguro"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4440
               TabIndex        =   56
               Top             =   480
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcSS 
               Bindings        =   "rw_ficha_rrhh.frx":B8D2
               DataField       =   "beneficiario_codigo_seguro"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   7800
               TabIndex        =   57
               Top             =   480
               Visible         =   0   'False
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   ""
            End
            Begin MSComCtl2.DTPicker DTP_FechaSS 
               DataField       =   "fecha_asegurado_caja"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2640
               TabIndex        =   58
               Top             =   465
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   115736577
               CurrentDate     =   40179
               MinDate         =   2
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Matrícula del Asegurado     Fecha Asegurado     Nombre Entidad"
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
               Left            =   240
               TabIndex        =   59
               Top             =   225
               Width           =   5730
            End
         End
         Begin VB.Frame FraBco 
            BackColor       =   &H00E0E0E0&
            Caption         =   "CUENTA(S) BANCARIA(S) PERSONAL(ES)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1005
            Left            =   0
            TabIndex        =   53
            Top             =   6960
            Width           =   12540
            Begin VB.CommandButton BtnModificar2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cuenta Banco"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   11640
               Picture         =   "rw_ficha_rrhh.frx":B8F1
               Style           =   1  'Graphical
               TabIndex        =   194
               ToolTipText     =   "Carga Foto de la Persona"
               Top             =   240
               Width           =   855
            End
            Begin MSDataGridLib.DataGrid DtgCuentaBanco 
               Bindings        =   "rw_ficha_rrhh.frx":C2F3
               Height          =   750
               Left            =   120
               TabIndex        =   193
               Top             =   240
               Width           =   11505
               _ExtentX        =   20294
               _ExtentY        =   1323
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   16777152
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
               ColumnCount     =   9
               BeginProperty Column00 
                  DataField       =   "bco_codigo"
                  Caption         =   "Banco1"
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
                  DataField       =   "cta_tipo"
                  Caption         =   "Tipo.Cta.1"
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
                  DataField       =   "cta_codigo"
                  Caption         =   "Cta.Bancaria1"
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
                  DataField       =   "bco_codigo2"
                  Caption         =   "Banco2"
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
                  DataField       =   "cta_tipo2"
                  Caption         =   "Tipo.Cta.2"
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
                  DataField       =   "cta_codigo2"
                  Caption         =   "Cta.Bancaria2"
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
                  DataField       =   "bco_codigo3"
                  Caption         =   "Banco3"
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
               BeginProperty Column07 
                  DataField       =   "cta_tipo"
                  Caption         =   "Tipo.Cta.3"
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
               BeginProperty Column08 
                  DataField       =   "cta_codigo3"
                  Caption         =   "Cta.Bancaria3"
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
                     ColumnWidth     =   675.213
                  EndProperty
                  BeginProperty Column01 
                     Object.Visible         =   -1  'True
                     ColumnWidth     =   884.976
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   1769.953
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   675.213
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   854.929
                  EndProperty
                  BeginProperty Column05 
                     Object.Visible         =   -1  'True
                     ColumnWidth     =   1814.74
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   689.953
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   854.929
                  EndProperty
                  BeginProperty Column08 
                     ColumnWidth     =   2445.166
                  EndProperty
               EndProperty
            End
         End
         Begin MSComCtl2.DTPicker DTP_FechaNac 
            DataField       =   "beneficiario_fecha_nacimiento"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   9360
            TabIndex        =   95
            Top             =   2550
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   115736577
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTP_FechaExpira 
            DataField       =   "Fecha_expiracion"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4560
            TabIndex        =   96
            Top             =   1200
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   115736577
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo TDBtipoben 
            Bindings        =   "rw_ficha_rrhh.frx":C311
            DataField       =   "tipoben_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   97
            Top             =   1200
            Visible         =   0   'False
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "tipoben_descripcion"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo TxtTipo 
            Bindings        =   "rw_ficha_rrhh.frx":C32A
            DataField       =   "tipoben_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3720
            TabIndex        =   98
            Top             =   1260
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "tipoben_codigo"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcEstCivDes 
            Bindings        =   "rw_ficha_rrhh.frx":C343
            DataField       =   "estado_civil_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7680
            TabIndex        =   99
            Top             =   2550
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "estado_civil_descripcion"
            BoundColumn     =   "estado_civil_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo DtcEstCiv 
            Bindings        =   "rw_ficha_rrhh.frx":C35D
            DataField       =   "estado_civil_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8400
            TabIndex        =   100
            Top             =   2700
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "estado_civil_codigo"
            BoundColumn     =   "estado_civil_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "rw_ficha_rrhh.frx":C377
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   101
            Top             =   1920
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "rw_ficha_rrhh.frx":C390
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4440
            TabIndex        =   102
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "unidad_codigo"
            BoundColumn     =   "unidad_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "rw_ficha_rrhh.frx":C3A9
            DataField       =   "genero_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   11100
            TabIndex        =   103
            Top             =   2550
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "genero_descripcion"
            BoundColumn     =   "genero_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "rw_ficha_rrhh.frx":C3C2
            DataField       =   "genero_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   11760
            TabIndex        =   104
            Top             =   2760
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "genero_codigo"
            BoundColumn     =   "genero_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            DataField       =   "fecha_ingreso"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4680
            TabIndex        =   105
            Top             =   2550
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   115736577
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "rw_ficha_rrhh.frx":C3DB
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8640
            TabIndex        =   107
            Top             =   1920
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "unidad_descripcion_pla"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "rw_ficha_rrhh.frx":C3F4
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7680
            TabIndex        =   108
            Top             =   1935
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "unidad_codigo_pla"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "rw_ficha_rrhh.frx":C40D
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5520
            TabIndex        =   109
            Top             =   1920
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "planilla_descripcion"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "rw_ficha_rrhh.frx":C426
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   11760
            TabIndex        =   110
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "planilla_codigo"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin VB.TextBox TxtRenca 
            DataField       =   "reg_profesional"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            MaxLength       =   20
            TabIndex        =   106
            Top             =   1200
            Visible         =   0   'False
            Width           =   2280
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Obs.                                                                           Planilla a la que corresponde"
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
            Left            =   120
            TabIndex        =   180
            Top             =   3600
            Width           =   420
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "beneficiario_telefono_of"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   118
            Top             =   3105
            Width           =   2130
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono Celular Personal  Teléfono Corporativo        Domicilio Actual:"
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
            Left            =   120
            TabIndex        =   117
            Top             =   2880
            Width           =   6180
         End
         Begin VB.Label txt_file 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "beneficiario_nro_file"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6720
            TabIndex        =   116
            Top             =   1200
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.Label txt_item 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "beneficiario_item"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6480
            TabIndex        =   115
            Top             =   2550
            Width           =   1095
         End
         Begin VB.Label TxtDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            DataField       =   "beneficiario_domicilio_legal"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   4800
            TabIndex        =   114
            Top             =   3105
            Width           =   7695
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"rw_ficha_rrhh.frx":C440
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
            Left            =   120
            TabIndex        =   113
            Top             =   1665
            Width           =   11715
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   $"rw_ficha_rrhh.frx":C507
            DataField       =   "  "
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
            Left            =   120
            TabIndex        =   112
            Top             =   2310
            Width           =   11880
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "beneficiario_telefono_cel"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   111
            Top             =   3105
            Width           =   2250
         End
      End
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H00404040&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   12555
         TabIndex        =   44
         Top             =   660
         Width           =   12615
         Begin VB.CommandButton BtnAux2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Modificar Persona"
            Height          =   600
            Left            =   6480
            MaskColor       =   &H00FFFFFF&
            Picture         =   "rw_ficha_rrhh.frx":C598
            Style           =   1  'Graphical
            TabIndex        =   181
            ToolTipText     =   "Modificar Datos Personales"
            Top             =   30
            Width           =   1260
         End
         Begin VB.CommandButton BtnAprobar 
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3900
            Picture         =   "rw_ficha_rrhh.frx":CCB3
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Aprueba Registro"
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton BtnAñadir 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "rw_ficha_rrhh.frx":D4E9
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton BtnModificar 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1260
            Picture         =   "rw_ficha_rrhh.frx":DCA8
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton BtnEliminar 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":E5BD
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Anula Registro Activo"
            Top             =   30
            Width           =   1245
         End
         Begin VB.CommandButton cmd_mas 
            BackColor       =   &H80000015&
            Caption         =   "Mas Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5280
            Picture         =   "rw_ficha_rrhh.frx":ED09
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Carga Foto de la Persona"
            Top             =   30
            Width           =   1215
         End
         Begin VB.CommandButton CmdDesapr 
            BackColor       =   &H0080C0FF&
            Caption         =   "Desapr"
            Height          =   600
            Left            =   3000
            Picture         =   "rw_ficha_rrhh.frx":F293
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Aprueba Registro"
            Top             =   60
            Width           =   740
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " I. DATOS PERSONALES GENERALES"
            DataSource      =   "adoLista"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   735
            Left            =   7680
            TabIndex        =   51
            Top             =   0
            Width           =   4935
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "EXPERIENCIA LABORAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4215
         Left            =   -74880
         TabIndex        =   36
         Top             =   5160
         Width           =   12615
         Begin VB.CommandButton Command3 
            BackColor       =   &H80000018&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":F49D
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H80000018&
            Caption         =   "AnuLar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":FA27
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H80000018&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":10429
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H80000018&
            Caption         =   "Modif."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":109B3
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Height          =   690
            Left            =   11595
            TabIndex        =   37
            Top             =   120
            Width           =   615
            Begin VB.Image Img_DocRespaldo 
               Height          =   540
               Left            =   15
               Picture         =   "rw_ficha_rrhh.frx":10F3D
               Top             =   105
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtgLaborales 
            Bindings        =   "rw_ficha_rrhh.frx":112C5
            Height          =   2865
            Left            =   165
            TabIndex        =   42
            Top             =   840
            Width           =   12330
            _ExtentX        =   21749
            _ExtentY        =   5054
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
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
            Caption         =   "EXPERIENCIA LABORAL (Empresas o Instituciones donde Trabajó)"
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "codigo_experiencia"
               Caption         =   "Nro."
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
               DataField       =   "fecha_inicio"
               Caption         =   "Fecha.Inicio"
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
               DataField       =   "fecha_fin"
               Caption         =   "Fecha Fin"
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
               DataField       =   "denominacion_institucion"
               Caption         =   "Institución Donde Trabajó"
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
               DataField       =   "cargo"
               Caption         =   "Cargo que Ocupó"
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
               DataField       =   "Tiempo_Meses"
               Caption         =   "Duracion"
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
               DataField       =   "tiempo_dmy"
               Caption         =   "Tiempo"
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
            BeginProperty Column07 
               DataField       =   "tipo_institucion"
               Caption         =   "Tipo Institucion"
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
            BeginProperty Column08 
               DataField       =   "funcion_general"
               Caption         =   "Función Principal"
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
            BeginProperty Column09 
               DataField       =   "pais"
               Caption         =   "País"
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
            BeginProperty Column10 
               DataField       =   "ciudad"
               Caption         =   "Ciudad"
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
            BeginProperty Column11 
               DataField       =   "presento_documento"
               Caption         =   "Docs"
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
                  Object.Visible         =   0   'False
                  WrapText        =   -1  'True
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   2025.071
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column07 
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_Laborales 
            Height          =   330
            Left            =   165
            Top             =   3720
            Width           =   12330
            _ExtentX        =   21749
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
            BackColor       =   -2147483624
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
            Caption         =   " <--- Empresas o Instituciones donde Trabajó  --->"
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
         Begin VB.Label LblResp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Hoja de Vida -->"
            DataField       =   "ARCHIVO_RESPALDO"
            DataSource      =   "adoLista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   10035
            TabIndex        =   43
            Top             =   540
            Width           =   1425
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESTUDIOS REALIZADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3975
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   12615
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":112E1
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00FFC0C0&
            Caption         =   "AnuLar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":1186B
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":1226D
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Modif."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":127F7
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Height          =   690
            Left            =   11595
            TabIndex        =   29
            Top             =   120
            Width           =   615
            Begin VB.Image Image1 
               Height          =   540
               Left            =   20
               Picture         =   "rw_ficha_rrhh.frx":12D81
               Top             =   100
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtgEducacionales 
            Bindings        =   "rw_ficha_rrhh.frx":13109
            Height          =   2625
            Left            =   165
            TabIndex        =   34
            Top             =   840
            Width           =   12330
            _ExtentX        =   21749
            _ExtentY        =   4630
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BackColor       =   16761024
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
            Caption         =   "ESTUDIOS REALIZADOS (Datos Educacionales)"
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "Codigo_Educacion"
               Caption         =   "Nro."
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
               DataField       =   "fecha_inicio"
               Caption         =   "Fecha Inicio"
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
               DataField       =   "Fecha_Fin"
               Caption         =   "Fecha Fin"
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
               DataField       =   "Carrera_Curso"
               Caption         =   "Carrera/Curso"
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
               DataField       =   "centro_educativo"
               Caption         =   "Centro Educativo"
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
               DataField       =   "duracion_tiempo"
               Caption         =   "Duracion"
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
               DataField       =   "tiempo_dmy"
               Caption         =   "Tiempo"
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
            BeginProperty Column07 
               DataField       =   "estado_codigo"
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
            BeginProperty Column08 
               DataField       =   "nivel_educacional"
               Caption         =   "Nivel Educ."
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
            BeginProperty Column09 
               DataField       =   "pais"
               Caption         =   "Pais"
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
            BeginProperty Column10 
               DataField       =   "ciudad"
               Caption         =   "Ciudad"
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
            BeginProperty Column11 
               DataField       =   "PRESENTO_DOCUMENTO"
               Caption         =   "Docs"
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
            BeginProperty Column12 
               DataField       =   "titulo_obtenido"
               Caption         =   "Titulo Obtenido"
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
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2369.764
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1769.953
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_Educacionales 
            Height          =   330
            Left            =   165
            Top             =   3480
            Width           =   12330
            _ExtentX        =   21749
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
            BackColor       =   16761024
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
            Caption         =   " <--- Estudios Realizados --->"
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Hoja de Vida -->"
            DataField       =   "ARCHIVO_HOJAVIDA"
            DataSource      =   "adoLista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   10080
            TabIndex        =   35
            Top             =   555
            Width           =   1425
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FINIQUITOS, QUINQUENIOS Y OTROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4215
         Left            =   -74880
         TabIndex        =   20
         Top             =   5160
         Width           =   12615
         Begin VB.CommandButton BtnImprimir4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Left            =   3480
            Picture         =   "rw_ficha_rrhh.frx":13129
            Style           =   1  'Graphical
            TabIndex        =   192
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00C0E0FF&
            Caption         =   "AnuLar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":136B3
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":140B5
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Aprueba Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Modif."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":1463F
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":14BC9
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   11595
            TabIndex        =   21
            Top             =   120
            Width           =   615
            Begin VB.Image Image3 
               Height          =   540
               Left            =   20
               Picture         =   "rw_ficha_rrhh.frx":15153
               Top             =   80
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtgLiquidacion 
            Bindings        =   "rw_ficha_rrhh.frx":154DB
            Height          =   2865
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   12330
            _ExtentX        =   21749
            _ExtentY        =   5054
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            Caption         =   "FINIQUITOS - QUINQUENIOS - OTRAS LIQUIDACIONES"
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "fecha_ingreso"
               Caption         =   "Fecha.Ingreso"
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
               DataField       =   "fecha_retiro"
               Caption         =   "Fecha. Retiro"
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
               DataField       =   "ges_gestion_ini"
               Caption         =   "Gestion.Inicial"
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
               DataField       =   "tipo_memo"
               Caption         =   "Motivo"
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
               DataField       =   "Fecha_Liquidacion"
               Caption         =   "Fecha Liquidacion"
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
               DataField       =   "Monto_Total"
               Caption         =   "Monto Liquidacion"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "id_liquidacion"
               Caption         =   "Nro.Liquidacion"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "cta_codigo"
               Caption         =   "Cta. Bancaria"
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
            BeginProperty Column08 
               DataField       =   "beneficiario_codigo"
               Caption         =   "Beneficiario"
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
            BeginProperty Column09 
               DataField       =   "estado_codigo"
               Caption         =   "Estado"
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
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1679.811
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   780.095
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoLiquidacion 
            Height          =   330
            Left            =   120
            Top             =   3720
            Width           =   12330
            _ExtentX        =   21749
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
            Caption         =   " <--- Finiquitos, Quinquenios y Otros --->"
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
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Ver Liquidación -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   9840
            TabIndex        =   27
            Top             =   240
            Width           =   1605
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CONTRATOS CON LA INSTITUCION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3975
         Left            =   -74880
         TabIndex        =   12
         Top             =   1200
         Width           =   12615
         Begin VB.CommandButton Command15 
            BackColor       =   &H00FFFFC0&
            Caption         =   "AnuLar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1800
            Picture         =   "rw_ficha_rrhh.frx":154F8
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Anula Registro Activo"
            Top             =   260
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Aprobar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            Picture         =   "rw_ficha_rrhh.frx":15EFA
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   260
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Modif."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   960
            Picture         =   "rw_ficha_rrhh.frx":16484
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   260
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "rw_ficha_rrhh.frx":16A0E
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Nuevo Registro"
            Top             =   260
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Frame Frame21 
            BackColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   11595
            TabIndex        =   13
            Top             =   120
            Width           =   615
            Begin VB.Image Image4 
               Height          =   540
               Left            =   20
               Picture         =   "rw_ficha_rrhh.frx":16F98
               Top             =   80
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtG_Contrato 
            Bindings        =   "rw_ficha_rrhh.frx":17320
            Height          =   2745
            Left            =   165
            TabIndex        =   18
            Top             =   840
            Width           =   12330
            _ExtentX        =   21749
            _ExtentY        =   4842
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BackColor       =   16777152
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
            Caption         =   "REGISTRO DE CONTRATOS - ADENDAS - MEMORANDAS DESIGNACION"
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "fecha_inicio"
               Caption         =   "Fecha.Inicio"
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
               DataField       =   "fecha_fin"
               Caption         =   "Fecha.Finaliz."
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
               DataField       =   "codigo_beneficiario"
               Caption         =   "Trabajador"
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
               DataField       =   "unidad_codigo"
               Caption         =   "Unidad"
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
               DataField       =   "cargo_codigo"
               Caption         =   "Cargo"
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
               DataField       =   "puesto_codigo"
               Caption         =   "Puesto"
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
               DataField       =   "fte_codigo"
               Caption         =   "Fte.Fin."
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
            BeginProperty Column07 
               DataField       =   "org_codigo"
               Caption         =   "Financiador"
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
            BeginProperty Column08 
               DataField       =   "pro_codigo"
               Caption         =   "Proyecto"
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
            BeginProperty Column09 
               DataField       =   "codigo_contrato"
               Caption         =   "Cod.Contrato"
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
            BeginProperty Column10 
               DataField       =   "estado_contrato"
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
            BeginProperty Column11 
               DataField       =   "fecha_firma"
               Caption         =   "Fecha.Firma"
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
            BeginProperty Column12 
               DataField       =   "fechas_confirmado"
               Caption         =   "Vigente"
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
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   555.024
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   884.976
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1260.284
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_Contrato 
            Height          =   330
            Left            =   165
            Top             =   3600
            Width           =   12330
            _ExtentX        =   21749
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
            BackColor       =   16777152
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
            Caption         =   " <--- Contratos con la Institución --->"
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
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Ver Contrato -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   10140
            TabIndex        =   19
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.PictureBox FraGrabarCancelar 
         BackColor       =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   12555
         TabIndex        =   169
         Top             =   660
         Width           =   12615
         Begin VB.CommandButton BtnGrabar 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3240
            Picture         =   "rw_ficha_rrhh.frx":1733B
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton BtnCancelar 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   4620
            MaskColor       =   &H00000000&
            Picture         =   "rw_ficha_rrhh.frx":17B11
            Style           =   1  'Graphical
            TabIndex        =   170
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   1485
         End
         Begin VB.Label lbl_titulo2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IDENTIFICACION DEL CLIENTE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   360
            Left            =   8250
            TabIndex        =   172
            Top             =   300
            Visible         =   0   'False
            Width           =   525
         End
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " II. - CONTROL DE ASISTENCIA "
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   390
         Left            =   -74880
         TabIndex        =   10
         Top             =   660
         Width           =   12645
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " III. - PERMISOS - VACACIONES "
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   390
         Left            =   -74880
         TabIndex        =   176
         Top             =   660
         Width           =   12540
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IV. - MOVILIDAD DE PERSONAL "
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   495
         Left            =   -74880
         TabIndex        =   175
         Top             =   660
         Width           =   12615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " V. ANTECEDENTES PROFESIONALES y DE TRABAJO "
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   495
         Left            =   -74880
         TabIndex        =   174
         Top             =   720
         Width           =   12615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VI. SITUACION DENTRO DE LA INSTITUCION "
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   495
         Left            =   -74880
         TabIndex        =   173
         Top             =   720
         Width           =   12615
      End
   End
   Begin VB.PictureBox fra_cabecera 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1065
      ScaleWidth      =   6465
      TabIndex        =   4
      Top             =   0
      Width           =   6495
      Begin VB.PictureBox btncumple 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -30
         Picture         =   "rw_ficha_rrhh.frx":183FD
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   187
         ToolTipText     =   "Busca Registros "
         Top             =   480
         Width           =   1400
      End
      Begin VB.PictureBox Command19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1290
         Picture         =   "rw_ficha_rrhh.frx":19056
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   191
         ToolTipText     =   "Busca Registros "
         Top             =   480
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2680
         Picture         =   "rw_ficha_rrhh.frx":19A70
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   190
         ToolTipText     =   "Busca Registros "
         Top             =   480
         Width           =   1215
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "rw_ficha_rrhh.frx":1A225
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   188
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   480
         Width           =   1245
      End
      Begin VB.PictureBox btnimprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3880
         Picture         =   "rw_ficha_rrhh.frx":1A9E7
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   189
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   480
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         Caption         =   "FICHA PERSONAL"
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
         Height          =   465
         Left            =   15
         TabIndex        =   5
         Top             =   40
         Width           =   6405
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LISTADOS"
      ForeColor       =   &H00C00000&
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6495
      Begin MSDataListLib.DataCombo dtc_buscar_desc 
         Bindings        =   "rw_ficha_rrhh.frx":1B2B4
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
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
      Begin MSDataListLib.DataCombo dtc_buscar_ci 
         Bindings        =   "rw_ficha_rrhh.frx":1B2D1
         DataField       =   "beneficiario_codigo"
         Height          =   315
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
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
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   4080
         TabIndex        =   3
         Top             =   7575
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Vigentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   1440
         TabIndex        =   2
         Top             =   7560
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   60
         Top             =   7500
         Width           =   6375
         _ExtentX        =   11245
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
         BackColor       =   16777152
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "rw_ficha_rrhh.frx":1B2EE
         Height          =   6855
         Left            =   60
         TabIndex        =   0
         Top             =   600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Apellidos y Nombres"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Doc. Identidad"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
            DataField       =   "tipoben_codigo"
            Caption         =   "Tipo_Benef"
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
            DataField       =   "beneficiario_telefono_fijo"
            Caption         =   "Telefono"
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
            DataField       =   "munic_codigo"
            Caption         =   "Procedencia"
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
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   4919.811
            EndProperty
         EndProperty
      End
      Begin VB.OLE OLE1 
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label52 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc AdoTip_ben 
      Height          =   330
      Left            =   10800
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AdoTip_ben"
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
   Begin MSAdodcLib.Adodc Ado_Depto 
      Height          =   330
      Left            =   120
      Top             =   9000
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_Depto"
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
   Begin MSAdodcLib.Adodc Ado_prov 
      Height          =   330
      Left            =   0
      Top             =   9360
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Prov"
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
   Begin MSAdodcLib.Adodc Ado_Muni 
      Height          =   330
      Left            =   2160
      Top             =   9000
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Muni"
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
      Left            =   2160
      Top             =   9360
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin MSAdodcLib.Adodc Ado_Depto2 
      Height          =   330
      Left            =   4320
      Top             =   9000
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Depto"
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
   Begin MSAdodcLib.Adodc Ado_prov2 
      Height          =   330
      Left            =   4320
      Top             =   9360
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Prov"
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
   Begin MSAdodcLib.Adodc Ado_Muni2 
      Height          =   330
      Left            =   6480
      Top             =   9000
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Muni"
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
   Begin MSAdodcLib.Adodc Ado_CtaPersonal 
      Height          =   330
      Left            =   6360
      Top             =   9360
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "Ado_CtaPersonal"
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
   Begin MSAdodcLib.Adodc Ado_TipoDocId 
      Height          =   330
      Left            =   8640
      Top             =   9000
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_TipoDocId"
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
   Begin MSAdodcLib.Adodc Ado_Depto3 
      Height          =   330
      Left            =   8640
      Top             =   9360
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Depto3"
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   10800
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_datos5"
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
   Begin MSAdodcLib.Adodc AdoNivelEducacional 
      Height          =   330
      Left            =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AdoNivelEducacional"
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
   Begin MSAdodcLib.Adodc Ado_TipoInstitucion 
      Height          =   330
      Left            =   2160
      Top             =   9720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_TipoInstitucion"
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
   Begin MSAdodcLib.Adodc Ado_Benef_seguro 
      Height          =   330
      Left            =   4320
      Top             =   9720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_Benef_seguro"
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   6480
      Top             =   9720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AdoCta"
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
   Begin MSAdodcLib.Adodc Ado_Ocupacion 
      Height          =   330
      Left            =   12960
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_Ocupacion"
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
   Begin MSAdodcLib.Adodc Ado_Benef_Afp 
      Height          =   330
      Left            =   12960
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_Benef_Afp"
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   8640
      Top             =   9720
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos6"
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
   Begin MSAdodcLib.Adodc AdoPuestoOrg 
      Height          =   330
      Left            =   2160
      Top             =   10080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoPuestoOrg"
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
   Begin MSAdodcLib.Adodc AdoOrg 
      Height          =   330
      Left            =   10800
      Top             =   9720
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoOrg"
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
   Begin MSAdodcLib.Adodc AdoPry 
      Height          =   330
      Left            =   12960
      Top             =   9720
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoPry"
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
   Begin MSAdodcLib.Adodc AdoCargo 
      Height          =   330
      Left            =   0
      Top             =   10080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoCargo"
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
   Begin MSAdodcLib.Adodc AdoFuente 
      Height          =   330
      Left            =   4320
      Top             =   10080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoFuente"
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
   Begin MSAdodcLib.Adodc AdoEstCivil 
      Height          =   330
      Left            =   6480
      Top             =   10080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoEstCivil"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   15120
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoPais 
      Height          =   330
      Left            =   8640
      Top             =   10080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoPais"
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
   Begin MSAdodcLib.Adodc AdoCalendario 
      Height          =   330
      Left            =   10800
      Top             =   10080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoCalendario"
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
   Begin MSAdodcLib.Adodc AdoHorarios 
      Height          =   330
      Left            =   12960
      Top             =   10080
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoHorarios"
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
   Begin Crystal.CrystalReport CR02 
      Left            =   15720
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR03 
      Left            =   16320
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos_busq 
      Height          =   330
      Left            =   15240
      Top             =   10080
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
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
      Caption         =   "Ado_datos_busq"
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   0
      Top             =   10440
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos7"
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   2160
      Top             =   10440
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos4"
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
   Begin MSAdodcLib.Adodc adoafp 
      Height          =   330
      Left            =   4320
      Top             =   10440
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "adoafp"
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
   Begin Crystal.CrystalReport CR04 
      Left            =   16920
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "rw_ficha_rrhh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento de Beneficiarios
Option Explicit
Dim rstbeneficiario As New ADODB.Recordset
Dim rst_ben, rsNada, rs_cumple As New ADODB.Recordset

Dim rs_datos4, rs_datos1, rs_datos2, rs_datos3 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rsauxiliar As New ADODB.Recordset
Dim rstbeneaux As New ADODB.Recordset
Dim rs_Depto, rs_Prov, rs_Muni, rs_comunid As New ADODB.Recordset
Dim rs_Depto2, rs_Prov2, rs_Muni2, rs_comunid2 As New ADODB.Recordset
Dim rs_Depto3 As New ADODB.Recordset
Dim rs_TipoDocId, rs_RepLegal As New ADODB.Recordset
Dim rs_Permisos, rs_Permiso_detalle  As New ADODB.Recordset
Dim rs_nivel_educacional As New ADODB.Recordset
Dim rs_laborales As New ADODB.Recordset
Dim rs_tipoInstitucion As New ADODB.Recordset
Dim rs_beneficiario As New ADODB.Recordset
Dim rs_CTA_BCO As New ADODB.Recordset
Dim rs_ocupacion As New ADODB.Recordset
Dim rs_beneficiario_Afp As New ADODB.Recordset
Dim rs_Dependiente, rs_pais As New ADODB.Recordset
Dim rs_contrato, rs_Puesto, rs_UNIDAD, rs_Org, rs_Pry, rs_CARGO As New ADODB.Recordset
Dim rs_correlativo, rsfuente As New ADODB.Recordset
Dim rs_EstCivil, rs_liquidacion As New ADODB.Recordset
Dim rs_Asistencia As New ADODB.Recordset
Dim rs_vacaciones_prog As New ADODB.Recordset
Dim rs_calendario, rs_calendario2 As New ADODB.Recordset
Dim rs_HORARIOS As New ADODB.Recordset
Dim rs_movilidad As New ADODB.Recordset
Dim VAR_COD2 As String

Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset
Dim rs_aux14 As New ADODB.Recordset
Dim rs_aux17 As New ADODB.Recordset
Dim rs_aux18 As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim permisos, totalminutos As Integer
Dim calretrasos As Double
Dim CAMPOS As ADODB.Field
Dim LISTAC As String

Dim rs_CtaPersonal As New ADODB.Recordset

Dim rs_datos_educacionales As New ADODB.Recordset
Dim rs_Puesto_Org As New ADODB.Recordset
Dim rstafp As New ADODB.Recordset

'BUSQUEDA
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim queryinicial As String
'OTROS
Dim SW As Boolean

Dim total, total2 As Double
Dim VAR_MES As Integer
Dim VAR_REG As Integer

Dim SQL_FOR As String
Dim CORREL, accion As Integer
Dim V_TIPO, V_TDOC As String
Dim sino As String
Dim marca1 As String
Dim NombreCarpeta, e As String

Dim imag2 As Long
Dim VARB, VARBD, VARG, VARS, VARU, VARP, varCat, VAR10, VAR11, VAR12, VAR13, VAR14, VAR15 As String
Dim VARPU, VARCAN, VARPT As Double
Dim sqlAux As String

Dim VAR_VAL As String
Private Sub Pagos(Unidad, formulario, org_codigo, solicitud_codigo, justificacion, observaciones, beneficiario_codigo, concepto_pago, obs_fo_gastos_detalle, mon_pago As String)
  'PAGOS
    'WWWWWWWWWWWWWW
    Dim VAR_CMPBTE As Integer
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        'VAR_COD4 = parametro    'UNIDAD
        'VAR_SOL = Ado_datos.Recordset!beneficiario_codigo  '
        'tipo_formulario TERCER PARAMETRO
        'org_codigo CUARTO PARAMETRO
        'ini generación de correlativo
        
        If Unidad = "DRRHH" Then
        Set rs_aux5 = New ADODB.Recordset
        If rs_aux5.State = 1 Then rs_aux5.Close
        rs_aux5.Open "select * from ro_rrhh_adjudica_personas where beneficiario_codigo = '" & beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic
        solicitud_codigo = rs_aux5!solicitud_codigo
        Unidad = rs_aux5!unidad_codigo
        End If
        
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from fc_organismo_financiamiento where org_codigo = '" & org_codigo & "'", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
           rs_aux2!correlativo_gasto = rs_aux2!correlativo_gasto + 1
           VAR_CMPBTE = rs_aux2!correlativo_gasto
           rs_aux2.Update
        End If
        
        'WWWWWWWWWWWWWWW
        'correlv = Ado_datos.Recordset!venta_codigo
        'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
        Set rs_aux3 = New ADODB.Recordset
        If rs_aux3.State = 1 Then rs_aux3.Close
        rs_aux3.Open "select * from  fo_gastos_cabecera where unidad_codigo = '" & Unidad & "' AND solicitud_codigo = " & solicitud_codigo & " ", db, adOpenKeyset, adLockOptimistic
        If rs_aux3.RecordCount = 0 Then
            rs_aux3.AddNew
            rs_aux3!ges_gestion = Year(Date) 'glGestion     'Year(Date)
            rs_aux3!org_codigo = org_codigo
            rs_aux3!gasto_codigo = VAR_CMPBTE
            rs_aux3!tipo_comp = "DEV"
            rs_aux3!gasto_codigo_anterior = VAR_CMPBTE
            rs_aux3!unidad_codigo = Unidad
            rs_aux3!solicitud_codigo = solicitud_codigo
            rs_aux3!tipo_formulario = formulario     'GC_TIPO_SOLICITUD
            rs_aux3!pago_codigo = rs_aux3.RecordCount + 1
            
            rs_aux3!proceso_codigo = "FIN"
            rs_aux3!subproceso_codigo = "FIN-03"
            rs_aux3!etapa_codigo = "FIN-03-03"
            rs_aux3!clasif_codigo = "ADM"
            rs_aux3!doc_codigo = "R-111"
            rs_aux3!doc_numero = 0
            rs_aux3!poa_codigo = "4.2.3"
            
            rs_aux3!fecha_egreso = Date
            rs_aux3!tipo_moneda = "BOB"
            rs_aux3!da_codigo = "1.1"
            
            rs_aux3!fte_codigo = "10"   'DEVISAR DE LA TABLA fc_organismo_financiamiento
            rs_aux3!monto_Bolivianos = 0
            rs_aux3!monto_dolares = 0
            rs_aux3!liquido_pagar = 0
            rs_aux3!monto_Bolivianos_pag = 0
            rs_aux3!monto_dolares_pag = 0
            rs_aux3!Deducciones = 0
            rs_aux3!fecha_autorizacion = Date
            rs_aux3!justificacion = justificacion
            rs_aux3!es_base = "S"
            
            rs_aux3!CODIGO_GRUPO = VAR_CMPBTE   'rs_aux3!pago_codigo
            rs_aux3!NUMERO_PAGO = 1
            rs_aux3!observaciones = observaciones
            
            'rs_aux3!edif_codigo = VAR_PROY2
            'rs_aux3!beneficiario_codigo = VAR_BENEF
            'rs_aux3!solicitud_tipo = "10"
            'rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
            
            rs_aux3!estado_devengado = "APR"
            rs_aux3!estado_pagado = "REG"
            rs_aux3!estado_contabilidad = "REG"
            
            rs_aux3!estado_codigo = "REG"
            rs_aux3!usr_codigo = glusuario
            rs_aux3!fecha_registro = Date
            rs_aux3!usr_codigo_aprueba = glusuario
            rs_aux3!fecha_aprueba = Date
            rs_aux3.Update
            
            'DETALLE Carga fo_gastos_detalle
            
            Set rstdestino = New ADODB.Recordset
            If rstdestino.State = 1 Then rstdestino.Close
            rstdestino.Open "select * from fo_gastos_detalle where org_codigo = '" & org_codigo & "' AND gasto_codigo= " & VAR_CMPBTE & "  ", db, adOpenKeyset, adLockBatchOptimistic
            If rstdestino.RecordCount > 0 Then
            End If

            Set rs_aux4 = New ADODB.Recordset
            If rs_aux4.State = 1 Then rs_aux4.Close
            rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & Unidad & "' AND solicitud_codigo = " & solicitud_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux4.RecordCount > 0 Then
               VAR_REG = 1
               rs_aux4.MoveFirst
               Dim bien_total_compra As Double
               Dim par_codigo As String
               If Unidad = "DRRHH" Then
               bien_total_compra = mon_pago
               par_codigo = ""
               Else
               bien_total_compra = rs_aux4!bien_total_compra
               par_codigo = rs_aux4!par_codigo
               End If
               While Not rs_aux4.EOF
               '     db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
               '     "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
               '     rs_aux4.MoveNext
                    db.Execute "INSERT INTO fo_gastos_detalle (ges_gestion, org_codigo, gasto_codigo, gasto_codigo_detalle, par_codigo, pro_codigo, codigo_beneficiario, concepto_pago, monto_total, monto_dolares_dev, tipo_cambio_dev, monto_Bolivianos, monto_Dolares, saldo_bolivianos, tipo_cambio, Porcentaje, deducciones, fecha_pago, depto_codigo, estado_aprobacion, fecha_autorizacion, Observacion, codigo_dev, usr_usuario, fecha_registro, hora_registro,  estado_conciliacion, codigo_poa " & _
                    "VALUES ('" & glGestion & "','" & org_codigo & "', " & VAR_CMPBTE & ", '" & VAR_REG & "', " & par_codigo & ", '8', '" & beneficiario_codigo & "', '" & concepto_pago & "', " & bien_total_compra & ", " & bien_total_compra * GlTipoCambioOficial & ", " & GlTipoCambioOficial & ", " & bien_total_compra & ", " & bien_total_compra * GlTipoCambioOficial & " , '0', " & GlTipoCambioOficial & ", '100', '0', '2', 'REG', '" & Date & "', '" & obs_fo_gastos_detalle & "', " & VAR_CMPBTE & ", '" & glusuario & "', '" & Date & "', 'REG', '4.2.3')"
                   VAR_REG = VAR_REG + 1
               '     'cta_codigo, cheque_o_trf, numero_cheque_trf, cta_codigo_destino, cheque_o_trf_destino, numero_cheque_trf_destino,
               '     'Fecha_Aprobacion_tesoreria, fecha_impresion_cheque, banco_destino,
               Wend
            End If
            'If rstdestino.State = 1 Then rstdestino.Close
 
        End If
        'db.Execute "update ao_compra_planilla_pagos set estado_codigo = 'APR' where compra_codigo = " & Ado_detalle2.Recordset!compra_codigo & " and adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & " and pago_codigo=" & Ado_detalle3.Recordset!pago_codigo & "   "
        'Call ABRIR_TABLA_DET
    Else
        MsgBox "NO se puede APROBAR un registro Anulado o previamente Aprobado. ", vbExclamation, "Atención!"
    End If
        'WWWWWWWWWW
End Sub


Private Sub Ado_VacacionesProg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If swnuevo = 2 Then
'    If rs_contrato!estAdo_Memo = "REG" Then
''        TxtAprob.ForeColor = &H80&
'        CmdAdd4.Visible = True
'        CmdMod4.Visible = True
'        CmdElim4.Visible = False
'        CmdApr4.Visible = True
'    Else
''        TxtAprob.ForeColor = &H4000&
'        CmdAdd4.Visible = True
'        CmdMod4.Visible = False
'        CmdElim4.Visible = False
'        CmdApr4.Visible = False
'    End If
  Else
'    If rs_contrato!estAdo_Memo = "REG" Then
''        TxtAprob.ForeColor = &H80&
''        lblARCH.ForeColor = &H80&
'    Else
''        TxtAprob.ForeColor = &H4000&
''        lblARCH.ForeColor = &H4000&
'    End If
  End If
  If Ado_VacacionesProg.Recordset.RecordCount > 0 Then
  
      If Ado_VacacionesProg.Recordset!estado_codigo = "REG" Then
         frm_ao_Vacacion_Prog.txtEstado.ForeColor = &H4000&
    '        CmdAdd4.Visible = True
    '        CmdMod4.Visible = True
    '        CmdElim4.Visible = False
    '        CmdApr4.Visible = True
      Else
         frm_ao_Vacacion_Prog.txtEstado.ForeColor = &H80&
    '        CmdAdd4.Visible = True '&H000000C0&
    '        CmdMod4.Visible = False
    '        CmdElim4.Visible = False
    '        CmdApr4.Visible = False
        End If
    
'       If Ado_VacacionesProg.Recordset("ARCHIVO") = "Cargar_Archivo" Then
'            frm_ao_Vacacion_Prog.lblARCH.ForeColor = &HC0&
'            LblCto.ForeColor = &HC0&
'        Else
'            frm_ao_Vacacion_Prog.lblARCH.ForeColor = &H8000&
'            LblCto.ForeColor = &H8000&
'        End If
  End If
End Sub



'Private Sub Ado_VacacionesProg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
''    If Ado_VacacionesProg.Recordset.RecordCount > 0 Then
''        Dim codig0 As Integer
''        codig0 = Ado_VacacionesProg.Recordset!Codigo_Educacion
''    End If
'End Sub

Private Sub AdoMovilidad_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  If AdoMovilidad.Recordset.RecordCount > 0 Then
'
'      If AdoMovilidad.Recordset!estado_codigo = "REG" Then
'         ro_Personal_Liquidacion.TxtAprob.ForeColor = &H4000&
'    '        CmdAdd4.Visible = True
'    '        CmdMod4.Visible = True
'    '        CmdElim4.Visible = False
'    '        CmdApr4.Visible = True
'      Else
'         ro_Personal_Liquidacion.TxtAprob.ForeColor = &H80&
'    '        CmdAdd4.Visible = True '&H000000C0&
'    '        CmdMod4.Visible = False
'    '        CmdElim4.Visible = False
'    '        CmdApr4.Visible = False
'        End If
'
'       If AdoMovilidad.Recordset("ARCHIVO") = "Cargar_Archivo" Then
'            ro_Personal_Liquidacion.lblARCH.ForeColor = &HC0&
'            LblLiq.ForeColor = &HC0&
'        Else
'            ro_Personal_Liquidacion.lblARCH.ForeColor = &H8000&
'            LblLiq.ForeColor = &H8000&
'        End If
'  End If

End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'If pRecordset.EOF Or pRecordset.BOF Then
  If Ado_datos.Recordset.EOF Or Ado_datos.Recordset.BOF Then
'      BtnModificar.Enabled = False
     ' BtnEliminar.Enabled = False
      'TxtTipo.Text = Empty
      TxtCodigo = Empty
  Else
'lblActivo.Caption = Ado_datos.Recordset!sigla_emprea

    If Ado_datos.Recordset!fecha_expiracion <= Date Then
        If OptFilGral1.Value = True Then
            Label18.Visible = True
            'Label18.Caption = "Dado de Baja el " & Ado_datos.Recordset!fecha_expiracion & " Se quitara de la pantaala de Vigentes Depues de generar la siguiente planilla"
            Label18.Caption = "Dado de Baja el " & Ado_datos.Recordset!fecha_expiracion
        Else
            Label18.Visible = False
            Label18.Caption = ""
        End If
    
        If OptFilGral1.Value = True Then
            Label18.Visible = True
            Label18.Caption = "Dado de Baja el " & Ado_datos.Recordset!fecha_expiracion
        Else
            Label18.Visible = False
            Label18.Caption = ""
        End If
    Else
        Label18.Visible = False
        Label18.Caption = ""
    End If
    '      Text2.Text = Empty
    '      Text3.Text = Empty
         ' txtDenominacion = Empty
    '      Exit Sub
  End If
  
   'BtnModificar.Enabled = True
   'BtnEliminar.Enabled = True
  If Ado_datos.Recordset.RecordCount > 0 Then
    Call filtrar_asistencia(txt_mes.Text, cbo_gestion.Text)
    '   Ado_datos.Recordset.MoveFirst
    '  txtDenominacion.Caption = Ado_datos.Recordset!beneficiario_denominacion
    
    Select Case Ado_datos.Recordset.EditMode
      Case adEditInProgress
        Frame2.Enabled = False            'Verif. Nombre Proveedor JQA NOV-2009
      Case adEditNone
        Set Img_Foto = Leer_Imagen(db, "Select Foto From rv_personal_contratado Where beneficiario_codigo= '" & Ado_datos.Recordset!beneficiario_codigo & "' ", "Foto")
        Image2 = Img_Foto
        'CmdFoto.Visible = True
        'Set Picture1 = Leer_Imagen(Cn, ComandoSQL, CampoImagen)
         'Call

'         Text1.Text = IIf(IsNull(pRecordset("paterno_beneficiario")), "", pRecordset("paterno_beneficiario"))
'         Text2.Text = IIf(IsNull(pRecordset("materno_beneficiario")), "", pRecordset("materno_beneficiario"))
'         Text3.Text = IIf(IsNull(pRecordset("nombres_beneficiario")), "", pRecordset("nombres_beneficiario"))
'         TxtTipo.Text = IIf(IsNull(pRecordset("tipoben_codigo")), "-", pRecordset("tipoben_codigo"))
         'lblActivo.Caption = IIf(IsNull(pRecordset("estado_codigo")), "", pRecordset("estado_codigo"))
         'lblActivo.Caption = IIf(IsNull(Ado_datos.Recordset!estado_codigo), "", Ado_datos.Recordset!estado_codigo)
         'lblActivo2.Caption = IIf(IsNull(pRecordset("estado_codigo")), "", pRecordset("estado_codigo"))
         'lblActivo2.Caption = IIf(IsNull(Ado_datos.Recordset!estado_codigo), "", Ado_datos.Recordset!estado_codigo)
'         txtDenominacion.Text = IIf(IsNull(pRecordset("beneficiario_denominacion")), "-", pRecordset("beneficiario_denominacion"))
'         TxtDireccion.Text = IIf(IsNull(pRecordset("direccion_benef")), "-", pRecordset("direccion_benef"))
'         TxtTelefono.Text = IIf(IsNull(pRecordset("Telefono")), "-", pRecordset("Telefono"))
'         DTP_FechaNac.Value = IIf(IsNull(pRecordset("fecha_nacimiento")), "", pRecordset("fecha_nacimiento"))
'         Txtfecha.Text = IIf(IsNull(pRecordset("fecha_registro")), "", pRecordset("fecha_registro"))
'         Txthora.Text = IIf(IsNull(pRecordset("hora_registro")), "", pRecordset("hora_registro"))
'         Txtusuario.Text = IIf(IsNull(pRecordset("usr_usuario")), "", pRecordset("usr_usuario"))
'         TxtCargo
'         TxtProfesion
'         TxtZona
'         Txt_mail
        
        If Ado_datos.Recordset("Fecha_expiracion") <= Date And Ado_datos.Recordset("tipoben_codigo") = "2" Then
'            Ado_datos.Recordset("estado_codigo") = "N"
'            MsgBox "La fecha de validez del CONTRATO ya expiro, será deshabilitado el Consultor:" + Ado_datos.Recordset("beneficiario_denominacion")
        End If
        'If pRecordset("Fecha_expiracion") <= Date And pRecordset("tipoben_codigo") = "2" Then
        '    pRecordset("estado_codigo") = "N"
        '    MsgBox "La fecha de validez del RENCA ya expiro, será deshabilitado el Consultor:" + pRecordset("beneficiario_denominacion")
        'End If
'        Set rs_vacaciones_prog = New ADODB.Recordset
        If GlSW <> "ADD" Then
            Call abrirtabla
'            Set rs_vacaciones_prog = New ADODB.Recordset'<>
'            rs_vacaciones_prog.Open "select * from rc_datos_educacionales where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", DB, adOpenKeyset, adLockOptimistic
'            Set Ado_VacacionesProg.Recordset = rs_vacaciones_prog
'
'            Set rs_laborales = New ADODB.Recordset
'            rs_laborales.Open "select * from rc_experiencia_laboral where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", DB, adOpenKeyset, adLockOptimistic
'            Set Ado_Laborales.Recordset = rs_laborales
        Else
'            'Set Ado_ProyUbic.Recordset = RSNADA
'            Set DtgVacacionesProg.DataSource = RSNADA
'            Set DtgVacaciones.DataSource = RSNADA
''            rs_ProyUbic.Open "select * from mo_proy_Id_Ubicacion  ", db, adOpenKeyset, adLockOptimistic
        End If
        
        If SSTab1.Tab = 0 Then
        'SSTab1.TabEnabled(0) = True
        'SSTab1.TabEnabled(1) = False
        Else
'           FrmEditaDet.Visible = False
'           DtGLista.Visible = False
'           adoao_solicitud_lista.Visible = False
        End If
        'If pRecordset("tipoben_codigo") = "6" Then
'        If Ado_datos.Recordset("tipoben_codigo") = "6" Then
'            SSTab1.Tab = 3
'            SSTab1.TabEnabled(0) = False
''            SSTab1.TabEnabled(3) = True
'        Else
            'SSTab1.Tab = 0
            SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = True
'        End If
'        If pRecordset("estado_codigo") = "REG" Then
'            BtnAprobar.Visible = True
'            CmdDesapr.Visible = False
'        Else
'            BtnAprobar.Visible = False
'            CmdDesapr.Visible = True
'        End If
        'JQA NOV-2010
'        If Ado_datos.Recordset("tipoben_codigo") = "0" Then
'            TxtRenca.Visible = False
'            'TxtRenca.BackColor =&H8000000B&
'            DTP_FechaExpira.Visible = False
'            Label13.Visible = False
'            Label10.Visible = False
'        Else
'            TxtRenca.Visible = True
'            'TxtRenca.BackColor =&H8000000B&
'            DTP_FechaExpira.Visible = True
'            Label13.Visible = True
'            Label10.Visible = True
'        End If

'        If pRecordset("tipoben_codigo") = "6" Then
'            TxtNIT.Text = pRecordset("beneficiario_codigo")
'            txtCodigo.Text = IIf(IsNull(pRecordset("Nit")), "", pRecordset("Nit"))
'        Else
'            txtCodigo.Text = pRecordset("beneficiario_codigo")
'            TxtNIT.Text = IIf(IsNull(pRecordset("Nit")), "", pRecordset("Nit"))
'        End If
      Case adEditDelete
      Case adEditAdd
        Frame2.Enabled = True            'Verif. Nombre Proveedor JQA NOV-2009
    End Select
'    If Ado_datos.Recordset("estado_codigo") = "APR" Then
'        lblActivo.ForeColor = &H8000&
'    Else
'        lblActivo.ForeColor = &HC0&
'    End If
    If Ado_datos.Recordset("ARCHIVO_FOTO") = "Cargar_Archivo" Then
        LblInicial.ForeColor = &HC0&
    Else
        LblInicial.ForeColor = &H8000&
    End If
    If Ado_datos.Recordset("ARCHIVO_HOJAVIDA") = "Cargar_Archivo" Then
        LblCV.ForeColor = &HC0&
    Else
        LblCV.ForeColor = &H8000&
    End If
'    If Ado_datos.Recordset("ARCHIVO_RESPALDO") = "Cargar_Archivo" Then
'        LblResp.ForeColor = &HC0&
'    Else
'        LblResp.ForeColor = &H8000&
'    End If
    If swnuevo = 0 Then
    'If Not (IsNull(AdoTip_ben.Recordset("tipoben_codigo"))) Then
    '            If Not (AdoTip_ben.Recordset.BOF) Then AdoTip_ben.Recordset.MoveFirst
    '            AdoTip_ben.Recordset.Find "tipoben_codigo='" & Ado_datos.Recordset!tipoben_codigo & "'", , adSearchForward
    '            If Not AdoTip_ben.Recordset.EOF Then
    '                'TDBC_marcas.Item(1) = AdoMarca.Recordset!descripcion
    '            End If
    'End If
      Ado_datos.Caption = CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & CStr(Ado_datos.Recordset.RecordCount)
    End If
  End If
End Sub
   
Private Sub AdoPermiso_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Not AdoPermiso.Recordset.EOF Then
        Set rs_Permiso_detalle = New ADODB.Recordset
        If rs_Permiso_detalle.State = 1 Then rs_Permiso_detalle.Close
        rs_Permiso_detalle.Open "select * from ro_Permisos_detalle where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and Correl = '" & AdoPermiso.Recordset!CORREL & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    End If
End Sub

Private Sub BtnAux2_Click()
On Error GoTo EditErr
    glPersNew = "FICHA"
    gw_p_gc_beneficiario_aux.Show vbModal
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub btncumple_Click()
LISTAC = ""
   Set rs_cumple = New ADODB.Recordset
   If rs_cumple.State = 1 Then rs_cumple.Close
   rs_cumple.Open "rp_cumple_prox 1", db, adOpenKeyset, adLockOptimistic, adCmdText
   If rs_cumple.RecordCount > 0 Then
   LISTAC = "HOY CUMPLEN" & vbCrLf & vbCrLf
   rs_cumple.MoveFirst
   While Not rs_cumple.EOF
   LISTAC = LISTAC & rs_cumple!beneficiario_denominacion & " EN " & rs_cumple!depto_descripcion & vbCrLf
   rs_cumple.MoveNext
   Wend
   End If
   
  If LISTAC = "" Then
  LISTAC = "NO HAY CUMPLEAÑOS HOY" & vbCrLf & vbCrLf
  End If
   
   Set rs_cumple = New ADODB.Recordset
   If rs_cumple.State = 1 Then rs_cumple.Close
   rs_cumple.Open "rp_cumple_prox 2", db, adOpenKeyset, adLockOptimistic, adCmdText
   If rs_cumple.RecordCount > 0 Then
   LISTAC = LISTAC & vbCrLf & "MAÑANA CUMPLEN" & vbCrLf & vbCrLf
   rs_cumple.MoveFirst
   While Not rs_cumple.EOF
   LISTAC = vbCrLf & LISTAC & rs_cumple!beneficiario_denominacion & " EN " & rs_cumple!depto_descripcion & vbCrLf
   rs_cumple.MoveNext
   Wend
   End If
   
  If LISTAC = "NO HAY CUMPLEAÑOS HOY" Then
  LISTAC = LISTAC & vbCrLf & "NO HAY CUMPLEAÑOS MAÑANA"
  End If
  
  MsgBox LISTAC
End Sub

Private Sub BtnGrabar_Click()
VAR_COD2 = Ado_datos.Recordset!beneficiario_codigo
If dtc_desc2.Text <> "PERSONAL A PRUEBA" Then
    If dtc_afp_des.Text = "" Then
        MsgBox "Elija la AFP a la que corresponde la persona", vbCritical, "Error"
        Exit Sub
    End If
    If txt_afp.Text = "" Then
        MsgBox "Tiene que llenar el NUA de la persona", vbCritical, "Error"
        Exit Sub
    End If
End If

If dtc_desc2.Text = "PERSONAL A PRUEBA" Then
    If dtc_afp_des.Text <> "" Then
        MsgBox "el PERSONAL A PRUEBA no puede aportar a una AFP, la asignacion sera borrada." & vbCrLf & "Todos los demas datos seran guardados", vbCritical, "Error"
        dtc_afp.Text = ""
        dtc_afp_des.Text = ""
    End If
End If
   V_TIPO = Trim(TxtTipo.Text)
   V_TDOC = Trim(Dtc_doc_id)
   'On Error GoTo errorAceptar
   With Ado_datos
     If swnuevo = 1 Then
       CORREL = 0
'       DE.dbo_fc_correl_ben CORREL
       Set rstbeneaux = New ADODB.Recordset
       'SQL_FOR = "select * from rv_personal_contratado where beneficiario_codigo= '" & TxtCodigo.Text & "' OR beneficiario_codigo= '" & txtCodigo2.Text & "' "
       SQL_FOR = "select * from Gc_beneficiario where beneficiario_codigo= '" & TxtCodigo & "'  "
       rstbeneaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
       'If rstbeneaux.RecordCount > 0 And txtCodigo.Enabled Then
       If rstbeneaux.RecordCount > 0 Then
            SW = True
            MsgBox " CODIGO DUPLICADO"
            ' TxtCodigo.SetFocus
            Exit Sub
       End If
     End If
     If TxtTipo < "20" Then
        If Trim(TxtCodigo) = "" Then
            MsgBox "Introduzca el No. Documento de Identidad :"
'                TxtCodigo.SetFocus
            Exit Sub
        End If
        If TxtTipo.Text = "" Then
            MsgBox "Introduzca el Tipo de Persona :"
            TDBtipoben.SetFocus
            Exit Sub
        End If
        If txtDenominacion = "" Then
            MsgBox "INTRODUZCA el nombre completo de la persona:"
'                txtDenominacion
            Exit Sub
        End If
     End If
        ' CORREL = CORREL + 1
        db.BeginTrans
        SW = False
        If TxtCodigo.Enabled And swnuevo = 1 Then
            .Recordset("beneficiario_codigo") = Trim(TxtCodigo)
'            If TxtTipo2.Text = "6" Then
'                 .Recordset("NIT") = TxtCodigo.Text
'            End If
            If TxtTipo.Text < "20" Then
                 .Recordset("NIT") = TxtNIT
            End If
            accion = 0
            .Recordset("estado_codigo").Value = "REG"
            .Recordset("ARCHIVO_FOTO") = "Cargar_Archivo"
            .Recordset("ARCHIVO_HOJAVIDA") = "Cargar_Archivo"
            .Recordset("ARCHIVO_RESPALDO") = "Cargar_Archivo"
'            rs_contrato!ARCHIVO_NOMB = Trim(DtcInicial.Text) & "_Contrato_" & rs_contrato!numero_consultoria & ".pdf"
            .Recordset("archivo_foto") = Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "_Foto.JPG"
            .Recordset("ARCHIVO_HV") = Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "_HojadeVida_1.pdf"
            .Recordset("ARCHIVO_RESP") = Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "_Respaldo_1.pdf"
            
'            Dim RUTA1 As String
'            RUTA1 = "PERSONAL" + "\" + Text1 + " " + Text2 + " " + Text3
'            MsgBox RUTA1
'            MkDir RUTA1
'
''            RUTA1 = "PERSONAL" + "\" + Trim(LblInicial)
''            MsgBox RUTA1
''            MkDir RUTA1
'
'            RUTA1 = "PERSONAL" + "\" + Trim(LblInicial) + "-" + Trim(txtCodigo)
'            MsgBox RUTA1
'            MkDir RUTA1
        End If
        If TxtTipo.Text < "20" Then
            .Recordset("Reg_Profesional") = TxtRenca.Text
            .Recordset("ocup_codigo") = IIf(Dtc_Ocup.Text = "", 0, Dtc_Ocup.Text)
            .Recordset("pais_codigo") = DtcPaisCod.Text
            .Recordset("depto_codigo") = IIf(Dtc_depto_cod.Text = "", "0", Dtc_depto_cod.Text)
            .Recordset("prov_codigo") = IIf(Dtc_prov_cod.Text = "", "0", Dtc_prov_cod.Text)
            .Recordset("munic_codigo") = IIf(Dtc_munic_cod.Text = "", "0", Dtc_munic_cod.Text)
            .Recordset("estado_civil_codigo") = DtcEstCiv.Text
            .Recordset("unidad_codigo_pla") = dtc_codigo2.Text
            .Recordset("beneficiario_haber_mensual") = txt_sueldo.Text
            .Recordset("beneficiario_otro_mensual") = txt_otro.Text
            .Recordset("fecha_ingreso") = DTPicker2.Value
            .Recordset("unidad_codigo") = dtc_codigo1.Text
'            Dim a As String, b As String, C As String
'            a = Left(Text1.Text, 1)
'            b = Left(Text2.Text, 1)
'            C = Left(Text3.Text, 1)
'            .Recordset("beneficiario_beneficiario_iniciales") = Trim(LblInicial.Caption)
'            RUTA1 = "PERSONAL" + "\" + Trim(LblInicial) + "-" + Trim(txtCodigo)
'            MsgBox RUTA1
'            MkDir RUTA1
        End If
            .Recordset("observaciones") = "* " + txt_obs.Text
            .Recordset("usr_codigo").Value = glusuario 'frmLogin.txtUserName.Text
            .Recordset("fecha_registro").Value = Date
            .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
            .Recordset("puesto_codigo").Value = dtc_codigo3.Text
            .Recordset("cargo_codigo").Value = dtc_cargo.Text
            .Recordset.Update
            '.Recordset.Requery
'            Dim ARCH_FOTO As String
'            ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset("inicial") + "\" + Ado_datos.Recordset("inicial") + "-FOTO.JPG"
'            If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where beneficiario_codigo= '" & Ado_datos.Recordset("beneficiario_codigo") & "' ", "Foto", ARCH_FOTO) Then
'                MsgBox "Se cargo la Imagen Correctamente !!"
'            Else
'                MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'            End If
            db.CommitTrans
      End With
   '
   swnuevo = 0
   fraOpciones.Visible = True
       fra_cabecera.Enabled = True
       
            SSTab1.TabEnabled(5) = True
       SSTab1.TabEnabled(4) = True
            SSTab1.TabEnabled(3) = True
         SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = True
       
       
       
   FraGrabarCancelar.Visible = False
   FraNavega.Enabled = True
   fraDatos.Enabled = False
   TxtCodigo.Enabled = True
'   FraSS_SS.Enabled = False

   Call Carga_Recor
     
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rstbeneficiario.Find "beneficiario_codigo = '" & VAR_COD2 & "'   ", , , 1
        dg_datos.SelBookmarks.Add (rstbeneficiario.Bookmark)
     Else
        rstbeneficiario.MoveLast
     End If
   'Call Carga_Beneficiario
'De.dbo_alGraba_rc_personal Accion, CORREL, txtCodigo.Text, Text1.Text, Text2.Text, Text3.Text, "2002"
 Exit Sub
errorAceptar:
   Call pErrorRst(db.Errors)
   Ado_datos.Recordset.CancelUpdate
   'db.RollbackTrans
End Sub

Private Sub BtnImprimir_Click()
    Dim iResult As Integer
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\Presupuesto\beneficiarios\crybeneficiario.rpt"
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\Generales\crybeneficiario.rpt"
    CrystalReport1.ReportFileName = App.Path & "\REPORTES\RRHH\rr_personal_cotratado.rpt"
    CrystalReport1.StoredProcParam(0) = "%"
    CrystalReport1.StoredProcParam(1) = "%"
    CrystalReport1.StoredProcParam(2) = "1 AND (ro_personal_contratado.codigo_empresa >= 1"
    iResult = CrystalReport1.PrintReport
    If iResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
    End If

    CrystalReport1.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir1_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR02.ReportFileName = App.Path & "\Reportes\RRHH\rr_vacaciones_programadas.rpt"
        CR02.WindowShowPrintSetupBtn = True
        CR02.WindowShowRefreshBtn = True
    
        CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!depto_codigo
        CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!beneficiario_codigo
        
        iResult = CR02.PrintReport
        If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
    CR02.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir2_Click()
On Error GoTo EditErr
 If Ado_datos.Recordset.RecordCount > 0 Then
    If AdoMovilidad.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\RRHH\rr_movilidad_personal.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
    
        CR03.StoredProcParam(0) = Me.AdoMovilidad.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.AdoMovilidad.Recordset!numero_cambio
        CR03.StoredProcParam(2) = Me.AdoMovilidad.Recordset!beneficiario_codigo
    
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
    CR03.WindowState = crptMaximized
 End If
 
 Exit Sub

EditErr:
  MsgBox Err.Description
 
End Sub

Private Sub BtnImprimir3_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_Memo.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR04.ReportFileName = App.Path & "\Reportes\RRHH\rr_memoranda.rpt"
        CR04.WindowShowPrintSetupBtn = True
        CR04.WindowShowRefreshBtn = True
    
        CR04.StoredProcParam(0) = Me.Ado_Memo.Recordset!ges_gestion
        CR04.StoredProcParam(1) = Me.Ado_Memo.Recordset!beneficiario_codigo
        CR04.StoredProcParam(2) = Me.Ado_Memo.Recordset!mes_descuento
        CR04.StoredProcParam(3) = UCase(MonthName(Month(Me.Ado_Memo.Recordset!fecha_memo)))
        CR04.StoredProcParam(4) = Me.Ado_Memo.Recordset!Numero
        
        iResult = CR04.PrintReport
        If iResult <> 0 Then MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical, "Error de impresión"
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
    CR04.WindowState = crptMaximized
 End If
End Sub

Private Sub BtnImprimir4_Click()
'    If Ado_datos.Recordset.RecordCount > 0 Then
'        Dim iResult As Integer
'        'Dim co As New ADODB.Command
'        CR05.ReportFileName = App.Path & "\Reportes\RRHH\rr_vacaciones_programadas.rpt"
'        CR05.WindowShowPrintSetupBtn = True
'        CR05.WindowShowRefreshBtn = True
'
'        CR05.StoredProcParam(0) = Me.Ado_datos.Recordset!depto_codigo
'        CR05.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
'        CR05.StoredProcParam(2) = Me.Ado_datos.Recordset!beneficiario_codigo
'
'        iResult = CR05.PrintReport
'        If iResult <> 0 Then MsgBox CR05.LastErrorNumber & " : " & CR05.LastErrorString, vbCritical, "Error de impresión"
'    Else
'        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
'    End If
'        v.WindowState = crptMaximized
End Sub

Private Sub cbo_gestion_Click()
    Call filtrar_asistencia(txt_mes.Text, cbo_gestion.Text)
End Sub

Private Sub cbo_mes_Click()
txt_mes.Text = cbo_mes.ListIndex
txt_mes.Text = Val(txt_mes.Text) + 1
Call filtrar_asistencia(txt_mes.Text, cbo_gestion.Text)
End Sub

Private Sub cmd_mas_Click()
    rw_datos_extra.txt_codigo = Ado_datos.Recordset!beneficiario_codigo
    rw_datos_extra.Txt_descripcion = Ado_datos.Recordset!beneficiario_denominacion
    rw_datos_extra.txt_ext = Ado_datos.Recordset!depto_sigla
    rw_datos_extra.txt_tipo_doc = Ado_datos.Recordset!tipodoc_codigo
    rw_datos_extra.dtc_desc.BoundText = Ado_datos.Recordset!codigo_empresa
    If Ado_datos.Recordset!discapacidad = 1 Then
        rw_datos_extra.cmb_discapacidad.Text = "SI"
    Else
        rw_datos_extra.cmb_discapacidad.Text = "NO"
    End If
    
    If Ado_datos.Recordset!tutor = 1 Then
        rw_datos_extra.cmb_tutor.Text = "SI"
    Else
        rw_datos_extra.cmb_tutor.Text = "NO"
    End If
    
    rw_datos_extra.Show vbModal
End Sub

Private Sub cmdAdd2_Click()
On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
       marca1 = Ado_datos.Recordset.Bookmark
       frm_ao_Vacacion_Prog.txtSW = "ADD"
       frm_ao_Vacacion_Prog.txtBenef = Ado_datos.Recordset!beneficiario_codigo
       frm_ao_Vacacion_Prog.txtEstado = "REG"
       frm_ao_Vacacion_Prog.TxtGestion.Text = Year(Date)
       frm_ao_Vacacion_Prog.txt_empresa.Text = Ado_datos.Recordset!codigo_empresa
'      Ado_VacacionesProg.Recordset.AddNew
'       frm_ao_Vacacion_Prog.lblbien(1).Visible = True
'       frm_ao_Vacacion_Prog.Txt02.Visible = True
         frm_ao_Vacacion_Prog.sel = 1
       frm_ao_Vacacion_Prog.Show vbModal
     
       Call abrirtabla
       'Ado_VacacionesProg.Refresh
   Else
       MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdAdd3_Click()
On Error GoTo EditErr

   If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ao_Permisos_js.txtSW = "ADD"
        frm_ao_Permisos_js.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ao_Permisos_js.txtEstado = "REG"
        'AdoPermiso.Recordset.AddNew
        frm_ao_Permisos_js.TxtGestion = Year(Date)
        frm_ao_Permisos_js.Show vbModal
        
        Call abrirtabla
        'AdoPermiso.Recordset.MoveLast
   Else
        MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdAdd4_Click()
On Error GoTo EditErr

   If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ao_memoranda.txtSW = "ADD"
        frm_ao_memoranda.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ao_memoranda.TxtInicial = IIf(IsNull(Ado_datos.Recordset!beneficiario_iniciales), "", Ado_datos.Recordset!beneficiario_iniciales)
        frm_ao_memoranda.txtEstado = "REG"
'        Ado_Memo.Recordset.AddNew
        frm_ao_memoranda.Show vbModal
        Call abrirtabla
        'Ado_Memo.Refresh
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description

End Sub

Private Sub CmdAdd1_Click()
On Error GoTo EditErr
    If Ado_datos.Recordset.RecordCount > 0 Then
      marca1 = Ado_datos.Recordset.Bookmark
      frm_ao_Asistencia.txtSW = "ADD"
      AdoAsistencia.Recordset.AddNew
      frm_ao_Asistencia.txtSW = "ADD"
      frm_ao_Asistencia.txtBenef = Ado_datos.Recordset!beneficiario_codigo
      frm_ao_Asistencia.DTPFec_Inicio = Date    'Ado_datos.Recordset!Fecha_control
      'frm_ao_Asistencia.lblcodigo_solicitud = Ado_datos.Recordset!codigo_solicitud
      frm_ao_Asistencia.txtEstado = "REG"
      frm_ao_Asistencia.Show vbModal
      Call abrirtabla
   Else
        MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdAdd5_Click()
On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ro_movilidad_personal.txtSW.Text = "ADD"
        frm_ro_movilidad_personal.txtBenef.Text = Ado_datos.Recordset!beneficiario_codigo
        'frm_ro_movilidad_personal.TxtInicial = Ado_datos.Recordset!beneficiario_iniciales
        frm_ro_movilidad_personal.TxtAprob = "REG"
        frm_ro_movilidad_personal.DtcPryDes = IIf(IsNull(Ado_datos.Recordset!puesto_descripcion), "", Ado_datos.Recordset!puesto_descripcion)
        frm_ro_movilidad_personal.DTPFelaboracion.Value = Date
        frm_ro_movilidad_personal.DtcPry.BoundText = frm_ro_movilidad_personal.DtcPryDes.BoundText
        frm_ro_movilidad_personal.DtcPryCargo.BoundText = frm_ro_movilidad_personal.DtcPryDes.BoundText
        frm_ro_movilidad_personal.DtcPryUni.BoundText = frm_ro_movilidad_personal.DtcPryDes.BoundText
        frm_ro_movilidad_personal.Dtc_codigo_ant.BoundText = frm_ro_movilidad_personal.DtcPryUni.Text
        frm_ro_movilidad_personal.Dtc_descrip_ant.BoundText = frm_ro_movilidad_personal.DtcPryUni.Text
        frm_ro_movilidad_personal.DtcOrgDes.BoundText = frm_ro_movilidad_personal.DtcPryCargo.Text
        frm_ro_movilidad_personal.DtcOrg.BoundText = frm_ro_movilidad_personal.DtcPryCargo.Text
        'AdoMovilidad.Recordset.AddNew
        frm_ro_movilidad_personal.DTPFcontrato.Value = Date
        frm_ro_movilidad_personal.Show vbModal
        Call abrirtabla
        'AdoMovilidad.Refresh
        
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
   
Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdApr3_Click()
 On Error GoTo EditErr
 If AdoPermiso.Recordset("estado_codigo") = "REG" Then '
 If AdoPermiso.Recordset("TipoPermiso") = "VC" Or AdoPermiso.Recordset("TipoPermiso") = "VP" Then
 Dim DIAS, DIASP, TOTALD As Double
 sino = MsgBox("Gestion: " & Ado_VacacionesProg.Recordset("ges_Gestion") & vbCrLf & "Mes: " & Ado_VacacionesProg.Recordset("Mes_control") & vbCrLf & "Días Programados: " & Ado_VacacionesProg.Recordset("dias_Programados") & vbCrLf & "¿Está seguro de descontar los dias de esta vacacion?", vbYesNo + vbQuestion, "Atención")
 If sino = vbYes Then
 DIASP = AdoPermiso.Recordset("dias_permiso") + Ado_VacacionesProg.Recordset("dias_utilizados")
 DIAS = Ado_VacacionesProg.Recordset("dias_Programados")
 
 If DIASP <= DIAS Then
 If AdoPermiso.Recordset("TipoPermiso") = "VP" Then 'pago por vp
 Dim PAGO As Double
 PAGO = (Ado_datos.Recordset!beneficiario_haber_mensual / 30) * Ado_VacacionesProg.Recordset!Dias_Pendientes
 'Call Pagos("DRRHH", "formulario", "org_codigo", "solicitud_codigo", "justificacion", "observaciones", Ado_datos.Recordset!beneficiario_codigo, "concepto_pago", "obs_fo_gastos_detalle", PAGO)
 End If
 Ado_VacacionesProg.Recordset("dias_utilizados") = DIASP
 Ado_VacacionesProg.Recordset("Dias_Pendientes") = DIAS - DIASP
 Ado_VacacionesProg.Recordset.Update
 Else
 sino = MsgBox("Los dias disponibles de esta vacacion son insuficientes", vbCritical, "Atención")
 Exit Sub
 End If
 Else
 Exit Sub
 End If
 Else
    sino = MsgBox("Está Seguro de APROBAR el Registro activo ? ", vbYesNo + vbQuestion, "Atención")
 End If
 
 
      If sino = vbYes Then
        'Call detalle_permiso
        AdoPermiso.Recordset("estado_codigo") = "APR"
        AdoPermiso.Recordset("fecha_registro") = Date
        AdoPermiso.Recordset("usr_usuario") = glusuario
        AdoPermiso.Recordset.Update
        '''''''Call opciones
         
         
          Dim rs_datos4 As New ADODB.Recordset
        If rs_datos4.State = 1 Then rs_datos4.Close
         rs_datos4.Open "select * from ro_pagos_cronograma_Detalle where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "'  and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic
         If rs_datos4.RecordCount <> 0 Then
        
         If rs_aux9.State = 1 Then rs_aux9.Close
            rs_aux9.Open "select sum(AtrasoMin1) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '" & RTrim(LTrim(AdoPermiso.Recordset!beneficiario_codigo)) & "' AND ges_gestion = '" & RTrim(LTrim(AdoPermiso.Recordset!ges_gestion)) & "' and Mes_control = '" & RTrim(LTrim(Month(AdoPermiso.Recordset!Fecha_control))) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        
        If rs_aux14.State = 1 Then rs_aux14.Close
          rs_aux14.Open "select sum(total_minuto) as PermisoMes from ro_permisos where beneficiario_codigo = '" & RTrim(LTrim(AdoPermiso.Recordset!beneficiario_codigo)) & "' AND ges_gestion = '" & RTrim(LTrim(AdoPermiso.Recordset!ges_gestion)) & "' AND Mes_control = '" & AdoPermiso.Recordset!mes_control & "'  AND estado_codigo = 'APR' and dias_permiso = 0 and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            
            If rs_aux14!PermisoMes <> "NULL" Then
                permisos = rs_aux14!PermisoMes
            Else
                permisos = "0"
            End If
       
        If rs_aux9!TardanzaMes <> "NULL" Then
             totalminutos = rs_aux9!TardanzaMes - permisos
                If totalminutos >= 45 And totalminutos <= 60 Then
                   calretrasos = ((rs_datos4!sueldo_basico / 30) / 2)
                Else
                    If totalminutos >= 61 And totalminutos <= 75 Then
                        calretrasos = (rs_datos4!sueldo_basico / 30)
                    Else
                        If totalminutos >= 76 And totalminutos <= 105 Then
                            calretrasos = ((rs_datos4!sueldo_basico / 30) * 2)
                        Else
                            If totalminutos >= 106 Then
                               calretrasos = ((rs_datos4!sueldo_basico / 30) * 3)
                            Else
                                calretrasos = 0
                            End If
                        End If
                    End If
                End If
           End If
     
     If rs_aux10.State = 1 Then rs_aux10.Close
          rs_aux10.Open "select sum(monto) as montomes from ro_memorandas where beneficiario_codigo = '" & RTrim(LTrim(AdoPermiso.Recordset!beneficiario_codigo)) & "' AND ges_gestion = '" & RTrim(LTrim(AdoPermiso.Recordset!ges_gestion)) & "' AND mes_descuento = '" & AdoPermiso.Recordset!mes_control & "'  AND estado_codigo = 'APR'  and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & ", db, adOpenKeyset, adLockOptimistic, adCmdText"
            If rs_aux10!montomes <> "NULL" Then
            Dim memo As Double
            memo = 0
            memo = rs_aux10!PermisoMes
            'db.Execute "update ro_pagos_cronograma_Detalle set otros_dsctos = " & memo & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "'"
            Else
            memo = 0
            'db.Execute "update ro_pagos_cronograma_Detalle set otros_dsctos = " & memo & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "'"
            End If
     
     
       total = 0
       total = memo + calretrasos
       total2 = rs_datos4!anticipo_sueldo + rs_datos4!anticipo_refrigerio + rs_datos4!prestamo + rs_datos4!afp1 + rs_datos4!afp2 + rs_datos4!rciva + total
           
     
      db.Execute "update ro_pagos_cronograma_Detalle set otros_dsctos = " & total & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & ""
      db.Execute "update ro_pagos_cronograma_Detalle set total_dsctos = " & total2 & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & ""
      total = 0
      total = rs_datos4!total_ganado - total2
      db.Execute "update ro_pagos_cronograma_Detalle set liquido_pagable_bs = " & total & ", liquido_pagable_us = " & (total2 / GlTipoCambioOficial) & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & ""
      
      
      
         
      End If
      End If
Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
End If


Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub detalle_permiso()
'    Dim fecha2 As Date
'    Dim horaIng, horaSal As Date
'    Dim dia2 As String
'    Dim NoHoras, NoMin As Integer
'    Dim DifHr1, DifHr2 As Integer
'    Dim rs_premisoCtrl As New ADODB.Recordset
'    fecha2 = AdoPermiso.Recordset("FechaDesde")
'    horaIng = AdoPermiso.Recordset("horadesde")
'    horaSal = AdoPermiso.Recordset("horahasta")
'    DifHr1 = DateDiff("h", CDate(GlHora1), frmBeneficiario_Control.AdoPermiso.Recordset("horadesde").Value)
'    DifHr2 = 4 - DateDiff("h", CDate(GlHora2), frmBeneficiario_Control.AdoPermiso.Recordset("horahasta").Value)
'    If DifHr1 > 0 Then
'      If DifHr1 > 4 Then
'          DifHr1 = 4
'      Else
'          DifHr1 = DifHr1
'      End If
'    Else
'       DifHr1 = 0
'    End If
'    If DifHr2 > 0 Then
'       DifHr2 = DifHr2
'    Else
'       DifHr2 = 0
'    End If
'    While fecha2 <= AdoPermiso.Recordset("FechaHasta")
'      Set rs_calendario2 = New ADODB.Recordset
'      rs_calendario2.Open "select * from gc_calendario where fecha = '" & fecha2 & "' and tipo = 'L' ", DB, adOpenKeyset, adLockOptimistic, adCmdText
'      If rs_calendario2.RecordCount > 0 Then
'        Set rs_premisoCtrl = New ADODB.Recordset
'        rs_premisoCtrl.Open "select * from ro_Permisos_detalle where beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "' and Fecha_control = '" & fecha2 & "' and Correl = '" & AdoPermiso.Recordset!CORREL & "' ", DB, adOpenKeyset, adLockOptimistic, adCmdText
'        If rs_premisoCtrl.RecordCount > 0 Then
'            rs_Permiso_detalle!Fecha_control = AdoPermiso.Recordset("fecha2")
'        Else
'            rs_Permiso_detalle.AddNew
'            rs_Permiso_detalle!Fecha_control = fecha2
'            rs_Permiso_detalle!beneficiario_codigo = AdoPermiso.Recordset("beneficiario_codigo")
'            rs_Permiso_detalle!CORREL = AdoPermiso.Recordset("Correl")
'            rs_Permiso_detalle!ges_gestion = AdoPermiso.Recordset("ges_gestion")
'        End If
'        dia2 = WeekdayName(Weekday(fecha2))
'        rs_Permiso_detalle!dia_control = dia2
'        If horaIng > GlHora1 Then
'            rs_Permiso_detalle!horadesde = horaIng
'            horaIng = GlHora1
'            NoHoras = 8 - (DifHr1 + DifHr2)
'        Else
'            rs_Permiso_detalle!horadesde = GlHora1
'            NoHoras = 8
'        End If
''        If horaSal > CDate("16:30:00") and horaSal < CDate("16:30:00") Then
''            rs_Permiso_detalle!horahasta = horaSal
''            horaSal = CDate("18:30:00")
''        Else
''            rs_Permiso_detalle!horahasta = CDate("18:30:00")
''        End If
'        NoMin = NoHoras * 60
'        If AdoPermiso.Recordset("TipoPermiso") = "VC" Then
'            rs_Permiso_detalle!Vacacion = NoMin
'        Else
'            rs_Permiso_detalle!Vacacion = 0
'        End If
'        rs_Permiso_detalle!horas_permiso = NoHoras
'        rs_Permiso_detalle!minutos_permiso = NoMin
'        rs_Permiso_detalle!usr_usuario = GlUsuario
'        rs_Permiso_detalle!fecha_registro = Date
'        rs_Permiso_detalle!hora_registro = "08:00"
'        rs_Permiso_detalle.Update
'      End If
'      fecha2 = fecha2 + 1
'    Wend
End Sub

Private Sub CmdApr4_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_Memo.Recordset("estado_codigo") = "REG" Then
      If Ado_Memo.Recordset!ARCHIVO = "Cargar_Archivo" Then   '---------------------------> No dejaba aprobar <>
        If sino = vbYes Then
            Ado_Memo.Recordset!estado_codigo = "APR"
            Ado_Memo.Recordset!fecha_registro = Date
            Ado_Memo.Recordset!usr_codigo = glusuario
            Ado_Memo.Recordset.Update
            
   If rs_datos3.State = 1 Then rs_datos3.Close
   rs_datos3.Open "select * from rc_tipo_memoranda where tipo_memo = '" & Ado_Memo.Recordset("tipo_memo") & "' ", db, adOpenKeyset, adLockOptimistic
   
   If rs_datos3!estado_baja = "S" Then
   db.Execute "update ro_personal_contratado set fecha_expiracion = '" & Ado_Memo.Recordset!fecha_aprobacion & "' WHERE beneficiario_codigo = '" & Ado_Memo.Recordset("beneficiario_codigo") & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & ""
   'db.Execute "update ro_personal_contratado set estado_codigo = 'ANL' WHERE beneficiario_codigo = '" & Ado_Memo.Recordset("beneficiario_codigo") & "'"
   db.Execute "update rc_puestos set puesto_vacante = 'SI' ,  beneficiario_codigo = '' WHERE puesto_codigo = '" & Ado_datos.Recordset("puesto_codigo") & "' AND cargo_codigo = " & Ado_datos.Recordset("cargo_codigo") & " AND unidad_codigo ='" & dtc_codigo1.Text & "'"
   Call Carga_Beneficiario(1)
   End If
    If Ado_Memo.Recordset("tipo_memo") = "JUB" Then
    db.Execute "UPDATE ro_personal_contratado set estado_jubilado = 'APR' WHERE beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & """"
    End If
    
    If Ado_Memo.Recordset("descuento_pla") = "SI" Then
    'REVISA SI ESTA APROBADA LA PLANILLA
    If rs_aux18.State = 1 Then rs_aux18.Close
    rs_aux18.Open "select * from ro_pagos_grupos where ges_gestion = '" & Year(Ado_Memo.Recordset("fecha_aprobacion")) & "' AND mes_grupo = " & Ado_Memo.Recordset("mes_descuento") & " AND planilla_codigo = '" & Left(Ado_datos.Recordset("unidad_codigo_pla"), 3) & "'", db, adOpenKeyset, adLockOptimistic
    If rs_aux18.RecordCount > 0 Then
    If rs_aux18!estado_codigo = "APR" Then
    sino = MsgBox("La planilla de " & UCase(MonthName(Ado_Memo.Recordset("mes_descuento"))) & " " & Year(Ado_Memo.Recordset("fecha_aprobacion")) & " ya fue aprobada y no se puede realizar ningun cambio en los descuentos", vbCritical, "Error")
    Ado_Memo.Recordset!estado_codigo = "REG"
    Ado_Memo.Recordset!fecha_registro = Date
    Ado_Memo.Recordset!usr_codigo = glusuario
    Ado_Memo.Recordset.Update
    Exit Sub
    End If
    End If
    total = 0
    total2 = 0
    VAR_MES = Month(Ado_Memo.Recordset!fecha_aprobacion)
    Dim rs_datos4 As New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
     rs_datos4.Open "select * from ro_pagos_cronograma_Detalle where ges_gestion = '" & Ado_Memo.Recordset!ges_gestion & "' AND mes_grupo = " & VAR_MES & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "' AND estado_codigo = 'REG'", db, adOpenKeyset, adLockOptimistic
     If rs_datos4.RecordCount <> 0 Then
     
     If Ado_Memo.Recordset("monto") > 0 Then
     total = Ado_Memo.Recordset("monto")
'     rs_datos!otros_dsctos = total
'     rs_datos!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!rciva + rs_datos2!otros_dsctos
     End If
         
    Dim rs_datos1 As New ADODB.Recordset
     If Ado_Memo.Recordset("dias") > 0 Then
     If rs_datos1.State = 1 Then rs_datos1.Close
     rs_datos1.Open "select * from ro_personal_contratado where beneficiario_codigo = '" & Ado_Memo.Recordset("beneficiario_codigo") & "'", db, adOpenKeyset, adLockOptimistic
     total = total + ((rs_datos1!beneficiario_haber_mensual / 30) * Ado_Memo.Recordset("dias"))
     'total = total + rs_datos4!otros_dsctos
'     rs_datos!otros_dsctos = total
'     rs_datos!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!rciva + rs_datos2!otros_dsctos
     End If

     If total > 0 Then
     total = total + rs_datos4!otros_dsctos
     
     db.Execute "update ro_pagos_cronograma_Detalle set otros_dsctos = " & total & "where ges_gestion = '" & Ado_Memo.Recordset!ges_gestion & "' AND mes_grupo = " & VAR_MES & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "'"
     'rs_datos4!otros_dsctos = total
'     total = 0
     total2 = rs_datos4!anticipo_sueldo + rs_datos4!anticipo_refrigerio + rs_datos4!prestamo + rs_datos4!afp1 + rs_datos4!afp2 + rs_datos4!rciva + total
     'rs_datos4!total_dsctos = rs_datos4!anticipo_sueldo + rs_datos4!anticipo_refrigerio + rs_datos4!prestamo + rs_datos4!afp1 + rs_datos4!afp2 + rs_datos4!rciva + rs_datos4!otros_dsctos
     db.Execute "update ro_pagos_cronograma_Detalle set total_dsctos = " & total2 & "where ges_gestion = '" & Ado_Memo.Recordset!ges_gestion & "' AND mes_grupo = " & VAR_MES & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "'"
     total = 0
     total = rs_datos4!total_ganado - total2
     db.Execute "update ro_pagos_cronograma_Detalle set liquido_pagable_bs = " & total & ", liquido_pagable_us = " & (total2 / GlTipoCambioOficial) & "where ges_gestion = '" & Ado_Memo.Recordset!ges_gestion & "' AND mes_grupo = " & VAR_MES & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "'"
     
     
     
     
     
     End If

     'db.Execute "update ro_pagos_cronograma_Detalle set otros_dsctos = " & total & "where beneficiario_codigo = '" & Ado_Memo.Recordset("beneficiario_codigo")
  

            End If
            End If
        '''''''Call opciones
        End If
      Else
            MsgBox "No se puede APROBAR. Previamente Debe cargar el archivo .PDF asociado al registro ... ", vbExclamation, "Validación de Registro"
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  
  MsgBox Err.Description

End Sub
Private Sub CmdApr1_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de APROBAR el Registro activo ? ", vbYesNo + vbQuestion, "Atención")
   If AdoAsistencia.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoAsistencia.Recordset("estado_codigo") = "APR"
        AdoAsistencia.Recordset("fecha_REGISTRO") = Date
        AdoAsistencia.Recordset("usr_usuario") = glusuario
        AdoAsistencia.Recordset.Update
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdApr2_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_VacacionesProg.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_VacacionesProg.Recordset("estado_codigo") = "APR"
        Ado_VacacionesProg.Recordset("fecha_registro") = Date
        Ado_VacacionesProg.Recordset("usr_usuario") = glusuario
        Ado_VacacionesProg.Recordset.Update
        '''''''''''Call opciones
        

      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If


Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdApr5_Click()
 On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If AdoMovilidad.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        VAR_COD2 = Ado_datos.Recordset!beneficiario_codigo
        AdoMovilidad.Recordset("estado_codigo") = "APR"
        AdoMovilidad.Recordset("fecha_registro") = Date
        AdoMovilidad.Recordset("usr_codigo") = glusuario
        AdoMovilidad.Recordset.Update
        
        If AdoMovilidad.Recordset!tipo_mov = "INTERCAMBIO" Then
            db.Execute "UPDATE ro_movilidad_personal set estado_codigo = 'APR' WHERE numero_intercambio = " & rs_movilidad!numero_intercambio & ""
            db.Execute "UPDATE ro_personal_contratado SET cargo_codigo = " & rs_movilidad!cargo_codigo & ", puesto_codigo = '" & rs_movilidad!puesto_codigo & "', unidad_codigo = '" & rs_movilidad!unidad_codigo & "' where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & " "
            db.Execute "UPDATE ro_personal_contratado SET cargo_codigo = " & rs_movilidad!cargo_anterior & ", puesto_codigo = '" & rs_movilidad!puesto_anterior & "', unidad_codigo = '" & rs_movilidad!unidad_anterior & "' where beneficiario_codigo = '" & rs_movilidad!beneficiario_codigo_int & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & ""
        Else
            db.Execute "update rc_puestos set puesto_vacante = 'SI', beneficiario_codigo = '' WHERE puesto_codigo = '" & Ado_datos.Recordset("puesto_codigo") & "' AND cargo_codigo = " & Ado_datos.Recordset("cargo_codigo") & " AND unidad_codigo ='" & dtc_codigo1.Text & "'"
            db.Execute "UPDATE ro_personal_contratado SET cargo_codigo = " & rs_movilidad!cargo_codigo & ", puesto_codigo = '" & rs_movilidad!puesto_codigo & "', unidad_codigo = '" & rs_movilidad!unidad_codigo & "' where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & " "
            db.Execute "update rc_puestos set puesto_vacante = 'NO', beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' WHERE puesto_codigo = '" & rs_movilidad!puesto_codigo & "' AND cargo_codigo = " & rs_movilidad!cargo_codigo & " AND unidad_codigo ='" & rs_movilidad!unidad_codigo & "'"
        End If
        If OptFilGral1.Value = True Then
           Call OptFilGral1_Click        'Pendientes
        Else
           Call OptFilGral2_Click        'TODOS
        End If
     
        If (dg_datos.SelBookmarks.Count <> 0) Then
           dg_datos.SelBookmarks.Remove 0
        End If
        If Ado_datos.Recordset.RecordCount > 0 Then
           rstbeneficiario.Find "beneficiario_codigo = " & VAR_COD2 & "   ", , , 1
           dg_datos.SelBookmarks.Add (rstbeneficiario.Bookmark)
        Else
           rstbeneficiario.MoveLast
        End If
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdDesapr_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_datos.Recordset("estado_codigo") = "APR" Then
      If sino = vbYes Then
'        Dim RUTA1, RUTA2 As String
'        RUTA1 = "PERSONAL" + "\" + Trim(Ado_datos.Recordset("beneficiario_beneficiario_iniciales")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
'        MsgBox RUTA1
'        MkDir RUTA1
'        MkDir RUTA1 + "\CONTRATOS"
'        MkDir RUTA1 + "\FINIQUITO"
'        MkDir RUTA1 + "\MEMORANDUMS"
'        MkDir RUTA1 + "\DOCUMENTOS_RESPALDO"
'        MkDir RUTA1 + "\HOJA_VIDA"
'        MkDir RUTA1 + "\OTROS"
'        MkDir RUTA1 + "\EVALUACIONES"
'        MkDir RUTA1 + "\LICENCIAS"
'        MkDir RUTA1 + "\VACACIONES"
''
''            RUTA1 = "PERSONAL" + "\" + Text1 + " " + Text2 + " " + Text3
''            MsgBox RUTA1
''            MkDir RUTA1
'
''            RUTA1 = "PERSONAL" + "\" + Trim(LblInicial)
''            MsgBox RUTA1
''            MkDir RUTA1
        Ado_datos.Recordset("estado_codigo") = "REG"
        Ado_datos.Recordset("fecha_aprueba") = Date
        Ado_datos.Recordset("usr_aprueba") = glusuario
        Ado_datos.Recordset.Update
        
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Anulado o sin Aprobar ...", vbExclamation, "Validación de Registro"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdElim1_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de ANULAR el Registro del Dependiente ? ", vbYesNo + vbQuestion, "Atención")
   If AdoAsistencia.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoAsistencia.Recordset("estado_codigo") = "ANL"
        AdoAsistencia.Recordset("fecha_registro") = Date
        AdoAsistencia.Recordset("usr_usuario") = glusuario
        AdoAsistencia.Recordset("Archivo") = "REG. ANULADO"
        AdoAsistencia.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If

Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdElim2_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de Eliminar el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_VacacionesProg.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
'        Ado_VacacionesProg.Recordset("estado_codigo") = "ANL"
'        Ado_VacacionesProg.Recordset("fecha_registro") = Date
'        Ado_VacacionesProg.Recordset("usr_usuario") = glusuario
'        'Ado_VacacionesProg.Recordset("centro_educativo") = "REG. ANULADO"
'        Ado_VacacionesProg.Recordset.Update  'Batch adAffectAll
        
        
        db.Execute "delete ro_vacaciones_programadas where ges_Gestion = '" & Ado_VacacionesProg.Recordset!ges_gestion & "' AND beneficiario_codigo = '" & Ado_VacacionesProg.Recordset!beneficiario_codigo & "' and Correl = " & Ado_VacacionesProg.Recordset!CORREL & " "
        Ado_VacacionesProg.Recordset.Update
        '''''Call opciones
        Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdElim3_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de BORRAR físicamente el Registro elegido ? ", vbYesNo + vbQuestion, "Atención")
   'If AdoPermiso.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        If AdoPermiso.Recordset!TipoPermiso = "VC" And AdoPermiso.Recordset!estado_codigo = "APR" Then
        sino = MsgBox("Se Restauraran los Días a la Vacacion Seleccionada ARRIBA, ¿Desea continuar?", vbYesNo + vbQuestion, "Atención")
         If sino = vbYes Then
        Ado_VacacionesProg.Recordset!dias_utilizados = Ado_VacacionesProg.Recordset!dias_utilizados - AdoPermiso.Recordset!dias_permiso
        Ado_VacacionesProg.Recordset!Dias_Pendientes = Ado_VacacionesProg.Recordset!dias_Programados - Ado_VacacionesProg.Recordset!dias_utilizados
        Ado_VacacionesProg.Recordset.Update
         Else
         Exit Sub
         End If
        End If
        AdoPermiso.Recordset.Delete adAffectCurrent
        
'        db.Execute "delete ro_permisos where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and Correl = " & AdoPermiso.Recordset!CORREL & " "
'        AdoPermiso.Recordset("estado_codigo") = "ANL"
'        AdoPermiso.Recordset("fecha_registro") = Date
'        AdoPermiso.Recordset("usr_usuario") = glusuario
'        'AdoPermiso.Recordset("Archivo") = "REG. ANULADO"
'        AdoPermiso.Recordset.Update  'Batch adAffectAll
        '''''''''''''''Call opciones
        Call abrirtabla
      End If
  ' Else
      '  MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
  'End If

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdElim4_Click()
'   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
'   If Ado_Memo.Recordset("estAdo_Memo") = "REG" Then
'      If sino = vbYes Then
'        Ado_Memo.Recordset("estAdo_Memo") = "ANL"
'        Ado_Memo.Recordset("fecha_registro") = Date
'        Ado_Memo.Recordset("usr_usuario") = glusuario
'        Ado_Memo.Recordset("codigo_unidad") = "REG. ANULADO"
'        Ado_Memo.Recordset.Update  'Batch adAffectAll
'      End If
'   Else
'        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
'   End If
   
If Ado_Memo.Recordset.RecordCount > 0 Then
On Error GoTo UpdateErr
If Ado_Memo.Recordset!tipo_memo = "SAD" Then
   If ExisteReg(" ges_gestion = '" & Ado_Memo.Recordset!ges_gestion & "' AND mes_grupo = " & Ado_Memo.Recordset!mes_descuento & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "'", "ro_pagos_cronograma_Detalle") Then
      sino = MsgBox("No se puede ELIMINAR porque ya fue Procesado en la planilla. Desea marcar como ERRADO ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_Memo.Recordset!estado_codigo = "ERR"
         Ado_Memo.Recordset!fecha_registro = Date
         Ado_Memo.Recordset!usr_codigo = glusuario
         Ado_Memo.Recordset.UpdateBatch adAffectAll
      End If
   Else
      sino = MsgBox("Está Seguro de ELIMINAR fisicamente el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          db.Execute "DELETE ro_memorandas where ges_gestion = " & Ado_Memo.Recordset!ges_gestion & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "' AND numero = " & Ado_Memo.Recordset!Numero
      '''''''''''Call opciones
      End If
   End If
  Call abrirtabla
  
   Exit Sub
Else
sino = MsgBox("Está Seguro de ELIMINAR fisicamente el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         db.Execute "DELETE ro_memorandas where ges_gestion = " & Ado_Memo.Recordset!ges_gestion & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "' AND numero = " & Ado_Memo.Recordset!Numero
      Call abrirtabla
       '''''''''''Call opciones
      End If
End If
Exit Sub
UpdateErr:
  MsgBox Err.Description
  Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
   
   
   
   
End Sub

Private Sub CmdElim5_Click()
 On Error GoTo EditErr
   sino = MsgBox("Está Seguro de BORRAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If AdoMovilidad.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        db.Execute "delete ro_movilidad_personal where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and numero_cambio = " & AdoMovilidad.Recordset!numero_cambio & " "
'        AdoMovilidad.Recordset("estado_codigo") = "ANL"
'        AdoMovilidad.Recordset("fecha_registro") = Date
'        AdoMovilidad.Recordset("usr_codigo") = glusuario
'        AdoMovilidad.Recordset.Update  'Batch adAffectAll
        Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub


'Private Sub CmdGraba_Click()
'   If Ado_datos.Recordset.RecordCount > 0 And Not IsNull(DtgVacacionesProg.Columns("nivel_educacional").Value) And (DtgVacacionesProg.Columns("nivel_educacional").Value) <> "" Then
'      If Ado_datos.Recordset!estado_codigo = "REG" Then
'        marca1 = Ado_datos.Recordset.Bookmark
'        VARB = DtgVacacionesProg.Columns("beneficiario_codigo").Value
'        VARBD = DtgVacacionesProg.Columns("Carrera_Curso").Value
'        VARG = DtgVacacionesProg.Columns("centro_educativo").Value
'        VARS = DtgVacacionesProg.Columns("titulo_obtenido").Value
'        VARU = DtgVacacionesProg.Columns("nivel_educacional").Value
'        VARPU = DtgVacacionesProg.Columns("duracion_años").Value
'        VAR10 = DtgVacacionesProg.Columns("pais").Value
'        VAR11 = DtgVacacionesProg.Columns("ciudad").Value
'        VAR12 = DtgVacacionesProg.Columns("fecha_inicio").Value
'        VAR13 = DtgVacacionesProg.Columns("fecha_fin").Value
'        VAR14 = DtgVacacionesProg.Columns("presento_documento").Value
''        MarcaB = adoao_solicitud_bien.Recordset.Bookmark
''        Call Abre_Sol_Bien
''        'MarcaB = rs_ao_solicitud_bien.Bookmark
''        adoao_solicitud_bien.Recordset.Bookmark = MarcaB
'        rs_datos_educacionales!beneficiario_codigo = VARB
'        rs_datos_educacionales!Carrera_Curso = VARBD
'        rs_datos_educacionales!centro_educativo = VARG
'        rs_datos_educacionales!titulo_obtenido = VARS
'        rs_datos_educacionales!nivel_educacional = VARU
'        rs_datos_educacionales!duracion_años = VARPU
'        rs_datos_educacionales!pais = VAR10
'        rs_datos_educacionales!ciudad = VAR11
'        rs_datos_educacionales!fecha_inicio = IIf(VAR12 = "", Date, VAR12)
'        rs_datos_educacionales!fecha_fin = VAR13
'        rs_datos_educacionales!presento_documento = VAR14
'        rs_datos_educacionales.Update
'        'Call Abre_Sol_Bien
'        rs_datos_educacionales.MoveLast
''        Call OptFilGral1_Click
'        'adosolicitud.Recordset.BookMark = marca1
'        'adosolicitud.Refresh
'        'swgrabar = 2
'        DtgVacacionesProg.AllowAddNew = False
'        DtgVacacionesProg.AllowDelete = False
'        DtgVacacionesProg.AllowUpdate = False
'        CmdAdd2.Visible = True
'        CmdMod2.Visible = True
'        CmdGraba.Visible = False
'      Else
'         MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Personal"
'      End If
'   Else
'         MsgBox "Verifique los datos para continuar ... ", vbInformation, "Personal"
'   End If
'
'End Sub


'Private Sub CmdGraba2_Click()
'   If Ado_datos.Recordset.RecordCount > 0 And Not IsNull(DtgVacaciones.Columns("tipo_institucion").Value) And (DtgVacaciones.Columns("tipo_institucion").Value) <> "" Then
'      If Ado_datos.Recordset!estado_codigo = "REG" Then
'        marca1 = Ado_datos.Recordset.Bookmark
'        VARB = DtgVacaciones.Columns("beneficiario_codigo").Value
'        VARBD = DtgVacaciones.Columns("denominacion_institucion").Value
'        VARG = DtgVacaciones.Columns("tipo_institucion").Value
'        VARS = DtgVacaciones.Columns("cargo").Value
'        VARU = DtgVacaciones.Columns("funcion_general").Value
'        VARPU = DtgVacaciones.Columns("Tiempo_Meses").Value
'        VAR10 = DtgVacaciones.Columns("pais").Value
'        VAR11 = DtgVacaciones.Columns("ciudad").Value
'        VAR12 = DtgVacaciones.Columns("fecha_inicio").Value
'        VAR13 = DtgVacaciones.Columns("fecha_fin").Value
'        VAR14 = DtgVacaciones.Columns("presento_documento").Value
''        MarcaB = adoao_solicitud_bien.Recordset.Bookmark
''        Call Abre_Sol_Bien
''        'MarcaB = rs_ao_solicitud_bien.Bookmark
''        adoao_solicitud_bien.Recordset.Bookmark = MarcaB
'        rs_laborales!beneficiario_codigo = VARB
'        rs_laborales!denominacion_institucion = VARBD
'        rs_laborales!tipo_institucion = VARG
'        rs_laborales!cargo = VARS
'        rs_laborales!funcion_general = VARU
'        rs_laborales!Tiempo_Meses = VARPU
'        rs_laborales!pais = VAR10
'        rs_laborales!ciudad = VAR11
'        rs_laborales!fecha_inicio = IIf(VAR12 = "", Date, VAR12)
'        rs_laborales!fecha_fin = VAR13
'        rs_laborales!presento_documento = VAR14
'        rs_laborales.Update
'        'Call Abre_Sol_Bien
'        rs_laborales.MoveLast
''        Call OptFilGral1_Click
'        'adosolicitud.Recordset.BookMark = marca1
'        'adosolicitud.Refresh
'        'swgrabar = 2
'        DtgVacaciones.AllowAddNew = False
'        DtgVacaciones.AllowDelete = False
'        DtgVacaciones.AllowUpdate = False
'        CmdAdd2.Visible = True
'        CmdMod2.Visible = True
'        CmdGraba2.Visible = False
'      Else
'         MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Personal"
'      End If
'   Else
'         MsgBox "Verifique los datos para continuar ... ", vbInformation, "Personal"
'   End If
'
'End Sub

'Private Sub CmdGrabaCto_Click()
'  On Error GoTo UpdateErr
'  VAR_VAL = "OK"
''  Call valida_campos
'  If VAR_VAL = "OK" Then
'    If GlSW = "ADD" Then
'      rs_contrato!codigo_contrato = txtCodigo.Text
'      rs_contrato!beneficiario_codigo = Ado_datos.Recordset("beneficiario_codigo") 'DtcBenef.Text
'      rs_contrato!ges_gestion = glGestion
'      rs_contrato!codigo_solicitud = rs_contrato.RecordCount
'
'      Set rs_correlativo = New ADODB.Recordset
'      rs_correlativo.Open "select * from ro_contratos_personas WHERE beneficiario_codigo = '" & Ado_datos.Recordset("beneficiario_codigo") & "'  ", DB, adOpenKeyset, adLockOptimistic
'      If rs_correlativo.RecordCount > 0 Then
'            rs_contrato!numero_consultoria = rs_correlativo.RecordCount
''            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
''            rs_correlativo.Update
''            rs_M1!Numero_FA = rs_correlativo!correlativo
'      Else
'            rs_contrato!numero_consultoria = 1
'      End If
'      rs_contrato!ARCHIVO = "Cargar_Archivo"
'      rs_contrato!ARCHIVO_NOMB = Trim(Ado_datos.Recordset("beneficiario_beneficiario_iniciales")) & "_Contrato_" & rs_contrato!numero_consultoria & ".pdf"
'      TxtAprob.Text = "REG"
'    End If
'      rs_contrato!objeto_contrato = txtObjContrato.Text
'      rs_contrato!codigo_puesto = DtcPuesto.Text
'      rs_contrato!codigo_unidad = Dtc_codigo.Text
'      rs_contrato!codigo_convenio = DtcOrg.Text
'      rs_contrato!pro_proyecto = DtcPry.Text
'      rs_contrato!fechas_confirmado = Txtestado
'      rs_contrato!estAdo_Memo = TxtAprob
'      rs_contrato!fecha_firma = DTPFFirma.Value
'      rs_contrato!fecha_inicio = DTPFInicio.Value
'      rs_contrato!fecha_fin = DTPFFin.Value
'      rs_contrato!monto_totalbs = TxtBs.Text
'      If GlTipoCambioOficial > 0 Then
'        rs_contrato!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      Else
'        GlTipoCambioOficial = 7.05
'        rs_contrato!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      End If
'      rs_contrato!observacion_contrato = "-"
'      rs_contrato!establece_multas = "N"
'      rs_contrato!cod_forma_inicio = "1"
'      rs_contrato!tiempo_num = 0
'      rs_contrato!tiempo_dmy = "-"
'      rs_contrato!tipo_moneda = "Bs"
'      rs_contrato!tc_us = GlTipoCambioOficial
'
'      rs_contrato!org_codigo = "111"
'      rs_contrato!porc_orgfin = 100
'      rs_contrato!porc_contra = 0
'      'rs_contrato!fechas_confirmado = "N"
'      rs_contrato!hora_registro = "8:00"
'      rs_contrato!fecha_registro = Date
'      rs_contrato!usr_usuario = "ADMIN" 'GlUsuario
'      rs_contrato.Update    'Batch adAffectAll
'
''      mbDataChanged = False
'      CmdAddCto.Visible = True
'      CmdModCto.Visible = True
'      CmdGrabaCto.Visible = False
'      CmdAprCto.Visible = True
'      TxtAprob.Enabled = True
'      Fra_ABM.Enabled = False
'      DtG_Auxiliar.Enabled = False
'      GlSW = " "
'
'  End If
'  Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'End Sub

Private Sub CmdMod2_Click()

On Error GoTo EditErr
     If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_VacacionesProg.Recordset!estado_codigo = "REG" Or glusuario = "VPAREDES" Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ao_Vacacion_Prog.txtSW = "MOD"
        frm_ao_Vacacion_Prog.sel = 1
        frm_ao_Vacacion_Prog.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ao_Vacacion_Prog.TxtGestion.Text = Ado_VacacionesProg.Recordset!ges_gestion
        
        frm_ao_Vacacion_Prog.Txt01.Text = Ado_VacacionesProg.Recordset!mes_control
        frm_ao_Vacacion_Prog.Txt02.Text = Ado_VacacionesProg.Recordset!dias_Programados
        frm_ao_Vacacion_Prog.txt03.Value = Ado_VacacionesProg.Recordset!fecha_ini_Prog
        frm_ao_Vacacion_Prog.txt04.Value = Ado_VacacionesProg.Recordset!fecha_fin_Prog
        frm_ao_Vacacion_Prog.DTPFec_Inicio.Value = Ado_VacacionesProg.Recordset!fecha_registro
        frm_ao_Vacacion_Prog.DtcFec_Fin.Value = Ado_VacacionesProg.Recordset!fecha_reincoporacion
        frm_ao_Vacacion_Prog.txt05.Value = Ado_VacacionesProg.Recordset!horadesde
        frm_ao_Vacacion_Prog.txt06.Value = Ado_VacacionesProg.Recordset!HoraHasta
        frm_ao_Vacacion_Prog.txt07.Value = Ado_VacacionesProg.Recordset!Hora_reincorporacion
        'frm_ao_Vacacion_Prog.Txt09.Text = Ado_VacacionesProg.Recordset!horas_Programadas
        frm_ao_Vacacion_Prog.txt10.Text = Ado_VacacionesProg.Recordset!minutos_programados
        'frm_ao_Vacacion_Prog.Dtc_Par.Text = Ado_VacacionesProg.Recordset!nivel_educacional
        'frm_ao_Vacacion_Prog.Dtc_ParDes.BoundText = ac_CapturaEstudiosRealizados.Dtc_Par.BoundText
        frm_ao_Vacacion_Prog.txtEstado.Text = Ado_VacacionesProg.Recordset!estado_codigo
        frm_ao_Vacacion_Prog.txt_dias_vac.Text = Ado_VacacionesProg.Recordset!dias_Programados
        frm_ao_Vacacion_Prog.txt_empresa.Text = Ado_datos.Recordset!codigo_empresa
        frm_ao_Vacacion_Prog.Show vbModal
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Personal"
      End If
      Ado_VacacionesProg.Refresh
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
'    CmdAdd.Visible = False
'    CmdMod.Visible = False
'    CmdGraba.Visible = True

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdMod3_Click()
 On Error GoTo EditErr
 If AdoPermiso.Recordset!estado_codigo = "REG" Then
    marca1 = Ado_datos.Recordset.Bookmark
    frm_ao_Permisos_js.txtSW = "MOD"
    frm_ao_Permisos_js.txtBenef = Ado_datos.Recordset!beneficiario_codigo
    'frm_ao_Permisos.TxtInicial = Ado_datos.Recordset!beneficiario_beneficiario_iniciales
    frm_ao_Permisos_js.lblARCH.Caption = AdoPermiso.Recordset!ARCHIVO
    frm_ao_Permisos_js.Dtc_Par.BoundText = AdoPermiso.Recordset!TipoPermiso
    frm_ao_Permisos_js.dt_fechasolicitusper = AdoPermiso.Recordset!Fecha_control
    frm_ao_Permisos_js.cmb_mescontrol = AdoPermiso.Recordset!mes_control
'    frm_ao_Permisos_js.txt02 = AdoPermiso.Recordset!dia_control
    frm_ao_Permisos_js.dt_fechadesde = AdoPermiso.Recordset!FechaDesde
    frm_ao_Permisos_js.dt_fechahasta = AdoPermiso.Recordset!FechaHasta
    frm_ao_Permisos_js.dt_fechareincorporacion = AdoPermiso.Recordset!fecha_reincorporacion
    frm_ao_Permisos_js.hr_horadesde = AdoPermiso.Recordset!horadesde
    frm_ao_Permisos_js.hr_horahasta = AdoPermiso.Recordset!HoraHasta
    frm_ao_Permisos_js.hr_horareincorporacion = AdoPermiso.Recordset!Hora_reincorporacion
    frm_ao_Permisos_js.TxtGestion = AdoPermiso.Recordset!ges_gestion
    frm_ao_Permisos_js.txt_nrodias = AdoPermiso.Recordset!dias_permiso
    frm_ao_Permisos_js.txt_nrohoras = AdoPermiso.Recordset!horas_permiso
    frm_ao_Permisos_js.txt_nrominutos = AdoPermiso.Recordset!minutos_permiso
    'frm_ao_Permisos.Dtc_ParDes = AdoPermiso.Recordset!nomb_pariente
    frm_ao_Permisos_js.txtEstado = AdoPermiso.Recordset!estado_codigo
    frm_ao_Permisos_js.cmb_tipopermiso.BoundText = frm_ao_Permisos_js.Dtc_Par.BoundText
    frm_ao_Permisos_js.cmb_tipopermiso.BoundText = AdoPermiso.Recordset!TipoPermiso
    frm_ao_Permisos_js.Show vbModal
    
 Else
        MsgBox "No se puede MODIFICAR un registro Aprobado o Anulado ...", vbExclamation, "Validación de Registro"
 End If
 Call abrirtabla
 
 Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdMod4_Click()
  On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_Memo.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ao_memoranda.txtSW = "MOD"
        frm_ao_memoranda.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        'frm_ao_memoranda.TxtInicial = Ado_datos.Recordset!beneficiario_beneficiario_iniciales
        frm_ao_memoranda.Txt_Correl.Text = IIf(IsNull(Ado_Memo.Recordset!CORREL), "1", Ado_Memo.Recordset!CORREL)
        frm_ao_memoranda.Dtc_Par.Text = Ado_Memo.Recordset!tipo_memo
        frm_ao_memoranda.Dtc_ParDes.BoundText = frm_ao_memoranda.Dtc_Par.BoundText
        frm_ao_memoranda.Txt09.Text = IIf(IsNull(Ado_Memo.Recordset!observaciones), "-", Ado_Memo.Recordset!observaciones)
        frm_ao_memoranda.DTPFec_Inicio.Value = Ado_Memo.Recordset!fecha_memo
        frm_ao_memoranda.DtcFec_Fin.Value = Ado_Memo.Recordset!fecha_aprobacion
        frm_ao_memoranda.TxtGestion.Text = Ado_Memo.Recordset!ges_gestion
        frm_ao_memoranda.txt08.Text = IIf(IsNull(Ado_Memo.Recordset!Monto), "0", Ado_Memo.Recordset!Monto)
        frm_ao_memoranda.txt10.Text = Ado_Memo.Recordset!minutos
        frm_ao_memoranda.TxtGestion2.Text = Ado_Memo.Recordset!gestion_descuento
        frm_ao_memoranda.Txt01.Text = Ado_Memo.Recordset!mes_descuento
        frm_ao_memoranda.lblARCH.Caption = Ado_Memo.Recordset!ARCHIVO
        frm_ao_memoranda.txt_memo.Caption = Ado_Memo.Recordset!CORREL
        frm_ao_memoranda.txtEstado = Ado_Memo.Recordset!estado_codigo
        frm_ao_memoranda.cbo_dias.Text = Ado_Memo.Recordset!DIAS
        If Ado_Memo.Recordset!descuento_pla = "SI" Then
        frm_ao_memoranda.optSi.Value = True
        Else
        frm_ao_memoranda.optNo.Value = True
        End If
        frm_ao_memoranda.Show vbModal
        'Ado_Memo.Refresh
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Personal"
      End If
      Call abrirtabla
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
   Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdMod1_Click()
On Error GoTo EditErr
 If AdoAsistencia.Recordset!estado_codigo = "REG" Then
    marca1 = Ado_datos.Recordset.Bookmark
    frm_ao_Asistencia.txtSW = "MOD"
    frm_ao_Asistencia.txtBenef = Ado_datos.Recordset!beneficiario_codigo
    frm_ao_Asistencia.DTPFec_Inicio = AdoAsistencia.Recordset!Fecha_control
    frm_ao_Asistencia.Txt01 = AdoAsistencia.Recordset!mes_control
    frm_ao_Asistencia.Txt02 = AdoAsistencia.Recordset!dia_control
    frm_ao_Asistencia.txt03 = AdoAsistencia.Recordset!HoraUno
    frm_ao_Asistencia.txt04 = AdoAsistencia.Recordset!HoraDos
    frm_ao_Asistencia.txt05 = AdoAsistencia.Recordset!Atraso
    frm_ao_Asistencia.txt06 = AdoAsistencia.Recordset!Falta
    frm_ao_Asistencia.txt07 = AdoAsistencia.Recordset!HoraTres
    frm_ao_Asistencia.txt08 = AdoAsistencia.Recordset!HoraCuatro
    frm_ao_Asistencia.Cmb01 = AdoAsistencia.Recordset!AtrasoI
    frm_ao_Asistencia.Cmb02 = AdoAsistencia.Recordset!Falta2
    frm_ao_Asistencia.Dtc_ParDes = IIf(IsNull(AdoAsistencia.Recordset!turno), "AM", AdoAsistencia.Recordset!turno)
    frm_ao_Asistencia.Dtc_Par.BoundText = frm_ao_Asistencia.Dtc_ParDes.BoundText
    frm_ao_Asistencia.Dtc_Par3 = IIf(IsNull(AdoAsistencia.Recordset!turno2), "PM", AdoAsistencia.Recordset!turno2)
    frm_ao_Asistencia.Dtc_Par2.BoundText = frm_ao_Asistencia.Dtc_Par3.BoundText
    frm_ao_Asistencia.txtEstado = AdoAsistencia.Recordset!estado_codigo
    'AdoAsistencia.Recordset.AddNew
    frm_ao_Asistencia.Show vbModal
 Else
        MsgBox "No se puede MODIFICAR un registro Aprobado o Anulado ...", vbExclamation, "Validación de Registro"
 End If
 Call abrirtabla
Exit Sub

EditErr:
  MsgBox Err.Description
  
End Sub

Private Sub CmdMod5_Click()
  On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
      If AdoMovilidad.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ro_movilidad_personal.txtSW = "MOD"
        
        frm_ro_movilidad_personal.TxtCodigo.Text = AdoMovilidad.Recordset!numero_cambio
          frm_ro_movilidad_personal.TxtForm.Text = AdoMovilidad.Recordset!numero_resolucion
          frm_ro_movilidad_personal.DtcRespaldoCod.Text = AdoMovilidad.Recordset!tipo_memo
          frm_ro_movilidad_personal.DtcRespaldo.BoundText = frm_ro_movilidad_personal.DtcRespaldoCod.BoundText
          frm_ro_movilidad_personal.txtObjContrato.Text = AdoMovilidad.Recordset!observaciones
          frm_ro_movilidad_personal.Dtc_codigo_ant.Text = AdoMovilidad.Recordset!unidad_anterior
          frm_ro_movilidad_personal.Dtc_descrip_ant.BoundText = frm_ro_movilidad_personal.Dtc_codigo_ant.BoundText
          frm_ro_movilidad_personal.DtcOrg.Text = AdoMovilidad.Recordset!cargo_anterior
          frm_ro_movilidad_personal.DtcOrgDes.BoundText = frm_ro_movilidad_personal.DtcOrg.BoundText
          frm_ro_movilidad_personal.DtcPry.Text = AdoMovilidad.Recordset!puesto_anterior
          frm_ro_movilidad_personal.DtcPryDes.BoundText = frm_ro_movilidad_personal.DtcPry.BoundText
          frm_ro_movilidad_personal.dtc_codigo.Text = AdoMovilidad.Recordset!unidad_codigo
          frm_ro_movilidad_personal.Dtc_descrip.BoundText = frm_ro_movilidad_personal.dtc_codigo.BoundText
          frm_ro_movilidad_personal.DtcCargo.Text = AdoMovilidad.Recordset!cargo_codigo
          frm_ro_movilidad_personal.DtcCargoDes.BoundText = frm_ro_movilidad_personal.DtcCargo.BoundText
          frm_ro_movilidad_personal.DtcPuesto.Text = AdoMovilidad.Recordset!puesto_codigo
          frm_ro_movilidad_personal.DtcPuestoDes.BoundText = frm_ro_movilidad_personal.DtcPuesto.BoundText
          
          frm_ro_movilidad_personal.dtc_beneficiario_den.BoundText = AdoMovilidad.Recordset!beneficiario_codigo_int
          frm_ro_movilidad_personal.txt_tipo_mov = AdoMovilidad.Recordset!tipo_mov
          
          frm_ro_movilidad_personal.DTPFelaboracion = AdoMovilidad.Recordset!fecha_elaboracion
          frm_ro_movilidad_personal.DTPFcontrato = AdoMovilidad.Recordset!fecha_inicio_contrato
    '      frm_ro_movilidad_personal.DTPFaprobacion = AdoMovilidad.Recordset!fecha_aprobacion
'          frm_ro_movilidad_personal.TxtBs.Text = AdoMovilidad.Recordset!Item
          frm_ro_movilidad_personal.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ro_movilidad_personal.Show vbModal
        'Ado_Memo.Refresh
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Personal"
      End If
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
   Exit Sub
EditErr:
  MsgBox Err.Description
End Sub


Private Sub CmdVerDisco_Click()
'  On Error GoTo Error_Sub
'    marca1 = Ado_datos.Recordset.Bookmark
'    FrmExplora.lblges_gestion = Ado_datos.Recordset!primer_apellido + " " + Ado_datos.Recordset!segundo_apellido + " " + Ado_datos.Recordset!NombreS
'    FrmExplora.LblFA = Ado_datos.Recordset!beneficiario_codigo
'    'FrmExplora.LblForm = Ado_datos.Recordset!tipo_formulario
''    sino = MsgBox("Elija <SI> para ver la Información de su Disco Local. , o del Servidor <NO> ", vbQuestion + vbYesNo, "Confirmando...")
''    If sino = vbYes Then
'    NombreCarpeta = App.Path & "\PERSONAL\" & Ado_datos.Recordset!beneficiario_codigo
'    e = App.Path & "\PERSONAL\" & Ado_datos.Recordset!beneficiario_codigo
''    NombreCarpeta = App.Path & "\PERSONAL\" & Ado_datos.Recordset!beneficiario_beneficiario_iniciales
''    e = App.Path & "\PERSONAL\" & Ado_datos.Recordset!beneficiario_beneficiario_iniciales
''    If MsgBox("- Elija 'Si' para ver la Información de su Disco Local ..." & vbCrLf & _
''             "- Elija 'No' para ver la Información del SERVIDOR ... ", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
'        FrmExplora.Dir1.Path = NombreCarpeta
'        FrmExplora.Label1 = NombreCarpeta
''    Else
'        FrmExplora.Dir1.Path = e
'        FrmExplora.Label1 = e
''    End If
'    FrmExplora.Show 'vdmodal
'Exit Sub
'Error_Sub:
' MsgBox Err.Description, vbCritical
End Sub

Private Sub graba_persona()
On Error GoTo EditErr
    Set rsauxiliar = New ADODB.Recordset
    'SQL_FOR = "select * from rc_personal where ci = '" & txtCodigo.Text & "'"
    'rsauxiliar.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
    rsauxiliar.Open "select * from rc_personal where ci = '" & TxtCodigo & "' ", db, adOpenKeyset, adLockOptimistic
    If rsauxiliar.RecordCount = 0 Then
        rsauxiliar.AddNew
        rsauxiliar!ci = TxtCodigo
        rsauxiliar!idfuncionario = CORREL
    Else
        'MsgBox " YA EXISTE EL CODIGO ..."
    End If
        rsauxiliar!tipoben_codigo = Trim(TxtTipo.Text) 'JQA NOV-2009
        rsauxiliar!ruc_id = TxtNIT
        rsauxiliar!Fecha_Nacimiento = DTP_FechaNac.Value
        'rsauxiliar!calle_domicilio = TxtDireccion
'        rsauxiliar!zona_domicilio = TxtZona.Text
'        rsauxiliar!Telefono = TxtTelefono.Text
        rsauxiliar!Status = "S"
        rsauxiliar!Activo = "S"
        rsauxiliar!usr_usuario = glusuario 'frmLogin.txtUserName.Text
        rsauxiliar!fecha_registro = Date
        rsauxiliar!hora_registro = Format(Time, "HH:mm:ss")
        rsauxiliar!departamento_nacimiento = dtc_depto.Text
        rsauxiliar!Procedencia = Dtc_prov.Text
        rsauxiliar!lugar_procedencia = Dtc_munic.Text
        'rsauxiliar!codigo_cargo = "-"   'TxtCargo.Text
'        rsauxiliar!numero_folder = Txt_mail.Text
        rsauxiliar!profesion = TxtProfesion.Text
        rsauxiliar.Update
        MkDir TxtCodigo
        If Guardar_Imagen(db, "Select Foto From rv_personal_contratado Where beneficiario_codigo= '" & Ado_datos.Recordset("beneficiario_codigo") & "' ", "Foto", App.Path) Then
            MsgBox "ok"
        Else
            MsgBox "ERR"
        End If
        'Guardar_Imagen(cn, Sql, Campo, Path_Imagen)
    
Exit Sub

EditErr:
  MsgBox Err.Descriptio
End Sub

Private Sub BtnAñadir_Click()
On Error GoTo EditErr
   swnuevo = 1
   Ado_datos.Recordset.AddNew
   Set rst_ben = New ADODB.Recordset
'   If MsgBox("- Elija 'Si' para registrar la ENTIDAD (Empresa o Institución) ..." & vbCrLf & _
'             "- Elija 'No' para registrar Consultores o Funcionarios  ", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
'      SSTab1.Tab = 1
'      SSTab1.TabEnabled(1) = True
'      SSTab1.TabEnabled(0) = False
''      fraDatos2.Enabled = True
'      Frame22.Enabled = True
'        txtCodigo2.Text = Empty
'        DtcRep_Paterno.Text = Empty
'        DtcRep_Materno.Text = Empty
'        DtcRep_Nombres.Text = Empty
'        TxtTipo.Text = Empty
'        txtDenominacion2.Text = Empty
'        txtDenominacion2.Enabled = True
'        'Carga_Recor
'        TDBtipoben2.SetFocus
'      rst_ben.Open "SELECT * FROM gc_Tipo_Beneficiario where estado_codigo ='B' ORDER BY descripcion ", db, adOpenStatic
'      Call rep_legal
'   Else
      SSTab1.Tab = 0
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = False
      TxtCodigo.Enabled = True
      fraDatos.Enabled = True
      Frame2.Enabled = True
      TxtCodigo = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
      TxtTipo.Text = Empty
      txtDenominacion = Empty
        'Carga_Recor
      TDBtipoben.SetFocus
      rst_ben.Open "SELECT * FROM gc_Tipo_Beneficiario where estado_codigo ='APR' ORDER BY descripcion ", db, adOpenStatic
'   End If
    Set AdoTip_ben.Recordset = rst_ben
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = False
'    FraSS_SS.Enabled = True
    CmdAdd1.Visible = False
    CmdMod1.Visible = False
    CmdElim1.Visible = False
    CmdApr1.Visible = False
    CmdAdd2.Visible = False
    CmdMod2.Visible = False
    CmdElim2.Visible = False
    CmdApr2.Visible = False
    CmdAdd3.Visible = False
    CmdMod3.Visible = False
    CmdElim3.Visible = False
    CmdApr3.Visible = False
    CmdAdd4.Visible = False
    CmdMod4.Visible = False
    CmdElim4.Visible = False
    CmdApr4.Visible = False
    CmdAdd5.Visible = False
    CmdMod5.Visible = False
    CmdElim5.Visible = False
    CmdApr5.Visible = False
'    CmdAdd6.Visible = False
'    CmdMod6.Visible = False
'    CmdElim6.Visible = False
'    CmdApr6.Visible = False
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

'Private Sub rep_legal()
'   Set rs_RepLegal = New ADODB.Recordset
'   If rs_RepLegal.State = 1 Then rs_RepLegal.Close
'   rs_RepLegal.Open "select * from gc_Beneficiario WHERE tipoben_codigo = '5' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'   rs_RepLegal.Sort = "beneficiario_denominacion"
'   'If rs_RepLegal.RecordCount > 0 Then
'    Set AdoRepLegal.Recordset = rs_RepLegal
'    AdoRepLegal.Refresh
'   'End If
'End Sub

'Private Sub BtnEliminar_Click()
'   Dim Mensaje As String
'
'On Error GoTo errorDelete
'
'   Mensaje = "¿Borrar: " & _
'               txtCodigo.Text & " " & _
'               Trim(txtDenominacion.Text) & "?"
'   If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar:") = vbYes Then
'      db.BeginTrans
'      Ado_datos.Recordset.Delete
'      db.CommitTrans
'   End If
'
'   Exit Sub
'errorDelete:
'
'   Dim e As ADODB.Error
'
'   For Each e In db.Errors
'      MsgBox "Error No. " & e.Number & " " & e.Description
'   Next
'
'   db.RollbackTrans
'
'End Sub

Private Sub BtnBuscar_Click()
On Error GoTo EditErr
'Fra_Busqueda.Visible = True
'fradatos.Enabled = True
' Set ClBuscaGrid = New ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = db
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = dg_datos
'    ClBuscaGrid.QueryUtilizado = queryinicial
'    Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = False
  'Dim GrSqlAux As String
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = dg_datos
  ClBuscaGrid.QueryUtilizado = queryinicial
  'Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
  Set ClBuscaGrid.RecordsetTrabajo = rstbeneficiario.DataSource
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_datos.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        If Ado_datos.Recordset("no_file") <> 1 Then
            Dim RUTA1, RUTA2 As String
            RUTA1 = "PERSONAL" + "\" + Trim(Ado_datos.Recordset("beneficiario_beneficiario_iniciales")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
            MsgBox RUTA1
            MkDir RUTA1
            MkDir RUTA1 + "\CONTRATOS"
            MkDir RUTA1 + "\FINIQUITO"
            MkDir RUTA1 + "\MEMORANDUMS"
            MkDir RUTA1 + "\DOCUMENTOS_RESPALDO"
            MkDir RUTA1 + "\HOJA_VIDA"
            MkDir RUTA1 + "\OTROS"
            MkDir RUTA1 + "\EVALUACIONES"
            MkDir RUTA1 + "\LICENCIAS"
            MkDir RUTA1 + "\VACACIONES"
'
'            RUTA1 = "PERSONAL" + "\" + Text1 + " " + Text2 + " " + Text3
'            MsgBox RUTA1
'            MkDir RUTA1
            
'            RUTA1 = "PERSONAL" + "\" + Trim(LblInicial)
'            MsgBox RUTA1
'            MkDir RUTA1
            Ado_datos.Recordset("no_file") = 1
        End If
        Ado_datos.Recordset("estado_codigo") = "APR"
        Ado_datos.Recordset("fecha_aprueba") = Date
        Ado_datos.Recordset("usr_aprueba") = glusuario
        Ado_datos.Recordset.Update
        
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub


Private Sub BtnEliminar_Click()
On Error GoTo EditErr
   sino = MsgBox("Está Seguro de ANULAR el Registro?", vbYesNo + vbQuestion, "Atención")
   If Ado_datos.Recordset("estado_codigo") = "APR" Or Ado_datos.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_datos.Recordset("estado_codigo") = "ANL"
        Ado_datos.Recordset("fecha_aprueba") = Date
        Ado_datos.Recordset("usr_codigo_apr") = glusuario
        Ado_datos.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnCancelar_Click()
On Error GoTo EditErr
  On Error Resume Next
VAR_COD2 = Ado_datos.Recordset!beneficiario_codigo
Ado_datos.Recordset.CancelUpdate
    swnuevo = 0
   fraOpciones.Visible = True
       fra_cabecera.Enabled = True
       
        SSTab1.TabEnabled(5) = True
       SSTab1.TabEnabled(4) = True
            SSTab1.TabEnabled(3) = True
         SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = True
            
   FraGrabarCancelar.Visible = False
   FraNavega.Enabled = True
   fraDatos.Enabled = False
   TxtCodigo.Enabled = True
'   FraSS_SS.Enabled = False
   CmdAdd1.Visible = False
   CmdMod1.Visible = False
   CmdElim1.Visible = False
   CmdApr1.Visible = False
   CmdAdd2.Visible = False
   CmdMod2.Visible = False
   CmdElim2.Visible = False
   CmdApr2.Visible = False
   CmdAdd3.Visible = False
   CmdMod3.Visible = False
   CmdElim3.Visible = False
   CmdApr3.Visible = False
   CmdAdd4.Visible = False
   CmdMod4.Visible = False
   CmdElim4.Visible = False
   CmdApr4.Visible = False
   CmdAdd5.Visible = False
   CmdMod5.Visible = False
   CmdElim5.Visible = False
   CmdApr5.Visible = False
'   CmdAdd6.Visible = False
'   CmdMod6.Visible = False
'   CmdElim6.Visible = False
'   CmdApr6.Visible = False

'   Call Carga_Recor
'   Call Carga_Beneficiario

 If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rstbeneficiario.Find "beneficiario_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rstbeneficiario.Bookmark)
     Else
        rstbeneficiario.MoveLast
     End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
 
  Public Sub encontrar()
  If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If

     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rstbeneficiario.Find "beneficiario_codigo = '" & rw_datos_extra.txt_nuevo_num.Text & "'   ", , , 1
        dg_datos.SelBookmarks.Add (rstbeneficiario.Bookmark)
     Else
        rstbeneficiario.MoveLast
     End If

  End Sub
 
 Public Sub opciones()
 On Error GoTo EditErr
   TxtCodigo.Enabled = True
   fraDatos.Enabled = False
   'Carga_Recor
   swnuevo = 0
   
   fraOpciones.Visible = True
   FraGrabarCancelar.Visible = False
   FraNavega.Enabled = True
'  FraSS_SS.Enabled = False
'''   CmdAdd1.Visible = False
'''   CmdMod1.Visible = False
'''   CmdElim1.Visible = False
'''   CmdApr1.Visible = False
'''   CmdAdd2.Visible = False
'''   CmdMod2.Visible = False
'''   CmdElim2.Visible = False
'''   CmdApr2.Visible = False
'''   CmdAdd3.Visible = False
'''   CmdMod3.Visible = False
'''   CmdElim3.Visible = False
'''   CmdApr3.Visible = False
'''   CmdAdd4.Visible = False
'''   CmdMod4.Visible = False
'''   CmdElim4.Visible = False
'''   CmdApr4.Visible = False
'''   CmdAdd5.Visible = False
'''   CmdMod5.Visible = False
'''   CmdElim5.Visible = False
'''   CmdApr5.Visible = False
'   CmdAdd6.Visible = False
'   CmdMod6.Visible = False
'   CmdElim6.Visible = False
'   CmdApr6.Visible = False
'   Call Carga_Recor
'   Call Carga_Beneficiario
'   Ado_datos.Recordset.Requery
'   Ado_datos.Refresh
'   dg_datos.ReBind
'   dg_datos.Refresh
Exit Sub

EditErr:
  MsgBox Err.Description
 End Sub
 
 
Private Sub BtnModificar_Click()
On Error GoTo EditErr
'  If Ado_datos.Recordset("estado_codigo") = "N" Then
     swnuevo = 2
     Set rst_ben = New ADODB.Recordset
     SSTab1.Tab = 0
     SSTab1.TabEnabled(0) = True
     SSTab1.TabEnabled(1) = True
     SSTab1.TabEnabled(2) = True
     If Ado_datos.Recordset("estado_codigo") = "APR" Then
        MsgBox "El registro está APROBADO, solo se puede modificar por usuarios Autorizados ..."
        Frame2.Enabled = False
'        Frame1.Enabled = False
     Else
        Frame2.Enabled = True
'        Frame1.Enabled = True
     End If
     fraDatos.Enabled = True
     DTP_FechaNac.Enabled = True
'     TxtRenca.SetFocus
'     FraSS_SS.Enabled = True
            
            SSTab1.TabEnabled(5) = False
            SSTab1.TabEnabled(4) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(1) = False
           ' SSTab1.TabEnabled(0) = False

 
     CmdAdd1.Visible = True
     CmdMod1.Visible = True
     CmdElim1.Visible = True
     CmdApr1.Visible = True
     CmdAdd2.Visible = True
     CmdMod2.Visible = True
     CmdElim2.Visible = True
     CmdApr2.Visible = True
     CmdAdd3.Visible = True
     CmdMod3.Visible = True
     CmdElim3.Visible = True
     CmdApr3.Visible = True
     CmdAdd4.Visible = True
     CmdMod4.Visible = True
     CmdElim4.Visible = True
     CmdApr4.Visible = True
     CmdAdd5.Visible = True
     CmdMod5.Visible = True
     CmdElim5.Visible = True
     CmdApr5.Visible = True
'     CmdAdd6.Visible = True
'     CmdMod6.Visible = True
'     CmdElim6.Visible = True
'     CmdApr6.Visible = True
   
     fraOpciones.Visible = False
         fra_cabecera.Enabled = False
     FraGrabarCancelar.Visible = True
     FraNavega.Enabled = False
'     FraSS_SS.Enabled = True
     TxtCodigo.Enabled = False
     rst_ben.Open "SELECT * FROM gc_Tipo_Beneficiario where estado_codigo ='APR' ORDER BY tipoben_descripcion ", db, adOpenStatic
     Set AdoTip_ben.Recordset = rst_ben
     
'   If Ado_datos.Recordset("tipoben_codigo") = "6" Then
'      SSTab1.Tab = 1
'      SSTab1.TabEnabled(1) = True
'      SSTab1.TabEnabled(0) = False
'      Frame12.Enabled = False
'      TxtNIT2.Enabled = False
'      txtCodigo2.Enabled = False
''      fraDatos2.Enabled = True
'      Frame22.Enabled = True
'      If Ado_datos.Recordset("estado_codigo") = "N" Then
'        txtDenominacion2.Enabled = True
'        DtcDepto32.Enabled = True
'        TxtCargo2.SetFocus
'        'TDBtipoben2.SetFocus
'      Else
'        txtDenominacion2.Enabled = False
'        DtcDepto32.Enabled = False
'        TxtCargo2.SetFocus
'      End If
'      rst_ben.Open "SELECT * FROM gc_Tipo_Beneficiario where estado_codigo ='B' ORDER BY descripcion ", db, adOpenStatic
'      Call rep_legal
'   End If
'   If (Ado_datos.Recordset("tipoben_codigo") = "1" Or Ado_datos.Recordset("tipoben_codigo") = "2" Or Ado_datos.Recordset("tipoben_codigo") = "7") Then
      
'  Else
'     MsgBox "El registro está APROBADO, solo se puede modificar por usuarios Autorizados ..."
'     '   Frame2.Enabled = False
'      '  Frame1.Enabled = False
'       ' TxtNIT.SetFocus
'  End If


Exit Sub

EditErr:
  MsgBox Err.Description
End Sub



Private Sub BtnSalir_Click()
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
   Unload Me
End Sub

Private Sub cmdDepurarBenef_Click()
'  If txtCodigo <> "" And txtDenominacion <> "" Then
'    FrmDepuradorBeneficiarios.Principal txtCodigo, txtDenominacion
'    FrmDepuradorBeneficiarios.Show vbModal
'  Else
'    MsgBox "El beneficiario no tiene Denominacion. Revise", vbInformation + vbOKOnly, "Atencion"
'  End If
End Sub

Private Sub Command10_Click()
On Error GoTo EditErr
     If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_Educacionales.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        ac_CapturaEstudiosRealizados.txtSW = "MOD"
        ac_CapturaEstudiosRealizados.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        ac_CapturaEstudiosRealizados.Txt01.Text = Ado_Educacionales.Recordset!Carrera_Curso
        ac_CapturaEstudiosRealizados.Txt02.Text = Ado_Educacionales.Recordset!centro_educativo
        ac_CapturaEstudiosRealizados.txt03.Text = Ado_Educacionales.Recordset!titulo_obtenido
        ac_CapturaEstudiosRealizados.Dtc_Par.Text = Ado_Educacionales.Recordset!nivel_educ_codigo
        ac_CapturaEstudiosRealizados.txt06.Text = Ado_Educacionales.Recordset!duracion_tiempo
        ac_CapturaEstudiosRealizados.txt07.Text = Ado_Educacionales.Recordset!tiempo_dmy
        ac_CapturaEstudiosRealizados.txt04.Text = Ado_Educacionales.Recordset!pais
        ac_CapturaEstudiosRealizados.txt05.Text = Ado_Educacionales.Recordset!ciudad
        ac_CapturaEstudiosRealizados.DTPFec_Inicio.Value = Ado_Educacionales.Recordset!fecha_inicio
        ac_CapturaEstudiosRealizados.DtcFec_Fin.Value = Ado_Educacionales.Recordset!fecha_fin
        ac_CapturaEstudiosRealizados.cboTDoc.Text = Ado_Educacionales.Recordset!presento_documento
        ac_CapturaEstudiosRealizados.txtEstado.Text = Ado_Educacionales.Recordset!estado_codigo
        ac_CapturaEstudiosRealizados.Dtc_ParDes.BoundText = ac_CapturaEstudiosRealizados.Dtc_Par.BoundText
        ac_CapturaEstudiosRealizados.Show vbModal
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Personal"
      End If
      Ado_Educacionales.Refresh
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
'    CmdAdd.Visible = False
'    CmdMod.Visible = False
'    CmdGraba.Visible = True

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command11_Click()
On Error GoTo EditErr
  sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If AdoLiquidacion.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoLiquidacion.Recordset("estado_codigo") = "ANL"
        AdoLiquidacion.Recordset("fecha_registro") = Date
        AdoLiquidacion.Recordset("usr_codigo") = glusuario
'        AdoLiquidacion.Recordset("centro_educativo") = "REG. ANULADO"
        AdoLiquidacion.Recordset.Update  'Batch adAffectAll
         Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
  
End Sub

Private Sub Command12_Click()
On Error GoTo EditErr
   
   sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If AdoLiquidacion.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoLiquidacion.Recordset("estado_codigo") = "APR"
        AdoLiquidacion.Recordset("fecha_registro") = Date
        AdoLiquidacion.Recordset("usr_usuario") = glusuario
        AdoLiquidacion.Recordset.Update
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command13_Click()
On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
      If AdoLiquidacion.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        ro_Personal_Liquidacion.txtSW = "MOD"
        ro_Personal_Liquidacion.TxtGestion = AdoLiquidacion.Recordset!ges_gestion
        ro_Personal_Liquidacion.TxtGestion_ini = IIf(IsNull(AdoLiquidacion.Recordset!ges_gestion_ini), Year(Date), AdoLiquidacion.Recordset!ges_gestion_ini)
        ro_Personal_Liquidacion.TxtGestion = AdoLiquidacion.Recordset!ges_gestion
        ro_Personal_Liquidacion.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        ro_Personal_Liquidacion.TxtInicial = Ado_datos.Recordset!beneficiario_iniciales
        ro_Personal_Liquidacion.TxtAprob = AdoLiquidacion.Recordset!estado_codigo
        ro_Personal_Liquidacion.TxtLquida.Text = AdoLiquidacion.Recordset!id_liquidacion
        ro_Personal_Liquidacion.DTPFInicio.Value = AdoLiquidacion.Recordset!fecha_ingreso
        ro_Personal_Liquidacion.DTPFFin.Value = AdoLiquidacion.Recordset!fecha_retiro
        ro_Personal_Liquidacion.DTCFInicio.Text = AdoLiquidacion.Recordset!fecha_ingreso
        ro_Personal_Liquidacion.DTCFFin.Text = AdoLiquidacion.Recordset!fecha_retiro
        ro_Personal_Liquidacion.DtcRetiro.Text = AdoLiquidacion.Recordset!tipo_memo
        ro_Personal_Liquidacion.CmbMes1.Text = IIf(IsNull(AdoLiquidacion.Recordset!Mes_Antepenultimo), "ENERO", AdoLiquidacion.Recordset!Mes_Antepenultimo)
        ro_Personal_Liquidacion.CmbMes2.Text = IIf(IsNull(AdoLiquidacion.Recordset!Mes_Penultimo), "FEBRERO", AdoLiquidacion.Recordset!Mes_Penultimo)
        ro_Personal_Liquidacion.CmbMes3.Text = IIf(IsNull(AdoLiquidacion.Recordset!Mes_Utimo), "MARZO", AdoLiquidacion.Recordset!Mes_Utimo)
        ro_Personal_Liquidacion.txtpago1.Text = IIf(IsNull(AdoLiquidacion.Recordset!Pago_Antepenultimo), "0", AdoLiquidacion.Recordset!Pago_Antepenultimo)
        ro_Personal_Liquidacion.TxtPago2.Text = IIf(IsNull(AdoLiquidacion.Recordset!Pago_Penultimo), "0", AdoLiquidacion.Recordset!Pago_Penultimo)
        ro_Personal_Liquidacion.Txtpago3.Text = IIf(IsNull(AdoLiquidacion.Recordset!Pago_Utimo), "0", AdoLiquidacion.Recordset!Pago_Utimo)
        ro_Personal_Liquidacion.txtpago4.Text = IIf(IsNull(AdoLiquidacion.Recordset!OtroPago_Antep), "0", AdoLiquidacion.Recordset!OtroPago_Antep)
        ro_Personal_Liquidacion.txtpago5.Text = IIf(IsNull(AdoLiquidacion.Recordset!OtroPago_Penul), "0", AdoLiquidacion.Recordset!OtroPago_Penul)
        ro_Personal_Liquidacion.txtpago6.Text = IIf(IsNull(AdoLiquidacion.Recordset!OtroPago_Utimo), "0", AdoLiquidacion.Recordset!OtroPago_Utimo)
        ro_Personal_Liquidacion.lblARCH.Caption = AdoLiquidacion.Recordset!ARCHIVO
        ro_Personal_Liquidacion.CmbAño.Text = IIf(IsNull(AdoLiquidacion.Recordset!Años), "0", AdoLiquidacion.Recordset!Años)
        ro_Personal_Liquidacion.CmbMes.Text = IIf(IsNull(AdoLiquidacion.Recordset!meses), "0", AdoLiquidacion.Recordset!meses)
        ro_Personal_Liquidacion.CmbDia.Text = IIf(IsNull(AdoLiquidacion.Recordset!DIAS), "0", AdoLiquidacion.Recordset!DIAS)
        ro_Personal_Liquidacion.TxtImdemAño.Text = IIf(IsNull(AdoLiquidacion.Recordset!Imdem_Año), "0", AdoLiquidacion.Recordset!Imdem_Año)
        ro_Personal_Liquidacion.TxtImdemMes.Text = IIf(IsNull(AdoLiquidacion.Recordset!Imdem_Mes), "0", AdoLiquidacion.Recordset!Imdem_Mes)
        ro_Personal_Liquidacion.TxtImdemDia.Text = IIf(IsNull(AdoLiquidacion.Recordset!Indem_dias), "0", AdoLiquidacion.Recordset!Indem_dias)
        ro_Personal_Liquidacion.TxtNavidad.Text = IIf(IsNull(AdoLiquidacion.Recordset!Aguin_Navidad), "0", AdoLiquidacion.Recordset!Aguin_Navidad)
        ro_Personal_Liquidacion.TxtVacacion.Text = IIf(IsNull(AdoLiquidacion.Recordset!Aguin_Vacacion), "0", AdoLiquidacion.Recordset!Aguin_Vacacion)
        ro_Personal_Liquidacion.TxtPrima.Text = IIf(IsNull(AdoLiquidacion.Recordset!Prima_Legal), "0", AdoLiquidacion.Recordset!Prima_Legal)
        ro_Personal_Liquidacion.TxtOtros.Text = IIf(IsNull(AdoLiquidacion.Recordset!Otros_Pagos), "0", AdoLiquidacion.Recordset!Otros_Pagos)
        ro_Personal_Liquidacion.CmbChq_Trf.Text = IIf(IsNull(AdoLiquidacion.Recordset!Forma_pago), "CHEQUE", AdoLiquidacion.Recordset!Forma_pago)
        ro_Personal_Liquidacion.TxtNo_Chq.Text = IIf(IsNull(AdoLiquidacion.Recordset!Num_chq_cmpbte), "0", AdoLiquidacion.Recordset!Num_chq_cmpbte)
        ro_Personal_Liquidacion.TxtCta.Text = IIf(IsNull(AdoLiquidacion.Recordset!cta_codigo), "0", AdoLiquidacion.Recordset!cta_codigo)
        ro_Personal_Liquidacion.TxtDeduccion.Text = IIf(IsNull(AdoLiquidacion.Recordset!Deducciones), "0", AdoLiquidacion.Recordset!Deducciones)
        ro_Personal_Liquidacion.TxtTotBenef.Text = IIf(IsNull(AdoLiquidacion.Recordset!monto_total), "0", AdoLiquidacion.Recordset!monto_total)
'        ro_Personal_Liquidacion.txt_dias_agui.Text = AdoLiquidacion.Recordset!dias_agui
'        ro_Personal_Liquidacion.txt_meses_agui.Text = AdoLiquidacion.Recordset!meses_agui
        If ro_Personal_Liquidacion.DtcRetiro.Text = "QUI" Then
        ro_Personal_Liquidacion.Frame4.Visible = False
        Else
        ro_Personal_Liquidacion.Frame4.Visible = True
        End If
        ro_Personal_Liquidacion.Show vbModal
        'Ado_Contrato.Refresh
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Personal"
      End If
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
   Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command14_Click()
On Error GoTo EditErr
 If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        'AdoLiquidacion.Recordset.AddNew
        ro_Personal_Liquidacion.txtSW = "ADD"
        ro_Personal_Liquidacion.TxtGestion = Year(Date)
        ro_Personal_Liquidacion.TxtGestion_ini = Year(Date) - 5
        ro_Personal_Liquidacion.txtBenef.Text = Ado_datos.Recordset!beneficiario_codigo
        ro_Personal_Liquidacion.TxtInicial = Ado_datos.Recordset!beneficiario_iniciales
        ro_Personal_Liquidacion.TxtAprob = "REG"
        ro_Personal_Liquidacion.txtpago1 = Ado_datos.Recordset!beneficiario_haber_mensual
        ro_Personal_Liquidacion.TxtPago2 = Ado_datos.Recordset!beneficiario_haber_mensual
        ro_Personal_Liquidacion.Txtpago3 = Ado_datos.Recordset!beneficiario_haber_mensual
        'frmBeneficiario_Admin.AdoLiquidacion.Recordset!tipo_memo = "REF"
        ro_Personal_Liquidacion.Show vbModal
        'Call abrirtabla
        'AdoLiquidacion.Refresh
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

Exit Sub

EditErr:
  MsgBox Err.Description

End Sub

Private Sub Command15_Click()
On Error GoTo EditErr
 sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_Contrato.Recordset("estado_contrato") = "REG" Then
      If sino = vbYes Then
        Ado_Contrato.Recordset("estado_contrato") = "ANL"
        Ado_Contrato.Recordset("fecha_registro") = Date
        Ado_Contrato.Recordset("usr_usuario") = glusuario
        Ado_Contrato.Recordset("observacion_contrato") = "REGISTRO ANULADO"
        Ado_Contrato.Recordset.Update  'Batch adAffectAll
         Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command16_Click()
 On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_Contrato.Recordset("estado_contrato") = "REG" Then
'      If Ado_Contrato.Recordset!ARCHIVO <> "Cargar_Archivo" Then
        If sino = vbYes Then
            Ado_Contrato.Recordset("estado_contrato") = "APR"
            Ado_Contrato.Recordset("fecha_aprueba") = Date
            Ado_Contrato.Recordset("usr_aprueba") = glusuario
            Ado_Contrato.Recordset("observacion_contrato") = "REGISTRO APROBADO"
            Ado_Contrato.Recordset.Update
        End If
'      Else
'            MsgBox "No se puede APROBAR. Previamente Debe cargar el archivo .PDF asociado al registro ... ", vbExclamation, "Validación de Registro"
'      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Command17_Click()
 On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_Contrato.Recordset!estado_contrato = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ro_personal_contrato.txtSW = "MOD"
        frm_ro_personal_contrato.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ro_personal_contrato.TxtInicial = Ado_datos.Recordset!beneficiario_iniciales
        frm_ro_personal_contrato.TxtForm = Ado_Contrato.Recordset!solicitud_codigo
        frm_ro_personal_contrato.TxtAprob = Ado_Contrato.Recordset!estado_contrato
        frm_ro_personal_contrato.lblARCH.Caption = Ado_Contrato.Recordset!ARCHIVO
        frm_ro_personal_contrato.TxtCodigo.Text = Ado_Contrato.Recordset!codigo_contrato
        frm_ro_personal_contrato.txtObjContrato.Text = IIf(IsNull(Ado_Contrato.Recordset!objeto_contrato), "-", Ado_Contrato.Recordset!objeto_contrato)
        frm_ro_personal_contrato.DTcFte.Text = IIf(IsNull(Ado_Contrato.Recordset!fte_codigo), "10", Ado_Contrato.Recordset!fte_codigo)
        frm_ro_personal_contrato.dtc_codigo.Text = Ado_Contrato.Recordset!unidad_codigo
        frm_ro_personal_contrato.DtcOrg.Text = IIf(IsNull(Ado_Contrato.Recordset!org_codigo), "111", Ado_Contrato.Recordset!org_codigo)
        frm_ro_personal_contrato.DtcCargo.Text = Ado_Contrato.Recordset!cargo_codigo
        frm_ro_personal_contrato.DtcPry.Text = Ado_Contrato.Recordset!pro_codigo
        frm_ro_personal_contrato.DTPFInicio.Value = Ado_Contrato.Recordset!fecha_inicio
        frm_ro_personal_contrato.DTPFFin.Value = Ado_Contrato.Recordset!fecha_fin
        frm_ro_personal_contrato.DtcPuesto.Text = Ado_Contrato.Recordset!puesto_codigo
        frm_ro_personal_contrato.txtEstado.Text = Ado_Contrato.Recordset!estado_confirmado
        frm_ro_personal_contrato.DTPFFirma.Value = Ado_Contrato.Recordset!fecha_firma
        
        frm_ro_personal_contrato.TxtBs.Text = Ado_Contrato.Recordset!monto_totalbs
        frm_ro_personal_contrato.txt_time.Text = Ado_Contrato.Recordset!tiempo_num
        frm_ro_personal_contrato.txtMensual_bs.Text = IIf(IsNull(Ado_Contrato.Recordset!monto_mensualBS), 0, Ado_Contrato.Recordset!monto_mensualBS)
        frm_ro_personal_contrato.txt_otro_bs.Text = IIf(IsNull(Ado_Contrato.Recordset!monto_otroBS), 0, Ado_Contrato.Recordset!monto_otroBS)
        frm_ro_personal_contrato.DtcRespaldoCod.Text = Ado_Contrato.Recordset!doc_codigo
        
        frm_ro_personal_contrato.DtcFteDes.BoundText = frm_ro_personal_contrato.DTcFte.BoundText
        frm_ro_personal_contrato.DtcOrgDes.BoundText = frm_ro_personal_contrato.DtcOrg.BoundText
        frm_ro_personal_contrato.DtcPryDes.BoundText = frm_ro_personal_contrato.DtcPry.BoundText
        frm_ro_personal_contrato.Dtc_descrip.BoundText = frm_ro_personal_contrato.dtc_codigo.BoundText
        frm_ro_personal_contrato.DtcCargoDes.BoundText = frm_ro_personal_contrato.DtcCargo.BoundText
        frm_ro_personal_contrato.DtcPuestoDes.BoundText = frm_ro_personal_contrato.DtcPuesto.BoundText
        frm_ro_personal_contrato.DtcRespaldo.BoundText = frm_ro_personal_contrato.DtcRespaldoCod.BoundText
        
        frm_ro_personal_contrato.Show vbModal
        'Ado_Contrato.Refresh
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Personal"
      End If
      Call abrirtabla
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
   Exit Sub
EditErr:
  MsgBox Err.Description
End Sub


Private Sub Command18_Click()
On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ro_personal_contrato.txtSW = "ADD"
        frm_ro_personal_contrato.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ro_personal_contrato.TxtInicial = Ado_datos.Recordset!beneficiario_iniciales
        frm_ro_personal_contrato.TxtAprob = "REG"
        'Ado_Contrato.Recordset.AddNew
        frm_ro_personal_contrato.Show vbModal
        'Call abrirtabla
        'Ado_Contrato.Refresh
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
   Exit Sub
AddErr:
  MsgBox Err.Description

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command19_Click()
    rw_contratacion_personal.Show
End Sub

Private Sub Command2_Click()
rw_reportes_asistencia.Show
End Sub

Private Sub Command1_Click()
On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
       marca1 = Ado_datos.Recordset.Bookmark
       frm_ao_Vacacion_Prog.txtSW = "ADD"
       frm_ao_Vacacion_Prog.txtBenef = Ado_datos.Recordset!beneficiario_codigo
       frm_ao_Vacacion_Prog.txtEstado = "REG"
       frm_ao_Vacacion_Prog.TxtGestion.Text = Year(Date)
'       frm_ao_Vacacion_Prog.lblbien(1).Visible = False
'       frm_ao_Vacacion_Prog.Txt02.Visible = False
       'Ado_VacacionesProg.Recordset.AddNew
       frm_ao_Vacacion_Prog.sel = 2
       frm_ao_Vacacion_Prog.Show vbModal
       Call abrirtabla
       'Ado_VacacionesProg.Refresh
   Else
       MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnModificar2_Click()
    glBenef = Ado_datos.Recordset!beneficiario_codigo
    gw_beneficiario_cuenta.Show 'vbModal
End Sub

Private Sub Command3_Click()
On Error GoTo EditErr
sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_Laborales.Recordset("estado_codigo") = "REG" Then
        If sino = vbYes Then
          Ado_Laborales.Recordset("estado_codigo") = "APR"
          Ado_Laborales.Recordset("fecha_aprueba") = Date
          Ado_Laborales.Recordset("usr_aprueba") = glusuario
          Ado_Laborales.Recordset.Update
        End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   
Exit Sub

EditErr:
  MsgBox Err.Description
  
End Sub

Private Sub Command4_Click()
 On Error GoTo EditErr
  sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_Laborales.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_Laborales.Recordset("estado_codigo") = "ANL"
        Ado_Laborales.Recordset("fecha_registro") = Date
        Ado_Laborales.Recordset("usr_usuario") = glusuario
'        Ado_Laborales.Recordset("cargo") = "REG. ANULADO"
        Ado_Laborales.Recordset.Update  'Batch adAffectAll
         Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
   
   Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command5_Click()
 On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        ac_CapturaExperienciaLaboral.txtSW = "ADD"
        ac_CapturaExperienciaLaboral.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        ac_CapturaExperienciaLaboral.txtEstado = "REG"
        Ado_Laborales.Recordset.AddNew
        ac_CapturaExperienciaLaboral.Show vbModal
        'Call abrirtabla
        'Ado_Laborales.Refresh
   Else
        MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command6_Click()
 On Error GoTo EditErr
     If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_Laborales.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        ac_CapturaExperienciaLaboral.txtSW = "MOD"
        ac_CapturaExperienciaLaboral.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        ac_CapturaExperienciaLaboral.Txt01.Text = Ado_Laborales.Recordset!denominacion_institucion
        ac_CapturaExperienciaLaboral.Txt02.Text = Ado_Laborales.Recordset!cargo
        ac_CapturaExperienciaLaboral.txt03.Text = Ado_Laborales.Recordset!funcion_general
        ac_CapturaExperienciaLaboral.Dtc_Par.Text = Ado_Laborales.Recordset!tipo_institucion
        ac_CapturaExperienciaLaboral.txt06.Text = Ado_Laborales.Recordset!Tiempo_Meses
        ac_CapturaExperienciaLaboral.txt04.Text = Ado_Laborales.Recordset!pais
        ac_CapturaExperienciaLaboral.txt05.Text = Ado_Laborales.Recordset!ciudad
        ac_CapturaExperienciaLaboral.DTPFec_Inicio.Value = Ado_Laborales.Recordset!fecha_inicio
        ac_CapturaExperienciaLaboral.DtcFec_Fin.Value = Ado_Laborales.Recordset!fecha_fin
        ac_CapturaExperienciaLaboral.cboTDoc.Text = Ado_Laborales.Recordset!presento_documento
        ac_CapturaExperienciaLaboral.txtEstado.Text = Ado_Laborales.Recordset!estado_codigo
        ac_CapturaExperienciaLaboral.Dtc_ParDes.BoundText = ac_CapturaExperienciaLaboral.Dtc_Par.BoundText
        ac_CapturaExperienciaLaboral.Show vbModal
        Ado_Laborales.Refresh
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Personal"
      End If
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If


Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command7_Click()
 On Error GoTo EditErr
 sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_Educacionales.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_Educacionales.Recordset("estado_codigo") = "APR"
        Ado_Educacionales.Recordset("fecha_registro") = Date
        Ado_Educacionales.Recordset("usr_usuario") = glusuario
        Ado_Educacionales.Recordset.Update
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command8_Click()
 On Error GoTo EditErr
 sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_Educacionales.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_Educacionales.Recordset("estado_codigo") = "ANL"
        Ado_Educacionales.Recordset("fecha_registro") = Date
        Ado_Educacionales.Recordset("usr_usuario") = glusuario
        Ado_Educacionales.Recordset("centro_educativo") = "REG. ANULADO"
        Ado_Educacionales.Recordset.Update  'Batch adAffectAll
         Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
   
Exit Sub

EditErr:
  MsgBox Err.Description
   
End Sub

Private Sub Command9_Click()

 On Error GoTo EditErr
 If Ado_datos.Recordset.RecordCount > 0 Then
       marca1 = Ado_datos.Recordset.Bookmark
       ac_CapturaEstudiosRealizados.txtSW = "ADD"
       ac_CapturaEstudiosRealizados.txtBenef = Ado_datos.Recordset!beneficiario_codigo
       ac_CapturaEstudiosRealizados.txtEstado = "REG"
       'Ado_Educacionales.Recordset.AddNew
       ac_CapturaEstudiosRealizados.Show vbModal
       'Call abrirtabla
       'Ado_Educacionales.Refresh
   Else
       MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
 Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub dtc_buscar_ci_Click(Area As Integer)
dtc_buscar_desc.BoundText = dtc_buscar_ci.BoundText
End Sub

Private Sub dtc_buscar_desc_Change()
 dtc_buscar_ci.BoundText = dtc_buscar_desc.BoundText
 If dtc_buscar_ci.SelectedItem <> "" Then
 'busq = busq + 1
 'If busq = 2 Then
 Call Carga_Beneficiario(3)
 'busq = 0
 'End If
 End If
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText

End Sub



Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
 dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    dtc_desc7.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
 dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub DtcPaisSigla_Click(Area As Integer)
    TxtNacionalidad.BoundText = DtcPaisSigla.BoundText
    DtcPaisCod.BoundText = DtcPaisSigla.BoundText
End Sub

Private Sub DtcEstCivDes_Click(Area As Integer)
 DtcEstCiv.BoundText = DtcEstCivDes.BoundText
End Sub

Private Sub DtcEstCiv_Click(Area As Integer)
    DtcEstCivDes.BoundText = DtcEstCiv.BoundText
End Sub

Private Sub Dtc_depto_Click(Area As Integer)
    Dtc_depto_cod.BoundText = dtc_depto.BoundText
    Call pProvincia(Dtc_depto_cod.BoundText)
End Sub

Private Sub Dtc_depto_cod_Click(Area As Integer)
    dtc_depto.BoundText = Dtc_depto_cod.BoundText
    Call pProvincia(dtc_depto.BoundText)
End Sub

Private Sub pProvincia(depto_codigo As String)
   Dim strConsultaP As String

   strConsultaP = "select * from GC_Provincia where depto_codigo='" & depto_codigo & "'"

   Set Dtc_prov_cod.RowSource = Nothing
   Set Dtc_prov_cod.RowSource = db.Execute(strConsultaP, , adCmdText)
   Dtc_prov_cod.ReFill
   Dtc_prov_cod.BoundText = Empty

   Set Dtc_prov.RowSource = Nothing
   Set Dtc_prov.RowSource = db.Execute(strConsultaP, , adCmdText)
   Dtc_prov.ReFill
   Dtc_prov.BoundText = Empty
End Sub

'Private Sub Dtc_depto_cod02_Click(Area As Integer)
'    Dtc_depto02.BoundText = Dtc_depto_cod02.BoundText
'    Call pProvincia02(Dtc_depto02.BoundText)
'End Sub

'Private Sub Dtc_depto_cod22_Click(Area As Integer)
'    Dtc_depto22.BoundText = Dtc_depto_cod22.BoundText
'    Call pProvincia22(Dtc_depto22.BoundText)
'End Sub

'Private Sub Dtc_depto02_Click(Area As Integer)
'    Dtc_depto_cod02.BoundText = Dtc_depto02.BoundText
'    Call pProvincia02(Dtc_depto_cod02.BoundText)
'End Sub

'Private Sub pProvincia02(depto_codigo02 As String)
'   Dim strConsultaP02 As String
'
'   strConsultaP02 = "select * from GC_Provincia where depto_codigo='" & depto_codigo02 & "'"
'
'   Set Dtc_prov_cod02.RowSource = Nothing
'   Set Dtc_prov_cod02.RowSource = DB.Execute(strConsultaP02, , adCmdText)
'   Dtc_prov_cod02.ReFill
'   Dtc_prov_cod02.BoundText = Empty
'
'   Set Dtc_prov02.RowSource = Nothing
'   Set Dtc_prov02.RowSource = DB.Execute(strConsultaP02, , adCmdText)
'   Dtc_prov02.ReFill
'   Dtc_prov02.BoundText = Empty
'End Sub

'Private Sub Dtc_depto2_Click(Area As Integer)
'    Dtc_depto_cod2.BoundText = Dtc_depto2.BoundText
'    Call pProvincia2(Dtc_depto_cod2.BoundText)
'End Sub

'Private Sub Dtc_depto_cod2_Click(Area As Integer)
'    Dtc_depto2.BoundText = Dtc_depto_cod2.BoundText
'    Call pProvincia2(Dtc_depto2.BoundText)
'End Sub
'
'Private Sub pProvincia2(depto_codigo2 As String)
'   Dim strConsultaP2 As String
'
'   strConsultaP2 = "select * from GC_Provincia where depto_codigo='" & depto_codigo2 & "'"
'
'   Set Dtc_prov_cod2.RowSource = Nothing
'   Set Dtc_prov_cod2.RowSource = DB.Execute(strConsultaP2, , adCmdText)
'   Dtc_prov_cod2.ReFill
'   Dtc_prov_cod2.BoundText = Empty
'
'   Set Dtc_prov2.RowSource = Nothing
'   Set Dtc_prov2.RowSource = DB.Execute(strConsultaP2, , adCmdText)
'   Dtc_prov2.ReFill
'   Dtc_prov2.BoundText = Empty
'End Sub

'Private Sub Dtc_depto22_Click(Area As Integer)
'    Dtc_depto_cod22.BoundText = Dtc_depto22.BoundText
'    Call pProvincia22(Dtc_depto_cod22.BoundText)
'End Sub

'Private Sub pProvincia22(depto_codigo22 As String)
'   Dim strConsultaP22 As String
'
'   strConsultaP22 = "select * from GC_Provincia where depto_codigo='" & depto_codigo22 & "'"
'
'   Set Dtc_prov_cod22.RowSource = Nothing
'   Set Dtc_prov_cod22.RowSource = DB.Execute(strConsultaP22, , adCmdText)
'   Dtc_prov_cod22.ReFill
'   Dtc_prov_cod22.BoundText = Empty
'
'   Set Dtc_prov22.RowSource = Nothing
'   Set Dtc_prov22.RowSource = DB.Execute(strConsultaP22, , adCmdText)
'   Dtc_prov22.ReFill
'   Dtc_prov22.BoundText = Empty
'End Sub

'Private Sub Dtc_local_cod02_Click(Area As Integer)
'    Dtc_local02.BoundText = Dtc_local_cod02.BoundText
'End Sub

'Private Sub Dtc_local_cod22_Click(Area As Integer)
'    Dtc_local22.BoundText = Dtc_local_cod22.BoundText
'End Sub

'Private Sub Dtc_local02_Click(Area As Integer)
'    Dtc_local_cod02.BoundText = Dtc_local02.BoundText
'End Sub

'Private Sub Dtc_local22_Click(Area As Integer)
'    Dtc_local_cod22.BoundText = Dtc_local22.BoundText
'End Sub

'Private Sub Dtc_munic_cod02_Click(Area As Integer)
'    Dtc_munic02.BoundText = Dtc_munic_cod02.BoundText
'    Call pComunidad02(Dtc_munic_cod02.BoundText)
'End Sub


'Private Sub Dtc_munic_cod2_Click(Area As Integer)
'    Dtc_munic2.BoundText = Dtc_munic_cod2.BoundText
'    Call pComunidad2(Dtc_munic_cod2.BoundText)
'End Sub

'Private Sub Dtc_munic_cod22_Click(Area As Integer)
'    Dtc_munic22.BoundText = Dtc_munic_cod22.BoundText
'    Call pComunidad22(Dtc_munic_cod22.BoundText)
'End Sub

'Private Sub Dtc_munic02_Click(Area As Integer)
'    Dtc_munic_cod02.BoundText = Dtc_munic02.BoundText
'    Call pComunidad02(Dtc_munic_cod02.BoundText)
'End Sub

'Private Sub pComunidad02(CodMunic02 As String)
'   Dim strConsultaC02 As String
'
'   strConsultaC02 = "select * from GC_comunidad where munic_codigo='" & CodMunic02 & "'"
'
'   Set Dtc_local_cod02.RowSource = Nothing
'   Set Dtc_local_cod02.RowSource = DB.Execute(strConsultaC02, , adCmdText)
'   Dtc_local_cod02.ReFill
'   Dtc_local_cod02.BoundText = Empty
'
'   Set Dtc_local02.RowSource = Nothing
'   Set Dtc_local02.RowSource = DB.Execute(strConsultaC02, , adCmdText)
'   Dtc_local02.ReFill
'   Dtc_local02.BoundText = Empty
'End Sub

'Private Sub Dtc_munic22_Click(Area As Integer)
'    Dtc_munic_cod22.BoundText = Dtc_munic22.BoundText
'    Call pComunidad22(Dtc_munic_cod22.BoundText)
'End Sub

'Private Sub pComunidad22(CodMunic22 As String)
'   Dim strConsultaC22 As String
'
'   strConsultaC22 = "select * from GC_comunidad where munic_codigo='" & CodMunic22 & "'"
'
'   Set Dtc_local_cod22.RowSource = Nothing
'   Set Dtc_local_cod22.RowSource = DB.Execute(strConsultaC22, , adCmdText)
'   Dtc_local_cod22.ReFill
'   Dtc_local_cod22.BoundText = Empty
'
'   Set Dtc_local22.RowSource = Nothing
'   Set Dtc_local22.RowSource = DB.Execute(strConsultaC22, , adCmdText)
'   Dtc_local22.ReFill
'   Dtc_local22.BoundText = Empty
'End Sub

Private Sub Dtc_Ocup_Click(Area As Integer)
    TxtProfesion.BoundText = Dtc_Ocup.BoundText
End Sub

Private Sub Dtc_prov_Click(Area As Integer)
    Dtc_prov_cod.BoundText = Dtc_prov.BoundText
    Call pMunicipio(Dtc_prov_cod.BoundText)
End Sub

Private Sub Dtc_prov_cod_Click(Area As Integer)
    Dtc_prov.BoundText = Dtc_prov_cod.BoundText
    Call pMunicipio(Dtc_prov.BoundText)
End Sub

Private Sub pMunicipio(CodProv As String)
   Dim strConsultaM As String

   strConsultaM = "select * from gc_Municipio where prov_codigo='" & CodProv & "'"

   Set Dtc_munic_cod.RowSource = Nothing
   Set Dtc_munic_cod.RowSource = db.Execute(strConsultaM, , adCmdText)
   Dtc_munic_cod.ReFill
   Dtc_munic_cod.BoundText = Empty

   Set Dtc_munic.RowSource = Nothing
   Set Dtc_munic.RowSource = db.Execute(strConsultaM, , adCmdText)
   Dtc_munic.ReFill
   Dtc_munic.BoundText = Empty
End Sub

'Private Sub Dtc_prov_cod02_Click(Area As Integer)
'    Dtc_prov02.BoundText = Dtc_prov_cod02.BoundText
'    Call pMunicipio02(Dtc_prov02.BoundText)
'End Sub


'Private Sub Dtc_prov_cod22_Click(Area As Integer)
'    Dtc_prov22.BoundText = Dtc_prov_cod22.BoundText
'    Call pMunicipio22(Dtc_prov22.BoundText)
'End Sub

'Private Sub Dtc_prov02_Click(Area As Integer)
'    Dtc_prov_cod02.BoundText = Dtc_prov02.BoundText
'    Call pMunicipio02(Dtc_prov_cod02.BoundText)
'End Sub

'Private Sub pMunicipio02(CodProv02 As String)
'   Dim strConsultaM02 As String
'
'   strConsultaM02 = "select * from gc_Municipio where prov_codigo='" & CodProv02 & "'"
'
'   Set Dtc_munic_cod02.RowSource = Nothing
'   Set Dtc_munic_cod02.RowSource = DB.Execute(strConsultaM02, , adCmdText)
'   Dtc_munic_cod02.ReFill
'   Dtc_munic_cod02.BoundText = Empty
'
'   Set Dtc_munic02.RowSource = Nothing
'   Set Dtc_munic02.RowSource = DB.Execute(strConsultaM02, , adCmdText)
'   Dtc_munic02.ReFill
'   Dtc_munic02.BoundText = Empty
'End Sub

'Private Sub Dtc_prov2_Click(Area As Integer)
'    Dtc_prov_cod2.BoundText = Dtc_prov2.BoundText
'    Call pMunicipio2(Dtc_prov_cod2.BoundText)
'End Sub

'Private Sub Dtc_prov_cod2_Click(Area As Integer)
'    Dtc_prov2.BoundText = Dtc_prov_cod2.BoundText
'    Call pMunicipio2(Dtc_prov2.BoundText)
'End Sub

'Private Sub pMunicipio2(CodProv2 As String)
'   Dim strConsultaM2 As String
'
'   strConsultaM2 = "select * from gc_Municipio where prov_codigo='" & CodProv2 & "'"
'
'   Set Dtc_munic_cod2.RowSource = Nothing
'   Set Dtc_munic_cod2.RowSource = DB.Execute(strConsultaM2, , adCmdText)
'   Dtc_munic_cod2.ReFill
'   Dtc_munic_cod2.BoundText = Empty
'
'   Set Dtc_munic2.RowSource = Nothing
'   Set Dtc_munic2.RowSource = DB.Execute(strConsultaM2, , adCmdText)
'   Dtc_munic2.ReFill
'   Dtc_munic2.BoundText = Empty
'End Sub

Private Sub Dtc_munic_Click(Area As Integer)
    Dtc_munic_cod.BoundText = Dtc_munic.BoundText
'    Call pComunidad(Dtc_munic_cod.BoundText)
End Sub

Private Sub Dtc_munic_cod_Click(Area As Integer)
    Dtc_munic.BoundText = Dtc_munic_cod.BoundText
    'Call pComunidad(Dtc_munic.BoundText)
End Sub

'Private Sub pComunidad(CodMunic As String)
'   Dim strConsultaC As String
'
'   strConsultaC = "select * from GC_comunidad where munic_codigo='" & CodMunic & "'"
'
'   Set Dtc_local_cod.RowSource = Nothing
'   Set Dtc_local_cod.RowSource = DB.Execute(strConsultaC, , adCmdText)
'   Dtc_local_cod.ReFill
'   Dtc_local_cod.BoundText = Empty
'
'   Set Dtc_local.RowSource = Nothing
'   Set Dtc_local.RowSource = DB.Execute(strConsultaC, , adCmdText)
'   Dtc_local.ReFill
'   Dtc_local.BoundText = Empty
'End Sub

'Private Sub Dtc_munic2_Click(Area As Integer)
'    Dtc_munic_cod2.BoundText = Dtc_munic2.BoundText
'    Call pComunidad2(Dtc_munic_cod2.BoundText)
'End Sub

'Private Sub Dtc_munic2_cod_Click(Area As Integer)
'    Dtc_munic2.BoundText = Dtc_munic_cod2.BoundText
'    Call pComunidad2(Dtc_munic2.BoundText)
'End Sub

'Private Sub pComunidad2(CodMunic2 As String)
'   Dim strConsultaC2 As String
'
'   strConsultaC2 = "select * from GC_comunidad where munic_codigo='" & CodMunic2 & "'"
'
'   Set Dtc_local_cod2.RowSource = Nothing
'   Set Dtc_local_cod2.RowSource = DB.Execute(strConsultaC2, , adCmdText)
'   Dtc_local_cod2.ReFill
'   Dtc_local_cod2.BoundText = Empty
'
'   Set Dtc_local2.RowSource = Nothing
'   Set Dtc_local2.RowSource = DB.Execute(strConsultaC2, , adCmdText)
'   Dtc_local2.ReFill
'   Dtc_local2.BoundText = Empty
'End Sub

'Private Sub Dtc_local_Click(Area As Integer)
'    Dtc_local_cod.BoundText = Dtc_local.BoundText
'End Sub
'
'Private Sub Dtc_local_cod_Click(Area As Integer)
'    Dtc_local.BoundText = Dtc_local_cod.BoundText
'End Sub

'Private Sub Dtc_local2_Click(Area As Integer)
'    Dtc_local_cod2.BoundText = Dtc_local2.BoundText
'End Sub

'Private Sub Dtc_local_cod2_Click(Area As Integer)
'    Dtc_local2.BoundText = Dtc_local_cod2.BoundText
'End Sub

'Private Sub Dtc_prov22_Click(Area As Integer)
'    Dtc_prov_cod22.BoundText = Dtc_prov22.BoundText
'    Call pMunicipio22(Dtc_prov_cod22.BoundText)
'End Sub

'Private Sub pMunicipio22(CodProv22 As String)
'   Dim strConsultaM22 As String
'
'   strConsultaM22 = "select * from gc_Municipio where prov_codigo='" & CodProv22 & "'"
'
'   Set Dtc_munic_cod22.RowSource = Nothing
'   Set Dtc_munic_cod22.RowSource = DB.Execute(strConsultaM22, , adCmdText)
'   Dtc_munic_cod22.ReFill
'   Dtc_munic_cod22.BoundText = Empty
'
'   Set Dtc_munic22.RowSource = Nothing
'   Set Dtc_munic22.RowSource = DB.Execute(strConsultaM22, , adCmdText)
'   Dtc_munic22.ReFill
'   Dtc_munic22.BoundText = Empty
'End Sub

Private Sub DtcPaisCod_Click(Area As Integer)
    DtcPaisSigla.BoundText = DtcPaisCod.BoundText
    TxtNacionalidad.BoundText = DtcPaisCod.BoundText
End Sub

'Private Sub DtcRep_Materno_Click(Area As Integer)
'    TxtNIT2.BoundText = DtcRep_Materno.BoundText
'    DtcRep_Paterno.BoundText = DtcRep_Materno.BoundText
'    DtcRep_Nombres.BoundText = DtcRep_Materno.BoundText
'End Sub

'Private Sub DtcRep_Nombres_Click(Area As Integer)
'    TxtNIT2.BoundText = DtcRep_Nombres.BoundText
'    DtcRep_Paterno.BoundText = DtcRep_Nombres.BoundText
'    DtcRep_Materno.BoundText = DtcRep_Nombres.BoundText
'End Sub

'Private Sub DtcRep_Paterno_Click(Area As Integer)
'    TxtNIT2.BoundText = DtcRep_Paterno.BoundText
'    DtcRep_Materno.BoundText = DtcRep_Paterno.BoundText
'    DtcRep_Nombres.BoundText = DtcRep_Paterno.BoundText
'End Sub

Private Sub DtcRep_Paterno_LostFocus()
'    Text102.Text = DtcRep_Paterno.Text
'    Text202.Text = DtcRep_Materno.Text
'    Text302.Text = DtcRep_Nombres.Text
End Sub

Private Sub Form_Load()
On Error GoTo EditErr
'Ado_VacacionesProg   'Label5.Caption = GlUsuario       'frmLogin.txtUserName.Text      'JQA NOV-2009
   fraDatos.Enabled = False
'   fraDatos2.Enabled = False
'   FraSS_SS.Enabled = False
     Call OptFilGral1_Click
    Call Carga_afp
   Call Carga_Recor
   'Call Carga_Beneficiario(1)
   
'   Call rep_legal
   cbo_gestion.Text = Year(Date)
   cbo_mes.Text = UCase(MonthName(Month(Date)))
   txt_mes.Text = Month(Date)
   Call abrirtabla
  
    Call Carga_Beneficiario(2)
   GlSW = ""
   swnuevo = 0
'   Fra_ABM.Enabled = False
   If Not Ado_datos.Recordset.EOF Then
'        If Ado_datos.Recordset("tipoben_codigo") = "6" Then
'            SSTab1.Tab = 3
'            SSTab1.TabEnabled(0) = False
'            SSTab1.TabEnabled(1) = False
'            SSTab1.TabEnabled(2) = False
''            SSTab1.TabEnabled(3) = True
'        Else
            SSTab1.Tab = 0
'            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = True
'        End Ifa

'        DtgVacacionesProg.AllowAddNew = False
'        DtgVacacionesProg.AllowDelete = False
'        DtgVacacionesProg.AllowUpdate = False
'        DtgVacaciones.AllowAddNew = False
'        DtgVacaciones.AllowDelete = False
'        DtgVacaciones.AllowUpdate = False


'aqui se hacen invisibles os botones del los grids

'        CmdAdd1.Visible = False
'        CmdMod1.Visible = False
'        CmdElim1.Visible = False
'        CmdApr1.Visible = False
'        CmdAdd2.Visible = False
'        CmdMod2.Visible = False
'        CmdElim2.Visible = False
'        CmdApr2.Visible = False
'        CmdAdd3.Visible = False
'        CmdMod3.Visible = False
'        CmdElim3.Visible = False
'        CmdApr3.Visible = False
'        CmdAdd4.Visible = False
'        CmdMod4.Visible = False
'        CmdElim4.Visible = False
'        CmdApr4.Visible = False
'        CmdAdd5.Visible = False
'        CmdMod5.Visible = False
'        CmdElim5.Visible = False
'        CmdApr5.Visible = False

'----------------------------------------------------


'        CmdAdd6.Visible = False
'        CmdMod6.Visible = False
'        CmdElim6.Visible = False
'        CmdApr6.Visible = False
   End If
   ' Set ClBuscaGrid = Nothing
 
Exit Sub

EditErr:
  MsgBox Err.Description
	Call SeguridadSet(Me)
End Sub

Private Sub Carga_Beneficiario(posicion As Integer)
 On Error GoTo UpdateErr
   Select Case posicion
   Case 1
   
   Set rstbeneficiario = New ADODB.Recordset
   If rstbeneficiario.State = 1 Then rstbeneficiario.Close
   queryinicial = "select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' AND estado_codigo <> 'ANL' "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rstbeneficiario.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstbeneficiario.Sort = "beneficiario_denominacion"
   Set Ado_datos.Recordset = rstbeneficiario
   
   Case 2
   
   Set rs_aux17 = New ADODB.Recordset
    If rs_aux17.State = 1 Then rs_aux17.Close
    
    If OptFilGral1.Value = True Then
    rs_aux17.Open "select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' AND estado_codigo <> 'ANL' order by beneficiario_denominacion asc", db, adOpenKeyset, adLockOptimistic, adCmdText
    Else
    rs_aux17.Open "select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' order by beneficiario_denominacion asc", db, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    
    Set Ado_datos_busq.Recordset = rs_aux17
    dtc_buscar_ci.BoundText = dtc_buscar_desc.BoundText
    If rs_aux17.RecordCount > 0 Then
    dtc_buscar_desc.Visible = True
    Label52.Visible = True
    Else
    dtc_buscar_desc.Visible = False
    Label52.Visible = False
    End If
   
   Case 3
''''
''''   Set rs_datos2 = New ADODB.Recordset
''''   If rs_datos2.State = 1 Then rs_datos2.Close
''''   rs_datos2.Open "select * from av_gc_beneficiario_vs_ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos.Recordset!mes_grupo & "AND beneficiario_codigo = '" & dtc_buscar_ci.Text & "' order by Numero_consultoriaHist asc", db, adOpenKeyset, adLockOptimistic, adCmdText
''''   Set Ado_datos2.Recordset = rs_datos2
''''   Set dg_det2.DataSource = Ado_datos2.Recordset
   
' Call ABRIR_TABLA_DET(1)
''dg_det1.SelBookmarks.Remove (0)
''dg_det1.ClearFields
' mover = 1
'Me.dgv.Currentcell = Nothing

   If (dg_datos.SelBookmarks.Count <> 0) Then
            dg_datos.SelBookmarks.Remove 0
   End If
   
   If rstbeneficiario.RecordCount > 0 Then
   
   rstbeneficiario.Find "beneficiario_codigo = '" & dtc_buscar_ci.Text & "'", , , 1
   
   dg_datos.SelBookmarks.Add (rstbeneficiario.Bookmark)
 
 Else
 sino = MsgBox("No se encontro a nadie con ese nombre", vbInformation, "Aviso")
 Call Carga_Beneficiario(1)
 dtc_buscar_desc.Text = ""
 End If
End Select
 Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub
Private Sub Carga_afp()
   Set rstafp = New ADODB.Recordset
   If rstafp.State = 1 Then rstafp.Close
   queryinicial = "select * from gc_beneficiario WHERE hora_registro = 'AFP' "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rstafp.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstafp.Sort = "beneficiario_denominacion"
   Set adoafp.Recordset = rstafp
End Sub

Private Sub filtrar_asistencia(mes As String, ges_gestion As String)
On Error GoTo EditErr
Set rs_Asistencia = New ADODB.Recordset
    Dim rs_AsisTT As ADODB.Recordset
    Set rs_AsisTT = New ADODB.Recordset
    If rs_AsisTT.State = 1 Then rs_AsisTT.Close
    If rs_Asistencia.State = 1 Then rs_Asistencia.Close
    If cbo_mes.Text = "TODO" Then
    sqlAux = "SELECT DATENAME(month, fecha_control ) AS MES, YEAR(fecha_control) AS GESTION,  * FROM ro_ControlAsistencia where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & ges_gestion & "'"
    Else
    sqlAux = "SELECT DATENAME(month, fecha_control ) AS MES, YEAR(fecha_control) AS GESTION,  * FROM ro_ControlAsistencia where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & ges_gestion & "' AND Mes_control = '" & mes & "'"
    End If
 
    rs_Asistencia.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_Asistencia.Sort = "Fecha_control"
    Set AdoAsistencia.Recordset = rs_Asistencia
    Set DtgAsistencia.DataSource = AdoAsistencia.Recordset
       If cbo_mes.Text = "TODO" Then
    sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & ges_gestion & "'"
       Else
    sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & ges_gestion & "' AND Mes_control = '" & mes & "' "
       End If
    
    rs_AsisTT.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_AsisTT.MoveFirst
    AdoAsistencia.Caption = CStr(rs_AsisTT!totHrs)
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Public Sub abrirtabla()
    Set rs_Asistencia = New ADODB.Recordset
    Dim rs_AsisTT As ADODB.Recordset
    Set rs_AsisTT = New ADODB.Recordset
    If rs_AsisTT.State = 1 Then rs_AsisTT.Close
    If rs_Asistencia.State = 1 Then rs_Asistencia.Close
    ' Asistencia.
    sqlAux = "SELECT DATENAME(month, fecha_control ) AS MES, YEAR(fecha_control) AS GESTION,  * FROM ro_ControlAsistencia where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & Year(Date) & "' AND Mes_control = '" & Month(Date) & "'"
    rs_Asistencia.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_Asistencia.Sort = "Fecha_control"
    Set AdoAsistencia.Recordset = rs_Asistencia
    'Set DtgAsistencia.DataSource = AdoAsistencia.Recordset
        
'    sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & Year(Date) & "' AND Mes_control = '" & Month(Date) & "'"
'    rs_AsisTT.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
'    If rs_AsisTT.RecordCount > 0 Then
'    rs_AsisTT.MoveFirst
'    AdoAsistencia.Caption = CStr(rs_AsisTT!totHrs)
'    Else
'    AdoAsistencia.Caption = "0"
'    End If
    Set rs_Permisos = New ADODB.Recordset
    If rs_Permisos.State = 1 Then rs_Permisos.Close
    rs_Permisos.Open "select * from ro_Permisos where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & " order by FechaDesde", db, adOpenKeyset, adLockOptimistic
    Set AdoPermiso.Recordset = rs_Permisos
    Set DtgPermiso.DataSource = AdoPermiso.Recordset
    
    Set rs_Permiso_detalle = New ADODB.Recordset
    If rs_Permiso_detalle.State = 1 Then rs_Permiso_detalle.Close
    rs_Permiso_detalle.Open "select * from ro_Permisos_detalle where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set AdoDependiente.Recordset = rs_Permiso_detalle
'    Set DtgDependiente.DataSource = AdoDependiente.Recordset
   
    Set rs_vacaciones_prog = New ADODB.Recordset
    If rs_vacaciones_prog.State = 1 Then rs_vacaciones_prog.Close
    rs_vacaciones_prog.Open "select * from ro_vacaciones_programadas where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & " order by fecha_ini_Prog desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_VacacionesProg.Recordset = rs_vacaciones_prog
    Set DtgVacacionesProg.DataSource = Ado_VacacionesProg.Recordset
  
    Set rs_HORARIOS = New ADODB.Recordset
    If rs_HORARIOS.State = 1 Then rs_HORARIOS.Close
    rs_HORARIOS.Open "select * from RC_HORARIOS  ", db, adOpenKeyset, adLockOptimistic
    Set AdoHorarios.Recordset = rs_HORARIOS
'    Set DtgVacaciones.DataSource = AdoHorarios.Recordset
    
    Set rs_contrato = New Recordset
    If rs_contrato.State = 1 Then rs_contrato.Close
    rs_contrato.Open "select * from ro_memorandas where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & " order by correl desc ", db, adOpenKeyset, adLockOptimistic
    
    Set Ado_Memo.Recordset = rs_contrato.DataSource
    Set DtG_Memo.DataSource = Ado_Memo.Recordset
    
    Set rs_movilidad = New Recordset
    If rs_movilidad.State = 1 Then rs_movilidad.Close
    rs_movilidad.Open "select * from rv_movilidad_personal_js where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic
    Set AdoMovilidad.Recordset = rs_movilidad.DataSource
    Set DtgMovilidad.DataSource = AdoMovilidad.Recordset
    
     Set rs_datos_educacionales = New ADODB.Recordset
    If rs_datos_educacionales.State = 1 Then rs_datos_educacionales.Close
    rs_datos_educacionales.Open "select * from ro_datos_educacionales where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' order by fecha_inicio desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Educacionales.Recordset = rs_datos_educacionales
    Set DtgEducacionales.DataSource = Ado_Educacionales.Recordset
  
    Set rs_laborales = New ADODB.Recordset
    If rs_laborales.State = 1 Then rs_laborales.Close
    rs_laborales.Open "select * from ro_experiencia_laboral where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' order by fecha_inicio desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Laborales.Recordset = rs_laborales
    Set DtgLaborales.DataSource = Ado_Laborales.Recordset
    
    Set rs_contrato = New Recordset
    If rs_contrato.State = 1 Then rs_contrato.Close
    rs_contrato.Open "select * from ro_contratos_personas where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' order by fecha_inicio desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Contrato.Recordset = rs_contrato.DataSource
    Set DtG_Contrato.DataSource = Ado_Contrato.Recordset
    
    Set rs_liquidacion = New Recordset
    If rs_liquidacion.State = 1 Then rs_liquidacion.Close
    rs_liquidacion.Open "select * from ro_liquidaciones where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & " order by fecha_ingreso desc ", db, adOpenKeyset, adLockOptimistic
    Set AdoLiquidacion.Recordset = rs_liquidacion.DataSource
    Set DtgLiquidacion.DataSource = AdoLiquidacion.Recordset

    Set rs_CtaPersonal = New Recordset
    If rs_CtaPersonal.State = 1 Then rs_CtaPersonal.Close
    rs_CtaPersonal.Open "select * from rv_personal_cuenta_bancaria3 where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'   ", db, adOpenKeyset, adLockOptimistic       'order by cta_para_abono desc
    Set Ado_CtaPersonal.Recordset = rs_CtaPersonal.DataSource
    Set DtgCuentaBanco.DataSource = Ado_CtaPersonal.Recordset

End Sub

Private Sub Form_Resize()
'   '  Centrear titulo
'   With lbl_titulo
'      .Left = (fraOpciones.Width - .Width) \ 2
'   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If glPersNew = "P" Then
'    FrmVentas.DtcNIT = rstbeneficiario!beneficiario_codigo
''    FrmVentas.Dtc_pers_1apell = rstbeneficiario!paterno_beneficiario
''    FrmVentas.Dtc_pers_2Apell = rstbeneficiario!materno_beneficiario
''    FrmVentas.Dtc_Pers_nombre = rstbeneficiario!nombres_beneficiario
'    FrmVentas.DtcdesNIT = rstbeneficiario!beneficiario_denominacion
  End If
  If glPersNew = "CMP" Then
     'frmComprasDirectas.DtcNIT = rstbeneficiario!beneficiario_codigo
     'frmComprasDirectas.DtcdesNIT = rstbeneficiario!beneficiario_denominacion
     
'    Set frmComprasDirectas.recSetAuxbenefi1 = New ADODB.Recordset
'    If frmComprasDirectas.recSetAuxbenefi1.State = 1 Then frmComprasDirectas.recSetAuxbenefi1.Close
'    frmComprasDirectas.recSetAuxbenefi1.Open "select * from Gc_beneficiario  ", db, adOpenKeyset, adLockReadOnly
'    frmComprasDirectas.adoProveedores.Recordset.Requery
    
'    deCD.dbo_cdListaProveedores
'    With deCD.rsdbo_cdListaProveedores
'        While Not .EOF
'            frmComprasDirectas.cboListaProv.AddItem !beneficiario_denominacion
'            frmComprasDirectas.cboListaProv2.AddItem !beneficiario_denominacion
'            .MoveNext
'        Wend
'    End With
'    deCD.rsdbo_cdListaProveedores.Close
    
'    frmmo_formulario_M1.Dtc_doc_id_lab = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell_lab = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2apell_lab = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre_lab = rs_Personal!pers_nombres
  End If
'  If glPersNew = "PL" Then
'    frmeo_Larvas_mosquitos.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_Larvas_mosquitos.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_Larvas_mosquitos.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_Larvas_mosquitos.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
'  If glPersNew = "PMA" Then
'    frmeo_mosquito_adulto.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_mosquito_adulto.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_mosquito_adulto.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_mosquito_adulto.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
  glPersNew = "N"
   
   If (rstbeneficiario.State = adStateClosed) Then rstbeneficiario.Close
   'Set rstbeneficiario = Nothing

End Sub

'Private Sub Option1_Click()
'   TxtTipo.Text = "R"
'   Frame2.Visible = False
'
'End Sub
'
'Private Sub Option2_Click()
'   TxtTipo.Text = "C"
'   Frame2.Visible = True
'End Sub

Private Sub Carga_Recor()
  'carga    fc_tipoben_codigo
    Set rst_ben = New ADODB.Recordset
    rst_ben.Open "SELECT * FROM gc_Tipo_Beneficiario ORDER BY tipoben_descripcion ", db, adOpenStatic
    Set AdoTip_ben.Recordset = rst_ben
    
    Set rs_Depto = New ADODB.Recordset
    rs_Depto.Open "select * from gc_Departamento", db, adOpenKeyset, adLockOptimistic
    Set Ado_Depto.Recordset = rs_Depto
    dtc_depto.BoundText = Dtc_depto_cod.BoundText
    
    Set rs_Prov = New ADODB.Recordset
    rs_Prov.Open "select * from GC_Provincia", db, adOpenKeyset, adLockOptimistic
    Set Ado_prov.Recordset = rs_Prov
    Dtc_prov.BoundText = Dtc_prov_cod.BoundText
    
    Set rs_Muni = New ADODB.Recordset
    rs_Muni.Open "select * from gc_Municipio ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Muni.Recordset = rs_Muni
    Dtc_munic.BoundText = Dtc_munic_cod.BoundText
    
    Set rs_aux7 = New ADODB.Recordset
    rs_aux7.Open "select * from gc_unidad_ejecutora ", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos1.Recordset = rs_aux7
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_CARGO = New ADODB.Recordset
    rs_CARGO.Open "select * from rc_cargos ", db, adOpenKeyset, adLockOptimistic
    Set AdoCargo.Recordset = rs_CARGO
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
'    Set rs_Puesto = New ADODB.Recordset
'    rs_Puesto.Open "select * from rc_puestos ", db, adOpenKeyset, adLockOptimistic
'    Set AdoPuestoOrg.Recordset = rs_Puesto
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'
'    Set rs_comunid = New ADODB.Recordset
'    rs_comunid.Open "select * from GC_comunidad ", DB, adOpenKeyset, adLockOptimistic
'    Set Ado_Comunid.Recordset = rs_comunid
'    Dtc_local.BoundText = Dtc_local_cod.BoundText

    Set rs_Depto2 = New ADODB.Recordset
    rs_Depto2.Open "select * from gc_Departamento", db, adOpenKeyset, adLockOptimistic
    Set Ado_Depto2.Recordset = rs_Depto2
'    Dtc_depto2.BoundText = Dtc_depto_cod2.BoundText
    
    Set rs_Prov2 = New ADODB.Recordset
    rs_Prov2.Open "select * from GC_Provincia", db, adOpenKeyset, adLockOptimistic
    Set Ado_prov2.Recordset = rs_Prov2
'    Dtc_prov2.BoundText = Dtc_prov_cod2.BoundText
    
    Set rs_Muni2 = New ADODB.Recordset
    rs_Muni2.Open "select * from gc_Municipio ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Muni2.Recordset = rs_Muni2
'    Dtc_munic2.BoundText = Dtc_munic_cod2.BoundText
    
'    Set rs_comunid2 = New ADODB.Recordset
'    rs_comunid2.Open "select * from GC_comunidad ", db, adOpenKeyset, adLockOptimistic
'    Set Ado_Comunid2.Recordset = rs_comunid2
''    Dtc_local2.BoundText = Dtc_local_cod2.BoundText
    
    Set rs_TipoDocId = New ADODB.Recordset
    rs_TipoDocId.Open "select * from gc_tipo_documento_id where estado_codigo ='APR' ", db, adOpenKeyset, adLockOptimistic
    Set Ado_TipoDocId.Recordset = rs_TipoDocId
    
    Set rs_Depto3 = New ADODB.Recordset
    rs_Depto3.Open "select * from gc_Departamento", db, adOpenKeyset, adLockOptimistic
    Set Ado_Depto3.Recordset = rs_Depto3
    'Dtc_depto2.BoundText = Dtc_depto_cod2.BoundText
    
    Set rs_nivel_educacional = New ADODB.Recordset
    rs_nivel_educacional.Open "select * from rc_nivel_educacional ", db, adOpenKeyset, adLockOptimistic
    Set AdoNivelEducacional.Recordset = rs_nivel_educacional
    
    Set rs_tipoInstitucion = New ADODB.Recordset
    rs_tipoInstitucion.Open "select * from rc_tipo_institucion ", db, adOpenKeyset, adLockOptimistic
    Set Ado_TipoInstitucion.Recordset = rs_tipoInstitucion
    
   Set rs_beneficiario = New ADODB.Recordset
   If rs_beneficiario.State = 1 Then rs_beneficiario.Close
   rs_beneficiario.Open "select * from gc_Beneficiario WHERE hora_registro = 'SS'", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_Benef_seguro.Recordset = rs_beneficiario
   
   Set rs_CTA_BCO = New ADODB.Recordset
   If rs_CTA_BCO.State = 1 Then rs_CTA_BCO.Close
   rs_CTA_BCO.Open "select * from fv_cuenta_bco WHERE cta_codigo_tgn = '000'", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set AdoCta.Recordset = rs_CTA_BCO
   
    Set rs_ocupacion = New ADODB.Recordset
    If rs_ocupacion.State = 1 Then rs_ocupacion.Close
    rs_ocupacion.Open "select * from gc_ocupacion_profesion WHERE estado_codigo = 'APR' ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Ocupacion.Recordset = rs_ocupacion
    TxtProfesion.BoundText = Dtc_Ocup.BoundText
    
   Set rs_beneficiario_Afp = New ADODB.Recordset
   If rs_beneficiario_Afp.State = 1 Then rs_beneficiario_Afp.Close
   rs_beneficiario_Afp.Open "select * from gc_Beneficiario WHERE tipoben_codigo = '22'", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_Benef_Afp.Recordset = rs_beneficiario_Afp
      
   Set rs_EstCivil = New ADODB.Recordset
   If rs_EstCivil.State = 1 Then rs_EstCivil.Close
   rs_EstCivil.Open "select * from rc_estado_civil ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set AdoEstCivil.Recordset = rs_EstCivil
   If AdoEstCivil.Recordset.RecordCount > 0 Then
   End If
   
   
   Set rs_pais = New ADODB.Recordset
   rs_pais.Open "SELECT * FROM gc_pais ORDER BY pais_descripcion ", db, adOpenStatic
   Set AdoPais.Recordset = rs_pais
   
   Set rs_calendario = New ADODB.Recordset
   rs_calendario.Open "SELECT * FROM gc_calendario ", db, adOpenStatic
   Set AdoCalendario.Recordset = rs_calendario
   
   'PLANILLA
'   Set rs_datos5 = New ADODB.Recordset
'   If rs_datos5.State = 1 Then rs_datos5.Close
'   rs_datos5.Open "select * from rc_planilla_grupo ", db, adOpenKeyset, adLockOptimistic, adCmdText
'   Set Ado_datos5.Recordset = rs_datos5
   'SUB PLANILLA
   Set rs_datos6 = New ADODB.Recordset
   If rs_datos6.State = 1 Then rs_datos6.Close
   rs_datos6.Open "select * from rv_rc_planilla_vs_rc_sub_planilla ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_datos6.Recordset = rs_datos6
   dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'GENERO
    Set rs_datos4 = New ADODB.Recordset
   If rs_datos4.State = 1 Then rs_datos4.Close
   rs_datos4.Open "select * from gc_genero ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_datos4.Recordset = rs_datos4
   dtc_desc4.BoundText = dtc_codigo4.BoundText
   
'  Set rs_datos7 = New ADODB.Recordset
'   If rs_datos7.State = 1 Then rs_datos7.Close
'   rs_datos7.Open "select * from rc_planilla_grupo ", db, adOpenKeyset, adLockOptimistic, adCmdText
'   Set Ado_datos7.Recordset = rs_datos7
'   dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub


Private Sub Img_03_Click()
 If AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo" Then
    MsgBox "No Existe el Archivo asociado al Registro, debe Cargarlo ...", vbExclamation, "Advertencia"
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
   If GlServidor = "SRVPRO" Then
      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      Else
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      End If
   Else
      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      Else
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      End If
   End If
 End If

End Sub

Private Sub Img_CTO_Click()
 If Ado_Memo.Recordset!ARCHIVO = "Cargar_Archivo" Then
    MsgBox "No Existe el Archivo Asociado al Contrato, debe Cargarlo ...", vbExclamation, "Advertencia"
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SRVPRO" Then
        'e = ShellExecute(Img_CTO, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Memo.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Memo.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        'e = ShellExecute(Img_CTO, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Memo.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Memo.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ShellExecute(0, vbNullString, "c:\Archivo.PDF", vbNullString, vbNullString, vbNormalFocus)
'System.Diagnostics.Process.Start("c:\Archivo.PDF")
End Sub

Private Sub Img_CV_Click()
'    Dim e As Long
  If swnuevo <> 0 Then
    If Ado_datos.Recordset!archivo_hojavida = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "C_V"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\"
         ' e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.AdoMovilidad.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      Else
         e = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "C_V"
          'If GlServidor <> GlMaquina Then      ' "-" Then
          If GlServidor = "SRVPRO" Then
            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\"
          Else
            e = NombreCarpeta
          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If
  End If
  'If GlServidor <> GlMaquina Then      ' "-" Then
  If GlServidor = "SRVPRO" Then
        'imag2 = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(Ado_datos.Recordset!ARCHIVO_HOJAVIDA), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
  Else
        'imag2 = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(Ado_datos.Recordset!ARCHIVO_HOJAVIDA), vbNullString, vbNullString, SW_SHOWNORMAL)
        'Call ShellExecute(Me.hwnd, "Open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(Ado_datos.Recordset!ARCHIVO_HOJAVIDA), vbNullString, vbNullString, SW_SHOWNORMAL)
        'imag2 = ShellExecute(Me.hwnd, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(Ado_datos.Recordset!ARCHIVO_HOJAVIDA), "", "", 1)
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
        'ShellExecute(0, vbNullString, "c:\Archivo.PDF", vbNullString, vbNullString, vbNormalFocus)
        'System.Diagnostics.Process.Start("c:\Archivo.PDF")
        'pdfshell.dll
        'support.microsoft.com/kb/238245/es
        'support.microsoft.com/kb/114038/es
        'http://www.mygnet.net/codigos/vbdotnet/manipulacion_objetos/abrir_un_archivo_excel_desde_visual_basic_dot_net.2509
  End If
End Sub

Private Sub btnEjecutar_Click()
    ' Ejecutar un acceso directo
'    Call ShellExecute(Me.hwnd, "Open", Text1.Text, "", "", 1)

End Sub


Private Sub Img_DocRespaldo_Click()
'  If swnuevo <> 0 Then
'    If Ado_datos.Recordset!ARCHIVO_RESPALDO = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "D_R"
'      'If GlServidor <> GlMaquina Then      ' "-" Then
'      If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
'      Else
'            e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "D_R"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'  End If
'    'If GlServidor <> GlMaquina Then      ' "-" Then
'    If GlServidor = "SRVPRO" Then
'        'e = ShellExecute(Img_DocRespaldo, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(Ado_datos.Recordset!ARCHIVO_RESPALDO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(Ado_datos.Recordset!ARCHIVO_RESPALDO), vbNullString, vbNullString, vbNormalFocus)
'    Else
'        'e = ShellExecute(Img_DocRespaldo, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(Ado_datos.Recordset!ARCHIVO_RESPALDO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(Ado_datos.Recordset!ARCHIVO_RESPALDO), vbNullString, vbNullString, vbNormalFocus)
'    End If
End Sub

Private Sub Img_Foto_Click()
  If swnuevo <> 0 Then
    If Ado_datos.Recordset!ARCHIVO_Foto = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FOT"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      Else
         e = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FOT"
          'If GlServidor <> GlMaquina Then      ' "-" Then
          If GlServidor = "SRVPRO" Then
            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          Else
            e = NombreCarpeta
          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If
  
    Dim ARCH_FOTO As String
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SRVPRO" Then
        ARCH_FOTO = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
    Else
        ARCH_FOTO = App.Path + "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
    End If
    'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!beneficiario_codigo + "\" + Ado_datos.Recordset("beneficiario_codigo") + "-FOTO.JPG"
    If Guardar_Imagen(db, "Select Foto From rv_personal_contratado Where beneficiario_codigo= '" & Ado_datos.Recordset("beneficiario_codigo") & "' ", "Foto", ARCH_FOTO) Then
        MsgBox "Se cargo la Imagen Correctamente !!"
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
  End If
End Sub

Private Sub ImgFiniquito_Click()
 If AdoMovilidad.Recordset!ARCHIVO = "Cargar_Archivo" Then
    MsgBox "No Existe el Archivo Asociado a la Liquidación, debe Cargarlo ...", vbExclamation, "Advertencia"
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SRVPRO" Then
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoMovilidad.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(AdoMovilidad.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        'e = ShellExecute(Img_CTO, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoMovilidad.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(AdoMovilidad.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoMovilidad.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(AdoMovilidad.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub

Private Sub ImgMemo_Click()
'    Dim e As Long
'    If GlServidor <> GlMaquina Then      ' "-" Then
'        e = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\MEMORANDUMS\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-Memo-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    Else
'        e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\MEMORANDUMS\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-Memo-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    End If
End Sub

Private Sub ImgVacacion_Click()
'    Dim e As Long
'    If GlServidor <> GlMaquina Then      ' "-" Then
'        e = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-Vacacion-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    Else
'        e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-Vacacion-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    End If
End Sub

Private Sub OptFilGral1_Click()
   Set rstbeneficiario = New ADODB.Recordset
   If rstbeneficiario.State = 1 Then rstbeneficiario.Close
   queryinicial = "select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' AND estado_codigo <> 'ANL' "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rstbeneficiario.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstbeneficiario.Sort = "beneficiario_denominacion"
   Set Ado_datos.Recordset = rstbeneficiario
   Set dg_datos.DataSource = Ado_datos.Recordset
   Call Carga_Beneficiario(2)
   dtc_buscar_desc.Text = ""
If Ado_datos.Recordset.RecordCount > 0 Then


   If Ado_datos.Recordset!fecha_expiracion <= Date Then
        Label18.Visible = True
        'Label18.Caption = "Dado de Baja el " & Ado_datos.Recordset!fecha_expiracion & " Se quitara de la pantalla de Vigentes Depues de generar la siguiente planilla"
        Label18.Caption = "Dado de Baja el " & Ado_datos.Recordset!fecha_expiracion
    Else
        Label18.Visible = False
        Label18.Caption = ""
    End If
    
End If
End Sub

Private Sub OptFilGral2_Click()
Set rstbeneficiario = New ADODB.Recordset
   If rstbeneficiario.State = 1 Then rstbeneficiario.Close
   queryinicial = "select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rstbeneficiario.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstbeneficiario.Sort = "beneficiario_denominacion"
   Set Ado_datos.Recordset = rstbeneficiario
   Set dg_datos.DataSource = Ado_datos.Recordset
   Call Carga_Beneficiario(2)
   Ado_datos.Recordset.MoveFirst
   dtc_buscar_desc.Text = ""
   
   If Ado_datos.Recordset.RecordCount > 0 Then


   If Ado_datos.Recordset!fecha_expiracion <= Date Then
        Label18.Visible = True
        Label18.Caption = "Dado de Baja el " & Ado_datos.Recordset!fecha_expiracion
        Else
        Label18.Visible = False
        Label18.Caption = ""
    End If
    
End If
End Sub

Private Sub SSTab1_DblClick()
    If SSTab1.Tab = 0 Then
    End If
End Sub

Private Sub TDBNivelEdu_DropDownClose()
'    DtgVacacionesProg.Columns("nivel_educacional").Value = TDBNivelEdu.Columns("nivel_educacional").Value
'    'DtgVacacionesProg.Columns("descripcion").Value = TDBNivelEdu.Columns("descripcion").Value
End Sub

Private Sub TDBtipoben_Click(Area As Integer)
    TxtTipo.BoundText = TDBtipoben.BoundText
End Sub

Private Sub TDBtipoben_LostFocus()
    If TxtTipo.Text = "6" Then
        txtDenominacion.Enabled = True
'        Label2.Caption = "Empresa/Instit."
'        Frame2.Caption = "Datos del Representante Legal"
'        Label14.Caption = "Fecha de Creación"
'        LlbCargo.Caption = "Actividad Principal"
'        LblProf_Asoc.Caption = "Camara/Asociación a la que Pertenece"
'        Label13.Caption = "Nro. Registro"
    Else
        txtDenominacion.Enabled = False
'        Label2.Caption = "Denominación"
'        Frame2.Caption = "Datos de la Persona"
'        Label14.Caption = "Fecha de Nacimiento"
'        LlbCargo.Caption = "Cargo que Ocupa"
'        LblProf_Asoc.Caption = "Profesion u Ocupacion:"
'        Label13.Caption = "Nro. Empresa"
    End If
        DtgVacacionesProg.AllowAddNew = False
        DtgVacacionesProg.AllowDelete = False
        DtgVacacionesProg.AllowUpdate = False
'        DtgVacaciones.AllowAddNew = False
'        DtgVacaciones.AllowDelete = False
'        DtgVacaciones.AllowUpdate = False
'        CmdAdd.Visible = False
'        CmdMod.Visible = False
'        CmdGraba.Visible = False
'        CmdAdd2.Visible = False
'        CmdMod2.Visible = False
'        CmdGraba2.Visible = False
'    If TxtTipo.Text = "1" Then
'        TxtRenca.Visible = False
'        'TxtRenca.BackColor =&H8000000B&
'        DTP_FechaExpira.Visible = False
'        Label13.Visible = False
'        Label10.Visible = False
'    Else
'        TxtRenca.Visible = True
'        'TxtRenca.BackColor =&H8000000B&
'        DTP_FechaExpira.Visible = True
'        Label13.Visible = True
'        Label10.Visible = True
'    End If
End Sub

Private Sub TDBTipoInst_DropDownClose()
'    DtgVacaciones.Columns("tipo_institucion").Value = TDBTipoInst.Columns("tipo_institucion").Value
End Sub

Private Sub TDBTipoInst_Click()

End Sub

'Private Sub Text102_Change()
'    Text102.BackColor = &H80000014
'End Sub

'Private Sub Text2_LostFocus()
'    txtDenominacion.Text = Text1.Text + " " + Text2.Text + " " + Text3.Text
'End Sub

'Private Sub Text202_Change()
'    Text202.BackColor = &H80000014
'End Sub

'Private Sub Text3_LostFocus()
'    txtDenominacion.Text = Text1.Text + " " + Text2.Text + " " + Text3.Text
'End Sub

Private Sub TxtNacionalidad_Click(Area As Integer)
    DtcPaisCod.BoundText = TxtNacionalidad.BoundText
    DtcPaisSigla.BoundText = TxtNacionalidad.BoundText
End Sub

'Private Sub Text302_Change()
'    Text302.BackColor = &H80000014
'End Sub

'Private Sub TxtNIT2_Click(Area As Integer)
'    DtcRep_Nombres.BoundText = TxtNIT2.BoundText
'    DtcRep_Paterno.BoundText = TxtNIT2.BoundText
'    DtcRep_Materno.BoundText = TxtNIT2.BoundText
'End Sub

Private Sub txtProfesion_Click(Area As Integer)
    Dtc_Ocup.BoundText = TxtProfesion.BoundText
End Sub

Private Sub TxtTipo_Click(Area As Integer)
    TDBtipoben.BoundText = TxtTipo.BoundText
End Sub




Private Function ExisteReg(where As String, tabla As String) As Boolean
        Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM " & tabla & " WHERE " & where & ""
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
    
End Function






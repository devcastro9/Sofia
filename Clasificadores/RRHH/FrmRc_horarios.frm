VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRc_Horarios 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Administrativos - Horario Laboral"
   ClientHeight    =   7680
   ClientLeft      =   1065
   ClientTop       =   2415
   ClientWidth     =   13680
   Icon            =   "FrmRc_horarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmRc_horarios.frx":0A02
   ScaleHeight     =   7680
   ScaleWidth      =   13680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "FrmRc_horarios.frx":6CA34
      ScaleHeight     =   960
      ScaleWidth      =   13395
      TabIndex        =   26
      Top             =   120
      Width           =   13460
      Begin VB.CommandButton cmdBorrar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "FrmRc_horarios.frx":D8A66
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmd_busqueda 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "FrmRc_horarios.frx":D9730
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   5160
         Picture         =   "FrmRc_horarios.frx":D9CE8
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   720
         Left            =   3480
         Picture         =   "FrmRc_horarios.frx":D9EF2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "FrmRc_horarios.frx":DA60D
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   1800
         Picture         =   "FrmRc_horarios.frx":DA817
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdIMPRIMIR 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "FrmRc_horarios.frx":DAA21
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "FrmRc_horarios.frx":DAFDE
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdAdicionar 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "FrmRc_horarios.frx":DB5BE
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORARIO LABORAL"
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
         Left            =   9015
         TabIndex        =   36
         Top             =   300
         Width           =   3075
      End
   End
   Begin VB.Frame FraHorario 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3150
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   13455
      Begin VB.Frame FraHora2 
         BackColor       =   &H00000000&
         Caption         =   "SEGUNDO HORARIO"
         ForeColor       =   &H000080FF&
         Height          =   1215
         Left            =   6700
         TabIndex        =   20
         Top             =   1800
         Width           =   6735
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            DataField       =   "hora_tope_ingreso2"
            DataSource      =   "adoLista"
            Height          =   285
            Left            =   3720
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            DataField       =   "hora_salida2"
            DataSource      =   "adoLista"
            Height          =   285
            Left            =   1920
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            DataField       =   "hora_ingreso2"
            DataSource      =   "adoLista"
            Height          =   285
            Left            =   240
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox Txt02 
            DataField       =   "minutos_tolerancia2"
            DataSource      =   "adoLista"
            Height          =   315
            ItemData        =   "FrmRc_horarios.frx":DBBE2
            Left            =   5520
            List            =   "FrmRc_horarios.frx":DBC07
            TabIndex        =   21
            Text            =   "10"
            Top             =   600
            Width           =   735
         End
         Begin MSComCtl2.DTPicker txt06 
            DataField       =   "hora_ingreso2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   4
            EndProperty
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   86769666
            UpDown          =   -1  'True
            CurrentDate     =   0.333333333333333
            MaxDate         =   0.999988425925926
            MinDate         =   4.16666666666667E-02
         End
         Begin MSComCtl2.DTPicker txt07 
            DataField       =   "hora_salida2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   4
            EndProperty
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   1920
            TabIndex        =   23
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   86769666
            UpDown          =   -1  'True
            CurrentDate     =   0.333333333333333
            MaxDate         =   0.999988425925926
            MinDate         =   4.16666666666667E-02
         End
         Begin MSComCtl2.DTPicker txt08 
            DataField       =   "hora_tope_ingreso2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   4
            EndProperty
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   3720
            TabIndex        =   24
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   86769666
            UpDown          =   -1  'True
            CurrentDate     =   0.333333333333333
            MaxDate         =   0.999988425925926
            MinDate         =   4.16666666666667E-02
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Ingreso 2:              Hora Salida 2:               Hora Tope Ingreso2:     Min.Tolerancia2"
            ForeColor       =   &H00FFFF80&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   6240
         End
      End
      Begin VB.Frame FraHora1 
         BackColor       =   &H00000000&
         Caption         =   "PRIMER HORARIO"
         ForeColor       =   &H000080FF&
         Height          =   1215
         Left            =   25
         TabIndex        =   14
         Top             =   1800
         Width           =   6670
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            DataField       =   "hora_tope_ingreso"
            DataSource      =   "adoLista"
            Height          =   285
            Left            =   3720
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            DataField       =   "hora_salida"
            DataSource      =   "adoLista"
            Height          =   285
            Left            =   1920
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            DataField       =   "hora_ingreso"
            DataSource      =   "adoLista"
            Height          =   285
            Left            =   240
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox txt01 
            DataField       =   "minutos_tolerancia"
            DataSource      =   "adoLista"
            Height          =   315
            ItemData        =   "FrmRc_horarios.frx":DBC37
            Left            =   5520
            List            =   "FrmRc_horarios.frx":DBC5C
            TabIndex        =   15
            Text            =   "10"
            Top             =   600
            Width           =   735
         End
         Begin MSComCtl2.DTPicker txt03 
            DataField       =   "hora_ingreso"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   4
            EndProperty
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   86769666
            UpDown          =   -1  'True
            CurrentDate     =   0.333333333333333
            MaxDate         =   0.999988425925926
            MinDate         =   4.16666666666667E-02
         End
         Begin MSComCtl2.DTPicker txt04 
            DataField       =   "hora_salida"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "hh:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   4
            EndProperty
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   86769666
            CurrentDate     =   0.5
            MaxDate         =   0.999988425925926
            MinDate         =   4.16666666666667E-02
         End
         Begin MSComCtl2.DTPicker txt05 
            DataField       =   "hora_tope_ingreso"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "hh:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   4
            EndProperty
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   3720
            TabIndex        =   18
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   86769666
            CurrentDate     =   0.333333333333333
            MaxDate         =   0.999988425925926
            MinDate         =   4.16666666666667E-02
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Ingreso:                 Hora Salida:                    Hora Tope Ingreso:      Min.Tolerancia"
            ForeColor       =   &H00FFFF80&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   6195
         End
      End
      Begin VB.ComboBox Txt09 
         DataField       =   "nro_horarios"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_horarios.frx":DBC8C
         Left            =   9360
         List            =   "FrmRc_horarios.frx":DBC99
         TabIndex        =   10
         Text            =   "1"
         Top             =   840
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPFec_Fin 
         DataField       =   "fecha_hasta"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   9360
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   86769665
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_horarios.frx":DBCA6
         Left            =   1080
         List            =   "FrmRc_horarios.frx":DBCBC
         TabIndex        =   2
         Text            =   "2015"
         Top             =   360
         Width           =   900
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         DataField       =   "fecha_desde"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   3480
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   86769665
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         Bindings        =   "FrmRc_horarios.frx":DBCE4
         DataField       =   "turno_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   10440
         TabIndex        =   4
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "turno_codigo"
         BoundColumn     =   "turno_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         Bindings        =   "FrmRc_horarios.frx":DBCFB
         DataField       =   "turno_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "turno_descripcion"
         BoundColumn     =   "turno_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_Hrs 
         Bindings        =   "FrmRc_horarios.frx":DBD12
         DataField       =   "turno_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   3480
         TabIndex        =   13
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "turno_nro_horas"
         BoundColumn     =   "turno_codigo"
         Text            =   ""
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Fin de Control (Hasta): "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   4
         Left            =   6360
         TabIndex        =   12
         Top             =   1335
         Width           =   2925
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.de Horarios (1er. - 2do.):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   2
         Left            =   6720
         TabIndex        =   11
         Top             =   855
         Width           =   2520
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Total de Horas p/día Laboral:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   855
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion:                                                        Tipo de Horario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   45
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4725
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio de Control (Desde): "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   40
         Left            =   240
         TabIndex        =   6
         Top             =   1335
         Width           =   3195
      End
   End
   Begin MSAdodcLib.Adodc adoLista 
      Height          =   330
      Left            =   120
      Top             =   3960
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   12648447
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
   Begin MSDataGridLib.DataGrid grdlista 
      Bindings        =   "FrmRc_horarios.frx":DBD29
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   4789
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "Ges_gestion"
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
      BeginProperty Column01 
         DataField       =   "horario_descripcion"
         Caption         =   "Turno"
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
         DataField       =   "fecha_desde"
         Caption         =   "Fecha Desde"
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
         DataField       =   "fecha_hasta"
         Caption         =   "Fecha Hasta"
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
         DataField       =   "hora_ingreso"
         Caption         =   "Hora Ingreso"
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
         DataField       =   "hora_salida"
         Caption         =   "Hora Salida"
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
         DataField       =   "minutos_tolerancia"
         Caption         =   "Min.Toler"
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
         DataField       =   "hora_tope_ingreso"
         Caption         =   "Tope Ingreso"
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
         DataField       =   "hora_ingreso2"
         Caption         =   "Hora Ingreso2"
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
         DataField       =   "hora_salida2"
         Caption         =   "Hora Salida2"
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
         DataField       =   "minutos_tolerancia2"
         Caption         =   "Min.Toler.2"
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
         DataField       =   "fecha_registro"
         Caption         =   "fecha_registro"
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
         DataField       =   "hora_registro"
         Caption         =   "hora_registro"
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
      BeginProperty Column13 
         DataField       =   "usr_usuario"
         Caption         =   "usr_usuario"
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
      BeginProperty Column14 
         DataField       =   "hora_tope_ingreso2"
         Caption         =   "Tope Ingreso2"
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
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTurno 
      Height          =   375
      Left            =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "AdoTurno"
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
      Left            =   2520
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRc_Horarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstorg As New ADODB.Recordset
Dim rs_turnos As New ADODB.Recordset

Dim CAMPOS As ADODB.Field
'Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim sql_financiador As String
Dim sw2 As String

Private Sub Adolista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'   If pRecordset.EOF Or pRecordset.BOF Then
'      cmdEditar.Enabled = False
'      cmdBorrar.Enabled = False
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
'      Text4.Text = Empty
'      dtcfu.Text = ""
'      dtcfue.Text = ""
'      Exit Sub
'   End If
   
'   cmdEditar.Enabled = True
'   cmdBorrar.Enabled = True
'
'   Select Case pRecordset.EditMode
'      Case adEditInProgress
'      Case adEditNone
'         Txt01.Text = IIf(IsNull(pRecordset("Dia_control")), "", pRecordset("Dia_control"))
'         TxtGestion.Text = IIf(IsNull(pRecordset("ges_gestion")), "", pRecordset("ges_gestion"))
'         txt02.Text = IIf(IsNull(pRecordset("minutos_tolerancia")), "", pRecordset("minutos_tolerancia"))
'         txt03.Value = IIf(IsNull(pRecordset("primer_ingreso")), "", pRecordset("primer_ingreso"))
'         txt04.Value = IIf(IsNull(pRecordset("primera_salida")), "", pRecordset("primera_salida"))
'         txt05.Value = IIf(IsNull(pRecordset("tope_primer_ingreso")), "", pRecordset("tope_primer_ingreso"))
'         Txt06.Text = IIf(IsNull(pRecordset("Hora_extra_AM")), "", pRecordset("Hora_extra_AM"))
'         txt07.Value = IIf(IsNull(pRecordset("segundo_ingreso")), "", pRecordset("segundo_ingreso"))
'         txt08.Value = IIf(IsNull(pRecordset("segunda_salida")), "", pRecordset("segunda_salida"))
'         txt09.Value = IIf(IsNull(pRecordset("tope_segundo_ingreso")), "", pRecordset("tope_segundo_ingreso"))
'         txt10.Text = IIf(IsNull(pRecordset("Hora_extra_PM")), "", pRecordset("Hora_extra_PM"))
'         DTPFec_Inicio.Value = IIf(IsNull(pRecordset("fecha_registro")), "", pRecordset("fecha_registro"))
'         TxtHora.Text = IIf(IsNull(pRecordset("hora_registro")), "", pRecordset("hora_registro"))
''         TxtUsuario.Text = IIf(IsNull(pRecordset("usr_usuario")), "", pRecordset("usr_usuario"))
'      Case adEditDelete
'      Case adEditAdd
'   End Select
   Adolista.Caption = CStr(Adolista.Recordset.AbsolutePosition) & " de " & CStr(Adolista.Recordset.RecordCount)
End Sub
   
Private Sub CmdAceptar_Click()
On Error GoTo errorAceptar
Dim sw As Boolean
Dim SQL_FOR As String
Dim ctrl1, ctrl2 As Integer
Dim RSTORAUX As New ADODB.Recordset

   With Adolista
        If TxtGestion = "" Then
              MsgBox "INTRODUZCA DATOS"
              TxtGestion.SetFocus
              Exit Sub
        End If
        If Txt01 = "" Then
              MsgBox "INTRODUZCA DATOS"
              Txt01.SetFocus
              Exit Sub
        End If
        If txt02 = "" Then
              MsgBox "INTRODUZCA DATOS"
              txt02.SetFocus
              Exit Sub
        End If
        If txt03 = "" Then
              MsgBox "INTRODUZCA DATOS"
              txt03.SetFocus
              Exit Sub
        End If
        If txt03 > txt04 Then
              MsgBox "La Hora de INGRESO, NO puede ser mayor a la fecha de SALIDA.."
              txt03.SetFocus
              Exit Sub
        End If
        If DTPFec_Inicio > DTPFec_Fin Then
              MsgBox "La fecha inicial de Control, NO puede ser mayor a la fecha Final de Control.."
              DTPFec_Inicio.SetFocus
              Exit Sub
        End If
        If (txt05 < txt03) Or (txt05 > txt04) Then
              MsgBox "La fecha tope debe ser mayor a la de INGRESO, ni menor a la de SALIDA.."
              txt05.SetFocus
              Exit Sub
        End If
'        If (txt09 < txt07) Or (txt09 > txt08) Then
'              MsgBox "La fecha tope debe ser mayor a la de INGRESO y menor a la de SALIDA.."
'              txt09.SetFocus
'              Exit Sub
'        End If
'
'    Set RSTORAUX = New ADODB.Recordset
'    SQL_FOR = "select * from Fc_ORGANISMO_FINANCIAMIENTO where ORG_CODIGO = '" & Text1.Text & "'"
'    RSTORAUX.Open SQL_FOR, DB, adOpenKeyset, adLockOptimistic, adCmdText
'    If RSTORAUX.RecordCount > 0 And Text1.Enabled Then
'      sw = True
'      MsgBox " CODIGO DUPLICADO"
'      Text1.SetFocus
'      Exit Sub
'    End If
    '
    'DB.BeginTrans
    sw = False
    If sw2 = "ADD" Then
'        .Recordset.AddNew
        .Recordset("ges_gestion").Value = TxtGestion.Text
        .Recordset("turno") = Dtc_Par.Text
    End If
    .Recordset("descripcion").Value = Dtc_ParDes.Text
    .Recordset("fecha_desde").Value = DTPFec_Inicio.Value
    .Recordset("fecha_hasta").Value = DTPFec_Fin.Value
    .Recordset("hora_extra").Value = "NO" '(txt01.Text)
    .Recordset("hora_ingreso").Value = Format(txt03.Value, "HH:mm:ss")
    .Recordset("hora_salida").Value = Format(txt04.Value, "HH:mm:ss")
    .Recordset("hora_tope_ingreso").Value = Format(txt05.Value, "HH:mm:ss")
    ctrl1 = DateDiff("n", .Recordset("hora_ingreso").Value, .Recordset("hora_tope_ingreso").Value)
    .Recordset("minutos_tolerancia").Value = ctrl1  'txt02.Text
    
    .Recordset("hora_ingreso2").Value = Format(txt06.Value, "HH:mm:ss")
    .Recordset("hora_salida2").Value = Format(txt07.Value, "HH:mm:ss")
    .Recordset("hora_tope_ingreso2").Value = Format(txt08.Value, "HH:mm:ss")
    ctrl2 = DateDiff("n", .Recordset("hora_ingreso2").Value, .Recordset("hora_tope_ingreso2").Value)
    .Recordset("minutos_tolerancia2").Value = ctrl2
    .Recordset("nro_horas").Value = Dtc_Hrs.Text
    .Recordset("nro_horarios").Value = Txt09.Text
    .Recordset("usr_usuario").Value = glusuario
    .Recordset("fecha_registro").Value = Date   'Format(Date, "dd/mm/aaaa")
    .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
    .Recordset.Update
    '.Recordset.Requery
    'DB.CommitTrans
            
   End With
   
   sw2 = "XX"
   Cmdadicionar.Visible = True
   cmdEditar.Visible = True
   Cmdborrar.Visible = True
   cmdRefresh.Visible = True
   Cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   cmdSalir.Visible = True
   Cmdaceptar.Visible = False
   CmdCancelar.Visible = False
   FraHorario.Enabled = False
   FraHora2.Enabled = True
   
   Adolista.Enabled = True
   grdlista.Enabled = True
   Call ABRIR_TABLA
   TxtGestion.Enabled = True
   Exit Sub

errorAceptar:
   
   Call pErrorRst(db.Errors)
   
   Adolista.Recordset.CancelUpdate
   
   'DB.RollbackTrans
End Sub
 Private Sub Cmdadicionar_Click()
   sw2 = "ADD"
   Cmdadicionar.Visible = False
   cmdEditar.Visible = False
   Cmdborrar.Visible = False
   cmdRefresh.Visible = False
   Cmd_busqueda.Visible = False
   'CmdIMPRIMIR.Visible = False
   cmdSalir.Visible = False
   Cmdaceptar.Visible = True
   CmdCancelar.Visible = True
   FraHorario.Enabled = True
   
   Adolista.Recordset.AddNew
   Adolista.Enabled = False
   grdlista.Enabled = False
   'TxtGestion.Text = Empty
   'Txt01.Text = Empty
'   Txt02.Text = Empty
'   Txt06.Text = Empty
'   txt10.Text = Empty
End Sub

'Private Sub Cmdborrar_Click()
'   Dim Mensaje As String
'
'On Error GoTo errorDelete
'
'   Mensaje = "¿Borrar: " & _
'               Text1.Text & " " & _
'               Trim(Text3.Text) & "?"
'   If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar:") = vbYes Then
'      db.BeginTrans
'      adoLista.Recordset.Delete
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

Private Sub Cmd_Busqueda_Click()
''BUSQUEDA.Visible = True
''fradatos.Enabled = True
' Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = DB
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = grdlista
'    ClBuscaGrid.QueryUtilizado = sql_financiador
'    Set ClBuscaGrid.RecordsetTrabajo = adoLista.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar

End Sub

Private Sub CmdCancelar_Click()
  On Error Resume Next
   sw2 = "XX"
   Cmdadicionar.Visible = True
   cmdEditar.Visible = True
   Cmdborrar.Visible = True
   cmdRefresh.Visible = True
   Cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   cmdSalir.Visible = True
   Cmdaceptar.Visible = False
   CmdCancelar.Visible = False
   FraHorario.Enabled = False
   
   Adolista.Enabled = True
   grdlista.Enabled = True
   Call ABRIR_TABLA
   TxtGestion.Enabled = True
End Sub

Private Sub cmdEditar_Click()
   sw2 = "MOD"
   Cmdadicionar.Visible = False
   cmdEditar.Visible = False
   Cmdborrar.Visible = False
   cmdRefresh.Visible = False
   Cmd_busqueda.Visible = False
   'CmdIMPRIMIR.Visible = False
   cmdSalir.Visible = False
   Cmdaceptar.Visible = True
   CmdCancelar.Visible = True
   FraHorario.Enabled = True
   
   Adolista.Enabled = False
   grdlista.Enabled = False

   TxtGestion.Enabled = False
'   txt01.Enabled = False
   Txt01.SetFocus
End Sub

Private Sub Cmdimprimir_Click()
  Dim iResult As Integer
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\bancos\crybancos.rpt"
     CrystalReport1.WindowShowPrintSetupBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
  'CrystalReport1.ReportFileName = "\SAF-2000\Reportes\RRHH\rr_horario_laboral.rpt"
      CrystalReport1.ReportFileName = App.Path & "\Reportes\RRHH\rr_horario_laboral.rpt"
  iResult = CrystalReport1.PrintReport
  If iResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
CrystalReport1.WindowState = crptMaximized
'REPORGFIN.Show
'   rptModalidadSeleccion.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
  If Adolista.Recordset!turno <> "AM" And Adolista.Recordset!turno <> "PM" Then
    GlHora1 = Adolista.Recordset!tope_hora_ingreso
    cnn.Execute "Update gc_parametros_sistema Set Hora_Ingreso1='" & GlHora1 & "' Where estado_codigo='APR'"
  Else
    If Adolista.Recordset!turno = "AM" Then
        GlHora1 = Adolista.Recordset!tope_hora_ingreso
        cnn.Execute "Update gc_parametros_sistema Set Hora_Ingreso1='" & GlHora1 & "' Where estado_codigo='APR'"
    End If
    If Adolista.Recordset!turno = "PM" Then
        GlHora2 = Adolista.Recordset!tope_hora_ingreso
        cnn.Execute "Update gc_parametros_sistema Set Hora_Ingreso2='" & GlHora2 & "' Where estado_codigo='APR'"
    End If
    
'    Set rs_PARAMETRO2 = New ADODB.Recordset
'    If rs_PARAMETRO2.State = 1 Then rs_PARAMETRO2.Close
'    rs_PARAMETRO2.Open "select * from gc_parametros_sistema where estado_codigo = 'APR' ", cnn, adOpenDynamic, adLockPessimistic
'    If rs_PARAMETRO2.RecordCount > 0 Then
'        'rs_PARAMETRO.MoveFirst
'        rs_PARAMETRO2!Hora_Ingreso1 = GlHora1
'        rs_PARAMETRO2!Hora_Ingreso2 = GlHora2
'    End If
  End If
  Adolista.Recordset!estado_codigo = "SI"
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Dtc_Hrs_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Hrs.BoundText
    Dtc_Par.BoundText = Dtc_Hrs.BoundText
End Sub

Private Sub Dtc_Par_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Par.BoundText
    Dtc_Hrs.BoundText = Dtc_Par.BoundText
End Sub

Private Sub Dtc_ParDes_Click(Area As Integer)
    Dtc_Par.BoundText = Dtc_ParDes.BoundText
    Dtc_Hrs.BoundText = Dtc_ParDes.BoundText
End Sub

Private Sub Form_Load()
   'Dim sql_fuente As String
   sw2 = "XX"
   Cmdadicionar.Visible = True
   cmdEditar.Visible = True
   Cmdborrar.Visible = True
   cmdRefresh.Visible = True
   Cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   cmdSalir.Visible = True
   Cmdaceptar.Visible = False
   CmdCancelar.Visible = False
   FraHorario.Enabled = False
   Adolista.Enabled = True
   grdlista.Enabled = True
   
   Call ABRIR_TABLA
   
   Set rs_turnos = New ADODB.Recordset
   'sql_fuente = "select * from rc_turnos" ' order by fte_codigo"
   rs_turnos.Open "select * from rc_turnos order by turno_codigo", db, adOpenKeyset, adLockOptimistic, adCmdText
   'rs_turnos.Sort = "horario_codigo"
'  ' MsgBox rstfue.RecordCount
   Set AdoTurno.Recordset = rs_turnos
'
   
'   Set rstorg = New ADODB.Recordset
'   sql_financiador = "select * from rc_horarios" 'order by org_codigo"
'   rstorg.Open sql_financiador, DB, adOpenKeyset, adLockOptimistic, adCmdText
''   rstorg.Sort = "Dia_control"
'   Set adoLista.Recordset = rstorg
'   'Set ClBuscaGrid = Nothing
  
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
    Set rstorg = New ADODB.Recordset
    If rstorg.State = 1 Then rstorg.Close
    sql_financiador = "select * from rc_horarios" 'order by org_codigo"
    rstorg.Open sql_financiador, db, adOpenKeyset, adLockOptimistic, adCmdText
    rstorg.Sort = "horario_codigo"
    Set Adolista.Recordset = rstorg
    Set grdlista.DataSource = Adolista.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (rstorg.State = adStateClosed) Then rstorg.Close
   'Set rstorg = Nothing

End Sub

Private Sub txt09_LostFocus()
    If Txt09.Text = "1" Then
        FraHora2.Enabled = False
    End If
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_fc_OrganismoFinanciador 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Financieros - Financiadores"
   ClientHeight    =   6780
   ClientLeft      =   1065
   ClientTop       =   2415
   ClientWidth     =   13080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_fc_OrganismoFinanciador.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   12000
      TabIndex        =   38
      Top             =   120
      Width           =   12060
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1680
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3720
         MaskColor       =   &H00000000&
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6C446
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdAdicionar 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6C650
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6CC74
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdBorrar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6D254
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6DF1E
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdIMPRIMIR 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6E128
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmd_busqueda 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6E6E5
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_fc_OrganismoFinanciador.frx":6EC9D
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FINANCIADOR"
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
         Left            =   8145
         TabIndex        =   47
         Top             =   300
         Width           =   2145
      End
   End
   Begin VB.PictureBox picButtons 
      BackColor       =   &H00C0FFC0&
      Height          =   660
      Left            =   120
      ScaleHeight     =   600
      ScaleWidth      =   10650
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   10710
      Begin VB.CommandButton cmdSalir99 
         Caption         =   "Cerrar"
         Height          =   480
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Salir de Personas"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton CmdIMPRIMIR99 
         Caption         =   "Imprimir"
         Height          =   480
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmd_busqueda99 
         Caption         =   "&Buscar"
         Height          =   480
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdAdicionar99 
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Nuevo Registro"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdEditar99 
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh99 
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Aprueba Registro"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdBorrar99 
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Anula Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdaceptar99 
         Caption         =   "Grabar"
         Height          =   480
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar99 
         Caption         =   "Cancelar"
         Height          =   480
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.PictureBox FRADATOS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5520
      Left            =   5475
      ScaleHeight     =   5460
      ScaleWidth      =   6660
      TabIndex        =   6
      Top             =   1200
      Width           =   6720
      Begin MSDataListLib.DataCombo dtcfue 
         Bindings        =   "frm_fc_OrganismoFinanciador.frx":6EEA7
         DataField       =   "fte_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   4350
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "fte_descripcion"
         BoundColumn     =   "fte_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcPais 
         Bindings        =   "frm_fc_OrganismoFinanciador.frx":6EEBF
         DataField       =   "pais_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   2880
         TabIndex        =   36
         Top             =   3600
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "pais_descripcion"
         BoundColumn     =   "pais_codigo"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Combo4 
         DataField       =   "beneficiario_codigo"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   2040
         TabIndex        =   34
         Text            =   "Representante"
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox Combo3 
         DataField       =   "beneficiario_cargo_representante"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   2040
         TabIndex        =   33
         Text            =   "Representante"
         Top             =   2640
         Width           =   4335
      End
      Begin VB.ComboBox Combo299 
         Height          =   315
         Left            =   4965
         TabIndex        =   4
         Text            =   "N"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtUsuario 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   5160
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.TextBox Combo2 
         BackColor       =   &H80000004&
         DataField       =   "estado_codigo"
         DataSource      =   "adoLista"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtFecha 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         DataField       =   "org_sigla"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "RR.PP."
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "org_descripcion"
         DataSource      =   "adoLista"
         Height          =   495
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "-"
         Top             =   1440
         Width           =   6255
      End
      Begin VB.TextBox Text2 
         DataField       =   "ges_gestion"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Text            =   "2011"
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         DataField       =   "org_codigo"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "111"
         Top             =   720
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtcfu 
         Bindings        =   "frm_fc_OrganismoFinanciador.frx":6EED5
         DataField       =   "fte_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   4350
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "fte_codigo"
         BoundColumn     =   "fte_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo Combo1 
         Bindings        =   "frm_fc_OrganismoFinanciador.frx":6EEED
         DataField       =   "pais_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   2040
         TabIndex        =   35
         Top             =   3600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483633
         ListField       =   "pais_codigo"
         BoundColumn     =   "pais_codigo"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label7 
         Caption         =   "Label3"
         Height          =   225
         Left            =   4440
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label11 
         Caption         =   "Label3"
         Height          =   225
         Left            =   4320
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FUENTE DE FINANCIAMIENTO"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   16
         Top             =   4080
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GESTION:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARGO RESPONSABLE:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   2670
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIGLA:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   2240
         Width           =   510
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "DENOMINACION"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE REGISTRO"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   4860
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "PAIS FINANCIAMIENTO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   3645
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE REPRESENT.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   7
         Top             =   3180
         Width           =   1830
      End
   End
   Begin MSAdodcLib.Adodc adoLista 
      Height          =   330
      Left            =   120
      Top             =   6375
      Width           =   5355
      _ExtentX        =   9446
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
      Bindings        =   "frm_fc_OrganismoFinanciador.frx":6EF03
      Height          =   5115
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   9022
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "Org_codigo"
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
         DataField       =   "Org_descripcion"
         Caption         =   "Descripcion"
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
         DataField       =   "estado_codigo"
         Caption         =   "Estado."
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
         DataField       =   "Org_sigla"
         Caption         =   "Sigla"
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
         DataField       =   "pais_codigo"
         Caption         =   "PAIS"
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
         DataField       =   "org_es_externo"
         Caption         =   "es_externo"
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
         DataField       =   "correlativo_ingreso"
         Caption         =   "Ingreso"
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
      BeginProperty Column08 
         DataField       =   "beneficiario_codigo"
         Caption         =   "Org_representante"
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
         DataField       =   "beneficiario_cargo_representant"
         Caption         =   "Org_cargo"
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
         DataField       =   "fte_codigo"
         Caption         =   "fte_codigo"
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
         DataField       =   "usr_codigo"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3585.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1005.165
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
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adofuente 
      Height          =   375
      Left            =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Adofuente"
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
   Begin MSAdodcLib.Adodc AdoPais 
      Height          =   375
      Left            =   2160
      Top             =   6600
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_fc_OrganismoFinanciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstorg As New ADODB.Recordset
Dim rstfuente As New ADODB.Recordset
Dim rsPais As New ADODB.Recordset
Dim CAMPOS As ADODB.Field
'Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim sql_financiador As String

Private Sub Adolista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     If pRecordset.EOF Or pRecordset.BOF Then
'      cmdEditar.Enabled = False
'      cmdBorrar.Enabled = False
      Text1.Text = Empty
      Text2.Text = Empty
      Text3.Text = Empty
      Text4.Text = Empty
      dtcfu.Text = ""
      dtcfue.Text = ""
      Exit Sub
   End If
   
'   cmdEditar.Enabled = True
'   cmdBorrar.Enabled = True
'
   Select Case pRecordset.EditMode
      Case adEditInProgress
      Case adEditNone
         Text1.Text = IIf(IsNull(pRecordset("org_codigo")), "", pRecordset("org_codigo"))
         Text2.Text = IIf(IsNull(pRecordset("ges_gestion")), "", pRecordset("ges_gestion"))
         Text3.Text = IIf(IsNull(pRecordset("org_descripcion")), "", pRecordset("org_descripcion"))
         Text4.Text = IIf(IsNull(pRecordset("org_sigla")), "", pRecordset("org_sigla"))
         Combo3.Text = IIf(IsNull(pRecordset("beneficiario_cargo_representante")), "", pRecordset("beneficiario_cargo_representante"))
         Combo4.Text = IIf(IsNull(pRecordset("beneficiario_codigo")), "", pRecordset("beneficiario_codigo"))
         Combo1.Text = IIf(IsNull(pRecordset("pais_codigo")), "", pRecordset("pais_codigo"))
         Combo2.Text = IIf(IsNull(pRecordset("estado_codigo")), "", pRecordset("estado_codigo"))
        'If rstfue.State = 1 Then rstfue.Close
         dtcfu.BoundText = pRecordset("fte_codigo")
         rstfuente.MoveFirst
         rstfuente.Find "Fte_codigo= '" & pRecordset!fte_codigo & "'"
         If Not rstfuente.EOF Then dtcfue.Text = rstfuente!fte_descripcion & "" Else dtcfue.Text = ""
         Txtfecha.Text = IIf(IsNull(pRecordset("fecha_registro")), "", pRecordset("fecha_registro"))
         'TxtHora.Text = IIf(IsNull(pRecordset("hora_registro")), "", pRecordset("hora_registro"))
         Txtusuario.Text = IIf(IsNull(pRecordset("usr_codigo")), "", pRecordset("usr_codigo"))
      Case adEditDelete
      Case adEditAdd
   End Select
   adoLista.Caption = CStr(adoLista.Recordset.AbsolutePosition) & " de " & CStr(adoLista.Recordset.RecordCount)
End Sub
   
Private Sub cmdAceptar_Click()
On Error GoTo errorAceptar
Dim SW As Boolean
Dim SQL_FOR As String
Dim RSTORAUX As New ADODB.Recordset

   With adoLista
              If Text1 = "" Then
                    MsgBox "INTRODUZCA DATOS"
                    Text1.SetFocus
                    Exit Sub
               End If
                 If Text2 = "" Then
                    MsgBox "INTRODUZCA DATOS"
                    Text2.SetFocus
                    Exit Sub
                 End If
                If Text3 = "" Then
                    MsgBox "INTRODUZCA DATOS"
                    Text3.SetFocus
                    Exit Sub
                 End If
                   If Text4 = "" Then
                    MsgBox "INTRODUZCA DATOS"
                    Text4.SetFocus
                    Exit Sub
                 End If
                                      
    Set RSTORAUX = New ADODB.Recordset
    SQL_FOR = "select * from Fc_ORGANISMO_FINANCIAMIENTO where ORG_CODIGO = '" & Text1.Text & "'"
    RSTORAUX.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic, adCmdText
    If RSTORAUX.RecordCount > 0 And Text1.Enabled Then
      SW = True
      MsgBox " CODIGO DUPLICADO"
      Text1.SetFocus
      Exit Sub
    End If
    '
    db.BeginTrans
    SW = False
    If Text1.Enabled Then
        .Recordset.AddNew
        .Recordset("org_codigo") = Text1.Text
    End If
            .Recordset("ges_gestion").Value = Text2.Text
            .Recordset("org_descripcion").Value = Text3.Text
            .Recordset("org_sigla").Value = Text4.Text
            .Recordset("beneficiario_cargo_representante").Value = Trim(Combo3.Text)
            .Recordset("beneficiario_codigo").Value = Trim(Combo4.Text)
            .Recordset("pais_codigo").Value = Trim(Combo1.Text)
            .Recordset("estado_codigo").Value = Trim(Combo2.Text)
            .Recordset("fte_codigo").Value = dtcfu.Text
            .Recordset("usr_codigo").Value = frmLogin.txtUserName.Text
            .Recordset("fecha_registro").Value = Date
            .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
            .Recordset.Update
            .Recordset.Requery
            db.CommitTrans
            
      End With
      
   Call Cmdadicionar_Click
    
   Call cmdCancelar_Click
   
   Exit Sub

errorAceptar:
   
   Call pErrorRst(db.Errors)
   
   adoLista.Recordset.CancelUpdate
   
   db.RollbackTrans
End Sub
 Private Sub Cmdadicionar_Click()
   Text1.Enabled = True
   adoLista.Enabled = False
   'grdlista.Enabled = False
   fradatos.Enabled = True
   
  ' cmdBorrar.Visible = False
   cmd_busqueda.Visible = False
   CmdIMPRIMIR.Visible = False
'   cmdSalir.Visible = False
   cmdEditar.Visible = False
  cmdAdicionar.Visible = False

   cmdaceptar.Visible = True
   cmdCancelar.Visible = True
   
   Text1.Text = Empty
   Text2.Text = Empty
   Text3.Text = Empty
   Text4.Text = Empty
   dtcfu.Text = ""
   dtcfue.Text = ""
   Combo3.Text = "" 'Combo3.List(0)
   Combo4.Text = "" 'Combo4.List(0)
   Combo1.Text = "" 'Combo1.List(0)
   Combo2.Text = "" 'Combo2.List(0)
   Text2.SetFocus
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

Private Sub cmdCancelar_Click()
  On Error Resume Next
   Text1.Enabled = True
   fradatos.Enabled = False
    adoLista.Recordset.Requery
   ' Grdlista.ReBind
  ' cmdBorrar.Visible = True
   cmd_busqueda.Visible = True
   CmdIMPRIMIR.Visible = True
'   cmdSalir.Visible = True
   cmdEditar.Visible = True
   cmdAdicionar.Visible = True
   cmdaceptar.Visible = False
   cmdCancelar.Visible = False
   adoLista.Enabled = True
   ' Grdlista.Enabled = True
   adoLista.Recordset.Requery
   'Grdlista.ReBind
'   Unload Me
End Sub

Private Sub cmdEditar_Click()
   If adoLista.Recordset!ESTADO_codigo = "REG" Then
       adoLista.Enabled = False
       ' Grdlista.Enabled = False
       fradatos.Enabled = True
       
       'cmdBorrar.Visible = False
       cmd_busqueda.Visible = False
       CmdIMPRIMIR.Visible = False
    '   cmdSalir.Visible = False
       cmdEditar.Visible = False
      cmdAdicionar.Visible = False
    
       cmdaceptar.Visible = True
       cmdCancelar.Visible = True
    '
       Text1.Enabled = False
       Text2.Enabled = True
       Text3.Enabled = True
       Text4.Enabled = True
       
       Text2.SetFocus
   Else
       MsgBox "No se puede modificar un registro Aprobado ...", , "Atencion"
   End If
End Sub

Private Sub CmdImprimir_Click()
  Dim IResult As Integer
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.ReportFileName = App.Path & "\REPORTES\clasificadores\fr_organismo_financiador.rpt"
  IResult = CrystalReport1.PrintReport
  If IResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  CrystalReport1.WindowState = crptMaximized
  
'  Dim IResult As Integer
'    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\bancos\crybancos.rpt"
'     CrystalReport1.WindowShowPrintSetupBtn = True
'     CrystalReport1.WindowShowRefreshBtn = True
'  CrystalReport1.ReportFileName = "\SAF-2000\Clasificadores\presupuesto\organismo financiador\cryorgfin.rpt"
'  IResult = CrystalReport1.PrintReport
'  If IResult <> 0 Then
'      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
'  End If
'
'CrystalReport1.WindowState = crptMaximized

'REPORGFIN.Show

'   rptModalidadSeleccion.Show vbModal
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub


Private Sub Combo1_Click(Area As Integer)
    DtcPais.BoundText = Combo1.BoundText
End Sub

Private Sub dtcfu_Click(Area As Integer)
dtcfue.BoundText = dtcfu.BoundText
End Sub

Private Sub dtcfue_Click(Area As Integer)
dtcfu.BoundText = dtcfue.BoundText
End Sub

Private Sub DtcPais_Click(Area As Integer)
    Combo1.BoundText = DtcPais.BoundText
End Sub

Private Sub Form_Load()
   
   Dim sql_fuente As String
   Label7.Caption = frmLogin.txtUserName.Text
'   Label9.Caption = Format(Time, "HH:mm:ss")
   Label11.Caption = Date
   
   fradatos.Enabled = False
   cmdBorrar.Visible = True
   cmd_busqueda.Visible = True
   CmdIMPRIMIR.Visible = True
'   cmdSalir.Visible = True
   cmdaceptar.Visible = False
   cmdCancelar.Visible = False
   
   Set rstfuente = New ADODB.Recordset
   sql_fuente = "select * from fc_fuente_financiamiento" ' order by fte_codigo"
   rstfuente.Open sql_fuente, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstfuente.Sort = "FTE_CODIGO"
  ' MsgBox rstfue.RecordCount
   Set Adofuente.Recordset = rstfuente
   
   Set rsPais = New ADODB.Recordset
   'sql_fuente = "select * from gc_pais" ' order by fte_codigo"
   rsPais.Open "select * from gc_pais order by pais_descripcion", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set AdoPais.Recordset = rsPais
   
   Set rstorg = New ADODB.Recordset
   sql_financiador = "select * from fc_organismo_financiamiento" 'order by org_codigo"
   rstorg.Open sql_financiador, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstorg.Sort = "org_codigo"
   Set adoLista.Recordset = rstorg
   'Set ClBuscaGrid = Nothing
  
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (rstorg.State = adStateClosed) Then rstorg.Close
   'Set rstorg = Nothing

End Sub

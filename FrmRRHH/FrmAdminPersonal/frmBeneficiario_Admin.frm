VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBeneficiario_Admin 
   BackColor       =   &H00000000&
   Caption         =   "Control de Personal - File Funcionario"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "frmBeneficiario_Admin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fra_cabecera 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frmBeneficiario_Admin.frx":0A02
      ScaleHeight     =   915
      ScaleWidth      =   6000
      TabIndex        =   193
      Top             =   120
      Width           =   6060
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
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
         Left            =   3720
         Picture         =   "frmBeneficiario_Admin.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
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
         Height          =   720
         Left            =   4440
         Picture         =   "frmBeneficiario_Admin.frx":6CFEC
         Style           =   1  'Graphical
         TabIndex        =   196
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
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
         Left            =   5280
         Picture         =   "frmBeneficiario_Admin.frx":6D5A9
         Style           =   1  'Graphical
         TabIndex        =   195
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   405
         Left            =   120
         TabIndex        =   198
         Top             =   240
         Width           =   2835
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
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
         Left            =   9900
         TabIndex        =   194
         Top             =   300
         Width           =   1305
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   0
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
      TabPicture(0)   =   "frmBeneficiario_Admin.frx":6D7B3
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label40"
      Tab(0).Control(1)=   "fraDatos"
      Tab(0).Control(2)=   "FraGrabarCancelar"
      Tab(0).Control(3)=   "fraOpciones"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "SEG. SOCIAL y AFP"
      TabPicture(1)   =   "frmBeneficiario_Admin.frx":6D7CF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame18"
      Tab(1).Control(1)=   "Frame19"
      Tab(1).Control(2)=   "Label44"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "CURRICULARES"
      TabPicture(2)   =   "frmBeneficiario_Admin.frx":6D7EB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame15"
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(2)=   "Label45"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "REL. CONTRACTUAL"
      TabPicture(3)   =   "frmBeneficiario_Admin.frx":6D807
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label46"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame16"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame17"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H00404040&
         Height          =   735
         Left            =   -75000
         Picture         =   "frmBeneficiario_Admin.frx":6D823
         ScaleHeight     =   675
         ScaleWidth      =   9195
         TabIndex        =   208
         Top             =   840
         Width           =   9255
         Begin VB.CommandButton CmdDesapr 
            BackColor       =   &H0080C0FF&
            Caption         =   "Desapr"
            Height          =   600
            Left            =   6600
            Picture         =   "frmBeneficiario_Admin.frx":D9855
            Style           =   1  'Graphical
            TabIndex        =   215
            ToolTipText     =   "Aprueba Registro"
            Top             =   60
            Width           =   740
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
            Left            =   4020
            Picture         =   "frmBeneficiario_Admin.frx":D9A5F
            Style           =   1  'Graphical
            TabIndex        =   214
            ToolTipText     =   "Aprueba Registro"
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton BtnAñadir 
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
            Left            =   120
            Picture         =   "frmBeneficiario_Admin.frx":DA295
            Style           =   1  'Graphical
            TabIndex        =   213
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton BtnModificar 
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
            Left            =   1380
            Picture         =   "frmBeneficiario_Admin.frx":DAA54
            Style           =   1  'Graphical
            TabIndex        =   212
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
            Left            =   2760
            Picture         =   "frmBeneficiario_Admin.frx":DB369
            Style           =   1  'Graphical
            TabIndex        =   211
            ToolTipText     =   "Anula Registro Activo"
            Top             =   30
            Width           =   1245
         End
         Begin VB.CommandButton CmdVerDisco 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Docs."
            Enabled         =   0   'False
            Height          =   600
            Left            =   7440
            Picture         =   "frmBeneficiario_Admin.frx":DBAB5
            Style           =   1  'Graphical
            TabIndex        =   210
            Top             =   60
            Width           =   740
         End
         Begin VB.CommandButton CmdFoto 
            BackColor       =   &H80000015&
            Caption         =   "&Foto"
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
            Left            =   5400
            Picture         =   "frmBeneficiario_Admin.frx":DBE3D
            Style           =   1  'Graphical
            TabIndex        =   209
            ToolTipText     =   "Carga Foto de la Persona"
            Top             =   30
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.PictureBox FraGrabarCancelar 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   -75000
         Picture         =   "frmBeneficiario_Admin.frx":DC3C7
         ScaleHeight     =   675
         ScaleWidth      =   9195
         TabIndex        =   204
         Top             =   840
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Admin.frx":1483F9
            Style           =   1  'Graphical
            TabIndex        =   206
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
            Picture         =   "frmBeneficiario_Admin.frx":148BCF
            Style           =   1  'Graphical
            TabIndex        =   205
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
            TabIndex        =   207
            Top             =   300
            Visible         =   0   'False
            Width           =   525
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   4335
         Left            =   0
         TabIndex        =   175
         Top             =   5040
         Width           =   9255
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
            Height          =   480
            Left            =   1800
            Picture         =   "frmBeneficiario_Admin.frx":1494BB
            Style           =   1  'Graphical
            TabIndex        =   180
            ToolTipText     =   "Anula Registro Activo"
            Top             =   260
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
            Height          =   480
            Left            =   2640
            Picture         =   "frmBeneficiario_Admin.frx":149EBD
            Style           =   1  'Graphical
            TabIndex        =   179
            ToolTipText     =   "Aprueba Registro"
            Top             =   260
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
            Height          =   480
            Left            =   960
            Picture         =   "frmBeneficiario_Admin.frx":14A447
            Style           =   1  'Graphical
            TabIndex        =   178
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   260
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
            Height          =   480
            Left            =   120
            Picture         =   "frmBeneficiario_Admin.frx":14A9D1
            Style           =   1  'Graphical
            TabIndex        =   177
            ToolTipText     =   "Nuevo Registro"
            Top             =   260
            Width           =   855
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00000000&
            Height          =   660
            Left            =   8360
            TabIndex        =   176
            Top             =   120
            Width           =   615
            Begin VB.Image ImgFiniquito 
               Height          =   540
               Left            =   20
               Picture         =   "frmBeneficiario_Admin.frx":14AF5B
               Top             =   80
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtgLiquidacion 
            Bindings        =   "frmBeneficiario_Admin.frx":14B2E3
            Height          =   2985
            Left            =   120
            TabIndex        =   181
            Top             =   840
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   5265
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
            ColumnCount     =   9
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
            BeginProperty Column03 
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
            BeginProperty Column04 
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
            BeginProperty Column05 
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
            BeginProperty Column06 
               DataField       =   "estado_registro"
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
               DataField       =   "codigo_beneficiario"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1379.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   569.764
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column08 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoLiquidacion 
            Height          =   375
            Left            =   120
            Top             =   3840
            Width           =   9015
            _ExtentX        =   15901
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
            Appearance      =   0
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
         Begin VB.Label Label35 
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
            ForeColor       =   &H00FFFF80&
            Height          =   195
            Left            =   6600
            TabIndex        =   182
            Top             =   240
            Width           =   1605
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   4215
         Left            =   0
         TabIndex        =   167
         Top             =   840
         Width           =   9255
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
            Height          =   480
            Left            =   1800
            Picture         =   "frmBeneficiario_Admin.frx":14B300
            Style           =   1  'Graphical
            TabIndex        =   172
            ToolTipText     =   "Anula Registro Activo"
            Top             =   260
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
            Height          =   480
            Left            =   2640
            Picture         =   "frmBeneficiario_Admin.frx":14BD02
            Style           =   1  'Graphical
            TabIndex        =   171
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   260
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
            Height          =   480
            Left            =   960
            Picture         =   "frmBeneficiario_Admin.frx":14C28C
            Style           =   1  'Graphical
            TabIndex        =   170
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   260
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
            Height          =   480
            Left            =   120
            Picture         =   "frmBeneficiario_Admin.frx":14C816
            Style           =   1  'Graphical
            TabIndex        =   169
            ToolTipText     =   "Nuevo Registro"
            Top             =   260
            Width           =   855
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00000000&
            Height          =   660
            Left            =   8360
            TabIndex        =   168
            Top             =   120
            Width           =   615
            Begin VB.Image Img_CTO 
               Height          =   540
               Left            =   20
               Picture         =   "frmBeneficiario_Admin.frx":14CDA0
               Top             =   80
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtG_Contrato 
            Bindings        =   "frmBeneficiario_Admin.frx":14D128
            Height          =   2985
            Left            =   165
            TabIndex        =   173
            Top             =   840
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   5265
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
            Top             =   3840
            Width           =   8970
            _ExtentX        =   15822
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
            Appearance      =   0
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
         Begin VB.Label Label34 
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
            ForeColor       =   &H00FFFF80&
            Height          =   195
            Left            =   6900
            TabIndex        =   174
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   4335
         Left            =   -75000
         TabIndex        =   159
         Top             =   5040
         Width           =   9255
         Begin VB.CommandButton CmdApr3 
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
            Picture         =   "frmBeneficiario_Admin.frx":14D143
            Style           =   1  'Graphical
            TabIndex        =   164
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdElim3 
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
            Picture         =   "frmBeneficiario_Admin.frx":14D6CD
            Style           =   1  'Graphical
            TabIndex        =   163
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdAdd3 
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
            Picture         =   "frmBeneficiario_Admin.frx":14E0CF
            Style           =   1  'Graphical
            TabIndex        =   162
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdMod3 
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
            Picture         =   "frmBeneficiario_Admin.frx":14E659
            Style           =   1  'Graphical
            TabIndex        =   161
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00000000&
            Height          =   690
            Left            =   8360
            TabIndex        =   160
            Top             =   120
            Width           =   615
            Begin VB.Image Img_DocRespaldo 
               Height          =   540
               Left            =   20
               Picture         =   "frmBeneficiario_Admin.frx":14EBE3
               Top             =   100
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtgLaborales 
            Bindings        =   "frmBeneficiario_Admin.frx":14EF6B
            Height          =   3105
            Left            =   165
            TabIndex        =   165
            Top             =   840
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   5477
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
            Top             =   3960
            Width           =   8970
            _ExtentX        =   15822
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
            Appearance      =   0
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
            ForeColor       =   &H00FFFF80&
            Height          =   195
            Left            =   6795
            TabIndex        =   166
            Top             =   540
            Width           =   1425
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   4215
         Left            =   -75000
         TabIndex        =   151
         Top             =   840
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Admin.frx":14EF87
            Style           =   1  'Graphical
            TabIndex        =   156
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdElim2 
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
            Picture         =   "frmBeneficiario_Admin.frx":14F511
            Style           =   1  'Graphical
            TabIndex        =   155
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
            Picture         =   "frmBeneficiario_Admin.frx":14FF13
            Style           =   1  'Graphical
            TabIndex        =   154
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
            Picture         =   "frmBeneficiario_Admin.frx":15049D
            Style           =   1  'Graphical
            TabIndex        =   153
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00000000&
            Height          =   690
            Left            =   8360
            TabIndex        =   152
            Top             =   120
            Width           =   615
            Begin VB.Image Img_CV 
               Height          =   540
               Left            =   20
               Picture         =   "frmBeneficiario_Admin.frx":150A27
               Top             =   100
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtgEducacionales 
            Bindings        =   "frmBeneficiario_Admin.frx":150DAF
            Height          =   2985
            Left            =   165
            TabIndex        =   157
            Top             =   840
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   5265
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
            Top             =   3840
            Width           =   8970
            _ExtentX        =   15822
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
            Appearance      =   0
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
         Begin VB.Label LblCV 
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
            ForeColor       =   &H00FFFF80&
            Height          =   195
            Left            =   6840
            TabIndex        =   158
            Top             =   560
            Width           =   1425
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00000000&
         Caption         =   "DEPENDIENTES DEL FUNCIONARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   3735
         Left            =   -75000
         TabIndex        =   145
         Top             =   720
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Admin.frx":150DCF
            Style           =   1  'Graphical
            TabIndex        =   149
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
            Picture         =   "frmBeneficiario_Admin.frx":151359
            Style           =   1  'Graphical
            TabIndex        =   148
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
            Picture         =   "frmBeneficiario_Admin.frx":151D5B
            Style           =   1  'Graphical
            TabIndex        =   147
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
            Picture         =   "frmBeneficiario_Admin.frx":1522E5
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DtgDependiente 
            Bindings        =   "frmBeneficiario_Admin.frx":15286F
            Height          =   2625
            Left            =   120
            TabIndex        =   150
            Top             =   720
            Width           =   9090
            _ExtentX        =   16034
            _ExtentY        =   4630
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
            Caption         =   "DEPENDIENTES (Hijos y Parientes)"
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "Cod_asegurado"
               Caption         =   "Cod.Asegurado"
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
               DataField       =   "denominacion_beneficiario"
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
            BeginProperty Column02 
               DataField       =   "fecha_nacimiento"
               Caption         =   "Fecha.Nacimiento"
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
               DataField       =   "Fecha_asegurado"
               Caption         =   "Fecha.Asegurado"
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
               DataField       =   "pariente_descripcion"
               Caption         =   "Parentesco"
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
            BeginProperty Column06 
               DataField       =   "ocupacion_pariente"
               Caption         =   "Ocupacion"
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
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1379.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1379.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column06 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoDependiente 
            Height          =   330
            Left            =   120
            Top             =   3360
            Width           =   9090
            _ExtentX        =   16034
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
            Caption         =   " <--- Dependientes --->"
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
      Begin VB.Frame Frame19 
         BackColor       =   &H00000000&
         Height          =   4935
         Left            =   -75000
         TabIndex        =   107
         Top             =   4440
         Width           =   9255
         Begin VB.Frame FraBco 
            BackColor       =   &H00000000&
            Caption         =   "CUENTA BANCARIA PERSONAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   1600
            Left            =   135
            TabIndex        =   117
            Top             =   3255
            Width           =   9060
            Begin VB.ComboBox DtcCtaTip 
               DataSource      =   "adoLista"
               Height          =   315
               ItemData        =   "frmBeneficiario_Admin.frx":15288C
               Left            =   6240
               List            =   "frmBeneficiario_Admin.frx":152896
               TabIndex        =   188
               Text            =   "CUENTA CORRIENTE"
               Top             =   480
               Width           =   2660
            End
            Begin VB.TextBox DtcCtaNom 
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               DataField       =   "beneficiario_denominacion"
               DataSource      =   "adoLista"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3360
               MaxLength       =   15
               TabIndex        =   187
               Top             =   1140
               Width           =   5540
            End
            Begin VB.TextBox DtcCta 
               BackColor       =   &H00FFFFFF&
               DataField       =   "cta_codigo"
               DataSource      =   "adoLista"
               Height          =   285
               Left            =   240
               MaxLength       =   15
               TabIndex        =   186
               Top             =   1140
               Width           =   3045
            End
            Begin MSDataListLib.DataCombo DtcBanco 
               Bindings        =   "frmBeneficiario_Admin.frx":1528BC
               DataField       =   "bco_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   3120
               TabIndex        =   14
               Top             =   240
               Visible         =   0   'False
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "bco_codigo"
               BoundColumn     =   "bco_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcBancoDes 
               Bindings        =   "frmBeneficiario_Admin.frx":1528D1
               DataField       =   "bco_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   240
               TabIndex        =   118
               Top             =   480
               Width           =   5895
               _ExtentX        =   10398
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "bco_descripcion"
               BoundColumn     =   "bco_codigo"
               Text            =   ""
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Cuenta Bancaria                                     Denominacion de la Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   240
               TabIndex        =   185
               Top             =   885
               Width           =   5625
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Entidad Financiera                                                                                                 Tipo de Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   240
               TabIndex        =   119
               Top             =   225
               Width           =   7425
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00404040&
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
            ForeColor       =   &H00FFFF80&
            Height          =   1515
            Left            =   135
            TabIndex        =   112
            Top             =   1695
            Width           =   9060
            Begin VB.TextBox Text7 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   280
               Left            =   8625
               TabIndex        =   121
               Top             =   1090
               Width           =   240
            End
            Begin VB.TextBox txt_afp 
               BackColor       =   &H00FFFFFF&
               DataField       =   "asegurado_codigo_afp"
               DataSource      =   "adoLista"
               Height          =   285
               Left            =   240
               MaxLength       =   15
               TabIndex        =   11
               Top             =   480
               Width           =   2205
            End
            Begin MSDataListLib.DataCombo dtc_afp_des 
               Bindings        =   "frmBeneficiario_Admin.frx":1528E6
               DataField       =   "beneficiario_codigo_afp"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   2640
               TabIndex        =   12
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
               Bindings        =   "frmBeneficiario_Admin.frx":152902
               DataField       =   "beneficiario_codigo_afp"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   7560
               TabIndex        =   113
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
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   240
               TabIndex        =   13
               Top             =   1080
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   90439681
               CurrentDate     =   40179
               MinDate         =   2
            End
            Begin MSDataListLib.DataCombo dtc_afp_dir 
               Bindings        =   "frmBeneficiario_Admin.frx":15291E
               DataField       =   "beneficiario_codigo_afp"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   1920
               TabIndex        =   114
               Top             =   1080
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   0
               ForeColor       =   16777215
               ListField       =   "beneficiario_domicilio_legal"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   ""
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H00404040&
               Caption         =   "Número Reg.APF (NUA)     Nombre de Entidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   240
               TabIndex        =   116
               Top             =   240
               Width           =   4170
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackColor       =   &H00404040&
               Caption         =   "Fecha Reg.AFP      Direccion Entidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   240
               TabIndex        =   115
               Top             =   840
               Width           =   3300
            End
         End
         Begin VB.Frame FraSS 
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFF80&
            Height          =   1515
            Left            =   120
            TabIndex        =   108
            Top             =   120
            Width           =   9060
            Begin VB.TextBox Text10 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   255
               Left            =   8400
               TabIndex        =   137
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txt_ss 
               BackColor       =   &H00FFFFFF&
               DataField       =   "asegurado_codigo_caja"
               DataSource      =   "adoLista"
               Height          =   285
               Left            =   240
               MaxLength       =   15
               TabIndex        =   7
               Top             =   465
               Width           =   2805
            End
            Begin MSDataListLib.DataCombo DtcSSEnt 
               Bindings        =   "frmBeneficiario_Admin.frx":15293A
               DataField       =   "beneficiario_codigo_seguro"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   240
               TabIndex        =   10
               Top             =   1080
               Width           =   4335
               _ExtentX        =   7646
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcSS 
               Bindings        =   "frmBeneficiario_Admin.frx":152959
               DataField       =   "beneficiario_codigo_seguro"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   3360
               TabIndex        =   109
               Top             =   720
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
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   4080
               TabIndex        =   8
               Top             =   465
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   90439681
               CurrentDate     =   40179
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker DTP_FechaSSExp 
               DataField       =   "fecha_asegurado_fin_caja"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   7155
               TabIndex        =   9
               Top             =   465
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   90439681
               CurrentDate     =   40179
               MinDate         =   2
            End
            Begin MSDataListLib.DataCombo DtcSSDir 
               Bindings        =   "frmBeneficiario_Admin.frx":152978
               DataField       =   "beneficiario_codigo_seguro"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   4560
               TabIndex        =   110
               Top             =   1080
               Width           =   4335
               _ExtentX        =   7646
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   4210752
               ForeColor       =   16777215
               ListField       =   "beneficiario_domicilio_legal"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   ""
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Matrícula del Asegurado                                     Fecha Asegurado                                Fecha Expiracion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   240
               TabIndex        =   184
               Top             =   220
               Width           =   8475
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Entidad Aseguradora                                                       Direccion Entidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   240
               TabIndex        =   111
               Top             =   795
               Width           =   5985
            End
         End
      End
      Begin VB.PictureBox fraDatos 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   7815
         Left            =   -75000
         ScaleHeight     =   7755
         ScaleWidth      =   9195
         TabIndex        =   20
         Top             =   1560
         Width           =   9255
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
            DataSource      =   "adoLista"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   223
            Top             =   4080
            Width           =   1440
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
            DataSource      =   "adoLista"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            MaxLength       =   20
            TabIndex        =   222
            Top             =   4080
            Width           =   1440
         End
         Begin VB.TextBox TxtRenca 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            DataSource      =   "adoLista"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   221
            Text            =   "0"
            Top             =   4080
            Width           =   1560
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   8520
            TabIndex        =   220
            Top             =   3375
            Width           =   375
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   4190
            TabIndex        =   203
            Top             =   2175
            Width           =   255
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   8620
            TabIndex        =   136
            Top             =   2175
            Width           =   255
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   7425
            TabIndex        =   126
            Top             =   2655
            Width           =   255
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00000000&
            Caption         =   "Aprobado"
            ForeColor       =   &H00FFFFFF&
            Height          =   600
            Left            =   6060
            TabIndex        =   27
            Top             =   120
            Width           =   980
            Begin VB.Label lblActivo 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "NO"
               DataField       =   "estado_registro"
               DataSource      =   "adoLista"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   300
               Left            =   120
               TabIndex        =   32
               Top             =   200
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00000000&
            ForeColor       =   &H00000040&
            Height          =   1800
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   9150
            Begin VB.PictureBox Img_Foto 
               AutoRedraw      =   -1  'True
               Height          =   1635
               Left            =   7215
               ScaleHeight     =   1575
               ScaleWidth      =   1815
               TabIndex        =   120
               Top             =   120
               Width           =   1875
               Begin VB.Image Image2 
                  Height          =   1575
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.Label TxtNIT 
               BackColor       =   &H00404040&
               Caption         =   "-"
               DataField       =   "beneficiario_nit"
               DataSource      =   "adoLista"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6045
               TabIndex        =   134
               Top             =   1020
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label DtcDepto3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "depto_sigla"
               DataSource      =   "adoLista"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4875
               TabIndex        =   133
               Top             =   1260
               Width           =   975
            End
            Begin VB.Label Dtc_doc_id 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "tipodoc_codigo"
               DataSource      =   "adoLista"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3105
               TabIndex        =   132
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label txtDenominacion 
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "beneficiario_denominacion"
               DataSource      =   "adoLista"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   130
               Top             =   600
               Width           =   5895
            End
            Begin VB.Label txtCodigo 
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "beneficiario_codigo"
               DataSource      =   "adoLista"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   122
               Top             =   1260
               Width           =   2055
            End
            Begin VB.Label LblInicial 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Label34"
               DataField       =   "ARCHIVO_FOTO"
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
               ForeColor       =   &H00FFFF80&
               Height          =   255
               Left            =   4965
               TabIndex        =   99
               Top             =   1425
               Width           =   2055
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Nombres y Apellidos del Funcionario"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   120
               TabIndex        =   30
               Top             =   330
               Width           =   3315
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Documento de Identidad                Tipo Doc.                  Expedido.en "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   120
               TabIndex        =   29
               Top             =   990
               Width           =   5805
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFF80&
            Height          =   1335
            Left            =   120
            TabIndex        =   21
            Top             =   5685
            Width           =   8835
            Begin MSDataListLib.DataCombo Dtc_prov_cod 
               Bindings        =   "frmBeneficiario_Admin.frx":152997
               DataField       =   "prov_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   2880
               TabIndex        =   22
               Top             =   600
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
            Begin MSDataListLib.DataCombo Dtc_munic_cod 
               Bindings        =   "frmBeneficiario_Admin.frx":1529AE
               DataField       =   "munic_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   7200
               TabIndex        =   23
               Top             =   555
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
            Begin MSDataListLib.DataCombo Dtc_prov 
               Bindings        =   "frmBeneficiario_Admin.frx":1529C5
               DataField       =   "prov_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   1020
               TabIndex        =   5
               Top             =   885
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "prov_descripcion"
               BoundColumn     =   "prov_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_munic 
               Bindings        =   "frmBeneficiario_Admin.frx":1529DC
               DataField       =   "munic_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   5280
               TabIndex        =   6
               Top             =   885
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "munic_descripcion"
               BoundColumn     =   "munic_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_depto_cod 
               Bindings        =   "frmBeneficiario_Admin.frx":1529F3
               DataField       =   "depto_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   7200
               TabIndex        =   24
               Top             =   120
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
            Begin MSDataListLib.DataCombo Dtc_depto 
               Bindings        =   "frmBeneficiario_Admin.frx":152A0B
               DataField       =   "depto_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   5280
               TabIndex        =   4
               Top             =   345
               Width           =   3375
               _ExtentX        =   5953
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
               Bindings        =   "frmBeneficiario_Admin.frx":152A23
               DataField       =   "pais_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   1400
               TabIndex        =   3
               Top             =   360
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "pais_descripcion"
               BoundColumn     =   "pais_codigo"
               Text            =   "DataCombo5"
            End
            Begin MSDataListLib.DataCombo DtcPaisCod 
               Bindings        =   "frmBeneficiario_Admin.frx":152A39
               DataField       =   "pais_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   3000
               TabIndex        =   138
               Top             =   120
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
               Bindings        =   "frmBeneficiario_Admin.frx":152A4F
               DataField       =   "pais_codigo"
               DataSource      =   "adoLista"
               Height          =   315
               Left            =   2040
               TabIndex        =   139
               Top             =   120
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
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Nacionalidad                                                                         Depto."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   3
               Left            =   120
               TabIndex        =   26
               Top             =   375
               Width           =   5100
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Provincia                                                                                 Municipio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   12
               Left            =   120
               TabIndex        =   25
               Top             =   915
               Width           =   5340
            End
         End
         Begin MSComCtl2.DTPicker DTP_FechaNac 
            DataField       =   "beneficiario_fecha_nacimiento"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   3600
            TabIndex        =   0
            Top             =   5240
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90439681
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo TxtProfesion 
            Bindings        =   "frmBeneficiario_Admin.frx":152A65
            DataField       =   "ocup_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   2160
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "ocup_descripcion"
            BoundColumn     =   "ocup_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_Ocup 
            Bindings        =   "frmBeneficiario_Admin.frx":152A81
            DataField       =   "ocup_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   3720
            TabIndex        =   100
            Top             =   1800
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "ocup_codigo"
            BoundColumn     =   "ocup_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "frmBeneficiario_Admin.frx":152A9D
            DataField       =   "planilla_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Visible         =   0   'False
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "planilla_descripcion"
            BoundColumn     =   "planilla_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo TxtTipo 
            Bindings        =   "frmBeneficiario_Admin.frx":152AB6
            DataField       =   "planilla_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   3720
            TabIndex        =   105
            Top             =   900
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "planilla_codigo"
            BoundColumn     =   "planilla_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcEstCivDes 
            Bindings        =   "frmBeneficiario_Admin.frx":152ACF
            DataField       =   "estado_civil_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   5240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "estado_civil_descripcion"
            BoundColumn     =   "estado_civil_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo DtcEstCiv 
            Bindings        =   "frmBeneficiario_Admin.frx":152AE9
            DataField       =   "estado_civil_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2520
            TabIndex        =   106
            Top             =   5220
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "estado_civil_codigo"
            BoundColumn     =   "estado_civil_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "frmBeneficiario_Admin.frx":152B03
            DataField       =   "puesto_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   8160
            TabIndex        =   123
            Top             =   1800
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "puesto_codigo"
            BoundColumn     =   "puesto_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "frmBeneficiario_Admin.frx":152B1C
            DataField       =   "unidad_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   127
            Top             =   2640
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "frmBeneficiario_Admin.frx":152B35
            DataField       =   "puesto_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4560
            TabIndex        =   128
            Top             =   2160
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "puesto_descripcion"
            BoundColumn     =   "puesto_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "frmBeneficiario_Admin.frx":152B4E
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   120
            TabIndex        =   129
            Top             =   3360
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "unidad_descripcion_pla"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "frmBeneficiario_Admin.frx":152B67
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   3840
            TabIndex        =   131
            Top             =   3120
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "unidad_codigo_pla"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "frmBeneficiario_Admin.frx":152B80
            DataField       =   "unidad_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   7800
            TabIndex        =   135
            Top             =   2700
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
            Bindings        =   "frmBeneficiario_Admin.frx":152B99
            DataField       =   "genero_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   6660
            TabIndex        =   2
            Top             =   5240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "genero_descripcion"
            BoundColumn     =   "genero_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "frmBeneficiario_Admin.frx":152BB2
            DataField       =   "genero_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5760
            TabIndex        =   141
            Top             =   5280
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
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
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5760
            TabIndex        =   192
            Top             =   4080
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90439681
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "frmBeneficiario_Admin.frx":152BCB
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   840
            TabIndex        =   216
            Top             =   0
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "planilla_descripcion"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "frmBeneficiario_Admin.frx":152BE4
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5520
            TabIndex        =   217
            Top             =   0
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "planilla_codigo"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmBeneficiario_Admin.frx":152BFE
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   218
            Top             =   3360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "planilla_descripcion"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frmBeneficiario_Admin.frx":152C17
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   7920
            TabIndex        =   219
            Top             =   3120
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "planilla_codigo"
            BoundColumn     =   "unidad_codigo_pla"
            Text            =   ""
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Profesion u Ocupacion Principal                                   Puesto Actual del Funcionario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   202
            Top             =   1905
            Width           =   7080
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            DataSource      =   "adoLista"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6720
            TabIndex        =   191
            Top             =   4540
            Width           =   2130
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            DataSource      =   "adoLista"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2520
            TabIndex        =   190
            Top             =   4540
            Width           =   2130
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Teléfono Celular Personal                                                    Teléfono Corporativo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   189
            Top             =   4560
            Width           =   6585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Unidad Organizacional"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   183
            Top             =   2625
            Width           =   2055
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
            DataSource      =   "adoLista"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6960
            TabIndex        =   144
            Top             =   -360
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.Label txt_item 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            DataSource      =   "adoLista"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7755
            TabIndex        =   143
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label TxtDireccion 
            BackColor       =   &H00404040&
            Caption         =   "-"
            DataField       =   "beneficiario_domicilio_legal"
            DataSource      =   "adoLista"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1680
            TabIndex        =   142
            Top             =   7200
            Width           =   7215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Estado Civil                                                        Fecha Nacimiento                            Género"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   140
            Top             =   4950
            Width           =   7155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Planilla a la que corresponde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   125
            Top             =   3105
            Width           =   2625
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Sueldo Básico           Refrigerio/Otro          Bono Antigüedad           Fecha de Ingreso      Correl.Planilla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   124
            Top             =   3825
            Width           =   8730
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Domicilio Actual:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   31
            Top             =   7245
            Width           =   1485
         End
      End
      Begin VB.Label Label46 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IV. SITUACION DENTRO DE LA INSTITUCION "
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
         Left            =   0
         TabIndex        =   104
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label Label45 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " III. ANTECEDENTES PROFESIONALES y DE TRABAJO "
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
         Left            =   -75000
         TabIndex        =   103
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label Label44 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " II. DEPENDIENTES,  SEGURO SOCIAL Y AFP "
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
         Height          =   375
         Left            =   -75000
         TabIndex        =   102
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label Label40 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " I. DATOS PERSONALES GENERALES"
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
         Left            =   -75000
         TabIndex        =   101
         Top             =   360
         Width           =   9255
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   8325
      Left            =   120
      ScaleHeight     =   8265
      ScaleWidth      =   5985
      TabIndex        =   17
      Top             =   1200
      Width           =   6045
      Begin MSDataGridLib.DataGrid Grdlista 
         Bindings        =   "frmBeneficiario_Admin.frx":152C31
         Height          =   7335
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   12938
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
         Caption         =   "LISTADOS"
         ColumnCount     =   6
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "estado_codigo"
            Caption         =   "Aprob."
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3660.095
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoLista 
         Height          =   330
         Left            =   0
         Top             =   7860
         Width           =   5985
         _ExtentX        =   10557
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
      Begin MSDataListLib.DataCombo dtc_buscar_desc 
         Bindings        =   "frmBeneficiario_Admin.frx":152C48
         Height          =   315
         Left            =   720
         TabIndex        =   199
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
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
         Bindings        =   "frmBeneficiario_Admin.frx":152C65
         DataField       =   "beneficiario_codigo"
         Height          =   315
         Left            =   4200
         TabIndex        =   200
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label52 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar..."
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   0
         TabIndex        =   201
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc AdoTip_ben 
      Height          =   330
      Left            =   10800
      Top             =   10080
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
      Left            =   4320
      Top             =   10440
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Left            =   0
      Top             =   10440
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin MSAdodcLib.Adodc Ado_Datos2 
      Height          =   330
      Left            =   2160
      Top             =   10440
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_Datos2"
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
      Left            =   12960
      Top             =   10080
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   6480
      Top             =   10440
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_datos3"
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   10800
      Top             =   10440
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   15000
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
   Begin VB.PictureBox fraDatos2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   7455
      Left            =   6240
      ScaleHeight     =   7395
      ScaleWidth      =   8685
      TabIndex        =   33
      Top             =   2160
      Visible         =   0   'False
      Width           =   8745
      Begin VB.TextBox txtCodigo2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "codigo_beneficiario"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   81
         Top             =   740
         Width           =   2205
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lugar (DOMICILIO PRINCIPAL) ------------------------ Lugar (DOMICILIO LEGAL)"
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
         Height          =   1720
         Left            =   120
         TabIndex        =   56
         Top             =   4485
         Width           =   8475
         Begin MSDataListLib.DataCombo Dtc_prov_cod22 
            DataField       =   "prov_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   57
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
         Begin MSDataListLib.DataCombo Dtc_munic_cod22 
            DataField       =   "munic_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   58
            Top             =   795
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
         Begin MSDataListLib.DataCombo Dtc_local_cod22 
            DataField       =   "comun_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   59
            Top             =   1140
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "comun_codigo"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_depto_cod22 
            DataField       =   "depto_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   60
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   741
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo Dtc_prov_cod02 
            DataField       =   "prov_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   61
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
         Begin MSDataListLib.DataCombo Dtc_munic_cod02 
            DataField       =   "munic_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   62
            Top             =   795
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
         Begin MSDataListLib.DataCombo Dtc_local_cod02 
            DataField       =   "comun_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   63
            Top             =   1140
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "comun_codigo"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_prov02 
            DataField       =   "prov_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   64
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_munic02 
            DataField       =   "munic_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   65
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_local02 
            DataField       =   "comun_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   66
            Top             =   1320
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "comun_descripcion"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_depto_cod02 
            DataField       =   "depto_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   67
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   741
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo Dtc_depto02 
            DataField       =   "depto_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   68
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   741
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo Dtc_depto22 
            DataField       =   "depto_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   69
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   741
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo Dtc_prov22 
            DataField       =   "prov_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   70
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_munic22 
            DataField       =   "munic_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   71
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_local22 
            DataField       =   "comun_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   72
            Top             =   1320
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "comun_descripcion"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Depto."
            Height          =   255
            Index           =   21
            Left            =   4320
            TabIndex        =   80
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Comuni."
            Height          =   255
            Index           =   20
            Left            =   4320
            TabIndex        =   79
            Top             =   1360
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Munic."
            Height          =   255
            Index           =   19
            Left            =   4320
            TabIndex        =   78
            Top             =   1000
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prov."
            Height          =   255
            Index           =   18
            Left            =   4320
            TabIndex        =   77
            Top             =   640
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prov."
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   76
            Top             =   640
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Munic."
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   75
            Top             =   1000
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Comuni."
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   74
            Top             =   1360
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Depto."
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   73
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtDireccion22 
         DataField       =   "domicilio_legal"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   1440
         TabIndex        =   55
         Top             =   4080
         Width           =   7155
      End
      Begin VB.TextBox TxtZona2 
         DataField       =   "zona"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   5280
         TabIndex        =   54
         Text            =   "-"
         Top             =   3240
         Width           =   3300
      End
      Begin VB.TextBox TxtProfesion2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "profesion"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   3000
         TabIndex        =   53
         Top             =   2640
         Width           =   5595
      End
      Begin VB.TextBox TxtCargo2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "cargo"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   2040
         TabIndex        =   52
         Top             =   1920
         Width           =   6540
      End
      Begin VB.TextBox TxtRenca2 
         DataField       =   "Reg_Profesional"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   51
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Persona Juridica"
         Height          =   675
         Left            =   120
         TabIndex        =   48
         Top             =   0
         Width           =   4635
         Begin MSDataListLib.DataCombo TDBtipoben2 
            DataField       =   "tipo_beneficiario"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   225
            TabIndex        =   49
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "descripcion"
            BoundColumn     =   "tipo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo TxtTipo2 
            DataField       =   "tipo_beneficiario"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   3360
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "tipo_beneficiario"
            Text            =   ""
         End
      End
      Begin VB.TextBox Txt_mail2 
         DataField       =   "correo_electronico"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   2520
         TabIndex        =   47
         Text            =   "-"
         Top             =   3240
         Width           =   2700
      End
      Begin VB.TextBox TxtTelefono2 
         DataField       =   "telefonos"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Top             =   3240
         Width           =   2355
      End
      Begin VB.TextBox TxtDireccion12 
         DataField       =   "direccion_domicilio"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   1440
         TabIndex        =   45
         Top             =   3660
         Width           =   7155
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado del Registro"
         Height          =   915
         Left            =   6720
         TabIndex        =   43
         Top             =   0
         Width           =   1860
         Begin VB.Label lblActivo2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "N"
            DataField       =   "estado_registro"
            DataSource      =   "adoLista"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            TabIndex        =   44
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.Frame Frame22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Representante Legal o Funcionario Responsable"
         ForeColor       =   &H00000080&
         Height          =   1020
         Left            =   120
         TabIndex        =   35
         Top             =   6240
         Width           =   8475
         Begin MSDataListLib.DataCombo DtcRep_Nombres 
            DataField       =   "NIT"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   36
            Top             =   585
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "nombres"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcRep_Paterno 
            DataField       =   "NIT"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   150
            TabIndex        =   37
            Top             =   585
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "primer_apellido"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcRep_Materno 
            DataField       =   "NIT"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2985
            TabIndex        =   38
            Top             =   585
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "segundo_apellido"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo TxtNIT2 
            DataField       =   "NIT"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   6720
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483637
            ListField       =   "codigo_beneficiario"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nombres:"
            Height          =   195
            Left            =   5745
            TabIndex        =   42
            Top             =   285
            Width           =   675
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Primer Apellido:"
            Height          =   195
            Left            =   225
            TabIndex        =   41
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Segundo Apellido:"
            Height          =   195
            Left            =   3105
            TabIndex        =   40
            Top             =   285
            Width           =   1290
         End
      End
      Begin VB.TextBox txtDenominacion2 
         DataField       =   "denominacion_beneficiario"
         DataSource      =   "adoLista"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   8475
      End
      Begin MSComCtl2.DTPicker DTP_FechaNac2 
         DataField       =   "fecha_nacimiento"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   120
         TabIndex        =   82
         Top             =   1920
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90439681
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTP_FechaExpira2 
         DataField       =   "Fecha_expiracion"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   1320
         TabIndex        =   83
         Top             =   2640
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90439681
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo DtcDepto32 
         DataField       =   "depto_sigla"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   5160
         TabIndex        =   84
         Top             =   740
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483628
         ListField       =   "depto_sigla"
         BoundColumn     =   "depto_sigla"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_doc_id2 
         DataField       =   "Tipo_Documento"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   4920
         TabIndex        =   85
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483628
         ListField       =   "Tipo_Documento"
         BoundColumn     =   "Tipo_Documento"
         Text            =   ""
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "NIT Expedido en:"
         Height          =   255
         Index           =   10
         Left            =   3720
         TabIndex        =   98
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Ingreso Asoc.:"
         Height          =   195
         Left            =   1365
         TabIndex        =   97
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Creación:"
         Height          =   195
         Left            =   165
         TabIndex        =   96
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Domicilio Legal:"
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   4125
         Width           =   1110
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cámara o Asociación a la que Pertenece:"
         Height          =   195
         Left            =   3120
         TabIndex        =   94
         Top             =   2415
         Width           =   2940
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Actividad Principal :"
         Height          =   195
         Left            =   2160
         TabIndex        =   93
         Top             =   1695
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Zona del Domicilio Principal:"
         Height          =   195
         Index           =   23
         Left            =   5400
         TabIndex        =   92
         Top             =   3015
         Width           =   1995
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "No. Registro:"
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   2415
         Width           =   930
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Correo Electrónico :"
         Height          =   195
         Index           =   22
         Left            =   2640
         TabIndex        =   90
         Top             =   3015
         Width           =   1395
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono (Fijo, Fax, Celular):"
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   3015
         Width           =   1965
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Domicilio Principal:"
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   3705
         Width           =   1320
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número NIT :"
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   780
         Width           =   960
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Denominación Unidad Productiva:"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   1080
         Width           =   2430
      End
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
      Left            =   8640
      Top             =   10440
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
      Left            =   15120
      Top             =   10080
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
      Width           =   2160
      _ExtentX        =   3810
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
      Left            =   8640
      Top             =   10080
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
   Begin MSAdodcLib.Adodc adoafp 
      Height          =   330
      Left            =   12960
      Top             =   10440
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin MSAdodcLib.Adodc Ado_datos_busq 
      Height          =   330
      Left            =   15120
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
End
Attribute VB_Name = "frmBeneficiario_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento de Beneficiarios
Option Explicit
Dim rstbeneficiario As New ADODB.Recordset
Dim rst_ben, rsNada As New ADODB.Recordset
Dim rsauxiliar As New ADODB.Recordset
Dim rstbeneaux As New ADODB.Recordset
Dim rs_Depto, rs_Prov, rs_Muni, rs_comunid As New ADODB.Recordset
Dim rs_datos1, rs_datos2, rs_datos3, rs_datos4 As New ADODB.Recordset
Dim rs_datos5, rs_datos6 As New ADODB.Recordset
Dim rs_Depto3 As New ADODB.Recordset
Dim rs_TipoDocId, rs_RepLegal As New ADODB.Recordset
Dim rs_datos_educacionales As New ADODB.Recordset
Dim rs_nivel_educacional As New ADODB.Recordset
Dim rs_laborales As New ADODB.Recordset
Dim rs_tipoInstitucion As New ADODB.Recordset
Dim rs_beneficiario As New ADODB.Recordset
Dim rs_CTA_BCO As New ADODB.Recordset
Dim rs_ocupacion As New ADODB.Recordset
Dim rs_beneficiario_Afp As New ADODB.Recordset
Dim rs_Dependiente, rs_pais As New ADODB.Recordset
Dim rs_contrato, rs_Puesto_Org, rs_UNIDAD, rs_Org, rs_Pry, rs_CARGO As New ADODB.Recordset
Dim rs_correlativo, rsfuente As New ADODB.Recordset
Dim rs_EstCivil, rs_liquidacion As New ADODB.Recordset
Dim rstafp As New ADODB.Recordset
Dim CAMPOS As ADODB.Field

Dim rs_aux17 As New ADODB.Recordset


'BUSQUEDA
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim queryinicial As String
'OTROS
Dim SW As Boolean
Dim SQL_FOR As String
Dim CORREL, accion As Integer
Dim swnuevo As String
Dim V_TIPO, V_TDOC As String
Dim sino As String
Dim marca1 As String
Dim NombreCarpeta, e As String
Dim imag2 As Long
Dim VARB, VARBD, VARG, VARS, VARU, VARP, varCat, VAR10, VAR11, VAR12, VAR13, VAR14, VAR15 As String
Dim VARPU, VARCAN, VARPT As Double

Dim VAR_VAL, CodBenef, VINICIAL As String

Private Sub Ado_Contrato_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If swnuevo = "M" Then
'    If rs_contrato!estado_contrato = "REG" Then
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
'    If rs_contrato!estado_contrato = "REG" Then
''        TxtAprob.ForeColor = &H80&
''        lblARCH.ForeColor = &H80&
'    Else
''        TxtAprob.ForeColor = &H4000&
''        lblARCH.ForeColor = &H4000&
'    End If
  End If
  If Ado_Contrato.Recordset.RecordCount > 0 Then
  
      If Ado_Contrato.Recordset!estado_contrato = "REG" Then
         frm_ro_personal_contrato.TxtAprob.ForeColor = &H4000&
    '        CmdAdd4.Visible = True
    '        CmdMod4.Visible = True
    '        CmdElim4.Visible = False
    '        CmdApr4.Visible = True
      Else
         frm_ro_personal_contrato.TxtAprob.ForeColor = &H80&
    '        CmdAdd4.Visible = True '&H000000C0&
    '        CmdMod4.Visible = False
    '        CmdElim4.Visible = False
    '        CmdApr4.Visible = False
        End If
    
       If Ado_Contrato.Recordset("ARCHIVO") = "Cargar_Archivo" Then
            frm_ro_personal_contrato.lblARCH.ForeColor = &HC0&
            'LblCto.ForeColor = &HC0&
        Else
            frm_ro_personal_contrato.lblARCH.ForeColor = &H8000&
            'LblCto.ForeColor = &H8000&
        End If
  End If
End Sub

Private Sub Ado_Educacionales_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Ado_Educacionales.Recordset.RecordCount > 0 Then
        Dim codig0 As Integer
        codig0 = Ado_Educacionales.Recordset!Codigo_Educacion
    End If
End Sub

Private Sub AdoLiquidacion_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If AdoLiquidacion.Recordset.RecordCount > 0 Then
  
      If AdoLiquidacion.Recordset!estado_codigo = "REG" Then
         ro_Personal_Liquidacion.TxtAprob.ForeColor = &H4000&
    '        CmdAdd4.Visible = True
    '        CmdMod4.Visible = True
    '        CmdElim4.Visible = False
    '        CmdApr4.Visible = True
      Else
         ro_Personal_Liquidacion.TxtAprob.ForeColor = &H80&
    '        CmdAdd4.Visible = True '&H000000C0&
    '        CmdMod4.Visible = False
    '        CmdElim4.Visible = False
    '        CmdApr4.Visible = False
        End If
    
       If AdoLiquidacion.Recordset("ARCHIVO") = "Cargar_Archivo" Then
            ro_Personal_Liquidacion.lblARCH.ForeColor = &HC0&
           'LblLiq.ForeColor = &HC0&
        Else
            ro_Personal_Liquidacion.lblARCH.ForeColor = &H8000&
            'LblLiq.ForeColor = &H8000&
        End If
  End If

End Sub

Private Sub Adolista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'If pRecordset.EOF Or pRecordset.BOF Then
  If Adolista.Recordset.EOF Or Adolista.Recordset.BOF Then
      BtnModificar.Enabled = False
     ' BtnEliminar.Enabled = False
      'TxtTipo.Text = Empty
      txtCodigo = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
      txtDenominacion = Empty
      Exit Sub
  End If
  
   'BtnModificar.Enabled = True
   'BtnEliminar.Enabled = True
  If Adolista.Recordset.RecordCount > 0 Then
    Select Case Adolista.Recordset.EditMode
      Case adEditInProgress
        Frame2.Enabled = False            'Verif. Nombre Proveedor JQA NOV-2009
      Case adEditNone
        Set Img_Foto = Leer_Imagen(db, "Select Foto From rv_personal_contratado Where beneficiario_codigo= '" & Adolista.Recordset!beneficiario_codigo & "' ", "Foto")
        Image2 = Img_Foto
            CmdFoto.Visible = True
                    
        If Adolista.Recordset("Fecha_expiracion") <= Date And Adolista.Recordset("tipoben_codigo") = "2" Then
'            adoLista.Recordset("estado_codigo") = "N"
'            MsgBox "La fecha de validez del CONTRATO ya expiro, será deshabilitado el Consultor:" + adoLista.Recordset("beneficiario_denominacion")
        End If
        'If pRecordset("Fecha_expiracion") <= Date And pRecordset("tipoben_codigo") = "2" Then
        '    pRecordset("estado_codigo") = "N"
        '    MsgBox "La fecha de validez del RENCA ya expiro, será deshabilitado el Consultor:" + pRecordset("beneficiario_denominacion")
        'End If
        Set rs_datos_educacionales = New ADODB.Recordset
        If GlSW <> "ADD" Then
            Call abrirtabla
'            Set rs_datos_educacionales = New ADODB.Recordset'<>
'            rs_datos_educacionales.Open "select * from rc_datos_educacionales where beneficiario_codigo = '" & adoLista.Recordset!beneficiario_codigo & "'  ", DB, adOpenKeyset, adLockOptimistic
'            Set Ado_Educacionales.Recordset = rs_datos_educacionales
'
'            Set rs_laborales = New ADODB.Recordset
'            rs_laborales.Open "select * from rc_experiencia_laboral where beneficiario_codigo = '" & adoLista.Recordset!beneficiario_codigo & "'  ", DB, adOpenKeyset, adLockOptimistic
'            Set Ado_Laborales.Recordset = rs_laborales
        Else
            'Set Ado_ProyUbic.Recordset = RSNADA
            Set DtgEducacionales.DataSource = rsNada
            Set DtgLaborales.DataSource = rsNada
'            rs_ProyUbic.Open "select * from mo_proy_Id_Ubicacion  ", db, adOpenKeyset, adLockOptimistic
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
'        If adoLista.Recordset("tipoben_codigo") = "6" Then
'            SSTab1.Tab = 3
'            SSTab1.TabEnabled(0) = False
''            SSTab1.TabEnabled(3) = True
'        Else
            'SSTab1.Tab = 0
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(3) = True
'        End If
        If Adolista.Recordset!estado_codigo = "REG" Then
            BtnAprobar.Visible = True
            CmdDesapr.Visible = False
        Else
            BtnAprobar.Visible = False
            CmdDesapr.Visible = True
        End If
        'JQA NOV-2010
'        If adoLista.Recordset("tipoben_codigo") = "0" Then
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
    If Adolista.Recordset("estado_codigo") = "APR" Then
        lblActivo.ForeColor = &H8000&
    Else
        lblActivo.ForeColor = &HC0&
    End If
    If Adolista.Recordset("ARCHIVO_FOTO") = "Cargar_Archivo" Then
        LblInicial.ForeColor = &HC0&
    Else
        LblInicial.ForeColor = &H8000&
    End If
    If Adolista.Recordset("ARCHIVO_HOJAVIDA") = "Cargar_Archivo" Then
        LblCV.ForeColor = &HC0&
    Else
        LblCV.ForeColor = &H8000&
    End If
    If Adolista.Recordset("ARCHIVO_RESPALDO") = "Cargar_Archivo" Then
        LblResp.ForeColor = &HC0&
    Else
        LblResp.ForeColor = &H8000&
    End If
    If swnuevo = "X" Then
    'If Not (IsNull(AdoTip_ben.Recordset("tipoben_codigo"))) Then
    '            If Not (AdoTip_ben.Recordset.BOF) Then AdoTip_ben.Recordset.MoveFirst
    '            AdoTip_ben.Recordset.Find "tipoben_codigo='" & adoLista.Recordset!tipoben_codigo & "'", , adSearchForward
    '            If Not AdoTip_ben.Recordset.EOF Then
    '                'TDBC_marcas.Item(1) = AdoMarca.Recordset!descripcion
    '            End If
    'End If
      Adolista.Caption = CStr(Adolista.Recordset.AbsolutePosition) & " de " & CStr(Adolista.Recordset.RecordCount)
    End If
  End If
End Sub
   
Private Sub BtnGrabar_Click()
'Frame2.Visible = True
'If TxtTipo = "6" Then
    V_TIPO = Trim(TxtTipo.Text)
    V_TDOC = Trim(Dtc_doc_id)
'Else
'    V_TIPO = Trim(TxtTipo2.Text)
'    V_TDOC = Trim(Dtc_doc_id2.Text)
'End If
'On Error GoTo errorAceptar
   
   'GC_BENEFICIARIO
   Set rsauxiliar = New ADODB.Recordset
   If rsauxiliar.State = 1 Then rsauxiliar.Close
   rsauxiliar.Open "select * from gc_beneficiario where beneficiario_codigo= '" & txtCodigo & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
   'Set Ado_datos1.Recordset = rsauxiliar
   If rsauxiliar.RecordCount > 0 Then
        VINICIAL = rsauxiliar!beneficiario_iniciales
   End If
   With Adolista
     If swnuevo = "A" Then
       CORREL = 0
'       DE.dbo_fc_correl_ben CORREL
       Set rstbeneaux = New ADODB.Recordset
       SQL_FOR = "select * from Gc_beneficiario where beneficiario_codigo= '" & txtCodigo & "' OR beneficiario_codigo= '" & txtCodigo2.Text & "' "
       rstbeneaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
       'If rstbeneaux.RecordCount > 0 And txtCodigo.Enabled Then
       If rstbeneaux.RecordCount > 0 Then
                SW = True
                MsgBox " CODIGO DUPLICADO"
                'txtCodigo.SetFocus
                Exit Sub
       End If
     End If
       If TxtTipo < "20" Then
            If Trim(txtCodigo) = "" Then
                MsgBox "Introduzca el No. Documento de Identidad :"
                'txtCodigo.SetFocus
                Exit Sub
            End If
'            If TxtTipo.Text = "" Then
'                MsgBox "Introduzca el Tipo de Persona :"
'                TDBtipoben.SetFocus
'                Exit Sub
'            End If
'            If Trim(Text1) = "" Then
'                 MsgBox "INTRODUZCA el Primer Apellido:"
'                 Text1.SetFocus
'                 Exit Sub
'            End If
'            If Trim(Text2) = "" Then
'                 MsgBox "INTRODUZCA el Segundo Apellido:"
'                 Text2.SetFocus
'                 Exit Sub
'            End If
'            If Trim(Text3) = "" Then
'                 MsgBox "INTRODUZCA los Nombres:"
'                 Text3.SetFocus
'                 Exit Sub
'            End If
            If txtDenominacion = "" Then
                MsgBox "INTRODUZCA el nombre completo de la persona:"
                'txtDenominacion.SetFocus
                Exit Sub
            End If
       End If
           ' CORREL = CORREL + 1
        db.BeginTrans
        SW = False
        If txtCodigo.Enabled And swnuevo = "A" Then
            .Recordset("beneficiario_codigo") = Trim(txtCodigo)
'            If TxtTipo2.Text = "22" Then
'                 .Recordset("NIT") = txtCodigo
'            End If
            If TxtTipo.Text < "20" Then
                 .Recordset("NIT") = TxtNIT
            End If
            accion = 0
'            Dim a As String, b As String, C As String, d As String
'            a = Left(Text1.Text, 1)
'            b = Left(Text2.Text, 1)
'            C = Left(Text3.Text, 1)
'            d = Trim(a) + Trim(b) + Trim(C)
'            LblInicial.Caption = Trim(d)
            .Recordset("estado_codigo").Value = "REG"
            .Recordset("archivo_foto_cargado") = "N"
            '.Recordset("ARCHIVO_FOTO") = "Cargar_Archivo"
            .Recordset("ARCHIVO_HOJAVIDA") = "Cargar_Archivo"
            .Recordset("ARCHIVO_RESPALDO") = "Cargar_Archivo"
'            rs_contrato!ARCHIVO_NOMB = Trim(DtcInicial.Text) & "_Contrato_" & rs_contrato!numero_consultoria & ".pdf"
            .Recordset("archivo_foto") = Trim(VINICIAL) & "_Foto.JPG"
            .Recordset("ARCHIVO_HV") = Trim(VINICIAL) & "_HojadeVida_1.pdf"
            .Recordset("ARCHIVO_RESP") = Trim(VINICIAL) & "_Respaldo_1.pdf"
            .Recordset("beneficiario_nro_file") = "0"
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
'            .Recordset("Fecha_expiracion") = DTP_FechaExpira    'IIf(DTP_FechaExpira = "", Format(Date, "dd/mm/yyyy"), DTP_FechaExpira)
            .Recordset("Reg_Profesional") = TxtRenca.Text
            .Recordset("beneficiario_fecha_nacimiento") = DTP_FechaNac.Value
            .Recordset("estado_civil_codigo") = DtcEstCiv.Text
            .Recordset("genero_codigo") = dtc_codigo4.Text
            .Recordset("pais_codigo") = DtcPaisCod.Text
            .Recordset("depto_codigo") = IIf(Dtc_depto_cod.Text = "", "-", Dtc_depto_cod.Text)
            .Recordset("prov_codigo") = IIf(Dtc_prov_cod.Text = "", "-", Dtc_prov_cod.Text)
            .Recordset("munic_codigo") = IIf(Dtc_munic_cod.Text = "", "-", Dtc_munic_cod.Text)
            .Recordset("asegurado_codigo_caja") = txt_ss.Text
            .Recordset("fecha_asegurado_caja") = DTP_FechaSS.Value
            .Recordset("fecha_asegurado_fin_caja") = DTP_FechaSSExp.Value
            .Recordset("beneficiario_codigo_seguro") = DtcSS.Text
            .Recordset("unidad_codigo_pla") = dtc_codigo2.Text
    
            .Recordset("asegurado_codigo_afp") = txt_afp.Text
            .Recordset("beneficiario_codigo_afp") = dtc_afp.Text
            .Recordset("fecha_asegurado_afp") = DTP_FechaAfp.Value
            
            .Recordset("bco_codigo") = DtcBanco.Text
            .Recordset("cta_codigo") = DtcCta.Text 'beneficiario_haber_mensual
            .Recordset("beneficiario_haber_mensual") = txt_sueldo.Text
            .Recordset("beneficiario_otro_mensual") = txt_otro.Text
            .Recordset("fecha_ingreso") = DTPicker2.Value
            If DtcCtaTip.Text = "CUENTA CORRIENTE" Then
                .Recordset("cta_tipo") = "CC"
            Else
                .Recordset("cta_tipo") = "CA"
            End If
            '.Recordset("ocup_codigo") = IIf(Dtc_Ocup.Text = "", 0, Dtc_Ocup.Text)
            .Recordset("archivo_foto") = Trim(VINICIAL) & "_Foto.JPG"
            
            .Recordset("ARCHIVO_hojavida") = Trim(VINICIAL) & "_HojadeVida_1.pdf"
            .Recordset("ARCHIVO_respaldo") = Trim(VINICIAL) & "_Respaldo_1.pdf"
'            Dim a As String, b As String, C As String
'            a = Left(Text1.Text, 1)
'            b = Left(Text2.Text, 1)
'            C = Left(Text3.Text, 1)
'            .Recordset("beneficiario_beneficiario_iniciales") = Trim(LblInicial.Caption)
'            RUTA1 = "PERSONAL" + "\" + Trim(LblInicial) + "-" + Trim(txtCodigo)
'            MsgBox RUTA1
'            MkDir RUTA1
        End If
            '.Recordset("activo").Value = "S"
            .Recordset("usr_codigo").Value = glusuario 'frmLogin.txtUserName.Text
            .Recordset("fecha_registro").Value = Date
            .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
            .Recordset.Update
            '.Recordset.Requery
'            Dim ARCH_FOTO As String
'            ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + adoLista.Recordset("inicial") + "\" + adoLista.Recordset("inicial") + "-FOTO.JPG"
'            If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where beneficiario_codigo= '" & adoLista.Recordset("beneficiario_codigo") & "' ", "Foto", ARCH_FOTO) Then
'                MsgBox "Se cargo la Imagen Correctamente !!"
'            Else
'                MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'            End If
            db.CommitTrans
      End With
   '
   swnuevo = "X"
   fraOpciones.Visible = True
   fra_cabecera.Enabled = True
   
        SSTab1.TabEnabled(1) = True
     SSTab1.TabEnabled(2) = True
     SSTab1.TabEnabled(3) = True
   
   FraGrabarCancelar.Visible = False
   Picture2.Enabled = True
   fraDatos.Enabled = False
   Frame19.Enabled = False
'   FraSS_SS.Enabled = False
'''''''   CmdAdd1.Visible = False
'''''''   CmdMod1.Visible = False
'''''''   CmdElim1.Visible = False
'''''''   CmdApr1.Visible = False
'''''''   CmdAdd2.Visible = False
'''''''   CmdMod2.Visible = False
'''''''   CmdElim2.Visible = False
'''''''   CmdApr2.Visible = False
'''''''   CmdAdd3.Visible = False
'''''''   CmdMod3.Visible = False
'''''''   CmdElim3.Visible = False
'''''''   CmdApr3.Visible = False
'''''''   CmdAdd4.Visible = False
'''''''   CmdMod4.Visible = False
'''''''   CmdElim4.Visible = False
'''''''   CmdApr4.Visible = False
'''''''   CmdAdd5.Visible = False
'''''''   CmdMod5.Visible = False
'''''''   CmdElim5.Visible = False
'''''''   CmdApr5.Visible = False

   Call Carga_Recor
   'Call Carga_Beneficiario
'De.dbo_alGraba_rc_personal Accion, CORREL, txtCodigo.Text, Text1.Text, Text2.Text, Text3.Text, "2002"
 Exit Sub
errorAceptar:
   Call pErrorRst(db.Errors)
   Adolista.Recordset.CancelUpdate
   'db.RollbackTrans
End Sub

Private Sub cmdAdd2_Click()
   If Adolista.Recordset.RecordCount > 0 Then
       marca1 = Adolista.Recordset.Bookmark
       ac_CapturaEstudiosRealizados.txtSW = "ADD"
       ac_CapturaEstudiosRealizados.txtBenef = Adolista.Recordset!beneficiario_codigo
       ac_CapturaEstudiosRealizados.txtEstado = "REG"
       'Ado_Educacionales.Recordset.AddNew
       ac_CapturaEstudiosRealizados.Show vbModal
       'Call abrirtabla
       'Ado_Educacionales.Refresh
   Else
       MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
End Sub

Private Sub cmdAdd3_Click()
   If Adolista.Recordset.RecordCount > 0 Then
        marca1 = Adolista.Recordset.Bookmark
        ac_CapturaExperienciaLaboral.txtSW = "ADD"
        ac_CapturaExperienciaLaboral.txtBenef = Adolista.Recordset!beneficiario_codigo
        ac_CapturaExperienciaLaboral.txtEstado = "REG"
        Ado_Laborales.Recordset.AddNew
        ac_CapturaExperienciaLaboral.Show vbModal
        'Call abrirtabla
        'Ado_Laborales.Refresh
   Else
        MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

End Sub

Private Sub CmdAdd4_Click()
   If Adolista.Recordset.RecordCount > 0 Then
        marca1 = Adolista.Recordset.Bookmark
        frm_ro_personal_contrato.txtSW = "ADD"
        frm_ro_personal_contrato.txtBenef = Adolista.Recordset!beneficiario_codigo
        frm_ro_personal_contrato.TxtInicial = Adolista.Recordset!beneficiario_iniciales
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

End Sub

Private Sub CmdAdd1_Click()
  marca1 = Adolista.Recordset.Bookmark
  ac_CapturaDatosPersonales.txtSW = "ADD"
  ac_CapturaDatosPersonales.txtBenef = Adolista.Recordset!beneficiario_codigo
  'ac_CapturaDatosPersonales.lblcodigo_unidad = adoLista.Recordset!unidad_codigo
  'ac_CapturaDatosPersonales.lblcodigo_solicitud = adoLista.Recordset!solicitud_codigo
  ac_CapturaDatosPersonales.txtEstado = "REG"
 ' AdoDependiente.Recordset.AddNew
  ac_CapturaDatosPersonales.Show vbModal
  'Call abrirtabla
'  AdoDependiente.Refresh
'  AdoDependiente.Recordset.Requery  ' .Refresh
''  adoLista.Recordset.Move marca1 - 1
End Sub

Private Sub CmdAdd5_Click()
   If Adolista.Recordset.RecordCount > 0 Then
        marca1 = Adolista.Recordset.Bookmark
        'AdoLiquidacion.Recordset.AddNew
        ro_Personal_Liquidacion.txtSW = "ADD"
         ro_Personal_Liquidacion.TxtGestion = Year(Date)
         ro_Personal_Liquidacion.TxtGestion_ini = Year(Date) - 5
        ro_Personal_Liquidacion.txtBenef.Text = Adolista.Recordset!beneficiario_codigo
        ro_Personal_Liquidacion.TxtInicial = Adolista.Recordset!beneficiario_iniciales
        ro_Personal_Liquidacion.TxtAprob = "REG"
        ro_Personal_Liquidacion.txtpago1 = Adolista.Recordset!beneficiario_haber_mensual
        ro_Personal_Liquidacion.TxtPago2 = Adolista.Recordset!beneficiario_haber_mensual
        ro_Personal_Liquidacion.Txtpago3 = Adolista.Recordset!beneficiario_haber_mensual
        'frmBeneficiario_Admin.AdoLiquidacion.Recordset!tipo_memo = "REF"
        ro_Personal_Liquidacion.Show vbModal
        'Call abrirtabla
        'AdoLiquidacion.Refresh
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

End Sub

Private Sub CmdApr3_Click()
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
End Sub

Private Sub CmdApr4_Click()
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

Private Sub CmdApr1_Click()
   sino = MsgBox("Está Seguro de APROBAR el Registro del Dependiente ? ", vbYesNo + vbQuestion, "Atención")
   If AdoDependiente.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoDependiente.Recordset("estado_codigo") = "APR"
        AdoDependiente.Recordset("fecha_REGISTRO") = Date
        AdoDependiente.Recordset("usr_usuario") = glusuario
        AdoDependiente.Recordset.Update
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub CmdApr2_Click()
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
End Sub

Private Sub CmdApr5_Click()
  On Error GoTo UpdateErr
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
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdDesapr_Click()
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If Adolista.Recordset("estado_codigo") = "APR" Then
      If sino = vbYes Then
'        Dim RUTA1, RUTA2 As String
'        RUTA1 = "PERSONAL" + "\" + Trim(adoLista.Recordset("beneficiario_iniciales")) + "-" + Trim(adoLista.Recordset("beneficiario_codigo"))
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
        Adolista.Recordset("estado_codigo") = "REG"
        Adolista.Recordset("fecha_aprueba") = Date
        Adolista.Recordset("usr_aprueba") = glusuario
        Adolista.Recordset.Update
        
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Anulado o sin Aprobar ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub CmdElim1_Click()
   sino = MsgBox("Está Seguro de ANULAR el Registro del Dependiente ? ", vbYesNo + vbQuestion, "Atención")
   If AdoDependiente.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoDependiente.Recordset("estado_codigo") = "ANL"
        AdoDependiente.Recordset("fecha_registro") = Date
        AdoDependiente.Recordset("usr_usuario") = glusuario
        AdoDependiente.Recordset("ocupacion_pariente") = "REG. ANULADO"
        AdoDependiente.Recordset.Update  'Batch adAffectAll
         Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub CmdElim2_Click()
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
End Sub

Private Sub CmdElim3_Click()
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
End Sub

Private Sub CmdElim4_Click()
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
End Sub

Private Sub CmdElim5_Click()
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If AdoLiquidacion.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoLiquidacion.Recordset("estado_codigo") = "ANL"
        AdoLiquidacion.Recordset("fecha_registro") = Date
        AdoLiquidacion.Recordset("usr_usuario") = glusuario
'        AdoLiquidacion.Recordset("centro_educativo") = "REG. ANULADO"
        AdoLiquidacion.Recordset.Update  'Batch adAffectAll
         Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub CmdFoto_Click()
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'CMACBCMCB
    On Error GoTo QError
    Dim ARCH_FOTO As String
    Dim SW0 As String
    'If adoLista.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
    If Adolista.Recordset!archivo_foto_cargado = "N" Then
      NombreCarpeta = App.Path & "\PERSONAL\" & Trim(Adolista.Recordset!beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FOT1"
'      If GlServidor = "SERVIDOR2" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
      SW0 = 1
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\PERSONAL\" & Trim(Adolista.Recordset!beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FOT1"
'          If GlServidor = "SERVIDOR2" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
          SW0 = 1
      Else
        SW0 = 0
      End If
    End If
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    If SW0 = 1 Then
    '    If GlServidor = "SRVPRO" Then
    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
    '    Else
            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(dtc_campo2.Text) + "\" + Trim(Ado_datos.Recordset("edif_codigo")) + "\" + Trim(Ado_datos.Recordset("edif_codigo")) + ".JPG"
        ARCH_FOTO = App.Path + "\PERSONAL\" + Trim(Adolista.Recordset!beneficiario_iniciales) + "-" + Trim(Adolista.Recordset("beneficiario_codigo")) + "\" + Trim(Adolista.Recordset!ARCHIVO_Foto)
            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!codigo1) + "\" + Trim(Ado_datos.Recordset!edif_codigo) + ".JPG"
    '    End If
        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
        CodBenef = Adolista.Recordset!beneficiario_codigo
        'CodBien = Ado_datos.Recordset!edif_codigo
        If Guardar_Imagen(db, "Select Foto From ro_personal_contratado Where beneficiario_codigo= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
            MsgBox "Se cargo la Imagen Correctamente !!"
        Else
            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
        End If
    Else
        Set Img_Foto = Leer_Imagen(db, "Select Foto From ro_personal_contratado Where beneficiario_codigo = '" & Adolista.Recordset("beneficiario_codigo") & "' ", "Foto")
        Image2 = Img_Foto
    End If
'  Else
'    MsgBox "Debe Aprobar el registro, para crear la carpeta correspondiente..."
'  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If

End Sub

'Private Sub CmdGraba_Click()
'   If adoLista.Recordset.RecordCount > 0 And Not IsNull(DtgEducacionales.Columns("nivel_educacional").Value) And (DtgEducacionales.Columns("nivel_educacional").Value) <> "" Then
'      If adoLista.Recordset!estado_codigo = "REG" Then
'        marca1 = adoLista.Recordset.Bookmark
'        VARB = DtgEducacionales.Columns("beneficiario_codigo").Value
'        VARBD = DtgEducacionales.Columns("Carrera_Curso").Value
'        VARG = DtgEducacionales.Columns("centro_educativo").Value
'        VARS = DtgEducacionales.Columns("titulo_obtenido").Value
'        VARU = DtgEducacionales.Columns("nivel_educacional").Value
'        VARPU = DtgEducacionales.Columns("duracion_años").Value
'        VAR10 = DtgEducacionales.Columns("pais").Value
'        VAR11 = DtgEducacionales.Columns("ciudad").Value
'        VAR12 = DtgEducacionales.Columns("fecha_inicio").Value
'        VAR13 = DtgEducacionales.Columns("fecha_fin").Value
'        VAR14 = DtgEducacionales.Columns("presento_documento").Value
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
'        DtgEducacionales.AllowAddNew = False
'        DtgEducacionales.AllowDelete = False
'        DtgEducacionales.AllowUpdate = False
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
'   If adoLista.Recordset.RecordCount > 0 And Not IsNull(DtgLaborales.Columns("tipo_institucion").Value) And (DtgLaborales.Columns("tipo_institucion").Value) <> "" Then
'      If adoLista.Recordset!estado_codigo = "REG" Then
'        marca1 = adoLista.Recordset.Bookmark
'        VARB = DtgLaborales.Columns("beneficiario_codigo").Value
'        VARBD = DtgLaborales.Columns("denominacion_institucion").Value
'        VARG = DtgLaborales.Columns("tipo_institucion").Value
'        VARS = DtgLaborales.Columns("cargo").Value
'        VARU = DtgLaborales.Columns("funcion_general").Value
'        VARPU = DtgLaborales.Columns("Tiempo_Meses").Value
'        VAR10 = DtgLaborales.Columns("pais").Value
'        VAR11 = DtgLaborales.Columns("ciudad").Value
'        VAR12 = DtgLaborales.Columns("fecha_inicio").Value
'        VAR13 = DtgLaborales.Columns("fecha_fin").Value
'        VAR14 = DtgLaborales.Columns("presento_documento").Value
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
'        DtgLaborales.AllowAddNew = False
'        DtgLaborales.AllowDelete = False
'        DtgLaborales.AllowUpdate = False
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
'      rs_contrato!beneficiario_codigo = adoLista.Recordset("beneficiario_codigo") 'DtcBenef.Text
'      rs_contrato!ges_gestion = glGestion
'      rs_contrato!solicitud_codigo = rs_contrato.RecordCount
'
'      Set rs_correlativo = New ADODB.Recordset
'      rs_correlativo.Open "select * from ro_contratos_personas WHERE beneficiario_codigo = '" & adoLista.Recordset("beneficiario_codigo") & "'  ", DB, adOpenKeyset, adLockOptimistic
'      If rs_correlativo.RecordCount > 0 Then
'            rs_contrato!numero_consultoria = rs_correlativo.RecordCount
''            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
''            rs_correlativo.Update
''            rs_M1!Numero_FA = rs_correlativo!correlativo
'      Else
'            rs_contrato!numero_consultoria = 1
'      End If
'      rs_contrato!ARCHIVO = "Cargar_Archivo"
'      rs_contrato!ARCHIVO_NOMB = Trim(adoLista.Recordset("beneficiario_beneficiario_iniciales")) & "_Contrato_" & rs_contrato!numero_consultoria & ".pdf"
'      TxtAprob.Text = "REG"
'    End If
'      rs_contrato!objeto_contrato = txtObjContrato.Text
'      rs_contrato!puesto_codigo = DtcPuesto.Text
'      rs_contrato!unidad_codigo = Dtc_codigo.Text
'      rs_contrato!codigo_convenio = DtcOrg.Text
'      rs_contrato!pro_codigo = DtcPry.Text
'      rs_contrato!fechas_confirmado = Txtestado
'      rs_contrato!estado_contrato = TxtAprob
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
     If Adolista.Recordset.RecordCount > 0 Then
      If Ado_Educacionales.Recordset!estado_codigo = "REG" Then
        marca1 = Adolista.Recordset.Bookmark
        ac_CapturaEstudiosRealizados.txtSW = "MOD"
        ac_CapturaEstudiosRealizados.txtBenef = Adolista.Recordset!beneficiario_codigo
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
End Sub

Private Sub CmdMod3_Click()
     If Adolista.Recordset.RecordCount > 0 Then
      If Ado_Laborales.Recordset!estado_codigo = "REG" Then
        marca1 = Adolista.Recordset.Bookmark
        ac_CapturaExperienciaLaboral.txtSW = "MOD"
        ac_CapturaExperienciaLaboral.txtBenef = Adolista.Recordset!beneficiario_codigo
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
End Sub

Private Sub CmdMod4_Click()
  On Error GoTo EditErr
   If Adolista.Recordset.RecordCount > 0 Then
      If Ado_Contrato.Recordset!estado_contrato = "REG" Then
        marca1 = Adolista.Recordset.Bookmark
        frm_ro_personal_contrato.txtSW = "MOD"
        frm_ro_personal_contrato.txtBenef = Adolista.Recordset!beneficiario_codigo
        frm_ro_personal_contrato.TxtInicial = Adolista.Recordset!beneficiario_iniciales
        frm_ro_personal_contrato.TxtForm = Ado_Contrato.Recordset!solicitud_codigo
        frm_ro_personal_contrato.TxtAprob = Ado_Contrato.Recordset!estado_contrato
        frm_ro_personal_contrato.lblARCH.Caption = Ado_Contrato.Recordset!ARCHIVO
        frm_ro_personal_contrato.txtCodigo.Text = Ado_Contrato.Recordset!codigo_contrato
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

Private Sub CmdMod1_Click()
 If AdoDependiente.Recordset!estado_codigo = "REG" Then
    marca1 = Adolista.Recordset.Bookmark
    ac_CapturaDatosPersonales.txtSW = "MOD"
    ac_CapturaDatosPersonales.txtBenef = Adolista.Recordset!beneficiario_codigo
    ac_CapturaDatosPersonales.txtCI = AdoDependiente.Recordset!cod_dependiente
    ac_CapturaDatosPersonales.TxtItem = AdoDependiente.Recordset!Cod_asegurado
    ac_CapturaDatosPersonales.DTPFec_Seguro = AdoDependiente.Recordset!Fecha_asegurado
    ac_CapturaDatosPersonales.txtNac = AdoDependiente.Recordset!Fecha_Nacimiento
    ac_CapturaDatosPersonales.txtPat = AdoDependiente.Recordset!primer_apellido
    ac_CapturaDatosPersonales.txtMat = AdoDependiente.Recordset!segundo_apellido
    ac_CapturaDatosPersonales.txtNom = AdoDependiente.Recordset!NombreS
    ac_CapturaDatosPersonales.Dtc_Par = AdoDependiente.Recordset!pariente_codigo
    ac_CapturaDatosPersonales.Dtc_ParDes = AdoDependiente.Recordset!pariente_descripcion
    ac_CapturaDatosPersonales.TxtOcupacion = AdoDependiente.Recordset!ocupacion_pariente
    ac_CapturaDatosPersonales.txtEstado = AdoDependiente.Recordset!estado_codigo
    'AdoDependiente.Recordset.AddNew
    ac_CapturaDatosPersonales.Show vbModal
 Else
        MsgBox "No se puede MODIFICAR un registro Aprobado o Anulado ...", vbExclamation, "Validación de Registro"
 End If
 'AdoDependiente.Refresh
End Sub

Private Sub CmdMod5_Click()
 On Error GoTo EditErr
   If Adolista.Recordset.RecordCount > 0 Then
      If AdoLiquidacion.Recordset!estado_codigo = "REG" Then
        marca1 = Adolista.Recordset.Bookmark
        ro_Personal_Liquidacion.txtSW = "MOD"
        ro_Personal_Liquidacion.TxtGestion = AdoLiquidacion.Recordset!ges_gestion
        ro_Personal_Liquidacion.TxtGestion_ini = IIf(IsNull(AdoLiquidacion.Recordset!ges_gestion_ini), Year(Date), AdoLiquidacion.Recordset!ges_gestion_ini)
        ro_Personal_Liquidacion.TxtGestion = AdoLiquidacion.Recordset!ges_gestion
        ro_Personal_Liquidacion.txtBenef = Adolista.Recordset!beneficiario_codigo
        ro_Personal_Liquidacion.TxtInicial = Adolista.Recordset!beneficiario_iniciales
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
        ro_Personal_Liquidacion.TxtCta.Text = IIf(IsNull(AdoLiquidacion.Recordset!Cta_Codigo), "0", AdoLiquidacion.Recordset!Cta_Codigo)
        ro_Personal_Liquidacion.TxtDeduccion.Text = IIf(IsNull(AdoLiquidacion.Recordset!Deducciones), "0", AdoLiquidacion.Recordset!Deducciones)
        ro_Personal_Liquidacion.TxtTotBenef.Text = IIf(IsNull(AdoLiquidacion.Recordset!monto_total), "0", AdoLiquidacion.Recordset!monto_total)
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

Private Sub CmdVerDisco_Click()
'  On Error GoTo Error_Sub
'    marca1 = adoLista.Recordset.Bookmark
'    FrmExplora.lblges_gestion = adoLista.Recordset!primer_apellido + " " + adoLista.Recordset!segundo_apellido + " " + adoLista.Recordset!NombreS
'    FrmExplora.LblFA = adoLista.Recordset!beneficiario_codigo
'    'FrmExplora.LblForm = adoLista.Recordset!tipo_formulario
''    sino = MsgBox("Elija <SI> para ver la Información de su Disco Local. , o del Servidor <NO> ", vbQuestion + vbYesNo, "Confirmando...")
''    If sino = vbYes Then
'    NombreCarpeta = App.Path & "\PERSONAL\" & adoLista.Recordset!beneficiario_codigo
'    e = App.Path & "\PERSONAL\" & adoLista.Recordset!beneficiario_codigo
''    NombreCarpeta = App.Path & "\PERSONAL\" & adoLista.Recordset!beneficiario_beneficiario_iniciales
''    e = App.Path & "\PERSONAL\" & adoLista.Recordset!beneficiario_beneficiario_iniciales
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
'    Set rsauxiliar = New ADODB.Recordset
'    'SQL_FOR = "select * from rc_personal where ci = '" & txtCodigo.Text & "'"
'    'rsauxiliar.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
'    rsauxiliar.Open "select * from rc_personal where ci = '" & TxtCodigo & "' ", db, adOpenKeyset, adLockOptimistic
'    If rsauxiliar.RecordCount = 0 Then
'        rsauxiliar.AddNew
'        rsauxiliar!ci = TxtCodigo
'        rsauxiliar!idfuncionario = CORREL
'    Else
'        'MsgBox " YA EXISTE EL CODIGO ..."
'    End If
'        rsauxiliar!tipoben_codigo = Trim(TxtTipo.Text) 'JQA NOV-2009
''        rsauxiliar!paterno = Text1.Text
''        rsauxiliar!materno = Text2.Text
''        rsauxiliar!NombreS = Text3.Text
''        Dim a As String, b As String, C As String
''        'a = Left(Linea, 5)
''        'b = Mid(Linea, 11, 4)
''        a = Left(Text1.Text, 1)
''        b = Left(Text2.Text, 1)
''        C = Left(Text3.Text, 1)
''        rsauxiliar!inicial = a + b + C
''        If txtDenominacion.Text = "" Then
''            rsauxiliar!nro_cite_no_objecion = Trim(Text1.Text) + " " + Trim(Text2.Text) + " " + Trim(Text3.Text)
''        Else
''            rsauxiliar!nro_cite_no_objecion = txtDenominacion.Text
''        End If
'        rsauxiliar!ruc_id = TxtNIT
'        rsauxiliar!Fecha_Nacimiento = DTP_FechaNac.Value
'        'rsauxiliar!calle_domicilio = TxtDireccion
''        rsauxiliar!zona_domicilio = TxtZona.Text
''        rsauxiliar!Telefono = TxtTelefono.Text
'        rsauxiliar!Status = "S"
'        rsauxiliar!Activo = "S"
'        rsauxiliar!usr_usuario = glusuario 'frmLogin.txtUserName.Text
'        rsauxiliar!fecha_registro = Date
'        rsauxiliar!hora_registro = Format(Time, "HH:mm:ss")
'        rsauxiliar!departamento_nacimiento = Dtc_depto.Text
'        rsauxiliar!Procedencia = Dtc_prov.Text
'        rsauxiliar!lugar_procedencia = Dtc_munic.Text
'        'rsauxiliar!cargo_codigo = "-"   'TxtCargo.Text
''        rsauxiliar!numero_folder = Txt_mail.Text
'        rsauxiliar!profesion = TxtProfesion.Text
'        rsauxiliar.Update
'        MkDir TxtCodigo
'        If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where beneficiario_codigo= '" & adoLista.Recordset("beneficiario_codigo") & "' ", "Foto", App.Path) Then
'            MsgBox "ok"
'        Else
'            MsgBox "ERR"
'        End If
'        'Guardar_Imagen(cn, Sql, Campo, Path_Imagen)
End Sub

Private Sub BtnAñadir_Click()
   swnuevo = "A"
   Adolista.Recordset.AddNew
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
      SSTab1.TabEnabled(2) = False
      SSTab1.TabEnabled(3) = False
      
'      TxtCodigo.Enabled = True
      fraDatos.Enabled = True
      Frame19.Enabled = True
      Frame2.Enabled = True
'      txtCodigo = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
      TxtTipo.Text = Empty
'      txtDenominacio = Empty
        'Carga_Recor
'      TDBtipoben.SetFocus
'      rst_ben.Open "SELECT * FROM gc_Tipo_Beneficiario where estado_codigo ='APR' ORDER BY descripcion ", db, adOpenStatic
'   End If
    Set AdoTip_ben.Recordset = rst_ben
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    Picture2.Enabled = False
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
    
    Dtc_depto.Enabled = True
    Dtc_prov.Enabled = False
    Dtc_munic.Enabled = False
'    Dtc_local.Enabled = False
End Sub

'Private Sub rep_legal()
'   Set rs_RepLegal = New ADODB.Recordset
'   If rs_RepLegal.State = 1 Then rs_RepLegal.Close
'   rs_RepLegal.Open "select * from gc_Beneficiario WHERE tipoben_codigo = '3' ", db, adOpenKeyset, adLockOptimistic, adCmdText
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

Private Sub BtnBuscar_Click()
'Fra_Busqueda.Visible = True
'fradatos.Enabled = True
' Set ClBuscaGrid = New ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = db
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = grdlista
'    ClBuscaGrid.QueryUtilizado = queryinicial
'    Set ClBuscaGrid.RecordsetTrabajo = adoLista.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = False
  'Dim GrSqlAux As String
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = grdlista
  ClBuscaGrid.QueryUtilizado = queryinicial
  'Set ClBuscaGrid.RecordsetTrabajo = adoLista.Recordset
  Set ClBuscaGrid.RecordsetTrabajo = rstbeneficiario.DataSource
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True

End Sub

Private Sub BtnAprobar_Click()
   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If Adolista.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        If Adolista.Recordset("no_file") <> 1 Then
            Dim RUTA1, RUTA2 As String
            RUTA1 = "PERSONAL" + "\" + Trim(Adolista.Recordset("beneficiario_beneficiario_iniciales")) + "-" + Trim(Adolista.Recordset("beneficiario_codigo"))
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
            Adolista.Recordset("no_file") = 1
        End If
        Adolista.Recordset("estado_codigo") = "APR"
        Adolista.Recordset("fecha_aprueba") = Date
        Adolista.Recordset("usr_aprueba") = glusuario
        Adolista.Recordset.Update
        
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub BtnEliminar_Click()
   sino = MsgBox("Está Seguro de ANULAR el Registro?", vbYesNo + vbQuestion, "Atención")
   If Adolista.Recordset("estado_codigo") = "APR" Or Adolista.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Adolista.Recordset("estado_codigo") = "ANL"
        Adolista.Recordset("fecha_aprueba") = Date
        Adolista.Recordset("usr_codigo_apr") = glusuario
        Adolista.Recordset.Update  'Batch adAffectAll
         Call abrirtabla
      End If
   Else
        MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
'   TxtCodigo.Enabled = True
   fraDatos.Enabled = False
   Frame19.Enabled = False
   'Carga_Recor
   swnuevo = "X"
   
   fraOpciones.Visible = True
       fra_cabecera.Enabled = True
       
            SSTab1.TabEnabled(1) = True
     SSTab1.TabEnabled(2) = True
     SSTab1.TabEnabled(3) = True
       
   FraGrabarCancelar.Visible = False
   Picture2.Enabled = True
'   FraSS_SS.Enabled = False
'''''''   CmdAdd1.Visible = False
'''''''   CmdMod1.Visible = False
'''''''   CmdElim1.Visible = False
'''''''   CmdApr1.Visible = False
'''''''   CmdAdd2.Visible = False
'''''''   CmdMod2.Visible = False
'''''''   CmdElim2.Visible = False
'''''''   CmdApr2.Visible = False
'''''''   CmdAdd3.Visible = False
'''''''   CmdMod3.Visible = False
'''''''   CmdElim3.Visible = False
'''''''   CmdApr3.Visible = False
'''''''   CmdAdd4.Visible = False
'''''''   CmdMod4.Visible = False
'''''''   CmdElim4.Visible = False
'''''''   CmdApr4.Visible = False
'''''''   CmdAdd5.Visible = False
'''''''   CmdMod5.Visible = False
'''''''   CmdElim5.Visible = False
'''''''   CmdApr5.Visible = False
   Call Carga_Recor
   'Call Carga_Beneficiario
'   adoLista.Recordset.Requery
'   adoLista.Refresh
'   Grdlista.ReBind
'   Grdlista.Refresh
End Sub

Private Sub BtnModificar_Click()
'  If adoLista.Recordset("estado_codigo") = "N" Then
     swnuevo = "M"
     Set rst_ben = New ADODB.Recordset
     SSTab1.Tab = 0
    
     SSTab1.TabEnabled(1) = True
     SSTab1.TabEnabled(2) = False
     SSTab1.TabEnabled(3) = False
     If Adolista.Recordset("estado_codigo") = "APR" Then
        MsgBox "El registro está APROBADO, solo se puede modificar por usuarios Autorizados ..."
        Frame2.Enabled = False
'        Frame1.Enabled = False
     Else
        Frame2.Enabled = True
'        Frame1.Enabled = True
     End If
     fraDatos.Enabled = True
     Frame19.Enabled = True
     DTP_FechaNac.Enabled = True
     'TxtRenca.SetFocus
'     FraSS_SS.Enabled = True
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
   
     fraOpciones.Visible = False
     fra_cabecera.Enabled = False
     FraGrabarCancelar.Visible = True
     Picture2.Enabled = False
'     FraSS_SS.Enabled = True
'     TxtCodigo.Enabled = False
     
     Dtc_depto.Enabled = True
     Dtc_prov.Enabled = False
     Dtc_munic.Enabled = False
'     Dtc_local.Enabled = False
    
     rst_ben.Open "SELECT * FROM gc_Tipo_Beneficiario where estado_codigo ='APR' ORDER BY tipoben_descripcion ", db, adOpenStatic
     Set AdoTip_ben.Recordset = rst_ben
     
'   If adoLista.Recordset("tipoben_codigo") = "6" Then
'      SSTab1.Tab = 1
'      SSTab1.TabEnabled(1) = True
'      SSTab1.TabEnabled(0) = False
'      Frame12.Enabled = False
'      TxtNIT2.Enabled = False
'      txtCodigo2.Enabled = False
''      fraDatos2.Enabled = True
'      Frame22.Enabled = True
'      If adoLista.Recordset("estado_codigo") = "N" Then
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
'   If (adoLista.Recordset("tipoben_codigo") = "1" Or adoLista.Recordset("tipoben_codigo") = "2" Or adoLista.Recordset("tipoben_codigo") = "7") Then
      
'  Else
'     MsgBox "El registro está APROBADO, solo se puede modificar por usuarios Autorizados ..."
'     '   Frame2.Enabled = False
'      '  Frame1.Enabled = False
'       ' TxtNIT.SetFocus
'  End If
    
End Sub

Private Sub BtnImprimir_Click()
'  Dim iResult As Integer
'     CrystalReport1.WindowShowPrintSetupBtn = True
'     CrystalReport1.WindowShowRefreshBtn = True
'    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\Presupuesto\beneficiarios\crybeneficiario.rpt"
'    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\Generales\crybeneficiario.rpt"
'    CrystalReport1.ReportFileName = App.Path & "\REPORTES\clasificadores\crybeneficiario.rpt"
'  iResult = CrystalReport1.PrintReport
'  If iResult <> 0 Then
'      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
'  End If
'
'CrystalReport1.WindowState = crptMaximized
'
''    repbeneficiario.Show
'    '   rptModalidadSeleccion.Show vbModal
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

Private Sub dtc_afp_Click(Area As Integer)
    dtc_afp_des.BoundText = dtc_afp.BoundText
    dtc_afp_dir.BoundText = dtc_afp.BoundText
End Sub
Private Sub dtc_afp_des_Click(Area As Integer)
    dtc_afp.BoundText = dtc_afp_des.BoundText
    dtc_afp_dir.BoundText = dtc_afp_des.BoundText
End Sub

Private Sub dtc_afp_dir_Click(Area As Integer)
    dtc_afp.BoundText = dtc_afp_dir.BoundText
    dtc_afp_des.BoundText = dtc_afp_dir.BoundText
End Sub

Private Sub dtc_buscar_ci_Click(Area As Integer)
dtc_buscar_desc.BoundText = dtc_buscar_ci.BoundText
End Sub

Private Sub dtc_buscar_desc_Click(Area As Integer)
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

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub DtcBanco_Click(Area As Integer)
    DtcBancoDes.BoundText = DtcBanco.BoundText
End Sub

Private Sub DtcBancoDes_Click(Area As Integer)
    DtcBanco.BoundText = DtcBancoDes.BoundText
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
    Dtc_depto_cod.BoundText = Dtc_depto.BoundText
    Call pProvincia(Dtc_depto_cod.BoundText)
    Dtc_prov.Enabled = True
    Dtc_munic.Enabled = False
'    Dtc_local.Enabled = False
End Sub

Private Sub Dtc_depto_cod_Click(Area As Integer)
    Dtc_depto.BoundText = Dtc_depto_cod.BoundText
    Call pProvincia(Dtc_depto.BoundText)
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

Private Sub Dtc_depto_cod02_Click(Area As Integer)
    Dtc_depto02.BoundText = Dtc_depto_cod02.BoundText
    Call pProvincia02(Dtc_depto02.BoundText)
End Sub

Private Sub Dtc_depto_cod22_Click(Area As Integer)
    Dtc_depto22.BoundText = Dtc_depto_cod22.BoundText
    Call pProvincia22(Dtc_depto22.BoundText)
End Sub

Private Sub Dtc_depto02_Click(Area As Integer)
    Dtc_depto_cod02.BoundText = Dtc_depto02.BoundText
    Call pProvincia02(Dtc_depto_cod02.BoundText)
End Sub

Private Sub pProvincia02(depto_codigo02 As String)
   Dim strConsultaP02 As String

   strConsultaP02 = "select * from GC_Provincia where depto_codigo='" & depto_codigo02 & "'"

   Set Dtc_prov_cod02.RowSource = Nothing
   Set Dtc_prov_cod02.RowSource = db.Execute(strConsultaP02, , adCmdText)
   Dtc_prov_cod02.ReFill
   Dtc_prov_cod02.BoundText = Empty

   Set Dtc_prov02.RowSource = Nothing
   Set Dtc_prov02.RowSource = db.Execute(strConsultaP02, , adCmdText)
   Dtc_prov02.ReFill
   Dtc_prov02.BoundText = Empty
End Sub

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

Private Sub Dtc_depto22_Click(Area As Integer)
    Dtc_depto_cod22.BoundText = Dtc_depto22.BoundText
    Call pProvincia22(Dtc_depto_cod22.BoundText)
End Sub

Private Sub pProvincia22(depto_codigo22 As String)
   Dim strConsultaP22 As String

   strConsultaP22 = "select * from GC_Provincia where depto_codigo='" & depto_codigo22 & "'"

   Set Dtc_prov_cod22.RowSource = Nothing
   Set Dtc_prov_cod22.RowSource = db.Execute(strConsultaP22, , adCmdText)
   Dtc_prov_cod22.ReFill
   Dtc_prov_cod22.BoundText = Empty

   Set Dtc_prov22.RowSource = Nothing
   Set Dtc_prov22.RowSource = db.Execute(strConsultaP22, , adCmdText)
   Dtc_prov22.ReFill
   Dtc_prov22.BoundText = Empty
End Sub

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

Private Sub Dtc_munic02_Click(Area As Integer)
    Dtc_munic_cod02.BoundText = Dtc_munic02.BoundText
    'Call pComunidad02(Dtc_munic_cod02.BoundText)
End Sub

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

Private Sub Dtc_munic22_Click(Area As Integer)
    Dtc_munic_cod22.BoundText = Dtc_munic22.BoundText
    'Call pComunidad22(Dtc_munic_cod22.BoundText)
End Sub

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
    Dtc_munic.Enabled = True
'    Dtc_local.Enabled = False
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

Private Sub Dtc_prov_cod02_Click(Area As Integer)
    Dtc_prov02.BoundText = Dtc_prov_cod02.BoundText
    Call pMunicipio02(Dtc_prov02.BoundText)
End Sub


Private Sub Dtc_prov_cod22_Click(Area As Integer)
    Dtc_prov22.BoundText = Dtc_prov_cod22.BoundText
    Call pMunicipio22(Dtc_prov22.BoundText)
End Sub

Private Sub Dtc_prov02_Click(Area As Integer)
    Dtc_prov_cod02.BoundText = Dtc_prov02.BoundText
    Call pMunicipio02(Dtc_prov_cod02.BoundText)
End Sub

Private Sub pMunicipio02(CodProv02 As String)
   Dim strConsultaM02 As String

   strConsultaM02 = "select * from gc_Municipio where prov_codigo='" & CodProv02 & "'"

   Set Dtc_munic_cod02.RowSource = Nothing
   Set Dtc_munic_cod02.RowSource = db.Execute(strConsultaM02, , adCmdText)
   Dtc_munic_cod02.ReFill
   Dtc_munic_cod02.BoundText = Empty

   Set Dtc_munic02.RowSource = Nothing
   Set Dtc_munic02.RowSource = db.Execute(strConsultaM02, , adCmdText)
   Dtc_munic02.ReFill
   Dtc_munic02.BoundText = Empty
End Sub

'Private Sub Dtc_prov2_Click(Area As Integer)
'    Dtc_prov_cod2.BoundText = Dtc_prov2.BoundText
'    Call pMunicipio2(Dtc_prov_cod2.BoundText)
'End Sub

'Private Sub Dtc_prov_cod2_Click(Area As Integer)
'    Dtc_prov2.BoundText = Dtc_prov_cod2.BoundText
'    Call pMunicipio2(Dtc_prov2.BoundText)
'End Sub

Private Sub Dtc_munic_Click(Area As Integer)
    Dtc_munic_cod.BoundText = Dtc_munic.BoundText
    'Call pComunidad(Dtc_munic_cod.BoundText)
    'Dtc_local.Enabled = True
End Sub

Private Sub Dtc_munic_cod_Click(Area As Integer)
    Dtc_munic.BoundText = Dtc_munic_cod.BoundText
    'Call pComunidad(Dtc_munic.BoundText)
    'tc_local.Enabled = True
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

Private Sub Dtc_prov22_Click(Area As Integer)
    Dtc_prov_cod22.BoundText = Dtc_prov22.BoundText
    Call pMunicipio22(Dtc_prov_cod22.BoundText)
End Sub

Private Sub pMunicipio22(CodProv22 As String)
   Dim strConsultaM22 As String

   strConsultaM22 = "select * from gc_Municipio where prov_codigo='" & CodProv22 & "'"

   Set Dtc_munic_cod22.RowSource = Nothing
   Set Dtc_munic_cod22.RowSource = db.Execute(strConsultaM22, , adCmdText)
   Dtc_munic_cod22.ReFill
   Dtc_munic_cod22.BoundText = Empty

   Set Dtc_munic22.RowSource = Nothing
   Set Dtc_munic22.RowSource = db.Execute(strConsultaM22, , adCmdText)
   Dtc_munic22.ReFill
   Dtc_munic22.BoundText = Empty
End Sub



Private Sub DtcPaisCod_Click(Area As Integer)
    DtcPaisSigla.BoundText = DtcPaisCod.BoundText
    TxtNacionalidad.BoundText = DtcPaisCod.BoundText
End Sub

Private Sub DtcRep_Materno_Click(Area As Integer)
    TxtNIT2.BoundText = DtcRep_Materno.BoundText
    DtcRep_Paterno.BoundText = DtcRep_Materno.BoundText
    DtcRep_Nombres.BoundText = DtcRep_Materno.BoundText
End Sub

Private Sub DtcRep_Nombres_Click(Area As Integer)
    TxtNIT2.BoundText = DtcRep_Nombres.BoundText
    DtcRep_Paterno.BoundText = DtcRep_Nombres.BoundText
    DtcRep_Materno.BoundText = DtcRep_Nombres.BoundText
End Sub

Private Sub DtcRep_Paterno_Click(Area As Integer)
    TxtNIT2.BoundText = DtcRep_Paterno.BoundText
    DtcRep_Materno.BoundText = DtcRep_Paterno.BoundText
    DtcRep_Nombres.BoundText = DtcRep_Paterno.BoundText
End Sub

Private Sub DtcRep_Paterno_LostFocus()
'    Text102.Text = DtcRep_Paterno.Text
'    Text202.Text = DtcRep_Materno.Text
'    Text302.Text = DtcRep_Nombres.Text
End Sub

Private Sub DtcSS_Click(Area As Integer)
    DtcSSEnt.BoundText = DtcSS.BoundText
    DtcSSDir.BoundText = DtcSS.BoundText
End Sub

Private Sub DtcSSDir_Click(Area As Integer)
    DtcSS.BoundText = DtcSSDir.BoundText
    DtcSSEnt.BoundText = DtcSSDir.BoundText
End Sub

Private Sub DtcSSEnt_Click(Area As Integer)
    DtcSS.BoundText = DtcSSEnt.BoundText
    DtcSSDir.BoundText = DtcSSEnt.BoundText
End Sub

Private Sub Form_Load()
'   Label5.Caption = GlUsuario       'frmLogin.txtUserName.Text      'JQA NOV-2009
   fraDatos.Enabled = False
   Frame19.Enabled = False
'   fraDatos2.Enabled = False
'   FraSS_SS.Enabled = False
   Call Carga_Recor             'Clasificadores
   Call Carga_Beneficiario(1)
   Call Carga_Beneficiario(2)
'   Call rep_legal
   Call abrirtabla              'Tablas del Proceso
  ' Call Carga_afp
   GlSW = ""
   swnuevo = "X"
'   Fra_ABM.Enabled = False
   If Not Adolista.Recordset.EOF Then
'        If adoLista.Recordset("tipoben_codigo") = "6" Then
'            SSTab1.Tab = 3
'            SSTab1.TabEnabled(0) = False
'            SSTab1.TabEnabled(1) = False
'            SSTab1.TabEnabled(2) = False
''            SSTab1.TabEnabled(3) = True
'        Else
            SSTab1.Tab = 0
            SSTab1.TabEnabled(3) = True
            SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = True
'        End If
'        DtgEducacionales.AllowAddNew = False
'        DtgEducacionales.AllowDelete = False
'        DtgEducacionales.AllowUpdate = False
'        DtgLaborales.AllowAddNew = False
'        DtgLaborales.AllowDelete = False
'        DtgLaborales.AllowUpdate = False


'se culta los botones
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
        '----------------------------------
   End If
   ' Set ClBuscaGrid = Nothing
	Call SeguridadSet(Me)
End Sub

Private Sub Carga_Beneficiario(posicion As Integer)
Select Case posicion
 Case 1
   Set rstbeneficiario = New ADODB.Recordset
   If rstbeneficiario.State = 1 Then rstbeneficiario.Close
   queryinicial = "select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rstbeneficiario.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstbeneficiario.Sort = "beneficiario_denominacion"
   Set Adolista.Recordset = rstbeneficiario
   Case 2
   
   Set rs_aux17 = New ADODB.Recordset
    If rs_aux17.State = 1 Then rs_aux17.Close
    rs_aux17.Open "select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' AND estado_codigo <> 'ANL' order by beneficiario_denominacion asc", db, adOpenKeyset, adLockOptimistic, adCmdText
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

   If (grdlista.SelBookmarks.Count <> 0) Then
            grdlista.SelBookmarks.Remove 0
   End If
   If rstbeneficiario.RecordCount > 0 Then

   rstbeneficiario.Find "beneficiario_codigo = '" & dtc_buscar_ci.Text & "'", , , 1

   grdlista.SelBookmarks.Add (rstbeneficiario.Bookmark)
 
 Else
 sino = MsgBox("No se encontro a nadie con ese nombre", vbInformation, "Aviso")
 Call Carga_Beneficiario(1)
 dtc_buscar_desc.Text = ""
 End If
End Select
  
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

Public Sub abrirtabla()
    Set rs_Dependiente = New ADODB.Recordset
    If rs_Dependiente.State = 1 Then rs_Dependiente.Close
    rs_Dependiente.Open "select * from ro_Beneficiario_Dependiente where beneficiario_codigo = '" & Adolista.Recordset!beneficiario_codigo & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set AdoDependiente.Recordset = rs_Dependiente
    Set DtgDependiente.DataSource = AdoDependiente.Recordset
   
    Set rs_datos_educacionales = New ADODB.Recordset
    If rs_datos_educacionales.State = 1 Then rs_datos_educacionales.Close
    rs_datos_educacionales.Open "select * from ro_datos_educacionales where beneficiario_codigo = '" & Adolista.Recordset!beneficiario_codigo & "' order by fecha_inicio desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Educacionales.Recordset = rs_datos_educacionales
    Set DtgEducacionales.DataSource = Ado_Educacionales.Recordset
  
    Set rs_laborales = New ADODB.Recordset
    If rs_laborales.State = 1 Then rs_laborales.Close
    rs_laborales.Open "select * from ro_experiencia_laboral where beneficiario_codigo = '" & Adolista.Recordset!beneficiario_codigo & "' order by fecha_inicio desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Laborales.Recordset = rs_laborales
    Set DtgLaborales.DataSource = Ado_Laborales.Recordset
    
    Set rs_contrato = New Recordset
    If rs_contrato.State = 1 Then rs_contrato.Close
    rs_contrato.Open "select * from ro_contratos_personas where beneficiario_codigo = '" & Adolista.Recordset!beneficiario_codigo & "' order by fecha_inicio desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Contrato.Recordset = rs_contrato.DataSource
    Set DtG_Contrato.DataSource = Ado_Contrato.Recordset
    
    Set rs_liquidacion = New Recordset
    If rs_liquidacion.State = 1 Then rs_liquidacion.Close
    rs_liquidacion.Open "select * from ro_liquidaciones where beneficiario_codigo = '" & Adolista.Recordset!beneficiario_codigo & "' order by fecha_ingreso desc ", db, adOpenKeyset, adLockOptimistic
    Set AdoLiquidacion.Recordset = rs_liquidacion.DataSource
    Set DtgLiquidacion.DataSource = AdoLiquidacion.Recordset

End Sub

Private Sub Form_Resize()
'   '  Centrear titulo
'   With lblTitulo
'      .Left = (fraTitulo.Width - .Width) \ 2
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
  'carga    fc_tipo_beneficiario
    Set rst_ben = New ADODB.Recordset
    rst_ben.Open "SELECT * FROM gc_tipo_beneficiario ORDER BY tipoben_descripcion ", db, adOpenStatic
    Set AdoTip_ben.Recordset = rst_ben
    
    Set rs_Depto = New ADODB.Recordset
    rs_Depto.Open "select * from gc_Departamento", db, adOpenKeyset, adLockOptimistic
    Set Ado_Depto.Recordset = rs_Depto
    Dtc_depto.BoundText = Dtc_depto_cod.BoundText
    
    Set rs_Prov = New ADODB.Recordset
    rs_Prov.Open "select * from GC_Provincia", db, adOpenKeyset, adLockOptimistic
    Set Ado_prov.Recordset = rs_Prov
    Dtc_prov.BoundText = Dtc_prov_cod.BoundText
    
    Set rs_Muni = New ADODB.Recordset
    rs_Muni.Open "select * from gc_Municipio ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Muni.Recordset = rs_Muni
    Dtc_munic.BoundText = Dtc_munic_cod.BoundText
    
    Set rs_TipoDocId = New ADODB.Recordset
    rs_TipoDocId.Open "select * from gc_tipo_documento_id where estado_codigo ='APR' ", db, adOpenKeyset, adLockOptimistic
    Set Ado_TipoDocId.Recordset = rs_TipoDocId
    
'    Set rs_Depto3 = New ADODB.Recordset
'    rs_Depto3.Open "select * from gc_Departamento", DB, adOpenKeyset, adLockOptimistic
'    Set Ado_Depto3.Recordset = rs_Depto3
'    'Dtc_depto2.BoundText = Dtc_depto_cod2.BoundText
    
    Set rs_nivel_educacional = New ADODB.Recordset
    rs_nivel_educacional.Open "select * from rc_nivel_educacional ", db, adOpenKeyset, adLockOptimistic
    Set AdoNivelEducacional.Recordset = rs_nivel_educacional
    
    Set rs_tipoInstitucion = New ADODB.Recordset
    rs_tipoInstitucion.Open "select * from rc_tipo_institucion ", db, adOpenKeyset, adLockOptimistic
    Set Ado_TipoInstitucion.Recordset = rs_tipoInstitucion
    
   Set rs_beneficiario = New ADODB.Recordset
   If rs_beneficiario.State = 1 Then rs_beneficiario.Close
   rs_beneficiario.Open "select * from gc_Beneficiario WHERE hora_registro = 'SS' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_Benef_seguro.Recordset = rs_beneficiario
   
   Set rs_CTA_BCO = New ADODB.Recordset
   If rs_CTA_BCO.State = 1 Then rs_CTA_BCO.Close
   'rs_CTA_BCO.Open "select * from fv_cuenta_bco WHERE cta_codigo_tgn = '000'", db, adOpenKeyset, adLockOptimistic, adCmdText
   rs_CTA_BCO.Open "select * from fc_bancos ORDER BY bco_descripcion", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set AdoCta.Recordset = rs_CTA_BCO
   
    Set rs_ocupacion = New ADODB.Recordset
    rs_ocupacion.Open "select * from gc_ocupacion_profesion ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Ocupacion.Recordset = rs_ocupacion
    
   Set rs_beneficiario_Afp = New ADODB.Recordset
   If rs_beneficiario_Afp.State = 1 Then rs_beneficiario_Afp.Close
   rs_beneficiario_Afp.Open "select * from gc_Beneficiario WHERE hora_registro = 'AFP'", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_Benef_Afp.Recordset = rs_beneficiario_Afp
      
   Set rs_EstCivil = New ADODB.Recordset
   If rs_EstCivil.State = 1 Then rs_EstCivil.Close
   rs_EstCivil.Open "select * from rc_estado_civil ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set AdoEstCivil.Recordset = rs_EstCivil
   
   Set rs_pais = New ADODB.Recordset
   rs_pais.Open "SELECT * FROM gc_pais ORDER BY pais_descripcion ", db, adOpenStatic
   Set AdoPais.Recordset = rs_pais
   
   'UNIDAD EJECUTORA
   Set rs_datos1 = New ADODB.Recordset
   If rs_datos1.State = 1 Then rs_datos1.Close
   rs_datos1.Open "select * from gc_unidad_ejecutora ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_datos1.Recordset = rs_datos1
   'CARGOS
   Set rs_datos2 = New ADODB.Recordset
   If rs_datos2.State = 1 Then rs_datos2.Close
   rs_datos2.Open "select * from rc_cargos ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_datos2.Recordset = rs_datos2
   'PUESTOS
   Set rs_datos3 = New ADODB.Recordset
   If rs_datos3.State = 1 Then rs_datos3.Close
   rs_datos3.Open "select * from rc_puestos ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_datos3.Recordset = rs_datos3
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
End Sub

Private Sub Img_CTO_Click()
 If Ado_Contrato.Recordset!ARCHIVO = "Cargar_Archivo" Then
    MsgBox "No Existe el Archivo Asociado al Contrato, debe Cargarlo ...", vbExclamation, "Advertencia"
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SERVIDOR2" Then
        'e = ShellExecute(Img_CTO, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        'e = ShellExecute(Img_CTO, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ShellExecute(0, vbNullString, "c:\Archivo.PDF", vbNullString, vbNullString, vbNormalFocus)
'System.Diagnostics.Process.Start("c:\Archivo.PDF")
End Sub

'Private Sub Image1_Click()
'    Dim e As Long
'    If GlServidor <> GlMaquina Then      ' "-" Then
'        e = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Permiso-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    Else
'        e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Permiso-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    End If
'End Sub

Private Sub Img_CV_Click()
'    Dim e As Long
  If swnuevo <> "X" Then
    If Adolista.Recordset!archivo_hojavida = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "C_V"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SERVIDOR2" Then
         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\"
         ' e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      Else
         e = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "C_V"
          'If GlServidor <> GlMaquina Then      ' "-" Then
          If GlServidor = "SERVIDOR2" Then
            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\"
          Else
            e = NombreCarpeta
          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If
  End If
  'If GlServidor <> GlMaquina Then      ' "-" Then
  If GlServidor = "SERVIDOR2" Then
        'imag2 = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(adoLista.Recordset!ARCHIVO_HOJAVIDA), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(Adolista.Recordset!archivo_hojavida), vbNullString, vbNullString, vbNormalFocus)
  Else
        'imag2 = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(adoLista.Recordset!ARCHIVO_HOJAVIDA), vbNullString, vbNullString, SW_SHOWNORMAL)
        'Call ShellExecute(Me.hwnd, "Open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(adoLista.Recordset!ARCHIVO_HOJAVIDA), vbNullString, vbNullString, SW_SHOWNORMAL)
        'imag2 = ShellExecute(Me.hwnd, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(adoLista.Recordset!ARCHIVO_HOJAVIDA), "", "", 1)
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\HOJA_VIDA\" & Trim(Adolista.Recordset!archivo_hojavida), vbNullString, vbNullString, vbNormalFocus)
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
  If swnuevo <> "X" Then
    If Adolista.Recordset!archivo_respaldo = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "D_R"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SERVIDOR2" Then
            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
      Else
            e = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "D_R"
          'If GlServidor <> GlMaquina Then      ' "-" Then
          If GlServidor = "SERVIDOR2" Then
            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\"
          Else
            e = NombreCarpeta
          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If
  End If
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SERVIDOR2" Then
        'e = ShellExecute(Img_DocRespaldo, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(adoLista.Recordset!ARCHIVO_RESPALDO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(Adolista.Recordset!archivo_respaldo), vbNullString, vbNullString, vbNormalFocus)
    Else
        'e = ShellExecute(Img_DocRespaldo, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(adoLista.Recordset!ARCHIVO_RESPALDO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Adolista.Recordset!beneficiario_codigo) & "\DOCUMENTOS_RESPALDO\" & Trim(Adolista.Recordset!archivo_respaldo), vbNullString, vbNullString, vbNormalFocus)
    End If
End Sub

'Private Sub Img_Foto_Click()
'  If swnuevo <> "X" Then
'    If adoLista.Recordset!ARCHIVO_Foto = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "FOT"
'      'If GlServidor <> GlMaquina Then      ' "-" Then
'      If GlServidor = "SERVIDOR2" Then
'         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\"
'      Else
'         e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "FOT"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SERVIDOR2" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'
'    Dim ARCH_FOTO As String
'    'If GlServidor <> GlMaquina Then      ' "-" Then
'    If GlServidor = "SERVIDOR2" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" + Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(adoLista.Recordset("beneficiario_codigo")) + "\" + Trim(adoLista.Recordset!ARCHIVO_Foto)
'    Else
'        ARCH_FOTO = App.Path + "\" & Trim(GLCarpeta2) & "\" + Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(adoLista.Recordset("beneficiario_codigo")) + "\" + Trim(adoLista.Recordset!ARCHIVO_Foto)
'    End If
'    'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + adoLista.Recordset!beneficiario_codigo + "\" + adoLista.Recordset("beneficiario_codigo") + "-FOTO.JPG"
'    If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where beneficiario_codigo= '" & adoLista.Recordset("beneficiario_codigo") & "' ", "Foto", ARCH_FOTO) Then
'        MsgBox "Se cargo la Imagen Correctamente !!"
'    Else
'        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'    End If
'  End If
'End Sub

Private Sub ImgEvaluacion_Click()
'    Dim e As Long
'    If GlServidor <> GlMaquina Then      ' "-" Then
'        e = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\EVALUACIONES\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Evaluacion-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    Else
'        e = ShellExecute(Img_CV, "open", App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\EVALUACIONES\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Evaluacion-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    End If
End Sub

Private Sub ImgFiniquito_Click()
 If AdoLiquidacion.Recordset!ARCHIVO = "Cargar_Archivo" Then
    MsgBox "No Existe el Archivo Asociado a la Liquidación, debe Cargarlo ...", vbExclamation, "Advertencia"
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SERVIDOR2" Then
        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoLiquidacion.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(AdoLiquidacion.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        'e = ShellExecute(Img_CTO, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(AdoLiquidacion.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Adolista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(AdoLiquidacion.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub

Private Sub ImgMemo_Click()
'    Dim e As Long
'    If GlServidor <> GlMaquina Then      ' "-" Then
'        e = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\MEMORANDUMS\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Memo-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    Else
'        e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\MEMORANDUMS\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Memo-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    End If
End Sub

Private Sub ImgVacacion_Click()
'    Dim e As Long
'    If GlServidor <> GlMaquina Then      ' "-" Then
'        e = ShellExecute(Img_CV, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Vacacion-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    Else
'        e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(adoLista.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(adoLista.Recordset!beneficiario_beneficiario_iniciales) & "-Vacacion-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
'    End If
End Sub

Private Sub SSTab1_DblClick()
    If SSTab1.Tab = 0 Then
    End If
End Sub

Private Sub TDBNivelEdu_DropDownClose()
'    DtgEducacionales.Columns("nivel_educacional").Value = TDBNivelEdu.Columns("nivel_educacional").Value
'    'DtgEducacionales.Columns("descripcion").Value = TDBNivelEdu.Columns("descripcion").Value
End Sub

'Private Sub TDBtipoben_LostFocus()
'    If TxtTipo.Text = "6" Then
'        txtDenominacion.Enabled = True
''        Label2.Caption = "Empresa/Instit."
''        Frame2.Caption = "Datos del Representante Legal"
''        Label14.Caption = "Fecha de Creación"
''        LlbCargo.Caption = "Actividad Principal"
''        LblProf_Asoc.Caption = "Camara/Asociación a la que Pertenece"
''        Label13.Caption = "Nro. Registro"
'    Else
'        txtDenominacion.Enabled = False
''        Label2.Caption = "Denominación"
''        Frame2.Caption = "Datos de la Persona"
''        Label14.Caption = "Fecha de Nacimiento"
''        LlbCargo.Caption = "Cargo que Ocupa"
''        LblProf_Asoc.Caption = "Profesion u Ocupacion:"
''        Label13.Caption = "Nro. Empresa"
'    End If
'        DtgEducacionales.AllowAddNew = False
'        DtgEducacionales.AllowDelete = False
'        DtgEducacionales.AllowUpdate = False
'        DtgLaborales.AllowAddNew = False
'        DtgLaborales.AllowDelete = False
'        DtgLaborales.AllowUpdate = False
''        CmdAdd.Visible = False
''        CmdMod.Visible = False
''        CmdGraba.Visible = False
''        CmdAdd2.Visible = False
''        CmdMod2.Visible = False
''        CmdGraba2.Visible = False
''    If TxtTipo.Text = "1" Then
''        TxtRenca.Visible = False
''        'TxtRenca.BackColor =&H8000000B&
''        DTP_FechaExpira.Visible = False
''        Label13.Visible = False
''        Label10.Visible = False
''    Else
''        TxtRenca.Visible = True
''        'TxtRenca.BackColor =&H8000000B&
''        DTP_FechaExpira.Visible = True
''        Label13.Visible = True
''        Label10.Visible = True
''    End If
'End Sub

Private Sub TDBTipoInst_DropDownClose()
'    DtgLaborales.Columns("tipo_institucion").Value = TDBTipoInst.Columns("tipo_institucion").Value
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
    Call pDepto(DtcPaisCod.BoundText)
    Dtc_depto.Enabled = True
End Sub

Private Sub pDepto(CodDepto As String)
   Dim strConsultaM2 As String

   strConsultaM2 = "select * from gc_departamento where pais_codigo='" & CodDepto & "'"

   Set Dtc_depto_cod.RowSource = Nothing
   Set Dtc_depto_cod.RowSource = db.Execute(strConsultaM2, , adCmdText)
   Dtc_depto_cod.ReFill
   Dtc_depto_cod.BoundText = Empty

   Set Dtc_depto.RowSource = Nothing
   Set Dtc_depto.RowSource = db.Execute(strConsultaM2, , adCmdText)
   Dtc_depto.ReFill
   Dtc_depto.BoundText = Empty
End Sub

'Private Sub Text302_Change()
'    Text302.BackColor = &H80000014
'End Sub

Private Sub TxtNIT2_Click(Area As Integer)
    DtcRep_Nombres.BoundText = TxtNIT2.BoundText
    DtcRep_Paterno.BoundText = TxtNIT2.BoundText
    DtcRep_Materno.BoundText = TxtNIT2.BoundText
End Sub

Private Sub txtProfesion_Click(Area As Integer)
    Dtc_Ocup.BoundText = TxtProfesion.BoundText
End Sub


VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBeneficiario_Control 
   BackColor       =   &H00000000&
   Caption         =   "Control de Personal - File Funcionario"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15810
   Icon            =   "frmBeneficiario_Control.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15810
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fra_cabecera 
      BackColor       =   &H00404040&
      Height          =   975
      Left            =   120
      Picture         =   "frmBeneficiario_Control.frx":0A02
      ScaleHeight     =   915
      ScaleWidth      =   6075
      TabIndex        =   122
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton btnimprimir 
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
         Index           =   0
         Left            =   4560
         Picture         =   "frmBeneficiario_Control.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
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
         Left            =   3800
         Picture         =   "frmBeneficiario_Control.frx":6CFF1
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Busca un Registro"
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
         Height          =   735
         Left            =   5320
         Picture         =   "frmBeneficiario_Control.frx":6D5A9
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FICHA PERSONAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   15
         TabIndex        =   126
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADOS"
      ForeColor       =   &H00FFFFC0&
      Height          =   8295
      Left            =   120
      TabIndex        =   46
      Top             =   1200
      Width           =   6135
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
         TabIndex        =   48
         Top             =   7935
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
         Left            =   2160
         TabIndex        =   47
         Top             =   7920
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   7860
         Width           =   5865
         _ExtentX        =   10345
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
         Bindings        =   "frmBeneficiario_Control.frx":6D7B3
         Height          =   7215
         Left            =   60
         TabIndex        =   45
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   12726
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
         ColumnCount     =   6
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1395.213
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
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_buscar_desc 
         Bindings        =   "frmBeneficiario_Control.frx":6D7CB
         Height          =   315
         Left            =   840
         TabIndex        =   128
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
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
         Bindings        =   "frmBeneficiario_Control.frx":6D7E8
         DataField       =   "beneficiario_codigo"
         Height          =   315
         Left            =   4320
         TabIndex        =   129
         Top             =   240
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
      Begin VB.OLE OLE1 
         Height          =   495
         Left            =   1920
         TabIndex        =   135
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label52 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar..."
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   240
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9360
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   16510
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   64
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
      TabPicture(0)   =   "frmBeneficiario_Control.frx":6D805
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraGrabarCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDatos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraOpciones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "CONTROL ASISTENCIA"
      TabPicture(1)   =   "frmBeneficiario_Control.frx":6D821
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame18"
      Tab(1).Control(1)=   "Label44"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "PERMISOS-VACACIONES"
      TabPicture(2)   =   "frmBeneficiario_Control.frx":6D83D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(2)=   "Label45"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "MOVILIDAD PERSONAL"
      TabPicture(3)   =   "frmBeneficiario_Control.frx":6D859
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label46"
      Tab(3).Control(1)=   "Frame16"
      Tab(3).Control(2)=   "Frame17"
      Tab(3).ControlCount=   3
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H00404040&
         Height          =   735
         Left            =   0
         Picture         =   "frmBeneficiario_Control.frx":6D875
         ScaleHeight     =   675
         ScaleWidth      =   9195
         TabIndex        =   109
         Top             =   840
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Control.frx":D98A7
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Carga Foto de la Persona"
            Top             =   30
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton CmdVerDisco 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Docs."
            Enabled         =   0   'False
            Height          =   600
            Left            =   7440
            Picture         =   "frmBeneficiario_Control.frx":D9E31
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   60
            Width           =   740
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
            Picture         =   "frmBeneficiario_Control.frx":DA1B9
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Anula Registro Activo"
            Top             =   30
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
            Picture         =   "frmBeneficiario_Control.frx":DA905
            Style           =   1  'Graphical
            TabIndex        =   113
            ToolTipText     =   "Modifica Registro Activo"
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
            Picture         =   "frmBeneficiario_Control.frx":DB21A
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Visible         =   0   'False
            Width           =   1245
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
            Picture         =   "frmBeneficiario_Control.frx":DB9D9
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Aprueba Registro"
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton CmdDesapr 
            BackColor       =   &H0080C0FF&
            Caption         =   "Desapr"
            Height          =   600
            Left            =   2760
            Picture         =   "frmBeneficiario_Control.frx":DC20F
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "Aprueba Registro"
            Top             =   60
            Width           =   740
         End
      End
      Begin VB.PictureBox fraDatos 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   7815
         Left            =   0
         ScaleHeight     =   7755
         ScaleWidth      =   9195
         TabIndex        =   49
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
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   144
            Top             =   3840
            Width           =   1440
         End
         Begin VB.TextBox Text4 
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
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3960
            MaxLength       =   20
            TabIndex        =   143
            Text            =   "0"
            Top             =   3840
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   8640
            TabIndex        =   142
            Top             =   3135
            Width           =   375
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
            TabIndex        =   127
            Top             =   3840
            Width           =   1440
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
            TabIndex        =   66
            Top             =   5445
            Width           =   8835
            Begin MSDataListLib.DataCombo Dtc_prov_cod 
               Bindings        =   "frmBeneficiario_Control.frx":DC419
               DataField       =   "prov_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3000
               TabIndex        =   67
               Top             =   840
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
               Bindings        =   "frmBeneficiario_Control.frx":DC430
               DataField       =   "munic_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   7200
               TabIndex        =   68
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
               Bindings        =   "frmBeneficiario_Control.frx":DC447
               DataField       =   "prov_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1020
               TabIndex        =   69
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
               Bindings        =   "frmBeneficiario_Control.frx":DC45E
               DataField       =   "munic_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   5280
               TabIndex        =   70
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
               Bindings        =   "frmBeneficiario_Control.frx":DC475
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   7200
               TabIndex        =   71
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
               Bindings        =   "frmBeneficiario_Control.frx":DC48D
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   5280
               TabIndex        =   72
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
               Bindings        =   "frmBeneficiario_Control.frx":DC4A5
               DataField       =   "pais_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1400
               TabIndex        =   73
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
               Bindings        =   "frmBeneficiario_Control.frx":DC4BB
               DataField       =   "pais_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3000
               TabIndex        =   74
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
               Bindings        =   "frmBeneficiario_Control.frx":DC4D1
               DataField       =   "pais_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2040
               TabIndex        =   75
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
               TabIndex        =   77
               Top             =   915
               Width           =   5340
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
               TabIndex        =   76
               Top             =   375
               Width           =   5100
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00000000&
            ForeColor       =   &H00000040&
            Height          =   1680
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   9150
            Begin VB.PictureBox Img_Foto 
               AutoRedraw      =   -1  'True
               Height          =   1515
               Left            =   7095
               ScaleHeight     =   1455
               ScaleWidth      =   1815
               TabIndex        =   56
               Top             =   120
               Width           =   1875
               Begin VB.Image Image2 
                  Height          =   1455
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Documento de Identidad    Tipo Documento      Expedido.en "
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
               TabIndex        =   64
               Top             =   870
               Width           =   5340
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
               TabIndex        =   63
               Top             =   210
               Width           =   3315
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
               ForeColor       =   &H00FFFF80&
               Height          =   255
               Left            =   4965
               TabIndex        =   62
               Top             =   1425
               Width           =   1815
            End
            Begin VB.Label txtCodigo 
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "beneficiario_codigo"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   61
               Top             =   1140
               Width           =   2055
            End
            Begin VB.Label txtDenominacion 
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "beneficiario_denominacion"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   60
               Top             =   480
               Width           =   6375
            End
            Begin VB.Label Dtc_doc_id 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "tipodoc_codigo"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2865
               TabIndex        =   59
               Top             =   1140
               Width           =   855
            End
            Begin VB.Label DtcDepto3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "depto_sigla"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4275
               TabIndex        =   58
               Top             =   1140
               Width           =   1095
            End
            Begin VB.Label TxtNIT 
               BackColor       =   &H00404040&
               Caption         =   "-"
               DataField       =   "beneficiario_nit"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   5445
               TabIndex        =   57
               Top             =   1140
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00000000&
            Caption         =   "Aprobado"
            ForeColor       =   &H00FFFFFF&
            Height          =   600
            Left            =   6060
            TabIndex        =   53
            Top             =   360
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
               TabIndex        =   54
               Top             =   200
               Width           =   735
            End
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   7425
            TabIndex        =   52
            Top             =   2415
            Width           =   255
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   8620
            TabIndex        =   51
            Top             =   1935
            Width           =   255
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   4185
            TabIndex        =   50
            Top             =   1935
            Width           =   255
         End
         Begin MSComCtl2.DTPicker DTP_FechaNac 
            DataField       =   "beneficiario_fecha_nacimiento"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3840
            TabIndex        =   78
            Top             =   4960
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90963969
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTP_FechaExpira 
            DataField       =   "Fecha_expiracion"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4560
            TabIndex        =   79
            Top             =   1320
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90963969
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo TxtProfesion 
            Bindings        =   "frmBeneficiario_Control.frx":DC4E7
            DataField       =   "ocup_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   80
            Top             =   1920
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
            Bindings        =   "frmBeneficiario_Control.frx":DC503
            DataField       =   "ocup_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3720
            TabIndex        =   81
            Top             =   1920
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
         Begin MSDataListLib.DataCombo TDBtipoben 
            Bindings        =   "frmBeneficiario_Control.frx":DC51F
            DataField       =   "tipoben_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   82
            Top             =   1320
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
            Bindings        =   "frmBeneficiario_Control.frx":DC538
            DataField       =   "tipoben_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3720
            TabIndex        =   83
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
            Bindings        =   "frmBeneficiario_Control.frx":DC551
            DataField       =   "estado_civil_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   4960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "estado_civil_descripcion"
            BoundColumn     =   "estado_civil_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo DtcEstCiv 
            Bindings        =   "frmBeneficiario_Control.frx":DC56B
            DataField       =   "estado_civil_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   85
            Top             =   4980
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
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "frmBeneficiario_Control.frx":DC585
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2160
            TabIndex        =   87
            Top             =   2400
            Width           =   5535
            _ExtentX        =   9763
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
            Bindings        =   "frmBeneficiario_Control.frx":DC59E
            DataField       =   "puesto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4440
            TabIndex        =   88
            Top             =   1920
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "puesto_descripcion"
            BoundColumn     =   "puesto_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "frmBeneficiario_Control.frx":DC5B9
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7680
            TabIndex        =   89
            Top             =   2400
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
            Bindings        =   "frmBeneficiario_Control.frx":DC5D2
            DataField       =   "genero_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6540
            TabIndex        =   90
            Top             =   4960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "genero_descripcion"
            BoundColumn     =   "genero_codigo"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "frmBeneficiario_Control.frx":DC5EB
            DataField       =   "genero_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5760
            TabIndex        =   91
            Top             =   4920
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
            Left            =   5760
            TabIndex        =   136
            Top             =   3840
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90963969
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin VB.TextBox TxtRenca 
            DataField       =   "reg_profesional"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   65
            Top             =   1320
            Visible         =   0   'False
            Width           =   2280
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "frmBeneficiario_Control.frx":DC604
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   137
            Top             =   3120
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
            Bindings        =   "frmBeneficiario_Control.frx":DC61D
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3840
            TabIndex        =   138
            Top             =   2775
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
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "frmBeneficiario_Control.frx":DC636
            DataField       =   "puesto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8160
            TabIndex        =   86
            Top             =   1920
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
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "frmBeneficiario_Control.frx":DC651
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4920
            TabIndex        =   140
            Top             =   3120
            Width           =   4095
            _ExtentX        =   7223
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
            Bindings        =   "frmBeneficiario_Control.frx":DC66A
            DataField       =   "unidad_codigo_pla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8040
            TabIndex        =   141
            Top             =   2880
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
         Begin VB.Label Label7 
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
            TabIndex        =   139
            Top             =   2880
            Width           =   2625
         End
         Begin VB.Label Label10 
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
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2520
            TabIndex        =   93
            Top             =   4300
            Width           =   2130
         End
         Begin VB.Label Label13 
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
            TabIndex        =   102
            Top             =   6885
            Width           =   1485
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   101
            Top             =   1665
            Width           =   7035
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Sueldo Básico             Refrigerio/Otro           Bono Antigüedad        Fecha de Ingreso        Correl.Planilla"
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
            TabIndex        =   100
            Top             =   3600
            Width           =   8820
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Unidad Oganizacional"
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
            TabIndex        =   99
            Top             =   2385
            Width           =   1995
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Estado Civil                                                               Fecha Nacimiento                    Género"
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
            TabIndex        =   98
            Top             =   4695
            Width           =   7110
         End
         Begin VB.Label TxtDireccion 
            BackColor       =   &H00404040&
            Caption         =   "-"
            DataField       =   "beneficiario_domicilio_legal"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   97
            Top             =   7200
            Width           =   8775
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
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7800
            TabIndex        =   96
            Top             =   3840
            Width           =   1095
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
            TabIndex        =   95
            Top             =   1200
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.Label Label1 
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
            TabIndex        =   94
            Top             =   4320
            Width           =   6585
         End
         Begin VB.Label Label11 
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
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6720
            TabIndex        =   92
            Top             =   4300
            Width           =   2130
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   3975
         Left            =   -75000
         TabIndex        =   36
         Top             =   5400
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Control.frx":DC684
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Nuevo Registro"
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
            Picture         =   "frmBeneficiario_Control.frx":DCC0E
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdElim3 
            BackColor       =   &H00C0C0FF&
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
            Picture         =   "frmBeneficiario_Control.frx":DD198
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
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
            Picture         =   "frmBeneficiario_Control.frx":DDB9A
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame19 
            BackColor       =   &H00000000&
            Height          =   650
            Left            =   8280
            TabIndex        =   37
            Top             =   140
            Width           =   615
            Begin VB.Image Img_03 
               Height          =   540
               Left            =   0
               Picture         =   "frmBeneficiario_Control.frx":DE124
               Top             =   80
               Width           =   555
            End
         End
         Begin MSDataGridLib.DataGrid DtgPermiso 
            Bindings        =   "frmBeneficiario_Control.frx":DE4AC
            Height          =   2625
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   4630
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
            Top             =   3480
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6840
            TabIndex        =   44
            Top             =   240
            Width           =   1305
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
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   7095
            TabIndex        =   43
            Top             =   555
            Width           =   1050
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   8895
         Left            =   -75000
         TabIndex        =   30
         Top             =   840
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Control.frx":DE4C5
            Style           =   1  'Graphical
            TabIndex        =   145
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cbo_mes 
            Height          =   315
            ItemData        =   "frmBeneficiario_Control.frx":DEA4F
            Left            =   5880
            List            =   "frmBeneficiario_Control.frx":DEA7A
            TabIndex        =   106
            Text            =   "MES"
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox cbo_gestion 
            Height          =   315
            ItemData        =   "frmBeneficiario_Control.frx":DEAE9
            Left            =   7680
            List            =   "frmBeneficiario_Control.frx":DEB0E
            TabIndex        =   105
            Text            =   "GESTION"
            Top             =   240
            Width           =   1335
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
            Picture         =   "frmBeneficiario_Control.frx":DEB54
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Nuevo Registro"
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
            Picture         =   "frmBeneficiario_Control.frx":DF0DE
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Modifica Registro Activo"
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
            Picture         =   "frmBeneficiario_Control.frx":DF668
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
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
            Picture         =   "frmBeneficiario_Control.frx":E006A
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DtgAsistencia 
            Bindings        =   "frmBeneficiario_Control.frx":E05F4
            Height          =   7185
            Left            =   180
            TabIndex        =   31
            Top             =   840
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   12674
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
         Begin MSAdodcLib.Adodc AdoAsistencia 
            Height          =   330
            Left            =   120
            Top             =   8040
            Width           =   9015
            _ExtentX        =   15901
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
         Begin VB.TextBox txt_mes 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFF00&
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   107
            Text            =   "0"
            Top             =   600
            Visible         =   0   'False
            Width           =   630
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
            ForeColor       =   &H00FFFF80&
            Height          =   240
            Left            =   5235
            TabIndex        =   108
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   4335
         Left            =   -75000
         TabIndex        =   21
         Top             =   5040
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Control.frx":E0610
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00C0C0C0&
            Height          =   690
            Left            =   8360
            TabIndex        =   28
            Top             =   100
            Width           =   615
            Begin VB.Image ImgFiniquito 
               Height          =   540
               Left            =   20
               Picture         =   "frmBeneficiario_Control.frx":E0B9A
               Top             =   100
               Width           =   555
            End
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
            Picture         =   "frmBeneficiario_Control.frx":E0F22
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Nuevo Registro"
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
            Picture         =   "frmBeneficiario_Control.frx":E14AC
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Modifica Registro Activo"
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
            Picture         =   "frmBeneficiario_Control.frx":E1A36
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Aprueba Registro"
            Top             =   240
            Width           =   855
         End
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
            Picture         =   "frmBeneficiario_Control.frx":E1FC0
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DtgMovilidad 
            Bindings        =   "frmBeneficiario_Control.frx":E29C2
            Height          =   3015
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5318
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12640511
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
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
            ColumnCount     =   8
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
               DataField       =   "unidad_anterior"
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
               DataField       =   "puesto_anterior"
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
               DataField       =   "unidad_codigo"
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
               DataField       =   "puesto_codigo"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1544.882
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   555.024
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoMovilidad 
            Height          =   330
            Left            =   120
            Top             =   3840
            Width           =   9015
            _ExtentX        =   15901
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6855
            TabIndex        =   27
            Top             =   240
            Width           =   1350
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
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   6660
            TabIndex        =   26
            Top             =   540
            Width           =   1560
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   4215
         Left            =   -75000
         TabIndex        =   12
         Top             =   840
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Control.frx":E29DF
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   690
            Left            =   8360
            TabIndex        =   20
            Top             =   100
            Width           =   615
            Begin VB.Image Img_CTO 
               Height          =   540
               Left            =   20
               Picture         =   "frmBeneficiario_Control.frx":E2F69
               Top             =   100
               Width           =   555
            End
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
            Picture         =   "frmBeneficiario_Control.frx":E32F1
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Nuevo Registro"
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
            Picture         =   "frmBeneficiario_Control.frx":E387B
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Modifica Registro Activo"
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
            Picture         =   "frmBeneficiario_Control.frx":E3E05
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
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
            Picture         =   "frmBeneficiario_Control.frx":E438F
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DtG_Memo 
            Bindings        =   "frmBeneficiario_Control.frx":E4D91
            Height          =   2865
            Left            =   165
            TabIndex        =   13
            Top             =   840
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   5054
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
            ColumnCount     =   7
            BeginProperty Column00 
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
            BeginProperty Column01 
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
            BeginProperty Column02 
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
            BeginProperty Column03 
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
            BeginProperty Column04 
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
            BeginProperty Column05 
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
            BeginProperty Column06 
               DataField       =   "estado_codigo"
               Caption         =   "Apr."
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
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   2940.095
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   434.835
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_Memo 
            Height          =   330
            Left            =   120
            Top             =   3720
            Width           =   9015
            _ExtentX        =   15901
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   7185
            TabIndex        =   19
            Top             =   240
            Width           =   1080
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
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   6720
            TabIndex        =   18
            Top             =   540
            Width           =   1560
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   4575
         Left            =   -75000
         TabIndex        =   4
         Top             =   840
         Width           =   9255
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
            TabIndex        =   133
            ToolTipText     =   "Nuevo Registro"
            Top             =   240
            Width           =   1455
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
            Picture         =   "frmBeneficiario_Control.frx":E4DAC
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00000000&
            Height          =   650
            Left            =   8360
            TabIndex        =   11
            Top             =   140
            Width           =   615
            Begin VB.Image Img_CV 
               Height          =   540
               Left            =   20
               Picture         =   "frmBeneficiario_Control.frx":E5336
               Top             =   60
               Width           =   555
            End
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
            Picture         =   "frmBeneficiario_Control.frx":E56BE
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Modifica Registro Activo"
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
            Picture         =   "frmBeneficiario_Control.frx":E5C48
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Nuevo Registro"
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
            Picture         =   "frmBeneficiario_Control.frx":E61D2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Anula Registro Activo"
            Top             =   240
            Width           =   855
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
            Picture         =   "frmBeneficiario_Control.frx":E6BD4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Aprueba Registro Activo"
            Top             =   240
            Width           =   855
         End
         Begin MSAdodcLib.Adodc Ado_VacacionesProg 
            Height          =   375
            Left            =   120
            Top             =   3720
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
            Bindings        =   "frmBeneficiario_Control.frx":E715E
            Height          =   2895
            Left            =   120
            TabIndex        =   132
            Top             =   840
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5106
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   4080
            Visible         =   0   'False
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6360
            TabIndex        =   10
            Top             =   240
            Width           =   1890
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
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   6660
            TabIndex        =   9
            Top             =   555
            Width           =   1605
         End
      End
      Begin VB.PictureBox FraGrabarCancelar 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         Picture         =   "frmBeneficiario_Control.frx":E717E
         ScaleHeight     =   675
         ScaleWidth      =   9195
         TabIndex        =   117
         Top             =   840
         Width           =   9255
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
            Picture         =   "frmBeneficiario_Control.frx":1531B0
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   1485
         End
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
            Picture         =   "frmBeneficiario_Control.frx":153A9C
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   30
            Width           =   1365
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
            TabIndex        =   120
            Top             =   300
            Visible         =   0   'False
            Width           =   525
         End
      End
      Begin VB.Label Label15 
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
         Left            =   0
         TabIndex        =   121
         Top             =   360
         Width           =   9255
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
         Left            =   -75000
         TabIndex        =   3
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IV. - PERMISOS - VACACIONES "
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
         TabIndex        =   2
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IV. - CONTROL DE ASISTENCIA "
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
         TabIndex        =   1
         Top             =   360
         Width           =   9255
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
   Begin MSAdodcLib.Adodc Ado_Comunid2 
      Height          =   330
      Left            =   6480
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
      Caption         =   "Ado_Comunid"
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
   Begin Crystal.CrystalReport CR03 
      Left            =   15120
      Top             =   9960
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
   Begin Crystal.CrystalReport CR04 
      Left            =   15600
      Top             =   4680
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
End
Attribute VB_Name = "frmBeneficiario_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento de Beneficiarios
Option Explicit
Dim rstbeneficiario As New ADODB.Recordset
Dim rst_ben, rsNada As New ADODB.Recordset

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


Dim rstdestino As New ADODB.Recordset

Dim permisos, totalminutos As Integer
Dim calretrasos As Double

Dim CAMPOS As ADODB.Field
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
            rs_aux3!Fecha_Registro = Date
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
      txtCodigo = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
      txtDenominacion = Empty
      Exit Sub
  End If
  
   'BtnModificar.Enabled = True
   'BtnEliminar.Enabled = True
  If Ado_datos.Recordset.RecordCount > 0 Then
    Select Case Ado_datos.Recordset.EditMode
      Case adEditInProgress
        Frame2.Enabled = False            'Verif. Nombre Proveedor JQA NOV-2009
      Case adEditNone
        Set Img_Foto = Leer_Imagen(db, "Select Foto From rv_personal_contratado Where beneficiario_codigo= '" & Ado_datos.Recordset!beneficiario_codigo & "' ", "Foto")
        Image2 = Img_Foto
        CmdFoto.Visible = True
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
        If pRecordset("estado_codigo") = "REG" Then
            BtnAprobar.Visible = True
            CmdDesapr.Visible = False
        Else
            BtnAprobar.Visible = False
            CmdDesapr.Visible = True
        End If
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
    If Ado_datos.Recordset("estado_codigo") = "APR" Then
        lblActivo.ForeColor = &H8000&
    Else
        lblActivo.ForeColor = &HC0&
    End If
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

Private Sub BtnGrabar_Click()
    V_TIPO = Trim(TxtTipo.Text)
    V_TDOC = Trim(Dtc_doc_id)
'On Error GoTo errorAceptar
   With Ado_datos
     If swnuevo = 1 Then
       CORREL = 0
'       DE.dbo_fc_correl_ben CORREL
       Set rstbeneaux = New ADODB.Recordset
       'SQL_FOR = "select * from rv_personal_contratado where beneficiario_codigo= '" & TxtCodigo.Text & "' OR beneficiario_codigo= '" & txtCodigo2.Text & "' "
       SQL_FOR = "select * from Gc_beneficiario where beneficiario_codigo= '" & txtCodigo & "'  "
       rstbeneaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
       'If rstbeneaux.RecordCount > 0 And txtCodigo.Enabled Then
       If rstbeneaux.RecordCount > 0 Then
                SW = True
                MsgBox " CODIGO DUPLICADO"
'                TxtCodigo.SetFocus
                Exit Sub
       End If
     End If
       If TxtTipo < "20" Then
            If Trim(txtCodigo) = "" Then
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
        If txtCodigo.Enabled And swnuevo = 1 Then
            .Recordset("beneficiario_codigo") = Trim(txtCodigo)
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
            .Recordset("depto_codigo") = IIf(Dtc_depto_cod.Text = "", "-", Dtc_depto_cod.Text)
            .Recordset("prov_codigo") = IIf(Dtc_prov_cod.Text = "", "-", Dtc_prov_cod.Text)
            .Recordset("munic_codigo") = IIf(Dtc_munic_cod.Text = "", "-", Dtc_munic_cod.Text)
            .Recordset("estado_civil_codigo") = DtcEstCiv.Text
            .Recordset("unidad_codigo_pla") = dtc_codigo2.Text
            .Recordset("beneficiario_haber_mensual") = txt_sueldo.Text
            .Recordset("beneficiario_otro_mensual") = txt_otro.Text
            .Recordset("fecha_ingreso") = DTPicker2.Value
'            Dim a As String, b As String, C As String
'            a = Left(Text1.Text, 1)
'            b = Left(Text2.Text, 1)
'            C = Left(Text3.Text, 1)
'            .Recordset("beneficiario_beneficiario_iniciales") = Trim(LblInicial.Caption)
'            RUTA1 = "PERSONAL" + "\" + Trim(LblInicial) + "-" + Trim(txtCodigo)
'            MsgBox RUTA1
'            MkDir RUTA1
        End If
            .Recordset("usr_codigo").Value = glusuario 'frmLogin.txtUserName.Text
            .Recordset("fecha_registro").Value = Date
            .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
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
       
         SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = True
       
       
       
   FraGrabarCancelar.Visible = False
   FraNavega.Enabled = True
   fraDatos.Enabled = False
'   FraSS_SS.Enabled = False
''''''   CmdAdd1.Visible = False
''''''   CmdMod1.Visible = False
''''''   CmdElim1.Visible = False
''''''   CmdApr1.Visible = False
''''''   CmdAdd2.Visible = False
''''''   CmdMod2.Visible = False
''''''   CmdElim2.Visible = False
''''''   CmdApr2.Visible = False
''''''   CmdAdd3.Visible = False
''''''   CmdMod3.Visible = False
''''''   CmdElim3.Visible = False
''''''   CmdApr3.Visible = False
''''''   CmdAdd4.Visible = False
''''''   CmdMod4.Visible = False
''''''   CmdElim4.Visible = False
''''''   CmdApr4.Visible = False
''''''   CmdAdd5.Visible = False
''''''   CmdMod5.Visible = False
''''''   CmdElim5.Visible = False
''''''   CmdApr5.Visible = False
'   CmdAdd6.Visible = False
'   CmdMod6.Visible = False
'   CmdElim6.Visible = False
'   CmdApr6.Visible = False

   Call Carga_Recor
   'Call Carga_Beneficiario
'De.dbo_alGraba_rc_personal Accion, CORREL, txtCodigo.Text, Text1.Text, Text2.Text, Text3.Text, "2002"
 Exit Sub
errorAceptar:
   Call pErrorRst(db.Errors)
   Ado_datos.Recordset.CancelUpdate
   'db.RollbackTrans
End Sub

Private Sub BtnImprimir_Click(Index As Integer)
Dim iResult As Integer
     CrystalReport1.WindowShowPrintSetupBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\Presupuesto\beneficiarios\crybeneficiario.rpt"
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\Generales\crybeneficiario.rpt"
    CrystalReport1.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_beneficiario_Personal_empresa.rpt"
  iResult = CrystalReport1.PrintReport
  If iResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

CrystalReport1.WindowState = crptMaximized

'    repbeneficiario.Show
    '   rptModalidadSeleccion.Show vbModal
End Sub

Private Sub BtnImprimir1_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR02.ReportFileName = App.Path & "\Reportes\clasificadores\rr_vacaciones_programadas.rpt"
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

Private Sub cbo_gestion_Click()
Call filtrar_asistencia(txt_mes.Text, cbo_gestion.Text)
End Sub

Private Sub cbo_mes_Click()
txt_mes.Text = cbo_mes.ListIndex
txt_mes.Text = Val(txt_mes.Text) + 1
Call filtrar_asistencia(txt_mes.Text, cbo_gestion.Text)
End Sub

Private Sub cmdAdd2_Click()
   If Ado_datos.Recordset.RecordCount > 0 Then
       marca1 = Ado_datos.Recordset.Bookmark
       frm_ao_Vacacion_Prog.txtSW = "ADD"
       frm_ao_Vacacion_Prog.txtBenef = Ado_datos.Recordset!beneficiario_codigo
       frm_ao_Vacacion_Prog.txtEstado = "REG"
       frm_ao_Vacacion_Prog.TxtGestion.Text = Year(Date)
       'Ado_VacacionesProg.Recordset.AddNew
'        frm_ao_Vacacion_Prog.lblbien(1).Visible = True
'       frm_ao_Vacacion_Prog.Txt02.Visible = True
         frm_ao_Vacacion_Prog.sel = 1
       frm_ao_Vacacion_Prog.Show vbModal
     
       Call abrirtabla
       'Ado_VacacionesProg.Refresh
   Else
       MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

End Sub

Private Sub cmdAdd3_Click()
   If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ao_Permisos_js.txtSW = "ADD"
        frm_ao_Permisos_js.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ao_Permisos_js.txtEstado = "REG"
        'AdoPermiso.Recordset.AddNew
        frm_ao_Permisos_js.TxtGestion = Year(Date)
        frm_ao_Permisos_js.Show vbModal
        
        Call abrirtabla
   Else
        MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
End Sub

Private Sub CmdAdd4_Click()
   If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ao_memoranda.txtSW = "ADD"
        frm_ao_memoranda.txtBenef = Ado_datos.Recordset!beneficiario_codigo
        frm_ao_memoranda.TxtInicial = Ado_datos.Recordset!beneficiario_iniciales
        frm_ao_memoranda.txtEstado = "REG"
'        Ado_Memo.Recordset.AddNew
        frm_ao_memoranda.Show vbModal
        Call abrirtabla
        'Ado_Memo.Refresh
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
   Exit Sub
AddErr:
  MsgBox Err.Description

End Sub

Private Sub CmdAdd1_Click()
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
End Sub

Private Sub CmdAdd5_Click()
   If Ado_datos.Recordset.RecordCount > 0 Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ro_movilidad_personal.txtSW = "ADD"
        frm_ro_movilidad_personal.txtBenef.Text = Ado_datos.Recordset!beneficiario_codigo
        frm_ro_movilidad_personal.TxtInicial = Ado_datos.Recordset!beneficiario_iniciales
        frm_ro_movilidad_personal.TxtAprob = "REG"
         frm_ro_movilidad_personal.DtcPryDes = Ado_datos.Recordset("puesto_descripcion")
         frm_ro_movilidad_personal.DTPFelaboracion.Value = Date
         
        'AdoMovilidad.Recordset.AddNew
        frm_ro_movilidad_personal.Show vbModal
        Call abrirtabla
        'AdoMovilidad.Refresh
        
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If
End Sub

Private Sub CmdApr3_Click()
 If AdoPermiso.Recordset("estado_codigo") = "REG" Then '
 If AdoPermiso.Recordset("TipoPermiso") = "VC" Or AdoPermiso.Recordset("TipoPermiso") = "VP" Then
 Dim DIAS, DIASP, TOTALD As Integer
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
         ''''''Call opciones
         
         
          Dim rs_datos4 As New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
     rs_datos4.Open "select * from ro_pagos_cronograma_Detalle where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic
     If rs_datos4.RecordCount <> 0 Then
        
         If rs_aux9.State = 1 Then rs_aux9.Close
            rs_aux9.Open "select sum(AtrasoMin1) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '" & RTrim(LTrim(AdoPermiso.Recordset!beneficiario_codigo)) & "' AND ges_gestion = '" & RTrim(LTrim(AdoPermiso.Recordset!ges_gestion)) & "' and Mes_control = '" & RTrim(LTrim(Month(AdoPermiso.Recordset!Fecha_control))) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        
        If rs_aux14.State = 1 Then rs_aux14.Close
          rs_aux14.Open "select sum(total_minuto) as PermisoMes from ro_permisos where beneficiario_codigo = '" & RTrim(LTrim(AdoPermiso.Recordset!beneficiario_codigo)) & "' AND ges_gestion = '" & RTrim(LTrim(AdoPermiso.Recordset!ges_gestion)) & "' AND Mes_control = '" & AdoPermiso.Recordset!mes_control & "'  AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic, adCmdText
            If rs_aux14!PermisoMes <> "NULL" Then
                permisos = rs_aux14!PermisoMes
            Else
                permisos = "0"
            End If
       
        If rs_aux9!TardanzaMes <> "NULL" Then
             totalminutos = rs_aux9!TardanzaMes - permisos
                If totalminutos >= 45 And rs_aux9!TardanzaMes <= 60 Then
                   calretrasos = ((rs_datos4!sueldo_basico / 30) / 2)
                Else
                    If totalminutos >= 61 And rs_aux9!TardanzaMes <= 75 Then
                        calretrasos = (rs_datos4!sueldo_basico / 30)
                    Else
                        If totalminutos >= 76 And rs_aux9!TardanzaMes <= 105 Then
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
          rs_aux10.Open "select sum(monto) as montomes from ro_memorandas where beneficiario_codigo = '" & RTrim(LTrim(AdoPermiso.Recordset!beneficiario_codigo)) & "' AND ges_gestion = '" & RTrim(LTrim(AdoPermiso.Recordset!ges_gestion)) & "' AND mes_descuento = '" & AdoPermiso.Recordset!mes_control & "'  AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic, adCmdText
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
           
     
      db.Execute "update ro_pagos_cronograma_Detalle set otros_dsctos = " & total & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "'"
      db.Execute "update ro_pagos_cronograma_Detalle set total_dsctos = " & total2 & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "'"
      total = 0
      total = rs_datos4!total_ganado - total2
      db.Execute "update ro_pagos_cronograma_Detalle set liquido_pagable_bs = " & total & ", liquido_pagable_us = " & (total2 / GlTipoCambioOficial) & "where ges_gestion = '" & AdoPermiso.Recordset!ges_gestion & "' AND mes_grupo = " & Month(AdoPermiso.Recordset!Fecha_control) & " AND beneficiario_codigo = '" & AdoPermiso.Recordset!beneficiario_codigo & "'"
      
      
      
         
      End If
      End If
Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
End If

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
            Ado_Memo.Recordset!Fecha_Registro = Date
            Ado_Memo.Recordset!usr_codigo = glusuario
            Ado_Memo.Recordset.Update
            
   If rs_datos3.State = 1 Then rs_datos3.Close
   rs_datos3.Open "select * from rc_tipo_memoranda where tipo_memo = '" & Ado_Memo.Recordset("tipo_memo") & "'", db, adOpenKeyset, adLockOptimistic
   
   If rs_datos3!estado_baja = "S" Then
   db.Execute "update ro_personal_contratado set fecha_expiracion = '" & Ado_Memo.Recordset!fecha_aprobacion & "' WHERE beneficiario_codigo = '" & Ado_Memo.Recordset("beneficiario_codigo") & "'"
   'db.Execute "update ro_personal_contratado set estado_codigo = 'ANL' WHERE beneficiario_codigo = '" & Ado_Memo.Recordset("beneficiario_codigo") & "'"
   Call Carga_Beneficiario(1)
   End If

    If Ado_Memo.Recordset("tipo_memo") = "SAD" Then
    total = 0
    total2 = 0
    VAR_MES = Month(Ado_Memo.Recordset!fecha_aprobacion)
    Dim rs_datos4 As New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
     rs_datos4.Open "select * from ro_pagos_cronograma_Detalle where ges_gestion = '" & Ado_Memo.Recordset!ges_gestion & "' AND mes_grupo = " & VAR_MES & " AND beneficiario_codigo = '" & Ado_Memo.Recordset!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic
     If rs_datos4.RecordCount <> 0 Then
     
     If Ado_Memo.Recordset("monto") > 0 Then
     total = rs_datos4!otros_dsctos + Ado_Memo.Recordset("monto")
'     rs_datos!otros_dsctos = total
'     rs_datos!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!rciva + rs_datos2!otros_dsctos
     End If

     If Ado_Memo.Recordset("dias") > 0 Then
     If rs_datos1.State = 1 Then rs_datos1.Close
     rs_datos1.Open "select * from ro_personal_contratado where beneficiario_codigo = '" & Ado_Memo.Recordset("beneficiario_codigo") & "'", db, adOpenKeyset, adLockOptimistic
     total = total + ((rs_datos1!beneficiario_haber_mensual / 30) * Ado_Memo.Recordset("dias"))
     total = total + rs_datos4!otros_dsctos
'     rs_datos!otros_dsctos = total
'     rs_datos!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!rciva + rs_datos2!otros_dsctos
     End If

     If total > 0 Then
     
     
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

End Sub

Private Sub CmdApr2_Click()
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
End Sub

Private Sub CmdApr5_Click()
 On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de APROBAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If AdoMovilidad.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        AdoMovilidad.Recordset("estado_codigo") = "APR"
        AdoMovilidad.Recordset("fecha_registro") = Date
        AdoMovilidad.Recordset("usr_codigo") = glusuario
        AdoMovilidad.Recordset.Update
        db.Execute "UPDATE ro_personal_contratado SET ro_personal_contratado = rs_movilidad!puesto_codigo where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'"
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdDesapr_Click()
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

End Sub

Private Sub CmdElim1_Click()
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
End Sub

Private Sub CmdElim2_Click()
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

End Sub

Private Sub CmdElim3_Click()
   sino = MsgBox("Está Seguro de BORRAR físicamente el Registro elegido ? ", vbYesNo + vbQuestion, "Atención")
   'If AdoPermiso.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        db.Execute "delete ro_permisos where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' and Correl = " & AdoPermiso.Recordset!CORREL & " "
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
   If ExisteReg(" ges_gestion = '" & Ado_Memo.Recordset!ges_gestion & "' AND mes_grupo = " & Ado_Memo.Recordset!mes_descuento & " AND beneficiario_codigo = " & Ado_Memo.Recordset!beneficiario_codigo, "ro_pagos_cronograma_Detalle") Then
      sino = MsgBox("No se puede ELIMINAR porque ya fue Procesado en la planilla. Desea marcar como ERRADO ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_contrato!estado_codigo = "ERR"
         rs_contrato!Fecha_Registro = Date
         rs_contrato!usr_codigo = glusuario
         rs_contrato.UpdateBatch adAffectAll
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
     If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_VacacionesProg.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ao_Vacacion_Prog.txtSW = "MOD"
        frm_ao_Vacacion_Prog.sel = 1
        frm_ao_Vacacion_Prog.txtBenef = Ado_datos.Recordset!beneficiario_codigo
         frm_ao_Vacacion_Prog.TxtGestion.Text = Ado_VacacionesProg.Recordset!ges_gestion
        
        frm_ao_Vacacion_Prog.Txt01.Text = Ado_VacacionesProg.Recordset!mes_control
        frm_ao_Vacacion_Prog.Txt02.Text = Ado_VacacionesProg.Recordset!dias_Programados
        frm_ao_Vacacion_Prog.txt03.Value = Ado_VacacionesProg.Recordset!fecha_ini_Prog
        frm_ao_Vacacion_Prog.txt04.Value = Ado_VacacionesProg.Recordset!fecha_fin_Prog
        frm_ao_Vacacion_Prog.DTPFec_Inicio.Value = Ado_VacacionesProg.Recordset!Fecha_Registro
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
End Sub

Private Sub CmdMod3_Click()
 If AdoPermiso.Recordset!estado_codigo = "REG" Then
    marca1 = Ado_datos.Recordset.Bookmark
    frm_ao_Permisos.txtSW = "MOD"
    frm_ao_Permisos.txtBenef = Ado_datos.Recordset!beneficiario_codigo
    'frm_ao_Permisos.TxtInicial = Ado_datos.Recordset!beneficiario_beneficiario_iniciales
    frm_ao_Permisos.lblARCH.Caption = AdoPermiso.Recordset!ARCHIVO
    frm_ao_Permisos.Dtc_Par = AdoPermiso.Recordset!TipoPermiso
    frm_ao_Permisos.dt_fechasolicitusper = AdoPermiso.Recordset!Fecha_control
    frm_ao_Permisos.cmb_mescontrol = AdoPermiso.Recordset!mes_control
'    frm_ao_Permisos.txt02 = AdoPermiso.Recordset!dia_control
    frm_ao_Permisos.dt_fechadesde = AdoPermiso.Recordset!FechaDesde
    frm_ao_Permisos.dt_fechahasta = AdoPermiso.Recordset!FechaHasta
    frm_ao_Permisos.dt_fechareincorporacion = AdoPermiso.Recordset!fecha_reincorporacion
    frm_ao_Permisos.hr_horadesde = AdoPermiso.Recordset!horadesde
    frm_ao_Permisos.hr_horahasta = AdoPermiso.Recordset!HoraHasta
    frm_ao_Permisos.hr_horareincorporacion = AdoPermiso.Recordset!Hora_reincorporacion
    frm_ao_Permisos.TxtGestion = AdoPermiso.Recordset!ges_gestion
    frm_ao_Permisos.txt_nrodias = AdoPermiso.Recordset!dias_permiso
    frm_ao_Permisos.txt_nrohoras = AdoPermiso.Recordset!horas_permiso
    frm_ao_Permisos.txt_nrominutos = AdoPermiso.Recordset!minutos_permiso
    'frm_ao_Permisos.Dtc_ParDes = AdoPermiso.Recordset!nomb_pariente
    frm_ao_Permisos.txtEstado = AdoPermiso.Recordset!estado_codigo
    frm_ao_Permisos.cmb_tipopermiso.BoundText = frm_ao_Permisos.Dtc_Par.BoundText
    frm_ao_Permisos.Show vbModal
    
 Else
        MsgBox "No se puede MODIFICAR un registro Aprobado o Anulado ...", vbExclamation, "Validación de Registro"
 End If
 Call abrirtabla
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
     
        frm_ao_memoranda.txtEstado = Ado_Memo.Recordset!estado_codigo
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
End Sub

Private Sub CmdMod5_Click()
  On Error GoTo EditErr
   If Ado_datos.Recordset.RecordCount > 0 Then
      If AdoMovilidad.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        frm_ro_movilidad_personal.txtSW = "MOD"
        
        frm_ro_movilidad_personal.txtCodigo.Text = AdoMovilidad.Recordset!numero_cambio
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
          frm_ro_movilidad_personal.DTPFelaboracion = AdoMovilidad.Recordset!fecha_elaboracion
          frm_ro_movilidad_personal.DTPFcontrato = AdoMovilidad.Recordset!fecha_inicio_contrato
    '      frm_ro_movilidad_personal.DTPFaprobacion = AdoMovilidad.Recordset!fecha_aprobacion
          frm_ro_movilidad_personal.TxtBs.Text = AdoMovilidad.Recordset!Item
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
    Set rsauxiliar = New ADODB.Recordset
    'SQL_FOR = "select * from rc_personal where ci = '" & txtCodigo.Text & "'"
    'rsauxiliar.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
    rsauxiliar.Open "select * from rc_personal where ci = '" & txtCodigo & "' ", db, adOpenKeyset, adLockOptimistic
    If rsauxiliar.RecordCount = 0 Then
        rsauxiliar.AddNew
        rsauxiliar!ci = txtCodigo
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
        rsauxiliar!Fecha_Registro = Date
        rsauxiliar!hora_registro = Format(Time, "HH:mm:ss")
        rsauxiliar!departamento_nacimiento = Dtc_depto.Text
        rsauxiliar!Procedencia = Dtc_prov.Text
        rsauxiliar!lugar_procedencia = Dtc_munic.Text
        'rsauxiliar!codigo_cargo = "-"   'TxtCargo.Text
'        rsauxiliar!numero_folder = Txt_mail.Text
        rsauxiliar!profesion = TxtProfesion.Text
        rsauxiliar.Update
        MkDir txtCodigo
        If Guardar_Imagen(db, "Select Foto From rv_personal_contratado Where beneficiario_codigo= '" & Ado_datos.Recordset("beneficiario_codigo") & "' ", "Foto", App.Path) Then
            MsgBox "ok"
        Else
            MsgBox "ERR"
        End If
        'Guardar_Imagen(cn, Sql, Campo, Path_Imagen)
End Sub

Private Sub BtnAñadir_Click()
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
      txtCodigo.Enabled = True
      fraDatos.Enabled = True
      Frame2.Enabled = True
      txtCodigo = Empty
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

End Sub

Private Sub BtnAprobar_Click()
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

End Sub

Private Sub BtnEliminar_Click()
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
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
    swnuevo = 0
   fraOpciones.Visible = True
       fra_cabecera.Enabled = True
       
         SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = True
       
   FraGrabarCancelar.Visible = False
   FraNavega.Enabled = True
   fraDatos.Enabled = False
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
End Sub
 
 Public Sub opciones()
   txtCodigo.Enabled = True
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
 End Sub
 
 
Private Sub BtnModificar_Click()
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

            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(0) = False

 
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
     txtCodigo.Enabled = False
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

Private Sub Command2_Click()
rw_reportes_asistencia.Show
End Sub

Private Sub Command1_Click()

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

Private Sub dtc_codigo7_Click(Area As Integer)
 dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    
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
    Dtc_depto_cod.BoundText = Dtc_depto.BoundText
    Call pProvincia(Dtc_depto_cod.BoundText)
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
'Ado_VacacionesProg   'Label5.Caption = GlUsuario       'frmLogin.txtUserName.Text      'JQA NOV-2009
   fraDatos.Enabled = False
'   fraDatos2.Enabled = False
'   FraSS_SS.Enabled = False
   Call Carga_Recor
   Call Carga_Beneficiario(1)
    Call Carga_Beneficiario(2)
'   Call rep_legal
   cbo_gestion.Text = Year(Date)
   cbo_mes.Text = UCase(MonthName(Month(Date)))
   txt_mes.Text = Month(Date)
   Call abrirtabla
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
'        End If
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
End Sub

Private Sub Carga_Beneficiario(posicion As Integer)
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

   
End Sub
Private Sub filtrar_asistencia(mes As String, ges_gestion As String)
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

End Sub
Private Sub abrirtabla()
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
    Set DtgAsistencia.DataSource = AdoAsistencia.Recordset
    cbo_gestion.Text = Year(Date)
    cbo_mes.Text = UCase(MonthName(Month(Date)))
    txt_mes.Text = Month(Date)
        
    sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & Year(Date) & "' AND Mes_control = '" & Month(Date) & "'"
    rs_AsisTT.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
    If rs_AsisTT.RecordCount > 0 Then
    rs_AsisTT.MoveFirst
    AdoAsistencia.Caption = CStr(rs_AsisTT!totHrs)
    Else
    AdoAsistencia.Caption = "0"
    End If
    Set rs_Permisos = New ADODB.Recordset
    If rs_Permisos.State = 1 Then rs_Permisos.Close
    rs_Permisos.Open "select * from ro_Permisos where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' order by FechaDesde ", db, adOpenKeyset, adLockOptimistic
    Set AdoPermiso.Recordset = rs_Permisos
    Set DtgPermiso.DataSource = AdoPermiso.Recordset
    
    Set rs_Permiso_detalle = New ADODB.Recordset
    If rs_Permiso_detalle.State = 1 Then rs_Permiso_detalle.Close
    rs_Permiso_detalle.Open "select * from ro_Permisos_detalle where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set AdoDependiente.Recordset = rs_Permiso_detalle
'    Set DtgDependiente.DataSource = AdoDependiente.Recordset
   
    Set rs_vacaciones_prog = New ADODB.Recordset
    If rs_vacaciones_prog.State = 1 Then rs_vacaciones_prog.Close
    rs_vacaciones_prog.Open "select * from ro_vacaciones_programadas where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' order by fecha_ini_Prog desc ", db, adOpenKeyset, adLockOptimistic
    Set Ado_VacacionesProg.Recordset = rs_vacaciones_prog
    Set DtgVacacionesProg.DataSource = Ado_VacacionesProg.Recordset
  
    Set rs_HORARIOS = New ADODB.Recordset
    If rs_HORARIOS.State = 1 Then rs_HORARIOS.Close
    rs_HORARIOS.Open "select * from RC_HORARIOS  ", db, adOpenKeyset, adLockOptimistic
    Set AdoHorarios.Recordset = rs_HORARIOS
'    Set DtgVacaciones.DataSource = AdoHorarios.Recordset
    
    Set rs_contrato = New Recordset
    If rs_contrato.State = 1 Then rs_contrato.Close
    rs_contrato.Open "select * from ro_memorandas where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' order by fecha_memo desc ", db, adOpenKeyset, adLockOptimistic
    
    Set Ado_Memo.Recordset = rs_contrato.DataSource
    Set DtG_Memo.DataSource = Ado_Memo.Recordset
    
    Set rs_movilidad = New Recordset
    If rs_movilidad.State = 1 Then rs_movilidad.Close
    rs_movilidad.Open "select * from ro_movilidad_personal where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
    Set AdoMovilidad.Recordset = rs_movilidad.DataSource
    Set DtgMovilidad.DataSource = AdoMovilidad.Recordset

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
    Dtc_depto.BoundText = Dtc_depto_cod.BoundText
    
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
    
    Set rs_Puesto = New ADODB.Recordset
    rs_Puesto.Open "select * from rc_puestos ", db, adOpenKeyset, adLockOptimistic
    Set AdoPuestoOrg.Recordset = rs_Puesto
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
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
   rs_beneficiario.Open "select * from gc_Beneficiario WHERE tipoben_codigo = '22'", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_Benef_seguro.Recordset = rs_beneficiario
   
   Set rs_CTA_BCO = New ADODB.Recordset
   If rs_CTA_BCO.State = 1 Then rs_CTA_BCO.Close
   rs_CTA_BCO.Open "select * from fv_cuenta_bco WHERE cta_codigo_tgn = '000'", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set AdoCta.Recordset = rs_CTA_BCO
   
    Set rs_ocupacion = New ADODB.Recordset
    rs_ocupacion.Open "select * from gc_ocupacion_profesion ", db, adOpenKeyset, adLockOptimistic
    Set Ado_Ocupacion.Recordset = rs_ocupacion
    
   Set rs_beneficiario_Afp = New ADODB.Recordset
   If rs_beneficiario_Afp.State = 1 Then rs_beneficiario_Afp.Close
   rs_beneficiario_Afp.Open "select * from gc_Beneficiario WHERE tipoben_codigo = '22'", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_Benef_Afp.Recordset = rs_beneficiario_Afp
      
   Set rs_EstCivil = New ADODB.Recordset
   If rs_EstCivil.State = 1 Then rs_EstCivil.Close
   rs_EstCivil.Open "select * from rc_estado_civil ", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set AdoEstCivil.Recordset = rs_EstCivil
   
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






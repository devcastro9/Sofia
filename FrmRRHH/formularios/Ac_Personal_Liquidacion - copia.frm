VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Ac_Personal_Liquidaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO DE LIQUIDACIONES"
   ClientHeight    =   8895
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   14085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14085
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   13320
      TabIndex        =   22
      Top             =   720
      Width           =   615
      Begin VB.Image ImgContrato 
         Height          =   540
         Left            =   15
         Picture         =   "Ac_Personal_Liquidacion.frx":0000
         Top             =   105
         Width           =   555
      End
   End
   Begin MSDataGridLib.DataGrid DtG_Auxiliar 
      Height          =   6105
      Left            =   15
      TabIndex        =   9
      Top             =   1755
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   10769
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "id_liquidacion"
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
         DataField       =   "Monto_Total"
         Caption         =   "Total Liquid."
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
         DataField       =   "codigo_motivo"
         Caption         =   "Motivo Retiro"
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
         DataField       =   "fecha_ingreso"
         Caption         =   "Fecha-Ingreso"
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
         DataField       =   "fecha_retiro"
         Caption         =   "Fecha-Retiro"
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
         DataField       =   "estado_registro"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOpciones 
      BackColor       =   &H80000018&
      Height          =   1140
      Left            =   20
      TabIndex        =   14
      Top             =   600
      Width           =   6090
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Contrato"
         Enabled         =   0   'False
         Height          =   720
         Left            =   3000
         Picture         =   "Ac_Personal_Liquidacion.frx":0388
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Carga Contrato"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "Ac_Personal_Liquidacion.frx":0710
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo Registro"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdMod 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   840
         Picture         =   "Ac_Personal_Liquidacion.frx":71FE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Anular"
         Height          =   720
         Left            =   1560
         Picture         =   "Ac_Personal_Liquidacion.frx":7AC8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Anula Registro Activo"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3720
         Picture         =   "Ac_Personal_Liquidacion.frx":8792
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Busca un Registro"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdSal 
         Caption         =   "Salir"
         Height          =   720
         Left            =   5160
         Picture         =   "Ac_Personal_Liquidacion.frx":905C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir de Contratos"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4440
         Picture         =   "Ac_Personal_Liquidacion.frx":9266
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprime Lista de Contratos"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton cmdAprueba 
         BackColor       =   &H0080C0FF&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2280
         Picture         =   "Ac_Personal_Liquidacion.frx":A9E8
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Aprueba Registro"
         Top             =   240
         Width           =   740
      End
   End
   Begin VB.Frame FraGrabarCancelar 
      BackColor       =   &H80000018&
      Height          =   1140
      Left            =   20
      TabIndex        =   16
      Top             =   600
      Width           =   6090
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   540
         Left            =   2760
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3360
         Picture         =   "Ac_Personal_Liquidacion.frx":ABF2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1920
         Picture         =   "Ac_Personal_Liquidacion.frx":ADFC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc Ado_Auxiliar 
      Height          =   330
      Left            =   0
      Top             =   7920
      Width           =   6105
      _ExtentX        =   10769
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
      BackColor       =   14737632
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
      Caption         =   "Navegar"
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
   Begin MSAdodcLib.Adodc AdoBeneficiario 
      Height          =   330
      Left            =   0
      Top             =   8280
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
      Caption         =   "AdoBeneficiario"
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
   Begin MSAdodcLib.Adodc AdoMotivos 
      Height          =   330
      Left            =   2160
      Top             =   8280
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
      Caption         =   "AdoMotivos"
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
      Left            =   4200
      Top             =   8280
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
   Begin MSAdodcLib.Adodc AdoOrg 
      Height          =   330
      Left            =   6240
      Top             =   8280
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
      Left            =   8280
      Top             =   8280
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
   Begin VB.Frame Fra_ABM 
      Height          =   7695
      Left            =   6120
      TabIndex        =   11
      Top             =   600
      Width           =   7935
      Begin VB.Frame Frame5 
         Caption         =   "IV. TOTAL BENEFICIOS SOCIALES"
         ForeColor       =   &H00C00000&
         Height          =   1360
         Left            =   120
         TabIndex        =   75
         Top             =   6240
         Width           =   7695
         Begin VB.ComboBox Combo7 
            DataField       =   "Forma_pago"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            ItemData        =   "Ac_Personal_Liquidacion.frx":B23E
            Left            =   120
            List            =   "Ac_Personal_Liquidacion.frx":B24B
            TabIndex        =   85
            Text            =   "CHEQUE"
            Top             =   540
            Width           =   1980
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Monto_Total"
            DataSource      =   "Ado_Auxiliar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5640
            MultiLine       =   -1  'True
            TabIndex        =   84
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Num_chq_cmpbte"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   540
            Width           =   1335
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Deducciones"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   78
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   4320
            MultiLine       =   -1  'True
            TabIndex        =   77
            Top             =   540
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total Beneficios Sociales"
            Height          =   195
            Index           =   27
            Left            =   3720
            TabIndex        =   83
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Deducciones"
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   82
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Bancaria"
            Height          =   195
            Index           =   25
            Left            =   4320
            TabIndex        =   81
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Nro.Cheq./Cmpbte."
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   79
            Top             =   300
            Width           =   1380
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   300
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "III. TOTAL REMUNERACION PROMEDIO INDEMNIZABLE"
         ForeColor       =   &H00C00000&
         Height          =   1935
         Left            =   120
         TabIndex        =   54
         Top             =   4200
         Width           =   7695
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Otros_Pagos"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   5880
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   1485
            Width           =   1455
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Prima_Legal"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   4080
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   1485
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "Años"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "Ac_Personal_Liquidacion.frx":B26F
            Left            =   5640
            List            =   "Ac_Personal_Liquidacion.frx":B2A0
            TabIndex        =   68
            Text            =   "0"
            Top             =   525
            Width           =   900
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "Años"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "Ac_Personal_Liquidacion.frx":B2E0
            Left            =   3480
            List            =   "Ac_Personal_Liquidacion.frx":B308
            TabIndex        =   67
            Text            =   "0"
            Top             =   525
            Width           =   900
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Utimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   5640
            MultiLine       =   -1  'True
            TabIndex        =   66
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Penul"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   65
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Antep"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   900
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "Años"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "Ac_Personal_Liquidacion.frx":B33C
            Left            =   1320
            List            =   "Ac_Personal_Liquidacion.frx":B36D
            TabIndex        =   60
            Text            =   "0"
            Top             =   525
            Width           =   900
         End
         Begin VB.TextBox TxtBs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Aguin_Vacacion"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   1485
            Width           =   1455
         End
         Begin VB.TextBox TxtMonto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Aguin_Navidad"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   1485
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Deshaucio 3 Meses por Retiro Forzoso:"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   2790
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Otros"
            Height          =   195
            Index           =   24
            Left            =   6000
            TabIndex        =   71
            Top             =   1245
            Width           =   375
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Prima Legal"
            Height          =   195
            Index           =   23
            Left            =   4080
            TabIndex        =   70
            Top             =   1245
            Width           =   825
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Vacaciones"
            Height          =   195
            Index           =   22
            Left            =   2160
            TabIndex        =   69
            Top             =   1245
            Width           =   840
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Dias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   6600
            TabIndex        =   63
            Top             =   555
            Width           =   390
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Meses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   4440
            TabIndex        =   62
            Top             =   555
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Años"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   2280
            TabIndex        =   61
            Top             =   555
            Width           =   435
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Aguinaldo Navidad"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   59
            Top             =   1245
            Width           =   1350
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Imdemnizacion . "
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   58
            Top             =   915
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Tiempo Servicio:"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   57
            Top             =   555
            Width           =   1185
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "II. LIQUIDACION PROMEDIO INDEMNIZABLE (3 Ultimos Meses)"
         ForeColor       =   &H00C00000&
         Height          =   1335
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         Width           =   7695
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Utimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   5640
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Penul"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   52
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Antep"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Txtpago3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Utimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   5640
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TxtPago2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Penultimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Combo6 
            DataField       =   "Mes_Utimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            ItemData        =   "Ac_Personal_Liquidacion.frx":B3AD
            Left            =   5640
            List            =   "Ac_Personal_Liquidacion.frx":B3D5
            TabIndex        =   47
            Text            =   "MARZO"
            Top             =   240
            Width           =   1860
         End
         Begin VB.ComboBox Combo5 
            DataField       =   "Mes_Penultimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            ItemData        =   "Ac_Personal_Liquidacion.frx":B43E
            Left            =   3480
            List            =   "Ac_Personal_Liquidacion.frx":B466
            TabIndex        =   46
            Text            =   "FEBRERO"
            Top             =   240
            Width           =   1860
         End
         Begin VB.ComboBox Combo4 
            DataField       =   "Mes_Antepenultimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            ItemData        =   "Ac_Personal_Liquidacion.frx":B4CF
            Left            =   1320
            List            =   "Ac_Personal_Liquidacion.frx":B4F7
            TabIndex        =   45
            Text            =   "ENERO"
            Top             =   240
            Width           =   1860
         End
         Begin VB.TextBox txtpago1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Antepenultimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Remuneracion . "
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Meses . . . . . . . ."
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Otros Pagos . . . "
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "I. DATOS GENERALES"
         ForeColor       =   &H00C00000&
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   7695
         Begin MSDataListLib.DataCombo DtcPuestoDes 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B560
            DataField       =   "codigo_motivo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   1540
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   741
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "denominacion_motivo"
            BoundColumn     =   "codigo_motivo"
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
         Begin MSDataListLib.DataCombo DtcPuesto 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B579
            DataField       =   "codigo_motivo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   1320
            TabIndex        =   25
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483624
            ListField       =   "codigo_motivo"
            BoundColumn     =   "codigo_motivo"
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
         Begin MSDataListLib.DataCombo DtcBenef 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B592
            DataField       =   "codigo_beneficiario"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   6240
            TabIndex        =   26
            Top             =   285
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "codigo_beneficiario"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcBenefDes 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B5B0
            DataField       =   "codigo_beneficiario"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   1320
            TabIndex        =   27
            Top             =   285
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "denominacion_beneficiario"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DTPFInicio 
            DataField       =   "fecha_ingreso"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   4560
            TabIndex        =   28
            Top             =   1540
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   41549825
            CurrentDate     =   40471
         End
         Begin MSComCtl2.DTPicker DTPFFin 
            DataField       =   "fecha_retiro"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   6240
            TabIndex        =   29
            Top             =   1540
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   41549825
            CurrentDate     =   40471
         End
         Begin MSDataListLib.DataCombo DtcInicial 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B5CE
            DataField       =   "codigo_beneficiario"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   1320
            TabIndex        =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "iniciales"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B5EC
            DataField       =   "codigo_beneficiario"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   5520
            TabIndex        =   31
            Top             =   885
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "estado_civil"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B60A
            DataField       =   "codigo_beneficiario"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   1320
            TabIndex        =   32
            Top             =   885
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "direccion_domicilio"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "Ac_Personal_Liquidacion.frx":B628
            DataField       =   "codigo_beneficiario"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   6240
            TabIndex        =   33
            Top             =   885
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "fecha_nacimiento"
            BoundColumn     =   "codigo_beneficiario"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha Retiro:"
            Height          =   195
            Index           =   1
            Left            =   6600
            TabIndex        =   40
            Top             =   1300
            Width           =   960
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Motivo de Retiro"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Top             =   1300
            Width           =   1170
         End
         Begin VB.Label lblLabels 
            Caption         =   "Funcionario / Trabajador . . . ."
            Height          =   435
            Index           =   20
            Left            =   120
            TabIndex        =   38
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ingreso:"
            Height          =   195
            Index           =   7
            Left            =   4680
            TabIndex        =   37
            Top             =   1300
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Dirección Dom. ."
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   36
            Top             =   885
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Est.Civil"
            Height          =   195
            Index           =   14
            Left            =   5520
            TabIndex        =   35
            Top             =   645
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Nacimiento"
            Height          =   195
            Index           =   15
            Left            =   6240
            TabIndex        =   34
            Top             =   645
            Width           =   1290
         End
      End
      Begin VB.TextBox TxtAprob 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "estado_registro"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   4125
         TabIndex        =   1
         Text            =   "NO"
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox Txtestado 
         DataField       =   "ges_gestion"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         ItemData        =   "Ac_Personal_Liquidacion.frx":B646
         Left            =   4680
         List            =   "Ac_Personal_Liquidacion.frx":B650
         TabIndex        =   0
         Text            =   "2011"
         Top             =   360
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox TxtLquida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "id_liquidacion"
         DataSource      =   "Ado_Auxiliar"
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
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Aprobado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   3120
         TabIndex        =   21
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblARCH 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         DataField       =   "ARCHIVO"
         DataSource      =   "Ado_Auxiliar"
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
         Height          =   195
         Left            =   6480
         TabIndex        =   20
         Top             =   405
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Registro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   375
         Width           =   1140
      End
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRO DE LIQUIDACIONES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   405
      Left            =   8640
      TabIndex        =   10
      Top             =   60
      Width           =   5220
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   525
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14055
   End
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   0
      Picture         =   "Ac_Personal_Liquidacion.frx":B660
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Ac_Personal_Liquidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_beneficiario As New ADODB.Recordset
Dim rs_Auxiliar As New ADODB.Recordset
Attribute rs_Auxiliar.VB_VarHelpID = -1
Dim rs_motivo As New ADODB.Recordset
Dim rs_Org As New ADODB.Recordset
Dim rs_Pry As New ADODB.Recordset
Dim rs_correlativo As New ADODB.Recordset
Dim e As Long

Dim var_cod As Integer
Dim VAR_VAL As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean


Private Sub cmdAprueba_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_Auxiliar!estado_registro = "NO" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "SI"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = GlUsuario
         rs_Auxiliar.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_Auxiliar.CancelUpdate
        If mvBookMark > 0 Then
          rs_Auxiliar.Bookmark = mvBookMark
        Else
          rs_Auxiliar.MoveFirst
        End If
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        DtG_Auxiliar.Enabled = True
    End If
End Sub

Private Sub CmdDel_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_Auxiliar!estado_registro = "SI" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "NL"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = GlUsuario
         rs_Auxiliar.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDesaprueba_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_Auxiliar!estado_registro = "SI" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "NO"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = GlUsuario
         rs_Auxiliar.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub


Private Sub CmdGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If GlSW = "ADD" Then
      rs_Auxiliar!fecha_ingreso = DTPFInicio.Value
      rs_Auxiliar!fecha_retiro = DTPFFin.Value
      rs_Auxiliar!codigo_beneficiario = DtcBenef.Text
      rs_Auxiliar!ges_gestion = glGestion
      rs_Auxiliar!codigo_motivo = DtcPuesto.Text
      
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ao_contrato_c WHERE codigo_beneficiario = '" & DtcBenef.Text & "'  ", DB, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            rs_Auxiliar!numero_consultoria = rs_correlativo.RecordCount
'            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
'            rs_correlativo.Update
'            rs_M1!Numero_FA = rs_correlativo!correlativo
      Else
            rs_Auxiliar!numero_consultoria = 1
      End If
      rs_Auxiliar!ARCHIVO = "Cargar_Archivo"
      rs_Auxiliar!ARCHIVO_NOMB = Trim(DtcInicial.Text) & "_Finiquito_" & rs_Auxiliar!numero_consultoria & ".pdf"
      TxtAprob.Text = "NO"
    End If
      rs_Auxiliar!monto_mensual = Txtpago3.Text
      rs_Auxiliar!Años = Combo1.Text
      rs_Auxiliar!meses = Combo2.Text
      rs_Auxiliar!dias = Combo3.Text
      rs_Auxiliar!Mes_Antepenultimo = Combo4.Text
      rs_Auxiliar!Mes_Penultimo = Combo5
      rs_Auxiliar!Mes_Utimo = Combo6
      rs_Auxiliar!Pago_Antepenultimo = txtpago1.Text
      rs_Auxiliar!Pago_Penultimo = TxtPago2
      rs_Auxiliar!Pago_Utimo = Txtpago3
      rs_Auxiliar!OtroPago_Antep = Text3.Text
'      If GlTipoCambioOficial > 0 Then
'        rs_Auxiliar!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      Else
'        GlTipoCambioOficial = 7.05
'        rs_Auxiliar!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      End If
      rs_Auxiliar!OtroPago_Penul = Text4
      rs_Auxiliar!OtroPago_Utimo = Text5
      rs_Auxiliar!Desah_3Meses = "0"
      rs_Auxiliar!Imdem_Año = Text6
      rs_Auxiliar!Imdem_Mes = Text7
      rs_Auxiliar!Indem_dias = Text8
      rs_Auxiliar!Aguin_Navidad = TxtMonto
      
      rs_Auxiliar!Aguin_Vacacion = TxtBs
      rs_Auxiliar!Prima_Legal = Text9
      rs_Auxiliar!Otros_Pagos = Text10
      rs_Auxiliar!Forma_pago = Combo7
      rs_Auxiliar!Num_chq_cmpbte = Text13
      rs_Auxiliar!cta_codigo = Text11
      rs_Auxiliar!Deducciones = Text12
      
      rs_Auxiliar!Monto_Total = Text14
      
      rs_Auxiliar!hora_registro = "8:00"
      rs_Auxiliar!fecha_registro = Date
      rs_Auxiliar!usr_usuario = "ADMIN" 'GlUsuario
      rs_Auxiliar.Update    'Batch adAffectAll
      
      mbDataChanged = False
    
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      DtG_Auxiliar.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If DtcBenef.Text = "" Then
    MsgBox "Debe registrar a la persona Beneficiaria ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If Text14.Text = "" Then
    MsgBox "Debe registrar el Monto Total de la Liquidacion ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If DTPFInicio.Value > DTPFFin.Value Then
    MsgBox "La Fecha de Retiro NO puede ser Mayor a la de Ingreso  ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
'  If DTPFInicio.Value > DTPFFin.Value Then
'    MsgBox "La Fecha de Inicio NO puede ser Mayor a la de Finalizacion del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If

End Sub

Private Sub CmdMod_Click()
  On Error GoTo EditErr
  If Ado_Auxiliar.Recordset!estado_registro = "SI" Then
    MsgBox "No se puede modificar un registro APROBADO ...", vbCritical + vbExclamation, "Validación de datos"
    Exit Sub
  Else
'  lblStatus.Caption = "Modificar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    DtG_Auxiliar.Enabled = False
    GlSW = "MOD"
    Exit Sub
  End If


EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdSal_Click()
'  If glPersNew = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_Auxiliar!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_Auxiliar!ocup_descripcion
'  End If
'  glPersNew = "N"
  Unload Me
End Sub

Private Sub CmdVerDisco_Click()
  On Error GoTo Error_Sub
    
    NombreCarpeta = App.Path & "\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\"
       'e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
   
'    Mensaje = NombreCarpeta
'    Call Eliminar_Directorio(NombreCarpeta)
'    Mensaje = e
'    Call Eliminar_Directorio(e)
    Frmexporta.DirDestino.Path = NombreCarpeta
'SERVIDOR
'    e = "\\SRVPRO\SIGPER\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\"
    'Frmexporta.DirDestino2.Path = e
    Frmexporta.Show vbModal
    'MsgBox "Coloque el CD, para volver a COPIAR su contenido ... ", vbCritical + vbExclamation, "Realiza la Copia de CD"
    'sino = MsgBox("Desea Borrar los datos copiados anteriormente en su computadora ? ", vbYesNo + vbQuestion, "Atención")
    'If sino = vbYes Then
    '    Kill NombreCarpeta & "\*.*"
    '    Kill e & "\*.*"
    '    My.Computer.FileSystem.DeleteFile (NombreCarpeta & "\*.*")
    '    'My.Computer.FileSystem.DeleteFile(NombreCarpeta & "\*.*", FileIO.UIOption.AllDialogs, FileIO.RecycleOption.DeletePermanently, FileIO.UICancelOption.DoNothing)

    '    'MkDir NombreCarpeta
    '    'MkDir e
    'End If
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'fs.CopyFile "G:\*.*", NombreCarpeta
    'fs.CopyFile "G:\*.*", e
    'fs.CopyFile "F:\WIN\*.*", NombreCarpeta
    'fs.CopyFile "F:\COPIA\*.*", e
  Exit Sub
Error_Sub:
  MsgBox Err.Description, vbCritical
    
End Sub

Private Sub DtcBenef_Click(Area As Integer)
    DtcBenefDes.BoundText = DtcBenef.BoundText
End Sub

Private Sub DtcBenefDes_Click(Area As Integer)
    DtcBenef.BoundText = DtcBenefDes.BoundText
End Sub

Private Sub DtcPuesto_Click(Area As Integer)
    DtcPuestoDes.BoundText = DtcPuesto.BoundText
End Sub

Private Sub DtcPuestoDes_Click(Area As Integer)
    DtcPuesto.BoundText = DtcPuestoDes.BoundText
End Sub

Private Sub Form_Load()

  Call abrirtabla
  
  Set rs_beneficiario = New ADODB.Recordset
  rs_beneficiario.Open "select * from gc_Beneficiario WHERE tipo_beneficiario='1' ORDER BY denominacion_beneficiario ", DB, adOpenKeyset, adLockOptimistic
  Set AdoBeneficiario.Recordset = rs_beneficiario.DataSource
  DtcBenefDes.BoundText = DtcBenef.BoundText
  
  Set rs_motivo = New ADODB.Recordset
  rs_motivo.Open "select * from ac_no_motivo WHERE estado_registro = 'L'  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoMotivos.Recordset = rs_motivo.DataSource
  DtcPuestoDes.BoundText = DtcPuesto.BoundText
  
'  Set rs_UNIDAD = New ADODB.Recordset
'  rs_UNIDAD.Open "select * from fc_unidad_ejecutora  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoUnidad.Recordset = rs_UNIDAD.DataSource
'  Dtc_descrip.BoundText = Dtc_codigo.BoundText
'
'  Set rs_Org = New ADODB.Recordset
'  rs_Org.Open "select * from fc_convenios  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoOrg.Recordset = rs_Org.DataSource
'  DtcOrgDes.BoundText = DtcOrg.BoundText
'
'  Set rs_Pry = New ADODB.Recordset
'  rs_Pry.Open "select * from fc_estructura_programatica  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoPry.Recordset = rs_Pry.DataSource
'  DtcPryDes.BoundText = DtcPry.BoundText
  
'  rs_Auxiliar.AddNew
'  txtParam.Text = GlParametro
'  TxtForm.Text = GlForm
'  TxtCorrel.Text = GlCorrel

  mbDataChanged = False
  Fra_ABM.Enabled = False
  DtG_Auxiliar.Enabled = True
  GlSW = "NADA"
	Call SeguridadSet(Me)
End Sub

Private Sub abrirtabla()
  Set rs_Auxiliar = New Recordset
  If rs_Auxiliar.State = 1 Then rs_Auxiliar.Close
  'queryinicial = "select * from rc_puesto_organizacional where param_codigo = '" & GlParametro & "' "
  queryinicial = "select * from ro_liquidaciones "
  rs_Auxiliar.Open queryinicial, DB, adOpenKeyset, adLockOptimistic
  rs_Auxiliar.Sort = "codigo_beneficiario, fecha_ingreso"
  Set Ado_Auxiliar.Recordset = rs_Auxiliar.DataSource
  Set DtG_Auxiliar.DataSource = Ado_Auxiliar.Recordset
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
'    frmeo_Larvas_mosquitos.Fra_detalle.Visible = False
End Sub

Private Sub Ado_Auxiliar_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Muestra la posición de registro actual para este Recordset
      Ado_Auxiliar.Caption = Ado_Auxiliar.Recordset.AbsolutePosition & " / " & Ado_Auxiliar.Recordset.RecordCount
End Sub

'Private Sub Ado_Auxiliar_WillChangeRecord(ByVal adReason As adodb.EventReasonEnum, ByVal cRecords As Long, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)
'  'Aquí se coloca el código de validación
'  'Se llama a este evento cuando ocurre la siguiente acción
'  Dim bCancel As Boolean
'
'  Select Case adReason
'  Case adRsnAddNew
'  Case adRsnClose
'  Case adRsnDelete
'  Case adRsnFirstChange
'  Case adRsnMove
'  Case adRsnRequery
'  Case adRsnResynch
'  Case adRsnUndoAddNew
'  Case adRsnUndoDelete
'  Case adRsnUndoUpdate
'  Case adRsnUpdate
'  End Select
'
'  If bCancel Then adStatus = adStatusCancel
'End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    'rs_Auxiliar.MoveLast
    rs_Auxiliar.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    DtG_Auxiliar.Enabled = False
    GlSW = "ADD"
'    rs_Auxiliar.AddNew
'    txtParam.Text = GlParametro
'    TxtForm.Text = "E-1" 'GlForm
'    TxtCorrel.Text = 1  'GlCorrel
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_Auxiliar.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub ImgContrato_Click()
 If lblARCH.Caption = "Cargar_Archivo" Then
    MsgBox ("No Existe el Archivo Asociado al Contrato, debe Cargarlo ...")
 Else
    If GlServidor = "SRVPRO" Then
        e = ShellExecute(Img_CV, "open", "\\SRVPRO\SIGPER\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\FINIQUITO\" & Trim(DtcInicial.Text) & "-Contrato-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
    Else
        e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\FINIQUITO\" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
    End If
 End If
End Sub


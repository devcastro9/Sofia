VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Ac_Personal_Contrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO DE CONTRATOS"
   ClientHeight    =   7455
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   13215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   13215
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   12480
      TabIndex        =   51
      Top             =   720
      Width           =   615
      Begin VB.Image ImgContrato 
         Height          =   540
         Left            =   20
         Picture         =   "Ac_Personal_Contrato.frx":0000
         Top             =   100
         Width           =   555
      End
   End
   Begin MSDataGridLib.DataGrid DtG_Auxiliar 
      Height          =   4665
      Left            =   15
      TabIndex        =   21
      Top             =   1755
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   8229
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "id_contrato"
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
         DataField       =   "codigo_contrato"
         Caption         =   "Cod-Contrato"
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
         DataField       =   "codigo_unidad"
         Caption         =   "Area"
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
         DataField       =   "fecha_inicio"
         Caption         =   "Fecha-Inicio"
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
         DataField       =   "fecha_fin"
         Caption         =   "Fecha-Final"
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
         DataField       =   "fechas_confirmado"
         Caption         =   "Vig."
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
         DataField       =   "cod_est_contrato"
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
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   329.953
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOpciones 
      BackColor       =   &H80000018&
      Height          =   1140
      Left            =   15
      TabIndex        =   31
      Top             =   600
      Width           =   6090
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cargar"
         Height          =   720
         Left            =   3000
         Picture         =   "Ac_Personal_Contrato.frx":0388
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Carga Contrato"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "Ac_Personal_Contrato.frx":0710
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Nuevo Registro"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdMod 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   840
         Picture         =   "Ac_Personal_Contrato.frx":71FE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Anular"
         Height          =   720
         Left            =   1560
         Picture         =   "Ac_Personal_Contrato.frx":7AC8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anula Registro Activo"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3720
         Picture         =   "Ac_Personal_Contrato.frx":8792
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Busca un Registro"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdSal 
         Caption         =   "Salir"
         Height          =   720
         Left            =   5160
         Picture         =   "Ac_Personal_Contrato.frx":905C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salir de Contratos"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4440
         Picture         =   "Ac_Personal_Contrato.frx":9266
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Imprime Lista de Contratos"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton cmdAprueba 
         BackColor       =   &H0080C0FF&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2280
         Picture         =   "Ac_Personal_Contrato.frx":A9E8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Aprueba Registro"
         Top             =   240
         Width           =   740
      End
   End
   Begin VB.Frame FraGrabarCancelar 
      BackColor       =   &H80000018&
      Height          =   1140
      Left            =   20
      TabIndex        =   33
      Top             =   600
      Width           =   6090
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   540
         Left            =   2760
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3360
         Picture         =   "Ac_Personal_Contrato.frx":ABF2
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1920
         Picture         =   "Ac_Personal_Contrato.frx":ADFC
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc Ado_Auxiliar 
      Height          =   330
      Left            =   0
      Top             =   6480
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
      Top             =   6960
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
   Begin MSAdodcLib.Adodc AdoPuestoOrg 
      Height          =   330
      Left            =   2160
      Top             =   6960
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
   Begin MSAdodcLib.Adodc AdoUnidad 
      Height          =   330
      Left            =   4200
      Top             =   6960
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
      Top             =   6960
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
      Top             =   6960
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
      Height          =   6255
      Left            =   6120
      TabIndex        =   23
      Top             =   600
      Width           =   7095
      Begin VB.TextBox TxtAprob 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "cod_est_contrato"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   3640
         TabIndex        =   1
         Text            =   "NO"
         Top             =   480
         Width           =   495
      End
      Begin MSDataListLib.DataCombo Dtc_descrip 
         Bindings        =   "Ac_Personal_Contrato.frx":B23E
         DataField       =   "codigo_unidad"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   3720
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Uni_descripcion_larga"
         BoundColumn     =   "codigo_unidad"
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
      Begin MSDataListLib.DataCombo DtcPryDes 
         Bindings        =   "Ac_Personal_Contrato.frx":B256
         DataField       =   "Pro_proyecto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   4920
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Pro_descripcion_larga"
         BoundColumn     =   "Pro_proyecto"
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
      Begin VB.ComboBox Txtestado 
         DataField       =   "fechas_confirmado"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         ItemData        =   "Ac_Personal_Contrato.frx":B26B
         Left            =   2160
         List            =   "Ac_Personal_Contrato.frx":B275
         TabIndex        =   0
         Text            =   "SI"
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox TxtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "codigo_contrato"
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
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPFFirma 
         DataField       =   "fecha_firma"
         DataSource      =   "Ado_Auxiliar"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   5760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   89260033
         CurrentDate     =   40471
      End
      Begin VB.TextBox txtObjContrato 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "objeto_contrato"
         DataSource      =   "Ado_Auxiliar"
         Height          =   645
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2160
         Width           =   5655
      End
      Begin VB.TextBox TxtForm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "id_contrato"
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
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtBs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_totalBS"
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
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcPuestoDes 
         Bindings        =   "Ac_Personal_Contrato.frx":B281
         DataField       =   "codigo_puesto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   3120
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "denominacion_puesto"
         BoundColumn     =   "codigo_puesto"
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
      Begin MSDataListLib.DataCombo DtcPuesto 
         Bindings        =   "Ac_Personal_Contrato.frx":B29C
         DataField       =   "codigo_puesto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   38
         Top             =   2840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_puesto"
         BoundColumn     =   "codigo_puesto"
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
      Begin MSDataListLib.DataCombo DtcBenef 
         Bindings        =   "Ac_Personal_Contrato.frx":B2B7
         DataField       =   "codigo_beneficiario"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   5640
         TabIndex        =   5
         Top             =   1560
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcBenefDes 
         Bindings        =   "Ac_Personal_Contrato.frx":B2D5
         DataField       =   "codigo_beneficiario"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1560
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPFInicio 
         DataField       =   "fecha_inicio"
         DataSource      =   "Ado_Auxiliar"
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   5760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   89260033
         CurrentDate     =   40471
      End
      Begin MSComCtl2.DTPicker DTPFFin 
         DataField       =   "fecha_fin"
         DataSource      =   "Ado_Auxiliar"
         Height          =   285
         Left            =   5400
         TabIndex        =   13
         Top             =   5760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   89260033
         CurrentDate     =   40471
      End
      Begin MSDataListLib.DataCombo Dtc_codigo 
         Bindings        =   "Ac_Personal_Contrato.frx":B2F3
         DataField       =   "codigo_unidad"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   41
         Top             =   3440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_unidad"
         BoundColumn     =   "codigo_unidad"
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
      Begin MSDataListLib.DataCombo DtcOrgDes 
         Bindings        =   "Ac_Personal_Contrato.frx":B30B
         DataField       =   "Codigo_Convenio"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   4320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Denominacion_Convenio"
         BoundColumn     =   "nivel_puestoCodigo_Convenio"
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
      Begin MSDataListLib.DataCombo DtcOrg 
         Bindings        =   "Ac_Personal_Contrato.frx":B320
         DataField       =   "Codigo_Convenio"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   44
         Top             =   4040
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "Codigo_Convenio"
         BoundColumn     =   "Codigo_Convenio"
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
      Begin MSDataListLib.DataCombo DtcPry 
         Bindings        =   "Ac_Personal_Contrato.frx":B335
         DataField       =   "Pro_proyecto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   46
         Top             =   4640
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "Pro_proyecto"
         BoundColumn     =   "Pro_proyecto"
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
      Begin MSDataListLib.DataCombo DtcInicial 
         Bindings        =   "Ac_Personal_Contrato.frx":B34A
         DataField       =   "codigo_beneficiario"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1320
         TabIndex        =   48
         Top             =   1280
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "iniciales"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   ""
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
         Left            =   3480
         TabIndex        =   50
         Top             =   240
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
         Left            =   5760
         TabIndex        =   49
         Top             =   400
         Width           =   585
      End
      Begin VB.Label lblLabels 
         Caption         =   "Puesto que Ocupa . . . . . . ."
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   3075
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Proyecto . . . . ."
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   47
         Top             =   4965
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Financiador . . . ."
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   45
         Top             =   4365
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cód.Contrato:"
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
         Index           =   9
         Left            =   120
         TabIndex        =   43
         Top             =   980
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         Caption         =   "Area Organizacional ."
         Height          =   435
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   3675
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio:"
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   40
         Top             =   5520
         Width           =   1380
      End
      Begin VB.Label lblLabels 
         Caption         =   "Monto Total del Contrato:"
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   39
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Vigente:"
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
         Index           =   5
         Left            =   2160
         TabIndex        =   37
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lblLabels 
         Caption         =   "Funcionario / Trabajador . . . ."
         Height          =   435
         Index           =   20
         Left            =   120
         TabIndex        =   30
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Objeto del Contrato . . . . . ."
         Height          =   435
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   2300
         Width           =   1170
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
         TabIndex        =   28
         Top             =   255
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Firma Contrato:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   5520
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Finalización:"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   25
         Top             =   5520
         Width           =   1365
      End
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRO DE CONTRATOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   405
      Left            =   8160
      TabIndex        =   22
      Top             =   60
      Width           =   4620
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   525
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   0
      Picture         =   "Ac_Personal_Contrato.frx":B368
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Ac_Personal_Contrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_beneficiario As New ADODB.Recordset
Dim rs_Auxiliar As New ADODB.Recordset
Attribute rs_Auxiliar.VB_VarHelpID = -1
Dim rs_Puesto_Org As New ADODB.Recordset
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
   If rs_Auxiliar!cod_est_contrato = "NO" Then
      If sino = vbYes Then
         rs_Auxiliar!cod_est_contrato = "SI"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = glusuario
         rs_Auxiliar.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancelar_Click()
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
   If rs_Auxiliar!estado_registro = "S" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "L"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = glusuario
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
   If rs_Auxiliar!estado_registro = "S" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "N"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = glusuario
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
      rs_Auxiliar!codigo_contrato = txtCodigo.Text
      rs_Auxiliar!codigo_beneficiario = DtcBenef.Text
      rs_Auxiliar!ges_gestion = glGestion
      rs_Auxiliar!codigo_solicitud = rs_Auxiliar.RecordCount
      
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ao_contrato_c WHERE codigo_beneficiario = '" & DtcBenef.Text & "'  ", db, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            rs_Auxiliar!numero_consultoria = rs_correlativo.RecordCount
'            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
'            rs_correlativo.Update
'            rs_M1!Numero_FA = rs_correlativo!correlativo
      Else
            rs_Auxiliar!numero_consultoria = 1
      End If
      rs_Auxiliar!ARCHIVO = "Cargar_Archivo"
      rs_Auxiliar!ARCHIVO_NOMB = Trim(DtcInicial.Text) & "_Contrato_" & rs_Auxiliar!numero_consultoria & ".pdf"
      TxtAprob.Text = "NO"
    End If
      rs_Auxiliar!objeto_contrato = txtObjContrato.Text
      rs_Auxiliar!codigo_puesto = DtcPuesto.Text
      rs_Auxiliar!codigo_unidad = dtc_codigo.Text
      rs_Auxiliar!codigo_convenio = DtcOrg.Text
      rs_Auxiliar!pro_proyecto = DtcPry.Text
      rs_Auxiliar!fechas_confirmado = txtEstado
      rs_Auxiliar!cod_est_contrato = TxtAprob
      rs_Auxiliar!fecha_firma = DTPFFirma.Value
      rs_Auxiliar!fecha_inicio = DTPFInicio.Value
      rs_Auxiliar!fecha_fin = DTPFFin.Value
      rs_Auxiliar!monto_totalbs = TxtBs.Text
      If GlTipoCambioOficial > 0 Then
        rs_Auxiliar!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      Else
        GlTipoCambioOficial = 7.05
        rs_Auxiliar!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      End If
      rs_Auxiliar!observacion_contrato = "-"
      rs_Auxiliar!establece_multas = "N"
      rs_Auxiliar!cod_forma_inicio = "1"
      rs_Auxiliar!tiempo_num = 0
      rs_Auxiliar!tiempo_dmy = "-"
      rs_Auxiliar!tipo_moneda = "Bs"
      rs_Auxiliar!tc_us = GlTipoCambioOficial
      
      rs_Auxiliar!org_codigo = "111"
      rs_Auxiliar!porc_orgfin = 100
      rs_Auxiliar!porc_contra = 0
      'rs_Auxiliar!fechas_confirmado = "N"
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
  If dtc_codigo.Text = "" Then
    MsgBox "Debe registrar el Código o Cite del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If TxtBs.Text = "" Then
    MsgBox "Debe registrar el Monto del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If DTPFFirma.Value > DTPFInicio.Value Then
    MsgBox "La Fecha de Firma NO puede ser Mayor a la de Inicio del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If DTPFInicio.Value > DTPFFin.Value Then
    MsgBox "La Fecha de Inicio NO puede ser Mayor a la de Finalizacion del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

End Sub

Private Sub CmdMod_Click()
  On Error GoTo EditErr
  If Ado_Auxiliar.Recordset!cod_est_contrato = "SI" Then
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
  If Ado_Auxiliar.Recordset!ARCHIVO = "Cargar_Archivo" Then
    NombreCarpeta = App.Path & "\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\"
    Frmexporta.DirDestino.Path = NombreCarpeta
    'e = "\\SRVPRO\SIGPER\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\"
    'Frmexporta.DirDestino2.Path = e
    Frmexporta.Show vbModal
  Else
'    MsgBox ""
    sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
        NombreCarpeta = App.Path & "\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        'e = "\\SRVPRO\SIGPER\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\"
        'Frmexporta.DirDestino2.Path = e
        Frmexporta.Show vbModal
    End If
  End If
'    Mensaje = NombreCarpeta
'    Call Eliminar_Directorio(NombreCarpeta)
'    Mensaje = e
'    Call Eliminar_Directorio(e)
'SERVIDOR
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

Private Sub dtc_codigo_Click(Area As Integer)
    Dtc_descrip.BoundText = dtc_codigo.BoundText
End Sub

Private Sub Dtc_descrip_Click(Area As Integer)
    dtc_codigo.BoundText = Dtc_descrip.BoundText
End Sub

Private Sub DtcBenef_Click(Area As Integer)
    DtcBenefDes.BoundText = DtcBenef.BoundText
End Sub

Private Sub DtcBenefDes_Click(Area As Integer)
    DtcBenef.BoundText = DtcBenefDes.BoundText
End Sub

Private Sub DtcOrg_Click(Area As Integer)
    DtcOrgDes.BoundText = DtcOrg.BoundText
End Sub

Private Sub DtcOrgDes_Click(Area As Integer)
    DtcOrg.BoundText = DtcOrgDes.BoundText
End Sub

Private Sub DtcPry_Click(Area As Integer)
    DtcPryDes.BoundText = DtcPry.BoundText
End Sub

Private Sub DtcPryDes_Click(Area As Integer)
    DtcPry.BoundText = DtcPryDes.BoundText
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
  rs_beneficiario.Open "select * from gc_Beneficiario WHERE tipo_beneficiario='1' ORDER BY denominacion_beneficiario ", db, adOpenKeyset, adLockOptimistic
  Set AdoBeneficiario.Recordset = rs_beneficiario.DataSource
  DtcBenefDes.BoundText = DtcBenef.BoundText
  
  Set rs_Puesto_Org = New ADODB.Recordset
  rs_Puesto_Org.Open "select * from rc_PUESTO_organizacional  ", db, adOpenKeyset, adLockOptimistic
  Set AdoPuestoOrg.Recordset = rs_Puesto_Org.DataSource
  DtcPuestoDes.BoundText = DtcPuesto.BoundText
  
  Set rs_UNIDAD = New ADODB.Recordset
  rs_UNIDAD.Open "select * from fc_unidad_ejecutora  ", db, adOpenKeyset, adLockOptimistic
  Set AdoUnidad.Recordset = rs_UNIDAD.DataSource
  Dtc_descrip.BoundText = dtc_codigo.BoundText
  
  Set rs_Org = New ADODB.Recordset
  rs_Org.Open "select * from fc_convenios  ", db, adOpenKeyset, adLockOptimistic
  Set AdoOrg.Recordset = rs_Org.DataSource
  DtcOrgDes.BoundText = DtcOrg.BoundText
  
  Set rs_Pry = New ADODB.Recordset
  rs_Pry.Open "select * from fc_estructura_programatica  ", db, adOpenKeyset, adLockOptimistic
  Set AdoPry.Recordset = rs_Pry.DataSource
  DtcPryDes.BoundText = DtcPry.BoundText
  
  
  
'  rs_Auxiliar.AddNew
'  txtParam.Text = GlParametro
'  TxtForm.Text = GlForm
'  TxtCorrel.Text = GlCorrel

  mbDataChanged = False
  Fra_ABM.Enabled = False
  DtG_Auxiliar.Enabled = True
  GlSW = "NADA"
End Sub

Private Sub abrirtabla()
  Set rs_Auxiliar = New Recordset
  If rs_Auxiliar.State = 1 Then rs_Auxiliar.Close
  'queryinicial = "select * from rc_puesto_organizacional where param_codigo = '" & GlParametro & "' "
  queryinicial = "select * from ao_contrato_c "
  rs_Auxiliar.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  rs_Auxiliar.Sort = "codigo_beneficiario, codigo_unidad"
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
  If Ado_Auxiliar.Recordset.RecordCount > 0 Then
    If Ado_Auxiliar.Recordset("cod_est_contrato") = "SI" Then
        TxtAprob.ForeColor = &H8000&
    Else
        TxtAprob.ForeColor = &HC0&
    End If
    If Ado_Auxiliar.Recordset("ARCHIVO") = "Cargar_Archivo" Then
        lblARCH.ForeColor = &HC0&
    Else
        lblARCH.ForeColor = &H8000&
    End If
      Ado_Auxiliar.Caption = Ado_Auxiliar.Recordset.AbsolutePosition & " / " & Ado_Auxiliar.Recordset.RecordCount
  End If
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
        e = ShellExecute(Img_CV, "open", "\\SRVPRO\SIGPER\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\" & Trim(DtcInicial.Text) & "-Contrato-1.pdf", vbNullString, vbNullString, SW_SHOWNORMAL)
    Else
        e = ShellExecute(Img_CV, "open", App.Path & "\PERSONAL\" & Trim(DtcInicial.Text) & "-" & Trim(Ado_Auxiliar.Recordset!codigo_beneficiario) & "\CONTRATOS\" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
    End If
 End If
End Sub


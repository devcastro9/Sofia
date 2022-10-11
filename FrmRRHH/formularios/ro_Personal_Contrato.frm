VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ro_Personal_Contrato 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administracion de Personal - Ficha Personal - Contratos, Adendas o Designaciones"
   ClientHeight    =   6360
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   9375
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "ro_Personal_Contrato.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   9075
      TabIndex        =   42
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   3000
         Picture         =   "ro_Personal_Contrato.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ver Contrato PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cargar"
         Height          =   680
         Left            =   2160
         Picture         =   "ro_Personal_Contrato.frx":6C3BA
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Carga Contrato en PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "ro_Personal_Contrato.frx":6C742
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "ro_Personal_Contrato.frx":6C94C
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRATOS Y DESIGNACIONES"
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
         Left            =   3990
         TabIndex        =   47
         Top             =   240
         Width           =   4935
      End
   End
   Begin MSAdodcLib.Adodc AdoBeneficiario 
      Height          =   330
      Left            =   0
      Top             =   6240
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
      Left            =   2040
      Top             =   6240
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
      Left            =   4080
      Top             =   6240
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
      Left            =   6120
      Top             =   6240
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
      Left            =   0
      Top             =   6600
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
      ForeColor       =   &H00FFFF80&
      Height          =   5295
      Left            =   135
      TabIndex        =   14
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   5760
         MaxLength       =   80
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   4080
         MaxLength       =   80
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtAprob 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5085
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "NO"
         Top             =   520
         Width           =   495
      End
      Begin MSDataListLib.DataCombo Dtc_descrip 
         DataField       =   "codigo_unidad"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   4680
         TabIndex        =   7
         Top             =   2640
         Width           =   4335
         _ExtentX        =   7646
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
         DataField       =   "Pro_proyecto"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   4335
         _ExtentX        =   7646
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
         Height          =   315
         Left            =   3360
         TabIndex        =   1
         Text            =   "SI"
         Top             =   520
         Width           =   660
      End
      Begin VB.TextBox TxtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "codigo_contrato"
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
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   520
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPFFirma 
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   4800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   116195329
         CurrentDate     =   40471
      End
      Begin VB.TextBox txtObjContrato 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1560
         Width           =   8535
      End
      Begin VB.TextBox TxtForm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   7560
         TabIndex        =   17
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtBs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   7215
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   4800
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DtcPuestoDes 
         DataField       =   "codigo_puesto"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   4080
         Width           =   4335
         _ExtentX        =   7646
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
         DataField       =   "codigo_puesto"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   6360
         TabIndex        =   21
         Top             =   3795
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_puesto"
         BoundColumn     =   "codigo_puesto"
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
      Begin MSComCtl2.DTPicker DTPFInicio 
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Top             =   4800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   116195329
         CurrentDate     =   40471
      End
      Begin MSComCtl2.DTPicker DTPFFin 
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Top             =   4800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   116195329
         CurrentDate     =   44196
      End
      Begin MSDataListLib.DataCombo Dtc_codigo 
         DataField       =   "codigo_unidad"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   6360
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_unidad"
         BoundColumn     =   "codigo_unidad"
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
      Begin MSDataListLib.DataCombo DtcOrgDes 
         DataField       =   "org_codigo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Org_descripcion"
         BoundColumn     =   "org_codigo"
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
         DataField       =   "org_codigo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   26
         Top             =   3075
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "org_codigo"
         BoundColumn     =   "org_codigo"
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
      Begin MSDataListLib.DataCombo DtcPry 
         DataField       =   "Pro_proyecto"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   27
         Top             =   3795
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "Pro_proyecto"
         BoundColumn     =   "Pro_proyecto"
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
      Begin MSDataListLib.DataCombo DtcFteDes 
         DataField       =   "Fte_codigo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Fte_descripcion_larga"
         BoundColumn     =   "Fte_codigo"
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
      Begin MSDataListLib.DataCombo DtcFte 
         DataField       =   "Fte_codigo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   36
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "Fte_codigo"
         BoundColumn     =   "Fte_codigo"
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
      Begin MSDataListLib.DataCombo DtcCargoDes 
         DataField       =   "codigo_cargo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   3345
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "descripcion_cargo"
         BoundColumn     =   "codigo_cargo"
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
      Begin MSDataListLib.DataCombo DtcCargo 
         DataField       =   "codigo_cargo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   6360
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_cargo"
         BoundColumn     =   "codigo_cargo"
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
      Begin MSDataListLib.DataCombo DtcRespaldo 
         DataField       =   "doc_codigo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   1920
         TabIndex        =   40
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
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
      Begin MSDataListLib.DataCombo DtcRespaldoCod 
         DataField       =   "doc_codigo"
         DataSource      =   "frmBeneficiario.Ado_Contrato"
         Height          =   315
         Left            =   4800
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
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
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFC0&
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   2280
         Y2              =   5280
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Documento:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   39
         Top             =   1005
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFC0&
         X1              =   0
         X2              =   9120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nombre de Archivo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   7485
         TabIndex        =   38
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Proyecto                                                                                       Puesto que Ocupa "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   3840
         Width           =   5925
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Organismo Financiador                                                                 Cargo que Ocupa"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   37
         Top             =   3120
         Width           =   5805
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fuente Financiamiento                                                                  Area Organizacional"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   28
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   6000
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Aprobado"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   4965
         TabIndex        =   29
         Top             =   300
         Width           =   690
      End
      Begin VB.Label lblARCH 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "-"
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
         Left            =   8790
         TabIndex        =   28
         Top             =   555
         Width           =   75
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Código Contrato"
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
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   300
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Inicio:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   23
         Top             =   4560
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Monto Total Contrato"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   7200
         TabIndex        =   22
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Vigente"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   20
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Objeto del Contrato"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1335
         Width           =   1530
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Correlativo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   6360
         TabIndex        =   18
         Top             =   1005
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Firma Contrato:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4920
         TabIndex        =   16
         Top             =   4560
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Finalización:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   15
         Top             =   4560
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc AdoFuente 
      Height          =   330
      Left            =   2040
      Top             =   6600
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
   Begin MSAdodcLib.Adodc AdoCargo 
      Height          =   330
      Left            =   4080
      Top             =   6600
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
   Begin MSAdodcLib.Adodc AdoRespaldo 
      Height          =   330
      Left            =   6120
      Top             =   6600
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
      Caption         =   "AdoRespaldo"
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
Attribute VB_Name = "ro_Personal_Contrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_FteFin As New ADODB.Recordset
Dim rs_Org As New ADODB.Recordset
Dim rs_Pry As New ADODB.Recordset
Dim rs_UNIDAD As New ADODB.Recordset
Attribute rs_UNIDAD.VB_VarHelpID = -1
Dim rs_CARGO As New ADODB.Recordset
Dim rs_Puesto_Org As New ADODB.Recordset

Dim rs_correlativo As New ADODB.Recordset

Dim e As Long
Dim DirCto As String
Dim var_cod As Integer
Dim VAR_VAL, IMG_CTR As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean


'Private Sub cmdAprueba_Click()
'  On Error GoTo UpdateErr
'   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'   If frmBeneficiario.Ado_Contrato!estado_contrato = "NO" Then
'      If sino = vbYes Then
'         frmBeneficiario.Ado_Contrato!estado_contrato = "SI"
'         frmBeneficiario.Ado_Contrato!fecha_registro = Date
'         frmBeneficiario.Ado_Contrato!usr_codigo = GlUsuario
'         frmBeneficiario.Ado_Contrato.UpdateBatch adAffectAll
'      End If
'   Else
'       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
'        frmBeneficiario.Ado_Contrato.Recordset.CancelUpdate
'        If mvBookMark > 0 Then
'          frmBeneficiario.Ado_Contrato.Recordset.Bookmark = mvBookMark
'        Else
'          frmBeneficiario.Ado_Contrato.Recordset.MoveFirst
'        End If
        mbDataChanged = False
        Unload Me
'        Fra_ABM.Enabled = False
'        fraOpciones.Visible = True
'        FraGrabarCancelar.Visible = False
'        DtG_Auxiliar.Enabled = True
    End If
End Sub

'Private Sub CmdDel_Click()
'  On Error GoTo UpdateErr
'   sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'   If frmBeneficiario.Ado_Contrato!estado_codigo = "S" Then
'      If sino = vbYes Then
'         frmBeneficiario.Ado_Contrato!estado_codigo = "L"
'         frmBeneficiario.Ado_Contrato!fecha_registro = Date
'         frmBeneficiario.Ado_Contrato!usr_codigo = GlUsuario
'         frmBeneficiario.Ado_Contrato.UpdateBatch adAffectAll
'      End If
'   Else
'      MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'End Sub

Private Sub cmdDesaprueba_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If frmBeneficiario.Ado_Contrato!estado_codigo = "S" Then
      If sino = vbYes Then
         frmBeneficiario.Ado_Contrato!estado_codigo = "N"
         frmBeneficiario.Ado_Contrato!fecha_registro = Date
         frmBeneficiario.Ado_Contrato!usr_codigo = glusuario
         frmBeneficiario.Ado_Contrato.Recordset.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
 'acepta las modificaciones realizadas
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    'If GlSW = "ADD" Then
    If txtSW = "ADD" Then '
      Dim tiempo2 As Integer
      frmBeneficiario.Ado_Contrato.Recordset!codigo_contrato = TxtCodigo.Text
      frmBeneficiario.Ado_Contrato.Recordset!beneficiario_codigo = txtBenef.Text
      frmBeneficiario.Ado_Contrato.Recordset!ges_gestion = glGestion
      frmBeneficiario.Ado_Contrato.Recordset!codigo_solicitud = frmBeneficiario.Ado_Contrato.Recordset.RecordCount
      TxtForm = frmBeneficiario.Ado_Contrato.Recordset!codigo_solicitud
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ro_contratos_personas WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "'  ", db, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            frmBeneficiario.Ado_Contrato.Recordset!numero_consultoria = rs_correlativo.RecordCount
'            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
'            rs_correlativo.Update
'            rs_M1!Numero_FA = rs_correlativo!correlativo
      Else
            frmBeneficiario.Ado_Contrato.Recordset!numero_consultoria = 1
      End If
      frmBeneficiario.Ado_Contrato.Recordset!ARCHIVO = "Cargar_Archivo"
      frmBeneficiario.Ado_Contrato.Recordset!ARCHIVO_NOMB = Trim(TxtInicial.Text) & "_Contrato_" & frmBeneficiario.Ado_Contrato.Recordset!numero_consultoria & ".pdf"
      TxtAprob.Text = "NO"
    End If
      frmBeneficiario.Ado_Contrato.Recordset!objeto_contrato = txtObjContrato.Text
      frmBeneficiario.Ado_Contrato.Recordset!fte_codigo = DTcFte.Text
      frmBeneficiario.Ado_Contrato.Recordset!codigo_unidad = dtc_codigo.Text
      frmBeneficiario.Ado_Contrato.Recordset!doc_codigo = DtcRespaldoCod.Text
      frmBeneficiario.Ado_Contrato.Recordset!org_codigo = DtcOrg.Text
      frmBeneficiario.Ado_Contrato.Recordset!codigo_cargo = DtcCargo.Text
      frmBeneficiario.Ado_Contrato.Recordset!pro_proyecto = DtcPry.Text
      frmBeneficiario.Ado_Contrato.Recordset!codigo_puesto = DtcPuesto.Text
      frmBeneficiario.Ado_Contrato.Recordset!fechas_confirmado = txtEstado
      frmBeneficiario.Ado_Contrato.Recordset!estado_contrato = TxtAprob
      If txtEstado.Text = "SI" Then
        frmBeneficiario.Ado_Contrato.Recordset!fecha_fin = Format("31/12/2020", "dd/mm/yyyy")
      Else
        frmBeneficiario.Ado_Contrato.Recordset!fecha_fin = DTPFFin.Value
      End If
      frmBeneficiario.Ado_Contrato.Recordset!fecha_firma = DTPFFirma.Value
      frmBeneficiario.Ado_Contrato.Recordset!fecha_inicio = DTPFInicio.Value
      
      frmBeneficiario.Ado_Contrato.Recordset!monto_totalbs = TxtBs.Text
      frmBeneficiario.Ado_Contrato.Recordset!tc_us = GlTipoCambioOficial
      If GlTipoCambioOficial > 0 Then
        frmBeneficiario.Ado_Contrato.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      Else
        GlTipoCambioOficial = 7.02
        frmBeneficiario.Ado_Contrato.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      End If
      frmBeneficiario.Ado_Contrato.Recordset!observacion_contrato = "-"
      frmBeneficiario.Ado_Contrato.Recordset!establece_multas = "N"
      frmBeneficiario.Ado_Contrato.Recordset!cod_forma_inicio = "1"
      tiempo2 = DTPFFin.Value - DTPFInicio.Value
      frmBeneficiario.Ado_Contrato.Recordset!tiempo_num = tiempo2
      frmBeneficiario.Ado_Contrato.Recordset!tiempo_dmy = "DIA"
      frmBeneficiario.Ado_Contrato.Recordset!tipo_moneda = "Bs"
      frmBeneficiario.Ado_Contrato.Recordset!org_codigo = "111"
      frmBeneficiario.Ado_Contrato.Recordset!porc_orgfin = 100
      frmBeneficiario.Ado_Contrato.Recordset!porc_contra = 0
      'frmBeneficiario.Ado_Contrato!fechas_confirmado = "N"
      frmBeneficiario.Ado_Contrato.Recordset!hora_registro = Format(Time, "HH:mm:ss")
      frmBeneficiario.Ado_Contrato.Recordset!fecha_registro = Date
      frmBeneficiario.Ado_Contrato.Recordset!usr_usuario = glusuario
      frmBeneficiario.Ado_Contrato.Recordset.Update 'Batch adAffectAll
      
      mbDataChanged = False
'      Call abrirtabla
      Unload Me
      'Fra_ABM.Enabled = False
'      fraOpciones.Visible = True
'      FraGrabarCancelar.Visible = False
'      DtG_Auxiliar.Enabled = True
  End If
  frmBeneficiario_Admin.abrirtabla
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If TxtCodigo.Text = "" Then
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

'Private Sub CmdMod_Click()
'  On Error GoTo EditErr
'  If Ado_Auxiliar.Recordset!estado_contrato = "SI" Then
'    MsgBox "No se puede modificar un registro APROBADO ...", vbCritical + vbExclamation, "Validación de datos"
'    Exit Sub
'  Else
''  lblStatus.Caption = "Modificar registro"
'    Fra_ABM.Enabled = True
'    fraOpciones.Visible = False
'    FraGrabarCancelar.Visible = True
'    DtG_Auxiliar.Enabled = False
'    GlSW = "MOD"
'    Exit Sub
'  End If
'
'
'EditErr:
'  MsgBox Err.Description
'End Sub
'
'Private Sub CmdSal_Click()
''  If glPersNew = "O" Then
''    frmmo_pacientes.Dtc_ocupac = frmBeneficiario.Ado_Contrato!ocup_codigo
''    frmmo_pacientes.Dtc_OcupacDes = frmBeneficiario.Ado_Contrato!ocup_descripcion
''  End If
''  glPersNew = "N"
'  Unload Me
'End Sub

Private Sub CmdVerDisco_Click()
  On Error GoTo Error_Sub
  
  If frmBeneficiario.Ado_Contrato.Recordset!ARCHIVO = "Cargar_Archivo" Then
     NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\"
     Frmexporta.DirDestino.Path = NombreCarpeta
     GlArch = "CTO"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
         DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(frmBeneficiario.Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\"
      Else
         DirCto = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = DirCto
     Frmexporta.Show vbModal
  Else
'    MsgBox ""
     sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
     If sino = vbYes Then
        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        GlArch = "CTO"
        'If GlServidor <> GlMaquina Then      ' "-" Then
        If GlServidor = "SRVPRO" Then
           DirCto = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(frmBeneficiario.Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\"
        Else
           DirCto = NombreCarpeta
        End If
        Frmexporta.DirDestino2.Path = DirCto
        Frmexporta.Show vbModal
     End If
  End If

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


Private Sub DtcCargo_Click(Area As Integer)
    DtcCargoDes.BoundText = DtcCargo.BoundText
    Call pCGO(DtcCargoDes.BoundText)
End Sub

Private Sub DtcCargoDes_Click(Area As Integer)
    DtcCargo.BoundText = DtcCargoDes.BoundText
    Call pCGO(DtcCargo.BoundText)
End Sub

Private Sub pCGO(CodCargo As String)
   Dim strConsulta As String
   
   strConsulta = "select * from rc_puestos where cargo_codigo = '" & CodCargo & "'"
   
   Set DtcPuesto.RowSource = Nothing
   Set DtcPuesto.RowSource = db.Execute(strConsulta, , adCmdText)
   DtcPuesto.ReFill
   DtcPuesto.BoundText = Empty
   
   Set DtcPuestoDes.RowSource = Nothing
   Set DtcPuestoDes.RowSource = db.Execute(strConsulta, , adCmdText)
   DtcPuestoDes.ReFill
   DtcPuestoDes.BoundText = Empty

End Sub

Private Sub DTcFte_Click(Area As Integer)
   DtcFteDes.BoundText = DTcFte.BoundText
   Call pOrganismo(DtcFteDes.BoundText)
End Sub

Private Sub DtcFteDes_Click(Area As Integer)
    DTcFte.BoundText = DtcFteDes.BoundText
    Call pOrganismo(DTcFte.BoundText)
End Sub

Private Sub pOrganismo(CodFuente As String)
   Dim strConsultaF As String
   strConsultaF = "select * from fc_organismo_financiamiento where fte_codigo='" & CodFuente & "'"
   Set DtcOrg.RowSource = Nothing
   Set DtcOrg.RowSource = db.Execute(strConsultaF, , adCmdText)
   DtcOrg.ReFill
   DtcOrg.BoundText = Empty
   Set DtcOrgDes.RowSource = Nothing
   Set DtcOrgDes.RowSource = db.Execute(strConsultaF, , adCmdText)
   DtcOrgDes.ReFill
   DtcOrgDes.BoundText = Empty
End Sub

Private Sub DtcOrg_Click(Area As Integer)
    DtcOrgDes.BoundText = DtcOrg.BoundText
    'Call pCat(DtcOrgDes.BoundText)
End Sub

Private Sub DtcOrgDes_Click(Area As Integer)
    DtcOrg.BoundText = DtcOrgDes.BoundText
    'Call pCat(DtcOrg.BoundText)
End Sub

Private Sub pCat(CodOrganismo As String)
   Dim strConsulta As String
   
   'strConsulta = "select * from fc_estructura_programatica where codigo_convenio='" & CodOrganismo & "'"
   strConsulta = "select * from fc_estructura_programatica where org_codigo ='" & CodOrganismo & "'"
   
   Set DtcPry.RowSource = Nothing
   Set DtcPry.RowSource = db.Execute(strConsulta, , adCmdText)
   DtcPry.ReFill
   DtcPry.BoundText = Empty
   
   Set DtcPryDes.RowSource = Nothing
   Set DtcPryDes.RowSource = db.Execute(strConsulta, , adCmdText)
   DtcPryDes.ReFill
   DtcPryDes.BoundText = Empty

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

Private Sub DtcRespaldo_Click(Area As Integer)
    DtcRespaldoCod.BoundText = DtcRespaldo.BoundText
End Sub

Private Sub DtcRespaldoCod_Click(Area As Integer)
    DtcRespaldo.BoundText = DtcRespaldoCod.BoundText
End Sub

Private Sub DTPFInicio_LostFocus()
    DTPFFirma.Value = DTPFInicio.Value
End Sub

Private Sub Form_Load()

'  Call abrirtabla
  
  Set rs_FteFin = New ADODB.Recordset
  rs_FteFin.Open "select * from fc_fuente_financiamiento WHERE estado_codigo = 'APR' ", db, adOpenKeyset, adLockOptimistic   'ORDER BY beneficiario_denominacion
  Set AdoFuente.Recordset = rs_FteFin.DataSource
  DtcFteDes.BoundText = DTcFte.BoundText
  
'  Set rs_Org = New ADODB.Recordset
'  rs_Org.Open "select * from fc_convenios  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoOrg.Recordset = rs_Org.DataSource
'  DtcOrgDes.BoundText = DtcOrg.BoundText
  
  Set rs_Org = New ADODB.Recordset
  rs_Org.Open "select * from fc_organismo_financiamiento  ", db, adOpenKeyset, adLockOptimistic
  Set AdoOrg.Recordset = rs_Org.DataSource
  DtcOrgDes.BoundText = DtcOrg.BoundText
  
  Set rs_Pry = New ADODB.Recordset
  rs_Pry.Open "select * from fc_estructura_programatica  ", db, adOpenKeyset, adLockOptimistic
  Set AdoPry.Recordset = rs_Pry.DataSource
  DtcPryDes.BoundText = DtcPry.BoundText
    
  Set rs_UNIDAD = New ADODB.Recordset
  rs_UNIDAD.Open "select * from gc_unidad_ejecutora  ", db, adOpenKeyset, adLockOptimistic
  Set AdoUnidad.Recordset = rs_UNIDAD.DataSource
  Dtc_descrip.BoundText = dtc_codigo.BoundText
  
  Set rs_CARGO = New ADODB.Recordset
  rs_CARGO.Open "select * from RC_CARGOS  ", db, adOpenKeyset, adLockOptimistic
  Set AdoCargo.Recordset = rs_CARGO.DataSource
  DtcCargoDes.BoundText = DtcCargo.BoundText
  If AdoCargo.Recordset.RecordCount > 0 Then
  End If
  
  Set rs_Puesto_Org = New ADODB.Recordset
  rs_Puesto_Org.Open "select * from rc_PUESTOs  ", db, adOpenKeyset, adLockOptimistic
  Set AdoPuestoOrg.Recordset = rs_Puesto_Org.DataSource
  DtcPuestoDes.BoundText = DtcPuesto.BoundText
  
  Set rs_Respaldo = New ADODB.Recordset
  rs_Respaldo.Open "select * from gc_documentos_respaldo where estado_codigo = 'APR'  ", db, adOpenKeyset, adLockOptimistic
  Set AdoRespaldo.Recordset = rs_Respaldo.DataSource
  DtcRespaldo.BoundText = DtcRespaldoCod.BoundText
  
'  frmBeneficiario.Ado_Contrato.AddNew
'  txtParam.Text = GlParametro
'  TxtForm.Text = GlForm
'  TxtCorrel.Text = GlCorrel

  mbDataChanged = False
'  Fra_ABM.Enabled = False
'  DtG_Auxiliar.Enabled = True
  GlSW = "NADA"
End Sub

'Private Sub abrirtabla()
'  Set frmBeneficiario.Ado_Contrato = New Recordset
'  If frmBeneficiario.Ado_Contrato.State = 1 Then frmBeneficiario.Ado_Contrato.Close
'  queryinicial = "select * from ro_contratos_personas "
'  frmBeneficiario.Ado_Contrato.Open queryinicial, DB, adOpenKeyset, adLockOptimistic
'  frmBeneficiario.Ado_Contrato.Sort = "beneficiario_codigo, codigo_unidad"
'  Set Ado_Auxiliar.Recordset = frmBeneficiario.Ado_Contrato.DataSource
'  Set DtG_Auxiliar.DataSource = Ado_Auxiliar.Recordset
'End Sub

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
    If Ado_Auxiliar.Recordset("estado_contrato") = "SI" Then
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

'Private Sub cmdAdd_Click()
'  On Error GoTo AddErr
'    'frmBeneficiario.Ado_Contrato.MoveLast
'    frmBeneficiario.Ado_Contrato.AddNew
'    'lblStatus.Caption = "Agregar registro"
'    Fra_ABM.Enabled = True
'    fraOpciones.Visible = False
'    FraGrabarCancelar.Visible = True
'    DtG_Auxiliar.Enabled = False
'    GlSW = "ADD"
''    frmBeneficiario.Ado_Contrato.AddNew
''    txtParam.Text = GlParametro
''    TxtForm.Text = "E-1" 'GlForm
''    TxtCorrel.Text = 1  'GlCorrel
'  Exit Sub
'AddErr:
'  MsgBox Err.Description
'End Sub

Private Sub cmdRefresh_Click()
 If lblARCH.Caption = "Cargar_Archivo" Then
    MsgBox ("No Existe el Archivo Asociado al Contrato, debe Cargarlo ...")
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
    If GlServidor = "SRVPRO" Then
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(frmBeneficiario.Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.Ado_Contrato.Recordset!beneficiario_codigo) & "\CONTRATOS\" & Trim(frmBeneficiario.Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub

Private Sub Txtestado_LostFocus()
    If txtEstado.Text = "SI" Then
        DTPFFin.Enabled = False
    Else
        DTPFFin.Enabled = True
    End If
End Sub

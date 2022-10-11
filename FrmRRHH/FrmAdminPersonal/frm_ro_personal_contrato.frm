VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ro_personal_contrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administracion de Personal - Ficha Personal - Contratos, Adendas o Designaciones"
   ClientHeight    =   7515
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   9360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ro_personal_contrato.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ro_personal_contrato.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   9075
      TabIndex        =   35
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "frm_ro_personal_contrato.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frm_ro_personal_contrato.frx":D665A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Contrato"
         Height          =   680
         Left            =   2160
         Picture         =   "frm_ro_personal_contrato.frx":D6864
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Carga Contrato"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   3000
         Picture         =   "frm_ro_personal_contrato.frx":D6BEC
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Carga Contrato"
         Top             =   120
         Width           =   720
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
         TabIndex        =   40
         Top             =   240
         Width           =   4935
      End
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
      Left            =   2040
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
      Left            =   4080
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
      Left            =   6120
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
      Left            =   0
      Top             =   7320
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
      Height          =   6015
      Left            =   135
      TabIndex        =   14
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox txt_time 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "tiempo_num"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
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
         Left            =   7920
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txt_otro_bs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_otroBS"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
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
         Left            =   4680
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox txtMensual_bs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_mensualBS"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
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
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   5760
         MaxLength       =   80
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   4080
         MaxLength       =   80
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtAprob 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "estado_contrato"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5000
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "NO"
         Top             =   520
         Width           =   600
      End
      Begin MSDataListLib.DataCombo Dtc_descrip 
         Bindings        =   "frm_ro_personal_contrato.frx":D6F74
         DataField       =   "unidad_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   4680
         TabIndex        =   7
         Top             =   2640
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D6F8C
         DataField       =   "pro_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "pro_descripcion"
         BoundColumn     =   "pro_codigo"
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
         DataField       =   "estado_confirmado"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
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
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
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
         DataField       =   "fecha_firma"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   91029505
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
         Top             =   1600
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
         Left            =   7800
         TabIndex        =   15
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtBs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_totalBS"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
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
         Left            =   7575
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   5400
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcPuestoDes 
         Bindings        =   "frm_ro_personal_contrato.frx":D6FA1
         DataField       =   "puesto_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   4080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "puesto_descripcion"
         BoundColumn     =   "puesto_codigo"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D6FBC
         DataField       =   "puesto_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   6360
         TabIndex        =   17
         Top             =   3795
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "puesto_codigo"
         BoundColumn     =   "puesto_codigo"
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
         DataField       =   "fecha_inicio"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   4800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   91029505
         CurrentDate     =   40471
      End
      Begin MSComCtl2.DTPicker DTPFFin 
         DataField       =   "fecha_fin"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
         Height          =   285
         Left            =   4680
         TabIndex        =   11
         Top             =   4800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   91029505
         CurrentDate     =   44196
      End
      Begin MSDataListLib.DataCombo Dtc_codigo 
         Bindings        =   "frm_ro_personal_contrato.frx":D6FD7
         DataField       =   "unidad_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   6360
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D6FEF
         DataField       =   "org_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D7004
         DataField       =   "org_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   21
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
         Bindings        =   "frm_ro_personal_contrato.frx":D7019
         DataField       =   "pro_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   22
         Top             =   3795
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "pro_codigo"
         BoundColumn     =   "pro_codigo"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D702E
         DataField       =   "Fte_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "fte_descripcion"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D7046
         DataField       =   "Fte_codigo"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   30
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
         Bindings        =   "frm_ro_personal_contrato.frx":D705E
         DataField       =   "cargo_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   3345
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "cargo_descripcion"
         BoundColumn     =   "cargo_codigo"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D7075
         DataField       =   "cargo_codigo"
         DataSource      =   "frmBeneficiario_Admin.Ado_Contrato"
         Height          =   315
         Left            =   6360
         TabIndex        =   26
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "cargo_codigo"
         BoundColumn     =   "cargo_codigo"
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
         Bindings        =   "frm_ro_personal_contrato.frx":D708C
         DataField       =   "doc_codigo"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
         Height          =   315
         Left            =   2040
         TabIndex        =   33
         Top             =   960
         Width           =   4215
         _ExtentX        =   7435
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
         Bindings        =   "frm_ro_personal_contrato.frx":D70A6
         DataField       =   "doc_codigo"
         DataSource      =   "frmBeneficiario_admin.Ado_Contrato"
         Height          =   315
         Left            =   4800
         TabIndex        =   34
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
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Sueldo Mensual                                      Refrigerio/Otro                                   Total Contrato"
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
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   5400
         Width           =   7320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFC0&
         X1              =   0
         X2              =   9120
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFC0&
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   2280
         Y2              =   4560
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Documento                                                                                               Nro. Correlativo"
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
         Index           =   11
         Left            =   240
         TabIndex        =   32
         Top             =   1005
         Width           =   7455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFC0&
         X1              =   0
         X2              =   9120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Proyecto                                                                                    Puesto que Ocupa "
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
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   3825
         Width           =   6300
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Organismo Financiador                                                       Cargo que Ocupa"
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
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   3105
         Width           =   6165
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fuente Financiamiento                                                        Unidad Organizacional"
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
         Index           =   28
         Left            =   120
         TabIndex        =   29
         Top             =   2385
         Width           =   6600
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
         TabIndex        =   23
         Top             =   555
         Width           =   75
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Código Contrato                       Vigente              Estado                   Nombre de Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         TabIndex        =   20
         Top             =   285
         Width           =   8640
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Inicio                                            Fecha de Fin                                        Tiempo (Meses)"
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
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   4800
         Width           =   7800
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         Caption         =   "Objeto del Contrato"
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
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   1335
         Width           =   1770
      End
   End
   Begin MSAdodcLib.Adodc AdoFuente 
      Height          =   330
      Left            =   2040
      Top             =   7320
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
      Top             =   7320
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
      Top             =   7320
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
Attribute VB_Name = "frm_ro_personal_contrato"
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
         frmBeneficiario.Ado_Contrato!Fecha_Registro = Date
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
      rw_ficha_rrhh.Ado_Contrato.Recordset!beneficiario_codigo = txtBenef.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!ges_gestion = glGestion
      rw_ficha_rrhh.Ado_Contrato.Recordset!solicitud_codigo = rw_ficha_rrhh.Ado_Contrato.Recordset.RecordCount
      TxtForm = rw_ficha_rrhh.Ado_Contrato.Recordset!solicitud_codigo
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ro_contratos_personas WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "'  ", db, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            rw_ficha_rrhh.Ado_Contrato.Recordset!numero_consultoria = rs_correlativo.RecordCount
'            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
'            rs_correlativo.Update
'            rs_M1!Numero_FA = rs_correlativo!correlativo
      Else
            rw_ficha_rrhh.Ado_Contrato.Recordset!numero_consultoria = 1
      End If
      rw_ficha_rrhh.Ado_Contrato.Recordset!ARCHIVO = "Cargar_Archivo"
      rw_ficha_rrhh.Ado_Contrato.Recordset!ARCHIVO_NOMB = Trim(TxtInicial.Text) & "_Contrato_" & rw_ficha_rrhh.Ado_Contrato.Recordset!numero_consultoria & ".pdf"
      TxtAprob.Text = "REG"
    End If
      rw_ficha_rrhh.Ado_Contrato.Recordset!codigo_contrato = txtCodigo.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!objeto_contrato = txtObjContrato.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!fte_codigo = DTcFte.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!unidad_codigo = dtc_codigo.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!doc_codigo = DtcRespaldoCod.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!org_codigo = DtcOrg.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!cargo_codigo = DtcCargo.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!pro_codigo = DtcPry.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!puesto_codigo = DtcPuesto.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!estado_confirmado = txtEstado
      rw_ficha_rrhh.Ado_Contrato.Recordset!estado_contrato = TxtAprob
      If rw_ficha_rrhh.Ado_Contrato.Recordset!fecha_fin >= Date Then
        txtEstado.Text = "SI"
      Else
        txtEstado.Text = "NO"
      End If
      'If Txtestado.Text = "SI" Then
      '  rw_ficha_rrhh.Ado_Contrato.Recordset!fecha_fin = Format("31/12/2020", "dd/mm/yyyy")
      'Else
        rw_ficha_rrhh.Ado_Contrato.Recordset!fecha_fin = DTPFFin.Value
      'End If
      rw_ficha_rrhh.Ado_Contrato.Recordset!fecha_firma = DTPFFirma.Value
      rw_ficha_rrhh.Ado_Contrato.Recordset!fecha_inicio = DTPFInicio.Value
      
      rw_ficha_rrhh.Ado_Contrato.Recordset!monto_totalbs = TxtBs.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!monto_mensualBS = txtMensual_bs.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!monto_otroBS = txt_otro_bs.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!tiempo_num = txt_time.Text
      rw_ficha_rrhh.Ado_Contrato.Recordset!tc_us = GlTipoCambioOficial
      If GlTipoCambioOficial > 0 Then
        rw_ficha_rrhh.Ado_Contrato.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      Else
        GlTipoCambioOficial = 6.96
        rw_ficha_rrhh.Ado_Contrato.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      End If
      rw_ficha_rrhh.Ado_Contrato.Recordset!observacion_contrato = "REGISTRO ELABORADO"
      rw_ficha_rrhh.Ado_Contrato.Recordset!establece_multas = "N"
      rw_ficha_rrhh.Ado_Contrato.Recordset!cod_forma_inicio = "1"
      tiempo2 = DTPFFin.Value - DTPFInicio.Value
      rw_ficha_rrhh.Ado_Contrato.Recordset!tiempo_num = tiempo2
      rw_ficha_rrhh.Ado_Contrato.Recordset!tiempo_dmy = "DIA"
      rw_ficha_rrhh.Ado_Contrato.Recordset!tipo_moneda = "Bs"
      'rw_ficha_rrhh.Ado_Contrato.Recordset!org_codigo = "111"
      rw_ficha_rrhh.Ado_Contrato.Recordset!porc_orgfin = 100
      rw_ficha_rrhh.Ado_Contrato.Recordset!porc_contra = 0
      'rw_ficha_rrhh.Ado_Contrato!estado_confirmado = "N"
      rw_ficha_rrhh.Ado_Contrato.Recordset!hora_registro = Format(Time, "HH:mm:ss")
      rw_ficha_rrhh.Ado_Contrato.Recordset!Fecha_Registro = Date
      rw_ficha_rrhh.Ado_Contrato.Recordset!usr_usuario = glusuario
      rw_ficha_rrhh.Ado_Contrato.Recordset.Update 'Batch adAffectAll
      
      mbDataChanged = False
'      Call abrirtabla
 rw_ficha_rrhh.abrirtabla
      Unload Me
      'Fra_ABM.Enabled = False
'      fraOpciones.Visible = True
'      FraGrabarCancelar.Visible = False
'      DtG_Auxiliar.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If txtCodigo.Text = "" Then
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
  rs_Respaldo.Open "select * from gc_documentos_respaldo where doc_copia_var = 'RRHH'  ", db, adOpenKeyset, adLockOptimistic
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
'  frmBeneficiario.Ado_Contrato.Sort = "beneficiario_codigo, unidad_codigo"
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

Private Sub txt_otro_bs_LostFocus()
    If txt_otro_bs.Text = "" Then
        txt_otro_bs.Text = "0"
    End If
    If txt_time.Text = "" Or txt_time.Text = "0" Then
        txt_time.Text = "1"
    End If
    TxtBs.Text = (CDbl(txtMensual_bs.Text) + CDbl(txt_otro_bs.Text)) * CDbl(txt_time.Text)
End Sub

Private Sub Txtestado_LostFocus()
    If txtEstado.Text = "SI" Then
        DTPFFin.Enabled = False
    Else
        DTPFFin.Enabled = True
    End If
End Sub

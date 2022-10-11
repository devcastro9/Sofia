VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ro_movilidad_personal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Personal - Ficha Personal - Movilidad de Personal"
   ClientHeight    =   7665
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10500
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraIntercambio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2380
      Left            =   5280
      TabIndex        =   56
      Top             =   4095
      Visible         =   0   'False
      Width           =   5145
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   70
         Top             =   2010
         Width           =   360
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   64
         Top             =   855
         Width           =   360
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   60
         Top             =   1440
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dtc_beneficiario_den 
         Bindings        =   "frm_ro_movilidad_personal.frx":0000
         DataField       =   "beneficiario_denominacion"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo dtc_puesto_den_int 
         Bindings        =   "frm_ro_movilidad_personal.frx":001E
         DataField       =   "puesto_codigo"
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Top             =   1995
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   16777215
         ListField       =   "puesto_descripcion"
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
      Begin MSDataListLib.DataCombo dtc_puesto_cod_int 
         Bindings        =   "frm_ro_movilidad_personal.frx":003C
         DataField       =   "puesto_codigo"
         Height          =   315
         Left            =   3240
         TabIndex        =   59
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "puesto_codigo"
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
      Begin MSDataListLib.DataCombo dtc_cargo_den_int 
         Bindings        =   "frm_ro_movilidad_personal.frx":005A
         DataField       =   "cargo_codigo"
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   1425
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   16777215
         ListField       =   "cargo_descripcion"
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
      Begin MSDataListLib.DataCombo dtc_beneficiario_cod 
         Bindings        =   "frm_ro_movilidad_personal.frx":0078
         DataField       =   "beneficiario_codigo"
         Height          =   315
         Left            =   2880
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo dtc_cargo_cod_int 
         Bindings        =   "frm_ro_movilidad_personal.frx":0096
         DataField       =   "cargo_codigo"
         Height          =   315
         Left            =   3360
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "cargo_codigo"
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
      Begin MSDataListLib.DataCombo dtc_unidad_den_int 
         Bindings        =   "frm_ro_movilidad_personal.frx":00B4
         DataField       =   "unidad_codigo"
         Height          =   315
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
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
      Begin MSDataListLib.DataCombo dtc_unidad_cod_int 
         Bindings        =   "frm_ro_movilidad_personal.frx":00D2
         DataField       =   "unidad_codigo"
         Height          =   315
         Left            =   3600
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "unidad_codigo"
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
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Puesto"
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
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   69
         Top             =   1770
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo "
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
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora "
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
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   1560
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   10275
      TabIndex        =   41
      Top             =   120
      Width           =   10335
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "frm_ro_movilidad_personal.frx":00F0
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   73
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   120
         Width           =   1245
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "frm_ro_movilidad_personal.frx":08C6
         ScaleHeight     =   615
         ScaleWidth      =   1365
         TabIndex        =   72
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton btn_intercambio 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Intercambiar puesto"
         Height          =   675
         Left            =   2760
         Picture         =   "frm_ro_movilidad_personal.frx":11B2
         TabIndex        =   54
         Top             =   120
         Width           =   1005
      End
      Begin VB.CommandButton btn_asignar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambiar puesto"
         Height          =   675
         Left            =   2760
         Picture         =   "frm_ro_movilidad_personal.frx":13BC
         TabIndex        =   55
         Top             =   120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Contrato"
         Height          =   680
         Left            =   3840
         Picture         =   "frm_ro_movilidad_personal.frx":15C6
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Carga Contrato"
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   4680
         Picture         =   "frm_ro_movilidad_personal.frx":194E
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Carga Contrato"
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOVILIDAD DE PERSONAL"
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
         Left            =   5745
         TabIndex        =   44
         Top             =   240
         Width           =   4065
      End
   End
   Begin MSAdodcLib.Adodc AdoBeneficiario 
      Height          =   330
      Left            =   120
      Top             =   6840
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
      Top             =   6840
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
      Top             =   6840
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
      Top             =   6840
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
      Left            =   120
      Top             =   7200
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
      BackColor       =   &H00C0C0C0&
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
      Height          =   5655
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   10320
      Begin VB.TextBox TxtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "numero_cambio"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   525
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         DataField       =   "numero_cambio"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
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
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   520
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         DataField       =   "numero_cambio"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
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
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   74
         Text            =   "frm_ro_movilidad_personal.frx":1CD6
         Top             =   520
         Width           =   375
      End
      Begin VB.TextBox txt_tipo_mov 
         Height          =   285
         Left            =   6720
         MaxLength       =   80
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   53
         Top             =   4090
         Width           =   365
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   52
         Top             =   3360
         Width           =   365
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   51
         Top             =   2660
         Width           =   365
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   50
         Top             =   3380
         Width           =   365
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   49
         Top             =   2660
         Width           =   365
      End
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   6000
         MaxLength       =   80
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   4320
         MaxLength       =   80
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   2640
         MaxLength       =   80
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtAprob 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "estado_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5445
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "REG"
         Top             =   520
         Width           =   495
      End
      Begin MSDataListLib.DataCombo Dtc_descrip 
         Bindings        =   "frm_ro_movilidad_personal.frx":1CDB
         DataField       =   "unidad_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Top             =   2640
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcPryDes 
         Bindings        =   "frm_ro_movilidad_personal.frx":1CF3
         DataField       =   "puesto_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   4080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         ListField       =   "puesto_descripcion"
         BoundColumn     =   "puesto_codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox Txtestado 
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Text            =   "SI"
         Top             =   520
         Visible         =   0   'False
         Width           =   660
      End
      Begin MSComCtl2.DTPicker DTPFaprobacion 
         DataField       =   "fecha_aprobacion"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   285
         Left            =   5640
         TabIndex        =   10
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   114098177
         CurrentDate     =   40471
      End
      Begin VB.TextBox txtObjContrato 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "Observaciones"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1560
         Width           =   9735
      End
      Begin VB.TextBox TxtForm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "numero_resolucion"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
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
         Left            =   8760
         TabIndex        =   17
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtBs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "item"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
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
         Left            =   8535
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DtcPuestoDes 
         Bindings        =   "frm_ro_movilidad_personal.frx":1D08
         DataField       =   "puesto_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   5280
         TabIndex        =   7
         Top             =   4080
         Width           =   4815
         _ExtentX        =   8493
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
         Bindings        =   "frm_ro_movilidad_personal.frx":1D23
         DataField       =   "puesto_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   8160
         TabIndex        =   21
         Top             =   3675
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "puesto_codigo"
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
      Begin MSComCtl2.DTPicker DTPFelaboracion 
         DataField       =   "fecha_elaboracion"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   5040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   114098177
         CurrentDate     =   40471
      End
      Begin MSComCtl2.DTPicker DTPFcontrato 
         DataField       =   "fecha_inicio_contrato"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   5040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   114098177
         CurrentDate     =   44196
      End
      Begin MSDataListLib.DataCombo Dtc_codigo 
         Bindings        =   "frm_ro_movilidad_personal.frx":1D3E
         DataField       =   "unidad_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   7920
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "unidad_codigo"
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
      Begin MSDataListLib.DataCombo DtcOrgDes 
         Bindings        =   "frm_ro_movilidad_personal.frx":1D56
         DataField       =   "cargo_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   3360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         ListField       =   "cargo_descripcion"
         BoundColumn     =   "cargo_codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcOrg 
         Bindings        =   "frm_ro_movilidad_personal.frx":1D6B
         DataField       =   "cargo_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   3360
         TabIndex        =   26
         Top             =   3075
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
      Begin MSDataListLib.DataCombo DtcPry 
         Bindings        =   "frm_ro_movilidad_personal.frx":1D80
         DataField       =   "puesto_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   3000
         TabIndex        =   27
         Top             =   3675
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "puesto_codigo"
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
      Begin MSDataListLib.DataCombo Dtc_descrip_ant 
         Bindings        =   "frm_ro_movilidad_personal.frx":1D95
         DataField       =   "unidad_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   2640
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dtc_codigo_ant 
         Bindings        =   "frm_ro_movilidad_personal.frx":1DB1
         DataField       =   "unidad_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   3840
         TabIndex        =   36
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "unidad_codigo"
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
      Begin MSDataListLib.DataCombo DtcCargoDes 
         Bindings        =   "frm_ro_movilidad_personal.frx":1DC9
         DataField       =   "cargo_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   3345
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         ListField       =   "cargo_descripcion"
         BoundColumn     =   "cargo_codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcCargo 
         Bindings        =   "frm_ro_movilidad_personal.frx":1DE0
         DataField       =   "cargo_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   7920
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "cargo_codigo"
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
      Begin MSDataListLib.DataCombo DtcRespaldo 
         Bindings        =   "frm_ro_movilidad_personal.frx":1DF7
         DataField       =   "tipo_memo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "descripcion"
         BoundColumn     =   "tipo_memo"
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
         Bindings        =   "frm_ro_movilidad_personal.frx":1E11
         DataField       =   "tipo_memo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   4800
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "tipo_memo"
         BoundColumn     =   "tipo_memo"
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
      Begin MSDataListLib.DataCombo DtcPryCargo 
         Bindings        =   "frm_ro_movilidad_personal.frx":1E2B
         DataField       =   "puesto_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   2160
         TabIndex        =   45
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "cargo_codigo"
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
      Begin MSDataListLib.DataCombo DtcPryUni 
         Bindings        =   "frm_ro_movilidad_personal.frx":1E40
         DataField       =   "puesto_anterior"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   3840
         TabIndex        =   46
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "unidad_codigo"
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
      Begin MSDataListLib.DataCombo DtcPuestoCargo 
         Bindings        =   "frm_ro_movilidad_personal.frx":1E55
         DataField       =   "puesto_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   7320
         TabIndex        =   47
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "cargo_codigo"
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
      Begin MSDataListLib.DataCombo DtcPuestoUni 
         Bindings        =   "frm_ro_movilidad_personal.frx":1E70
         DataField       =   "puesto_codigo"
         DataSource      =   "frmBeneficiario_control.AdoMovilidad"
         Height          =   315
         Left            =   6360
         TabIndex        =   48
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "unidad_codigo"
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
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFC0&
         X1              =   0
         X2              =   10320
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFC0&
         BorderWidth     =   2
         X1              =   5160
         X2              =   5160
         Y1              =   2280
         Y2              =   4560
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Documento:"
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
         Index           =   11
         Left            =   240
         TabIndex        =   39
         Top             =   1005
         Width           =   1830
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFC0&
         X1              =   0
         X2              =   10320
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Archivo"
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
         Index           =   10
         Left            =   7710
         TabIndex        =   38
         Top             =   300
         Width           =   1740
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Puesto que Ocupa (ORIGEN)                                                       Puesto Nuevo que Ocupará (DESTINO)"
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
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   3840
         Width           =   8640
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo que Ocupa (ORIGEN)                                                        Cargo Nuevo que Ocupará (DESTINO)"
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
         Index           =   8
         Left            =   240
         TabIndex        =   37
         Top             =   3120
         Width           =   8535
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora (ORIGEN)                                                         Unidad Ejecutora (DESTINO)"
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
         Index           =   28
         Left            =   240
         TabIndex        =   35
         Top             =   2400
         Width           =   7695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Index           =   12
         Left            =   5445
         TabIndex        =   29
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblARCH 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   9870
         TabIndex        =   28
         Top             =   555
         Width           =   75
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Memo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
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
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Elaboracion "
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
         Index           =   7
         Left            =   360
         TabIndex        =   23
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Height          =   255
         Index           =   6
         Left            =   8640
         TabIndex        =   22
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Vigente"
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
         Index           =   5
         Left            =   3600
         TabIndex        =   20
         Top             =   300
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Objeto del Proceso"
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
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1335
         Width           =   1890
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Correlativo:"
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
         Left            =   7200
         TabIndex        =   18
         Top             =   1005
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Aprobacion"
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
         Index           =   0
         Left            =   5640
         TabIndex        =   16
         Top             =   4800
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Reasignación  "
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
         Index           =   1
         Left            =   2760
         TabIndex        =   15
         Top             =   4800
         Width           =   1965
      End
   End
   Begin MSAdodcLib.Adodc adounidad_ant 
      Height          =   330
      Left            =   2160
      Top             =   7200
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
      Caption         =   "adounidad_ant"
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
      Left            =   4200
      Top             =   7200
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
      Left            =   6240
      Top             =   7200
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
   Begin MSAdodcLib.Adodc ado_intercambio 
      Height          =   330
      Left            =   8280
      Top             =   6840
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
      Caption         =   "ado_intercambio"
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
Attribute VB_Name = "frm_ro_movilidad_personal"
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

Dim rs_numero_int As New ADODB.Recordset

Dim rs_correlativo As New ADODB.Recordset

Dim rs_intercambio As New ADODB.Recordset

Dim e As Long
Dim DirCto, tipo_cam As String
Dim var_cod, numero_int As Integer
Dim VAR_VAL, IMG_CTR As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean


Private Sub btn_asignar_Click()
btn_intercambio.Visible = True
btn_asignar.Visible = False
FraIntercambio.Visible = False
tipo_cam = "CAMBIO"
End Sub

Private Sub btn_intercambio_Click()
btn_intercambio.Visible = False
btn_asignar.Visible = True
FraIntercambio.Visible = True

tipo_cam = "INTERCAMBIO"
End Sub

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
    If txtSW = "ADD" Then
        rw_ficha_rrhh.AdoMovilidad.Recordset.AddNew  '
        Dim tiempo2 As Integer
        rw_ficha_rrhh.AdoMovilidad.Recordset!ges_gestion = Year(DTPFcontrato.Value)
        rw_ficha_rrhh.AdoMovilidad.Recordset!numero_cambio = TxtCodigo.Text
    End If
    TxtCodigo.Enabled = False
      
    If tipo_cam = "INTERCAMBIO" Then
        Set rs_numero_int = New ADODB.Recordset
        rs_numero_int.Open "select max(numero_intercambio) as numero_int from ro_movilidad_personal", db, adOpenKeyset, adLockOptimistic
        If rs_numero_int!numero_int > 0 Then
            numero_int = rs_numero_int!numero_int + 1
        Else
            numero_int = 1
        End If
        rw_ficha_rrhh.AdoMovilidad.Recordset!tipo_mov = tipo_cam
        rw_ficha_rrhh.AdoMovilidad.Recordset!beneficiario_codigo_int = dtc_beneficiario_cod.Text
        rw_ficha_rrhh.AdoMovilidad.Recordset!unidad_codigo = dtc_unidad_cod_int.Text
        rw_ficha_rrhh.AdoMovilidad.Recordset!cargo_codigo = dtc_cargo_cod_int.Text
        rw_ficha_rrhh.AdoMovilidad.Recordset!puesto_codigo = dtc_puesto_cod_int.Text
        If txtSW = "ADD" Then
            rw_ficha_rrhh.AdoMovilidad.Recordset!numero_intercambio = numero_int
            db.Execute "INSERT ro_movilidad_personal (ges_gestion,            beneficiario_codigo,              numero_cambio,              tipo_memo,          fecha_elaboracion,         fecha_inicio_contrato,     unidad_codigo,                                                               cargo_codigo,                                                   puesto_codigo,        unidad_anterior,                  puesto_anterior,                  cargo_anterior,                  Observaciones,    estado_codigo, fecha_registro,        hora_registro,         usr_codigo,            tipo_mov,          beneficiario_codigo_int, numero_intercambio)" & _
                           "VALUES('" & Year(DTPFcontrato.Value) & "' , '" & dtc_beneficiario_cod.Text & "','" & TxtCodigo.Text & " ','" & DtcRespaldoCod.Text & "','" & Date & "' ,'" & DTPFcontrato.Value & "','" & IIf(DtcPryUni.Text = "", Dtc_codigo_ant.Text, DtcPryUni.Text) & " ','" & IIf(DtcPryCargo.Text = "", DtcOrg.Text, DtcPryCargo.Text) & "','" & DtcPry.Text & "','" & dtc_unidad_cod_int.Text & "','" & dtc_puesto_cod_int.Text & "','" & dtc_cargo_cod_int.Text & "','" & txtObjContrato.Text & "','REG','" & Date & "','" & Format(Time, "HH:mm:ss") & "','" & glusuario & "','" & tipo_cam & "','" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "' ,'" & numero_int & "')"
        Else
            db.Execute "UPDATE ro_movilidad_personal SET ges_gestion = '" & Year(DTPFcontrato.Value) & "', beneficiario_codigo = '" & dtc_beneficiario_cod.Text & "', tipo_memo = '" & DtcRespaldoCod.Text & "', fecha_elaboracion = '" & Date & "', fecha_inicio_contrato = '" & DTPFcontrato.Value & "', unidad_codigo = '" & IIf(DtcPryUni.Text = "", Dtc_codigo_ant.Text, DtcPryUni.Text) & "',cargo_codigo = '" & IIf(DtcPryCargo.Text = "", DtcOrg.Text, DtcPryCargo.Text) & "', puesto_codigo = '" & DtcPry.Text & "', unidad_anterior = '" & dtc_unidad_cod_int.Text & "', puesto_anterior = '" & dtc_puesto_cod_int.Text & "', cargo_anterior = '" & dtc_cargo_cod_int.Text & "', Observaciones = '" & txtObjContrato.Text & "', estado_codigo = 'REG', fecha_registro = '" & Date & "', hora_registro = '" & Format(Time, "HH:mm:ss") & "', usr_codigo = '" & glusuario & "' WHERE beneficiario_codigo_int = '" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "' AND numero_intercambio = " & _
            rw_ficha_rrhh.AdoMovilidad.Recordset!numero_intercambio & ""
            rw_ficha_rrhh.AdoMovilidad.Recordset!numero_cambio = TxtCodigo.Text
            rw_ficha_rrhh.AdoMovilidad.Recordset!ges_gestion = Year(DTPFcontrato.Value)
        End If
    Else
        rw_ficha_rrhh.AdoMovilidad.Recordset!tipo_mov = "CAMBIO"
        rw_ficha_rrhh.AdoMovilidad.Recordset!beneficiario_codigo_int = "0"
        rw_ficha_rrhh.AdoMovilidad.Recordset!unidad_codigo = IIf(DtcPuestoUni.Text = "", dtc_codigo.Text, DtcPuestoUni.Text)
        rw_ficha_rrhh.AdoMovilidad.Recordset!cargo_codigo = IIf(DtcPuestoCargo.Text = "", DtcCargo.Text, DtcPuestoCargo.Text)
        rw_ficha_rrhh.AdoMovilidad.Recordset!puesto_codigo = DtcPuesto.Text
    End If
    'rw_ficha_rrhh.AdoMovilidad.Recordset!tipo_memo = DtcRespaldoCod.Text
    rw_ficha_rrhh.AdoMovilidad.Recordset!observaciones = txtObjContrato.Text
    rw_ficha_rrhh.AdoMovilidad.Recordset!unidad_anterior = IIf(DtcPryUni.Text = "", Dtc_codigo_ant.Text, DtcPryUni.Text)
    rw_ficha_rrhh.AdoMovilidad.Recordset!cargo_anterior = IIf(DtcPryCargo.Text = "", DtcOrg.Text, DtcPryCargo.Text)
    rw_ficha_rrhh.AdoMovilidad.Recordset!puesto_anterior = DtcPry.Text
    
    'rw_ficha_rrhh.AdoMovilidad.Recordset!fecha_elaboracion = Date
    rw_ficha_rrhh.AdoMovilidad.Recordset!fecha_inicio_contrato = DTPFcontrato.Value    'Date
    rw_ficha_rrhh.AdoMovilidad.Recordset!fecha_aprobacion = DTPFcontrato.Value          'Date
    rw_ficha_rrhh.AdoMovilidad.Recordset!Item = "0"     'TxtBs.Text
    
    rw_ficha_rrhh.AdoMovilidad.Recordset!beneficiario_codigo = rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo
    
    rw_ficha_rrhh.AdoMovilidad.Recordset!ges_gestion = Year(DTPFcontrato.Value)
    rw_ficha_rrhh.AdoMovilidad.Recordset!hora_registro = Format(Time, "HH:mm:ss")
    rw_ficha_rrhh.AdoMovilidad.Recordset!fecha_registro = Date
    rw_ficha_rrhh.AdoMovilidad.Recordset!usr_codigo = glusuario
    rw_ficha_rrhh.AdoMovilidad.Recordset!codigo_empresa = rw_ficha_rrhh.Ado_datos.Recordset!codigo_empresa
    rw_ficha_rrhh.AdoMovilidad.Recordset.Update 'Batch adAffectAll

'      mbDataChanged = False
'      Call abrirtabla
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
  If TxtCodigo.Text = "" Then
    MsgBox "Debe registrar el Código o Cite del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
'  If TxtBs.Text = "" Then
'    MsgBox "Debe registrar el Item ...", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  
  If DTPFaprobacion.Value > DTPFaprobacion.Value Then
    MsgBox "La Fecha de Aprobacion NO puede ser Mayor a la de Inicio del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If DTPFelaboracion.Value > DTPFelaboracion.Value Then
    MsgBox "La Fecha de Elaboracion NO puede ser Mayor a la de Finalizacion del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
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

Private Sub DataCombo2_Click(Area As Integer)

End Sub

Private Sub dtc_beneficiario_cod_Click(Area As Integer)
dtc_beneficiario_den.BoundText = dtc_beneficiario_cod.BoundText
End Sub

Private Sub dtc_beneficiario_den_Change()
dtc_beneficiario_cod.BoundText = dtc_beneficiario_den.BoundText
dtc_unidad_den_int.BoundText = dtc_beneficiario_cod.BoundText
dtc_unidad_cod_int.BoundText = dtc_beneficiario_cod.BoundText

dtc_cargo_den_int.BoundText = dtc_beneficiario_cod.BoundText
dtc_cargo_cod_int.BoundText = dtc_beneficiario_cod.BoundText

dtc_puesto_den_int.BoundText = dtc_beneficiario_cod.BoundText
dtc_puesto_cod_int.BoundText = dtc_beneficiario_cod.BoundText
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
    'Call pCGO(DtcCargo.BoundText)
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

Private Sub Dtc_codigo_ant_Click(Area As Integer)
   Dtc_descrip_ant.BoundText = Dtc_codigo_ant.BoundText
   'Call pOrganismo(Dtc_descrip_ant.BoundText)
End Sub

Private Sub Dtc_descrip_ant_Click(Area As Integer)
    Dtc_codigo_ant.BoundText = Dtc_descrip_ant.BoundText
    'Call pOrganismo(Dtc_codigo_ant.BoundText)
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
    DtcPryCargo.BoundText = DtcPry.BoundText
    DtcPryUni.BoundText = DtcPry.BoundText
End Sub

Private Sub DtcPryCargo_Click(Area As Integer)
    DtcPryDes.BoundText = DtcPryCargo.BoundText
    DtcPry.BoundText = DtcPryCargo.BoundText
    DtcPryUni.BoundText = DtcPryCargo.BoundText

End Sub

Private Sub DtcPryDes_Click(Area As Integer)
    DtcPry.BoundText = DtcPryDes.BoundText
    DtcPryCargo.BoundText = DtcPryDes.BoundText
    DtcPryUni.BoundText = DtcPryDes.BoundText
    Dtc_codigo_ant.BoundText = DtcPryUni.Text
    Dtc_descrip_ant.BoundText = DtcPryUni.Text
    DtcOrgDes.BoundText = DtcPryCargo.Text
    DtcOrg.BoundText = DtcPryCargo.Text
   
End Sub

Private Sub DtcPryUni_Click(Area As Integer)
    DtcPryDes.BoundText = DtcPryUni.BoundText
    DtcPry.BoundText = DtcPryUni.BoundText
    DtcPryCargo.BoundText = DtcPryUni.BoundText
   
End Sub

Private Sub DtcPuesto_Click(Area As Integer)
    DtcPuestoDes.BoundText = DtcPuesto.BoundText
    DtcPuestoCargo.BoundText = DtcPuesto.BoundText
    DtcPuestoUni.BoundText = DtcPuesto.BoundText
End Sub

Private Sub DtcPuestoCargo_Click(Area As Integer)
    DtcPuestoDes.BoundText = DtcPuestoCargo.BoundText
    DtcPuesto.BoundText = DtcPuestoCargo.BoundText
    DtcPuestoUni.BoundText = DtcPuestoCargo.BoundText
End Sub

Private Sub DtcPuestoDes_Click(Area As Integer)
    DtcPuesto.BoundText = DtcPuestoDes.BoundText
    DtcPuestoCargo.BoundText = DtcPuestoDes.BoundText
    DtcPuestoUni.BoundText = DtcPuestoDes.BoundText
    
    DtcCargoDes.BoundText = DtcPuestoCargo.Text
    DtcCargo.BoundText = DtcPuestoCargo.Text
    dtc_codigo.BoundText = DtcPuestoUni.Text
    Dtc_descrip.BoundText = DtcPuestoUni.Text
    dtc_codigo.BoundText = DtcPuestoUni.Text
    
    
End Sub

Private Sub DtcPuestoUni_Click(Area As Integer)
    DtcPuestoDes.BoundText = DtcPuestoUni.BoundText
    DtcPuesto.BoundText = DtcPuestoUni.BoundText
    DtcPuestoCargo.BoundText = DtcPuestoUni.BoundText
End Sub

Private Sub DtcRespaldo_Click(Area As Integer)
    DtcRespaldoCod.BoundText = DtcRespaldo.BoundText
End Sub

Private Sub DtcRespaldoCod_Click(Area As Integer)
    DtcRespaldo.BoundText = DtcRespaldoCod.BoundText
End Sub

Private Sub DTPFelaboracion_LostFocus()
    DTPFaprobacion.Value = DTPFelaboracion.Value
End Sub

Private Sub Form_Load()
Text6.Text = "/" & Format(Date, "yy")

'  Call abrirtabla
  
  Set rs_FteFin = New ADODB.Recordset
  rs_FteFin.Open "select * from gc_unidad_ejecutora WHERE estado_codigo = 'APR' ", db, adOpenKeyset, adLockOptimistic   'ORDER BY beneficiario_denominacion
  Set adounidad_ant.Recordset = rs_FteFin.DataSource
  Dtc_descrip_ant.BoundText = Dtc_codigo_ant.BoundText
  
'  Set rs_Org = New ADODB.Recordset
'  rs_Org.Open "select * from fc_convenios  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoOrg.Recordset = rs_Org.DataSource
'  DtcOrgDes.BoundText = DtcOrg.BoundText
  
  Set rs_Org = New ADODB.Recordset
  rs_Org.Open "select * from RC_CARGOS  ", db, adOpenKeyset, adLockOptimistic
  Set AdoOrg.Recordset = rs_Org.DataSource
  DtcOrgDes.BoundText = DtcOrg.BoundText
  
  Set rs_Pry = New ADODB.Recordset
  rs_Pry.Open "select * from rc_PUESTOs  ", db, adOpenKeyset, adLockOptimistic
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
  rs_Puesto_Org.Open "select * from rc_PUESTOS where puesto_vacante = 'SI' and estado_codigo = 'APR' ", db, adOpenKeyset, adLockOptimistic
  Set AdoPuestoOrg.Recordset = rs_Puesto_Org.DataSource
  DtcPuestoDes.BoundText = DtcPuesto.BoundText
  
  Set rs_Respaldo = New ADODB.Recordset
  rs_Respaldo.Open "select * from rc_tipo_memoranda where uso = 'B' ORDER BY descripcion  ", db, adOpenKeyset, adLockOptimistic
  Set AdoRespaldo.Recordset = rs_Respaldo.DataSource
  DtcRespaldo.BoundText = DtcRespaldoCod.BoundText
  
  Set rs_Respaldo = New ADODB.Recordset
  rs_Respaldo.Open "select * from rc_tipo_memoranda where uso = 'B' ORDER BY descripcion  ", db, adOpenKeyset, adLockOptimistic
  Set AdoRespaldo.Recordset = rs_Respaldo.DataSource
  DtcRespaldo.BoundText = DtcRespaldoCod.BoundText
  
  Set rs_intercambio = New ADODB.Recordset
  rs_intercambio.Open "select * from rv_intercambio_puesto WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' AND estado_codigo <> 'ANL' AND beneficiario_codigo <> '" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "' ORDER BY beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic
  Set ado_intercambio.Recordset = rs_intercambio.DataSource
  dtc_beneficiario_den.BoundText = dtc_beneficiario_cod.BoundText
  
  'select * from rv_personal_contratado WHERE tipoben_codigo < '20' and beneficiario_codigo <> '0' AND estado_codigo <> 'ANL'
  
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
If txtSW.Text <> "ADD" Then

If rw_ficha_rrhh.AdoMovilidad.Recordset!tipo_mov = "INTERCAMBIO" Then
FraIntercambio.Visible = True
btn_intercambio.Visible = False
btn_asignar.Visible = True
Else
FraIntercambio.Visible = False
btn_intercambio.Visible = True
btn_asignar.Visible = False

End If

End If

  
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
        DTPFcontrato.Enabled = False
    Else
        DTPcontrato.Enabled = True
    End If
End Sub

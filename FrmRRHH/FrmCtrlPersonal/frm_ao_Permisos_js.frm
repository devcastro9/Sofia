VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_Permisos_js 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Personal - File Funcionario - Permisos"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9810
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame2 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9675
      TabIndex        =   45
      Top             =   0
      Width           =   9735
      Begin VB.PictureBox Cmdimprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2520
         Picture         =   "frm_ao_Permisos_js.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   51
         ToolTipText     =   "Cancela sin Guardar"
         Top             =   120
         Width           =   1335
      End
      Begin VB.PictureBox cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1200
         Picture         =   "frm_ao_Permisos_js.frx":08CD
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   50
         ToolTipText     =   "Cancela sin Guardar"
         Top             =   120
         Width           =   1335
      End
      Begin VB.PictureBox cmdOk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "frm_ao_Permisos_js.frx":11B9
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   49
         ToolTipText     =   "Graba Registro"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cargar"
         Height          =   680
         Left            =   4140
         Picture         =   "frm_ao_Permisos_js.frx":198F
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Carga Contrato en PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   4920
         Picture         =   "frm_ao_Permisos_js.frx":1D17
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ver Contrato PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERMISOS"
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
         Left            =   7680
         TabIndex        =   48
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame FraProyecto 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00008000&
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9735
      Begin VB.Frame frmResultado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RESULTADO CALCULO"
         ForeColor       =   &H00C00000&
         Height          =   975
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   9375
         Begin VB.TextBox txt_nrodias 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   44
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox txt_nrohoras 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   4560
            MaxLength       =   80
            TabIndex        =   43
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox txt_nrominutos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   6720
            MaxLength       =   80
            TabIndex        =   42
            Top             =   480
            Width           =   1000
         End
         Begin VB.CommandButton btnCalcular 
            Caption         =   "Calcular"
            Height          =   285
            Left            =   240
            TabIndex        =   41
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Minutos  "
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
            Index           =   15
            Left            =   6000
            TabIndex        =   40
            Top             =   480
            Width           =   780
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Horas "
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
            Index           =   14
            Left            =   3960
            TabIndex        =   39
            Top             =   480
            Width           =   600
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias"
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
            Left            =   2040
            TabIndex        =   38
            Top             =   480
            Width           =   420
         End
      End
      Begin VB.TextBox TxtInicial 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7080
         MaxLength       =   80
         TabIndex        =   24
         Text            =   "lkhdkdh"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.ComboBox Txt02 
         Height          =   315
         Left            =   6120
         TabIndex        =   21
         Text            =   "LUNES"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         Height          =   315
         ItemData        =   "frm_ao_Permisos_js.frx":209F
         Left            =   240
         List            =   "frm_ao_Permisos_js.frx":20C4
         TabIndex        =   20
         Text            =   "2015"
         Top             =   480
         Width           =   900
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   5040
         MaxLength       =   80
         TabIndex        =   17
         Top             =   5880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   3840
         MaxLength       =   80
         TabIndex        =   16
         Top             =   5880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8520
         MaxLength       =   80
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cmb_mescontrol 
         Height          =   315
         ItemData        =   "frm_ao_Permisos_js.frx":210A
         Left            =   240
         List            =   "frm_ao_Permisos_js.frx":2132
         TabIndex        =   1
         Text            =   "ENERO"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2300
      End
      Begin MSComCtl2.DTPicker dt_fechasolicitusper 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   2300
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109314049
         CurrentDate     =   44562
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker dt_fechadesde 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109314049
         CurrentDate     =   42005
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker hr_horadesde 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109314050
         CurrentDate     =   0.333333333333333
         MinDate         =   4.16666666666667E-02
      End
      Begin MSComCtl2.DTPicker dt_fechahasta 
         Height          =   315
         Left            =   3600
         TabIndex        =   5
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109314049
         CurrentDate     =   44562
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker hr_horahasta 
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109314050
         CurrentDate     =   0.770833333333333
         MaxDate         =   0.999305555555556
         MinDate         =   4.16666666666667E-02
      End
      Begin MSComCtl2.DTPicker dt_fechareincorporacion 
         Height          =   315
         Left            =   6960
         TabIndex        =   7
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109314049
         CurrentDate     =   44562
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         DataField       =   "TipoPermiso"
         DataSource      =   "frmBeneficiario_control.AdoPermiso"
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
         ListField       =   "TipoPermiso"
         BoundColumn     =   "TipoPermiso"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmb_tipopermiso 
         Bindings        =   "frm_ao_Permisos_js.frx":219B
         DataField       =   "TipoPermiso"
         DataSource      =   "frmBeneficiario_control.AdoPermiso"
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "TipoPermiso"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker hr_horareincorporacion 
         Height          =   315
         Left            =   5760
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   109314050
         CurrentDate     =   0.957106481481481
         MaxDate         =   0.999988425925926
         MinDate         =   4.16666666666667E-02
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   9720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "REINCORPORACION"
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
         Height          =   255
         Left            =   6960
         TabIndex        =   36
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "HASTA"
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
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   1800
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         Index           =   1
         X1              =   6360
         X2              =   6360
         Y1              =   1680
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         Index           =   0
         X1              =   3120
         X2              =   3120
         Y1              =   1680
         Y2              =   3840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "DESDE"
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
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Archivo:"
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
         Index           =   13
         Left            =   5760
         TabIndex        =   33
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado    "
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
         Left            =   8640
         TabIndex        =   32
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Permiso:      "
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
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Reincorporaci�n:"
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
         Left            =   5760
         TabIndex        =   30
         Top             =   3000
         Width           =   2040
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   " Hasta Hora:"
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
         Index           =   6
         Left            =   3000
         TabIndex        =   29
         Top             =   3000
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Reincorporacion:"
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
         Left            =   6960
         TabIndex        =   28
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta Fecha:"
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
         Index           =   4
         Left            =   3600
         TabIndex        =   27
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitud Permiso: "
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
         Left            =   3000
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblARCH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5670
         TabIndex        =   23
         Top             =   480
         Width           =   3705
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Gesti�n        "
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
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Benef"
         Height          =   195
         Index           =   9
         Left            =   4440
         TabIndex        =   19
         Top             =   5880
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "SW"
         Height          =   195
         Index           =   10
         Left            =   3480
         TabIndex        =   18
         Top             =   5880
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes de Control:                               "
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Fecha: "
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
         TabIndex        =   12
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Hora:       "
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
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1470
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciales"
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
         Index           =   10
         Left            =   6480
         TabIndex        =   10
         Top             =   5400
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   6600
      Top             =   4680
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
      Caption         =   "Ado_Clasificador"
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
   Begin MSDataGridLib.DataGrid DtgPermiso 
      Height          =   1665
      Left            =   120
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   2937
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
      Caption         =   "DETALLE DE PERMISOS"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Fecha_control"
         Caption         =   "Fecha Control"
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
         DataField       =   "Dia_control"
         Caption         =   "Dia Control"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "horas_permiso"
         Caption         =   "Hrs.Permiso"
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
         DataField       =   "minutos_permiso"
         Caption         =   "Min.Permiso"
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
         DataField       =   "Vacacion"
         Caption         =   "Hrs.Vacaci�n"
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
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1124.787
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoPermisoDetalle 
      Height          =   330
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   8565
      _ExtentX        =   15108
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
      Caption         =   " <--- Detalle de Permisos --->"
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
Attribute VB_Name = "frm_ao_Permisos_js"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' NOTA 27/07/16: Mejora para fecha de reincorporacion implementar calendario de feriados y dias habiles

Public Para_Aceptado As String
Dim rs_Clasificador As New ADODB.Recordset
Dim rs_correlativo As New ADODB.Recordset
Dim rs_correl_vac As New ADODB.Recordset
Dim rs_Permiso_detalle As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim sqlAux As String
Dim nomb2 As String
Dim hora01, hora02, hora03, hora04 As String
Dim fecha1 As String
Dim DirLic, DirVac As String
Dim totHrs, totMin, totVac As Integer
Dim numminutosTT As Integer

Private Sub btnCalcular_Click()
  Call ObtenerTiempoCal
End Sub

Private Sub cmb_tipopermiso_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 0 Then
        KeyAscii = 0
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdCancel_Click()
    'cancela la edicion de datos
    Para_Aceptado = "N"
    Unload Me
    'Me.Hide
End Sub

Private Sub cmdOk_Click()
On Error GoTo EditErr
'abcd
    cmb_mescontrol.Text = UCase(MonthName(Month(dt_fechadesde.Value)))
    dt_fechasolicitusper.Value = dt_fechadesde.Value
    TxtGestion.Text = Year(dt_fechadesde.Value)
 'acepta las modificaciones realizadas
' Dim NoDias, NoHoras, NoMin As Integer
' Dim DifHr1, DifHr2 As Integer
' If ValidaMontos Then
'   Call ObtenerTiempoCal
'   Dim SQLS As String
'   SQLS = ""
'   If txtSW = "ADD" Then
'
'      rw_ficha_rrhh.AdoPermiso.Recordset("beneficiario_codigo").Value = txtBenef.Text
'      rw_ficha_rrhh.AdoPermiso.Recordset("ges_gestion").Value = TxtGestion.Text
'      rw_ficha_rrhh.AdoPermiso.Recordset("mes_control") = cmb_mescontrol.Text
'      Set rs_correlativo = New ADODB.Recordset
'      rs_correlativo.Open "select * from ro_Permisos WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "'  ", db, adOpenKeyset, adLockOptimistic
'      If rs_correlativo.RecordCount > 0 Then
'            rw_ficha_rrhh.AdoPermiso.Recordset("CORREL") = rs_correlativo.RecordCount + 1
'      Else
'            rw_ficha_rrhh.AdoPermiso.Recordset!CORREL = 1
'      End If
'      rw_ficha_rrhh.AdoPermiso.Recordset("TipoPermiso").Value = Dtc_Par.Text
'
'      rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo"
'      If Trim(Dtc_Par.Text) = "VC" Then
'        Set rs_correl_vac = New ADODB.Recordset
'        rs_correl_vac.Open "select * from ro_Permisos WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "' and TipoPermiso = 'VC' ", db, adOpenKeyset, adLockOptimistic
'        If rs_correl_vac.RecordCount > 0 Then
'              rw_ficha_rrhh.AdoPermiso.Recordset!Vacacion = rs_correl_vac.RecordCount
'        Else
'              rw_ficha_rrhh.AdoPermiso.Recordset!Vacacion = 1
'        End If
'        'rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO_NOMB = Trim(rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "_Vacaciones_" & rw_ficha_rrhh.AdoPermiso.Recordset!Vacacion & ".pdf"
'      Else
'        'rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO_NOMB = Trim(rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "_Licencias_" & rw_ficha_rrhh.AdoPermiso.Recordset!CORREL & ".pdf"
'      End If
'      txtEstado.Text = "REG"
'   End If
'      rw_ficha_rrhh.AdoPermiso.Recordset("Fecha_control").Value = dt_fechasolicitusper.Value
''      rw_ficha_rrhh.AdoPermiso.Recordset("dia_control").Value = txt02.Text
'      rw_ficha_rrhh.AdoPermiso.Recordset("FechaDesde").Value = dt_fechasolicitusper.Value
'      rw_ficha_rrhh.AdoPermiso.Recordset("FechaHasta").Value = dt_fechahasta.Value
'      rw_ficha_rrhh.AdoPermiso.Recordset("fecha_reincorporacion").Value = dt_fechareincorporacion.Value
'      rw_ficha_rrhh.AdoPermiso.Recordset("horadesde").Value = Format(hr_horadesde.Value, "HH:mm:ss")
'      rw_ficha_rrhh.AdoPermiso.Recordset("horahasta").Value = Format(hr_horahasta.Value, "HH:mm:ss")
'      rw_ficha_rrhh.AdoPermiso.Recordset("Hora_reincorporacion").Value = Format(hr_horareincorporacion.Value, "HH:mm:ss")
'      NoDias = DateDiff("d", rw_ficha_rrhh.AdoPermiso.Recordset("FechaDesde").Value, rw_ficha_rrhh.AdoPermiso.Recordset("FechaHasta").Value)
'      rw_ficha_rrhh.AdoPermiso.Recordset("dias_permiso").Value = IIf(IsNull(NoDias), 0, NoDias) + 1 'txt_nrodias.Text
'      'NoHoras = DateDiff("h", rw_ficha_rrhh.AdoPermiso.Recordset("FechaDesde").Value, rw_ficha_rrhh.AdoPermiso.Recordset("FechaHasta").Value)
'      GlHora1 = "08:00"
'      GlHora2 = "14:30"
'      DifHr1 = DateDiff("h", CDate(GlHora1), rw_ficha_rrhh.AdoPermiso.Recordset("horadesde").Value)
'      DifHr2 = 4 - DateDiff("h", CDate(GlHora2), rw_ficha_rrhh.AdoPermiso.Recordset("horahasta").Value)
'      If DifHr1 > 0 Then
'        If DifHr1 > 4 Then
'            DifHr1 = 4
'        Else
'            DifHr1 = DifHr1
'        End If
'      Else
'         DifHr1 = 0
'      End If
'      If DifHr2 > 0 Then
'         DifHr2 = DifHr2
'      Else
'         DifHr2 = 0
'      End If
'      NoHoras = (rw_ficha_rrhh.AdoPermiso.Recordset("dias_permiso").Value * 8) - (DifHr1 + DifHr2)
'
'      NoMin = NoHoras * 60
'      'NoMin = DateDiff("n", rw_ficha_rrhh.AdoPermiso.Recordset("FechaDesde").Value, rw_ficha_rrhh.AdoPermiso.Recordset("FechaHasta").Value)
'      rw_ficha_rrhh.AdoPermiso.Recordset("dias_permiso").Value = txt_nrodias.Text ' ======
'      rw_ficha_rrhh.AdoPermiso.Recordset("horas_permiso").Value = txt_nrohoras.Text ' ===NoHoras     'txt_nrohoras.Text
'      rw_ficha_rrhh.AdoPermiso.Recordset("minutos_permiso").Value = txt_nrominutos.Text '==== NoMin     'txt_nrominutos.Text
'      rw_ficha_rrhh.AdoPermiso.Recordset("total_minuto").Value = numminutosTT '===== Total de minutos
'
'      rw_ficha_rrhh.AdoPermiso.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "NO", txtEstado.Text)
'      rw_ficha_rrhh.AdoPermiso.Recordset("fecha_registro") = Date
'      rw_ficha_rrhh.AdoPermiso.Recordset("usr_usuario").Value = glusuario
'      rw_ficha_rrhh.AdoPermiso.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
'      rw_ficha_rrhh.AdoPermiso.Recordset.Update
'   'End If
'   Para_Aceptado = "S"
'   Call detalle_permiso
'   Call ABRE_DETALLE
'   MsgBox "Los datos se guardaron con �xito ...", , "Atenci�n"
''   sino = MsgBox("Se guardaron los datos con �xito, desea Salir de la Pantalla actual ? ", vbYesNo + vbQuestion, "Atenci�n")
''   If sino = vbYes Then
''      Unload Me
''   End If
'   'Call acumulaMont(adoVentas.Recordset("ges_gestion"), adoVentas.Recordset("correl_venta"), adoVentas.Recordset("nro_venta"))
'  Set rstacumdet = New ADODB.Recordset
'  If rstacumdet.State = 1 Then rstacumdet.Close
'
'  ' ++++++++++++   COMENTADO FALTA DEFINIR ESTRUCTURAS DE TABLA 27/07/16  ++++++++++++++++++++
'
''  sqlAux = "select sum(horas_permiso) as totHrs, sum (minutos_permiso) as totMin , sum (Vacacion) as totVac from ro_Permisos_detalle where beneficiario_codigo = '" & txtBenef.Text & "' and ges_gestion = '" & TxtGestion.Text & "' and Correl = " & rw_ficha_rrhh.AdoPermiso.Recordset!CORREL & "  "
''  rstacumdet.Open sqlAux, db, adOpenKeyset, adLockOptimistic
''  sqlAux = "update ro_Permisos set horas_permiso = " & rstacumdet!totHrs & " , minutos_permiso = " & rstacumdet!totMin & ", Vacacion = " & rstacumdet!totVac & "  Where beneficiario_codigo = '" & txtBenef.Text & "' and ges_gestion = '" & TxtGestion.Text & "' and Correl = '" & rw_ficha_rrhh.AdoPermiso.Recordset!CORREL & "'"
''  db.Execute sqlAux
''  'DB.Execute "update ro_Permisos set ro_Permisos.horas_permiso = " & rstacumdet!totHrs & " , ro_Permisos.minutos_permiso = " & rstacumdet!totMin & ", ro_Permisos.Vacacion = " & rstacumdet!Vacacion & "  Where beneficiario_codigo = '" & rstacumdet!beneficiario_codigo & "' and ges_gestion = '" & rstacumdet!ges_gestion & "' and Correl = '" & rstacumdet!CORREL & "'"
''  If rstacumdet.State = 1 Then rstacumdet.Close
'rw_ficha_rrhh.opciones
'   Me.Hide
'
' End If

'
'If txtSW = "ADD" Then
' rw_ficha_rrhh.AdoPermiso.Recordset.AddNew
'End If
If Dtc_Par.Text = "" Then
    sino = MsgBox("Escoja el TIPO de PERMISO", vbInformation, "Permiso")
    Exit Sub
End If

If txt_nrodias.Text > 30 Then
    sino = MsgBox("El numero de dias no puede ser mayor a 30", vbInformation, "Permiso")
    Exit Sub
End If

If txt_nrohoras.Text > 7 Then
    sino = MsgBox("El numero de horas no puede ser mayor a 7", vbInformation, "Permiso")
    Exit Sub
End If

If txt_nrominutos.Text > 59 Then
    sino = MsgBox("El numero de minutos no puede ser mayor a 59", vbInformation, "Permiso")
    Exit Sub
End If

If txtSW = "ADD" Then
    rw_ficha_rrhh.AdoPermiso.Recordset.AddNew
End If

    numminutosTT = 0
    rw_ficha_rrhh.AdoPermiso.Recordset("beneficiario_codigo").Value = txtBenef.Text
    rw_ficha_rrhh.AdoPermiso.Recordset("ges_gestion").Value = TxtGestion.Text
    rw_ficha_rrhh.AdoPermiso.Recordset("mes_control") = cmb_mescontrol.Text
    Set rs_correlativo = New ADODB.Recordset
    rs_correlativo.Open "select MAX(Correl) AS NUM from ro_Permisos WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_correlativo!num <> "NULL" Then
        rw_ficha_rrhh.AdoPermiso.Recordset("Correl") = rs_correlativo!num + 1
    Else
        rw_ficha_rrhh.AdoPermiso.Recordset("Correl") = "1"
    End If
    rw_ficha_rrhh.AdoPermiso.Recordset("TipoPermiso").Value = Dtc_Par.Text
    rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo"
    If Trim(Dtc_Par.Text) = "VC" Then
        Set rs_correl_vac = New ADODB.Recordset
        rs_correl_vac.Open "select * from ro_Permisos WHERE beneficiario_codigo = '" & Trim(txtBenef.Text) & "' and TipoPermiso = 'VC' ", db, adOpenKeyset, adLockOptimistic
        If rs_correl_vac.RecordCount > 0 Then
              rw_ficha_rrhh.AdoPermiso.Recordset!Vacacion = rs_correl_vac.RecordCount
        Else
              rw_ficha_rrhh.AdoPermiso.Recordset!Vacacion = 1
        End If
        'rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO_NOMB = Trim(rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "_Vacaciones_" & rw_ficha_rrhh.AdoPermiso.Recordset!Vacacion & ".pdf"
    Else
        'rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO_NOMB = Trim(rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "_Licencias_" & rw_ficha_rrhh.AdoPermiso.Recordset!CORREL & ".pdf"
    End If
    txtEstado.Text = "REG"
   
    rw_ficha_rrhh.AdoPermiso.Recordset("Fecha_control").Value = dt_fechasolicitusper.Value
'      rw_ficha_rrhh.AdoPermiso.Recordset("dia_control").Value = txt02.Text
    rw_ficha_rrhh.AdoPermiso.Recordset("FechaDesde").Value = dt_fechasolicitusper.Value
    rw_ficha_rrhh.AdoPermiso.Recordset("FechaHasta").Value = dt_fechahasta.Value
    rw_ficha_rrhh.AdoPermiso.Recordset("fecha_reincorporacion").Value = dt_fechareincorporacion.Value
    rw_ficha_rrhh.AdoPermiso.Recordset("horadesde").Value = "08:00:00"
    rw_ficha_rrhh.AdoPermiso.Recordset("horahasta").Value = "18:30:00"
    rw_ficha_rrhh.AdoPermiso.Recordset("Hora_reincorporacion").Value = "14:30:00"
    rw_ficha_rrhh.AdoPermiso.Recordset("dias_permiso").Value = txt_nrodias.Text
    rw_ficha_rrhh.AdoPermiso.Recordset("horas_permiso").Value = txt_nrohoras.Text
    rw_ficha_rrhh.AdoPermiso.Recordset("minutos_permiso").Value = txt_nrominutos.Text
    If rw_ficha_rrhh.AdoPermiso.Recordset("dias_permiso") > 0 Then
        numminutosTT = rw_ficha_rrhh.AdoPermiso.Recordset("dias_permiso") * 480
        numminutosTT = numminutosTT + rw_ficha_rrhh.AdoPermiso.Recordset("horas_permiso") * 60
        numminutosTT = numminutosTT + rw_ficha_rrhh.AdoPermiso.Recordset("minutos_permiso")
    Else
        numminutosTT = numminutosTT + rw_ficha_rrhh.AdoPermiso.Recordset("horas_permiso") * 60
        numminutosTT = numminutosTT + rw_ficha_rrhh.AdoPermiso.Recordset("minutos_permiso")
    End If
      rw_ficha_rrhh.AdoPermiso.Recordset("total_minuto").Value = numminutosTT '===== Total de minutos
      rw_ficha_rrhh.AdoPermiso.Recordset("estado_codigo").Value = IIf(txtEstado.Text = "", "NO", txtEstado.Text)
      rw_ficha_rrhh.AdoPermiso.Recordset("fecha_registro") = Date
      rw_ficha_rrhh.AdoPermiso.Recordset("usr_usuario").Value = glusuario
      rw_ficha_rrhh.AdoPermiso.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
      rw_ficha_rrhh.AdoPermiso.Recordset("codigo_empresa").Value = rw_ficha_rrhh.Ado_datos.Recordset!codigo_empresa
      rw_ficha_rrhh.AdoPermiso.Recordset.Update
      rw_ficha_rrhh.opciones
    Me.Hide
Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub detalle_permiso()
    Dim fecha2 As Date
    Dim horaIng, horaSal As Date
    Dim dia2 As String
    Dim NoHoras, NoMin As Integer
    Dim DifHr1, DifHr2 As Integer
    Dim rs_premisoCtrl As New ADODB.Recordset
    fecha2 = dt_fechasolicitusper.Value    'rw_ficha_rrhh.AdoPermiso.Recordset("FechaDesde")
    horaIng = hr_horadesde.Value   'rw_ficha_rrhh.AdoPermiso.Recordset("horadesde")
    horaSal = hr_horahasta.Value   'rw_ficha_rrhh.AdoPermiso.Recordset("horahasta")
    DifHr1 = DateDiff("h", CDate(GlHora1), horaIng)
    If horaSal > GlHora2 Then
        DifHr2 = 4 - DateDiff("h", CDate(GlHora2), horaSal)
    Else
        DifHr2 = 4
    End If
    If DifHr1 > 0 Then
      If DifHr1 > 4 Then
          DifHr1 = 4
      Else
          DifHr1 = DifHr1
      End If
    Else
       DifHr1 = 0
    End If
    If DifHr2 > 0 Then
       DifHr2 = DifHr2
    Else
       DifHr2 = 0
    End If
    
    While fecha2 <= dt_fechahasta.Value
      Set rs_calendario2 = New ADODB.Recordset
      rs_calendario2.Open "select * from gc_calendario where fecha = '" & fecha2 & "' and tipo = 'L' and ges_gestion = '" & TxtGestion.Text & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
      If rs_calendario2.RecordCount > 0 Then
        Set rs_Permiso_detalle = New ADODB.Recordset
        rs_Permiso_detalle.Open "select * from ro_Permisos_detalle where beneficiario_codigo = '" & txtBenef.Text & "' and Fecha_control = '" & fecha2 & "' and Correl = '" & rw_ficha_rrhh.AdoPermiso.Recordset!CORREL & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'Set AdoPermisoDetalle = rs_Permiso_detalle
        If rs_Permiso_detalle.RecordCount > 0 Then
'            If rs_premisoCtrl.RecordCount = 1 Then
'                AdoPermisoDetalle.Recordset.MoveFirst
'            Else
'                AdoPermisoDetalle.Recordset.MoveNext
'            'rs_Permiso_detalle!Fecha_control = fecha2
'            'rs_Permiso_detalle.MoveNext
'            End If
        Else
            AdoPermisoDetalle.Recordset.AddNew
            AdoPermisoDetalle.Recordset!Fecha_control = fecha2
            AdoPermisoDetalle.Recordset!beneficiario_codigo = txtBenef.Text
            AdoPermisoDetalle.Recordset!CORREL = rw_ficha_rrhh.AdoPermiso.Recordset("Correl")
            AdoPermisoDetalle.Recordset!ges_gestion = rw_ficha_rrhh.AdoPermiso.Recordset("ges_gestion")
        End If
        'If rs_premisoCtrl.State = 1 Then rs_premisoCtrl.Close
        dia2 = WeekdayName(Weekday(fecha2))
        AdoPermisoDetalle.Recordset!dia_control = dia2
        If horaIng > GlHora1 Then
            AdoPermisoDetalle.Recordset!horadesde = horaIng
            horaIng = GlHora1
        Else
            AdoPermisoDetalle.Recordset!horadesde = GlHora1
            'NoHoras = 8
        End If
        If horaSal >= CDate("14:20:00") And horaSal <= CDate("16:30:00") Then
            AdoPermisoDetalle.Recordset!HoraHasta = horaSal
            horaSal = CDate("18:30:00")
        Else
            AdoPermisoDetalle.Recordset!HoraHasta = horaSal
            horaSal = CDate("18:30:00")
        End If
        NoHoras = 8 - (DifHr1 + DifHr2)
        NoMin = NoHoras * 60
        If rw_ficha_rrhh.AdoPermiso.Recordset("TipoPermiso") = "VC" Then
            AdoPermisoDetalle.Recordset!Vacacion = NoMin
        Else
            AdoPermisoDetalle.Recordset!Vacacion = 0
        End If
        AdoPermisoDetalle.Recordset!horas_permiso = NoHoras
        DifHr1 = 0
        DifHr2 = 0
        AdoPermisoDetalle.Recordset!minutos_permiso = NoMin
        AdoPermisoDetalle.Recordset!usr_usuario = glusuario
        AdoPermisoDetalle.Recordset!fecha_registro = Date
        AdoPermisoDetalle.Recordset!hora_registro = "08:00"
        AdoPermisoDetalle.Recordset.Update
        
      End If
      fecha2 = fecha2 + 1
    Wend
End Sub

Function ValidaMontos()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
ValidaMontos = True
'If Val(Me.mskMonto) > Val(Me.mskMonto_pendiente) Then
'    ValidaMontos = False
'    MsgBox "El monto indicado sobrepasa el monto pendiente de pago", vbInformation
'    Me.mskMonto.SelStart = 0
'    Me.mskMonto.SelLength = Len(Me.mskMonto)
'    Me.mskMonto.SetFocus
'End If
'    If Txt01 = "" Then
'        ValidaMontos = False
'    End If
    If Txt02 = "" Then
        ValidaMontos = False
        MsgBox " 'Dia' es requerido."
    End If
    If dt_fechasolicitusper = "" Then
        ValidaMontos = False
        MsgBox " 'FechaSolicitud' es requerido."
    End If
    If dt_fechahasta = "" Then
        ValidaMontos = False
        MsgBox " 'FechaHasta' es requerido."
    End If
    If dt_fechadesde = "" Then
        ValidaMontos = False
        MsgBox " 'FechaDesde' es requerido."
    End If
    If cmb_tipopermiso = "" Then
        ValidaMontos = False
        MsgBox " 'TipoPermiso' es requerido."
    End If
    If dt_fechahasta < dt_fechadesde Then
         ValidaMontos = False
         MsgBox " 'FechaHasta' no puede ser menor a 'FechaDesde'"
    End If
    
    If dt_fechadesde = dt_fechahasta Then
       If hr_horahasta < hr_horadesde Then
         ValidaMontos = False
         MsgBox " 'HoraHasta' no puede ser menor a 'DesdeHora'"
       End If
    End If
    
End Function


Private Sub cmdRefresh_Click()
 If lblARCH.Caption = "Cargar_Archivo" Then
    MsgBox "No Existe el Archivo asociado al Registro, debe Cargarlo ...", vbExclamation, "Advertencia"
 Else
    'If GlServidor <> GlMaquina Then      ' "-" Then
   If GlServidor = "SRVPRO" Then
      If rw_ficha_rrhh.AdoPermiso.Recordset!TipoPermiso = "VC" Then
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      Else
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      End If
   Else
      If rw_ficha_rrhh.AdoPermiso.Recordset!TipoPermiso = "VC" Then
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
      End If
   End If
 End If
 
End Sub

Private Sub CmdVerDisco_Click()
  On Error GoTo Error_Sub
    
  If rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo" Then
     If AdoPermiso.Recordset!TipoPermiso = "VC" Then
        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        GlArch = "VAC"
        'e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
     Else
        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        GlArch = "LIC"
        'e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
     End If
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
        If AdoPermiso.Recordset!TipoPermiso = "VC" Then
            DirVac = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\"
        Else
            DirLic = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
        End If
      Else
        If AdoPermiso.Recordset!TipoPermiso = "VC" Then
            DirVac = NombreCarpeta
        Else
            DirLic = NombreCarpeta
        End If
      End If
      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
         Frmexporta.DirDestino2.Path = DirVac
      Else
         Frmexporta.DirDestino2.Path = DirLic
      End If
      Frmexporta.Show vbModal
  Else
'    MsgBox ""
     sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atenci�n")
     If sino = vbYes Then
        If AdoPermiso.Recordset!TipoPermiso = "VC" Then
            NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\"
            Frmexporta.DirDestino.Path = NombreCarpeta
            GlArch = "VAC"
        Else
            NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
            Frmexporta.DirDestino.Path = NombreCarpeta
            GlArch = "LIC"
        End If
        'If GlServidor <> GlMaquina Then      ' "-" Then
        If GlServidor = "SRVPRO" Then
            If AdoPermiso.Recordset!TipoPermiso = "VC" Then
                DirVac = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\VACACIONES\"
            Else
                DirLic = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo) & "\LICENCIAS\"
            End If
        Else
            If AdoPermiso.Recordset!TipoPermiso = "VC" Then
                DirVac = NombreCarpeta
            Else
                DirLic = NombreCarpeta
            End If
        End If
        If AdoPermiso.Recordset!TipoPermiso = "VC" Then
            Frmexporta.DirDestino2.Path = DirVac
        Else
            Frmexporta.DirDestino2.Path = DirLic
        End If
        Frmexporta.Show vbModal
     End If
  End If

  Exit Sub
Error_Sub:
  MsgBox Err.Description, vbCritical

End Sub



Private Sub Dtc_Par_Click(Area As Integer)
    cmb_tipopermiso.BoundText = Dtc_Par.BoundText
End Sub

Private Sub cmb_tipopermiso_Click(Area As Integer)
    Dtc_Par.BoundText = cmb_tipopermiso.BoundText
End Sub

Private Sub Form_Load()
TxtGestion.Text = Year(Date)
  numminutosTT = 0
  dt_fechadesde = Date
  dt_fechahasta = Date
  dt_fechareincorporacion = Date
  dt_fechasolicitusper = Date
  txt_nrodias.Text = 0
  txt_nrohoras.Text = 0
  txt_nrominutos.Text = 0
'If glProceso = "CONSULTORIA" Then
'    Me.Caption = "Consultor�a - Captura de datos personales"
'Else
'    Me.Caption = "Recursos Humanos - Captura de datos personales"
'End If
Para_Aceptado = "N"
'LOS DATOS PERSONALES SE CARGAN EN EL FORMULARIO QUE LO LLAMA
'AQUI SE JALA LOS MONTOS REGISTRADOS EN AO_ADJUDICA_C
Dim Xmbe As Double, Xmde As Double, Xmbn As Double, Xmdn As Double
Dim XAbe As Double, XAde As Double, XAbn As Double, XAdn As Double
'With ac_Adjudicacion_c.adoSec.Recordset
'    Me.labTipoMoneda = !tipo_moneda
'    DE.dbo_edCmprSumaMontosLimiteBen1 !ges_gestion, !codigo_unidad, !codigo_solicitud, !numero_consultoria, Xmbe, Xmde, Xmbn, Xmdn, XAbe, XAde, XAbn, XAdn
'    If !tipo_moneda = "$US" Then
'        Me.mskMonto = Round(!monto_dolares_ext + !monto_dolares_nal, 2)
'        Me.mskMonto_ext = !monto_dolares_ext
'        Me.mskMonto_nal = !monto_dolares_nal
'        Me.mskMonto_limite = Xmde + Xmdn
'        Me.mskMonto_pendiente = Round(Xmde + Xmdn - XAde - XAdn + Val(Me.mskMonto), 2)
'        Me.labPorcExt = CStr(Format(Xmde / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        Me.labPorcNal = CStr(Format(Xmdn / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        Me.mskMonto = Round(!monto_dolares_ext + !monto_dolares_nal, 2)
'    Else
'        Me.mskMonto = Round(!monto_bolivianos_ext + !monto_bolivianos_nal)
'        Me.mskMonto_ext = !monto_bolivianos_ext
'        Me.mskMonto_nal = !monto_bolivianos_nal
'        Me.mskMonto_limite = Xmbe + Xmbn
'        Me.mskMonto_pendiente = Xmbe + Xmbn - XAbe - XAbn + Val(Me.mskMonto)
'        If Val(Me.mskMonto_limite) = 0 Then
'            Me.labPorcExt = "0 %"
'            Me.labPorcNal = "0 %"
'        Else
'            Me.labPorcExt = CStr(Format(Xmbe / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'            Me.labPorcNal = CStr(Format(Xmbn / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        End If
'        Me.mskMonto = Round(!monto_bolivianos_ext + !monto_bolivianos_nal)
'    End If
'End With
    GlHora1 = "08:00"
    GlHora2 = "14:30"
    Set rs_Clasificador = New ADODB.Recordset
    rs_Clasificador.Open "SELECT * FROM rc_TipoPermiso WHERE estado_codigo = 'APR' ORDER BY descripcion ", db, adOpenStatic
    Set Ado_Clasificador.Recordset = rs_Clasificador
        Dtc_Par.BoundText = cmb_tipopermiso.BoundText
    Call ABRE_DETALLE

'mskMonto.SetFocus
End Sub

Private Sub ABRE_DETALLE()
    Set rs_Permiso_detalle = New ADODB.Recordset
    'rs_Permiso_detalle.Open "SELECT * FROM ro_Permisos_detalle where beneficiario_codigo = '" & rw_ficha_rrhh.AdoPermiso.Recordset!beneficiario_codigo & "' and ges_gestion = '" & rw_ficha_rrhh.AdoPermiso.Recordset!ges_gestion & "' and Correl = '" & rw_ficha_rrhh.AdoPermiso.Recordset!CORREL & "' ", DB, adOpenKeyset, adLockOptimistic, adCmdText
    rs_Permiso_detalle.Open "SELECT * FROM ro_Permisos_detalle where beneficiario_codigo = '" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "'   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set AdoPermisoDetalle.Recordset = rs_Permiso_detalle
End Sub
'Private Sub mskMonto_KeyPress(KeyAscii As Integer)
'If Val(Chr(KeyAscii)) <> 0 Or Chr(KeyAscii) = "." Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "0" Or KeyAscii = 8 Then
'    'asdfasdf
'Else
'    KeyAscii = 0
'End If
'End Sub


Private Sub ObtenerTiempoCal()
   Dim Formato As String
   Formato = "#,##0"
   Dim cinim, cfinm, cinit, cfint As String
   cinim = "08:00:00"
   cfinm = "12:00:00"
   cinit = "14:30:00"
   cfint = "18:30:00"
   Dim hrini, hrfin, fini, FFin As String
   hrini = hr_horadesde
   hrfin = hr_horahasta
   fini = dt_fechadesde
   FFin = dt_fechahasta
   Dim numdias, numhoras, auxnumehoras, numminutos, minutosdia  As Integer
   numdias = Format(DateDiff("y", fini, FFin), Formato)
   minutosdia = Int(Format(DateDiff("n", cinim, cfinm), Formato))
   minutosdia = minutosdia + Int(Format(DateDiff("n", cinit, cfint), Formato))
   numhoras = 0
   numminutos = 0
   auxnumehoras = 0
   If fini <> FFin Then ' Fechas distintas
       If hrini = hrfin Then ' Mismo inicio fin
         ' numdias = (Format(DateDiff("y", fini, FFin), Formato)) ' Solo dias
       End If
       
       If hrfin >= hrini Then ' Hora fin mayor
          If hrini >= TimeValue(cinim) And hrini < TimeValue(cfinm) Then ' 1er intervalo ini
               If hrfin >= TimeValue(cinim) And hrfin < TimeValue(cfinm) Then ' 1er intervalo fin
                   numminutos = Int(Format(DateDiff("n", hrini, hrfin), Formato))
               Else  ' 2do intervalo fin
                   numminutos = Int(Format(DateDiff("n", cinit, hrfin), Formato))
                   numminutos = numminutos + Int(Format(DateDiff("n", cinim, cfinm), Formato))
               End If
          Else
               numminutos = Int(Format(DateDiff("n", hrini, hrfin), Formato))
          End If
       Else ' Hora inicio mayor
           numdias = numdias - 1
          
           If hrini >= TimeValue(cinim) And hrini < TimeValue(cfinm) Then ' 1er intervalo ini
               numminutos = numminutos + Int(Format(DateDiff("n", hrini, cfinm), Formato))
               numminutos = numminutos + Int((Format(DateDiff("n", cinit, cfint), Formato)))
           Else
               numminutos = numminutos + Int(Format(DateDiff("n", hrini, cfint), Formato))
           End If
           
           If hrfin >= TimeValue(cinim) And hrfin < TimeValue(cfinm) Then ' 1er intervalo fin
               numminutos = numminutos + Int(Format(DateDiff("n", cinim, hrfin), Formato))
           Else
               numminutos = numminutos + Int(Format(DateDiff("n", cinit, hrfin), Formato))
               numminutos = numminutos + Int(Format(DateDiff("n", cinim, cfinm), Formato))
           End If
       End If
   Else ' Fecha inicio y fin son iguales
           If hrini >= TimeValue(cinim) And hrini < TimeValue(cfinm) Then ' 1er intervalo ini
               If hrfin >= TimeValue(cinim) And hrfin < TimeValue(cfinm) Then ' 1er intervalo fin
                    numminutos = numminutos + Int(Format(DateDiff("n", hrini, hrfin), Formato))
               Else
                    numminutos = numminutos + Int(Format(DateDiff("n", hrini, cfinm), Formato))
                    numminutos = numminutos + Int(Format(DateDiff("n", cinit, hrfin), Formato))
               End If
           Else ' 2do intervalo ini
               numminutos = numminutos + Int(Format(DateDiff("n", hrini, hrfin), Formato))
           End If
   End If
      
   ' Obtiene horas por minutos
   If numminutos >= 60 Then
         numhoras = Int(numminutos / 60)
        If numminutos Mod 60 = 0 Then ' Son horas exactas
           numminutos = 0
        Else
           numminutos = numminutos - (numhoras * 60)
        End If
   End If
   ' Total de minutos
   numminutosTT = (minutosdia * Int(numdias)) + (numhoras * 60) + numminutos
   
   'Debug.Print "dias: " + CStr(numdias) + " hrs: " + CStr(numhoras) + " mns: " + CStr(numminutos) + " TTmns: " + CStr(numminutosTT)
   ' Carga de valores de controles.
   txt_nrodias.Text = numdias
   txt_nrohoras.Text = numhoras
   txt_nrominutos.Text = numminutos
   ' Si hrfin es menor a cfint
   If hrfin < cfint Then
       dt_fechareincorporacion = dt_fechahasta
       hr_horareincorporacion = hr_horahasta
   Else
       dt_fechareincorporacion = DateAdd("y", 1, dt_fechahasta)
       hr_horareincorporacion = cinim
   End If
End Sub
 

Private Sub txt_nrodias_KeyPress(KeyAscii As Integer)
If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub txt_nrohoras_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub txt_nrominutos_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

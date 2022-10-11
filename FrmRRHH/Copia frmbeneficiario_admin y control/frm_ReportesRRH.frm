VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ReportesRRH 
   BackColor       =   &H00000000&
   Caption         =   "Reportes RRHH"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   9975
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txt_aguinaldo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   104
      Text            =   "0"
      Top             =   5640
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Parámetros para Planilla por Financiamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   915
      Left            =   120
      TabIndex        =   90
      Top             =   8320
      Visible         =   0   'False
      Width           =   9300
      Begin MSDataListLib.DataCombo dtc_orgDes 
         Bindings        =   "frm_ReportesRRH.frx":0000
         DataField       =   "org_codigo"
         Height          =   315
         Left            =   1800
         TabIndex        =   91
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "org_descripcion"
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
      Begin MSDataListLib.DataCombo dtc_org 
         Bindings        =   "frm_ReportesRRH.frx":0016
         DataField       =   "org_codigo"
         Height          =   315
         Left            =   6600
         TabIndex        =   92
         Top             =   -120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   0
         ForeColor       =   16777215
         ListField       =   "org_codigo"
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
      Begin MSAdodcLib.Adodc Ado_org 
         Height          =   330
         Left            =   4200
         Top             =   240
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
         Caption         =   "Ado_org"
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
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Financiador"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   840
         TabIndex        =   93
         Top             =   495
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Parámetros para Migrar al Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1755
      Left            =   120
      TabIndex        =   78
      Top             =   6555
      Visible         =   0   'False
      Width           =   9300
      Begin VB.TextBox txt_archivo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   88
         Text            =   "LPCGI0716"
         Top             =   360
         Width           =   1590
      End
      Begin VB.TextBox txt_convenio 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   79
         Text            =   "0"
         Top             =   360
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker DTp_FCarga 
         Height          =   300
         Left            =   1845
         TabIndex        =   80
         Top             =   1320
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   95879169
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker DTP_Fvigencia 
         Height          =   300
         Left            =   7485
         TabIndex        =   81
         Top             =   1320
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   95879169
         CurrentDate     =   42735
      End
      Begin MSDataListLib.DataCombo dtc_ctades 
         Bindings        =   "frm_ReportesRRH.frx":002C
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   4200
         TabIndex        =   82
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   0
         ForeColor       =   16777215
         ListField       =   "cta_descripcion"
         BoundColumn     =   "cta_codigo"
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
      Begin MSDataListLib.DataCombo dtc_cta 
         Bindings        =   "frm_ReportesRRH.frx":0045
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   1800
         TabIndex        =   83
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
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
      Begin MSAdodcLib.Adodc Ado_cuenta 
         Height          =   330
         Left            =   3360
         Top             =   360
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
         Caption         =   "Ado_cuenta"
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Límite Vigencia Planilla:"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   5280
         TabIndex        =   89
         Top             =   1365
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Carga Datos:"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   1365
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Número de convenio"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Nombre del Archivo a Enviar"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   5400
         TabIndex        =   85
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Cuenta Bancaria CGI"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   855
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1515
      Left            =   120
      TabIndex        =   68
      Top             =   5040
      Visible         =   0   'False
      Width           =   9300
      Begin VB.ComboBox cb_aguinaldo 
         Height          =   315
         ItemData        =   "frm_ReportesRRH.frx":005E
         Left            =   5280
         List            =   "frm_ReportesRRH.frx":0068
         TabIndex        =   105
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "TODAS INTERIOR"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6480
         TabIndex        =   97
         Top             =   1080
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "TODAS LAS PLANILLAS"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6480
         TabIndex        =   94
         Top             =   840
         Width           =   2115
      End
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox cmb_gestion_rep 
         Height          =   315
         ItemData        =   "frm_ReportesRRH.frx":0086
         Left            =   1920
         List            =   "frm_ReportesRRH.frx":00AB
         TabIndex        =   72
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbo_mes_rep 
         Height          =   315
         ItemData        =   "frm_ReportesRRH.frx":00F1
         Left            =   5280
         List            =   "frm_ReportesRRH.frx":011F
         TabIndex        =   71
         Top             =   360
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   7605
         TabIndex        =   69
         Top             =   360
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   95879169
         CurrentDate     =   42005
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   7605
         TabIndex        =   70
         Top             =   840
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   95879169
         CurrentDate     =   42369
      End
      Begin MSDataListLib.DataCombo dtc_rep_det 
         Bindings        =   "frm_ReportesRRH.frx":01A2
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   2880
         TabIndex        =   74
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_descripcion"
         BoundColumn     =   "planilla_codigo"
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
      Begin MSDataListLib.DataCombo dtc_rep_cod 
         Bindings        =   "frm_ReportesRRH.frx":01BE
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   1920
         TabIndex        =   75
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_codigo"
         BoundColumn     =   "planilla_codigo"
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
      Begin MSAdodcLib.Adodc Ado_datos_rep 
         Height          =   330
         Left            =   120
         Top             =   600
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "Ado_datos_rep"
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
      Begin MSDataListLib.DataCombo dtc_depto 
         Bindings        =   "frm_ReportesRRH.frx":01DA
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   1920
         TabIndex        =   101
         Top             =   600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "depto_codigo"
         BoundColumn     =   "planilla_codigo"
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
      Begin VB.Label Label34 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   66
         Top             =   980
         Width           =   855
      End
      Begin VB.Label Label33 
         BackColor       =   &H00000000&
         Caption         =   "MES"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   4800
         TabIndex        =   77
         Top             =   380
         Width           =   735
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Caption         =   "GESTIÓN"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   76
         Top             =   380
         Width           =   735
      End
   End
   Begin VB.Frame fmrTipoReporte 
      BackColor       =   &H00000000&
      Caption         =   "------------------------Generales--------------------------------------------------------------Planillas"
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
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9300
      Begin VB.OptionButton optrep006D 
         BackColor       =   &H00000000&
         Caption         =   "VISION"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   3600
         TabIndex        =   106
         Top             =   2280
         Width           =   1125
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA AGUINALDO"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   240
         TabIndex        =   103
         Top             =   3740
         Width           =   2700
      End
      Begin VB.OptionButton optRep020 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA PERSONAL A PRUEBA"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   102
         Top             =   3345
         Width           =   2820
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "SANCIONES "
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   240
         TabIndex        =   100
         Top             =   3480
         Width           =   2700
      End
      Begin VB.OptionButton optRep019 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA PERSONAL A PRUEBA DETALLE"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   99
         Top             =   3690
         Width           =   3900
      End
      Begin VB.OptionButton optRep010_B 
         BackColor       =   &H00000000&
         Caption         =   "MIGRAR"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   8160
         TabIndex        =   96
         Top             =   240
         Width           =   1020
      End
      Begin VB.OptionButton optrep006B 
         BackColor       =   &H00000000&
         Caption         =   "VALORES"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   3600
         TabIndex        =   95
         Top             =   1812
         Width           =   1125
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00000000&
         Caption         =   "Libre 3"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   8040
         TabIndex        =   64
         Top             =   3360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.OptionButton optRep015 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA IMPOSITIVA - RC-IVA"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   63
         Top             =   1965
         Width           =   3525
      End
      Begin VB.OptionButton optRep016 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA POR FINANCIAMIENTO"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   62
         Top             =   2310
         Width           =   3525
      End
      Begin VB.OptionButton optRep012 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA AFP PREVISION (PARA MIGRAR)"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   61
         Top             =   930
         Width           =   4005
      End
      Begin VB.OptionButton optRep011 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA TRIMESTRAL MIN. (PARA MIGRAR)"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   60
         Top             =   585
         Width           =   3885
      End
      Begin VB.OptionButton optRep014 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA GENERAL AFP'S"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   59
         Top             =   1620
         Width           =   3540
      End
      Begin VB.OptionButton optRep013 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA AFP FUTURO (PARA MIGRAR)"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   58
         Top             =   1275
         Width           =   3540
      End
      Begin VB.OptionButton optRep017 
         BackColor       =   &H00000000&
         Caption         =   "CTAS. PARA MIGRAR AL BANCO"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   57
         Top             =   2655
         Width           =   2820
      End
      Begin VB.OptionButton optRep010 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA MINISTERIO DE TRABAJO"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   56
         Top             =   240
         Width           =   3180
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00000000&
         Caption         =   "-"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   8040
         TabIndex        =   55
         Top             =   3015
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OptionButton optRep008 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR TIPO DE VENTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2820
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.OptionButton optrep006 
         BackColor       =   &H00000000&
         Caption         =   "BOLETAS DE PAGO POR PLANILLA"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2070
         Width           =   3045
      End
      Begin VB.OptionButton optRep004 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA DETALLE ASISTENCIA"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1290
         Width           =   3285
      End
      Begin VB.OptionButton optRep003 
         BackColor       =   &H00000000&
         Caption         =   "LISTADO DEL PERSONAL"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   945
         Width           =   3525
      End
      Begin VB.OptionButton optrep005 
         BackColor       =   &H00000000&
         Caption         =   "RESUMEN ASISTENCIA "
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   2580
      End
      Begin VB.OptionButton optrep006C 
         BackColor       =   &H00000000&
         Caption         =   "POLITICA"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   2023
         Width           =   1155
      End
      Begin VB.OptionButton optRep009 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR COBRADOR"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3165
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.OptionButton optRep007 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR BIEN O SERVICIO"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2550
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.OptionButton optRep001 
         BackColor       =   &H00000000&
         Caption         =   "ESTRUCTURA ORGANIZACIONAL"
         ForeColor       =   &H00FFFFC0&
         Height          =   390
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   3780
      End
      Begin VB.OptionButton optRep002 
         BackColor       =   &H00000000&
         Caption         =   "ESTRUCTURA DE PUESTOS"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   615
         Width           =   3285
      End
      Begin VB.OptionButton optRep018 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA REFRIGERIO"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5040
         TabIndex        =   98
         Top             =   3000
         Width           =   3780
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   2040
         Y2              =   2280
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   120
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   4800
         X2              =   3480
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   4800
         X2              =   4800
         Y1              =   1800
         Y2              =   2520
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   4800
         X2              =   3480
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   3480
         Y1              =   2280
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   3480
         Y1              =   1800
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   4920
         X2              =   4920
         Y1              =   120
         Y2              =   3960
      End
   End
   Begin VB.Frame FrameConDet 
      Caption         =   "Con Detalle"
      ForeColor       =   &H000000C0&
      Height          =   600
      Left            =   3480
      TabIndex        =   52
      Top             =   7080
      Visible         =   0   'False
      Width           =   2040
      Begin VB.OptionButton optNo 
         Caption         =   "No"
         Height          =   195
         Left            =   945
         TabIndex        =   54
         Top             =   250
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optSi 
         Caption         =   "Si"
         Height          =   225
         Left            =   105
         TabIndex        =   53
         Top             =   255
         Width           =   705
      End
   End
   Begin VB.Frame FrameTipo 
      BackColor       =   &H00404040&
      Caption         =   "Estado de la Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   2400
      TabIndex        =   47
      Top             =   5400
      Visible         =   0   'False
      Width           =   4335
      Begin VB.OptionButton Opt_3 
         BackColor       =   &H00404040&
         Caption         =   "Ventas Facturadas y Cobradas"
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
         Height          =   255
         Left            =   300
         TabIndex        =   51
         Top             =   880
         Width           =   3855
      End
      Begin VB.OptionButton opt_4 
         BackColor       =   &H00404040&
         Caption         =   "Todas (Las 3 anteriores)"
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
         Height          =   285
         Left            =   300
         TabIndex        =   50
         Top             =   1180
         Value           =   -1  'True
         Width           =   2805
      End
      Begin VB.OptionButton Opt_1 
         BackColor       =   &H00404040&
         Caption         =   "Ventas Realizadas No Facturadas"
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
         Left            =   300
         TabIndex        =   49
         Top             =   255
         Width           =   3630
      End
      Begin VB.OptionButton opt_2 
         BackColor       =   &H00404040&
         Caption         =   "Ventas Facturadas No Cobradas"
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
         Height          =   285
         Left            =   300
         TabIndex        =   48
         Top             =   550
         Width           =   3750
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000006&
      Height          =   780
      Left            =   120
      ScaleHeight     =   720
      ScaleWidth      =   9240
      TabIndex        =   43
      Top             =   120
      Width           =   9300
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1800
         Picture         =   "frm_ReportesRRH.frx":01F6
         ScaleHeight     =   615
         ScaleWidth      =   1365
         TabIndex        =   67
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   60
         Width           =   1365
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frm_ReportesRRH.frx":09B8
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   65
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   60
         Width           =   1455
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00808000&
         Caption         =   "&Docs."
         Enabled         =   0   'False
         Height          =   720
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   740
      End
      Begin VB.CommandButton CmdFoto 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Foto"
         Height          =   720
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Carga Foto de la Persona"
         Top             =   120
         Visible         =   0   'False
         Width           =   740
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECURSOS HUMANOS"
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
         Left            =   4665
         TabIndex        =   46
         Top             =   180
         Width           =   3435
      End
   End
   Begin VB.Frame ConProy00 
      BackColor       =   &H00404040&
      Caption         =   "Selecciona Parametro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3600
      Left            =   120
      TabIndex        =   9
      Top             =   5445
      Visible         =   0   'False
      Width           =   9300
      Begin MSDataListLib.DataCombo DtcProvCod 
         Height          =   315
         Left            =   6720
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCliCod 
         Height          =   315
         Left            =   6720
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcVenCod 
         Height          =   315
         Left            =   6720
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCbrCod 
         Height          =   315
         Left            =   6720
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCbrDes 
         Height          =   315
         Left            =   2520
         TabIndex        =   26
         Top             =   1440
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcProvDes 
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCliDes 
         Height          =   315
         Left            =   2520
         TabIndex        =   17
         Top             =   720
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcVenDes 
         Height          =   315
         Left            =   2520
         TabIndex        =   20
         Top             =   1080
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc Ado_proveedor 
         Height          =   330
         Left            =   7920
         Top             =   360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
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
         Caption         =   "Ado_proveedor"
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
      Begin MSAdodcLib.Adodc ado_Cliente 
         Height          =   330
         Left            =   7920
         Top             =   720
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "ado_Cliente"
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
      Begin MSAdodcLib.Adodc Ado_vendedor 
         Height          =   330
         Left            =   7920
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Ado_vendedor"
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
      Begin MSAdodcLib.Adodc Ado_Cobrador 
         Height          =   330
         Left            =   7920
         Top             =   1440
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Ado_Cobrador"
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
      Begin MSDataListLib.DataCombo DtcTipo 
         Height          =   315
         Left            =   6720
         TabIndex        =   28
         Top             =   1800
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_venta"
         BoundColumn     =   "tipo_venta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcTipoDes 
         Height          =   315
         Left            =   2520
         TabIndex        =   29
         Top             =   1800
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_descripcion"
         BoundColumn     =   "tipo_venta"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc Ado_Tipo 
         Height          =   330
         Left            =   7920
         Top             =   1800
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Ado_Cobrador"
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
      Begin MSDataListLib.DataCombo DtcTipoCli 
         Height          =   315
         Left            =   6720
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Tipo_Beneficiario"
         BoundColumn     =   "Tipo_Beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcDepto 
         Height          =   315
         Left            =   6720
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Depto"
         BoundColumn     =   "Depto"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcTipoCliDes 
         Height          =   315
         Left            =   2520
         TabIndex        =   34
         Top             =   2160
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Tipo_Beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCiu 
         Height          =   315
         Left            =   2520
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "municipio"
         BoundColumn     =   "Depto"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc Ado_TipoBenef 
         Height          =   330
         Left            =   7920
         Top             =   2160
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Ado_Cobrador"
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
      Begin MSAdodcLib.Adodc Ado_Ciudad 
         Height          =   330
         Left            =   7920
         Top             =   2520
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Ado_Ciudad"
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
      Begin MSDataListLib.DataCombo DtcMesC 
         Height          =   315
         Left            =   6240
         TabIndex        =   38
         Top             =   2880
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "nom_periodo"
         BoundColumn     =   "nom_periodo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcProdC 
         Height          =   315
         Left            =   6720
         TabIndex        =   39
         Top             =   3240
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ccodDetalle"
         BoundColumn     =   "codDetalle"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcProd 
         Height          =   315
         Left            =   2520
         TabIndex        =   40
         Top             =   3240
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "concepto_venta"
         BoundColumn     =   "codDetalle"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcMes 
         Height          =   315
         Left            =   2520
         TabIndex        =   41
         Top             =   2880
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "nom_periodo"
         BoundColumn     =   "nom_periodo"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc Ado_Meses 
         Height          =   330
         Left            =   7920
         Top             =   2880
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Ado_Meses"
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
      Begin MSAdodcLib.Adodc Ado_Producto 
         Height          =   330
         Left            =   7920
         Top             =   3240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Ado_Producto"
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
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes Final . . . :"
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
         Height          =   255
         Left            =   5040
         TabIndex        =   42
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Producto (Bien). . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   37
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes Inicial . . .  . . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   36
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad (Cliente). . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   2565
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cliente . . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   2205
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Venta . . . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   1845
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrador . . . . . . . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   1485
         Width           =   1695
      End
      Begin VB.Label lblConv 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor . . . . . . . . .:"
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
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   1125
         Width           =   1695
      End
      Begin VB.Label lblOrg 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente . . . . . . . . . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   765
         Width           =   1620
      End
      Begin VB.Label lblFuente 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor . . . . . . . . :"
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
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   405
         Width           =   1575
      End
   End
   Begin VB.TextBox txtSubProg 
      Height          =   315
      Left            =   7680
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton butEstProg 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<-- Elige Estruc. Prog."
      Height          =   315
      Left            =   6960
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5535
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   8355
      TabIndex        =   3
      Top             =   5625
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "fte_codigo"
      BoundColumn     =   "fte_descripcion_larga"
      Text            =   "DataCombo1"
   End
   Begin Crystal.CrystalReport CryReporte 
      Left            =   9480
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryReporte4 
      Left            =   9480
      Top             =   1710
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryDetalle 
      Left            =   9495
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryReporte3 
      Left            =   9480
      Top             =   1155
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
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryReporte2 
      Left            =   9495
      Top             =   645
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryReporte5 
      Left            =   9480
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryReporte6 
      Left            =   9480
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "frm_ReportesRRH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iResult As Integer
Public vProg As String
Public vSubProg As String
Public vProy As String
Public vActi As String
Public glRepPresup As String
Public conDetalle As Boolean

Dim rs_proveedor, rs_cliente, rs_vendedor, rs_cobrador As New ADODB.Recordset
Dim rs_tipo, rs_tipoBenef, rs_ciudad As New ADODB.Recordset
Dim rs_meses, rs_producto As New ADODB.Recordset
Dim rs_aux7, rs_aux8, rs_aux9  As New ADODB.Recordset
Dim sino As String

Public Sub inicio(Usuario, Proceso As String)
  glRepPresup = Proceso
  Call llena_datos
  DTPfecha1.Value = Format("01/01/2016", "dd/mm/yyyy")
  dtpFecha2.Value = Format(Date, "dd/mm/yyyy")
  'dtpFecha2.Value = Date
'  frmRepPresupuesto.Show
End Sub

Private Sub BtnImprimir_Click()
If Option4.Value = True Then
    If cmb_gestion_rep.Text = "" Or dtc_rep_cod.Text = "" Or cb_aguinaldo.Text = "" Then
        sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
        Exit Sub
    End If
End If

If optRep001.Value = False And optRep002.Value = False And optRep003.Value = False And Option4.Value = False Then
    If cmb_gestion_rep.Text = "" Or dtc_rep_cod.Text = "" Or cbo_mes_rep.Text = "" Then
        sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
        Exit Sub
    End If
End If
If optRep016.Value = True Then
    If cmb_gestion_rep.Text = "" Or dtc_rep_cod.Text = "" Or cbo_mes_rep.Text = "" Or dtc_orgDes.Text = "" Then
        sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
        Exit Sub
    End If
End If
If optRep017.Value = True Then
    If cmb_gestion_rep.Text = "" Or dtc_rep_cod.Text = "" Or cbo_mes_rep.Text = "" Or txt_convenio.Text = "" Or dtc_cta.Text = "" Or txt_archivo.Text = "" Then
        sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
        Exit Sub
    End If
End If

If optRep018.Value = True Then
     If cmb_gestion_rep.Text = "" Or dtc_rep_cod.Text = "" Or cbo_mes_rep.Text = "" Or txt_convenio.Text = "" Or dtc_cta.Text = "" Or txt_archivo.Text = "" Then
        sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
        Exit Sub
    End If
End If
    
If optRep019.Value = True Then
    If cmb_gestion_rep.Text = "" Or dtc_rep_cod.Text = "" Or cbo_mes_rep.Text = "" Or dtc_orgDes.Text = "" Then
        sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
        Exit Sub
    End If

End If

If optRep020.Value = True Then
     If cmb_gestion_rep.Text = "" Or dtc_rep_cod.Text = "" Or cbo_mes_rep.Text = "" Or txt_convenio.Text = "" Or dtc_cta.Text = "" Or txt_archivo.Text = "" Then
        sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
        Exit Sub
    End If
End If




CryReporte.Reset
CryReporte2.Reset
CryReporte3.Reset
CryReporte4.Reset
CryReporte5.Reset
CryReporte6.Reset
CryDetalle.Reset
CryReporte.WindowState = crptMaximized
CryReporte2.WindowState = crptMaximized
CryReporte3.WindowState = crptMaximized
CryReporte4.WindowState = crptMaximized
CryReporte5.WindowState = crptMaximized
CryReporte6.WindowState = crptMaximized
CryDetalle.WindowState = crptMaximized

CryReporte.WindowShowSearchBtn = True
CryReporte2.WindowShowSearchBtn = True
CryReporte3.WindowShowSearchBtn = True
CryReporte4.WindowShowSearchBtn = True
CryReporte5.WindowShowSearchBtn = True
CryReporte6.WindowShowSearchBtn = True
CryDetalle.WindowShowSearchBtn = True

CryReporte.WindowShowRefreshBtn = True
CryReporte2.WindowShowRefreshBtn = True
CryReporte3.WindowShowRefreshBtn = True
CryReporte4.WindowShowRefreshBtn = True
CryReporte5.WindowShowRefreshBtn = True
CryReporte6.WindowShowRefreshBtn = True
CryDetalle.WindowShowRefreshBtn = True

CryReporte.WindowShowPrintSetupBtn = True
CryReporte2.WindowShowPrintSetupBtn = True
CryReporte3.WindowShowPrintSetupBtn = True
CryReporte4.WindowShowPrintSetupBtn = True
CryReporte5.WindowShowPrintSetupBtn = True
CryReporte6.WindowShowPrintSetupBtn = True
CryDetalle.WindowShowPrintSetupBtn = True



 '------------ REPORTES GENERALES --------------
    'ESTRURA ORGANIZACIONAL
  If optRep001.Value = True Then
    Call Reportes("\REPORTES\clasificadores\gr_unidad_ejecutora.rpt")
    
    'ESTRUCTURA DE PUESTOS
  ElseIf optRep002.Value = True Then
    Call Reportes("\REPORTES\RRHH\rr_puestos_organizacionales.rpt")
    
    'LISTADO DEL PERSONAL
  ElseIf optRep003.Value = True Then
    Call Reportes("\REPORTES\clasificadores\gr_beneficiario_Persona.rpt")
    
    'PLANILLA ASISTENCIA DETALLE
  ElseIf optRep004.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_asistencia.rpt")
    
    'PLANILLA ASISTENCIA RESUMEN
  ElseIf optrep005.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_asistencia_totales.rpt")
    
    'PLANILLA BOLETAS DE PAGO
  ElseIf optrep006.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_boleta_pago_general.rpt")
    
    'REVERSO DE BOLETAS
    ElseIf optrep006B.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_boleta_pago_general_rev.rpt")
    
    ElseIf optrep006C.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_boleta_pago_general_rev_2.rpt")
    
    ElseIf optrep006D.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_boleta_pago_general_rev1.rpt")
    
    
    
 '------------ REPORTES PLANILLAS --------------
   'PLANILLA MINISTERIO DE TRABAJO
  ElseIf optRep010.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_ministerio.rpt")
    
     ElseIf Option4.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_auginaldo.rpt")
    
     ElseIf Option3.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_sanciones.rpt")
    
    
   ElseIf optRep010_B.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_ministerio_migrar_2.rpt")
    
    
   'PLANILLA TRIMESTRAL MIN. (PARA MIGRAR)
  ElseIf optRep011.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_ministerio_migrar.rpt")
    
   'PLANILLA AFP PREVISION (PARA MIGRAR)
  ElseIf optRep012.Value = True Then
    Call Reportes4("\REPORTES\RRHH\rr_planilla_prevision.rpt")
    
   'PLANILLA AFP FUTURO (PARA MIGRAR)
  ElseIf optRep013.Value = True Then
    Call Reportes4("\REPORTES\RRHH\rr_planilla_futuro.rpt")
    
   'PLANILLA GENERAL AFP'S
  ElseIf optRep014.Value = True Then
    Call Reportes4("\REPORTES\RRHH\rr_liatado_afps.rpt")
    
   'PLANILLA IMPOSITIVA - RC-IVA
  ElseIf optRep015.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_rciva.rpt")
    
   'PLANILLA POR FINANCIAMIENTO
  ElseIf optRep016.Value = True Then
    Call Reportes6("\REPORTES\RRHH\rr_planilla_por_financiamiento.rpt")
    
   'CTAS. PARA MIGRAR AL BANCO
  ElseIf optRep017.Value = True Then
    Call Reportes5("\REPORTES\RRHH\rr_envio_banco_mercantil.rpt")

  'PLANILLAR EFRIGERIO
  ElseIf optRep018.Value = True Then
    Call Reportes5("\REPORTES\RRHH\rr_refrigerio.rpt")

  ElseIf optRep019.Value = True Then
    Call Reportes2("\REPORTES\RRHH\rr_planilla_personal_a_prueba.rpt")
    
  ElseIf optRep020.Value = True Then
    Call Reportes5("\REPORTES\RRHH\rr_personal_a_prueba_migrar.rpt")
    
End If

  
'    Call Reportes("CONSALDO", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep001.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep001.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
'  ElseIf optRep001.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
'    'VENTAS ACUMULADAS POR MES
'  ElseIf optRep002.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\VENTAS_MENSUALES.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep002.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\VENTAS_MENSUALES.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep002.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\VENTAS_MENSUALES.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep002.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\VENTAS_MENSUALES.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS POR PROVEEDOR Y LINEA
'  ElseIf optRep003.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\COMISION_VENTA_prov.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep003.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\COMISION_VENTA_prov.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep003.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\COMISION_VENTA_prov.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep003.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\COMISION_VENTA_prov.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS Y COBRANZAS POR CLIENTE (Detalle)
'  ElseIf optRep004.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\COMISION_VENTA_CLI.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep004.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\COMISION_VENTA_CLI.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep004.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\COMISION_VENTA_CLI.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep004.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\COMISION_VENTA_CLI.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'VENTAS Y COBRANZAS POR CLIENTE (Totales)
'  ElseIf optrep005.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\COMISION_VENTA_CLI_tot.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optrep005.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\COMISION_VENTA_CLI_tot.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optrep005.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\COMISION_VENTA_CLI_tot.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optrep005.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\COMISION_VENTA_CLI_tot.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'COMISIONES POR VENTAS Y COBRANZAS
'  ElseIf optrep006.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\COMISION_VENTA.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optrep006.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\COMISION_VENTA.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optrep006.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\COMISION_VENTA.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optrep006.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\COMISION_VENTA.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
'    'SEGUIMIENTO DE VENTAS POR PRODUCTO
'  ElseIf optRep007.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep007.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\VENTAS_PRODUCTO.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep007.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\VENTAS_PRODUCTO.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep007.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\VENTAS_PRODUCTO.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'
''  ElseIf optRep008.Value = True Then
''    'Call Reportes("\RRHH\Reportes\COMISION_VENTA_HIST_cli.rpt ")
''  ElseIf optRep009.Value = True Then
''    'Call Reportes("\RRHH\Reportes\COMISION_VENTA_HIST.rpt ")
''  ElseIf optRep0010.Value = True Then
''    '
'''  ElseIf optRep0011.Value = True Then
''    '
''  'End If
'
'  'LISTADO GENERAL DE VENTAS
'  ElseIf optRep001.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep001.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep001.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
'  ElseIf optRep001.Value = True And opt_4.Value = True Then
'    Call Reportes("TODAS", "\Reportes\RRHH\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")

'  'PLANILLA MINISTERIO DE TRABAJO 'rr_planilla_ministerio
'  ElseIf optRep010.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\ar_lista_cobranzas_facturadas_dol.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep010.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\ar_lista_cobranzas_facturadas_dol.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep010.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\ar_lista_cobranzas_facturadas_dol.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
'  ElseIf optRep010.Value = True And opt_4.Value = True Then
'    'Call Reportes("TODAS", "\Reportes\RRHH\ar_lista_cobranzas_facturadas_dol.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'    CryUnidad.ReportFileName = App.Path & "\Reportes\RRHH\ar_lista_cobranzas_facturadas_dol.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If

'  'LIBRO DE VENTAS
'  ElseIf optRep011.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\ar_libro_ventas.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep011.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep011.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep011.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\RRHH\ar_libro_ventas.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'    'Call Reportes("TODAS", "\Reportes\RRHH\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'  'End If
'
'  'COBRANZAS POR FACTURA
'  ElseIf optRep012.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep012.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep012.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep012.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\RRHH\ar_lista_cobranzas_solo_facturadas.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'
'  'COBRANZAS POR COBRADOR
'  ElseIf optRep015.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep015.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep015.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep015.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\RRHH\ar_lista_cobranzas_facturadas_Cobrador.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'
'        'COBRANZAS POR RECIBO
'  ElseIf optRep016.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep016.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep016.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep016.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\RRHH\ar_lista_cobranzas_solo_recibo.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'
'    'Call Reportes("TODAS", "\Reportes\RRHH\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'  End If

End Sub

'Private Sub BtnImprimir2_Click()
'    'LIBRO DE VENTAS
'  If optRep011.Value = True And Opt_1.Value = True Then
'    Call Reportes("CONSALDO", "\Reportes\RRHH\ar_libro_ventas_txt.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep011.Value = True And opt_2.Value = True Then
'    Call Reportes("SINSALDO", "\Reportes\RRHH\ar_libro_ventas_txt.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep011.Value = True And Opt_3.Value = True Then
'    Call Reportes("MONTOCERO", "\Reportes\RRHH\ar_libro_ventas_txt.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep011.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\RRHH\ar_libro_ventas_txt.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'    'Call Reportes("TODAS", "\Reportes\RRHH\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'  End If
'End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub butEstProg_Click()
'  frmListaEstProg.Show
End Sub

Private Sub Reportes(ArchRep As String)
'Private Sub Reportes(tipoRep As String, ArchRep As String, titulo1 As String)
'ini reporte
  CryReporte.ReportFileName = App.Path & ArchRep
'  CryReporte.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryReporte.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryReporte.StoredProcParam(0) = tipoRep
''  Call setParametros(CryReporte)
'  CryReporte.Formulas(0) = "FFInicio ='" & dtpFecha1.Value & "'"
'  CryReporte.Formulas(1) = "FFFinal ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryReporte.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryReporte.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If
'fin reporte
  iResult = CryReporte.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte.LastErrorNumber & " : " & CryReporte.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub Reportes2(ArchRep As String)
''Private Sub Reportes(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte2.ReportFileName = App.Path & ArchRep
'  CryReporte2.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryReporte2.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryReporte2.StoredProcParam(0) = tipoRep

   CryReporte2.StoredProcParam(0) = cmb_gestion_rep.Text
  If dtc_rep_cod.Text = "" Then
        CryReporte2.StoredProcParam(1) = "%"
  Else
        CryReporte2.StoredProcParam(1) = dtc_rep_cod.Text
  End If
   CryReporte2.StoredProcParam(2) = txt_mes.Text
   
 If Option3.Value = True Then
 CryReporte2.StoredProcParam(1) = "%"
  CryReporte2.StoredProcParam(2) = txt_mes.Text
 End If
 
 
   
  If optrep006.Value = True Then
  CryReporte2.StoredProcParam(3) = "1"
  CryReporte2.StoredProcParam(4) = "%"

  End If
  
   If optrep006B.Value = True Or optrep006C.Value Or optrep006D.Value Then
  CryReporte2.StoredProcParam(3) = "1"
  CryReporte2.StoredProcParam(4) = "%"

  End If
  
   If optRep018.Value = True Then
  CryReporte2.StoredProcParam(1) = Dtc_depto.Text
  CryReporte2.StoredProcParam(2) = txt_mes.Text
  End If
  If Option4.Value = True Then
  If cb_aguinaldo.Text = "AGUINALDO 1" Then
  CryReporte2.StoredProcParam(2) = "13"
  End If
  If cb_aguinaldo.Text = "AGUINALDO 2" Then
  CryReporte2.StoredProcParam(2) = "14"
  End If
  End If
   'CryReporte2.StoredProcParam(2) = dtc_rep_cod.Text
'ini reporte
'fin reporte
'  Call setParametros(CryReporte2)
'  CryReporte2.Formulas(0) = "FFInicio ='" & dtpFecha1.Value & "'"
'  CryReporte2.Formulas(1) = "FFFinal ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryReporte2.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryReporte2.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If

  iResult = CryReporte2.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte2.LastErrorNumber & " : " & CryReporte2.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub Reportes3(ArchRep As String)
'Private Sub Reportes(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte3.ReportFileName = App.Path & ArchRep
'  CryReporte3.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryReporte3.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryReporte3.StoredProcParam(0) = tipoRep

   CryReporte3.StoredProcParam(0) = cmb_gestion_rep.Text
  If dtc_rep_cod.Text = "" Then
        CryReporte3.StoredProcParam(1) = "%"
  Else
        CryReporte3.StoredProcParam(1) = dtc_rep_cod.Text
  End If
   CryReporte3.StoredProcParam(2) = txt_mes.Text
  CryReporte3.StoredProcParam(3) = txt_convenio.Text
  CryReporte3.StoredProcParam(4) = dtc_cta.Text
  CryReporte3.StoredProcParam(5) = txt_archivo
  CryReporte3.StoredProcParam(6) = Format(DTp_FCarga.Value, "dd/mm/yyyy")
  CryReporte3.StoredProcParam(7) = Format(DTP_Fvigencia.Value, "dd/mm/yyyy")

'  CryReporte3.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'
'@codigo_cgi VARCHAR(15),
'@cuenta_bancaria VARCHAR(15),
'@nombre_envio VARCHAR(15),
'@fecha_solicitud DATE,
'@fecha_autorizada DATE
'ini reporte
  If optRep007.Value = True Then
'    If DtcProdC.Text = "" Then
'        CryReporte3.StoredProcParam(7) = "%"
'    Else
'        CryReporte3.StoredProcParam(7) = DtcProdC.Text
'    End If
  End If
''fin reporte
''  Call setParametros(CryReporte3)
'  CryReporte3.Formulas(0) = "FFInicio ='" & dtpFecha1.Value & "'"
'  CryReporte3.Formulas(1) = "FFFinal ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryReporte3.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryReporte3.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If
  iResult = CryReporte3.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte3.LastErrorNumber & " : " & CryReporte3.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub Reportes4(ArchRep As String)
''Private Sub Reportes(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte4.ReportFileName = App.Path & ArchRep

   CryReporte4.StoredProcParam(0) = cmb_gestion_rep.Text
  If dtc_rep_cod.Text = "" Then
        CryReporte4.StoredProcParam(1) = "%"
  Else
        CryReporte4.StoredProcParam(1) = dtc_rep_cod.Text
  End If
   CryReporte4.StoredProcParam(2) = txt_mes.Text
   
'ini reporte
  
  If optRep012.Value = True Then
    CryReporte4.StoredProcParam(3) = "4"
  End If
  If optRep013.Value = True Then
    CryReporte4.StoredProcParam(3) = "3"
  End If
  If optRep014.Value = True Then
    'CryReporte4.StoredProcParam(3) = "3"
  End If
  
  iResult = CryReporte4.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte4.LastErrorNumber & " : " & CryReporte4.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub Reportes5(ArchRep As String)
''Private Sub Reportes(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte5.ReportFileName = App.Path & ArchRep
'  CryReporte2.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryReporte2.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryReporte2.StoredProcParam(0) = tipoRep

   CryReporte5.StoredProcParam(0) = cmb_gestion_rep.Text
  If dtc_rep_cod.Text = "" Then
        CryReporte5.StoredProcParam(1) = "%"
  Else
        CryReporte5.StoredProcParam(1) = dtc_rep_cod.Text
  End If
   If optRep018.Value = True Then
        CryReporte5.StoredProcParam(1) = Dtc_depto.Text
  End If
  If optRep020.Value = True Then
        CryReporte5.StoredProcParam(1) = Dtc_depto.Text
  End If
  
   CryReporte5.StoredProcParam(2) = txt_mes.Text
   
 
'ini reporte
'fin reporte
'  Call setParametros(CryReporte2)
'  CryReporte2.Formulas(0) = "FFInicio ='" & dtpFecha1.Value & "'"
'  CryReporte2.Formulas(1) = "FFFinal ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryReporte2.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryReporte2.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If

If optRep018.Value = True Or optRep020.Value = True Then
CryReporte5.Formulas(1) = "cuenta ='" & dtc_cta.Text & "'"
CryReporte5.Formulas(2) = "fecha_carga ='" & DTp_FCarga.Value & "'"
CryReporte5.Formulas(3) = "fecha_vigencia ='" & DTP_Fvigencia.Value & "'"
CryReporte5.Formulas(4) = "nombre ='" & txt_convenio.Text & "'"
CryReporte5.Formulas(5) = "numero_envio ='" & txt_archivo.Text & "'"
If txt_mes.Text > 12 Then
CryReporte5.Formulas(6) = "AGUINALDO"
Else
CryReporte5.Formulas(6) = "mes ='" & UCase(MonthName(txt_mes.Text)) & "'"
End If
CryReporte5.Formulas(7) = "gestion ='" & cmb_gestion_rep.Text & "'"
Else
   CryReporte5.StoredProcParam(3) = txt_convenio.Text
   CryReporte5.StoredProcParam(4) = dtc_cta.Text
   CryReporte5.StoredProcParam(5) = txt_archivo.Text
   CryReporte5.StoredProcParam(6) = DTp_FCarga.Value
   CryReporte5.StoredProcParam(7) = DTP_Fvigencia.Value
End If

  iResult = CryReporte5.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte5.LastErrorNumber & " : " & CryReporte5.LastErrorString, vbCritical + vbOKOnly, "Error..."
  
  End If
End Sub

Private Sub Reportes6(ArchRep As String)
''Private Sub Reportes(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte6.ReportFileName = App.Path & ArchRep

  CryReporte6.StoredProcParam(0) = cmb_gestion_rep.Text
  If dtc_rep_cod.Text = "" Then
        CryReporte6.StoredProcParam(1) = "%"
  Else
        CryReporte6.StoredProcParam(1) = dtc_rep_cod.Text
  End If
   CryReporte6.StoredProcParam(2) = txt_mes.Text
   
   CryReporte6.StoredProcParam(3) = dtc_org.Text
'ini reporte
'fin reporte
'  Call setParametros(CryReporte2)
'  CryReporte2.Formulas(0) = "FFInicio ='" & dtpFecha1.Value & "'"
'  CryReporte2.Formulas(1) = "FFFinal ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryReporte2.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryReporte2.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If
  iResult = CryReporte6.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte6.LastErrorNumber & " : " & CryReporte6.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub Rep001(tipoRep As String, ArchRep As String, titulo1 As String)
  CryReporte.ReportFileName = App.Path & ArchRep
'  CryReporte.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryReporte.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryReporte.StoredProcParam(2) = tipoRep
'  Call setParametros(CryReporte)
'  CryReporte.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
'  CryReporte.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
'  If titulo1 <> "" Then
'    CryReporte.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
'  End If
'  If ArchRep = "\rep002.rpt" Then
'     CryReporte.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If
  iResult = CryReporte.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte.LastErrorNumber & " : " & CryReporte.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub RepDetalle(tipoRep As String, ArchRep As String, titulo1 As String)
  CryDetalle.ReportFileName = App.Path & ArchRep
  CryDetalle.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
  CryDetalle.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryDetalle.StoredProcParam(2) = tipoRep
  Call setParametros(CryDetalle)
  CryDetalle.Formulas(0) = "fFecha1 ='" & DTPfecha1.Value & "'"
  CryDetalle.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  If titulo1 <> "" Then
    CryDetalle.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
  End If
  If ArchRep = "\rep002.rpt" Then
     CryDetalle.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
  End If
  iResult = CryDetalle.PrintReport
  If iResult <> 0 Then
    MsgBox CryDetalle.LastErrorNumber & " : " & CryDetalle.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub setParametros(objCryRep As Object)
'  If dcmFte_codigo.Text = "" Then
'    objCryRep.StoredProcParam(3) = "%"
'  Else
'    objCryRep.StoredProcParam(3) = dcmFte_codigo.BoundText
'  End If
'  If cdmOrganismo.Text = "" Then
'    objCryRep.StoredProcParam(4) = "%"
'  Else
'    objCryRep.StoredProcParam(4) = cdmOrganismo.BoundText
'  End If
'  If dtcboconvenio.Text = "" Then
'    objCryRep.StoredProcParam(5) = "%"
'  Else
'    objCryRep.StoredProcParam(5) = dtcboconvenio.BoundText
'  End If
'  If txtProg.Text = "" Then
'    objCryRep.StoredProcParam(6) = "%"
'  Else
'    objCryRep.StoredProcParam(6) = txtProg.Text
'  End If
'  If txtSubProg.Text = "" Then
'    objCryRep.StoredProcParam(7) = "%"
'  Else
'    objCryRep.StoredProcParam(7) = txtSubProg.Text
'  End If
'  If TxtProy.Text = "" Then
'    objCryRep.StoredProcParam(8) = "%"
'  Else
'    objCryRep.StoredProcParam(8) = TxtProy.Text
'  End If
'  If txtAct.Text = "" Then
'    objCryRep.StoredProcParam(9) = "%"
'  Else
'    objCryRep.StoredProcParam(9) = txtAct.Text
'  End If
'  If txtpartida.Text = "" Then
'    objCryRep.StoredProcParam(10) = "%"
'  Else
'    objCryRep.StoredProcParam(10) = txtpartida.Text
'  End If
End Sub

Private Sub llena_datos()
'  Set tFc_fuente_financiamiento = New ADODB.Recordset
'  If tFc_fuente_financiamiento.State = 1 Then tFc_fuente_financiamiento.Close
'    tFc_fuente_financiamiento.Open "SELECT fte_codigo, fte_codigo + '  ' + fte_descripcion_larga as fte_descripcion_larga FROM fc_fuente_financiamiento order by fte_codigo ", db, adOpenDynamic, adLockOptimistic
'  Set frmRepPresupuesto.Adodc_p.Recordset = tFc_fuente_financiamiento

    
    Set rs_proveedor = New ADODB.Recordset
    If rs_proveedor.State = 1 Then rs_proveedor.Close
    rs_proveedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=3 OR tipoben_codigo=22) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    'rs_proveedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=2 OR tipoben_codigo=22) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_proveedor.Recordset = rs_proveedor
    Ado_proveedor.Refresh

    Set rs_cliente = New ADODB.Recordset
    If rs_cliente.State = 1 Then rs_cliente.Close
    rs_cliente.Open "select * from gc_beneficiario WHERE (tipoben_codigo <> 1 AND tipoben_codigo <> 23) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    'rs_cliente.Open "select * from gc_beneficiario WHERE (tipoben_codigo <> 2 AND tipoben_codigo <> 22)  ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    Set ado_Cliente.Recordset = rs_cliente
    ado_Cliente.Refresh

    Set rs_vendedor = New ADODB.Recordset
    If rs_vendedor.State = 1 Then rs_vendedor.Close
    'rs_vendedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=6 OR tipoben_codigo=10) and (beneficiario_deudor = 'SI') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    rs_vendedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=1 OR tipoben_codigo=0) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_vendedor.Recordset = rs_vendedor
    Ado_vendedor.Refresh

    Set rs_cobrador = New ADODB.Recordset
    If rs_cobrador.State = 1 Then rs_cobrador.Close
    'rs_cobrador.Open "select * from gc_beneficiario WHERE (tipoben_codigo=7 OR tipoben_codigo=10) and (beneficiario_deudor = 'SI') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    rs_cobrador.Open "select * from gc_beneficiario WHERE (tipoben_codigo=1 OR tipoben_codigo=0) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Cobrador.Recordset = rs_cobrador
    Ado_Cobrador.Refresh

    Set rs_tipo = New ADODB.Recordset
    If rs_tipo.State = 1 Then rs_tipo.Close
    rs_tipo.Open "select venta_tipo, venta_tipo_descripcion from ac_tipo_compra_venta WHERE estado_codigo='APR' ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Tipo.Recordset = rs_tipo
    Ado_Tipo.Refresh

    Set rs_tipoBenef = New ADODB.Recordset
    If rs_tipoBenef.State = 1 Then rs_tipoBenef.Close
    rs_tipoBenef.Open "select tipoben_codigo, tipoben_Descripcion from gc_tipo_beneficiario WHERE (ESTADO_codigo='APR') ", db, adOpenKeyset, adLockReadOnly
    Set Ado_TipoBenef.Recordset = rs_tipoBenef
    Ado_TipoBenef.Refresh

    Set rs_ciudad = New ADODB.Recordset
    If rs_ciudad.State = 1 Then rs_ciudad.Close
    'rs_ciudad.Open "select Depto AS procedencia, municipio AS lugar_procedencia from gc_beneficiario WHERE (tipoben_codigo<>'B' and tipoben_codigo<>'O' and tipoben_codigo<>'P') and (activo = 'S') group BY Depto, municipio ", DB, adOpenKeyset, adLockReadOnly
    rs_ciudad.Open "select Depto_codigo , munic_codigo from gc_beneficiario WHERE (tipoben_codigo <>0 ) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') group BY Depto_codigo, munic_codigo ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Ciudad.Recordset = rs_ciudad
    Ado_Ciudad.Refresh
    
    ' rc_planilla_grupo
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "SELECT * FROM rc_planilla_grupo", db, adOpenStatic
    Set Ado_datos_rep.Recordset = rs_aux7
    dtc_rep_det.BoundText = dtc_rep_cod.BoundText
  
    'fc_cuenta_bancaria
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "SELECT * FROM fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_cuenta.Recordset = rs_aux8
    dtc_ctades.BoundText = dtc_cta.BoundText
  
    'fc_organismo_financiamiento
    Set rs_aux9 = New ADODB.Recordset
    If rs_aux9.State = 1 Then rs_aux8.Close
    rs_aux9.Open "SELECT * FROM fc_organismo_financiamiento", db, adOpenStatic
    Set Ado_org.Recordset = rs_aux9
    dtc_orgDes.BoundText = dtc_org.BoundText
  
'    Set rs_meses = New ADODB.Recordset
'    If rs_meses.State = 1 Then rs_meses.Close
'    rs_meses.Open "select * from gc_periodos WHERE (estado_registro = 'S') ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_Meses.Recordset = rs_meses
'    Ado_Meses.Refresh
    
    Set rs_producto = New ADODB.Recordset
    If rs_producto.State = 1 Then rs_producto.Close
    rs_producto.Open "select bien_codigo, concepto_venta from ao_ventas_detalle group BY bien_codigo, concepto_venta ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Producto.Recordset = rs_producto
    Ado_Producto.Refresh
    
    DtcProvCod.Enabled = False
    DtcProvDes.Enabled = False
    DtcCliCod.Enabled = True
    DtcCliDes.Enabled = True
    DtcVenCod.Enabled = True
    DtcVenDes.Enabled = True
    DtcCbrCod.Enabled = False
    DtcCbrDes.Enabled = False
    DtcMes.Enabled = False
    DtcMesC.Enabled = False
    DtcProd.Enabled = False
    DtcProdC.Enabled = False
End Sub
Private Sub showEtiquetas(mostrar As Boolean)
  If mostrar Then
    lblFuente.Visible = True
    lblOrg.Visible = True
    lblConv.Visible = True

  Else
    lblFuente.Visible = False
    lblOrg.Visible = False
    lblConv.Visible = False
    
  End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub cbo_mes_rep_LostFocus()
   ' BtnImprimir.Visible = True
    txt_mes.Text = cbo_mes_rep.ListIndex
    txt_mes.Text = Val(txt_mes.Text) + 1
End Sub

Private Sub dtc_cta_Click(Area As Integer)
    dtc_ctades.BoundText = dtc_cta.BoundText
End Sub

Private Sub dtc_ctades_Click(Area As Integer)
    dtc_cta.BoundText = dtc_ctades.BoundText
End Sub

Private Sub Dtc_Org_Click(Area As Integer)
    dtc_orgDes.BoundText = dtc_org.BoundText
End Sub

Private Sub dtc_orgDes_Click(Area As Integer)
    dtc_org.BoundText = dtc_orgDes.BoundText
End Sub

Private Sub dtc_rep_cod_Click(Area As Integer)
    dtc_rep_det.BoundText = dtc_rep_cod.BoundText
    Dtc_depto.BoundText = dtc_rep_cod.BoundText
    Option1.Value = False
'    dtc_rep_cod.Text = ""
'    dtc_rep_det.Text = ""
End Sub

Private Sub dtc_rep_det_Click(Area As Integer)
    dtc_rep_cod.BoundText = dtc_rep_det.BoundText
    Dtc_depto.BoundText = dtc_rep_det.BoundText
    Option1.Value = False
     Option2.Value = False
'    dtc_rep_cod.Text = ""
'    dtc_rep_det.Text = ""
End Sub

Private Sub DtcCbrCod_Click(Area As Integer)
    DtcCbrDes.BoundText = DtcCbrCod.BoundText
End Sub

Private Sub DtcCbrDes_Click(Area As Integer)
    DtcCbrCod.BoundText = DtcCbrDes.BoundText
End Sub

Private Sub DtcCliCod_Click(Area As Integer)
    DtcCliDes.BoundText = DtcCliCod.BoundText
End Sub

Private Sub DtcCliDes_Click(Area As Integer)
    DtcCliCod.BoundText = DtcCliDes.BoundText
End Sub

Private Sub DtcCiu_Click(Area As Integer)
    DtcDepto.BoundText = DtcCiu.BoundText
End Sub

Private Sub DtcDepto_Click(Area As Integer)
    DtcCiu.BoundText = DtcDepto.BoundText
End Sub

Private Sub DtcProvCod_Click(Area As Integer)
    DtcProvDes.BoundText = DtcProvCod.BoundText
End Sub

Private Sub DtcProvDes_Click(Area As Integer)
    DtcProvCod.BoundText = DtcProvDes.BoundText
End Sub

Private Sub DtcTipo_Click(Area As Integer)
    DtcTipoDes.BoundText = DtcTipo.BoundText
End Sub

Private Sub dtctipoDes_Click(Area As Integer)
    DtcTipo.BoundText = DtcTipoDes.BoundText
End Sub

Private Sub DtcTipoCli_Click(Area As Integer)
    DtcTipoCliDes.BoundText = DtcTipoCli.BoundText
End Sub

Private Sub DtcTipoCliDes_Click(Area As Integer)
    DtcTipoCli.BoundText = DtcTipoCliDes.BoundText
End Sub

Private Sub DtcVenCod_Click(Area As Integer)
    DtcVenDes.BoundText = DtcVenCod.BoundText
End Sub

Private Sub DtcVenDes_Click(Area As Integer)
    DtcVenCod.BoundText = DtcVenDes.BoundText
End Sub


Private Sub Form_Load()
'    BtnImprimir2.Visible = False
    Call llena_datos
    cmb_gestion_rep.Text = Year(Date)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Private Sub opt_rep009_Click()
  Call SetControles(False, True)
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
dtc_rep_cod.Text = "%"
dtc_rep_det.Text = "TODAS LAS PLANILLAS"
Dtc_depto.Text = "%"
Else
dtc_rep_cod.Text = ""
dtc_rep_det.Text = ""
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
dtc_rep_cod.Text = "INT"
dtc_rep_det.Text = "TODAS LAS DEL INTERIOR"
Dtc_depto.Text = "INT"
Else
dtc_rep_cod.Text = ""
dtc_rep_det.Text = ""
Dtc_depto.Text = ""
End If
End Sub

Private Sub Option3_Click()
 Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = True
      Option2.Visible = False
      Frame3.Visible = False
      
      
       cbo_mes_rep.Visible = True
 
 cb_aguinaldo.Visible = False
 
End Sub



Private Sub Option4_Click()
 Frame1.Visible = True
 Frame2.Visible = False
 Frame3.Visible = True
 Option2.Visible = False
 Frame3.Visible = False
 cbo_mes_rep.Visible = False
 
 cb_aguinaldo.Visible = True
 
 
End Sub

Private Sub optRep001_Click()
  Frame1.Visible = False
  Frame2.Visible = False
  BtnImprimir.Visible = True
  Option2.Visible = False
  Call SetControles(True, False)
  DtcProvCod.Enabled = False
  DtcProvDes.Enabled = False
  DtcCliCod.Enabled = True
  DtcCliDes.Enabled = True
  DtcVenCod.Enabled = True
  DtcVenDes.Enabled = True
  DtcCbrCod.Enabled = False
  DtcCbrDes.Enabled = False
  DtcTipo.Enabled = True
  DtcTipoDes.Enabled = True
  DtcTipoCliDes.Enabled = True
  DtcCiu.Enabled = True
  DtcMes.Enabled = False
  DtcMesC.Enabled = False
  DtcProd.Enabled = False
  DtcProdC.Enabled = False
End Sub

Private Sub optRep002_Click()
  Frame1.Visible = False
  Frame2.Visible = False
  BtnImprimir.Visible = True
    Option2.Visible = False
  
  Call SetControles(True, False)
  DtcProvCod.Enabled = False
  DtcProvDes.Enabled = False
  DtcCliCod.Enabled = False
  DtcCliDes.Enabled = False
  DtcVenCod.Enabled = False
  DtcVenDes.Enabled = False
  DtcCbrCod.Enabled = False
  DtcCbrDes.Enabled = False
  DtcTipo.Enabled = False
  DtcTipoDes.Enabled = False
  DtcTipoCliDes.Enabled = True
  DtcCiu.Enabled = True
  DtcMes.Enabled = True
  DtcMesC.Enabled = True
  DtcProd.Enabled = False
  DtcProdC.Enabled = False
End Sub

Private Sub optRep003_Click()
  Frame1.Visible = False
  Frame2.Visible = False
  BtnImprimir.Visible = True
    Option2.Visible = False
  
  Call SetControles(True, False)
  DtcProvCod.Enabled = True
  DtcProvDes.Enabled = True
  DtcCliCod.Enabled = True
  DtcCliDes.Enabled = True
  DtcVenCod.Enabled = False
  DtcVenDes.Enabled = False
  DtcCbrCod.Enabled = False
  DtcCbrDes.Enabled = False
  DtcTipo.Enabled = True
  DtcTipoDes.Enabled = True
  DtcTipoCliDes.Enabled = True
  DtcCiu.Enabled = True
  DtcMes.Enabled = False
  DtcMesC.Enabled = False
  DtcProd.Enabled = False
  DtcProdC.Enabled = False
End Sub

Private Sub optRep004_Click()
    Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
    'BtnImprimir.Visible = False
     cbo_mes_rep.Visible = True
 
 cb_aguinaldo.Visible = False
 
End Sub

Private Sub optRep005_Click()
    Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
'    BtnImprimir.Visible = False
 cbo_mes_rep.Visible = True
 
 cb_aguinaldo.Visible = False
 
End Sub

Private Sub optRep006_Click()
    Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
'    BtnImprimir.Visible = False
End Sub
Private Sub optRep006B_Click()
    Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
'    BtnImprimir.Visible = False
End Sub

Private Sub optrep006C_Click()
 Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
End Sub

Private Sub optrep006D_Click()
 Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
End Sub

Private Sub optRep007_Click()
  Call SetControles(True, False)
  DtcProvCod.Enabled = False
  DtcProvDes.Enabled = False
  DtcCliCod.Enabled = False
  DtcCliDes.Enabled = False
  DtcVenCod.Enabled = False
  DtcVenDes.Enabled = False
  DtcCbrCod.Enabled = False
  DtcCbrDes.Enabled = False
  DtcTipo.Enabled = False
  DtcTipoDes.Enabled = False
  DtcTipoCliDes.Enabled = False
  DtcCiu.Enabled = False
  DtcMes.Enabled = False
  DtcMesC.Enabled = False
  DtcProd.Enabled = True
  DtcProdC.Enabled = True
    Option2.Visible = False
End Sub

Private Sub optRep008_Click()
  Call SetControles(False, False)
    Option2.Visible = False
End Sub
Private Sub optRep0010_Click()
  FrameTipo.Visible = False
    Option2.Visible = False
End Sub
Private Sub optRep0011_Click()
  Call SetControles(False, True)
End Sub
Private Sub optRep002Finanzas_Click()
  Call SetControles(False, True)
End Sub
Private Sub SetControles(tipo, conDet As Boolean)
'  FrameTipo.Visible = tipo
'  FrameConDet.Visible = conDet
End Sub

Private Sub optRep009_Click()
  Option2.Visible = False
End Sub

'Private Sub RepVsLeyFinanciador(tipoRep As String, ArchRep As String, titulo1 As String)
'  CryRep002_financiador.ReportFileName = App.Path & ArchRep
'  CryRep002_financiador.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
'  CryRep002_financiador.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'  CryRep002_financiador.StoredProcParam(2) = tipoRep
'  Call setParametros(CryRep002_financiador)
'  CryRep002_financiador.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
'  CryRep002_financiador.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
'  CryRep002_financiador.Formulas(2) = "conDetalle = " & IIf(optSi.Value = True, "true", "false")
'  iResult = CryRep002_financiador.PrintReport
'  If iResult <> 0 Then
'    MsgBox CryRep002_financiador.LastErrorNumber & " : " & CryRep002_financiador.LastErrorString, vbCritical + vbOKOnly, "Error..."
'  End If
'End Sub

Private Sub optRep010_Click()
    Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
'    BtnImprimir.Visible = False
 cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False

End Sub
Private Sub optRep010_B_Click()
    Frame1.Visible = True
    Frame2.Visible = False
      Option2.Visible = False
       cbo_mes_rep.Visible = True
 
 cb_aguinaldo.Visible = False

'    BtnImprimir.Visible = False
End Sub
Private Sub optRep011_Click()
  Frame1.Visible = True
  Frame2.Visible = False
'  BtnImprimir.Visible = False
   cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False

  Call SetControles(True, False)
  DtcProvCod.Enabled = False
  DtcProvDes.Enabled = False
  DtcCliCod.Enabled = True
  DtcCliDes.Enabled = True
  DtcVenCod.Enabled = True
  DtcVenDes.Enabled = True
  DtcCbrCod.Enabled = False
  DtcCbrDes.Enabled = False
  DtcTipo.Enabled = True
  DtcTipoDes.Enabled = True
  DtcTipoCliDes.Enabled = True
  DtcCiu.Enabled = True
  DtcMes.Enabled = False
  DtcMesC.Enabled = False
  DtcProd.Enabled = False
  DtcProdC.Enabled = False
'  BtnImprimir2.Visible = True
  Option2.Visible = False
End Sub

Private Sub optRep012_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
      Option2.Visible = False
'    BtnImprimir.Visible = False
 cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False

End Sub

Private Sub optRep013_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
      Option2.Visible = False
        Option2.Visible = False
'    BtnImprimir.Visible = False
 cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False

End Sub

Private Sub optRep014_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
'    BtnImprimir.Visible = False
 cbo_mes_rep.Visible = True
 
 cb_aguinaldo.Visible = False
 
End Sub

Private Sub optRep015_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Option2.Visible = False
'    BtnImprimir.Visible = False
 cbo_mes_rep.Visible = True
 
 cb_aguinaldo.Visible = False
 
End Sub

Private Sub optRep016_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = True
      Option2.Visible = False
'    BtnImprimir.Visible = False
 
  cbo_mes_rep.Visible = True
 
 cb_aguinaldo.Visible = False
 
End Sub

Private Sub optRep017_Click()
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
      Option2.Visible = True
'    BtnImprimir.Visible = False
 cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False
End Sub

Private Sub optRep018_Click()
 Frame1.Visible = True
    Frame2.Visible = True
      Frame3.Visible = False
      Option2.Visible = False
         Option2.Visible = True

 cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False

End Sub

Private Sub optRep019_Click()
Frame1.Visible = True
    Frame2.Visible = False
      Frame3.Visible = False
      Option2.Visible = False
         Option2.Visible = True

 cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False

End Sub

Private Sub optRep020_Click()
Frame1.Visible = True
    Frame2.Visible = True
      Frame3.Visible = False
      Option2.Visible = False
         Option2.Visible = True

 cbo_mes_rep.Visible = True

 cb_aguinaldo.Visible = False

End Sub

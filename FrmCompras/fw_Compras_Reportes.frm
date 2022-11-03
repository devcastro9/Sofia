VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_Compras_Reportes 
   BackColor       =   &H00000000&
   Caption         =   "REPORTES COMPRAS Y PAGOS"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   Icon            =   "fw_Compras_Reportes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Fechas de Ventas y/o Cobranzas"
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
      Height          =   795
      Left            =   120
      TabIndex        =   83
      Top             =   5760
      Visible         =   0   'False
      Width           =   9900
      Begin VB.ComboBox cmb_gestion_rep 
         Height          =   315
         Left            =   5400
         TabIndex        =   89
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbo_mes_rep 
         Height          =   315
         Left            =   1320
         TabIndex        =   87
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Todos"
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   8160
         TabIndex        =   85
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GESTION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   4440
         TabIndex        =   88
         Top             =   405
         Width           =   1350
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   840
         TabIndex        =   84
         Top             =   405
         Width           =   1350
      End
   End
   Begin VB.Frame FrameConDet 
      Caption         =   "Con Detalle"
      ForeColor       =   &H000000C0&
      Height          =   600
      Left            =   3480
      TabIndex        =   65
      Top             =   7080
      Visible         =   0   'False
      Width           =   2040
      Begin VB.OptionButton optNo 
         Caption         =   "No"
         Height          =   195
         Left            =   945
         TabIndex        =   67
         Top             =   250
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optSi 
         Caption         =   "Si"
         Height          =   225
         Left            =   105
         TabIndex        =   66
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
      TabIndex        =   60
      Top             =   5640
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
         TabIndex        =   64
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
         Left            =   360
         TabIndex        =   63
         Top             =   1200
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
         TabIndex        =   62
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
         TabIndex        =   61
         Top             =   550
         Width           =   3750
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   900
      Left            =   120
      ScaleHeight     =   840
      ScaleWidth      =   9840
      TabIndex        =   54
      Top             =   40
      Width           =   9900
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H80000015&
         Height          =   600
         Left            =   0
         Picture         =   "fw_Compras_Reportes.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00808000&
         Caption         =   "&Docs."
         Enabled         =   0   'False
         Height          =   720
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   740
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H80000015&
         Height          =   600
         Left            =   8430
         Picture         =   "fw_Compras_Reportes.frx":12CF
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton CmdFoto 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Foto"
         Height          =   720
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Carga Foto de la Persona"
         Top             =   120
         Visible         =   0   'False
         Width           =   740
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
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
         Left            =   4380
         TabIndex        =   59
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame fmrTipoReporte 
      BackColor       =   &H00000000&
      Caption         =   "------------------------Compras---------------------------------------------------------Pagos"
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
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   9900
      Begin VB.OptionButton optRep024 
         BackColor       =   &H00000000&
         Caption         =   "por Fechas"
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
         Left            =   8280
         TabIndex        =   90
         Top             =   960
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optRep023 
         BackColor       =   &H00000000&
         Caption         =   "en Dolares"
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
         Left            =   7680
         TabIndex        =   82
         Top             =   280
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.OptionButton optRep022 
         BackColor       =   &H00000000&
         Caption         =   "para Exportar"
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
         Left            =   7800
         TabIndex        =   81
         Top             =   3400
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optRep021 
         BackColor       =   &H00000000&
         Caption         =   "FACTURACION p/MES c/Equipos "
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4680
         TabIndex        =   80
         Top             =   3380
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.OptionButton optRep020 
         BackColor       =   &H00000000&
         Caption         =   "para Exportar"
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
         Left            =   7800
         TabIndex        =   79
         Top             =   3060
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton BtnImprimir2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Facilito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   620
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optRep019 
         BackColor       =   &H00000000&
         Caption         =   "en Mora"
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
         Left            =   7800
         TabIndex        =   77
         Top             =   2020
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.OptionButton optRep015 
         BackColor       =   &H00000000&
         Caption         =   "COBRANZAS POR COBRADOR"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4680
         TabIndex        =   76
         Top             =   1980
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.OptionButton optRep016 
         BackColor       =   &H00000000&
         Caption         =   "COBRANZAS POR RECIBO"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4680
         TabIndex        =   75
         Top             =   2325
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.OptionButton optRep012 
         BackColor       =   &H00000000&
         Caption         =   "COBRANZAS POR FACTURA (General)"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4680
         TabIndex        =   74
         Top             =   930
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.OptionButton optRep011 
         BackColor       =   &H00000000&
         Caption         =   "LIBRO DE COMPRAS"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4680
         TabIndex        =   73
         Top             =   585
         Width           =   2085
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00000000&
         Caption         =   "COBRANZAS POR CLIENTE"
         ForeColor       =   &H00FFFFC0&
         Height          =   390
         Left            =   4680
         TabIndex        =   72
         Top             =   1635
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00000000&
         Caption         =   "COBRANZAS POR REGIONAL"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   4680
         TabIndex        =   71
         Top             =   1335
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.OptionButton optRep025 
         BackColor       =   &H00000000&
         Caption         =   "COBRANZAS PARA CONCILIACION"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4680
         TabIndex        =   70
         Top             =   2670
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.OptionButton optRep010 
         BackColor       =   &H00000000&
         Caption         =   "PAGOS POR UNIDAD"
         ForeColor       =   &H00FFFFC0&
         Height          =   390
         Left            =   4680
         TabIndex        =   69
         Top             =   240
         Width           =   2460
      End
      Begin VB.OptionButton optRep018 
         BackColor       =   &H00000000&
         Caption         =   "COBRANZAS POR MES"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4680
         TabIndex        =   68
         Top             =   3015
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.OptionButton optRep008 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR COBRADOR"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   1620
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.OptionButton optrep006 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR EJECUTIVO DE VENTAS"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   1965
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.OptionButton optRep004 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR CLIENTE"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   930
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.OptionButton optRep003 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR UNIDAD EJECUTORA"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   585
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.OptionButton optrep005 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR EDIFICIO"
         ForeColor       =   &H00FFFFC0&
         Height          =   390
         Left            =   600
         TabIndex        =   19
         Top             =   1275
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optRep0010 
         BackColor       =   &H00000000&
         Caption         =   "LIBRO DE COMPRAS"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   3360
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.OptionButton optRep009 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR TIPO DE VENTA"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   3040
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.OptionButton optRep007 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS POR EQUIPO"
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   2310
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optRep001 
         BackColor       =   &H00000000&
         Caption         =   "LISTADO GENERAL DE VENTAS"
         ForeColor       =   &H00FFFFC0&
         Height          =   390
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.OptionButton optRep002 
         BackColor       =   &H00000000&
         Caption         =   "VENTAS ACUMULADAS POR MES"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   2655
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   3840
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
      TabIndex        =   18
      Top             =   5685
      Visible         =   0   'False
      Width           =   9900
      Begin MSDataListLib.DataCombo DtcProvCod 
         Height          =   315
         Left            =   6720
         TabIndex        =   32
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
         TabIndex        =   33
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
         TabIndex        =   34
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
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   26
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
         TabIndex        =   28
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
         TabIndex        =   31
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   38
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
         TabIndex        =   35
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   27
         Top             =   405
         Width           =   1575
      End
   End
   Begin VB.TextBox txtPartida 
      Height          =   315
      Left            =   2250
      TabIndex        =   14
      Top             =   2505
      Width           =   1095
   End
   Begin VB.TextBox txtAct 
      Height          =   315
      Left            =   3365
      TabIndex        =   13
      Top             =   2175
      Width           =   510
   End
   Begin VB.TextBox txtProy 
      Height          =   315
      Left            =   2805
      TabIndex        =   12
      Top             =   2175
      Width           =   510
   End
   Begin VB.TextBox txtSubProg 
      Height          =   315
      Left            =   7680
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtProg 
      Height          =   315
      Left            =   2250
      TabIndex        =   10
      Top             =   2175
      Width           =   510
   End
   Begin VB.CommandButton butEstProg 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<-- Elige Estruc. Prog."
      Height          =   315
      Left            =   6960
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5535
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   8355
      TabIndex        =   8
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Fechas de Compras y/o Pagos"
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
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   9900
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   2085
         TabIndex        =   1
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   93913089
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   5565
         TabIndex        =   2
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   93913089
         CurrentDate     =   42735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   405
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   405
         Width           =   1350
      End
   End
   Begin Crystal.CrystalReport CryReporte 
      Left            =   10440
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryVsLey 
      Left            =   10440
      Top             =   1710
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryDetalle 
      Left            =   10455
      Top             =   -30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryUnidad 
      Left            =   10440
      Top             =   1275
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
   Begin Crystal.CrystalReport CryRep002_financiador 
      Left            =   10455
      Top             =   405
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lblEstr 
      Caption         =   "Estructura Programatica :"
      Height          =   255
      Left            =   165
      TabIndex        =   21
      Top             =   2265
      Width           =   1935
   End
   Begin VB.Label lblPartida 
      Caption         =   "Partida :"
      Height          =   255
      Left            =   165
      TabIndex        =   20
      Top             =   2595
      Width           =   855
   End
End
Attribute VB_Name = "fw_Compras_Reportes"
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

Dim titulo2, subtitulo2 As String

Public Sub inicio(Usuario, Proceso As String)
  glRepPresup = Proceso
  Call llena_datos
  dtpFecha1.Value = Format("01/01/2017", "dd/mm/yyyy")
  dtpFecha2.Value = Format(Date, "dd/mm/yyyy")
  'dtpFecha2.Value = Date
'  frmRepPresupuesto.Show
End Sub

Private Sub BtnImprimir_Click()
    'LISTADO GENERAL DE VENTAS
    CryUnidad.Reset
    CryUnidad.WindowShowSearchBtn = True
    CryUnidad.WindowShowRefreshBtn = True
    CryUnidad.WindowShowPrintSetupBtn = True
    CryUnidad.WindowShowGroupTree = True
    CryUnidad.WindowShowZoomCtl = True
    CryUnidad.WindowState = crptMaximized

  If optRep001.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
    
  ElseIf optRep011.Value = True Then
  
      CryUnidad.ReportFileName = App.Path & "\Reportes\Contabilidad\fr_libro_compras.rpt"
        CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
        CryUnidad.StoredProcParam(1) = txt_mes.Text
        If txt_mes.Text = "13" Then
        CryUnidad.StoredProcParam(2) = "0"
        Else
        CryUnidad.StoredProcParam(2) = "1"
        End If
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
  
  
  ElseIf optRep001.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
  ElseIf optRep001.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS ACUMULADAS POR MES
  ElseIf optRep002.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep002.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep002.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optRep002.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS POR PROVEEDOR Y LINEA
  ElseIf optRep003.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep003.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep003.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optRep003.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS Y COBRANZAS POR CLIENTE (Detalle)
  ElseIf optRep004.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep004.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep004.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optRep004.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS Y COBRANZAS POR CLIENTE (Totales)
  ElseIf optrep005.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optrep005.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optrep005.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optrep005.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'COMISIONES POR VENTAS Y COBRANZAS
  ElseIf optrep006.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optrep006.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optrep006.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optrep006.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'SEGUIMIENTO DE VENTAS POR PRODUCTO
'  ElseIf optRep007.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep007.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep007.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep007.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "TODAS LAS VENTAS Y COBRANZAS")

 'VENTAS POR EQUIPO
  ElseIf optRep007.Value = True Then
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt", "VENTAS Y COBRANZAS")
    CryUnidad.ReportFileName = App.Path & "\Reportes\Comercial\ar_bienes_equipos_OTIS.rpt"
'        titulo2 = "MODULO VENTAS"
'        subtitulo2 = "VENTAS POR COBRADOR"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
    'VENTAS POR COBRADOR
  ElseIf optRep008.Value = True Then
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt", "VENTAS Y COBRANZAS")
    CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "VENTAS POR COBRADOR"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
'  ElseIf optRep008.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep008.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep008.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep008.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'        titulo2 = "MODULO VENTAS"
'        subtitulo2 = "COBRANZAS POR COBRADOR"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'  ElseIf optRep008.Value = True Then
'    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST_cli.rpt ")
'  ElseIf optRep009.Value = True Then
'    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST.rpt ")
'  ElseIf optRep0010.Value = True Then
'    '
''  ElseIf optRep0011.Value = True Then
'    '
'  'End If

  'LISTADO GENERAL DE VENTAS
  ElseIf optRep001.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep001.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep001.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
  ElseIf optRep001.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")

  'LISTADO GENERAL DE COBRANZAS
'  ElseIf optRep010.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep010.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep010.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
  ElseIf optRep010.Value = True And opt_4.Value = True Then
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_unidad.rpt"
    titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR UNIDAD"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    
  ElseIf optRep023.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR UNIDAD (Dolares)"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

  'LIBRO DE VENTAS
  ElseIf optRep011.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS NO FACTURADAS")
  ElseIf optRep011.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
  ElseIf optRep011.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y COBRADAS")
  ElseIf optRep011.Value = True And opt_4.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_libro_ventas.rpt"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
  'End If
  
  'COBRANZAS POR FACTURA
  ElseIf optRep012.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS NO FACTURADAS")
  ElseIf optRep012.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
  ElseIf optRep012.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y COBRADAS")
  ElseIf optRep012.Value = True And opt_4.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_facturadas.rpt"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
  'COBRANZAS POR FACTURA - por FECHAS
  ElseIf optRep024.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_diaria_facturas_vs_cobros.rpt"
        titulo2 = "LISTADO DE FACTURACION"
        subtitulo2 = "MODULO COBRANZAS"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
  'COBRANZAS POR COBRADOR
  ElseIf optRep015.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS NO FACTURADAS")
  ElseIf optRep015.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
  ElseIf optRep015.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y COBRADAS")
  ElseIf optRep015.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR COBRADOR"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
 'COBRANZAS POR COBRADOR en MORA
'  ElseIf optRep015.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep015.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep015.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y COBRADAS")
  ElseIf optRep019.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobr_Mora.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS EN MORA POR COBRADOR"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
'COBRANZAS POR MES
  ElseIf optRep018.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR MES"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
'COBRANZAS POR MES (PARA MIGRAR)
  ElseIf optRep020.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes_txt.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR MES (MIGRAR)"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

'COBRANZAS POR MES con EQUIPOS
  ElseIf optRep021.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes_eqp.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS P/MES C/EQUIPOS"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
'COBRANZAS POR MES CON EQUIPOS (PARA MIGRAR)
  ElseIf optRep022.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes_eqp_txt.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS P/MES c/EQUIPOS(MIGRAR)"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
        'COBRANZAS POR RECIBO
'  ElseIf optRep016.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep016.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep016.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y COBRADAS")
    ElseIf optRep016.Value = True And opt_4.Value = True Then
        'CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt"
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_recibo_mes.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR COBRADOR"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
        'COBRANZAS PARA CONCILIACION
    ElseIf optRep025.Value = True And opt_4.Value = True Then
        'CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt"
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_seguimiento_cobranza_concilia.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS PARA CONCILIACION"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
  End If

End Sub

Private Sub BtnImprimir2_Click()
    'LIBRO DE VENTAS
  If optRep011.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_libro_ventas_txt.rpt", "VENTAS NO FACTURADAS")
  ElseIf optRep011.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_libro_ventas_txt.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
  ElseIf optRep011.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_libro_ventas_txt.rpt", "VENTAS FACTURADAS Y COBRADAS")
  ElseIf optRep011.Value = True And opt_4.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_libro_ventas_txt.rpt"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
  End If
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub butEstProg_Click()
'  frmListaEstProg.Show
End Sub

Private Sub cmdAcepta_Click()
    'LISTADO GENERAL DE VENTAS
  If optRep001.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep001.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep001.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "DONACION, OBSEQUIO, PERDIDA (MONTO CERO)")
  ElseIf optRep001.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_CLI_VEN2.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS ACUMULADAS POR MES
  ElseIf optRep002.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep002.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep002.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optRep002.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_MENSUALES.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS POR PROVEEDOR Y LINEA
  ElseIf optRep003.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep003.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep003.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optRep003.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_prov.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS Y COBRANZAS POR CLIENTE (Detalle)
  ElseIf optRep004.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep004.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep004.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optRep004.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'VENTAS Y COBRANZAS POR CLIENTE (Totales)
  ElseIf optrep005.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optrep005.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optrep005.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optrep005.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA_CLI_tot.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'COMISIONES POR VENTAS Y COBRANZAS
  ElseIf optrep006.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optrep006.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\COMISION_VENTA.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optrep006.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\COMISION_VENTA.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optrep006.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\COMISION_VENTA.rpt", "TODAS LAS VENTAS Y COBRANZAS")
    
    'SEGUIMIENTO DE VENTAS POR PRODUCTO
  ElseIf optRep007.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
  ElseIf optRep007.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
  ElseIf optRep007.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
  ElseIf optRep007.Value = True And opt_4.Value = True Then
    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "TODAS LAS VENTAS Y COBRANZAS")

  ElseIf optRep008.Value = True Then
    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST_cli.rpt ")
  ElseIf optRep009.Value = True Then
    'Call RepUnidad("\Ventas\Reportes\COMISION_VENTA_HIST.rpt ")
  ElseIf optRep0010.Value = True Then
    '
'  ElseIf optRep0011.Value = True Then
    '
  End If
End Sub

'Private Sub RepUnidad(tipoRep As String, ArchRep As String)
Private Sub RepUnidad(tipoRep As String, ArchRep As String, titulo1 As String)
  CryUnidad.ReportFileName = App.Path & ArchRep
  If optRep008.Value <> True Then
    CryUnidad.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
    CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
    CryUnidad.StoredProcParam(0) = tipoRep
  End If
'ini reporte
  If DtcProvCod.Text = "" Then
        CryUnidad.StoredProcParam(2) = "%"
  Else
        CryUnidad.StoredProcParam(2) = DtcProvCod.Text
  End If
  If DtcCliCod.Text = "" Then
        CryUnidad.StoredProcParam(3) = "%"
  Else
        CryUnidad.StoredProcParam(3) = DtcCliCod.Text
  End If
  If DtcVenCod.Text = "" Then
        CryUnidad.StoredProcParam(4) = "%"
  Else
        CryUnidad.StoredProcParam(4) = DtcVenCod.Text
  End If
  If DtcCbrCod.Text = "" Then
        CryUnidad.StoredProcParam(5) = "%"
  Else
        CryUnidad.StoredProcParam(5) = DtcCbrCod.Text
  End If
'  If DtcTipo.Text = "" Then
'        CryUnidad.StoredProcParam(6) = "%"
'  Else
'        CryUnidad.StoredProcParam(6) = DtcTipo.Text
'  End If
  CryUnidad.StoredProcParam(6) = tipoRep
  If optRep007.Value = True Then
    If DtcProdC.Text = "" Then
        CryUnidad.StoredProcParam(7) = "%"
    Else
        CryUnidad.StoredProcParam(7) = DtcProdC.Text
    End If
  End If
'fin reporte
'  Call setParametros(CryUnidad)
  CryUnidad.Formulas(0) = "FFInicio ='" & dtpFecha1.Value & "'"
  CryUnidad.Formulas(1) = "FFFinal ='" & dtpFecha2.Value & "'"
  If optRep008.Value = True Then
    subtitulo2 = "VENTAS VS COBRADORES"
    CryUnidad.Formulas(2) = "Titulo = '" & titulo1 & "'"
    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
  Else
    If titulo1 <> "" Then
      CryUnidad.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
    End If
  End If
  
  
'  If ArchRep = "\rep002.rpt" Then
'     CryUnidad.Formulas(2) = "conDetalle = " & IIf(conDetalle, "true", "false")
'  End If
  iResult = CryUnidad.PrintReport
  If iResult <> 0 Then
    MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
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
Private Sub RepVsLey(tipoRep As String, ArchRep As String, titulo1 As String)
  CryVsLey.ReportFileName = App.Path & ArchRep
  CryVsLey.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryVsLey.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryVsLey.StoredProcParam(2) = tipoRep
  Call setParametros(CryVsLey)
  CryVsLey.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
  CryVsLey.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  If titulo1 <> "" Then
    CryVsLey.Formulas(2) = "Titulo1 = '" & titulo1 & "'"
  End If
  If ArchRep = "\Reportes\Presupuesto\rep002.rpt" Or ArchRep = "\Reportes\Presupuesto\Rep002Finanzas.rpt" Then
     CryVsLey.Formulas(2) = "conDetalle = " & IIf(optSi, "true", "false")
  End If
  iResult = CryVsLey.PrintReport
  If iResult <> 0 Then
    MsgBox CryVsLey.LastErrorNumber & " : " & CryVsLey.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub RepDetalle(tipoRep As String, ArchRep As String, titulo1 As String)
  CryDetalle.ReportFileName = App.Path & ArchRep
  CryDetalle.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryDetalle.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryDetalle.StoredProcParam(2) = tipoRep
  Call setParametros(CryDetalle)
  CryDetalle.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
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

Private Sub Command1_Click()
'ok = frmListaEstProg.getcodigo(valor, valor)
'frmListaEstProg.Show
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
    lblEstr.Visible = True
    lblPartida.Visible = True
'    dcmFte_codigo.Visible = True
'    cdmOrganismo.Visible = True
'    dtcboconvenio.Visible = True
    txtProg.Visible = True
    txtSubProg.Visible = True
    txtProy.Visible = True
    txtAct.Visible = True
    butEstProg.Visible = True
    txtPartida.Visible = True
  Else
    lblFuente.Visible = False
    lblOrg.Visible = False
    lblConv.Visible = False
    lblEstr.Visible = False
    lblPartida.Visible = False
'    dcmFte_codigo.Visible = False
'    cdmOrganismo.Visible = False
'    dtcboconvenio.Visible = False
    txtProg.Visible = False
    txtSubProg.Visible = False
    txtProy.Visible = False
    txtAct.Visible = False
    butEstProg.Visible = False
    txtPartida.Visible = False
  End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub



Private Sub cbo_mes_rep_DblClick()
cbo_mes_rep.Visible = True
End Sub

Private Sub cbo_mes_rep_LostFocus()
BtnImprimir.Visible = True
    txt_mes.Text = cbo_mes_rep.ListIndex
    txt_mes.Text = Val(txt_mes.Text) + 1
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
    BtnImprimir2.Visible = False
    Call llena_datos
	Call SeguridadSet(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Private Sub opt_rep009_Click()
  Call SetControles(False, True)
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
cbo_mes_rep.Visible = False

End If

End Sub

Private Sub Option5_Click()
Frame2.Visible = False
End Sub

Private Sub Option7_Click()
Frame2.Visible = False
End Sub

Private Sub Option8_Click()

    Frame2.Visible = False
End Sub

Private Sub optRep001_Click()
Frame2.Visible = False
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

Private Sub optRep0010_LostFocus()
'Frame2.Visible = False
End Sub

Private Sub optRep002_Click()
Frame2.Visible = False
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
Frame2.Visible = False
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
Frame2.Visible = False
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

Private Sub optRep005_Click()
Frame2.Visible = False
  Call SetControles(True, False)
  DtcProvCod.Enabled = False
  DtcProvDes.Enabled = False
  DtcCliCod.Enabled = True
  DtcCliDes.Enabled = True
  DtcVenCod.Enabled = True
  DtcVenDes.Enabled = True
  DtcCbrCod.Enabled = False
  DtcCbrDes.Enabled = False
  DtcTipoCliDes.Enabled = True
  DtcCiu.Enabled = True
  DtcMes.Enabled = False
  DtcMesC.Enabled = False
  DtcProd.Enabled = False
  DtcProdC.Enabled = False
End Sub

Private Sub optRep006_Click()
Frame2.Visible = False
  Call SetControles(True, False)
  DtcProvCod.Enabled = False
  DtcProvDes.Enabled = False
  DtcCliCod.Enabled = True
  DtcCliDes.Enabled = True
  DtcVenCod.Enabled = True
  DtcVenDes.Enabled = True
  DtcCbrCod.Enabled = True
  DtcCbrDes.Enabled = True
  DtcTipo.Enabled = True
  DtcTipoDes.Enabled = True
  DtcTipoCliDes.Enabled = True
  DtcCiu.Enabled = True
  DtcMes.Enabled = False
  DtcMesC.Enabled = False
  DtcProd.Enabled = False
  DtcProdC.Enabled = False
End Sub

Private Sub optRep007_Click()
Frame2.Visible = False
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
End Sub

Private Sub optRep008_Click()
Frame2.Visible = False
  Call SetControles(False, False)
End Sub
Private Sub optRep0010_Click()
  FrameTipo.Visible = False
  Frame2.Visible = True
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

Private Sub RepVsLeyFinanciador(tipoRep As String, ArchRep As String, titulo1 As String)
  CryRep002_financiador.ReportFileName = App.Path & ArchRep
  CryRep002_financiador.StoredProcParam(0) = Format(dtpFecha1.Value, "dd/mm/yyyy")
  CryRep002_financiador.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
  CryRep002_financiador.StoredProcParam(2) = tipoRep
  Call setParametros(CryRep002_financiador)
  CryRep002_financiador.Formulas(0) = "fFecha1 ='" & dtpFecha1.Value & "'"
  CryRep002_financiador.Formulas(1) = "fFecha2 ='" & dtpFecha2.Value & "'"
  CryRep002_financiador.Formulas(2) = "conDetalle = " & IIf(optSi.Value = True, "true", "false")
  iResult = CryRep002_financiador.PrintReport
  If iResult <> 0 Then
    MsgBox CryRep002_financiador.LastErrorNumber & " : " & CryRep002_financiador.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub optRep009_Click()
Frame2.Visible = False
End Sub

Private Sub optRep010_Click()
    Frame2.Visible = False
    optRep023.Visible = True
End Sub

Private Sub optRep011_Click()
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
  Frame2.Visible = False
  DtcTipoCliDes.Enabled = True
  DtcCiu.Enabled = True
  DtcMes.Enabled = False
  DtcMesC.Enabled = False
  DtcProd.Enabled = False
  DtcProdC.Enabled = False
  Frame2.Visible = False
  BtnImprimir2.Visible = True
End Sub

Private Sub optRep012_Click()
    optRep024.Visible = True
End Sub

Private Sub optRep015_Click()
    Frame2.Visible = False
    optRep019.Visible = True
End Sub

Private Sub optRep016_Click()
    Frame2.Visible = False
End Sub

Private Sub optRep018_Click()
    Frame2.Visible = False
    optRep020.Visible = True
End Sub

Private Sub optRep021_Click()
    Frame2.Visible = False
'    optRep022.Visible = True
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Fw_ReportesVentas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "REPORTES DE VENTAS"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Fw_ReportesVentas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox fraOpciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   120
      ScaleHeight     =   630
      ScaleWidth      =   15270
      TabIndex        =   54
      Top             =   45
      Width           =   15300
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   13320
         Picture         =   "Fw_ReportesVentas.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   96
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   960
         Picture         =   "Fw_ReportesVentas.frx":11C4
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   95
         ToolTipText     =   "Kardex por Contrato Seleccionado"
         Top             =   0
         Width           =   1400
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00808000&
         Caption         =   "&Docs."
         Enabled         =   0   'False
         Height          =   720
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   740
      End
      Begin VB.CommandButton CmdFoto 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Foto"
         Height          =   720
         Left            =   2880
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
         Left            =   7020
         TabIndex        =   57
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame fmrTipoReporte 
      BackColor       =   &H00C0C0C0&
      Caption         =   "---- HISTORIAL VENTAS ACUMULADAS ------------ HISTORIAL DETALLE DE VENTAS ------------------- OTROS para Exportar a EXCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1635
      Width           =   15195
      Begin VB.OptionButton optrep005 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seguimiento VENTAS, FACTURACION y COBRANZA"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   10200
         TabIndex        =   19
         Top             =   960
         Width           =   4260
      End
      Begin VB.OptionButton optRep014 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas por UNIDAD EJECUTORA con Detalle de Bienes - para Exportar a Excel"
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
         Height          =   450
         Left            =   10200
         TabIndex        =   105
         Top             =   360
         Width           =   4155
      End
      Begin VB.OptionButton optRep022 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COBRANZAS por Gestion y Mes (con Equipos)"
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
         Height          =   325
         Left            =   10200
         TabIndex        =   104
         Top             =   2400
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.OptionButton optRep013 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COBRANZAS por Gestion y Mes (Ver.3)"
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
         Height          =   325
         Left            =   10200
         TabIndex        =   103
         Top             =   2040
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.OptionButton optRep026 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COBRANZAS por Gestion y Mes (Ver.2)"
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
         Height          =   325
         Left            =   10200
         TabIndex        =   102
         Top             =   1680
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.OptionButton optRep020 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COBRANZAS por Gestion y Mes (Ver.1)"
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
         Height          =   325
         Left            =   10200
         TabIndex        =   101
         Top             =   1320
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.OptionButton optRep0015 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por SERVICIO y REGIONAL"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5160
         TabIndex        =   100
         Top             =   3720
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.OptionButton optRep0011 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por SERVICIO y REGIONAL"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   99
         Top             =   3840
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.OptionButton optRep032 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por PROCESO ISO y GESTION"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   98
         Top             =   3000
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.OptionButton optRep031 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por PROCESOS ISO y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   97
         Top             =   2640
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.OptionButton optRep028 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por CLIENTE"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   94
         Top             =   1440
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
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
         Height          =   2175
         Left            =   4800
         TabIndex        =   93
         Top             =   4680
         Width           =   10215
      End
      Begin VB.OptionButton optRep047 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por: GESTION, MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   92
         Top             =   3375
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.OptionButton optRep040 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por: REGIONAL"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   91
         Top             =   1095
         Width           =   3420
      End
      Begin VB.OptionButton optRep046 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por UNIDAD EJECUTORA (Vigentes)"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   90
         Top             =   1830
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.OptionButton optRep042 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por: REGIONAL"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   89
         Top             =   1080
         Width           =   3540
      End
      Begin VB.OptionButton optRep043 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por CLIENTE"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   88
         Top             =   1500
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optRep041 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por REGIONAL (Vigentes)"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   87
         Top             =   2610
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optRep045 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por CLIENTE (Vigentes)"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   86
         Top             =   3000
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optRep044 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por SERVICIO (Proceso ISO) - VIGENTES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   85
         Top             =   2220
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.OptionButton opt_mcobrado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cobrado"
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
         Height          =   285
         Left            =   360
         TabIndex        =   76
         Top             =   4980
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.OptionButton opt_mfacturado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturado"
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
         Left            =   360
         TabIndex        =   75
         Top             =   4680
         Visible         =   0   'False
         Width           =   3630
      End
      Begin VB.OptionButton opt_mdevengado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Devengado"
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
         Left            =   360
         TabIndex        =   74
         Top             =   5310
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.OptionButton optRep008 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por SERVICIO (Proceso ISO) - VIGENTES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   25
         Top             =   2235
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optrep006 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seguimiento VENTAS, FACTURACION y COBRANZA (Migrar)"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   10200
         TabIndex        =   24
         Top             =   3795
         Width           =   4725
      End
      Begin VB.OptionButton optRep004 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por SERVICIO (Proceso ISO)"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   23
         Top             =   375
         Width           =   3420
      End
      Begin VB.OptionButton optRep003 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por UNIDAD EJECUTORA"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   22
         Top             =   720
         Width           =   3420
      End
      Begin VB.OptionButton optRep0010 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LIBRO DE COMPRAS"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   10080
         TabIndex        =   17
         Top             =   4440
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.OptionButton optRep009 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por GESTION Y MES"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   16
         Top             =   3405
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optRep007 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por UNIDAD EJECUTORA (Vigentes)"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   15
         Top             =   1875
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton optRep001 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por UNIDAD EJECUTORA"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   7
         Top             =   720
         Width           =   3540
      End
      Begin VB.OptionButton optRep002 
         BackColor       =   &H00C0C0C0&
         Caption         =   "por SERVICIO (Proceso ISO)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   5280
         TabIndex        =   6
         Top             =   375
         Width           =   3540
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos Para Migrar (Macont)"
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
         Height          =   1575
         Left            =   120
         TabIndex        =   73
         Top             =   4320
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         X1              =   4800
         X2              =   4800
         Y1              =   120
         Y2              =   5640
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         X1              =   9840
         X2              =   9840
         Y1              =   120
         Y2              =   5640
      End
   End
   Begin VB.Frame ConProy00 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H000000C0&
      Height          =   3720
      Left            =   120
      TabIndex        =   18
      Top             =   5970
      Visible         =   0   'False
      Width           =   11700
      Begin MSDataListLib.DataCombo DtcCbrDes 
         Height          =   315
         Left            =   4560
         TabIndex        =   37
         Top             =   1080
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcProvDes 
         Height          =   315
         Left            =   4560
         TabIndex        =   26
         Top             =   2880
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCliDes 
         Height          =   315
         Left            =   4560
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcVenDes 
         Height          =   315
         Left            =   4560
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcTipoDes 
         Height          =   315
         Left            =   4560
         TabIndex        =   40
         Top             =   3240
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_descripcion"
         BoundColumn     =   "tipo_venta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcTipoCliDes 
         Height          =   315
         Left            =   4560
         TabIndex        =   45
         Top             =   1440
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "subproceso_descripcion"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCiu 
         Height          =   315
         Left            =   4560
         TabIndex        =   46
         Top             =   1800
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "municipio"
         BoundColumn     =   "Depto"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcProd 
         Height          =   315
         Left            =   4560
         TabIndex        =   51
         Top             =   2520
         Visible         =   0   'False
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "concepto_venta"
         BoundColumn     =   "codDetalle"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcProvCod 
         Height          =   315
         Left            =   2880
         TabIndex        =   32
         Top             =   2880
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCliCod 
         Height          =   315
         Left            =   2880
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcVenCod 
         Height          =   315
         Left            =   2880
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcCbrCod 
         Height          =   315
         Left            =   2880
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc Ado_proveedor 
         Height          =   330
         Left            =   11640
         Top             =   2880
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   11640
         Top             =   360
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   11640
         Top             =   720
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   11640
         Top             =   1080
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   2880
         TabIndex        =   39
         Top             =   3240
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_venta"
         BoundColumn     =   "tipo_venta"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc Ado_Tipo 
         Height          =   330
         Left            =   11640
         Top             =   3240
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   2880
         TabIndex        =   43
         Top             =   1440
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "subproceso_codigo"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcDepto 
         Height          =   315
         Left            =   2880
         TabIndex        =   44
         Top             =   1800
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Depto"
         BoundColumn     =   "Depto"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc Ado_TipoBenef 
         Height          =   330
         Left            =   11640
         Top             =   1440
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Caption         =   "Ado_TipoBenef"
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
         Left            =   11640
         Top             =   1800
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   8520
         TabIndex        =   49
         Top             =   2160
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
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
         Left            =   2880
         TabIndex        =   50
         Top             =   2520
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ccodDetalle"
         BoundColumn     =   "codDetalle"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtcMes 
         Height          =   315
         Left            =   2880
         TabIndex        =   52
         Top             =   2160
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
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
         Left            =   11640
         Top             =   2160
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   11640
         Top             =   2520
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   53
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bien / Servicio. . . . :"
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
         Left            =   720
         TabIndex        =   48
         Top             =   2520
         Visible         =   0   'False
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento (Bolivia):"
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
         Left            =   720
         TabIndex        =   42
         Top             =   1845
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Procesos ISO  . . . . . . . ."
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
         Left            =   720
         TabIndex        =   41
         Top             =   1485
         Visible         =   0   'False
         Width           =   2055
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   3285
         Visible         =   0   'False
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   1125
         Visible         =   0   'False
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   765
         Visible         =   0   'False
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   405
         Visible         =   0   'False
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   2895
         Visible         =   0   'False
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija las Fechas, luego una Opci?n y el Bot?n Imprimir"
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
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   15180
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   3645
         TabIndex        =   1
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   110297089
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   7125
         TabIndex        =   2
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   110297089
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5880
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   405
         Width           =   1350
      End
   End
   Begin Crystal.CrystalReport CryReporte 
      Left            =   1440
      Top             =   10575
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
      Left            =   2880
      Top             =   10560
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
      Left            =   135
      Top             =   10530
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
      Left            =   2160
      Top             =   10560
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
      Left            =   840
      Top             =   10560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Par?metros de Ventas y/o Cobranzas"
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
      Height          =   5115
      Left            =   120
      TabIndex        =   63
      Top             =   7650
      Visible         =   0   'False
      Width           =   16980
      Begin VB.OptionButton optRep033 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por REGIONAL y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   84
         Top             =   1560
         Width           =   3765
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturaci?n Acumulada por REGIONAL y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   83
         Top             =   4320
         Width           =   3765
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cobranzas Acumuladas por REGIONAL y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   82
         Top             =   3960
         Width           =   3765
      End
      Begin VB.OptionButton optRep034 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por REGIONAL y GESTION"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   81
         Top             =   1920
         Width           =   4125
      End
      Begin VB.OptionButton optRep035 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por Proceso VENTAS NUEVAS y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   80
         Top             =   2280
         Width           =   4725
      End
      Begin VB.OptionButton optRep036 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por Proceso MANTENIMIENTO y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   79
         Top             =   2640
         Width           =   4725
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por Proceso REPARACIONES y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   78
         Top             =   3000
         Width           =   4725
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas Acumuladas por Proceso REPARACIONES y MES"
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   480
         TabIndex        =   77
         Top             =   3360
         Width           =   4725
      End
      Begin VB.Frame FrameConDet 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Elija Opcion:"
         ForeColor       =   &H000000C0&
         Height          =   600
         Left            =   8280
         TabIndex        =   70
         Top             =   160
         Visible         =   0   'False
         Width           =   4320
         Begin VB.OptionButton optSi 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Importes (Bolivianos)"
            Height          =   195
            Left            =   225
            TabIndex        =   72
            Top             =   255
            Width           =   1905
         End
         Begin VB.OptionButton optNo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cantidades"
            Height          =   195
            Left            =   2745
            TabIndex        =   71
            Top             =   250
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin VB.ComboBox cmb_gestion_rep 
         Height          =   315
         Left            =   5400
         TabIndex        =   69
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbo_mes_rep 
         Height          =   315
         Left            =   1320
         TabIndex        =   67
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   15120
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   915
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   68
         Top             =   405
         Width           =   990
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   64
         Top             =   405
         Width           =   630
      End
   End
   Begin VB.Frame FrameTipo 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   5640
      TabIndex        =   58
      Top             =   7680
      Visible         =   0   'False
      Width           =   4335
      Begin VB.OptionButton Opt_3 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   62
         Top             =   880
         Width           =   3855
      End
      Begin VB.OptionButton opt_4 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   300
         TabIndex        =   61
         Top             =   1200
         Value           =   -1  'True
         Width           =   2805
      End
      Begin VB.OptionButton Opt_1 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   300
         TabIndex        =   60
         Top             =   255
         Width           =   3630
      End
      Begin VB.OptionButton opt_2 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   300
         TabIndex        =   59
         Top             =   550
         Width           =   3750
      End
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
Attribute VB_Name = "Fw_ReportesVentas"
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
  dtpFecha1.Value = Format("01/01/2020", "dd/mm/yyyy")
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

  'Detalle de Ventas por UNIDAD EJECUTORA
  If optRep001.Value = True Then        'And Opt_1.Value = True
    'Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_de_ventas_unidad.rpt", "VENTAS EN GENERAL")
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_unidad.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "DETALLE VENTAS POR UNIDAD"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
  'Detalle de Ventas por SERVICIO (SUB-PROCESO ISO)
  ElseIf optRep002.Value = True Then        'And Opt_1.Value = True     '
    'Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_de_ventas_unidad.rpt", "VENTAS EN GENERAL")
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_servicio_regional.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "DETALLE VENTAS POR SERVICIO"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
  'Detalle de Ventas por REGIONAL
  ElseIf optRep042.Value = True And opt_4.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_regional.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "DETALLE VENTAS POR REGIONAL"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

    'LIBRO DE COMPRAS
  ElseIf optRep0010.Value = True Then
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
  
  'Ventas Acumuladas por UNIDAD EJECUTORA
  ElseIf optRep003.Value = True Then        'And Opt_1.Value = True
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_ventas_unidad_acum.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "VENTAS ACUMULADAS POR UNIDAD"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
  
    'Ventas Acumuladas por SERVICIO (PROCESO ISO)
  ElseIf optRep004.Value = True Then        'And Opt_1.Value = True
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_ventas_servicio_acum.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "VENTAS ACUMULADAS POR SERVICIO"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
    'Ventas Acumuladas por REGIONAL
  ElseIf optRep040.Value = True And opt_4.Value = True Then
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "TODAS LAS VENTAS Y COBRANZAS")
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_ventas_regional_acum.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "VENTAS ACUMULADAS POR PROCESO ISO"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

    ' SEGUIMIENTO DE VENTAS, FACTURACION Y COBRANZAS
  ElseIf optrep005.Value = True Then        'And Opt_1.Value = True     'optrep005
    'Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_de_ventas_unidad.rpt", "VENTAS EN GENERAL")
    
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_seguimiento_venta_fac_cobro.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "SEGUIMIENTO VENTAS, FACTURACION Y COBRANZAS"
        'CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        'CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
       
    ' SEGUIMIENTO DE VENTAS, FACTURACION Y COBRANZAS
  ElseIf optrep006.Value = True Then        'And Opt_1.Value = True     'optrep005
    'Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_de_ventas_unidad.rpt", "VENTAS EN GENERAL")
    
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_seguimiento_venta_fac_cobro_txt.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "SEGUIMIENTO VENTAS, FACTURACION Y COBRANZAS"
        'CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        'CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
 'VENTAS SIN FACTURAR
  ElseIf optRep007.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_seguimiento_venta_sin_fac.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "VENTAS PENDIENTES DE FACTURACION"
        'CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        'CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = "%"           'cmb_gestion_rep.Text
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
'    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_de_ventas_cobador.rpt", "VENTAS Y COBRANZAS")
'    CryUnidad.ReportFileName = App.Path & "\Reportes\Comercial\ar_bienes_equipos_OTIS.rpt"
''        titulo2 = "MODULO VENTAS"
''        subtitulo2 = "VENTAS POR COBRADOR"
''        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
''        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
    
    'SEGUIMIENTO DE VENTAS POR PRODUCTO
'  ElseIf optRep007.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
'  ElseIf optRep007.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
'  ElseIf optRep007.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "OBSEQUIO, DONACION, DEGUSTACION (MONTO CERO)")
'  ElseIf optRep007.Value = True And opt_4.Value = True Then
'    Call RepUnidad("TODAS", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "TODAS LAS VENTAS Y COBRANZAS")

        
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

'  'LISTADO GENERAL DE COBRANZAS
'  ElseIf optRep010.Value = True And opt_4.Value = True Then
'    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_dol.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'    CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_unidad.rpt"
'    titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS POR UNIDAD"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
    
'  ElseIf optRep023.Value = True And opt_4.Value = True Then
'    'db.Execute ""
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_unidad_dol.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS POR UNIDAD (Dolares)"
''        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
''        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
''        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
''        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
    
'  'LIBRO DE VENTAS
'  ElseIf optRep011.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep011.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep011.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_libro_ventas.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep011.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_libro_ventas.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'  'End If
  
'  'COBRANZAS POR FACTURA
'  ElseIf optRep012.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\fr_cobranzas_solo_facturadas.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep012.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\fr_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep012.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\fr_cobranzas_solo_facturadas.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep012.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_solo_facturadas.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
'  'COBRANZAS POR FACTURA - por FECHAS
'  ElseIf optRep024.Value = True And opt_4.Value = True Then
'    'db.Execute ""
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturas_vs_cobros.rpt"
'        titulo2 = "FACTURACION POR DEPARTAMENTO"
'        subtitulo2 = "MODULO COBRANZAS"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
'  'COBRANZAS POR COBRADOR
'  ElseIf optRep015.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep015.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep015.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep015.Value = True And opt_4.Value = True Then
'    'db.Execute ""
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS POR COBRADOR"
''        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
''        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
''        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
''        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
'COBRANZAS POR COBRADOR (ACUMULADO)
  ElseIf optRep044.Value = True And Opt_1.Value = True Then
    Call RepUnidad("CONSALDO", "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_Cobrador_Unidad_Depto.rpt", "VENTAS NO FACTURADAS")
  ElseIf optRep044.Value = True And opt_2.Value = True Then
    Call RepUnidad("SINSALDO", "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_Cobrador_Unidad_Depto.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
  ElseIf optRep044.Value = True And Opt_3.Value = True Then
    Call RepUnidad("MONTOCERO", "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_Cobrador_Unidad_Depto.rpt", "VENTAS FACTURADAS Y COBRADAS")
  ElseIf optRep044.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_Cobrador_Unidad_Depto.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR COBRADOR"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
' 'COBRANZAS POR COBRADOR en MORA
''  ElseIf optRep015.Value = True And Opt_1.Value = True Then
''    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS NO FACTURADAS")
''  ElseIf optRep015.Value = True And opt_2.Value = True Then
''    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
''  ElseIf optRep015.Value = True And Opt_3.Value = True Then
''    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobrador.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep019.Value = True And opt_4.Value = True Then
'    'db.Execute ""
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_facturadas_Cobr_Mora.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS EN MORA POR COBRADOR"
''        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
''        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
''        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
''        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If

''COBRANZAS POR MES
'  ElseIf optRep018.Value = True And opt_4.Value = True Then
'    'db.Execute ""
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS POR MES"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
''        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
''        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
'COBRANZAS POR MES (TOTALES)
  ElseIf optRep047.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_Gestion_mes.rpt"
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
        
  'COBRANZAS POR MES (PARA MIGRAR VER 1 - C/GLOSA)
  ElseIf optRep020.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes_txt.rpt"
        'CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_seguimiento_cobranza_kardex_mc.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "CUENTAS POR COBRAR Bs."
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
'        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

'    'COBRANZAS POR CONTRATO - CLIENTE
'  ElseIf optRep027.Value = True And opt_4.Value = True Then
'    'db.Execute ""
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_seguimiento_cobranza_kardex_mc.rpt"
'
'        titulo2 = "MODULO VENTAS"
'        subtitulo2 = "CONTRATOS POR CLIENTE"
''        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
''        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
''        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
''        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'        'CryUnidad.StoredProcParam(0) = "2017"   'Format(dtpFecha1.Value, "dd/mm/yyyy")
'
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
      'COBRANZAS POR CONTRATO - CLIENTE (TOTALES)
  ElseIf optRep043.Value = True And opt_4.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_Depto_Cliente.rpt"
        titulo2 = "MODULO VENTAS"
        subtitulo2 = "CONTRATOS POR CLIENTE"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

  'COBRANZAS POR MES (PARA MIGRAR VER 2 - C/COBRADOR Y FACTURADO A:)
  ElseIf optRep026.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes_lm.rpt"
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
        
  'COBRANZAS POR GESTIO Y MES (PARA MIGRAR VER 3 - C/COBRADOR Y FACTURADO A:)
  ElseIf optRep013.Value = True And opt_4.Value = True Then
    'db.Execute ""
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_gestion_mes_v3.rpt"
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

''COBRANZAS POR MES con EQUIPOS
'  ElseIf optRep021.Value = True And opt_4.Value = True Then
'    'db.Execute ""
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_cobranzas_facturadas_mes_eqp.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS P/MES C/EQUIPOS"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
''        CryUnidad.StoredProcParam(0) = Format(DTPfecha1.Value, "dd/mm/yyyy")
''        CryUnidad.StoredProcParam(1) = Format(dtpFecha2.Value, "dd/mm/yyyy")
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
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
        
'        'COBRANZAS POR REGIONAL, COBRADOR
''  ElseIf optRep016.Value = True And Opt_1.Value = True Then
''    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS NO FACTURADAS")
''  ElseIf optRep016.Value = True And opt_2.Value = True Then
''    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
''  ElseIf optRep016.Value = True And Opt_3.Value = True Then
''    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt", "VENTAS FACTURADAS Y COBRADAS")
'    ElseIf optRep016.Value = True And opt_4.Value = True Then
'        'CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt"
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_recibo_mes.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS POR COBRADOR"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
        'COBRANZAS POR RECIBO
    ElseIf optRep014.Value = True And opt_4.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_de_ventas_BIENES_txt.rpt"
        titulo2 = "CONTRATOS DE VENTAS "
        subtitulo2 = "por UNIDAD EJECUTORA"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    
'        'COBRANZAS POR REGIONAL, TESORERO, COBRADOR
'    ElseIf optRep017.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_tesoreria_regional_detalle.rpt"
'        titulo2 = "MODULO TESORERIA"
'        subtitulo2 = "SEGUIMIENTO POR REGIONAL"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
        'COBRANZAS POR REGIONAL, TESORERO, COBRADOR
    ElseIf optRep028.Value = True And opt_4.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_tesoreria_regional_seg.rpt"
        titulo2 = "MODULO TESORERIA"
        subtitulo2 = "SEGUIMIENTO POR REGIONAL"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
        'COBRANZAS POR RECIBO (TOTALES)
    ElseIf optRep045.Value = True And opt_4.Value = True Then
        'CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt"
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_por_depto_y_cobrador.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS POR COBRADOR"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
'        'COBRANZAS PARA CONCILIACION
'    ElseIf optRep025.Value = True And opt_4.Value = True Then
'        'CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt"
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_seguimiento_cobranza_concilia.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "COBRANZAS PARA CONCILIACION"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
        
        'COBRANZAS PARA CONCILIACION (TOTALES)
    ElseIf optRep046.Value = True And opt_4.Value = True Then
        'CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_lista_cobranzas_solo_recibo.rpt"
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\Gerenciales\gr_cobranzas_CuentaBancaria.rpt"
        titulo2 = "MODULO COBRANZAS"
        subtitulo2 = "COBRANZAS PARA CONCILIACION"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")

'----- REPORTES GERENCIALES
    'Ventas Acumuladas por PROCESOS ISO y MES (Bs)
  ElseIf optRep031.Value = True And optSi.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes_Bs.rpt"
        titulo2 = "PROCESOS ISO VS. MESES"
        subtitulo2 = "VENTAS ACUMULADAS EN Bs."
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
        CryUnidad.StoredProcParam(1) = 1
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    'Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
    'Ventas Acumuladas por PROCESOS ISO y MES (Cantidad de Contratos)
  ElseIf optRep031.Value = True And optNo.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes.rpt"
        titulo2 = "PROCESOS ISO VS. MESES"
        subtitulo2 = "VENTAS ACUMULADAS p/Contrato"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
        CryUnidad.StoredProcParam(1) = 1
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    'Call RepUnidad("SINSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS SIN SALDO DEUDOR (CANCELADAS)")
 
      'Ventas Acumuladas por DEPARTAMENTOS y MES (Bs)
  ElseIf optRep033.Value = True And optSi.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_depto_y_mes_Bs.rpt"
        titulo2 = "DEPARTAMENTOS VS. MESES"
        subtitulo2 = "VENTAS ACUMULADAS EN Bs."
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
        CryUnidad.StoredProcParam(1) = 1
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    'Call RepUnidad("CONSALDO", "\Reportes\Ventas\VENTAS_PRODUCTO.rpt", "VENTAS CON SALDO DEUDOR")
    'Ventas Acumuladas por DEPARTAMENTOS y MES (Cantidad Contratos)
  ElseIf optRep033.Value = True And optNo.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_depto_y_mes.rpt"
        titulo2 = "DEPARTAMENTOS VS. MESES"
        subtitulo2 = "VENTAS ACUMULADAS p/Contrato"
        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
        CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
        CryUnidad.StoredProcParam(1) = 1
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
   ' Macont
        'Facturado
     ElseIf opt_mfacturado.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\mr_facturado.rpt"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
        'Cobrado
    ElseIf opt_mcobrado.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\mr_cobrado.rpt"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        
        'Devengado
    ElseIf opt_mdevengado.Value = True Then
        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\mr_devengado.rpt"
        iResult = CryUnidad.PrintReport
        If iResult <> 0 Then
            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
        

  
  
  End If

End Sub

Private Sub BtnImprimir2_Click()
'    'LIBRO DE VENTAS
'  If optRep011.Value = True And Opt_1.Value = True Then
'    Call RepUnidad("CONSALDO", "\Reportes\Ventas\ar_libro_ventas_txt.rpt", "VENTAS NO FACTURADAS")
'  ElseIf optRep011.Value = True And opt_2.Value = True Then
'    Call RepUnidad("SINSALDO", "\Reportes\Ventas\ar_libro_ventas_txt.rpt", "VENTAS FACTURADAS Y NO COBRADAS")
'  ElseIf optRep011.Value = True And Opt_3.Value = True Then
'    Call RepUnidad("MONTOCERO", "\Reportes\Ventas\ar_libro_ventas_txt.rpt", "VENTAS FACTURADAS Y COBRADAS")
'  ElseIf optRep011.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\ar_libro_ventas_txt.rpt"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'        End If
'    'Call RepUnidad("TODAS", "\Reportes\Ventas\ar_libro_ventas.rpt", "TODAS LAS VENTAS Y COBRANZAS")
'  End If
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
    'Ado_proveedor.Refresh

    Set rs_cliente = New ADODB.Recordset
    If rs_cliente.State = 1 Then rs_cliente.Close
    rs_cliente.Open "select * from gc_beneficiario WHERE (tipoben_codigo <> 1 AND tipoben_codigo <> 23) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    'rs_cliente.Open "select * from gc_beneficiario WHERE (tipoben_codigo <> 2 AND tipoben_codigo <> 22)  ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    Set ado_Cliente.Recordset = rs_cliente
    'ado_Cliente.Refresh

    Set rs_vendedor = New ADODB.Recordset
    If rs_vendedor.State = 1 Then rs_vendedor.Close
    'rs_vendedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=6 OR tipoben_codigo=10) and (beneficiario_deudor = 'SI') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    rs_vendedor.Open "select * from gc_beneficiario WHERE (tipoben_codigo=1 OR tipoben_codigo=0) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_vendedor.Recordset = rs_vendedor
    'Ado_vendedor.Refresh

    Set rs_cobrador = New ADODB.Recordset
    If rs_cobrador.State = 1 Then rs_cobrador.Close
    'rs_cobrador.Open "select * from gc_beneficiario WHERE (tipoben_codigo=7 OR tipoben_codigo=10) and (beneficiario_deudor = 'SI') ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    rs_cobrador.Open "select * from gc_beneficiario WHERE (tipoben_codigo=1 OR tipoben_codigo=0) ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Cobrador.Recordset = rs_cobrador
    'Ado_Cobrador.Refresh

    Set rs_tipo = New ADODB.Recordset
    If rs_tipo.State = 1 Then rs_tipo.Close
    rs_tipo.Open "select venta_tipo, venta_tipo_descripcion from ac_tipo_compra_venta WHERE estado_codigo='APR' ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Tipo.Recordset = rs_tipo
    'Ado_Tipo.Refresh

    Set rs_tipoBenef = New ADODB.Recordset
    If rs_tipoBenef.State = 1 Then rs_tipoBenef.Close
    rs_tipoBenef.Open "select subproceso_codigo, subproceso_descripcion from gc_proceso_nivel2 WHERE (estado_codigo='APR' and subproceso_parametro_menor=1) ", db, adOpenKeyset, adLockReadOnly
    Set Ado_TipoBenef.Recordset = rs_tipoBenef
    'Ado_TipoBenef.Refresh

    Set rs_ciudad = New ADODB.Recordset
    If rs_ciudad.State = 1 Then rs_ciudad.Close
    'rs_ciudad.Open "select Depto AS procedencia, municipio AS lugar_procedencia from gc_beneficiario WHERE (tipoben_codigo<>'B' and tipoben_codigo<>'O' and tipoben_codigo<>'P') and (activo = 'S') group BY Depto, municipio ", DB, adOpenKeyset, adLockReadOnly
    rs_ciudad.Open "select Depto_codigo , munic_codigo from gc_beneficiario WHERE (tipoben_codigo <>0 ) and (beneficiario_deudor = 'SI' OR beneficiario_deudor = 'NO') group BY Depto_codigo, munic_codigo ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Ciudad.Recordset = rs_ciudad
    'Ado_Ciudad.Refresh
    
'    Set rs_meses = New ADODB.Recordset
'    If rs_meses.State = 1 Then rs_meses.Close
'    rs_meses.Open "select * from gc_periodos WHERE (estado_registro = 'S') ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_Meses.Recordset = rs_meses
'    Ado_Meses.Refresh
    
    Set rs_producto = New ADODB.Recordset
    If rs_producto.State = 1 Then rs_producto.Close
    rs_producto.Open "select bien_codigo, concepto_venta from ao_ventas_detalle group BY bien_codigo, concepto_venta ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Producto.Recordset = rs_producto
    'Ado_Producto.Refresh
    
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
'    BtnImprimir2.Visible = False
'    Call llena_datos
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
'    optRep023.Visible = True
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
'  BtnImprimir2.Visible = True
End Sub

Private Sub optRep012_Click()
'    optRep024.Visible = True
End Sub

Private Sub optRep015_Click()
    Frame2.Visible = False
'    optRep019.Visible = True
End Sub

Private Sub optRep016_Click()
    Frame2.Visible = False
End Sub

Private Sub optRep018_Click()
    Frame2.Visible = False
    optRep020.Visible = True
    optRep026.Visible = True
End Sub

Private Sub optRep021_Click()
    Frame2.Visible = False
'    optRep022.Visible = True
End Sub

Private Sub optRep031_Click()
    'ConProy00.Visible = True
    Label6.Visible = True
    DtcTipoCli.Visible = True
    DtcTipoCliDes.Visible = True
    Frame2.Visible = True
    FrameConDet.Visible = True
End Sub


Private Sub optRep033_Click()
    'ConProy00.Visible = True
    Label6.Visible = True
    DtcTipoCli.Visible = True
    DtcTipoCliDes.Visible = True
    Frame2.Visible = True
    FrameConDet.Visible = True
End Sub


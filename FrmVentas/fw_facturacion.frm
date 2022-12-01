VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_facturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financiero - Facturación y Cobranzas -  Facturación"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   17115
   Icon            =   "fw_facturacion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   2.3891e6
   ScaleMode       =   0  'User
   ScaleWidth      =   3.82152e7
   WindowState     =   2  'Maximized
   Begin VB.Frame frm_benef 
      BackColor       =   &H00404040&
      Caption         =   "Elije un Nuevo Beneficiario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2535
      Left            =   3600
      TabIndex        =   45
      Top             =   9240
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CommandButton BtnGrabarBen 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   3240
         Picture         =   "fw_facturacion.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1440
         Width           =   1365
      End
      Begin VB.CommandButton BtnCancelarBen 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   5280
         MaskColor       =   &H00000000&
         Picture         =   "fw_facturacion.frx":11F0
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Cancelar"
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9000
         TabIndex        =   46
         Top             =   735
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "fw_facturacion.frx":1ADC
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   6960
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "00000000000004"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "fw_facturacion.frx":1AF5
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   720
         TabIndex        =   50
         Top             =   720
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux8 
         Bindings        =   "fw_facturacion.frx":1B0E
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   6960
         TabIndex        =   51
         Top             =   720
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "beneficiario_nit"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "00000000000004"
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "NIT"
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
         Left            =   7080
         TabIndex        =   53
         Top             =   480
         Width           =   2025
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Factura a Nombre de:"
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
         Left            =   720
         TabIndex        =   52
         Top             =   465
         Width           =   2025
      End
   End
   Begin VB.Frame FrmCobros 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5700
      Left            =   2640
      TabIndex        =   13
      Top             =   480
      Width           =   12015
      Begin VB.TextBox TxtAutorizacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "dosifica_autorizacion"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos1"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   2280
         TabIndex        =   63
         Text            =   "0"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.PictureBox FraGrabarCancelar 
         BackColor       =   &H00404040&
         FillColor       =   &H00FFFFFF&
         Height          =   900
         Left            =   40
         ScaleHeight     =   840
         ScaleWidth      =   11880
         TabIndex        =   59
         Top             =   4680
         Width           =   11940
         Begin VB.CommandButton BtnImprimir3 
            BackColor       =   &H80000018&
            Caption         =   "Emitir.Factura"
            Height          =   635
            Left            =   4320
            Picture         =   "fw_facturacion.frx":1B27
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Emite e Imprime Factura"
            Top             =   120
            Width           =   1365
         End
         Begin VB.CommandButton BtnCancelar 
            BackColor       =   &H00E0E0E0&
            Height          =   635
            Left            =   6120
            MaskColor       =   &H00000000&
            Picture         =   "fw_facturacion.frx":2529
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   120
            Width           =   1365
         End
         Begin VB.CommandButton BtnGrabar 
            BackColor       =   &H00E0E0E0&
            Height          =   635
            Left            =   4320
            Picture         =   "fw_facturacion.frx":2F03
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   120
            Visible         =   0   'False
            Width           =   1365
         End
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   420
         Left            =   5040
         ScaleHeight     =   375
         ScaleMode       =   0  'User
         ScaleWidth      =   352.941
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton cmd_benef 
         BackColor       =   &H00E0E0E0&
         Height          =   555
         Left            =   10840
         Picture         =   "fw_facturacion.frx":37C3
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Buscar Beneficiario"
         Top             =   3150
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtObs 
         CausesValidation=   0   'False
         DataField       =   "glosa_Descripcion"
         DataSource      =   "Ado_datos1"
         Height          =   585
         Left            =   1440
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1740
         Width           =   9975
      End
      Begin VB.TextBox TxtMontoDol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "total_dol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos1"
         Height          =   285
         Left            =   10040
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   4080
         Width           =   1395
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   10380
         TabIndex        =   19
         Top             =   3645
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   11160
         TabIndex        =   17
         Top             =   2670
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txt_tdc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "cambio_oficial"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos1"
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Text            =   "6.96"
         Top             =   4080
         Width           =   915
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "total_bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos1"
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
         Left            =   5445
         TabIndex        =   15
         Text            =   "0"
         Top             =   4080
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpbte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "nro_factura"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   420
         Left            =   9960
         TabIndex        =   14
         Text            =   "0"
         Top             =   1040
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dtc_codigo4A 
         Bindings        =   "fw_facturacion.frx":3C05
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9720
         TabIndex        =   18
         Top             =   2655
         Visible         =   0   'False
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "12345678901234"
      End
      Begin MSDataListLib.DataCombo dtc_desc5xx 
         Bindings        =   "fw_facturacion.frx":3C1E
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos1"
         Height          =   315
         Left            =   2325
         TabIndex        =   23
         Top             =   3630
         Visible         =   0   'False
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux5 
         Bindings        =   "fw_facturacion.frx":3C37
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos1"
         Height          =   315
         Left            =   8355
         TabIndex        =   24
         Top             =   3630
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "beneficiario_nit"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "00000000000004"
      End
      Begin MSDataListLib.DataCombo dtc_desc4A 
         Bindings        =   "fw_facturacion.frx":3C50
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5295
         TabIndex        =   25
         Top             =   2655
         Visible         =   0   'False
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo lbl_nit 
         Bindings        =   "fw_facturacion.frx":3C69
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos1"
         Height          =   315
         Left            =   8400
         TabIndex        =   26
         Top             =   3000
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "00000000000004"
      End
      Begin MSComCtl2.DTPicker DTPFechaCobro 
         DataField       =   "fecha_fac"
         DataSource      =   "Ado_datos1"
         Height          =   300
         Left            =   2400
         TabIndex        =   27
         Top             =   2655
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CheckBox        =   -1  'True
         Format          =   274595841
         CurrentDate     =   44699
      End
      Begin VB.Label dtc_desc5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "beneficiario_RazonSocial"
         DataSource      =   "Ado_datos1"
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
         Height          =   300
         Left            =   2400
         TabIndex        =   76
         Top             =   3240
         Width           =   5325
      End
      Begin VB.Label lbl_doc1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "beneficiario_nit"
         DataSource      =   "Ado_datos1"
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
         Height          =   300
         Left            =   8355
         TabIndex        =   31
         Top             =   3300
         Width           =   2325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "NIT"
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
         Height          =   240
         Left            =   7920
         TabIndex        =   62
         Top             =   3285
         Width           =   330
      End
      Begin VB.Label lbl_factura 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   $"fw_facturacion.frx":3C82
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
         TabIndex        =   35
         Top             =   1080
         Width           =   9510
      End
      Begin VB.Label TxtMonto2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "cobranza_deuda_bs2"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
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
         Height          =   300
         Left            =   7725
         TabIndex        =   44
         Top             =   4515
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label dtc_codigo5 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "beneficiario_codigo_fac"
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos1"
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
         Height          =   300
         Left            =   6000
         TabIndex        =   43
         Top             =   3000
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label DTPFechaProg 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "cobranza_fecha_sol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Ado_datos"
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
         Height          =   300
         Left            =   9660
         TabIndex        =   42
         Top             =   2280
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label TxtNroVentaC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "venta_codigo"
         DataSource      =   "Ado_datos1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6120
         TabIndex        =   41
         Top             =   375
         Width           =   1245
      End
      Begin VB.Label lbl_fac 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "doc_codigo_fac"
         DataSource      =   "Ado_datos1"
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
         Height          =   300
         Left            =   7260
         TabIndex        =   40
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "edif_codigo_corto"
         DataSource      =   "Ado_datos1"
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
         Height          =   300
         Left            =   10200
         TabIndex        =   39
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label lbl_obs 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto:"
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
         Height          =   240
         Left            =   360
         TabIndex        =   38
         Top             =   1875
         Width           =   960
      End
      Begin VB.Label Lbl_Cobrador 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrador:"
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
         Height          =   240
         Left            =   4320
         TabIndex        =   37
         Top             =   2640
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lbl_fechas 
         Alignment       =   2  'Center
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Facturación"
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
         Height          =   240
         Left            =   120
         TabIndex        =   36
         Top             =   2640
         Width           =   2385
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFF80&
         X1              =   0
         X2              =   12000
         Y1              =   1605
         Y2              =   1560
      End
      Begin VB.Label Lbl_nombre_fac 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Factura a Nombre de:"
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
         Height          =   240
         Left            =   360
         TabIndex        =   34
         Top             =   3285
         Width           =   1950
      End
      Begin VB.Label Txt_cod_cobro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "IdFactura"
         DataSource      =   "Ado_datos1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1995
         TabIndex        =   33
         Top             =   375
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   $"fw_facturacion.frx":3D23
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
         Height          =   240
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Width           =   9810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   $"fw_facturacion.frx":3DAC
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
         Height          =   240
         Left            =   360
         TabIndex        =   30
         Top             =   4080
         Width           =   9615
      End
      Begin VB.Label TxtDscto 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "cobranza_deuda_dol2"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
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
         Height          =   300
         Left            =   5805
         TabIndex        =   29
         Top             =   2295
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFF80&
         X1              =   0
         X2              =   12000
         Y1              =   780
         Y2              =   780
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00C00000&
      Height          =   5835
      Left            =   120
      TabIndex        =   54
      Top             =   480
      Width           =   16545
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000018&
         Caption         =   "Solicitudes Anteriores"
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
         Left            =   5160
         TabIndex        =   75
         Top             =   4035
         Width           =   2175
      End
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H00404040&
         Height          =   900
         Left            =   120
         ScaleHeight     =   840
         ScaleWidth      =   16275
         TabIndex        =   66
         Top             =   0
         Width           =   16335
         Begin VB.CommandButton BtnDesAprobar 
            BackColor       =   &H80000018&
            Caption         =   "Devolver."
            Height          =   720
            Left            =   2760
            Picture         =   "fw_facturacion.frx":3E57
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Devuelve a Cobradores (Solicitud de Facturación)"
            Top             =   60
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton BtnAñadir 
            BackColor       =   &H80000018&
            Caption         =   "Facturas ANTERIORES"
            Height          =   720
            Left            =   8640
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Facturas ANTERIORES"
            Top             =   60
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "Iniciar Facturación"
            Height          =   720
            Left            =   1440
            Picture         =   "fw_facturacion.frx":4299
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Emitir Factura con los Registros Elegidos..."
            Top             =   60
            Width           =   1245
         End
         Begin VB.CommandButton BtnBuscar 
            BackColor       =   &H80000018&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   120
            Picture         =   "fw_facturacion.frx":4823
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Busca Registro para Facturar"
            Top             =   60
            Width           =   1245
         End
         Begin VB.CommandButton CmdFoto 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Reportes"
            Height          =   720
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Carga Imagen QR"
            Top             =   60
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton BtnSalir 
            BackColor       =   &H80000018&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   15000
            Picture         =   "fw_facturacion.frx":4B2D
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Cerrar Ventana"
            Top             =   60
            Width           =   1005
         End
         Begin VB.CommandButton BtnEliminar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "Anular"
            Height          =   720
            Left            =   7545
            Picture         =   "fw_facturacion.frx":2E73F
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Anula Factura Emitida"
            Top             =   60
            Width           =   1120
         End
         Begin VB.CommandButton BtnImprimir5 
            BackColor       =   &H80000018&
            Caption         =   "Re-Imprimir"
            Height          =   720
            Left            =   4080
            Picture         =   "fw_facturacion.frx":2F141
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Re-Imprime Factura"
            Top             =   60
            Width           =   1125
         End
         Begin VB.CommandButton BtnImprimir6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Re-Imprime PDF"
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
            Left            =   5160
            Picture         =   "fw_facturacion.frx":2F6CB
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Re-Imprime Factura"
            Top             =   60
            Width           =   1125
         End
         Begin VB.CommandButton BtnModDetalle2 
            BackColor       =   &H80000018&
            Caption         =   "Facturas Emitidas"
            Height          =   720
            Left            =   6255
            Picture         =   "fw_facturacion.frx":31193
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Ver Detalle del Bien ..."
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lbl_titulo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FACTURACION"
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
            Left            =   11040
            TabIndex        =   73
            Top             =   255
            Visible         =   0   'False
            Width           =   2265
         End
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "fw_facturacion.frx":315D5
         Height          =   1380
         Left            =   120
         TabIndex        =   65
         Top             =   4320
         Width           =   16365
         _ExtentX        =   28866
         _ExtentY        =   2434
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SOLICITUDES DE FACTURACION (DETALLE)"
         ColumnCount     =   17
         BeginProperty Column00 
            DataField       =   "cobranza_fecha_sol"
            Caption         =   "F.Solicit.Fac"
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
            DataField       =   "cobranza_prog_codigo"
            Caption         =   "Nro.Cuota"
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
            DataField       =   "cobranza_fecha_prog"
            Caption         =   "Mes.Cuota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "YYYY-MMMM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cobranza_codigo"
            Caption         =   "No.Cobranza"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Cobrador"
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
            DataField       =   "cobranza_total_bs"
            Caption         =   "Solicitado.Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "edif_codigo_corto"
            Caption         =   "Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre del Edificio"
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
         BeginProperty Column08 
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto de la Cuota"
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
            DataField       =   "estado_codigo_fac"
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
         BeginProperty Column10 
            DataField       =   "cobranza_fecha_fac"
            Caption         =   "Fecha.Factura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "cobranza_total_dol"
            Caption         =   "Cobrado en Dol."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc.Respaldo"
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
         BeginProperty Column13 
            DataField       =   "cobranza_nro_factura"
            Caption         =   "Nro. Factura"
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
         BeginProperty Column14 
            DataField       =   "cobranza_nro_factura"
            Caption         =   "#Factura"
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
         BeginProperty Column15 
            DataField       =   "beneficiario_codigo"
            Caption         =   "NIT/CI del Cliente"
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
         BeginProperty Column16 
            DataField       =   ""
            Caption         =   ""
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
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3750.236
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   6570.142
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_datos1 
         Bindings        =   "fw_facturacion.frx":315ED
         Height          =   2700
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   16365
         _ExtentX        =   28866
         _ExtentY        =   4763
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "FACTURAS EMITIDAS"
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "IdFactura"
            Caption         =   "Id"
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
            DataField       =   "dosifica_autorizacion"
            Caption         =   "#.Autorizacion"
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
            DataField       =   "nro_factura"
            Caption         =   "#Factura"
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
            DataField       =   "fecha_fac"
            Caption         =   "Fecha.Factura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "total_bs"
            Caption         =   "Facturado.Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "beneficiario_codigo_fac"
            Caption         =   "Factura a Nombre de:"
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
            DataField       =   "beneficiario_nit"
            Caption         =   "NIT/CI.Cliente"
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
            DataField       =   "edif_codigo_corto"
            Caption         =   "Edificio"
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
         BeginProperty Column08 
            DataField       =   "estado_codigo_fac"
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
         BeginProperty Column09 
            DataField       =   "beneficiario_RazonSocial"
            Caption         =   "Nombre del Cliente"
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
            DataField       =   "glosa_Descripcion"
            Caption         =   "Concepto de la Factura"
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
         BeginProperty Column11 
            DataField       =   "total_dol"
            Caption         =   "Facurado Dol."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "doc_numero"
            Caption         =   "#.OrdenCobro"
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
         BeginProperty Column13 
            DataField       =   "doc_codigo_fac"
            Caption         =   "Fact/Recibo"
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
         BeginProperty Column14 
            DataField       =   "beneficiario_email"
            Caption         =   "Correo Electónico"
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
         BeginProperty Column15 
            DataField       =   "JustificaAnulacionFac"
            Caption         =   "Justificacion.de.la.Anulacion"
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
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   3764.977
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1980.284
            EndProperty
            BeginProperty Column15 
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H80000018&
         Caption         =   "Solicitudes de Hoy"
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
         Left            =   2040
         TabIndex        =   57
         Top             =   4035
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H80000018&
         Caption         =   "Facturados y No Cobrados"
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
         Left            =   9000
         TabIndex        =   56
         Top             =   4035
         Width           =   2475
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000018&
         Caption         =   "Facturados y Cobrados"
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
         Left            =   13440
         TabIndex        =   55
         Top             =   4035
         Width           =   2235
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5160
         Width           =   8940
         _ExtentX        =   15769
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
      Begin MSDataGridLib.DataGrid dg_datos2 
         Bindings        =   "fw_facturacion.frx":31606
         Height          =   1380
         Left            =   1440
         TabIndex        =   64
         Top             =   4380
         Visible         =   0   'False
         Width           =   15060
         _ExtentX        =   26564
         _ExtentY        =   2434
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "REGISTROS ELEGIDOS PARA FACTURAR"
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "cobranza_fecha_sol"
            Caption         =   "F.Solicit.Fac"
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
            DataField       =   "cobranza_codigo"
            Caption         =   "No.Cobranza"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Cobrador"
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
            DataField       =   "edif_codigo_corto"
            Caption         =   "Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre del Edificio"
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
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto de la Cuota"
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
            DataField       =   "cobranza_fecha_fac"
            Caption         =   "Fecha.Factura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "cobranza_total_bs"
            Caption         =   "Facturado.Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "cobranza_total_dol"
            Caption         =   "Cobrado en Dol."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc.Respaldo"
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
            DataField       =   "cobranza_nro_factura"
            Caption         =   "Nro. Factura"
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
         BeginProperty Column11 
            DataField       =   "cobranza_nro_factura"
            Caption         =   "#Factura"
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
         BeginProperty Column12 
            DataField       =   "estado_codigo_fac"
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
         BeginProperty Column13 
            DataField       =   "beneficiario_codigo"
            Caption         =   "NIT/CI del Cliente"
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
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos1 
         Height          =   330
         Left            =   120
         Top             =   3960
         Width           =   16335
         _ExtentX        =   28813
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
      Begin MSAdodcLib.Adodc Ado_datos2 
         Height          =   330
         Left            =   120
         Top             =   5400
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "Ado_datos2"
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
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00404040&
      FillColor       =   &H00FFFFFF&
      Height          =   3075
      Left            =   120
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   5
      Top             =   6405
      Width           =   1935
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Kardex.Dol."
         Height          =   750
         Left            =   480
         Picture         =   "fw_facturacion.frx":3161F
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   1965
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnImprimir4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cronograma"
         Height          =   750
         Left            =   480
         Picture         =   "fw_facturacion.frx":31BA9
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Imprime Cronograma de Cobranzas ..."
         Top             =   195
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Kardex Bs."
         Height          =   750
         Left            =   480
         Picture         =   "fw_facturacion.frx":32133
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   45
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   688
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "FACTURACION CGI SOFIA"
      TabPicture(0)   =   "fw_facturacion.frx":326BD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "FACTURACION ON LINE CGE"
      TabPicture(1)   =   "fw_facturacion.frx":326D9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "FACTURACION ONLINE CGI"
      TabPicture(2)   =   "fw_facturacion.frx":326F5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DATOS DE LA VENTA (Para consultar)"
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
      Height          =   1335
      Left            =   2160
      TabIndex        =   4
      Top             =   6345
      Width           =   14535
      Begin MSDataGridLib.DataGrid dg_datos16 
         Bindings        =   "fw_facturacion.frx":32711
         Height          =   930
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   1640
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483644
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
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Tramite"
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
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Denominacion del Edificio"
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
            DataField       =   "zona_denominacion"
            Caption         =   "Zona"
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
            DataField       =   "calle_tipo"
            Caption         =   "Via.Acceso"
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
            DataField       =   "calle_denominacion"
            Caption         =   "Nombre de Calle, Av u otro"
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
            DataField       =   "edif_nro"
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
         BeginProperty Column07 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
         BeginProperty Column08 
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Cliente/Representante.Legal"
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
            DataField       =   "venta_fecha_inicio"
            Caption         =   "F.Inicio.Contrato"
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
            DataField       =   "venta_fecha_fin"
            Caption         =   "F.Fin.Contrato"
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
            DataField       =   "venta_cantidad_total"
            Caption         =   "Cantidad"
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
            DataField       =   "unimed_codigo"
            Caption         =   "Periodicidad"
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
         BeginProperty Column13 
            DataField       =   "venta_monto_total_bs"
            Caption         =   "Total,Contrato.Bs"
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
         BeginProperty Column14 
            DataField       =   "venta_monto_cobrado_bs"
            Caption         =   "Cobrado.Bs"
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
         BeginProperty Column15 
            DataField       =   "venta_saldo_p_cobrar_bs"
            Caption         =   "Saldo.P/Cobar"
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
         BeginProperty Column16 
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad.E."
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
         BeginProperty Column17 
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Tramite"
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
         BeginProperty Column18 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3284.788
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2234.835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2129.953
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
            BeginProperty Column15 
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
            EndProperty
            BeginProperty Column18 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE BIENES / SERVICIOS VENDIDOS"
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
      Height          =   1725
      Left            =   2160
      TabIndex        =   3
      Top             =   7725
      Width           =   14535
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "fw_facturacion.frx":3272B
         Height          =   1380
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   2434
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Bien"
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
            DataField       =   "concepto_venta"
            Caption         =   "Descripcion y Características del Bien"
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
            DataField       =   "venta_det_cantidad"
            Caption         =   "Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "venta_precio_unitario_bs"
            Caption         =   "Prec.Unitario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "venta_descuento_bs"
            Caption         =   "Descuento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "venta_precio_total_bs"
            Caption         =   "Precio Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Vendido"
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
         BeginProperty Column08 
            DataField       =   "almacen_codigo"
            Caption         =   "Almacen"
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
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   5400
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   720
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   240
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6840
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Ado_datos4"
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
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   0
      Top             =   10800
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "ado_datos14"
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
   Begin MSAdodcLib.Adodc ado_datos17 
      Height          =   330
      Left            =   9120
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "ado_datos17"
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   -120
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Ado_datos11"
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
   Begin MSAdodcLib.Adodc Ado_datos16 
      Height          =   330
      Left            =   2280
      Top             =   10800
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Ado_datos16"
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
   Begin MSAdodcLib.Adodc ado_datos15 
      Height          =   330
      Left            =   6840
      Top             =   10440
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "ado_datos15"
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
   Begin MSAdodcLib.Adodc AdoDsctos 
      Height          =   330
      Left            =   11400
      Top             =   10080
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "AdoDsctos"
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2280
      Top             =   10440
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "Ado_Datos12"
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
   Begin MSAdodcLib.Adodc Ado_datos13 
      Height          =   330
      Left            =   4560
      Top             =   10440
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "Ado_datos13"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   13680
      Top             =   10080
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "AdoAux"
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4560
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Ado_datos3"
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
   Begin MSAdodcLib.Adodc ado_datos4A 
      Height          =   330
      Left            =   9120
      Top             =   10080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "ado_datos4A"
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
   Begin Crystal.CrystalReport CryR01 
      Left            =   720
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos20 
      Height          =   330
      Left            =   4560
      Top             =   10800
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Ado_datos20"
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
   Begin Crystal.CrystalReport CryF01 
      Left            =   1200
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   6840
      Top             =   10800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   9120
      Top             =   10800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   11400
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   13680
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Ado_datos8"
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
   Begin Crystal.CrystalReport CryF02 
      Left            =   1680
      Top             =   9840
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
   Begin Crystal.CrystalReport CryQ01 
      Left            =   2160
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Height          =   1560
      Left            =   16680
      ScaleHeight     =   1500
      ScaleWidth      =   1695
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1755
   End
   Begin Crystal.CrystalReport crRecibo 
      Left            =   2640
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00404040&
      FillColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton BntImprimir2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cobranzas"
         Height          =   765
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   120
         Width           =   1005
      End
      Begin VB.CommandButton BntImprimir3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cobranzas Dolares"
         Height          =   765
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   840
         Width           =   1005
      End
   End
End
Attribute VB_Name = "fw_facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ventas
'INI QR
Enum TQRCodeEncoding
    ceALPHA
    ceBYTE
    ceNUMERIC
    ceKANJI
    ceAUTO
End Enum
Enum TQRCodeECLevel
    LEVEL_L
    LEVEL_M
    LEVEL_Q
    LEVEL_H
End Enum
Private Declare Sub FullQRCode Lib "QRCodeLib.dll" _
(ByVal autoConfigurate As Boolean, _
 ByVal AutoFit As Boolean, _
 ByVal backColor As Long, _
 ByVal barColor As Long, _
 ByVal Texto As String, _
 ByVal correctionLevel As TQRCodeECLevel, _
 ByVal encoding As TQRCodeEncoding, _
 ByVal marginpixels As Integer, _
 ByVal moduleWidth As Integer, _
 ByVal Height As Integer, _
 ByVal Width As Integer, _
 ByVal FileName As String)
Private Declare Sub FastQRCode Lib "QRCodeLib.dll" _
(ByVal Texto As String, _
 ByVal FileName As String)
Private Declare Function QRCodeLibVer Lib "QRCodeLib.dll" () As String
Dim sFile As String
Dim CadenaQ As String
'FIN QR
Dim rs_datos0 As New ADODB.Recordset     'FACTURACION
Dim rs_datos As New ADODB.Recordset     '
Dim rs_datos01 As New ADODB.Recordset     'INICIO COBRANZAS
Dim rs_datos02 As New ADODB.Recordset     'REG. COBRANZAS
Dim rs_datos1 As New ADODB.Recordset    'Detalle de Facturas por Bloques
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos4A As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset   ' Cotiza_venta
Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset   'Acumula Cobranzas para Factura
Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
Dim rs_datos20 As New ADODB.Recordset   'Cta Bancaria

Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset     'ao_ventas_cobranza_fac (para ANL Factura)
Dim rs_aux14 As New ADODB.Recordset
Dim rs_aux20 As New ADODB.Recordset

Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset

'CLASIFICADORES
Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset
'IMAGENES
Dim m_stream    As ADODB.Stream
'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir, Caracter As String
'Dim queryinicial As String
Dim queryinicial1 As String
Dim queryinicial2 As String

'Dim descri_bien As String
'VARIABLES
Dim iResult As Variant  ', i%, y%
Dim marca1 As Variant

Dim correlativo1, VAR_ID As Long
Dim nroventa, correlv, NRO_COBR As Long

Dim VAR_CANT, varTipo As Integer         'Cant_Alm,
Dim swgrabar, swnuevo, deta2 As Integer
Dim VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SW, VAR_TSOL As Integer
Dim VAR_SOL, VAR_TIPOS As Integer
Dim i As Integer
Dim VAR_COMPM As Integer
Dim VAR_DOCR, VAR_DDIF As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, COBR_BS As Double
Dim VAR_CONTAB As Double
Dim VAR_PORC As Double
Dim VAR_13, VAR_87 As Double
Dim VAR_13DOL, VAR_87DOL As Double

Dim var_literal, VAR_PROY2, VAR_CTA, VAR_PROY3 As String
Dim VAR_CODTIPO, VAR_BENEF, VAR_GLOSA, VAR_MONEDA As String
Dim VAR_COD1, VAR_COD2, VAR_COD3 As String
Dim VAR_ANIO, VAR_MES, VAR_DIA, VAR_FECHA, VAR_FFAC As String
Dim VAR_COD4, VAR_TIPOV, VAR_CITE  As String
Dim DESAUX, VARAUX, VARCODIG As String
Dim VAR_EST, VAR_FAC, VAR_DOC As String
Dim VAR_ORG, VAR_FTE, VAR_PARTIDA As String
Dim VAR_ETAPA, VAR_TCOMP, EST_PROG As String
Dim gestion0, VAR_JQ, VAR_VTIPO As String

Dim TIPOTRANS, TIPOPROC, VARDOCFAC As String
Dim VARPROC, VARSUB, VARETAPA As String
Dim VARAutor, VARFactura, VARFACIMPR, VARESTADO As String
Dim VARFactura2 As String
Dim VAR_ARCHIVO As String

Dim codigo_doc As String
Dim Numero As String
Dim Autorizacion As String
Dim NroFactura As String
Dim NitCi, VAR_NIT As String
Dim Fecha As String
Dim Monto As String
Dim Llave As String
Dim CodigoContro As String
Dim VAR_NOMD, VAR_NOMH As String
Dim VAR_DCORR, VAR_HCORR As String
Dim VAR_DGRAL, VAR_ARCH, VAR_DEPTO As String

Dim VARFECHA As Date

'Dim Exel As New Excel.Application
Dim fs As FileSystemObject      'Variable de tipo file System Object

Private Sub CmdDetalle_Click()
    FrmCobranza.Visible = True
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
''  Dim descri_bien As String
''  Dim Cant_Alm As Integer
'  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then   'EOF
'     If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then            'venta_codigo
'        If (Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG        'Ado_datos.Recordset("estado_codigo_sol") = "APR" And
''            BtnModificar.Visible = True
'            If Ado_datos.Recordset!doc_codigo_fac = "R-101" Then
''               BtnImprimir3.Visible = True
'               BtnImprimir2.Visible = False
'               'BtnImprimir3.Caption = "Facturar"
''               lbl_factura.Caption = "Nro.de Factura"
''               TxtCmpbte.Visible = True
''               TxtCmpbte.Locked = True
''               lbl_docnro.Visible = False
'               'TxtCmpbte.backColor = &H404040
'               'TxtCmpbte.ForeColor = &HFFFFFF
'               Lbl_nombre_fac.Caption = "Factura a Nombre de:                                                                                                      NIT/CI"
'               lbl_fechas.Caption = "Fecha Facturación"
'            Else
''               BtnImprimir3.Visible = False
'               BtnImprimir2.Visible = True
'               'BtnImprimir3.Caption = "Recibo"
''               lbl_factura.Caption = "Nro.de Recibo"
'               lbl_docnro.Visible = True
''               TxtCmpbte.Visible = False
'               'TxtCmpbte.Locked = False     ' CAMBIAR DE Objeto
'               'TxtCmpbte.backColor = &H80000005
'               'TxtCmpbte.ForeColor = &H80000008
'               Lbl_nombre_fac.Caption = "Recibo a Nombre de:                                                                                                       NIT/CI"
'               lbl_fechas.Caption = "Fecha de Recibo"
'            End If
''            If (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 16) Then
'''                TxtDsctoTot.backColor = &HFF&             'ROJO
''                DTPFechaProg.backColor = &HFF&             'ROJO
''            Else
''                If (Ado_datos.Recordset("cobranza_fecha_sol") > Date - 16) And (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 1) Then
'''                    TxtDsctoTot.backColor = &H80FF&           'NARANJA
''                    DTPFechaProg.backColor = &H80FF&           'NARANJA
''                Else
'''                    TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
''                    DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
''                End If
''            End If
'        Else
''            BtnModificar.Visible = False
''            BtnEliminar.Visible = False
''            BtnAprobar.Visible = False
''            BtnVer.Visible = True
''            FrmABMDet.Visible = False
''            FrmABMDet2.Visible = True
''            FrmCobranza.Visible = True
''            TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
'            DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
''            BtnImprimir3.Visible = False
'        End If
'
'        Set rs_datos2 = New Recordset
'        If rs_datos2.State = 1 Then rs_datos2.Close
'        If glusuario = "VPAREDES" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Or glusuario = "ADMIN" Then
'            rs_datos2.Open "select * From av_ventas_cobranza WHERE venta_codigo_new = " & Ado_datos.Recordset!IdFactura & " ", db, adOpenKeyset, adLockOptimistic           '(estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' AND estado_codigo1 = 'APR' and doc_codigo_fac = 'R-101' AND trans_codigo = 'X' )
'        Else
'            rs_datos2.Open "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'A' AND estado_codigo_fac = 'A') ", db, adOpenKeyset, adLockOptimistic
'        End If
'        'rs_datos2.Open .Sort = "cobranza_fecha_sol"
'        Set Ado_datos2.Recordset = rs_datos2.DataSource
'        Set dg_datos2.DataSource = Ado_datos2.Recordset
'
'
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and correl_venta = " & Ado_datos.Recordset!correl_venta & " "
'        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'        Set ado_datos14.Recordset = rs_datos14
'        ado_datos14.Recordset.Requery
'        If ado_datos14.Recordset.RecordCount > 0 Then
'            deta2 = 1
'        Else
'            deta2 = 0
'        End If
'
'        Set rs_datos16 = New ADODB.Recordset
'        If rs_datos16.State = 1 Then rs_datos16.Close
'        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_datos16.Recordset = rs_datos16
'        Ado_datos16.Recordset.Requery
'        If Ado_datos16.Recordset.RecordCount > 0 Then
'            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
'            FrmCobranza.Visible = True
'            'BtnImprimir2.Visible = True
'            'BtnImprimir3.Visible = True
'        Else
'            FrmCobranza.Visible = False
'            'BtnImprimir2.Visible = False
'            'BtnImprimir3.Visible = False
'        End If
'
'        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
'        Set rs_datos5 = New ADODB.Recordset
'        If rs_datos5.State = 1 Then rs_datos5.Close
'        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
'        Set Ado_datos5.Recordset = rs_datos5
'        dtc_desc5.BoundText = dtc_codigo5.BoundText
'        dtc_aux5.BoundText = dtc_codigo5.BoundText
'        If glusuario = "ADMIN" Or glusuario = "RVALDIVIEZO" Or glusuario = "VPAREDES" Or glusuario = "FACTURACION" Or glusuario = "SQUISPE" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Then
'            BtnImprimir5.Visible = True
'            BtnImprimir6.Visible = True
'        Else
'            BtnImprimir5.Visible = False
'            BtnImprimir6.Visible = False
'        End If
'        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'
'        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'
''        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos.Recordset!cobranza_codigo & "' ", "Foto")
''        Image2 = Img_Foto
''        'If adoLista.Recordset!estado_codigo = "APR" Then
''        CmdFoto.Visible = True
'     End If                         'venta_codigo
'     FrmDetalle.Enabled = True
'     FrmCobranza.Visible = True
'  Else
''    BtnImprimir3.Visible = False
''                BtnDesAprobar.Visible = True
''    BtnModificar.Visible = False
''    BtnEliminar.Visible = False
''    BtnVer.Visible = False
'    FrmDetalle.Enabled = False
'    FrmCobranza.Visible = False
''    FrmABMDet.Visible = False
'    FrmABMDet2.Visible = False
'  End If                            'EOF
End Sub

Private Sub Ado_datos1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Ado_datos1.Recordset.RecordCount > 0 Then
'        Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic

        'rs_datos.Close
        Dim rs_datos As New ADODB.Recordset     'FACTURACION
        Set rs_datos = New ADODB.Recordset
        If rs_datos.State = 1 Then rs_datos.Close
        If glusuario = "VPAREDES" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Or glusuario = "ADMIN" Then
            'If rs_datos.State = 1 Then rs_datos.Close
            rs_datos.Open "select * From av_ventas_cobranza WHERE venta_codigo = " & Ado_datos1.Recordset!venta_codigo & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " order by cobranza_prog_codigo", db, adOpenKeyset, adLockReadOnly    ', adLockOptimistic
        Else
            rs_datos.Open "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'A' AND estado_codigo_fac = 'A') ", db, adOpenKeyset, adLockOptimistic
        End If
        'rs_datos2.Open .Sort = "cobranza_fecha_sol"
        Set Ado_datos.Recordset = rs_datos.DataSource
        Set dg_datos.DataSource = Ado_datos.Recordset
        'If rs_datos.State = 1 Then rs_datos.Close

        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos1.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and correl_venta = " & Ado_datos.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
        Else
            deta2 = 0
        End If

        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos1.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!EDIF_CODIGO
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
    End If
End Sub

Private Sub Ado_datos2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not Ado_datos2.Recordset.BOF) And (Not Ado_datos2.Recordset.EOF) Then   'EOF
     If Ado_datos2.Recordset.RecordCount > 0 Then
        BtnModificar.Visible = True
     Else
        BtnModificar.Visible = False
     End If
  End If

End Sub

Private Sub BntImprimir2_Click()
    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
        'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
        CryF02.ReportFileName = App.Path & "\reportes\ventas\fr_cobranzas_facturadas_unidad.rpt"
        CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'MODULO DE COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If
End Sub

Private Sub BtnAñadir_Click()
 If glusuario = "ADMIN" Or glusuario = "FDELGADILLO" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "HMARIN" Then
    If Ado_datos.Recordset.RecordCount > 0 Then
       Set rs_datos2 = New Recordset
       If rs_datos2.State = 1 Then rs_datos2.Close
       rs_datos2.Open "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' AND estado_codigo1 = 'APR' and doc_codigo_fac = 'R-101' AND trans_codigo = 'X' ) ", db, adOpenKeyset, adLockOptimistic
       Set Ado_datos2.Recordset = rs_datos2.DataSource
       Set dg_datos2.DataSource = Ado_datos2.Recordset
       If Ado_datos2.Recordset.RecordCount > 0 Then
            If Ado_datos2.Recordset!venta_codigo = Ado_datos.Recordset!venta_codigo Then
                 db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'X'  WHERE cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
            Else
                 MsgBox "Error, debe Elegir un registro del mismo Edificio, para FACTURAR en bloque, verifique los datos y vuelva a intentar ...", , "Atención"
                 Exit Sub
            End If
        Else
            db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'X'  WHERE cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
            Ado_datos2.Recordset.Requery
        End If
    End If
 End If
        'select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' AND estado_codigo1 = 'APR' and doc_codigo_fac = 'R-101' AND trans_codigo = 'X' )
End Sub

Private Sub BtnAprobar_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'     If IsNull(Ado_datos02.Recordset("cobranza_observaciones")) Or (Ado_datos02.Recordset("cobranza_deuda_bs") = 0) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'        Exit Sub
'     Else
'        If Ado_datos02.Recordset("estado_codigo") = "REG" Then
'           sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'           If sino = vbYes Then
'               'If Ado_datos02.Recordset("venta_tipo") = "C" Or Ado_datos02.Recordset("venta_tipo") = "V" Then
'               '     db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
'               'End If
'               gestion0 = glGestion                 'Ado_datos02.Recordset("ges_gestion")
'               correlv = Ado_datos02.Recordset("venta_codigo")
'               nroventa = Ado_datos02.Recordset("venta_codigo")
'
'               VAR_BENEF = Ado_datos02.Recordset!beneficiario_codigo
'               VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'               VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) '+ " - Nro.: " + Trim(VAR_CITE)
'               VAR_DOL2 = Round(Ado_datos02.Recordset!cobranza_deuda_dol, 2)
'               VAR_BS2 = Round(Ado_datos02.Recordset!cobranza_deuda_bs, 2)
'               VAR_CTA = IIf(Ado_datos02.Recordset!Cta_Codigo = "", "NN", Ado_datos02.Recordset!Cta_Codigo)
'               VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'               VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'               VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'               VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'               If Ado_datos02.Recordset!Cta_Codigo <> "NN" Then
'                    VAR_FFAC = Ado_datos02.Recordset!cobranza_fecha_cobro1
'               Else
'                    VAR_FFAC = Ado_datos02.Recordset!cobranza_fecha_cobro
'               End If
'               NRO_COBR = Me.Ado_datos02.Recordset!cobranza_codigo
'               var_literal = Ado_datos02.Recordset!Literal
'               VAR_MONEDA = Ado_datos02.Recordset!tipo_moneda
'               VAR_CODTIPO = "REC"
'               VAR_DOC = "R-110"
'               VAR_ETAPA = "FIN-02-03"
'               VAR_TCOMP = "RECAUDADO (INGRESOS)"
'               VAR_ANIO = Year(VAR_FFAC)
'               VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
'
''               'Correlativo por Mes y Tipo de Comprobante
''                Set rs_aux2 = New ADODB.Recordset
''                SQL_FOR = "select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & Trim(VAR_MES) & "' and cta_codigo2 = '" & VAR_DOC & "' ) "
''                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''                If rs_aux2.RecordCount > 0 Then
''                      rs_aux2!numero_correlativo = rs_aux2!numero_correlativo + 1
''                      VAR_DOCR = rs_aux2!numero_correlativo
''                      rs_aux2.Update
''                End If
''                'Correlativo General por Documento I-E-T
''                Set rs_aux2 = New ADODB.Recordset
''                If rs_aux2.State = 1 Then rs_aux2.Close
''                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos02.Recordset!doc_codigo & "'  "
''                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''                If rs_aux2.RecordCount > 0 Then
''                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
''                    'Ado_datos02.Recordset!doc_numero = rs_aux2!correl_doc
''                    'Txt_campo1.Caption = rs_aux2!correl_doc
''                    rs_aux2.Update
''                End If
'                ' GRABA Nombre de Archivo en ao_ventas_cabecera
'
'                'Llave = Trim(rs_aux1!dosifica_llave)
'                'NitCi = Ado_datos.Recordset!beneficiario_codigo_fac     'VAR_BENEF
'                'Autorizacion = rs_aux1!dosifica_autorizacion
'
'               ' APRUEBA ao_ventas_cabecera
'               'db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'
'                'VAR_ARCH = RTrim(RTrim(Ado_datos02.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos02.Recordset!doc_numero))
'                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos02.Recordset("venta_codigo") & " "
'                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo_cargado = 'N' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos02.Recordset("venta_codigo") & " "
'
'
'               'marca1 = Ado_datos02.Recordset.Bookmark
'               'Ado_datos02.Recordset.Requery
'        '       Ado_datos02.Refresh
'               'Ado_datos02.Recordset.Move marca1 - 1
'
'               '  Set rstacumdet = New ADODB.Recordset
'                '  If rstacumdet.State = 1 Then rstacumdet.Close
'                '  rstacumdet.Open "select sum(deuda_cobrada) as Cobrobs from ao_ventas_cobranzas where ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' and venta_codigo = " & Ado_datos02.Recordset("venta_codigo"), db, adOpenKeyset, adLockOptimistic
'                '
'                '  Set rstdestino = New ADODB.Recordset
'                '  If rstdestino.State = 1 Then rstdestino.Close
'                '  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & gestion0 & "' and venta_codigo = " & nroventa, db, adOpenKeyset, adLockOptimistic
'                '  If rstdestino.RecordCount > 0 Then
'                '    rstdestino!deuda_cobrada = rstacumdet!Cobrobs
'                '    rstdestino!saldo_p_cobrar = (rstdestino!monto_total_Bs - rstdestino!monto_cobrado - rstdestino!deuda_cobrada)
'                '    rstdestino.Update
'                '  End If
'                '  If rstdestino.State = 1 Then rstdestino.Close
'                '  If rstacumdet.State = 1 Then rstacumdet.Close
'               VAR_SW = 2
'               Call Contabiliza_venta
'               db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where cobranza_codigo = " & NRO_COBR & " "
'               db.Execute "UPDATE co_diario SET co_diario.estado_codigo = co_comprobante_m.estado_codigo FROM co_diario INNER JOIN co_comprobante_m ON co_diario.Cod_Comp =co_comprobante_m.Cod_Comp where co_diario.estado_codigo Is Null "
'               Call OptFilGral1_Click
'           End If
'        End If
'     End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
'  End If
End Sub


Private Sub BtnAprobar2_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'   If Ado_datos02.Recordset!cmpbte_deposito <> "0" And Ado_datos02.Recordset!cmpbte_deposito <> "" Then
'      If Ado_datos02.Recordset!Cta_Codigo <> "NN" And Ado_datos02.Recordset!Cta_Codigo <> "" Then
'        COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
'        If IsNull(Ado_datos02.Recordset!cobranza_deuda_bs) Or (COBR_BS = 0) Then
'           MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'           Exit Sub
'        Else
'           If COBR_BS < Ado_datos02.Recordset!cobranza_total_bs Then
'               'MsgBox "No se puede APROBAR, hasta que el Monto Cobrado sea igual al Monto Facturado. Vuelva a intentar ...", , "Atención"
'               MsgBox "No se puede APROBAR hasta que el Total Monto Cobrado sea igual al Monto Facturado ...", , "Atención"
'               Ado_datos02.Recordset!cobranza_fecha_cobro1 = DTPFechaCobro2.Value
'               Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'               Ado_datos02.Recordset!estado_codigo = "REG"
'               Ado_datos02.Recordset.Update
'               'Exit Sub
'           Else
'               If Ado_datos02.Recordset("estado_codigo_bco") = "REG" Then
'                  sino = MsgBox("Esta seguro de Verificar la Cobranza ?", vbYesNo, "Confirmando")
'                  If sino = vbYes Then
'                    If TxtDscto2.Text = "0.00" Or TxtDscto2.Text = "" Then
'                       Ado_datos02.Recordset!cobranza_fecha_cobro = DTPFechaCobro2.Value
'                    Else
'                       Ado_datos02.Recordset!cobranza_fecha_cobro = DTPFechaCobro02.Value
'                    End If
'                    Ado_datos02.Recordset!cobranza_fecha_cobro1 = DTPFechaCobro2.Value
'                    Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'                    Ado_datos02.Recordset!estado_codigo_bco = "APR"
'                    Ado_datos02.Recordset!estado_codigo = "REG"
'                    Ado_datos02.Recordset.Update
'                     'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
'                  End If
'               Else
'                   MsgBox "No se puede APROBAR, el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
'               End If
'           End If
'        End If
'      Else
'        MsgBox "No se puede APROBAR, debe elegir una Cuenta Bancaria !! ", vbExclamation, "Atención!"
'      End If
'   Else
'    MsgBox "No se puede APROBAR, debe registrar el Comprobante (Cpbte) de Depósito !! ", vbExclamation, "Atención!"
'   End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
' End If
End Sub

Private Sub BtnAprobar3_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'     COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs '+ Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
'     If IsNull(Ado_datos02.Recordset!cobranza_deuda_bs) Or (COBR_BS = 0) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'        Exit Sub
'     Else
'        If COBR_BS <= Ado_datos02.Recordset!cobranza_total_bs Then
'            If Ado_datos02.Recordset("estado_codigo_bco1") = "REG" Then
'               sino = MsgBox("Esta seguro de Verificar la Cobranza 1 ?", vbYesNo, "Confirmando")
'               If sino = vbYes Then
'                  db.Execute "UPDATE ao_ventas_cobranza SET  "
'                    Ado_datos02.Recordset!cobranza_fecha_cobro = Date
'                    Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'                    Ado_datos02.Recordset!estado_codigo = "REG"
'                    Ado_datos02.Recordset.Update
'                  'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
'               End If
'            Else
'                MsgBox "No se puede APROBAR, el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
'            End If
'
'        Else
'            MsgBox "No se puede APROBAR, un Monto Cobrado Mayor al Monto Facturado. Vuelva a intentar ...", , "Atención"
'            Exit Sub
'        End If
'     End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
' End If
End Sub

Private Sub BtnBuscar_Click()
'JQA
 If Ado_datos1.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos1
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = rs_datos0.DataSource   'Ado_datos1.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnCancelar_Click()
  'Ado_datos.Refresh
  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  marca1 = Ado_datos.Recordset.Bookmark
  'If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") Then
  'OptFilGral1
  'Option2
  'OptFilGral2
  '
  If (Ado_datos1.Recordset!estado_codigo_fac = "REG") Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
  FraNavega.Enabled = True
  FrmCobros.Enabled = False
  'Fra_datos.Enabled = True
  FrmDetalle.Enabled = True
  FrmCobranza.Visible = True
  FrmCobros.Visible = False
  dg_datos.Visible = True
'  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True

  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True
  'Ado_datos.Recordset.Move marca1 - 1
'  BtnImprimir2.Visible = True
'  BtnImprimir3.Visible = True

  swnuevo = 0

End Sub

Private Sub BtnCancelarBen_Click()
    frm_benef.Visible = False
    FraGrabarCancelar.Enabled = True
End Sub

Private Sub BtnEliminar_Click()
    NumComp = Ado_datos1.Recordset!venta_codigo
    VAR_ID = Ado_datos1.Recordset!IdFactura
  If Ado_datos1.Recordset.RecordCount > 0 Then
    If (Ado_datos1.Recordset!estado_codigo_fac = "APR") And (glusuario = "SQUISPE" Or glusuario = "ADMIN" Or glusuario = "CSALINAS") Then
      sino = MsgBox("Esta seguro de ANULAR la facturación registrada ?", vbYesNo, "Confirmando")
      If sino = vbYes Then
        sino = MsgBox("Volverá a emitir otra FACTURA con el mismo DETALLE ? (Si elige NO, se cierra el registro)", vbYesNo, "Confirmando")
        If sino = vbYes Then
          'GRABA CABECERA DE FACTURACION NUEVA (ao_ventas_cobranza_fac)
'          db.Execute "INSERT INTO ao_ventas_cobranza_fac (ges_gestion, venta_codigo, doc_codigo_fac,              beneficiario_codigo_fac,                                beneficiario_nit,           glosa_Descripcion,                                  beneficiario_RazonSocial, nro_dui,      total_bs,                                       total_dol,                                      cambio_oficial, " & _
'                        " Importe_ICE, Exportaciones_Exentas, Ventas_tasa_0, Subtotal_ICE, Descuentos_Bonos, Importe_Base_Debito_Fiscal,                    factura_87_bs,                                                      factura_87_dol,                                                 debito_fiscal_13_bs,                                                debito_fiscal_13_dol,                                               literal, " & _
'                        " clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, correl_contab, estado_fac, estado_codigo_fac, estado_codigo,  " & _
'                        " usr_codigo, fecha_registro, edif_codigo_corto, edif_codigo, codigo_empresa ) " & _
'                " VALUES ('" & glGestion & "',  " & nroventa & ", '" & Ado_datos16.Recordset!doc_codigo_fac & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & dtc_codigo2A.Text & "', '" & Ado_datos16.Recordset!cobranza_concepto_plazo & "', '" & dtc_desc2A.Text & "',  '0', " & Ado_datos16.Recordset!cobranza_total_bs & ",  " & Ado_datos16.Recordset!cobranza_total_dol & ",  " & GlTipoCambioOficial & ",  " & _
'                        " '0',          '0',                    '0',            '0',            '0',    " & Ado_datos16.Recordset!cobranza_total_bs & ", " & Round(Ado_datos16.Recordset!cobranza_total_bs * 0.87, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_dol * 0.87, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_bs * 0.13, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_dol * 0.13, 2) & ", '" & Ado_datos16.Recordset!Literal & "',  " & _
'                        " 'ADM',        'R-103',        '0',        'N',            'BOB',      'NN',           'NN',        '0',            'REG',      'REG',          'REG',  " & _
'                        " '" & glusuario & "', '" & CDate(Date) & "', " & Ado_datos.Recordset!edif_codigo_corto & ", '" & Ado_datos.Recordset!EDIF_CODIGO & "', " & Ado_datos.Recordset!codigo_empresa & "  ) "
'
'            'Actualiza CORREO ELECTRONICO
'            db.Execute "UPDATE ao_ventas_cobranza_fac SET ao_ventas_cobranza_fac.beneficiario_email  = gc_beneficiario.beneficiario_email FROM ao_ventas_cobranza_fac INNER JOIN gc_beneficiario ON ao_ventas_cobranza_fac.beneficiario_codigo_fac = gc_beneficiario.beneficiario_codigo where ao_ventas_cobranza_fac.beneficiario_email Is Null "

'            Set rs_aux20 = New ADODB.Recordset
'            If rs_aux20.State = 1 Then rs_aux20.Close
'            rs_aux20.Open "Select max(IdFactura) as Codigo3 from ao_ventas_cobranza_fac  ", db, adOpenKeyset, adLockOptimistic
'            If IsNull(rs_aux20!codigo3) Then
'               VAR_IDFAC = 1
'            Else
'               VAR_IDFAC = rs_aux20!codigo3
'            End If
'            'GRABA CABECERA DE LA FACTURA (QR)
'            db.Execute "INSERT INTO ao_ventas_cobranza_fac_QR (IdFactura, archivo_foto_cargado, estado_codigo, usr_codigo, fecha_registro ) " & _
'                " VALUES ('" & VAR_IDFAC & "',  'N',            'REG',   '" & glusuario & "', '" & CDate(Date) & "' ) "
          'SE GENERAN CON LA FACTURA (dosifica_autorizacion, nro_factura, fecha_fac, codigo_control, archivo_foto, depto_codigo, Gestion, mes, edif_codigo_corto)
          db.Execute "INSERT INTO ao_ventas_cobranza_fac (ges_gestion, venta_codigo, doc_codigo_fac,              beneficiario_codigo_fac,                                      beneficiario_nit,                               glosa_Descripcion,                                  beneficiario_RazonSocial,               nro_dui,            total_bs,                                       total_dol,                                      cambio_oficial, " & _
                        " Importe_ICE, Exportaciones_Exentas, Ventas_tasa_0, Subtotal_ICE, Descuentos_Bonos,    Importe_Base_Debito_Fiscal,                                 factura_87_bs,                                       factura_87_dol,                              debito_fiscal_13_bs,                              debito_fiscal_13_dol,                             literal, " & _
                        " clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, correl_contab, estado_fac, estado_codigo_fac, estado_codigo,  " & _
                        " usr_codigo, fecha_registro, edif_codigo_corto, edif_codigo, codigo_empresa ) " & _
                " VALUES ('" & glGestion & "',  " & NumComp & ", '" & Ado_datos1.Recordset!doc_codigo_fac & "', '" & Ado_datos1.Recordset!beneficiario_codigo_fac & "', '" & Ado_datos1.Recordset!beneficiario_nit & "', '" & Ado_datos1.Recordset!glosa_Descripcion & "', '" & Ado_datos1.Recordset!beneficiario_RazonSocial & "',  '0', " & Ado_datos1.Recordset!total_bs & ",  " & Ado_datos1.Recordset!total_dol & ",  " & Ado_datos1.Recordset!cambio_oficial & ",  " & _
                        " '0',          '0',                    '0',            '0',            '0',    " & Ado_datos1.Recordset!Importe_Base_Debito_Fiscal & ", " & Ado_datos1.Recordset!factura_87_bs & ", " & Ado_datos1.Recordset!factura_87_dol & ", " & Ado_datos1.Recordset!debito_fiscal_13_bs & ", " & Ado_datos1.Recordset!debito_fiscal_13_dol & ", '" & Ado_datos1.Recordset!Literal & "',  " & _
                        " 'ADM',        'R-103',        '0',        'N',            'BOB',      'NN',           'NN',        '0',            'REG',      'REG',          'REG',  " & _
                        " '" & glusuario & "', '" & CDate(Date) & "', " & Ado_datos1.Recordset!edif_codigo_corto & ", '" & Ado_datos1.Recordset!EDIF_CODIGO & "', " & Ado_datos1.Recordset!codigo_empresa & "  ) "

          ' Actualiza IDFACTURA al DETALLE
          Set rs_aux10 = New ADODB.Recordset
          If rs_aux10.State = 1 Then rs_aux10.Close
          rs_aux10.Open "Select MAX(IdFactura) AS maxId from ao_ventas_cobranza_fac ", db, adOpenStatic
          If rs_aux10.RecordCount > 0 Then
                VAR_IDFAC = rs_aux10!maxId
                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'REG' Where venta_codigo = " & NumComp & "  and venta_codigo_new = " & VAR_ID & "  "
                db.Execute "update ao_ventas_cobranza set venta_codigo_new = " & rs_aux10!maxId & " Where venta_codigo = " & NumComp & "  and venta_codigo_new = " & VAR_ID & "  "
                db.Execute "UPDATE ao_ventas_cobranza_fac SET estado_codigo_fac = 'ANL', estado_codigo = 'ANL', estado_fac = 'ANL'  WHERE IdFactura = " & VAR_ID & "  "
          End If
          'GRABA CABECERA DE LA FACTURA (QR)
            db.Execute "INSERT INTO ao_ventas_cobranza_fac_QR (IdFactura, archivo_foto_cargado, estado_codigo, usr_codigo, fecha_registro ) " & _
                " VALUES ('" & VAR_IDFAC & "',  'N',            'REG',   '" & glusuario & "', '" & CDate(Date) & "' ) "
        Else
            db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'ANL', estado_codigo = 'ANL', estado_codigo_bco1 = 'ANL', estado_codigo_bco = 'ANL', estado_codigo_anl = 'APR'  Where venta_codigo = " & NumComp & "  and venta_codigo_new = " & VAR_ID & "  "
            db.Execute "UPDATE ao_ventas_cobranza_fac SET estado_codigo_fac = 'ANL', estado_codigo = 'ANL', estado_fac = 'ANL'  WHERE IdFactura = " & VAR_ID & "  "
        End If
'          db.Execute "update ao_ventas_cobranza set cobranza_nro_factura_anl = '" & Ado_datos.Recordset!cobranza_nro_factura & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_fecha_anl = '" & Format(Date, "dd/mm/yyyy") & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set usr_codigo_anl = '" & glusuario & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set estado_codigo_anl = 'APR' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_fecha_ant = cobranza_fecha_fac Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_codigo_control_anl = cobranza_codigo_control Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set correl_contab_anl = correl_contab Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'          db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion_anl = cobranza_nro_autorizacion Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
'
'          Set rs_datos12 = New ADODB.Recordset
'          If rs_datos12.State = 1 Then rs_datos12.Close
'          rs_datos12.Open "Select * from ao_ventas_cobro_anl where cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " and cobranza_nro_factura_anl = " & Ado_datos.Recordset!cobranza_nro_factura & " ", db, adOpenKeyset, adLockOptimistic
'          If rs_datos12.RecordCount > 0 Then
'            MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
'          Else
'            'wwwwwwwwwwwwwwwwwwwww
'              ' hora_registro
'            rs_datos12.AddNew
'            rs_datos12!ges_gestion = glGestion
'            rs_datos12!cobranza_codigo = Ado_datos.Recordset!cobranza_codigo
'            rs_datos12!venta_codigo = Ado_datos.Recordset!venta_codigo
'
'            rs_datos12!cobranza_nro_factura_anl = Ado_datos.Recordset!cobranza_nro_factura
'            rs_datos12!cobranza_prog_codigo = Ado_datos.Recordset!cobranza_prog_codigo
'            rs_datos12!beneficiario_codigo_fac = Ado_datos.Recordset!beneficiario_codigo_fac
'            rs_datos12!cobranza_anuladal_bs = Ado_datos.Recordset!cobranza_total_bs
'            rs_datos12!cobranza_anulada_dol = Ado_datos.Recordset!cobranza_total_dol
'
'            rs_datos12!cobranza_fecha_anl = Ado_datos.Recordset!cobranza_fecha_fac      'Format(Date, "dd/mm/yyyy")
'            rs_datos12!cobranza_fecha_fac2 = Ado_datos.Recordset!cobranza_fecha_fac2
'            rs_datos12!cobranza_observaciones = Ado_datos.Recordset!cobranza_observaciones
'            rs_datos12!cobranza_codigo_control_anl = Ado_datos.Recordset!cobranza_codigo_control
'            rs_datos12!Literal = Ado_datos.Recordset!Literal
'
'            rs_datos12!cobranza_nro_autorizacion_anl = Ado_datos.Recordset!cobranza_nro_autorizacion
'            rs_datos12!correl_contab_anl = Ado_datos.Recordset!correl_contab
'            rs_datos12!estado_codigo_anl = "APR"            'Ado_datos.Recordset!estado_codigo_anl
'            rs_datos12!usr_codigo_anl = glusuario           'Ado_datos.Recordset!usr_codigo_anl
'            rs_datos12!fecha_registro = Ado_datos.Recordset!fecha_registro
'
'            rs_datos12!trans_codigo = Ado_datos.Recordset!trans_codigo
'            rs_datos12!cmpbte_deposito = Ado_datos.Recordset!cmpbte_deposito
'            rs_datos12!cta_codigo = Ado_datos.Recordset!cta_codigo
'            rs_datos12.Update
'          End If
      End If
        '  rs_datos12!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
          'wwwwwwwwwwwwwwwwwwwww
          'marca1 = Ado_datos.Recordset.Bookmark
          'Call OptFilGral2_Click
          'Ado_datos.Recordset.Move marca1 - 1
    Else
      MsgBox "NO se puede ANULAR por una de las 3 razones: 1. El registro NO fue Facturado. 2. Ya fue Cobrado 3. NO tiene Permisos...", , "Atencion"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub cambiarEtiquetaFactura()
'    If lbl_fac.Caption <> "R-101" Then
'       TxtCmpbte = False
'       TxtCmpbte.backColor = &H80000005
'       TxtCmpbte.ForeColor = &H80000008
''       lbl_factura.Caption = "Nro.de Recibo"
'    Else
'       TxtCmpbte = True
'       TxtCmpbte.backColor = &H404040
'       TxtCmpbte.ForeColor = &HFFFFFF
''       lbl_factura.Caption = "Nro.de Factura"
'    End If
End Sub

Private Sub BtnGrabar_Click()
'  Call cambiarEtiquetaFactura
  If dtc_codigo4A.Text = "" Then
    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If dtc_codigo5 = "" Then
    MsgBox "Debe Elejir <<Factura a Nombre de:>> !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If TxtMonto = "" Or TxtMonto = "0" Or TxtMonto = "0.00" Then
    MsgBox "Debe Registrar el " + lbl_monto.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If TxtObs = "" Then
    MsgBox "Debe Registrar el " + lbl_obs.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
'       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    'db.BeginTrans
    If swnuevo = 1 Then
'      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
'      Set Ado_datos16.Recordset = rstdestino
'      Ado_datos16.Recordset.AddNew
      nroventa = Ado_datos.Recordset!venta_codigo
      Ado_datos.Recordset!ges_gestion = glGestion       'Ado_datos.Recordset("ges_gestion")
      'Ado_datos.Recordset!cobranza_fecha_prog = DTPFechaProg                                'Fecha Programada a Cobrar
    End If
    nroventa = Ado_datos.Recordset!venta_codigo
    NRO_COBR = Ado_datos.Recordset!cobranza_codigo
    'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
    VAR_ANIO = CStr(glGestion)
    VAR_MES = UCase(MonthName(Month(Date)))
    VAR_13 = Round(CDbl(TxtMonto.Text) * 0.13, 2)                   'Monto Cobrado Bs. * 13%
    VAR_87 = Round(CDbl(TxtMonto.Text) - VAR_13, 2)                 'Monto Cobrado Bs. * 87%
    var_literal = Literal(CStr(Ado_datos.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
    VARDOCFAC = "R-101"

    Select Case SSTab1.Tab
        Case 0
            'GRABA CGI
            TIPOTRAM = "O"
            TIPOPROC = "CGI_" + VAR_ANIO + VAR_MES
            VARPROC = "FIN"
            VARSUB = "FIN-02"
            VARETAPA = "FIN-02-02"
            VARAutor = "0"
            VARFactura = "0"
            VARFACIMPR = "N"
            VARFECHA = Format(Date, "dd/mm/yyyy")
        Case 1
            'GRABA CGE
            TIPOTRAM = "G"
            TIPOPROC = "CGE_" + VAR_ANIO + VAR_MES
            VARPROC = "CGE"
            VARSUB = "CGE-01"
            VARETAPA = "CGE-01-05"
            VARAutor = "0"
            VARFactura = "0"
            VARFACIMPR = "N"
            VARFECHA = Format(Date, "dd/mm/yyyy")
        Case 2
            'GRABA FE (Facturacion Electronica)
            TIPOTRAM = "L"
            TIPOPROC = "FE_" + VAR_ANIO + VAR_MES
            VARPROC = "FIN"
            VARSUB = "FIN-02"
            VARETAPA = "FIN-02-02"
            VARAutor = TxtAutorizacion.Text
            VARFactura = TxtCmpbte.Text
            VARFACIMPR = "S"
            VARFECHA = Format(DTPFechaCobro.Value, "dd/mm/yyyy")        '"08/04/2021"
    End Select
    db.Execute "update ao_ventas_cobranza set beneficiario_codigo_resp='" & dtc_codigo4A.Text & "', beneficiario_codigo_fac='" & dtc_codigo5 & "', cobranza_tdc=" & Txt_tdc.Text & ", cobranza_total_bs=" & TxtMonto.Text & ", cobranza_total_dol=" & TxtMontoDol.Text & ", cobranza_observaciones = '" & VAR_GLOSA & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
    db.Execute "update ao_ventas_cobranza set cta_codigo2 = '" & TIPOPROC & "', trans_codigo = '" & TIPOTRAM & "', proceso_codigo='" & VARPROC & "', subproceso_codigo = '" & VARSUB & "', etapa_codigo = '" & VARETAPA & "', Literal = '" & var_literal & "', cobranza_nro_factura = '" & VARFactura & "', cobranza_fecha_fac = '" & VARFECHA & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
    db.Execute "update ao_ventas_cobranza set cta_codigo = 'NN', cobranza_deuda_bs = '0', cobranza_deuda_dol = '0', cobranza_descuento_bs = " & VAR_13 & ", cobranza_descuento_dol = " & VAR_87 & ", cmpbte_deposito = '0', factura_impresa = '" & VARFACIMPR & "', poa_codigo = '3.1.2', estado_codigo_fac = '" & VARESTADO & "', cobranza_fecha_fac2 ='', usr_codigo = '" & glusuario & "', Fecha_Registro = '" & Format(Date, "dd/mm/yyyy") & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "

      'Ado_datos.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW

      'VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) + " - Nro.: " + Trim(VAR_CITE)
      'Ado_datos.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value                                'Fecha de Cobranza
      'Call acumulaMont(Ado_datos.Recordset!ges_gestion, Ado_datos.Recordset!correl_venta, Ado_datos.Recordset!venta_codigo)
      'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))

      Ado_datos.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion = "", "0", Trim(TxtAutorizacion))

        'VAR_ANIO = CStr(glGestion)
        'VAR_MES = CStr(Month(Date))
        'VAR_DIA = CStr(Day(Date))
       'Ado_datos.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      'Ado_datos.Recordset.Update
    'db.CommitTrans
    MsgBox "El registro se guardo correctamente"
    'Ado_datos.Recordset!doc_numero = Ado_datos.Recordset!cobranza_codigo       'Txt_cod_cobro.Text     ' "0"

    'SSTab1.Tab = 1
    'SSTab1.TabEnabled(0) = True
    'SSTab1.TabEnabled(1) = True
    FraNavega.Enabled = True
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FrmDetalle.Enabled = True
    FrmCobranza.Visible = True
    'FrmCobros.Enabled = False
    FrmCobros.Visible = False
'    BtnImprimir3.Visible = True

    swnuevo = 0

  'Else
  '  MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
  'End If

End Sub


Private Sub BtnGrabarBen_Click()
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' and beneficiario_codigo = '" & dtc_codigo8.Text & "'  ", db, adOpenStatic
    If rs_datos10.RecordCount = 0 Then
        'abrir gc_edificio_vs_beneficiario
        db.Execute "INSERT INTO gc_edificio_vs_beneficiario (edif_codigo, beneficiario_codigo, estado_codigo, fecha_registro, usr_codigo) VALUES ('" & VAR_PROY3 & "', '" & dtc_codigo8.Text & "', 'APR', '" & Date & "', '" & glusuario & "')"
        'Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
'        dtc_desc5.BoundText = dtc_codigo5.BoundText
'        dtc_aux5.BoundText = dtc_codigo5.BoundText
        'FraGrabarCancelar.Enabled = True
'        lbl_nit_fac.Caption = dtc_codigo8.Text
'        lbl_benef_fac.Caption = dtc_desc8.Text

'        lbl_nit_fac.Visible = True
'        lbl_benef_fac.Visible = True

    Else
        MsgBox "Ya existe el Beneficiario relacionado, en: <<Facturado a Nombre de>>. Vuelva a intentar ...", , "Atención"
    End If
    FraGrabarCancelar.Enabled = True
    frm_benef.Visible = False
End Sub

Private Sub BtnImprimir_Click()
    Select Case SSTab1.Tab
        Case 0
          If Ado_datos01.Recordset.RecordCount > 0 Then
''            Dim iResult As Variant  ', i%, y%
'            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
'            CryR01.WindowShowRefreshBtn = True
'            CryR01.StoredProcParam(0) = Me.Ado_datos01.Recordset!venta_codigo
'            CryR01.StoredProcParam(1) = Me.Ado_datos01.Recordset!cobranza_codigo
'            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos01.Recordset!Literal & "' "
'            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_codigo & "' "
'            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'            iResult = CryR01.PrintReport
'            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
        Case 1
          If Ado_datos.Recordset.RecordCount > 0 Then
''            Dim iResult As Variant  ', i%, y%
'            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
'            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'            CryR01.WindowShowRefreshBtn = True
'            CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
'            CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
'            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
'            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'            iResult = CryR01.PrintReport
'            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
        Case 2
'          If Ado_datos02.Recordset.RecordCount > 0 Then
''            Dim iResult As Variant  ', i%, y%
'            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
'            CryR01.WindowShowRefreshBtn = True
'            CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
'            CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
'            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos02.Recordset!Literal & "' "
'            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_codigo & "' "
'            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'            iResult = CryR01.PrintReport
'            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'          End If
    End Select

'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant, i%, y%
'    Dim co As New ADODB.Command
'
''    Dim rs As New ADODB.Recordset
''    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
''            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
''    i = 1
''    y = 1
'    CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_nota_de_venta.rpt"
'    CryV01.WindowShowRefreshBtn = True
'    CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'    CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
'    iResult = CryV01.PrintReport
'    If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
End Sub

Private Sub BtnImprimir1_Click()
    Select Case SSTab1.Tab
        Case 0
          If Ado_datos01.Recordset.RecordCount > 0 Then
''            Dim iResult As Variant  ', i%, y%
'            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_dol.rpt"
'            CryR01.WindowShowRefreshBtn = True
'            CryR01.StoredProcParam(0) = Me.Ado_datos01.Recordset!venta_codigo
'            CryR01.StoredProcParam(1) = Me.Ado_datos01.Recordset!cobranza_codigo
'            var_literal = Literal(CStr(Ado_datos01.Recordset!cobranza_programada_dol)) + " DOLARES "
'            CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_codigo & "' "
'            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'            iResult = CryR01.PrintReport
'            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
        Case 1
          If Ado_datos.Recordset.RecordCount > 0 Then
''            Dim iResult As Variant  ', i%, y%
'            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_dol.rpt"
'            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'            CryR01.WindowShowRefreshBtn = True
'            CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
'            CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
'            var_literal = Literal(CStr(Ado_datos.Recordset!cobranza_programada_dol)) + " DOLARES "
'            CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'            iResult = CryR01.PrintReport
'            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
        Case 2
'          If Ado_datos02.Recordset.RecordCount > 0 Then
''            Dim iResult As Variant  ', i%, y%
'            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_dol.rpt"
'            CryR01.WindowShowRefreshBtn = True
'            CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
'            CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
'            var_literal = Literal(CStr(Ado_datos02.Recordset!cobranza_programada_dol)) + " DOLARES "
'            CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_codigo & "' "
'            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'            iResult = CryR01.PrintReport
'            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'          End If
    End Select

End Sub

'Private Sub BtnImprimir1_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'    CryR01.WindowShowRefreshBtn = True
''    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'    CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
'    CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
'
'    CryR01.Formulas(1) = "literalcobro = '" & Ado_datos02.Recordset!Literal & "' "
'    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_codigo & "' "
'    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'    iResult = CryR01.PrintReport
'    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'
'End Sub

Private Sub BtnImprimir3_Click()
'IMPRIME FACTURA
'If Ado_datos2.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
'  If (Ado_datos2.Recordset!factura_impresa = "N") Then
If Ado_datos1.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
  If (Ado_datos1.Recordset!factura_impresa = "N") Then
    'If Ado_datos2.Recordset!cobranza_total_bs >= 3000 And dtc_aux5.Text = "0" Then
    If TxtMonto.Text >= 1000 And dtc_aux5.Text = "0" Then
        MsgBox "No se puede Imprimir una Factura >= Bs.1000, sin NIT, debe registrar el NIT del Beneficiario... ", , "Atención"
    Else
      If Ado_datos1.Recordset!doc_codigo_fac = "R-101" Then
        '===== ini GENERA EL CODIGO DE FACTURA ====             ONLINE ---- NOOO  --- ELEGIR NUMERO AUTORIZACION
        Set rs_aux1 = New ADODB.Recordset
        rs_aux1.CursorLocation = adUseClient
        If rs_aux1.State = 1 Then rs_aux1.Close
        rs_aux1.Open "select * from fc_dosificacion_docs where doc_codigo = 'R-101' AND estado_codigo = 'APR' AND dgral_codigo= '" & VAR_DGRAL & "' ", db, adOpenDynamic, adLockOptimistic
        'rs_aux1.Open "select * from fc_dosificacion_docs  where doc_codigo = 'R-101'  ", db, adOpenDynamic, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            If Date > rs_aux1!dosifica_fecha_limite Then
                MsgBox "No se puede EMITIR mas Facturas, la fecha límite de emisión EXPIRÓ, debe realizar una nueva Dosificación ... ", , "Atención"
                Exit Sub
            End If
            ' GENERA NUMERO DE FACTURA
            VAR_COD1 = CDbl(rs_aux1!CORREL) + 1
            'VALIDA SI EXISTE
            Set rs_aux12 = New ADODB.Recordset
            If rs_aux12.State = 1 Then rs_aux12.Close
            rs_aux12.Open "select * from ao_ventas_cobranza_fac where dosifica_autorizacion = '" & rs_aux1!dosifica_autorizacion & "' AND estado_codigo <> 'ERR' AND nro_factura = " & VAR_COD1 & " ", db, adOpenDynamic, adLockOptimistic
            If rs_aux12.RecordCount > 0 Then
                MsgBox "No se puede EMITIR, la Facturas YA Existe, Consulte con el Administrador del Sistema ... ", , "Atención"
                Exit Sub
            End If
            ' EMITE NRO. DE FACTURA
            db.Execute "UPDATE fc_dosificacion_docs SET CORREL = '" & Trim(Str(VAR_COD1)) & "' where doc_codigo = 'R-101' AND estado_codigo = 'APR' AND dgral_codigo= '" & VAR_DGRAL & "' "
            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
            correlv = Ado_datos1.Recordset("venta_codigo")
            nroventa = Ado_datos1.Recordset("venta_codigo")
            NRO_COBR = Ado_datos.Recordset!cobranza_codigo
            VAR_BENEF = Ado_datos1.Recordset!beneficiario_codigo_fac
            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
            'VAR_GLOSA = Trim(Ado_datos.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
            VAR_DOL2 = Round(CDbl(TxtMontoDol.Text), 2)     'Dolares
            VAR_BS2 = Round(CDbl(TxtMonto.Text), 2)         'Bolivianos
            VAR_13 = Round(VAR_BS2 * 0.13, 2)               '
            VAR_87 = Round(VAR_BS2 * 0.87, 2)
            VAR_13DOL = Round(VAR_DOL2 * 0.13, 2)              '
            VAR_87DOL = Round(VAR_DOL2 * 0.87, 2)
            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
            var_literal = Literal(CStr(VAR_BS2)) + " BOLIVIANOS"
            'var_literal = Literal(CStr(Ado_datos1.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
            VAR_FFAC = Format((Date), "DD/MM/YYYY")
            VAR_CODTIPO = "REF"     'Tipo Comprobante (paralelo VAR_DOC)
            VAR_DOC = "R-112"       'Doc. Respaldo
            VAR_ETAPA = "FIN-02-02"
            VAR_TCOMP = "RECAUDADO (FACTURACION)"
            Llave = Trim(rs_aux1!dosifica_llave)
            'If dtc_aux5.Text Like " " Then
            If Ado_datos1.Recordset!beneficiario_nit Like " " Then
                MsgBox "Error en el NIT del Cliente, Contactese con el Administrador y vuelva a intentar ...", , "Atención"
                Exit Sub
            Else
                If IsNumeric(Ado_datos1.Recordset!beneficiario_nit) Then
                    VAR_NIT = Ado_datos1.Recordset!beneficiario_nit
                           'IIf(dtc_aux5.Text = "", Ado_datos1.Recordset!beneficiario_nit, dtc_aux5.Text)    'VAR_BENEF
                Else
                    MsgBox "Error en el NIT del Cliente, Contactese con el Administrador y vuelva a intentar ...", , "Atención"
                    Exit Sub
                End If
                
            End If
            NitCi = Ado_datos1.Recordset!beneficiario_nit
            Autorizacion = rs_aux1!dosifica_autorizacion
            'Fecha = Val(Format((Date), "YYYYMMDD"))
            'Monto = Redondeo((VAR_BS2), 0)
            'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
            VAR_PROY2 = Ado_datos16.Recordset!EDIF_CODIGO
            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
            VAR_MONEDA = Ado_datos1.Recordset!tipo_moneda
            VAR_VTIPO = Ado_datos16.Recordset!venta_tipo
            VAR_DEPTO = Ado_datos16.Recordset!depto_codigo
            'CodigoContro = CodigoControl(NroFactura)
            If Autorizacion <> "" And NitCi <> "" And Llave <> "" And VAR_BS2 <> "0" And rs_aux1!CORREL >= 0 Then
                VAR_SW = 1
            Else
                VAR_SW = 0
                MsgBox "Error en Autorizacion, NIT o Llave, Contactese con el Administrador y vuelva a intentar ...", , "Atención"
                Exit Sub
            End If
'            'VAR_COD1 = CDbl(rs_aux1!CORREL) + 1         ' NRO. DE FACTURA
'            sino = MsgBox("Esta seguro(a) de EMITIR e IMPRIMIR la Factura Nro. " + Str(CDbl(rs_aux1!CORREL) + 1) + " ?", vbYesNo, "Confirmando")
'            'sino = MsgBox("Esta seguro(a) de EMITIR e IMPRIMIR la Factura ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
'                VAR_COD1 = CDbl(rs_aux1!CORREL) + 1         ' NRO. DE FACTURA
'                db.Execute "UPDATE fc_dosificacion_docs SET CORREL = '" & Trim(Str(VAR_COD1)) & "' where doc_codigo = 'R-101' AND estado_codigo = 'APR' AND dgral_codigo= '" & VAR_DGRAL & "' "
'                'rs_aux1!CORREL = Trim(Str(VAR_COD1))
'                'rs_aux1.Update
                VAR_ANIO = Year(VAR_FFAC)
                VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
                'VAR_COD1 = "4083"
                'GENERA CORREL NOTA DEBITO POR DEPTO INI
                Set rs_aux5 = New ADODB.Recordset
                If rs_aux5.State = 1 Then rs_aux5.Close
                'rs_aux5.Open "Select correl_contab as Codigo from gc_departamento where depto_codigo = '" & Left(VAR_PROY3, 1) & "'    ", db, adOpenStatic
                rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
                If Not rs_aux5.EOF Then
                    VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
                Else
                    VAR_CONTAB = 1
                End If
                db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
                'VAR_CONTAB = "17094"
                VAR_COD2 = rs_aux1!dosifica_autorizacion
                NroFactura = Trim(Str(VAR_COD1))
                VARFactura2 = NroFactura
                Fecha = Val(Format((Date), "YYYYMMDD"))
                Monto = Redondeo((VAR_BS2), 0)
                'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
                CodigoContro = CodigoControl(Autorizacion, NroFactura, VAR_NIT, Fecha, Monto, Llave)
                If CodigoContro = "" Or CodigoContro = "0" Then
                    VAR_SW = 0
                    MsgBox "Error en Codigo de Control, Contactese con el Administrador o vuelva a intentar ...", , "Atención"
                    Exit Sub
                Else
                    VAR_SW = 1
                End If
                ' GLOSA CONTABILIDAD CORRELATIVO SIC
                VAR_GLOSA = TxtObs.Text
                'VAR_ARCH = RTrim(RTrim(Ado_datos02.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos02.Recordset!doc_numero))
                VAR_ARCH = "R101" + "-" + LTrim(Str(VARFactura2))
                'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW GRABA REGISTRO WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'                ges_gestion, dosifica_autorizacion, nro_factura, correl_fac, doc_codigo_fac, fecha_fac, venta_codigo, beneficiario_codigo_fac, beneficiario_nit, glosa_factura, beneficiario_facturacion, nro_dui, total_bs, total_dol,
'                         cambio_oficial, Importe_ICE, Exportaciones_Exentas, Ventas_tasa_0, Subtotal_ICE, Descuentos_Bonos, Importe_Base_Debito_Fiscal, factura_87_bs, factura_87_dol, debito_fiscal_13_bs, debito_fiscal_13_dol, codigo_control,
'                         literal, clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, foto, archivo_foto, archivo_foto_cargado, correl_contab, Gestion, Mes, depto_codigo, etapa_codigo, estado_fac,
'                         estado_codigo_fac , estado_codigo, usr_codigo, fecha_registro, hora_registro, usr_codigo_apr, fecha_aprueba, usr_codigo_anl, fecha_anula

'                         literal, clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, foto, archivo_foto, archivo_foto_cargado, correl_contab, Gestion, Mes, depto_codigo, etapa_codigo, estado_fac,
'                         estado_codigo_fac , estado_codigo, usr_codigo, fecha_registro, hora_registro, usr_codigo_apr, fecha_aprueba, usr_codigo_anl, fecha_anula

                db.Execute "UPDATE ao_ventas_cobranza_fac SET dosifica_autorizacion = '" & VAR_COD2 & "', nro_factura = " & CDbl(VARFactura2) & ", fecha_fac = '" & VAR_FFAC & "', glosa_Descripcion =  '" & Left(VAR_GLOSA, 249) & "', codigo_control = '" & CodigoContro & "', literal = '" & var_literal & "' WHERE IdFactura = " & Ado_datos1.Recordset!IdFactura & " "

'                db.Execute "INSERT INTO ao_ventas_cobranza_fac (ges_gestion, dosifica_autorizacion, nro_factura, doc_codigo_fac, fecha_fac, venta_codigo, beneficiario_codigo_fac, beneficiario_nit, glosa_Descripcion, beneficiario_RazonSocial, nro_dui, total_bs, total_dol, cambio_oficial, " & _
'                        " Importe_ICE, Exportaciones_Exentas, Ventas_tasa_0, Subtotal_ICE, Descuentos_Bonos, Importe_Base_Debito_Fiscal, factura_87_bs, factura_87_dol, debito_fiscal_13_bs, debito_fiscal_13_dol, codigo_control, literal, " & _
'                        " clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, archivo_foto, archivo_foto_cargado, correl_contab, estado_fac, estado_codigo_fac, estado_codigo, depto_codigo, Gestion, " & _
'                        " mes , usr_codigo, fecha_registro, edif_codigo_corto) " & _
'                " VALUES ('" & gestion0 & "', '" & VAR_COD2 & "', " & CDbl(VARFactura2) & ", '" & Ado_datos1.Recordset!doc_codigo_fac & "', '" & VAR_FFAC & "', " & nroventa & ", '" & dtc_codigo5.Text & "', '" & dtc_aux5.Text & "', '" & Left(VAR_GLOSA, 249) & "', '" & dtc_desc5.Text & "', '0', " & CDbl(VAR_BS2) & ", " & CDbl(VAR_DOL2) & ", " & CDbl(txt_tdc.Text) & ",  " & _
'                        " '0', '0', '0', '0', '0', " & CDbl(VAR_BS2) & ", " & VAR_87 & ", " & VAR_87DOL & ", " & VAR_13 & ", " & VAR_13DOL & ", '" & CodigoContro & "', '" & var_literal & "',  " & _
'                        " 'ADM', 'R-103', '0', 'S', 'BOB', 'NN', 'NN',  '" & VAR_ARCH & "', 'N', " & VAR_CONTAB & ", 'REG', 'APR', 'REG', '" & VAR_DEPTO & "', '" & VAR_ANIO & "', " & _
'                        "  " & Month(VAR_FFAC) & ", '" & glusuario & "', '" & Date & "', " & Ado_datos.Recordset!edif_codigo_corto & "  ) "

'                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos02.Recordset("venta_codigo") & " "

                db.Execute "UPDATE ao_ventas_cobranza SET correl_contab = " & VAR_CONTAB & " WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "         'AND trans_codigo = 'X' AND estado_codigo1 = 'APR' "
                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "', cobranza_nro_factura = " & VAR_COD1 & " WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & "   "      '"
                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "         '
                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S', estado_codigo_fac = 'APR', estado_codigo_bco = 'REG', estado_codigo_bco1 = 'REG' WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "         '
                db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "         '
                db.Execute "update ao_ventas_cobranza set cobranza_deuda_bs = '0', cobranza_deuda_dol = '0', cobranza_deuda_bs2 = '0', cobranza_deuda_dol2 = '0' WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "         '
                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "         '
                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R101-' + '" & Str(VAR_COD1) & "' + '.JPG' WHERE venta_codigo = " & nroventa & " AND trans_codigo = 'X' AND estado_codigo1 = 'APR' "
                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' WHERE venta_codigo = " & nroventa & " AND venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "         '
                db.Execute "UPDATE ao_ventas_cobranza_fac SET edif_codigo = '" & VAR_PROY2 & "' WHERE IdFactura = " & Ado_datos1.Recordset!IdFactura & " "
                db.Execute "UPDATE ao_ventas_cobranza_fac SET estado_codigo_fac = 'APR', usr_codigo_apr = '" & glusuario & "' WHERE IdFactura = " & Ado_datos1.Recordset!IdFactura & " "
                db.Execute "UPDATE ao_ventas_cobranza_fac SET factura_impresa = 'S' WHERE IdFactura = " & Ado_datos1.Recordset!IdFactura & " "
'                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos1.Recordset!venta_codigo & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos1.Recordset!cobranza_codigo & " "
'
'                'Ado_datos.Recordset!correl_contab = VAR_CONTAB
''                If VAR_CONTAB < 10 Then
''                    'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
''                    VAR_GLOSA = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
''                End If
''                If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
''                   'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
''                   VAR_GLOSA = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
''                End If
''                If VAR_CONTAB > 99 Then
'''                    If VAR_CONTAB > 1200 Then
'''                        MsgBox "El ND Finaliza en 6564 ... ", , "Atención"
'''                    End If
''                   'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
''                   VAR_GLOSA = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
''                End If


'                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & NroVenta & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos1.Recordset!cobranza_codigo & " "
'               'GENERA CORREL NOTA DEBITO POR DEPTO FIN
                '===== ini nombre archivo de la FACTURA ====
                'db.Execute "update ao_ventas_cobranza set archivo_foto = '" & doc_codigo & "' + '-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "

                '===== fin nombre archivo de la FACTURA ====
                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza

                'IMPRIMIR FACTURA
'                VAR_ANIO = CStr(glGestion)
'                VAR_MES = CStr(Month(Date))
'                VAR_DIA = CStr(Day(Date))
'                VAR_FECHA = VAR_ANIO & VAR_MES & VAR_DIA

                'Dim F1
                'FI = Ado_datos.Recordset!cobranza_fecha_cobro
                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
                'frm_qr.Show vbModal
                'NIT del emisor, Nombre o Razón Social del emisor, Número correlativo de Factura, Número de Autorización, Fecha de emisión, Importe de la compra, Código de Control, Fecha Límite de Emisión, 0, 0, NIT / NDI Comprador, Nombre o Razón Social del comprador

                'MsgBox "Se está Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atención"

' WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW VERIFICADO
'                Select Case rs_aux1!dgral_codigo
'                    Case "1"
'                        db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'O'  WHERE cobranza_codigo = " & Ado_datos1.Recordset!cobranza_codigo & " "
'                    Case "2"
'                        db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'O'  WHERE cobranza_codigo = " & Ado_datos1.Recordset!cobranza_codigo & " "
'                    Case "0"
'                        db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'L', cta_codigo2 = 'FE_2021MAYO'  WHERE cobranza_codigo = " & Ado_datos1.Recordset!cobranza_codigo & " "
'                    Case "1"
'                        db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'O'  WHERE cobranza_codigo = " & Ado_datos1.Recordset!cobranza_codigo & " "
'                End Select

'                Ado_datos.Recordset!estado_codigo_fac = "APR"
'                Ado_datos.Recordset.Update
                'INI QR
                'sFile = "C:\Tmp\QRCode.bmp"
                '1003579028
                '& "|" & Format(Trim("0"), "###0.00") _
                '

                'NitCi = VAR_NIT
                
                
                sFile = App.Path & "\CLIENTES\QRCode.bmp"
                CadenaQ = Trim("1018533029") _
                & "|" & Trim(VAR_COD1) _
                & "|" & Trim(VAR_COD2) _
                & "|" & Format(Trim(Date), "DD/MM/YYYY") _
                & "|" & Format(Trim(VAR_BS2), "###0.00") _
                & "|" & Format(Trim(VAR_BS2), "###0.00") _
                & "|" & Trim(CodigoContro) _
                & "|" & Trim(Ado_datos1.Recordset!beneficiario_nit) _
                & "|" & Trim("0") _
                & "|" & Trim("0") _
                & "|" & Trim("0") _
                & "|" & Trim("0")

                'dtc_aux5.Text
                ' NitCi
                ' Ado_datos1.Recordset!beneficiario_nit
                
                'CadenaQ = Trim(txtNitEmisor.Text) _
                '& "|" & Trim(txtNumeroFactura.Text) _
                '& "|" & Trim(txtNumeroAutorizacion.Text) _
                '& "|" & Format(Trim(txtFechaEmision.Text), "DD/MM/YYYY") _
                '& "|" & Format(Trim(txtImporteCompra.Text), "###0.00") _
                '& "|" & Format(Trim(txtFiscal.Text), "###0.00") _
                '& "|" & Trim(txtCodigoControl.Text) _
                '& "|" & Trim(txtNitComprador.Text) _
                '& "|" & Trim(txtImporteICE.Text) _
                '& "|" & Trim(txtGravadas.Text) _
                '& "|" & Trim(txtNoFiscal) _
                '& "|" & Trim(TxtDescuento)
'                MsgBox CadenaQ
                FastQRCode CadenaQ, sFile
                Set Picture1.Picture = LoadPicture(sFile)
                'FIN QR
                'Call IMPRIME_FACTURA
                Call IMPRIME_QR
                'MsgBox CadenaQ
'                'If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                'End If
'                db.Execute "UPDATE co_diario SET co_diario.estado_codigo = co_comprobante_m.estado_codigo FROM co_diario INNER JOIN co_comprobante_m ON co_diario.Cod_Comp =co_comprobante_m.Cod_Comp where co_diario.estado_codigo Is Null "
'            Else
'                VAR_COD1 = "0"
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
        End If
        If rs_aux1.State = 1 Then rs_aux1.Close
        '===== fin TERMINA GENERACION DE FACTURA =====
'        '===== ini GENERA NRO. AUTORIZACION DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FAC_AUTORIZA'", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'          VAR_COD2 = CDbl(rs_aux1!numero_correlativo)
'          'rs_aux1!numero_correlativo = Trim(Str(VAR_COD2))
'          'rs_aux1.Update
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION NRO. AUTORIZACION DE FACTURA =====
'        Dim iResult As Variant  ', i%, y%
'        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R-101_factura.rpt"
'        CryF01.WindowShowRefreshBtn = True
'        CryF01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'        CryF01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'        CryF01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'
'        CryF01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
'        CryF01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'        iResult = CryF01.PrintReport
'        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"

          'Ado_datos.Refresh
          fraOpciones.Visible = True
          'FraGrabarCancelar.Visible = False
          marca1 = Ado_datos1.Recordset.Bookmark
          'If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") Then
          If (Ado_datos1.Recordset!estado_codigo_fac = "REG") Then
            Call OptFilGral1_Click
          Else
            Call OptFilGral2_Click
          End If
          FraNavega.Enabled = True
          FrmCobros.Enabled = False
          FrmDetalle.Enabled = True
          FrmCobranza.Visible = True
          FrmCobros.Visible = False
          dg_datos.Visible = True
        '  FrmABMDet.Visible = True
          FrmABMDet2.Visible = True

          SSTab1.Tab = 0
          SSTab1.TabEnabled(0) = True
          SSTab1.TabEnabled(1) = True
          'Ado_datos.Recordset.Move marca1 - 1
          swnuevo = 0

        TxtCmpbte = VAR_COD1
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
          Call OptFilGral1_Click
        Else
          Call OptFilGral2_Click
        End If
      Else
'        Call generarRepRecibo
      End If
'      If Ado_datos.Recordset!doc_codigo_fac = "R-103" Then
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'        '===== ini GENERA EL CODIGO DE RECIBO ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from gc_documentos_respaldo where doc_codigo = 'R-103' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
'            correlv = Ado_datos.Recordset("venta_codigo")
'            nroventa = Ado_datos.Recordset("venta_codigo")
'            NRO_COBR = Me.Ado_datos.Recordset!cobranza_codigo
'            VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
'            VAR_GLOSA = Trim(Ado_datos.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
'            VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_deuda_dol, 2)
'            VAR_BS2 = Round(Ado_datos.Recordset!cobranza_deuda_bs, 2)
'            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
'            var_literal = Ado_datos.Recordset!Literal
'            'Llave = Trim(rs_aux1!dosifica_llave)
'            NitCi = IIf(dtc_aux5.Text = "", Ado_datos.Recordset!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
'            'Autorizacion = rs_aux1!dosifica_autorizacion
'            VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'            VAR_MONEDA = Ado_datos.Recordset!tipo_moneda
'
'            VAR_COD1 = CDbl(rs_aux1!correl_doc) + 1
'            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Recibo Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
'                rs_aux1!correl_doc = Trim(Str(VAR_COD1))
'                rs_aux1.Update
'                'GENERA CORREL NOTA DEBITO POR DEPTO INI
''                Set rs_aux5 = New ADODB.Recordset
''                If rs_aux5.State = 1 Then rs_aux5.Close
''                rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
''                If Not rs_aux5.EOF Then
''                    VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
''                End If
''                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
''                db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
''                If VAR_CONTAB < 10 Then
''                    VAR_GLOSA = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
''                End If
''                If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
''                   VAR_GLOSA = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
''                End If
''                If VAR_CONTAB > 99 And VAR_CONTAB < 6564 Then
''                    If VAR_CONTAB > 1200 Then
''                        MsgBox "El ND Finaliza en 6564 ... ", , "Atención"
''                    End If
''                   VAR_GLOSA = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
''                End If
'                VAR_GLOSA = TxtObs.Text
'                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
''                'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'
'                VAR_COD2 = "0"  'rs_aux1!dosifica_autorizacion
'                NroFactura = Trim(Str(VAR_COD1))
'                '===== ini nombre archivo de la FACTURA ====
'                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R103-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                '===== fin nombre archivo de la FACTURA ====
'                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                'IMPRIMIR FACTURA
'                Fecha = Val(Format((Date), "YYYYMMDD"))
'                Monto = Redondeo((VAR_BS2), 0)
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
'                'Dim F1
'                'FI = Ado_datos.Recordset!cobranza_fecha_cobro
'                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
'                'frm_qr.Show vbModal
'                'NIT del emisor, Nombre o Razón Social del emisor, Número correlativo de Factura, Número de Autorización, Fecha de emisión, Importe de la compra, Código de Control, Fecha Límite de Emisión, 0, 0, NIT / NDI Comprador, Nombre o Razón Social del comprador
'
'                'MsgBox "Se está Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atención"
'                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'
'                VAR_SW = 1
'                'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
'                'db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'                Call IMPRIME_RECIBO
'                'If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                'End If
'            Else
'                VAR_COD1 = "0"
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
'        End If

'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION DE FACTURA =====
'        TxtCmpbte = VAR_COD1
'        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'      End If

    End If
  Else
        MsgBox "La Factura Nro. " + Ado_datos.Recordset!cobranza_nro_factura + " ya fue Impresa", , "Atención"
        'Call IMPRIME_FACTURA
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
          Call OptFilGral1_Click
        Else
          Call OptFilGral2_Click
        End If
  End If
Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
End If
End Sub

Private Sub generar(Autorizacion As String, Numero As String, NitCi As String, Fecha As String, Monto As String, Llave As String)
' paso 1
'    Dim suma As String
'    Dim digitos As String
'    Dim digitossum(4) As Integer
'    Dim cadenas(4) As String
'    Dim inicio As Integer
'    Dim x As Integer
'
'    Dim arc4 As String
'    Dim suma_total As Long
'    Dim sumas(4) As Long
'    Dim strlen_arc4 As Integer
'    Dim i As Integer
'    Dim total As Long
'
'    Dim mensaje As String
'    Dim last As String
'
'        numero = verhoeff_add_recursive(numero, 2)
'        nitci = verhoeff_add_recursive(nitci, 2)
'        fecha = verhoeff_add_recursive(fecha, 2)
'        monto = verhoeff_add_recursive(monto, 2)
''            Dim suma As String = CType((Long.Parse(numero) _
''                        + (Long.Parse(nitci) _
''                        + (Long.Parse(fecha) + Long.Parse(monto)))),Long).ToString
'        suma = (CStr(numero) + (CStr(nitci) + (Trim(fecha) + CStr(monto))))
'        suma = verhoeff_add_recursive(suma, 5)
'' paso2
''            Dim digitos As String = ("" + suma.Substring((suma.Length - 5), 5))
''            Dim digitossum() As Integer = New Integer() {0, 0, 0, 0, 0}
''            Dim cadenas() As String = New String() {"", "", "", "", ""}
''            Dim inicio As Integer = 0
''            Dim x As Integer = 0
'    digitos = ("" + suma.Substring((suma.Length - 5), 5))
'    digitossum(0) = 0
'    digitossum(1) = 0
'    digitossum(2) = 0
'    digitossum(3) = 0
'    digitossum(4) = 0
'    cadenas(0) = ""
'    cadenas(1) = ""
'    cadenas(2) = ""
'    cadenas(3) = ""
'    cadenas(4) = ""
'    inicio = 0
'    x = 0
''    For Each d As Char In digitos.ToCharArray
''                digitossum(x) = (Integer.Parse(d.ToString) + 1)
''                cadenas(x) = llave.Substring(inicio, (Integer.Parse(d.ToString) + 1))
''                inicio = (inicio _
''                            + (Integer.Parse(d.ToString) + 1))
''                x = (x + 1)
''    Next
'    For x = 0 To Len(digitos)
'        digitossum(x) = (CInt(digitos) + 1)
'        cadenas(x) = llave.Substring(inicio, (CInt(digitos) + 1))
'        inicio = (inicio + (CInt(digitos) + 1))
'        x = (x + 1)
'    Next x
'            autorizacion = (autorizacion + cadenas(0))
'            numero = (numero + cadenas(1))
'            nitci = (nitci + cadenas(2))
'            fecha = (fecha + cadenas(3))
'            monto = (monto + cadenas(4))
'' paso3
'    arc4 = allegedrc4((autorizacion + (numero + (nitci + (fecha + monto)))), (llave + digitos))
'' paso4
'    suma_total = 0
'    sumas(0) = 0
'    sumas(1) = 0
'    sumas(2) = 0
'    sumas(3) = 0
'    sumas(4) = 0
'    strlen_arc4 = Len(arc4)
'    i = 0
'    Do While (i < strlen_arc4)
'                x = CInt(arc4(i))
'                sumas((i Mod 5)) = (sumas((i Mod 5)) + x)
'                suma_total = (suma_total + x)
'                i = (i + 1)
'    Loop
'' paso5
'    total = 0
'    i = 0
'    Do While (i < Len(sumas))
'                total = (total + (suma_total * (sumas(i) / digitossum(i))))
'                i = (i + 1)
'    Loop
'    mensaje = big_base_convert(total, 64)
'    last = allegedrc4(mensaje, (llave + digitos)).Insert(2, "-").Insert(5, "-").Insert(8, "-")
'            If (last.Length > 11) Then
'                last = last.Insert(11, "-")
'            End If
'    'Return last

End Sub

Private Sub big_base_convert(ByVal Numero As Long, ByVal baseconv As Long)
'    Dim dic(63) As Char
'    Dim cociente As Long
'    Dim resto As Long
'    Dim palabra As String
'
'    dic(0) = Microsoft.VisualBasic.ChrW(48)
'    dic(1) = Microsoft.VisualBasic.ChrW(49)
'    dic(2) = Microsoft.VisualBasic.ChrW(50)
'    dic(3) = Microsoft.VisualBasic.ChrW(51)
'    dic(4) = Microsoft.VisualBasic.ChrW(52)
'    dic(5) = Microsoft.VisualBasic.ChrW(53)
'    dic(6) = Microsoft.VisualBasic.ChrW(54)
'    dic(7) = Microsoft.VisualBasic.ChrW(55)
'    dic(8) = Microsoft.VisualBasic.ChrW(56)
'    dic(9) = Microsoft.VisualBasic.ChrW(57)
'    dic(10) = Microsoft.VisualBasic.ChrW(65)
'    dic(11) = Microsoft.VisualBasic.ChrW(66)
'    dic(12) = Microsoft.VisualBasic.ChrW(67)
'    dic(13) = Microsoft.VisualBasic.ChrW(68)
'    dic(14) = Microsoft.VisualBasic.ChrW(69)
'    dic(15) = Microsoft.VisualBasic.ChrW(70)
'    dic(16) = Microsoft.VisualBasic.ChrW(71)
'    dic(17) = Microsoft.VisualBasic.ChrW(72)
'    dic(18) = Microsoft.VisualBasic.ChrW(73)
'    dic(19) = Microsoft.VisualBasic.ChrW(74)
'    dic(20) = Microsoft.VisualBasic.ChrW(75)
'    dic(21) = Microsoft.VisualBasic.ChrW(76)
'    dic(22) = Microsoft.VisualBasic.ChrW(77)
'    dic(23) = Microsoft.VisualBasic.ChrW(78)
'    dic(24) = Microsoft.VisualBasic.ChrW(79)
'    dic(25) = Microsoft.VisualBasic.ChrW(80)
'    dic(26) = Microsoft.VisualBasic.ChrW(81)
'    dic(27) = Microsoft.VisualBasic.ChrW(82)
'    dic(28) = Microsoft.VisualBasic.ChrW(83)
'    dic(29) = Microsoft.VisualBasic.ChrW(84)
'    dic(30) = Microsoft.VisualBasic.ChrW(85)
'    dic(31) = Microsoft.VisualBasic.ChrW(86)
'    dic(32) = Microsoft.VisualBasic.ChrW(87)
'    dic(33) = Microsoft.VisualBasic.ChrW(88)
'    dic(34) = Microsoft.VisualBasic.ChrW(89)
'    dic(35) = Microsoft.VisualBasic.ChrW(90)
'    dic(36) = Microsoft.VisualBasic.ChrW(97)
'    dic(37) = Microsoft.VisualBasic.ChrW(98)
'    dic(38) = Microsoft.VisualBasic.ChrW(99)
'    dic(39) = Microsoft.VisualBasic.ChrW(100)
'    dic(40) = Microsoft.VisualBasic.ChrW(101)
'    dic(41) = Microsoft.VisualBasic.ChrW(102)
'    dic(42) = Microsoft.VisualBasic.ChrW(103)
'    dic(43) = Microsoft.VisualBasic.ChrW(104)
'    dic(44) = Microsoft.VisualBasic.ChrW(105)
'    dic(45) = Microsoft.VisualBasic.ChrW(106)
'    dic(46) = Microsoft.VisualBasic.ChrW(107)
'    dic(47) = Microsoft.VisualBasic.ChrW(108)
'    dic(48) = Microsoft.VisualBasic.ChrW(109)
'    dic(49) = Microsoft.VisualBasic.ChrW(110)
'    dic(50) = Microsoft.VisualBasic.ChrW(111)
'    dic(51) = Microsoft.VisualBasic.ChrW(112)
'    dic(52) = Microsoft.VisualBasic.ChrW(113)
'    dic(53) = Microsoft.VisualBasic.ChrW(114)
'    dic(54) = Microsoft.VisualBasic.ChrW(115)
'    dic(55) = Microsoft.VisualBasic.ChrW(116)
'    dic(56) = Microsoft.VisualBasic.ChrW(117)
'    dic(57) = Microsoft.VisualBasic.ChrW(118)
'    dic(58) = Microsoft.VisualBasic.ChrW(119)
'    dic(59) = Microsoft.VisualBasic.ChrW(120)
'    dic(60) = Microsoft.VisualBasic.ChrW(121)
'    dic(61) = Microsoft.VisualBasic.ChrW(122)
'    dic(62) = Microsoft.VisualBasic.ChrW(43)
'    dic(63) = Microsoft.VisualBasic.ChrW(47)
'
'    cociente = 1
'    resto = 0
'    palabra = ""
'    While (cociente > 0)
'                cociente = (numero / baseconv)
'                resto = (numero Mod baseconv)
'                palabra = (dic(resto) + palabra)
'                numero = cociente
'
'    End
'    '        Return palabra
End Sub

Private Sub SWAP(ByRef num1 As Integer, ByRef num2 As Integer)
    Dim temp As Integer
    temp = num2
    num2 = num1
    num1 = temp
End Sub

'Private Sub allegedrc4(mensaje As String, llaverc4 As String)
'            Dim state() As Integer = New Integer((256) - 1) {}
'            Dim x As Integer = 0
'            Dim y As Integer = 0
'            Dim index1 As Integer = 0
'            Dim index2 As Integer = 0
'            Dim nmen As Integer = 0
'            Dim i As Integer = 0
'            Dim cifrado As String = ""
'            i = 0
'            Do While (i < 256)
'                state(i) = i
'                i = (i + 1)
'            Loop
'            Dim strlen_llave As Integer = llaverc4.Length
'            Dim strlen_mensaje As Integer = mensaje.Length
'            i = 0
'            Do While (i < 256)
'                index2 = ((CType(llaverc4(index1),Integer) _
'                            + (state(i) + index2)) _
'                            Mod 256)
'                swap(state(index2), state(i))
'                index1 = ((index1 + 1) _
'                            Mod strlen_llave)
'                i = (i + 1)
'            Loop
'            Dim cadtemp As String = ""
'            i = 0
'            Do While (i < strlen_mensaje)
'                x = ((x + 1) _
'                            Mod 256)
'                y = ((state(x) + y) _
'                            Mod 256)
'                swap(state(y), state(x))
'                ' ^ = XOR function
'                nmen = (CType(mensaje(i),Integer) Or state(((state(x) + state(y)) _
'                            Mod 256)))
'                'The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
'                cadtemp = ("0" + big_base_convert(nmen, 16))
'                cifrado = (cifrado + cadtemp.Substring((cadtemp.Length - 2), 2))
'                i = (i + 1)
'            Loop
'            Return cifrado
'End Sub
'
'Private Shared Function calcsum(ByVal number As String) As Integer
'            Dim c As Integer = 0
'            Dim n As String = reverse(number)
'            Dim len As Integer = n.Length
'            Dim nchar() As Char = n.ToCharArray
'            Dim i As Integer = 0
'            Do While (i < len)
'                c = table_d(c, table_p(((i + 1) _
'                            Mod 8), Integer.Parse(nchar(i).ToString)))
'                i = (i + 1)
'            Loop
'            Return table_inv(c)
'End Sub
'
'Private Shared Function verhoeff_add_recursive(ByVal number As String, ByVal digits As Integer) As String
'            Dim temp As String = number
'
'            While (digits > 0)
'                temp = (temp + calcsum(temp))
'                digits = (digits - 1)
'
'            End While
'            Return temp
'End Sub
'
'Private Shared Function reverse(ByVal cadena As String) As String
'            Dim str() As Char = cadena.ToCharArray
'            Array.Reverse(str)
'            Return New String(str)
'End Sub

Private Sub IMPRIME_FACTURA()
        'IMPRIMIR FACTURA
    Dim iResult As Variant  ', i%, y%
    sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
    If sino = vbYes Then
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior_rep.rpt"
    Else
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior.rpt"
    End If
        CryF01.WindowShowRefreshBtn = True
        CryF01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryF01.PrintReport
        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"

End Sub

Private Sub IMPRIME_QR()
    'IMPRIMIR FACTURA con QR
    'Dim Exel As Object
    'Set Exel = CreateObject("Excel.Application")
    'Exel.Workbooks.Open "c:\tmp\Factura.xlt", , , , "123", "123"
    'Exel.Visible = True
    Call CmdFoto_Click
    ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"

    Picture2.AutoRedraw = True
    Picture2.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight

    ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
    ' MsgBox CadenaQr
    FastQRCode CadenaQr, ImagenQr
    Picture1.AutoRedraw = True
    Picture1.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Clipboard.Clear
    Clipboard.SetData Picture2.Image
'    Exel.Application.Range("a2").Select
'    Exel.Application.ActiveSheet.Paste
    NRO_COBR = Ado_datos1.Recordset!IdFactura
    'NRO_COBR = Ado_datos.Recordset!cobranza_codigo
    Dim iResult As Variant  ', i%, y%
'    sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
'    If sino = vbYes Then
'        If VAR_COD4 = "DNMAN" Then
'            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_man.rpt"
'        Else
'            'CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
'            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_PRUEBA.rpt"
'        End If
'    Else
        'CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_PRUEBA.rpt"
        CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_facturaNEW.rpt"
'    End If
        CryQ01.WindowShowRefreshBtn = True
        CryQ01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryQ01.PrintReport
        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"

End Sub

Private Sub IMPRIME_RECIBO()
        'IMPRIMIR FACTURA
        Dim iResult As Variant  ', i%, y%
        'CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R-101_factura.rpt"
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_oficial.rpt"
        CryF01.WindowShowRefreshBtn = True
        CryF01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryF01.PrintReport
        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"

End Sub
Private Sub BtnImprimir4_Click()
    Select Case SSTab1.Tab
        Case 0
            If Ado_datos16.Recordset.RecordCount > 0 Then
              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
              CryV01.WindowShowRefreshBtn = True
              CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion            'glGestion
              CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo           'nroventa        '
              CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_prog_codigo   'NRO_COBR        '
              'Literal por el Total de la Compra
              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_prog_codigo & "' "
              iResult = CryV01.PrintReport
              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
            Else
              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
            End If
        Case 1
            If Ado_datos16.Recordset.RecordCount > 0 Then
              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
              CryV01.WindowShowRefreshBtn = True
              CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion            'glGestion
              CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo           'nroventa        '
              CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_prog_codigo   'NRO_COBR        '
              'Literal por el Total de la Compra
              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_prog_codigo & "' "
              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
              iResult = CryV01.PrintReport
              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
            Else
              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
            End If
        Case 2  'Ado_datos02
'            If Ado_datos16.Recordset.RecordCount > 0 Then
'              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
'              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
'              CryV01.WindowShowRefreshBtn = True
'              CryV01.StoredProcParam(0) = Me.Ado_datos02.Recordset!ges_gestion            'glGestion
'              CryV01.StoredProcParam(1) = Me.Ado_datos02.Recordset!venta_codigo           'nroventa        '
'              CryV01.StoredProcParam(2) = Me.Ado_datos02.Recordset!cobranza_prog_codigo   'NRO_COBR        '
'              'Literal por el Total de la Compra
'              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
'              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
'              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_prog_codigo & "' "
'              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'              iResult = CryV01.PrintReport
'              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'            Else
'              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'            End If
    End Select

End Sub

Private Sub BtnImprimir5_Click()
  'RE-IMPRIME FACTURA
  'If Ado_datos.Recordset.RecordCount > 0 And (Ado_datos.Recordset!cobranza_nro_factura > 0) Then         'And (dtc_aux5.Text <> "")
  If Ado_datos1.Recordset.RecordCount > 0 And (Ado_datos1.Recordset!nro_factura > 0) Then         '
        gestion0 = Ado_datos1.Recordset!ges_gestion
        nroventa = Ado_datos.Recordset!venta_codigo
        'NRO_COBR = Ado_datos.Recordset!cobranza_codigo
        'VAR_COD1 = Ado_datos.Recordset!cobranza_nro_factura
        NRO_COBR = Ado_datos1.Recordset!IdFactura
        'VAR_COD1 = Ado_datos1.Recordset!nro_factura
        VAR_COD4 = Ado_datos.Recordset!unidad_codigo
        VAR_BS2 = Round(CDbl(Ado_datos1.Recordset!total_bs), 2)
        var_literal = Literal(CStr(VAR_BS2)) + " BOLIVIANOS"
        db.Execute "UPDATE ao_ventas_cobranza_fac SET literal = '" & var_literal & "' WHERE IdFactura = " & NRO_COBR & " "
'        'Dim Exel As Object
'        'Set Exel = CreateObject("Excel.Application")
'        'Exel.Workbooks.Open "c:\tmp\Factura.xlt", , , , "123", "123"
'        'Exel.Visible = True
'        Call CmdFoto_Click
'        'ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'        ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'
'        Picture2.AutoRedraw = True
'        Picture2.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
'
'        ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'        ' MsgBox CadenaQr
'        FastQRCode CadenaQr, ImagenQr
'        Picture1.AutoRedraw = True
'        Picture1.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'        Clipboard.Clear
'        Clipboard.SetData Picture2.Image
'    '    Exel.Application.Range("a2").Select
'    '    Exel.Application.ActiveSheet.Paste

        Dim iResult As Variant  ', i%, y%
'        sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
'
'        If sino = vbYes Then
'            If VAR_COD4 = "DNMAN" Then
'                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_man.rpt"
'                'CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_PRUEBA.rpt"
'            Else
'                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
'            End If
'        Else
            'CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_PRUEBA.rpt"
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_facturaNEW.rpt"
'        End If
        CryQ01.WindowShowRefreshBtn = True
        CryQ01.StoredProcParam(0) = gestion0       'Me.Ado_datos.Recordset!ges_gestion
        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        var_literal = Ado_datos.Recordset!Literal
        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryQ01.PrintReport
        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"
  Else
      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir6_Click()

'RE-IMPRIME FACTURA
'If Not Ado_datos.Recordset.BOF And Not Ado_datos.Recordset.EOF Then
  If Ado_datos.Recordset.RecordCount > 0 And (Ado_datos.Recordset!cobranza_nro_factura > 0) Then         'And (dtc_aux5.Text <> "")
        gestion0 = Ado_datos.Recordset!ges_gestion
        nroventa = Ado_datos.Recordset!venta_codigo
        'NRO_COBR = Ado_datos.Recordset!cobranza_codigo
        NRO_COBR = Ado_datos1.Recordset!IdFactura
        'VAR_COD1 = Ado_datos.Recordset!cobranza_nro_factura
        VAR_COD1 = Ado_datos1.Recordset!nro_factura
        VAR_COD4 = Ado_datos.Recordset!unidad_codigo
        VAR_BS2 = Round(CDbl(Ado_datos1.Recordset!total_bs), 2)
        var_literal = Literal(CStr(VAR_BS2)) + " BOLIVIANOS"
        db.Execute "UPDATE ao_ventas_cobranza_fac SET literal = '" & var_literal & "' WHERE IdFactura = " & NRO_COBR & " "
        
        Dim iResult As Variant  ', i%, y%
        'sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")

'        If sino = vbYes Then
'            If VAR_COD4 = "DNMAN" Then
'                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_man_DPF.rpt"
'            Else
'                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep_PDF.rpt"
'            End If
'        Else
'            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_PDF.rpt"
'        End If
        If Ado_datos.Recordset.RecordCount > 1 Then
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_PDF_VARIOS.rpt"
        Else
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_PDF.rpt"
        End If
        CryQ01.WindowShowRefreshBtn = True
        CryQ01.StoredProcParam(0) = gestion0       'Me.Ado_datos.Recordset!ges_gestion
        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        var_literal = Ado_datos.Recordset!Literal
        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryQ01.PrintReport
        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"
  Else
      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
'End If
End Sub

Private Sub BtnModificar_Click()
    VAR_ID = Ado_datos1.Recordset!IdFactura
  codigo_doc = lbl_fac.Caption
  If Ado_datos1.Recordset.RecordCount > 0 Then
'    If codigo_doc <> "R-101" Then
''         Call cambiarEtiquetaFactura
'         Dim Cmd1 As ADODB.Command
'         Dim rs  As ADODB.Recordset
'         Set Cmd1 = New ADODB.Command
'         Set rs = New ADODB.Recordset
'         If glusuario = "RVALDIVIEZO" Or glusuario = "GSOLIZ" Then
'            DTPFechaCobro.Enabled = True
'         Else
'            DTPFechaCobro.Enabled = False
'         End If
'         Cmd1.ActiveConnection = db 'sqlServer
'         Cmd1.CommandType = adCmdStoredProc
'         Cmd1.CommandText = "ap_genera_codigoregistro"
'         Set Parm1 = Cmd1.CreateParameter("@codigo_doc", adVarChar, adParamInput, 200, codigo_doc)
'         Cmd1.Parameters.Append Parm1
'         rs.Open Cmd1
'         rs.MoveFirst
'         TxtCmpbte.Text = rs!Codigo
'         rs.Close
'    Else
'        Call cambiarEtiquetaFactura
'    End If
    'If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") And (Ado_datos.Recordset!venta_tipo = "E" Or Ado_datos.Recordset!venta_tipo = "V" Or Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "L") Then
    'If (Ado_datos1.Recordset!estado_codigo1 = "APR" And Ado_datos1.Recordset!trans_codigo = "X") Then
    If (Ado_datos1.Recordset!estado_codigo_fac = "REG") Then
       Nro = Ado_datos1.Recordset!venta_codigo
       FraNavega.Enabled = False
       fraOpciones.Visible = False
       FraGrabarCancelar.Visible = True
       FrmDetalle.Enabled = False
       FrmCobranza.Enabled = False
       BtnImprimir3.Visible = True
       'swgrabar = 0
       swnuevo = 2
       Set rs_datos18 = New Recordset
       If rs_datos18.State = 1 Then rs_datos18.Close
       'rs_datos2.Open "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' AND estado_codigo1 = 'APR' and doc_codigo_fac = 'R-101' AND trans_codigo = 'X' ) ", db, adOpenKeyset, adLockOptimistic
       rs_datos18.Open "select sum(cobranza_total_bs) as totbs2, sum(cobranza_total_dol) as totdl2 from ao_ventas_cobranza where venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " AND venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic           'estado_codigo1 = 'APR' AND trans_codigo = 'X'
        If IsNull(rs_datos18!totbs2) Then
            TxtMonto = 0
            TxtMontoDol = 0
        Else
            TxtMonto = Round(rs_datos18!totbs2, 2)
            TxtMontoDol = Round(rs_datos18!totdl2, 2)
        End If
'        Set rs_aux8 = New Recordset
'        If rs_aux8.State = 1 Then rs_aux8.Close
'        rs_aux8.Open "select * From ao_ventas_cobranza WHERE venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " AND venta_codigo = " & Nro & " ", db, adOpenKeyset, adLockOptimistic     'estado_codigo1 = 'APR' AND trans_codigo = 'X'
'        If rs_aux8.RecordCount > 0 Then
'            rs_aux8.MoveFirst
'            VAR_GLOSA = ""
'            While Not rs_aux8.EOF
'                VAR_GLOSA = VAR_GLOSA + Trim(rs_aux8!cobranza_observaciones) + ". "    ' + " - Tram.: " + Trim(VAR_CITE)
'                rs_aux8.MoveNext
'            Wend
'            TxtObs.Text = VAR_GLOSA
'        Else
'            TxtObs.Text = Ado_datos2.Recordset!cobranza_observaciones
'        End If
        TxtObs.Text = Ado_datos1.Recordset!glosa_Descripcion
        Txt_tdc.Text = GlTipoCambioMercado    'GlTipoCambioOficial
'      SSTab1.Tab = 0
'      SSTab1.TabEnabled(0) = False
'      SSTab1.TabEnabled(1) = True
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
'        dtc_codigo5 = Ado_datos1.Recordset!beneficiario_codigo_fac
'        dtc_desc5.BoundText = dtc_codigo5
'        dtc_aux5.BoundText = dtc_codigo5.BoundText
'        dtc_codigo5.Text = Ado_datos1.Recordset!beneficiario_codigo_fac
'      BtnImprimir2.Visible = False
'      BtnImprimir3.Visible = False
        CmdFoto.Visible = False
        Select Case SSTab1.Tab
        Case 0
            'GRABA CGI
            DTPFechaCobro.Visible = True
            DTPFechaCobro.Value = Date
            TxtCmpbte.Text = "0"
            TxtCmpbte.Locked = True
            VAR_DGRAL = "1"
        Case 1
            'GRABA CGE
            DTPFechaCobro.Visible = True
            DTPFechaCobro.Value = Date
            TxtCmpbte.Text = "0"
            TxtCmpbte.Locked = True
            VAR_DGRAL = "2"
        Case 2
            'GRABA FE (Facturacion Electronica)
            'DTPFechaCobro.Visible = False
            DTPFechaCobro.Visible = True
            TxtCmpbte.Enabled = True
            TxtCmpbte.Locked = False
            VAR_DGRAL = "0"
      End Select
      'TxtMonto.Text = CDbl(TxtDsctoTot)
      TxtMonto.SetFocus
    Else
        MsgBox "La Venta NO tiene saldo para procesar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
    End If
  Else
        MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

'Private Sub Cmd_Cliente_Click()
'    glPersNew = "P"
'    frmBeneficiario.Show 'vbModal
'End Sub


Private Sub BtnModDetalle2_Click()
'  If ado_datos14.Recordset.RecordCount > 0 Then
'    SSTab1.Tab = 2
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(1) = False
'
'    FrmEdita.Visible = True
''    BtnImprimir2.Visible = False
''    BtnImprimir3.Visible = False
'  Else
'    MsgBox "No Existen Bienes Registrados, Verifique por favor !! ", vbExclamation, "Atención!"
'  End If

    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'LISTADO DE FACTURACION' "
        CryF02.Formulas(2) = "subtitulo = 'MODULO COBRANZAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If

End Sub

Private Sub BtnDesAprobar_Click()
'  sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
'  If sino = vbYes Then
'    Dim rstdestino As New ADODB.Recordset
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " ", db, adOpenDynamic, adLockOptimistic
'    If Not rstdestino.BOF Then rstdestino.MoveFirst
'    If Not rstdestino.BOF And Not rstdestino.EOF Then
'      rstdestino("estado_codigo") = "REG"
'      rstdestino.Update
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
'    marca1 = Ado_datos.Recordset.Bookmark
'    Call OptFilGral1_Click
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo_fac = "REG" And Ado_datos.Recordset!factura_impresa = "N" Then
        Ado_datos.Recordset!estado_codigo_sol = "REG"
        Ado_datos.Recordset!estado_codigo_fac = "REG"
        Ado_datos.Recordset.Update
          'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
    Else
        MsgBox "No se puede DEVOLVER, el registro ya fue FACTURADO, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
 End If
End Sub

'Private Sub CmdDetallePoa_Click()
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'   MsgBox "No Existen Registros ", vbInformation, "Formulario 11"
'  Else
'    marca1 = Ado_datos.Recordset.BookMark
'    FrmPoasCapturaALB.Lblformulario = "F11"
'    FrmPoasCapturaALB.lblges_gestion = Ado_datos.Recordset!ges_gestion
'    FrmPoasCapturaALB.lblcodigo_unidad = Ado_datos.Recordset!codigo_unidad
'    FrmPoasCapturaALB.lblcodigo_solicitud = Ado_datos.Recordset!codigo_solicitud
'    FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'Ado_datos.Recordset!tipoben_codigo
'    FrmPoasCapturaALB.Show vbModal
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'    '
'  Else
'    Ado_datos.Refresh
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  End If
'End Sub

'Private Sub cmdElige_Click()
'  With ALFrmMateriales
'        .ALPrincipal
'        If .QResp Then
'            txtCodigo.Text = .QCodigo
'            txtDesc.Text = .QItem
'        End If
'    End With
'    Txtcant_alm = 0
'    Cant_Alm = 0
'    DE.dbo_albSacaDetalleMaterial Mid(txtCodigo, 3, 12), descri_bien, Cant_Alm
'    Txtcant_alm = Cant_Alm
'    If Cant_Alm >= TxtCantPedi Then
'        optSi = True
'    Else
'        optNo = True
'    End If
'End Sub

Private Sub Contabiliza_venta()
'    Call graba_proyecto
    If VAR_SW = 1 Then
        Call graba_ingreso
    End If
    'If VAR_SW = 1 Then
        Set rstdestino = New ADODB.Recordset
        If VAR_TIPOV = "L" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
        Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
        End If
        If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo

            'Modificar con CASE WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW MAY-2015
            If VAR_SW = 1 Then
                VAR_TSOL = VAR_TIPOS
            Else
                VAR_TSOL = rstdestino!solicitud_tipo
                VAR_TIPOS = rstdestino!solicitud_tipo
                VAR_PARTIDA = rstdestino!rubro_codigo
            End If
        Else

        End If
    'End If
  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  'If sino = vbYes Then
    ' INI CORRECCION 18-JUN-2014
    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

    fte_codigo1 = VAR_FTE
    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DEI", "DEY"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
        Case "REC"
            If VAR_MONEDA = "BOB" Then
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and SubCta_Deb1 = '01' ", db, adOpenKeyset, adLockReadOnly
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            Else
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  and SubCta_Deb1 = '02' ", db, adOpenKeyset, adLockReadOnly
            End If
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

            If VAR_JQ = "" Then
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
                'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
                If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
                  If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                    MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                    'JQA FEB-2016
                    'Exit Sub
                  End If
                End If
                If rs_aux1.State = 1 Then rs_aux1.Close
            End If
        Case "REF"
            If VAR_VTIPO = "L" Then     'Importación Directa
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " AND correlativo <> '6' ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_VTIPO = "V" Then     'Facturacion Local
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_VTIPO <> "L" And VAR_VTIPO <> "V" Then       'Mant, Rep, Inst, etc.
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_JQ = "" Then
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
                'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
                If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
                  If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                    MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                     'Exit Sub
                  End If
                End If
                If rs_aux1.State = 1 Then rs_aux1.Close
            End If
        Case "DYR"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DES"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "ANI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DVI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

            '' 02/07/2014 VERIFICAR
            'If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
            'If rstdestino2.State = 1 Then rstdestino2.Close
            'rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            'If rstdestino.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
            '  MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
            '  Exit Sub
            'End If
        Case Else
            MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
            If rstdestino.State = 1 Then rstdestino.Close
            Exit Sub
    End Select
    'If rstdestino.State = 1 Then rstdestino.Close
    '**** FIN VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************

    Dim cta_deb1 As String
    Dim Subcta_deb11 As String
    Dim Subcta_deb21 As String

    Dim cta_credito1 As String
    Dim Subcta_cred11 As String
    Dim Subcta_cred21 As String

    Dim cod_ant As Integer
    Dim org_ant As String

    v_Tipo_Comp(1, 1) = VAR_CODTIPO

    db.BeginTrans
'    Frmmensaje.Visible = True
'    LblMensaje.Caption = "Este proceso tomará solo unos segundos, gracias"
    '========================================
    '==== verifica si ya fue contabilizado
      yacontabilizo = 0
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '" & VAR_CODANT & "' and org_codigo = '" & VAR_ORG & "' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      'rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '2' and org_codigo = '111' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      If rs_aux2.RecordCount > 0 Then
        ' revisar para validar mejor si YA contabilizo !!
        'yacontabilizo = 1
        yacontabilizo = 0
      Else
        yacontabilizo = 0
      End If
      If yacontabilizo = 1 Then
        'MsgBox "aqui recontabilizar" & rstdestino!Cod_trans & " -- " & rstdestino!org_codigo & " / " & rstdestino!Cod_Comp
        Var_Comp = rs_aux2!Cod_Comp
      Else
        '===== ini GENERA EL CODIGO DE COMPROBANTE ====
        Set rstCodComp = New ADODB.Recordset
        rstCodComp.CursorLocation = adUseClient
        If rstCodComp.State = 1 Then rstCodComp.Close
        rstCodComp.Open "select * from fc_Correl where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
        If rstCodComp.RecordCount > 0 Then
          Var_Comp = CDbl(rstCodComp!numero_correlativo)
          Var_Comp = Var_Comp + 1
          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
          rstCodComp.Update
        End If
        If rstCodComp.State = 1 Then rstCodComp.Close

        'Correlativo por Mes y Tipo de Comprobante
        Set rs_aux14 = New ADODB.Recordset
        SQL_FOR = "select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & Trim(VAR_MES) & "' and cta_codigo2 = '" & VAR_DOC & "' ) "
        rs_aux14.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux14.RecordCount > 0 Then
              rs_aux14!numero_correlativo = rs_aux14!numero_correlativo + 1
              VAR_COMPM = rs_aux14!numero_correlativo    'VAR_DOCR
              rs_aux14.Update
        End If
        'R-112, R-110, R-111
         Set rs_aux14 = New ADODB.Recordset
          If rs_aux14.State = 1 Then rs_aux14.Close
          SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOC & "'  "  ''R-112' "          '  '" & txt_codigo1 & "' "
          rs_aux14.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
          If rs_aux14.RecordCount > 0 Then
                rs_aux14!correl_doc = rs_aux14!correl_doc + 1
                'VAR_COMPM = rs_aux14!correl_doc
                rs_aux14.Update
          End If
        '===== fin TERMINA GENERACION DE COMPROBANTE =====

      '==== ini registro co_comprobante_m

        rs_aux2.AddNew
        rs_aux2("cod_comp") = Var_Comp
      End If
    '========================================
    'anterior
    '      If rstdestino.State = 1 Then rstdestino.Close
    '      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
    '      If rstdestino.RecordCount > 0 Then
    '      End If
    '      rstdestino.AddNew

    '      rstdestino("cod_comp") = Var_Comp
    'anterior
      rs_aux2("Tipo_Comp") = VAR_CODTIPO        'v_Tipo_Comp(1, i)
      rs_aux2("cod_trans") = VAR_CODANT
      rs_aux2("org_codigo") = VAR_ORG
      'If yacontabilizo = 0 Then
      '  rs_aux2("Fecha_transacion") = Date
      'Else
        rs_aux2!Fecha_transacion = IIf(VAR_FFAC = "", Date, VAR_FFAC)
      'End If
      rs_aux2("mes_trasaccion") = UCase(MonthName(Month(Date)))
      rs_aux2("ges_gestion") = IIf(gestion0 = "", Year(Date), gestion0)  'glGestion
      rs_aux2("beneficiario_codigo") = VAR_BENEF
      rs_aux2("glosa") = VAR_TCOMP + "- " + VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4           'Ado_datos16.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = VAR_SOL         'Ado_datos16.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE
      'rs_aux2!Cobranza_aux = NRO_COBR
      rs_aux2("proceso_codigo") = Left(VAR_ETAPA, 3)        '"FIN"
      rs_aux2("subproceso_codigo") = Left(VAR_ETAPA, 6)     '"FIN-02"
      rs_aux2("etapa_codigo") = VAR_ETAPA

      rs_aux2("clasif_codigo") = "ADM"
      'rs_aux2("doc_codigo") = "R-112"
      rs_aux2("doc_codigo") = VAR_DOC       '"R-110" o "R-112"
      rs_aux2("doc_numero") = VAR_COMPM         'Var_Comp   VAR_COMPM
      rs_aux2("pro_codigo_det") = VAR_PROY2

      rs_aux2("estado_codigo") = "APR"
      rs_aux2("usr_codigo") = glusuario
      rs_aux2!venta_compra = nroventa
      rs_aux2!cobranza_pago = NRO_COBR
      'rs_aux2!Factura_cheque= NroFactura
      'If yacontabilizo = 0 Then
      '  rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
      '  'rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
      'Else
      If VAR_FFAC = "" Then
        VAR_FFAC = Date
      End If
        rs_aux2("Fecha_registro") = Format(VAR_FFAC, "dd/mm/yyyy")
      'End If
      rs_aux2.Update
      db.Execute "UPDATE co_comprobante_m SET edificio = gc_edificaciones.edif_descripcion FROM co_comprobante_m INNER JOIN gc_edificaciones ON co_comprobante_m.pro_codigo_det =gc_edificaciones.edif_codigo where co_comprobante_m.edificio Is Null "

      db.Execute "UPDATE co_comprobante_m SET cliente = gc_beneficiario.beneficiario_denominacion FROM co_comprobante_m INNER JOIN gc_beneficiario ON co_comprobante_m.beneficiario_codigo  =gc_beneficiario.beneficiario_codigo where co_comprobante_m.cliente Is Null "

      db.Execute "UPDATE co_comprobante_m SET departamento = gc_departamento.depto_descripcion FROM co_comprobante_m INNER JOIN gc_departamento ON LEFT(co_comprobante_m.pro_codigo_det,1)  =gc_departamento.depto_codigo where co_comprobante_m.departamento Is Null "

      If VAR_TCOMP = "REF" Then
        db.Execute "UPDATE co_comprobante_m SET glosa_contab = 'Fac: ' NroFactura + ' - '+ unidad_codigo + ' -Edif: ' + rtrim(edificio) + ' - Benef: ' + rtrim(cliente) + ' - ' + departamento + ' - ' + right(glosa,50) where co_comprobante_m.glosa_contab is null "
      Else
        db.Execute "UPDATE co_comprobante_m SET glosa_contab = unidad_codigo + ' -Edif: ' + rtrim(edificio) + ' - Benef: ' + rtrim(cliente) + ' - ' + departamento + ' - ' + right(glosa,50) where co_comprobante_m.glosa_contab is null "
      End If
      '==== fin registro co_comprobantre_m

    Dim d_cta_nombre_1 As String
    Dim d_aux1_1 As String
    Dim d_aux2_1 As String
    Dim d_aux3_1 As String
    Dim h_cta_nombre_1 As String
    Dim h_aux1_1 As String
    Dim h_aux2_1 As String
    Dim h_aux3_1 As String
    'If rstdestino.State = 1 Then rstdestino.Close

    For i = 1 To j
'    ' nuevo ini

      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")

        cta_credito1 = rstdestino("cta_cred")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")

        VAR_PORC = rstdestino!porcentaje
      Else
        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
        Subcta_deb11 = rstdestino!Subcta_cred1
        Subcta_deb21 = rstdestino!Subcta_cred2

        cta_credito1 = rstdestino!cta_deb
        Subcta_cred11 = rstdestino!Subcta_deb1
        Subcta_cred21 = rstdestino!Subcta_deb2

        VAR_PORC = rstdestino!porcentaje
      End If

      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        d_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
        VAR_DCORR = rs_aux1("correl")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        h_aux1_1 = rs_aux1("aux1")
        h_aux2_1 = rs_aux1("aux2")
        h_aux3_1 = rs_aux1("aux3")
        VAR_HCORR = rs_aux1("correl")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        VAR_NOMD = rs_aux1("NombreCta")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        VAR_NOMH = rs_aux1("NombreCta")
      End If
    ' nuevo fin

      '===== ini registra CO_diaRIO =========
      Set rstdestino2 = New ADODB.Recordset
      If rstdestino2.State = 1 Then rstdestino2.Close
      rstdestino2.Open "select * from co_diario where Cod_Comp = " & Var_Comp, db, adOpenKeyset, adLockOptimistic
      'If rstdestino2.RecordCount > 0 Then
      '  MsgBox "Ya Existe el asiento, se reemplazará con los nuevos datos..."
      'Else
        rstdestino2.AddNew
        rstdestino2("Cod_Comp") = Var_Comp
      'End If
        rstdestino2("Cod_Comp_Detalle") = rstdestino2.RecordCount
      'rstdestino2("Tipo_Comp") = "DEI"   'v_Tipo_Comp(1, i)
      'rstdestino2("Cod_Comp_C") = Var_Comp
      'If v_Tipo_Comp(1, i) = "DEI" Or v_Tipo_Comp(1, i) = "REC" Then
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_deb11
        rstdestino2("D_SubCta2") = Subcta_deb21
        rstdestino2("D_Aux1") = d_aux1_1
        rstdestino2("D_Aux2") = d_aux2_1
        rstdestino2("D_Aux3") = d_aux3_1
        rstdestino2("NOMCTADEBE") = VAR_NOMD
        rstdestino2("d_Correl") = VAR_DCORR
        ' para Aux1
        ' ini PARA EL FUTURO ******** REVISAR
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        SQL_FOR = "select * from cc_tipo_auxiliar where aux = '" & d_aux1_1 & "' "
'        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux4.RecordCount > 0 Then
'            Set rs_aux1 = New ADODB.Recordset
'            If rs_aux1.State = 1 Then rs_aux1.Close
'            SQL_FOR = "select * from " + rs_aux4!NombreTabla + " where " + rs_aux4!nombre_codigo + " = " + VAR_COD1
'            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'            If rs_aux1.RecordCount > 0 Then
'        Else
'        End If
        ' fin PARA EL FUTURO ******** REVISAR
        Select Case d_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(d_aux1_1, CStr(VAR_BENEF))    'DESAUX =
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                Call DESCAUX(d_aux1_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(d_aux1_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux1_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux1_1, rstdestino2!D_Cta_Aux1)
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                Call DESCAUX(d_aux1_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux1 = DESAUX
        Select Case d_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(d_aux2_1, CStr(VAR_BENEF))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                Call DESCAUX(d_aux2_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(d_aux2_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux2_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux2_1, rstdestino2!D_Cta_Aux2)
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                Call DESCAUX(d_aux2_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux2 = DESAUX
        Select Case d_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(d_aux3_1, CStr(VAR_BENEF))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                Call DESCAUX(d_aux3_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(d_aux3_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux3_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux3_1, rstdestino2!D_Cta_Aux3)
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                Call DESCAUX(d_aux3_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux3 = DESAUX

'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR

        'VAR_PORC Definido en el Relacionador
        If VAR_PORC = "0.87" Then
            rstdestino2("D_MontoBs") = VAR_87       'VAR_BS2 * VAR_PORC
            rstdestino2("D_MontoDl") = VAR_87 * GlTipoCambioOficial  'VAR_DOL2 * VAR_PORC
        End If
        If VAR_PORC = "0.13" Then
            rstdestino2("D_MontoBs") = VAR_13       'VAR_BS2 * VAR_PORC
            rstdestino2("D_MontoDl") = VAR_13 * GlTipoCambioOficial  'VAR_DOL2 * VAR_PORC
        End If
        If VAR_PORC <> "0.87" And VAR_PORC <> "0.13" Then
            rstdestino2("D_MontoBs") = VAR_BS2 * VAR_PORC
            rstdestino2("D_MontoDl") = VAR_DOL2 * VAR_PORC
        End If
        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        'AQUI MONEDA 02/07/01
        'rstdestino2("D_Cambio") = GlTipoCambioMercado
        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
        rstdestino2("H_Cuenta") = cta_credito1
        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_cred11
        rstdestino2("H_SubCta2") = Subcta_cred21
        rstdestino2("H_Aux1") = h_aux1_1
        rstdestino2("H_Aux2") = h_aux2_1
        rstdestino2("H_Aux3") = h_aux3_1
        rstdestino2("NOMCTAHABER") = VAR_NOMH
        rstdestino2("h_Correl") = VAR_HCORR
        'rstdestino2("H_Cta_Aux1") = ""
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(h_aux1_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                Call DESCAUX(h_aux1_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(h_aux1_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux1_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux1_1, rstdestino2!H_Cta_Aux1)
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                Call DESCAUX(h_aux1_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux1 = DESAUX

        Select Case h_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(h_aux2_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                Call DESCAUX(h_aux2_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(h_aux2_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux2_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux2_1, rstdestino2!H_Cta_Aux2)
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                Call DESCAUX(h_aux2_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux2 = DESAUX
        Select Case h_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(h_aux3_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                Call DESCAUX(h_aux3_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(h_aux3_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux3_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux3_1, rstdestino2!H_Cta_Aux3)
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                Call DESCAUX(h_aux3_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux3 = DESAUX

'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino2("H_MontoBs") = VAR_BS2 * VAR_PORC
        rstdestino2("H_MontoDl") = VAR_DOL2 * VAR_PORC
        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        rstdestino2!cobranza_pago = NRO_COBR
      End If

      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
      If (VAR_CODTIPO = "DES") Or (VAR_CODTIPO = "ANI") Or (VAR_CODTIPO = "DVI") Then
        'desafecta un devengado
        rstdestino2("D_Cuenta") = cta_credito1
        rstdestino2("D_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_cred11
        rstdestino2("D_SubCta2") = Subcta_cred21
        rstdestino2("D_Aux1") = h_aux1_1
        rstdestino2("D_Aux2") = h_aux2_1
        rstdestino2("D_Aux3") = h_aux3_1
'        rstdestino2("D_Cta_Aux1") = "VESCT"
        Select Case h_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = ""
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
        End Select

        Select Case h_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = ""
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
        End Select

        Select Case h_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = ""
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
        End Select
'        If h_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("D_Cambio") = GlTipoCambioMercado

        rstdestino2("H_Cuenta") = cta_deb1
        rstdestino2("H_Nombre") = d_cta_nombre_1  ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_deb11
        rstdestino2("H_SubCta2") = Subcta_deb21
        rstdestino2("H_Aux1") = d_aux1_1
        rstdestino2("H_Aux2") = d_aux2_1
        rstdestino2("H_Aux3") = d_aux3_1
'        rstdestino2("H_Cta_Aux1") = "VESCT"
        Select Case d_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = ""
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
        End Select

        Select Case d_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = ""
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
        End Select

        Select Case d_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = ""
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("H_Cambio") = GlTipoCambioMercado
        rstdestino2!cobranza_pago = NRO_COBR
      End If

'      '==== INI DVI ====
'      If (VAR_CODTIPO = "DVI") Then
'        rstdestino2("D_Cuenta") = cta_deb1
''        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_deb11
'        rstdestino2("D_SubCta2") = Subcta_deb21
'        rstdestino2("D_Aux1") = d_aux1_1
'        rstdestino2("D_Aux2") = d_aux2_1
'        rstdestino2("D_Aux3") = d_aux3_1
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("D_Cambio") = GlTipoCambioMercado
'        rstdestino2("H_Cuenta") = cta_credito1
''        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("H_SubCta1") = Subcta_cred11
'        rstdestino2("H_SubCta2") = Subcta_cred21
'        rstdestino2("H_Aux1") = h_aux1_1
'        rstdestino2("H_Aux2") = h_aux2_1
'        rstdestino2("H_Aux3") = h_aux3_1
'        'rstdestino2("H_Cta_Aux1") = "VESCT"
'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("H_Cambio") = GlTipoCambioMercado
'      End If
'      '==== FIN DVI ====
      rstdestino2("Usr_codigo") = glusuario
      'If yacontabilizo = 0 Then
      '  rstdestino2("Fecha_registro") = Date
      ''  rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
      'Else
        rstdestino2("Fecha_registro") = VAR_FFAC
      'End If

      rstdestino2.Update
      If rstdestino2.State = 1 Then rstdestino2.Close
      '======= fin registra co_diario ==========
      rstdestino.MoveNext
    Next i
    '======= inI Actualiza campos de estatus de ingresos ==========
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '" & correlativo1 & "' and org_codigo = '" & VAR_ORG & "' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' ", db, adOpenDynamic, adLockOptimistic
'    rstdestino.MoveFirst
'    If Not (rstdestino.EOF) Then
'      rstdestino("estado_aprobacion") = "S"
'        If VAR_CODTIPO = "DEI" Then
'          rstdestino("estado_devengado") = "S"
'        End If
'        If VAR_CODTIPO = "REC" Then
'          rstdestino("estado_recaudado") = "S"
'        End If
'        If VAR_CODTIPO = "DYR" Then
'          rstdestino("estado_devengado") = "S"
'          rstdestino("estado_recaudado") = "S"
'        End If
'
'        If VAR_CODTIPO = "DES" Then
'          rstdestino("estado_desafectado") = "S"
'        End If
'        If VAR_CODTIPO = "ANI" Then
'          rstdestino("estado_anulado") = "S"
'        End If
'        If VAR_CODTIPO = "DVI" Then
'          rstdestino!estado_desafectado = "S"
'          rstdestino!estado_anulado = "S"
'        End If
'       rstdestino.Update
'       If rstdestino.State = 1 Then rstdestino.Close
'    End If
    '======= fin Actualiza campos de estatus de ingresos ==========
    ' AAAAAAAAAQQQQQQQQQQQUUUUUUUUUUUIIIIIIIIIII
    cod_ant = 0
    org_ant = ""
    '======= ini Actualiza el monto recaudado  ==========
    If (VAR_CODTIPO = "REC") Then
      '      If rstdestino.State = 1 Then rstdestino.Close
      '      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      '      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
      '        cod_ant = rstdestino("ingreso_codigo_anterior")
      '        org_ant = rstdestino("org_codigo")
      '      End If
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") + VAR_DOL2
          rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") + VAR_BS2
          rstdestino.Update
      End If
      If rstdestino.State = 1 Then rstdestino.Close
    End If

    If (VAR_CODTIPO = "DES") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      Print VAR_CODANT
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
'        org_ant = rstdestino("org_codigo")
'      End If

      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "DEI" Then 'And VAR_CODTIPO = "DES"
'          rstdestino!estado_desafectado = "S" 02/07/01
          rstdestino!estado_codigo = "DES"
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        Else
          rstdestino("estado_codigo") = "DES"
'          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
          cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
          org_ant = rstdestino("org_codigo")
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
          'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & cod_ant & " and org_codigo = '" & org_ant & "' ", db, adOpenKeyset, adLockOptimistic
          rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
          If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
            rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
            rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") - VAR_BS2
          End If
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        End If
      End If
    End If

    If (VAR_CODTIPO = "ANI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "REC" Then
'          rstdestino("estado_desafectado") = ""
          rstdestino("estado_codigo") = "ANI"
'          rstdestino("estado_devengado") = "S" 02/07/01
'          rstdestino("estado_anulado") = ""
'          rstdestino("codigo_tipo") = "DEI" 02/07/01
          rstdestino("monto_recaudado_dolares") = 0
        End If
      End If
      rstdestino.Update
'      Print rstdestino!ingreso_codigo_anterior
'      Print rstdestino!monto_recaudado
      cod_ant = 0
      org_ant = ""

      'Call f_actual_rec(rstdestino!org_codigo, rstdestino!ingreso_codigo_anterior)
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    If (VAR_CODTIPO = "DVI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        rstdestino!estado_codigo = "DVI"
      End If
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    '======= fin Actualiza el monto recaudado  ==========

    '======= ini Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    If VAR_CODTIPO = "REC" Or VAR_CODTIPO = "DYR" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    If VAR_CODTIPO = "ANI" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    '======= fin Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    'LblMensaje.Caption = "El proceso concluyó exitosamente, gracias"
    'Frmmensaje.Visible = False
    db.CommitTrans
  'End If
'  'marca1 = Ado_datos.Recordset.Bookmark
'  rs_datos.Update
'  rs_datos.Requery
'  Set Ado_datos.Recordset = rs_datos
'  If rs_datos.RecordCount > 0 Then
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"

End Sub

Private Function DESCAUX(VARAUX As String, VARCODIG As String)
    Set rsAuxDetalle = New ADODB.Recordset
    If rsAuxDetalle.State = 1 Then rsAuxDetalle.Close
    Select Case VARAUX
        Case "01"
            rsAuxDetalle.Open "SELECT beneficiario_denominacion AS DESAUX2 FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT beneficiario_denominacion AS DESAUX FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' "
        Case "02"
            rsAuxDetalle.Open "SELECT cta_descripcion AS DESAUX2 FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT cta_descripcion AS DESAUX FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "' "
        Case "03"
            rsAuxDetalle.Open "SELECT pro_codigo_det_descripcion AS DESAUX2 FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT pro_codigo_det_descripcion AS DESAUX FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "' "
        Case "04"
            rsAuxDetalle.Open "SELECT unidad_descripcion AS DESAUX2 FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "05"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "06"
            rsAuxDetalle.Open "SELECT depto_descripcion AS DESAUX2 FROM gc_departamento where depto_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT depto_descripcion AS DESAUX FROM gc_departamento where depto_codigo = '" & VARCODIG & "' "
        Case "07"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "08"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "09"
            rsAuxDetalle.Open "SELECT Org_descripcion AS DESAUX2 FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT Org_descripcion AS DESAUX FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' "
        Case "10"
            'db.Execute "SELECT impuesto_descripcion AS DESAUX FROM fc_impuestos where impuesto_codigo = '" & VARCODIG & "' "
        Case "11"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "12"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "00"
            DESAUX = ""
    End Select
    If rsAuxDetalle.RecordCount > 0 Then
      DESAUX = RTrim(rsAuxDetalle!DESAUX2)
    Else
      DESAUX = ""
    End If
End Function

'Private Sub f_actual_rec(org, codant)
'  Dim acumDl As Double
'  Dim rsrecalc As New ADODB.Recordset
'  Set rsrecalc = New ADODB.Recordset
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select sum(monto_dolares) as acumDl from fo_ingresos_cabecera where org_codigo = '" & org & "' and  correlativo_anterior = '" & codant & "' and codigo_tipo = 'REC' and estado_recaudado= 'S'", db, adOpenKeyset, adLockReadOnly
'  If rsrecalc.RecordCount > 0 Then
'    acumDl = IIf(IsNull(rsrecalc!acumDl), 0, rsrecalc!acumDl)
'  Else
'    acumDl = 0
'  End If
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select * from fo_ingresos_cabecera where org_codigo = '" & org & "' and correlativo_ingreso = '" & codant & "' ", db, adOpenKeyset, adLockOptimistic
'  If rsrecalc.RecordCount > 0 Then
'    rsrecalc!monto_recaudado_dolares = acumDl
'  End If
'  rsrecalc.Update
'  If rsrecalc.State = 1 Then rsrecalc.Close
'
'End Sub

Private Sub graba_proyecto()
'    Select Case Ado_datos.Recordset!unidad_codigo
'        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
'            VAR_PROY = 12
'        Case "UCOM"
'            VAR_PROY = 17
'        Case "DVTA"
'            VAR_PROY = 18
'
'    End Select
'
'    Set rs_aux1 = New ADODB.Recordset
'    If rs_aux1.State = 1 Then rs_aux1.Close
'    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    Else
'        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
'           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & Ado_datos.Recordset!ges_gestion & ", 'APR', '" & GlUsuario & "', '" & Date & "')"
'    End If
End Sub

Private Sub graba_ingreso()
    '======= Ini grabado de datos
   'swgraba = 0
   'Call valida

'   If swgraba = 1 Then
'      FraOpciones2.Visible = False
'      fraOpciones.Visible = True
'      FraIngresosNav.Enabled = True
'      FraIngresosDat.Enabled = False

      'If v_añadir = 1 Then
        'EFECTIVO o a CREDITO
         'db.BeginTrans
         'Call add_correl
         Set rstdestino = New ADODB.Recordset
         If VAR_TIPOV = "L" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
         Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
         End If
         VAR_CODANT = 0
         If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo
            VAR_TIPOS = rstdestino!solicitud_tipo
            VAR_PARTIDA = rstdestino!rubro_codigo

         Else
             Select Case VAR_COD4
                 Case "DVTA", "DCOMB", "DCOMC", "DCOMS"             'INI COMERCIAL
                     VAR_ORG = "111"
                     VAR_FTE = "10"
                     VAR_TIPOS = 3
                     EST_PROG = 18      'Activ=17, Proy=18
                     If VAR_TIPOV = "L" Then
                         VAR_PARTIDA = "11360"
                     Else
                         VAR_PARTIDA = "11310"
                     End If
                 Case "COMEX"            'INI COMEX
                     VAR_ORG = "111"
                     VAR_FTE = "10"
                     EST_PROG = 15      'Activ=14, Proy=15
                     VAR_PARTIDA = "11310"
                 Case "DNMAN", "DMANB", "DMANC", "DMANS"            'INI MANTENIMIENTO
                     VAR_ORG = "112"
                     VAR_FTE = "10"
                     VAR_TIPOS = 10
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11320"
                 Case "DNREP", "DNEME", "DREPB", "DREPC", "DREPS"            'INI REPARACIONES 'INI EMERGENCIAS
                     VAR_ORG = "113"
                     VAR_FTE = "10"
                     VAR_TIPOS = 7
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11330"
                 Case "DNMOD"            'INI MODERNIZACION
                     VAR_ORG = "114"
                     VAR_FTE = "10"
                     VAR_TIPOS = 9
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11340"
                 Case "DNINS", "DNAJS", "DINSB", "DINSC", "DINSS"            'INI INSTALACIONES    'INI AJUSTE
                     VAR_ORG = "111"
                     VAR_FTE = "10"
                     VAR_TIPOS = 4
                     EST_PROG = 18      'Activ=17, Proy=18
                     VAR_PARTIDA = "11350"
                 Case "DCONT"            'INI EMERGENCIAS
                     VAR_ORG = "112"
                     VAR_FTE = "10"
                     VAR_TIPOS = 10
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11320"
                 Case Else               'INI COMPRAS
                     VAR_ORG = "311"
                     VAR_FTE = "30"
                     VAR_TIPOS = 6
                     EST_PROG = 18      'Activ=17, Proy=18
                     VAR_PARTIDA = "11320"
            End Select
            'Call add_correl
            'EXEPCION PARA GRABAR CONTRATO EN INGRESOS
'             rstdestino.AddNew
'             rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
'             rstdestino("ingreso_codigo") = correlativo1
'             rstdestino("org_codigo") = VAR_ORG
'             If VAR_CODANT = 0 Then
'                VAR_CODANT = correlativo1
'             End If
'             rstdestino("ingreso_codigo_anterior") = VAR_CODANT
'
'             rstdestino("proceso_codigo") = "FIN"
'             rstdestino("subproceso_codigo") = "FIN-01"
'             rstdestino("etapa_codigo") = "FIN-01-02"
'             rstdestino("clasif_codigo") = "ADM"
'             rstdestino("doc_codigo") = "R-110"
'             rstdestino("doc_numero") = correlativo1
'             rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos16.Recordset("unidad_codigo")
'             rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos16.Recordset("solicitud_codigo")
'             '
'             rstdestino("solicitud_tipo") = VAR_TIPOS
'
'             If VAR_COD4 = "DVTA" Then
'                rstdestino("tipo_comp") = "DEY"
'                rstdestino("Codigo_tipo") = "DEY"
'             Else
'                rstdestino("tipo_comp") = "DEI"
'                rstdestino("Codigo_tipo") = "DEI"
'             End If
'             'OJO JQA
'             rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
'             rstdestino("fecha_ingreso") = Date
'             rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
'             rstdestino("tipo_moneda") = VAR_MONEDA
'             'VAR_MONEDA = "BOB"
'             rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA       'Ado_datos.Recordset("cobranza_observaciones")
'             'CAMBIAR FTE
'             rstdestino("fte_codigo") = VAR_FTE
'             'CAMBIAR RUBROS
'             rstdestino("rubro_codigo") = VAR_PARTIDA
'             'CAMBIAR RUBROS
'             rstdestino("cheque_o_trf") = "T"
'             'CAMBIAR CTA
'             rstdestino("cta_codigo") = VAR_CTA
'             If VAR_CTA = "NN" Then
'                rstdestino("Bco_codigo") = "BCP"
'             Else
'                rstdestino("Bco_codigo") = "BMS"
'             End If
'             'CAMBIAR CTA
'             rstdestino("numero_documento") = VAR_COD1
'             rstdestino("unidad_codigo_ant") = VAR_CITE
'             rstdestino("monto_dolares") = VAR_DOL2 * 12
'             rstdestino("monto_bolivianos") = VAR_BS2 * 12
'             rstdestino("monto_recaudado_dolares") = VAR_DOL2 * 12 'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
'             rstdestino("monto_recaudado_bolivianos") = VAR_BS2 * 12   'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
'             rstdestino("convenio_codigo") = "NN"
'             rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
'             rstdestino("estado_CODIGO") = "APR"
'             'rstdestino("estado_codigo_dr") = "DEI"
'
'             rstdestino("usr_CODIGO") = glusuario
'             rstdestino("fecha_registro") = Date
'             rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
'
'             rstdestino.Update
'             VAR_CODANT = rstdestino!ingreso_codigo
'             VAR_ORG = rstdestino!org_codigo
'             VAR_FTE = rstdestino!fte_codigo
'             If rstdestino.State = 1 Then rstdestino.Close
'             If VAR_TIPOV = "L" Then
'                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
'             Else
'                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
'             End If
         End If
         Call add_correl
         ' OJO CAMBIA FINANCIADOR WWWWWWWWWWWWWWWWWWWWW
         rstdestino.AddNew
         rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
         rstdestino("ingreso_codigo") = correlativo1
         rstdestino("org_codigo") = VAR_ORG
         If VAR_CODANT = 0 Then
            VAR_CODANT = correlativo1
         End If
         rstdestino("ingreso_codigo_anterior") = VAR_CODANT
         rstdestino("Codigo_tipo") = VAR_CODTIPO
         rstdestino("proceso_codigo") = "FIN"
         rstdestino("subproceso_codigo") = "FIN-02"
         rstdestino("etapa_codigo") = VAR_ETAPA
         rstdestino("clasif_codigo") = "ADM"
         rstdestino("doc_codigo") = VAR_DOC
         rstdestino("doc_numero") = correlativo1
         rstdestino("unidad_codigo") = VAR_COD4
         rstdestino("solicitud_codigo") = VAR_SOL
         rstdestino("solicitud_tipo") = VAR_TIPOS
         'OJO JQA
         rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
         rstdestino("fecha_ingreso") = Date
         rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
         rstdestino("tipo_moneda") = VAR_MONEDA
         'VAR_MONEDA = "BOB"
         rstdestino("ingreso_concepto") = VAR_TCOMP + ": " + VAR_GLOSA      'Ado_datos.Recordset("cobranza_observaciones")
         'VAR_GLOSA = "INGRESO POR: " + Ado_datos.Recordset("cobranza_observaciones")
         If VAR_TIPOV = "E" Then
            rstdestino("tipo_comp") = "DYR"
         Else
            rstdestino("tipo_comp") = VAR_CODTIPO
         End If
         rstdestino("fte_codigo") = VAR_FTE
         rstdestino("rubro_codigo") = VAR_PARTIDA
         rstdestino("cheque_o_trf") = "T"
         'CAMBIAR CTA
         rstdestino("cta_codigo") = VAR_CTA
         If VAR_CTA = "2015046557-03-054" Then
            rstdestino("Bco_codigo") = "BCP"
         Else
            rstdestino("Bco_codigo") = "BMS"
         End If
         'CAMBIAR CTA
         NroFactura = Trim(Str(VAR_COD1))
         rstdestino("numero_documento") = NroFactura        'Ado_datos.Recordset!cobranza_nro_factura
         rstdestino("unidad_codigo_ant") = VAR_CITE
         rstdestino("monto_dolares") = VAR_DOL2
         rstdestino("monto_bolivianos") = VAR_BS2
         rstdestino("monto_recaudado_dolares") = VAR_DOL2   'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
         rstdestino("monto_recaudado_bolivianos") = VAR_BS2     'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
         rstdestino("convenio_codigo") = "NN"
         rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
         rstdestino("estado_CODIGO") = "APR"
         'rstdestino("estado_codigo_dr") = "DEI"

         rstdestino("usr_CODIGO") = glusuario
         rstdestino("fecha_registro") = Date
         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")

         rstdestino.Update
         If rstdestino.State = 1 Then rstdestino.Close
        'db.CommitTrans

'          If rstIngresos.State = 1 Then rstIngresos.Close
'          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'          rstIngresos.Sort = "ingreso_codigo"
'          rstIngresos.Requery

'          rstIngresos.Requery
'          Set AdoIngresos.Recordset = rstIngresos
'          AdoIngresos.Refresh
'          AdoIngresos.Recordset.Find "ultimo = 'S'"
'          If Not (AdoIngresos.Recordset.EOF) Then
'            marca1 = AdoIngresos.Recordset.Bookmark
'            AdoIngresos.Recordset("ultimo") = "N"
'            AdoIngresos.Recordset.Update
'          End If

'          AdoIngresos.Recordset.Move marca1 - 1

'          marca1 = 0
      'End If
'   Else*
'      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
''      FraOpciones2.Visible = False
''      FraOpciones.Visible = True
''      FraIngresosNav.Enabled = True
''      FraIngresosDat.Enabled = False
''      AdoIngresos.Refresh
'   End If
'   LblAccion = ""
'AAQQQQQUIIIIIIIIII    JQA

End Sub

Private Sub add_correl()
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' ", db, adOpenDynamic, adLockOptimistic
  If rstcorrel_ing.RecordCount = 0 Then
     VAR_ORG = "112"
     VAR_FTE = "10"
     If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
     rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "'  ", db, adOpenDynamic, adLockOptimistic
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = VAR_ORG   'Trim(DtCorg_codigo.Text)
'     rstcorrel_ing("ges_gestion") = Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
'     rstcorrel_ing("fte_codigo") = "10"
'     'rstcorrel_ing("correlativo") = 1
'     rstcorrel_ing("correlativo_ingreso") = 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
  Else
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update

     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
  End If
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close

End Sub

'Private Sub CmdGrabaCobranza()
'    If swnuevo = 1 Then
''      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
''      Set Ado_datos16.Recordset = rstdestino
''      Ado_datos16.Recordset.AddNew
'      Ado_datos16.Recordset!correl_venta = Val(lblcorrelVenta.Caption)
'      Ado_datos16.Recordset!venta_codigo = Val(TxtNroVenta.Text)
'      Ado_datos16.Recordset!ges_gestion = Year(Date)    'Trim(LblGestion.Caption)
'    End If
'      Ado_datos16.Recordset!beneficiario_codigo = dtc_codigo2A.Text                                 'Codigo Beneficiario/Cliente
'      Ado_datos16.Recordset!ci = dtc_codigo4A.Text                                                     'Codigo Cobrador
'      Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text + " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
'      Ado_datos16.Recordset!deuda_cobrada = Val(TxtMonto.Text)                                  'Monto Cobrado
'      Ado_datos16.Recordset!deuda_cobrada_dol = Val(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
'      Ado_datos16.Recordset!fecha_cobranza = DTPFechaCobro.Value                                'Fecha de Cobranza
'      'Call acumulaMont(Ado_datos16.Recordset!ges_gestion, Ado_datos16.Recordset!correl_venta, Ado_datos16.Recordset!venta_codigo)
'      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))
'
'      Ado_datos16.Recordset!obs_cobranza = TxtObs
'      Ado_datos16.Recordset!nro_cmpbte = Trim(TxtCmpbte)
'      Ado_datos16.Recordset!usr_usuario = GlUsuario
'      Ado_datos16.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'      Ado_datos16.Recordset.Update
'End Sub

'Private Sub CmdModDetalle_Click()
'  FraDetalle.Visible = True
'  FraDetalle.Enabled = True
'  txtnosolicitud1.Enabled = False
'  txtcorrdet.Enabled = False
'  dtccodpar.SetFocus
'  CmdGraDetalle.Enabled = True
'  CmdAddDetalle.Enabled = False
'  CmdModDetalle.Enabled = False
'  CmdSalDetalle.Enabled = False
'  CmdCanDetalle.Enabled = True
'  swgrabar = 2
'End Sub

'Private Sub CmdGraDetalle_Click()
'    If swgrabar = 1 Then
'        Dim rstdestino As New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle_correl where formulario = '" & "F11" & "' and correl_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("correl_solicitud_detalle") = rstdestino("correl_solicitud_detalle") + 1
'        Else
'            rstdestino.AddNew
'            rstdestino("formulario") = "F11"
'            rstdestino("correl_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correl_solicitud_detalle") = 1
'        End If
'        correldetalle = rstdestino("correl_solicitud_detalle")
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correlativo_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        rstdestino.AddNew
'        rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'        rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'        rstdestino("correlativo_detalle") = correldetalle
'        rstdestino("Par_codigo") = dtccodpar.Text
'        rstdestino("Importe_nacional") = txtsolpeso.Text
'        rstdestino("formulario") = "F11"
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    If swgrabar = 2 Then
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoDetalleSolicitud.Recordset("ges_gestion") & "' and correlativo_solicitud = " & adoDetalleSolicitud.Recordset("correlativo_solicitud") & " and correlativo_detalle =" & adoDetalleSolicitud.Recordset("correlativo_detalle"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'            rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correlativo_detalle") = correldetalle
'            rstdestino("Par_codigo") = dtccodpar.Text
'            rstdestino("Importe_nacional") = txtsolpeso.Text
'            rstdestino("formulario") = "F11"
'            rstdestino.Update
'        End If
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    CmdGraDetalle.Enabled = False
'    CmdAddDetalle.Enabled = True
'    CmdModDetalle.Enabled = True
'    CmdSalDetalle.Enabled = True
'    CmdCanDetalle.Enabled = False
'    FraDetalle.Enabled = False
'    swgrabar = 0
'End Sub

Private Sub CmdNOunidad_Click()
    swunidad = 0
    Frmunidad.Visible = False
End Sub

Private Sub CmdOKunidad_Click()
    swunidad = 1
        If swunidad = 1 Then
            Dim rstpagos As New ADODB.Recordset
            Set rstpagos = New ADODB.Recordset
            If rstpagos.State = 1 Then rstpagos.Close
            rstpagos.Open "select * from pagos where GES_gestion = '5000'", db, adOpenKeyset, adLockOptimistic
            rstpagos.AddNew
                rstpagos("ges_gestion") = Ado_datos.Recordset("ges_gestion")
                rstpagos("org_codigo") = DataCombo1.Text   'Ado_datos.Recordset("formulario")
                rstpagos("codigo_pago") = "" 'genera jorge
                rstpagos("codigo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
                rstpagos("formulario") = Ado_datos.Recordset("formulario")
                rstpagos("codigo_unidad") = Ado_datos.Recordset("codigo_unidad")
                rstpagos("monto_bolivianos") = Ado_datos.Recordset("monto_bolivianos")
                rstpagos("estado_compromiso") = "N"
                rstpagos("justificacion") = Ado_datos.Recordset("justificacion_solicitud")
             rstpagos.Update
        End If
End Sub

Private Sub CmdGrabaCobro_Click()
End Sub


'Private Sub BtnImprimir2_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'    CryR01.WindowShowRefreshBtn = True
''    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
'    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
'
'    CryR01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
'    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'    iResult = CryR01.PrintReport
'    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'End Sub

Private Sub cmd_benef_Click()
    Set rs_datos8 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_beneficiario where tipoben_codigo <> '0' and tipoben_codigo <> '1' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    If Ado_datos8.Recordset.RecordCount > 0 Then
        dtc_desc8.BoundText = dtc_codigo8.BoundText
        FraGrabarCancelar.Enabled = False
        frm_benef.Visible = True
    End If
End Sub

Private Sub cmd_moneda1_LostFocus()
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria where tipo_moneda = '" & cmd_moneda1.Text & "' ", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
'    dtc_ctades.BoundText = dtc_cta.BoundText
End Sub

Private Sub CmdFoto_Click()
'    Frm_Imprime_Factura.Show

    On Error GoTo QError
    Set fs = New FileSystemObject   'Creamos la Nueva referencia Fso

    Set rs_aux6 = New ADODB.Recordset     'Iniciales del Cliente - gc_beneficiario
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos1.Recordset!beneficiario_codigo_fac & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        'db.Execute "update ao_ventas_cobranza set beneficiario_iniciales = '" & rs_aux6!beneficiario_iniciales & "'   Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
        db.Execute "update ao_ventas_cobranza set beneficiario_iniciales = '" & Left(rs_aux6!beneficiario_iniciales, 4) & "' Where venta_codigo_new = " & Ado_datos1.Recordset!IdFactura & " "
    End If
    Dim VAR_FACQR As Long
    Set rs_aux11 = New ADODB.Recordset     'Iniciales del Cliente - gc_beneficiario
    If rs_aux11.State = 1 Then rs_aux11.Close
    rs_aux11.Open "Select * from ao_ventas_cobranza_fac_QR where IdFactura = " & Ado_datos1.Recordset!IdFactura & " ", db, adOpenStatic
    If rs_aux11.RecordCount > 0 Then
        VAR_FACQR = rs_aux11!IdFactura
    End If
'    'If Ado_datos.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
'    'If Ado_datos1.Recordset!archivo_foto_cargado = "N" Or IsNull(Ado_datos1.Recordset!archivo_foto_cargado) Then
'    If rs_aux11!archivo_foto_cargado = "N" Or IsNull(rs_aux11!archivo_foto_cargado) Then
      NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos1.Recordset!beneficiario_codigo_fac) & "\"
      DirOrigen = App.Path & "\CLIENTES\"
      DirDestino = App.Path & "\CLIENTES\"
      'DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos1.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG"       'Ado_datos.Recordset!cobranza_nro_factura        'ARCHIVO_Foto
      'fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos1.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG"       'Ado_datos.Recordset!cobranza_nro_factura        'ARCHIVO_Foto
      'Ado_datos1.Recordset!archivo_foto_cargado = "S"
      VAR_ARCHIVO = Trim(Ado_datos1.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG")
      db.Execute "UPDATE ao_ventas_cobranza_fac_QR SET archivo_foto_cargado= 'S', ARCHIVO_Foto= '" & VAR_ARCHIVO & "', estado_codigo= 'APR', usr_codigo = '" & glusuario & "', fecha_registro = '" & CDate(Date) & "' WHERE IdFactura= " & VAR_FACQR & "  "
                
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "Q_R"
''      If GlServidor = "SERVIDOR2" Then
''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
''      Else
'         e = NombreCarpeta
''      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          'NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
'          NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos1.Recordset!beneficiario_codigo_fac) & "\"
'          DirOrigen = App.Path & "\CLIENTES\"
'          DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
'          'fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos.Recordset!ARCHIVO_Foto
'          fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & rs_aux6!ARCHIVO_Foto
'          'frmBeneficiario_Admin.Adolista.Recordset!archivo_foto_cargado = "S"
'          db.Execute "UPDATE ao_ventas_cobranza_fac_QR SET archivo_foto_cargado= 'S', ARCHIVO_Foto= '" & VAR_ARCHIVO & "', estado_codigo= 'APR', usr_codigo = '" & glusuario & "', fecha_registro = '" & CDate(Date) & "'  "
'    '      Frmexporta.DirDestino.Path = NombreCarpeta
'    '      GlArch = "Q_R"
'    ''      If GlServidor = "SERVIDOR2" Then
'    ''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
'    ''      Else
'    '         e = NombreCarpeta
'    ''      End If
'    '      Frmexporta.DirDestino2.Path = e
'    '      Frmexporta.Show vbModal      End If
'      End If
'    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SERVIDOR2" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" + Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        'ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(rs_aux6!beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
        'ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(Ado_datos1.Recordset!ARCHIVO_Foto)
        ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(VAR_ARCHIVO)
        
'    End If
    'ARCH_FOTO = App.Path + "\" + "CLIENTES" + "\" + Ado_datos.Recordset!beneficiario_codigo + "\" + Ado_datos.Recordset("beneficiario_codigo") + "-FOTO.JPG"
    'CodBenef = Ado_datos.Recordset!cobranza_codigo
    CodBenef = Ado_datos1.Recordset!IdFactura
    'If Guardar_Imagen(db, "Select Foto From ao_ventas_cobranza_fac Where IdFactura= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
    If Guardar_Imagen(db, "Select Foto From ao_ventas_cobranza_fac_QR Where IdFactura= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
        'MsgBox "Se cargo la Imagen Correctamente !!"
        'Exit Sub
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
    If Guardar_Imagen(db, "Select Foto From ao_ventas_cobranza_fac_QR Where IdFactura= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
        'MsgBox "Se cargo la Imagen Correctamente !!"
        MsgBox "Se Emitió Correctamente la Factura Nro. " + Str(CDbl(VARFactura2))
        Exit Sub
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    db.RollbackTrans

    'Screen.MousePointer = vbDefault
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub BntImprimir3_Click()
    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas_dol.rpt"
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If
End Sub


Private Sub dtc_aux8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_aux8.BoundText
    dtc_codigo8.BoundText = dtc_aux8.BoundText
End Sub

Private Sub dtc_codigo4A1_Click(Area As Integer)
    dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''    dtc_aux5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    dtc_aux8.BoundText = dtc_codigo8.BoundText
End Sub

'Private Sub dtc_cta_Click(Area As Integer)
'    dtc_ctades.BoundText = dtc_cta.BoundText
'End Sub

'Private Sub dtc_ctades_Click(Area As Integer)
'    dtc_cta.BoundText = dtc_ctades.BoundText
'End Sub

Private Sub dtc_desc4A1_Click(Area As Integer)
    dtc_codigo4A1.BoundText = dtc_desc4A1.BoundText
End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'    dtc_codigo6.BoundText = dtc_desc6.BoundText
'End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    dtc_aux8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
End Sub

'Private Sub DataCombo1_Click(Area As Integer)
'    DataCombo2.Text = DataCombo1.BoundText
'End Sub

'Private Sub DataCombo2_Click(Area As Integer)
'    DataCombo1.Text = DataCombo2.BoundText
'End Sub

Private Sub dtccodmanejo_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodmanejo.BoundText
    DtCDescripcion.BoundText = dtccodmanejo.BoundText
    dtcunidadmedida.BoundText = dtccodmanejo.BoundText
    dtccodpeso.BoundText = dtccodmanejo.BoundText
End Sub

Private Sub dtccodpeso_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodpeso.BoundText
    DtCDescripcion.BoundText = dtccodpeso.BoundText
    dtcunidadmedida.BoundText = dtccodpeso.BoundText
    dtccodmanejo.BoundText = dtccodpeso.BoundText
End Sub


Private Sub dtccodpar_Click(Area As Integer)
    dtcdescripar.Text = dtccodpar.BoundText
End Sub

Private Sub dtccodpoa_Click(Area As Integer)
    dtcdespoa.Text = dtccodpoa.BoundText
End Sub

Private Sub dtccodpuesto_Click(Area As Integer)
    dtcdenopuesto.Text = dtccodpuesto.BoundText
End Sub

Private Sub dtccodtipoid_Click(Area As Integer)
    dtcdescrtipoid.BoundText = dtccodtipoid.BoundText
End Sub

Private Sub dtccoduni_Click(Area As Integer)
    dtcdescripuni.Text = dtccoduni.BoundText
End Sub

Private Sub dtccorrcompromiso_Click(Area As Integer)
    dtcfechacompromiso.BoundText = dtccorrcompromiso.BoundText
End Sub

Private Sub dtccorrsol_Click(Area As Integer)
 dtcfechasol.BoundText = dtccorrsol.BoundText
End Sub

Private Sub dtcdenominacionruc_Click(Area As Integer)
    dtcnroruc.BoundText = dtcdenominacionruc.BoundText
End Sub

Private Sub dtcdenopuesto_Click(Area As Integer)
    dtccodpuesto.Text = dtcdenopuesto.BoundText
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
    DtCCodigo.BoundText = DtCDescripcion.BoundText
    dtcunidadmedida.BoundText = DtCDescripcion.BoundText
    dtccodmanejo.BoundText = DtCDescripcion.BoundText
    dtccodpeso.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub dtcdescripar_Click(Area As Integer)
    dtccodpar.Text = dtcdescripar.BoundText
End Sub

Private Sub dtcdescripuni_Click(Area As Integer)
    dtccoduni.Text = dtcdescripuni.BoundText
End Sub

Private Sub dtcdescrtipoid_Click(Area As Integer)
    dtccodtipoid.BoundText = dtcdescrtipoid.BoundText
End Sub

Private Sub dtcfechacompromiso_Click(Area As Integer)
    dtccorrcompromiso.BoundText = dtcfechacompromiso.BoundText
End Sub

Private Sub dtcfechasol_Click(Area As Integer)
    dtccorrsol.BoundText = dtcfechasol.BoundText
End Sub

Private Sub dtcnroruc_Click(Area As Integer)
    dtcdenominacionruc.Text = dtcnroruc.BoundText
End Sub


Private Sub dtc_desc4A_Click(Area As Integer)
    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
End Sub

Private Sub dtctipodoc_Click(Area As Integer)
    dtcdenodoc.Text = dtctipodoc.BoundText
End Sub

Private Sub dtcunidadmedida_Click(Area As Integer)
    DtCCodigo.BoundText = dtcunidadmedida.BoundText
    DtCDescripcion.BoundText = dtcunidadmedida.BoundText
    dtccodmanejo.BoundText = dtcunidadmedida.BoundText
    dtccodpeso.BoundText = dtcunidadmedida.BoundText
End Sub

Private Sub dtcdespoa_Click(Area As Integer)
    dtccodpoa.Text = dtcdespoa.BoundText
End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
''    dtc_codigo5.BoundText = dtc_desc5.BoundText
''    dtc_aux5.BoundText = dtc_desc5.BoundText
'End Sub

Private Sub Form_Load()
'On Error GoTo QError2
    swnuevo = 0
    VAR_SW = 0
    parametro = Aux
    VAR_JQ = ""
    VAR_DGRAL = "1"
    'BtnVer
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
'    FrmCabecera.Enabled = False
    FrmCobros.Enabled = False
'    FrmCobros1.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
'    FraGrabarCancelar1.Visible = False
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
'    BtnImprimir2.Visible = True
    If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "FDELGADILLO" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "HMARIN" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
    Else
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
    End If
'    FrmEdita.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
    FraNavega.Caption = lbl_titulo.Caption
    FrmCobros.Visible = False
    'lbl_titulo2.Caption = lbl_titulo.Caption
    'lbl_titulo1.Caption = lbl_titulo.Caption
    'db.Execute "UPDATE ao_ventas_cobranza_fac SET ao_ventas_cobranza_fac.edif_codigo_corto  = ao_ventas_cabecera.edif_codigo_corto FROM ao_ventas_cobranza_fac INNER JOIN ao_ventas_cabecera ON ao_ventas_cobranza_fac.venta_codigo = ao_ventas_cabecera.venta_codigo WHERE ao_ventas_cobranza_fac.edif_codigo_corto IS NULL "
    
    'db.Execute "UPDATE ao_ventas_cobranza_fac SET ao_ventas_cobranza_fac.beneficiario_email  = gc_beneficiario.beneficiario_email FROM ao_ventas_cobranza_fac INNER JOIN gc_beneficiario ON ao_ventas_cobranza_fac.beneficiario_codigo_fac = gc_beneficiario.beneficiario_codigo where ao_ventas_cobranza_fac.beneficiario_email Is Null "
    
'QError2:
'    ' Manejo de errores
'    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    'Set Ado_datos1.Recordset = rs_datos1
    'dtc_desc1.BoundText = dtc_codigo1.BoundText

    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
'    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText

    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText

    Set rs_datos4 = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador en Fac.
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    'rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText

    Set rs_datos4A = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador
    If rs_datos4A.State = 1 Then rs_datos4A.Close
    'rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario ", db, adOpenStatic  '4333735
    'rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set ado_datos4A.Recordset = rs_datos4A
'    dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText

    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_tipo_transaccion order by trans_descripcion", db, adOpenStatic
    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText

    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "ac_tipo_compra_venta", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText

    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh

    'Solo para Equipos (*)
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    'rs_datos15.Open "select * from av_lista_productos where saldo_actual >= 0 order by DescDetalle ", db, adOpenKeyset, adLockReadOnly  'JQA 06/2008
    rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
    Set ado_datos15.Recordset = rs_datos15
    ado_datos15.Refresh

   'wwwwwwwwwwwwwwwwwwww
    'db.Execute "DELETE ao_ventas_cabecera where venta_codigo = 0 "
    'Call ABREVENTAS

'    Set rs_Dsctos = New ADODB.Recordset
'    If rs_Dsctos.State = 1 Then rs_Dsctos.Close
'    rs_Dsctos.Open "select * from ac_ventas_descuentos ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
'    Set AdoDsctos.Recordset = rs_Dsctos
'    AdoDsctos.Refresh

    Set rs_datos17 = New ADODB.Recordset
    If rs_datos17.State = 1 Then rs_datos17.Close
    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
    Set ado_datos17.Recordset = rs_datos17
    ado_datos17.Refresh

    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
'    dtc_ctades.BoundText = dtc_cta.BoundText

    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
'  End If
'  If glPersNew = "L" Then
'    frmmo_formulario_M1.Dtc_doc_id_lab = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell_lab = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2apell_lab = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre_lab = rs_Personal!pers_nombres
'  End If
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
'  glPersNew = "N"

End Sub

Private Sub OpMod1_Click()
'    Txt_modelo.Text = Txt_modelo1.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs
'    End If
'    'Set ado_datos17.Recordset = rs_datos18
'    'ado_datos17.Refresh
End Sub

Private Sub OpMod2_Click()
'    Txt_modelo.Text = Txt_modelo2.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_h
'    End If
End Sub

Private Sub OpMod3_Click()
'    Txt_modelo.Text = Txt_modelo3.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_x
'    End If
End Sub

'Private Sub OptFilGral01_Click()
'  '===== Proceso para filtrado general de datos(registros no aprobados)
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
'    'Set Ado_datos9.Recordset = rs_datos9
'    'dtc_desc1.BoundText = dtc_codigo1.BoundText
'    Set rs_datos01 = New Recordset
'    If rs_datos01.State = 1 Then rs_datos01.Close
'        If glusuario = "VPAREDES" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Then
'            queryinicial1 = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and doc_codigo_fac <> 'R-103') "      'ORDER BY cobranza_fecha_prog
'        Else
'            If glusuario = "ADMIN" Then
'                queryinicial1 = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG') "      'ORDER BY cobranza_fecha_prog
'            Else
'                queryinicial1 = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'A' AND estado_codigo_fac = 'E' and doc_codigo_fac = 'R') "      'ORDER BY cobranza_fecha_prog
'            End If
'        End If
'
''    If glusuario = "ADMIN" Then
''        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'REG' "
''    Else
''        If glusuario = "FDELGADILLO" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "RVALDIVIEZO" Or glusuario = "SQUISPE" Or glusuario = "VPAREDES" Or glusuario = "RTORREZ" Or glusuario = "HMARIN" Then
''            queryinicial = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'REG' AND (unidad_codigo = 'DVTA' or unidad_codigo ='DCOMS' or unidad_codigo ='DCOMB' or unidad_codigo ='DCOMC')) "
''        Else
''            queryinicial = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "') "
''        End If
''    End If
'    'queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos01.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    rs_datos01.Sort = "cobranza_fecha_prog"
'    Set Ado_datos01.Recordset = rs_datos01.DataSource
'    Set dg_datos1.DataSource = Ado_datos01.Recordset
'
'End Sub

'Private Sub OptFilGral02_Click()
''===== Proceso para filtrado general de datos (todos los registros )
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
'
'    Set rs_datos01 = New Recordset
'    If rs_datos01.State = 1 Then rs_datos01.Close
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
'
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    If glusuario = "FACTURACION" Or glusuario = "SQUISPE" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Then
'        queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' and doc_codigo_fac <> 'R-103' AND estado_codigo_fac1 = 'APR' AND trans_codigo= 'L' "     'ORDER BY cobranza_fecha_prog
'    Else
'        If glusuario = "ADMIN" Or glusuario = "VPAREDES" Then
'                queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' AND estado_codigo_fac1 = 'APR' AND trans_codigo= 'L' "
'            Else
'                queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'A' AND estado_codigo_fac = 'A' AND estado_codigo_bco = 'R' "      'ORDER BY cobranza_fecha_prog
'            End If
'    '    queryinicial = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
'    End If
''    'queryinicial1 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
''    If glusuario = "ADMIN" Then
''        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG'  "
''        'queryinicial = "SELECT ao_ventas_cobranza.*, ao_ventas_cabecera.* FROM ao_ventas_cobranza INNER JOIN ao_ventas_cabecera ON ao_ventas_cobranza.venta_codigo = ao_ventas_cabecera.venta_codigo"
''    Else
''        If glusuario = "RVALDIVIEZO" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "SQUISPE" Or glusuario = "FDELGADILLO" Or glusuario = "RTORREZ" Then
''            queryinicial = "select * From av_ventas_cobranza WHERE ((unidad_codigo = 'DVTA' or unidad_codigo ='DCOMS' or unidad_codigo ='DCOMB' or unidad_codigo ='DCOMC') and estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG') "
''        Else
''            queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
''        End If
''    End If
'    rs_datos01.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    rs_datos01.Sort = "cobranza_fecha_prog"
'    Set Ado_datos01.Recordset = rs_datos01.DataSource
'    Set dg_datos1.DataSource = Ado_datos01.Recordset
'End Sub

Private Sub OptFilGral04_Click()
  '===== Proceso para filtrado general de datos(Todos los registros)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText

    Set rs_datos02 = New Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
    If glusuario = "RVALDIVIEZO" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "VPAREDES" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "HMARIN" Then
        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR'  "
    Else
        If glusuario = "ADMIN" Then
            queryinicial2 = "select * From av_ventas_cobranza estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG'  "
        Else
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
    rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos02.Sort = "cobranza_fecha_fac"
    Set Ado_datos02.Recordset = rs_datos02.DataSource
    Set dg_datos2.DataSource = Ado_datos02.Recordset
End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
'    'Set Ado_datos9.Recordset = rs_datos9
'    'dtc_desc1.BoundText = dtc_codigo1.BoundText
        Set rs_datos0 = New Recordset
        If rs_datos0.State = 1 Then rs_datos0.Close
        If glusuario = "VPAREDES" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Or glusuario = "ADMIN" Then
            queryinicial = "select * From ao_ventas_cobranza_fac WHERE estado_codigo_fac = 'REG'  AND fecha_registro = '" & Date & "' "     '
        Else
            queryinicial = "select * From ao_ventas_cobranza_fac WHERE (estado_codigo_fac = 'A') "      'ORDER BY cobranza_fecha_prog
        End If
        rs_datos0.Open queryinicial, db, adOpenKeyset, adLockOptimistic
        rs_datos0.Sort = "idfactura"
        Set Ado_datos1.Recordset = rs_datos0.DataSource
        Set dg_datos1.DataSource = Ado_datos1.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic

    Set rs_datos0 = New Recordset
    If rs_datos0.State = 1 Then rs_datos0.Close
    If glusuario = "FACTURACION" Or glusuario = "SQUISPE" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Or glusuario = "ADMIN" Then
        queryinicial = "select * From ao_ventas_cobranza_fac WHERE ((estado_codigo_fac = 'APR' or estado_codigo_fac = 'ANL')  and doc_codigo_fac = 'R-101' AND estado_cobrado = 'REG') "     'ORDER BY cobranza_fecha_prog     '<>
    Else
        queryinicial = "select * From ao_ventas_cobranza_fac WHERE (estado_codigo_fac = 'A') "      'ORDER BY cobranza_fecha_prog
    End If
    'queryinicial = "select * From ao_ventas_cobranza  ORDER BY cobranza_fecha_prog "
    rs_datos0.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos0.Sort = "idfactura"
    Set Ado_datos1.Recordset = rs_datos0.DataSource
    Set dg_datos1.DataSource = Ado_datos1.Recordset
End Sub

Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares_contra.Text = 0
    End If
  End If
End Sub

Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares.Text = 0
    End If
  End If

End Sub

Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos_contra.Text = 0
    End If
  End If
End Sub

Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos.Text = 0
    End If
  End If
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtsolpeso_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then

    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtterref_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Then
        KeyAscii = Asc(UCase(Chr(0)))
    Else
        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "N" Or KeyAscii = 8 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            KeyAscii = Asc(UCase(Chr(0)))
            MsgBox "Debe escribir solo 'N' o 'S'", vbOKOnly, "Error..."
        End If
    End If
End Sub

Private Sub cerea()
  txt_venta = " "
  dtc_codigo4.Text = " "
  Dtcpaternosol.Text = " "  'dtc_codigo4.BoundText
'  dtcmaternosol.Text = " "
'  dtcnombresol.Text = " "
  txtCantTotal = "0"
  TxtMontoBs = "0"
  TxtMontoUs = "0"
  TxtConcepto = ""
  dtc_codigo2 = ""
  dtc_desc2 = ""
  txtTDC.Text = GlTipoCambioMercado ' GlTipoCambioOficial

'  DtCDenominacion_moneda = ""
'  TxtMonto_bolivianos = 0
'  Txtmonto_dolares = 0
'  TxtMonto_bolivianos_contra = 0
'  Txtmonto_dolares_contra = 0
'  DtCOrg_descripcion = ""
'  txtjustifica = ""
'  txt_venta = ""
'  txtterref = ""
End Sub
'Private Sub fbuscaunidad()
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  If rstFc_unidad_ejecutora.RecordCount > 0 Then
'    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
'  Else
'    LblUni_descripcion_larga.Caption = ""
'  End If
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'End Sub

Sub creaVista()
db.Execute "drop view vwF04"

db.Execute "create view vwF04 as " & _
            "select  ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.tipoben_codigo, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, ao_solicitud_lista.telefono, ao_solicitud_lista.razon_s, ao_solicitud.codigo_solicitud, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_numero, ao_solicitud.por_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.caracteristicas, ao_solicitud.duracion_estimada_tiempo, " & _
            "ao_solicitud.tr_adjuntos AS docAdjunta, " & _
            "ao_solicitud.codigo_bien, ac_bienes.bie_descripcion , ao_solicitud.observaciones, fc_unidad_ejecutora.uni_descripcion_larga, ao_solicitud.fecha_solicitud, " & _
            "(rc_personal.paterno) + ' ' + (rc_personal.materno) + ' ' +(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
            "from ao_solicitud_lista  ,     " & _
                 "ao_solicitud       ,     " & _
                 "fc_unidad_ejecutora,     " & _
                 "rc_personal,             " & _
                 "ac_bienes                " & _
            "where  ao_solicitud_lista.ges_Gestion       = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
                    "ao_solicitud_lista.codigo_unidad    = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
                    "ao_solicitud_lista.codigo_solicitud =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
                    "ao_solicitud_lista.ges_Gestion      = ao_solicitud.ges_gestion            and " & _
                    "ao_solicitud_lista.codigo_unidad    = ao_solicitud.codigo_unidad          and " & _
                    "ao_solicitud_lista.codigo_solicitud = ao_solicitud.codigo_solicitud       and " & _
                    "ao_solicitud.codigo_unidad          = fc_unidad_ejecutora.codigo_unidad   and " & _
                    "ao_solicitud.codigo_bien            = ac_bienes.codigo_bien               and " & _
                    "ao_solicitud.ci                     = rc_personal.ci                      " & _
            "GROUP BY ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.tipoben_codigo, " & _
            "ao_solicitud.codigo_solicitud, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.razon_s, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, " & _
            "ao_solicitud_lista.telefono, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.nacional_extranjero, ao_solicitud.por_tiempo, ao_solicitud.codigo_bien, ac_bienes.bie_descripcion, ao_solicitud.duracion_estimada_numero, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.esparaRH, ao_solicitud.tr_adjuntos, ao_solicitud.observaciones, ao_solicitud.caracteristicas, fc_unidad_ejecutora.Uni_descripcion_larga, ao_solicitud.fecha_solicitud, (rc_personal.paterno)+' '+(rc_personal.materno)+' '+(rc_personal.nombres)+' ['+ao_solicitud.ci+']', ao_solicitud_lista.id_beneficiario "

'            "trim$(rc_personal.paterno) + ' ' + trim$(rc_personal.materno) + ' ' +trim$(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _

'''db.Execute "create view vwF05 as " & _
'''            "select  ao_solicitud_lista.* " & _
'''            "from ao_solicitud_lista"
End Sub

Sub CREAVISTAF11()
db.Execute "drop view VWF11"
db.Execute "create view VWF11 as " & _
    "SELECT ao_Solicitud.Ges_Gestion, ao_Solicitud.codigo_unidad, " & _
    "ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, " & _
    "ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, " & _
    "ao_Solicitud.fecha_solicitud, ao_Solicitud.codigo_bien, " & _
    "ac_bienes_grupo.DescGrupo, RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres, " & _
    "ao_Solicitud.observaciones, ao_Solicitud.caracteristicas, " & _
    "ao_Solicitud.tr_adjuntos, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, " & _
    "ao_Solicitud.duracion_estimada_numero, ao_Solicitud.duracion_estimada_tiempo, " & _
    "ao_solicitud_lista.codDetalle AS ci_material,  ao_solicitud_lista.profesion, ao_solicitud_lista.Aplanilla, " & _
    "ao_solicitud_lista.razon_s, ao_solicitud_lista.Nro_pagos, ao_solicitud_lista.Monto_solicitud_dl, ao_solicitud_lista.AUnidad " & _
"FROM ao_Solicitud, ao_Solicitud_detalle, ac_bienes_grupo, RC_Personal, ao_solicitud_lista " & _
"WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    "(ao_Solicitud.codigo_unidad) = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
    "(ao_Solicitud.codigo_solicitud) =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
    "ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_lista.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_lista.codigo_solicitud AND " & _
    "ao_Solicitud.CodGrupo = ac_bienes_grupo.CodGrupo AND " & _
    "ao_Solicitud.ci = RC_Personal.ci"
End Sub

Private Sub acumulaMont(ges, Nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  Set rs_datos19 = New ADODB.Recordset
  If rs_datos19.State = 1 Then rs_datos19.Close
'  LblGestion
'  lblcorrelVenta
'  lblNroVenta
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!CANTOT
  End If

  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If

  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & Nro & " "

'  TxtMontoBs.Text = VAR_AUX
'  TxtCobrado.Text = Cobrobs
'  TxtBstotal.Text = VAR_Bs

  If rstacumdet.State = 1 Then rstacumdet.Close

End Sub

Private Sub Option1_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic

    Set rs_datos0 = New Recordset
    If rs_datos0.State = 1 Then rs_datos0.Close
    If glusuario = "FACTURACION" Or glusuario = "SQUISPE" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Or glusuario = "ADMIN" Then
        queryinicial = "select * From ao_ventas_cobranza_fac WHERE estado_codigo_fac = 'APR'  and doc_codigo_fac = 'R-101' AND estado_cobrado = 'APR' "     'ORDER BY cobranza_fecha_prog     '<> 'R-103'    'estado_codigo_sol = 'APR' AND AND estado_codigo_bco = 'REG'
    Else
        queryinicial = "select * From ao_ventas_cobranza_fac WHERE (estado_codigo_fac = 'A') "      'ORDER BY cobranza_fecha_prog
    End If
    rs_datos0.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos0.Sort = "idfactura"
    Set Ado_datos1.Recordset = rs_datos0.DataSource
    Set dg_datos1.DataSource = Ado_datos1.Recordset
End Sub

Private Sub Option2_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
        Set rs_datos0 = New Recordset
        If rs_datos0.State = 1 Then rs_datos0.Close
        If glusuario = "VPAREDES" Or glusuario = "SQUISPE" Or glusuario = "FACTURACION" Or glusuario = "FDELGADILLO" Or glusuario = "HMARIN" Or glusuario = "ADMIN" Then
            queryinicial = "select * From ao_ventas_cobranza_fac WHERE estado_codigo_fac = 'REG'  AND fecha_registro <> '" & Date & "' "     '
        Else
            queryinicial = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'A' AND estado_codigo_fac = 'A') "      'ORDER BY cobranza_fecha_prog
        End If
        rs_datos0.Open queryinicial, db, adOpenKeyset, adLockOptimistic
        rs_datos0.Sort = "Idfactura"
        Set Ado_datos1.Recordset = rs_datos0.DataSource
        Set dg_datos1.DataSource = Ado_datos1.Recordset
End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
            lbl_titulo.Caption = SSTab1.Caption
            FraNavega.Caption = SSTab1.Caption
            'SSTab1.Tab = 0
            varTipo = 1
            fraOpciones.backColor = &H404040
            FraNavega.backColor = &H404040
            FraNavega.Caption = &H0&
            lbl_titulo.ForeColor = &HFFFF80
            Call OptFilGral1_Click
        Case 1
            lbl_titulo.Caption = SSTab1.Caption
            FraNavega.Caption = SSTab1.Caption
            varTipo = 2
            fraOpciones.backColor = &HC0C0C0
            FraNavega.backColor = &HC0C0C0
            lbl_titulo.ForeColor = &H0&
            Call OptFilGral1_Click

        Case 2
            lbl_titulo.Caption = SSTab1.Caption
            FraNavega.Caption = SSTab1.Caption
            varTipo = 0
            fraOpciones.backColor = &HC0FFFF
            FraNavega.backColor = &HC0FFFF
            lbl_titulo.ForeColor = &H0&
            Call OptFilGral1_Click
    End Select
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtCantidad_LostFocus()
  If (TxtCantidad.Text) = "" Then
    TxtCantidad.Text = 1
  End If
  If dtc_codigo11.Text = "E" Then
    If (dtc_codigo12.Text) = "" Or IsNull(dtc_codigo12.Text) Then
        TxtDescuento.Text = "0"
    Else
        TxtDescuento.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) * CDbl(Dtc_aux12.Text))
    End If
    'TxtPrecioU.Text = dtc_precioventabase15.Text
    'TxtTotal.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) - CDbl(TxtDescuento.Text))
  End If
  If dtc_codigo11.Text = "C" Then
     TxtDescuento.Text = "0"
     'TxtDescuento.Text = CDbl(Dtc_aux12) * (CDbl(TxtCantidad) * CDbl(TxtPrecioU))
     TxtPrecioU.Text = dtc_precioventafinal15.Text
  End If
  If (dtc_codigo11.Text <> "E" And dtc_codigo11.Text <> "C") Then
     TxtDescuento.Text = "0"
     TxtPrecioU.Text = "0"
  End If
  TxtTotal.Text = (CDbl(TxtCantidad.Text) * CDbl(TxtPrecioU.Text)) - CDbl(TxtDescuento.Text)

End Sub

Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtDscto1_LostFocus()
    TxtMonto1.Text = Round(CDbl(TxtDscto1.Text) * GlTipoCambioMercado, 2)
End Sub

Private Sub TxtDscto2_LostFocus()
    TxtDscto2D.Text = Round(CDbl(TxtDscto2.Text) / Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtDscto2D_LostFocus()
    TxtDscto2.Text = Round(CDbl(TxtDscto2D.Text) * Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtMontoDol = "0"
    Else
        'TxtMontoDol = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
        TxtMontoDol = Round(CDbl(TxtMonto.Text) / CDbl(Txt_tdc), 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub
'adelante
Private Function CodigoControl(NAuto As String, NFactura As String, Nit As String, Fecha As String, Monto As String, Key As String) As String
Dim Suma As Currency
'Dim Suma As String
'
Dim CodControl As String, Cadena As String, NroVer As String
Dim Pos As Integer, i As Integer, Nro As Integer, j As Integer
Dim SumTot As Long, SumPar(1 To 5) As Currency


  Suma = 0
  Cadena = NFactura
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i

  NFactura = Cadena
  Suma = Suma + CDbl(Cadena)

  'MsgBox NFactura
  'Para el Nit o CI del Cliente.
  Cadena = Nit
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Nit = Cadena
  'se cambio de & a + en 13/04/2017
  Suma = Suma + CDbl(Cadena)
  'Suma = Suma & Cadena
  'MsgBox Nit
  'Para la Fecha de transaccion.
  Cadena = Fecha
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Fecha = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox Fecha
  'Para el monto de transaccion.
  Cadena = Monto
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Monto = Cadena
  'MsgBox Monto
  Suma = Suma + CDbl(Cadena)
  'MsgBox Suma

  'Para Obtener los 5 numeros Verhoeff.
  Cadena = Str(Suma)
  For i = 1 To 5
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  NroVer = Right(Cadena, 5)
  'MsgBox NroVer

  'Para obtener las nuevas cadenas.
  Cadena = ""
  Pos = 1
  For i = 1 To 5
    Nro = (Val(Mid(NroVer, i, 1)) + 1)
    Select Case i
      Case 1: Cadena = LTrim(Cadena) & LTrim(NAuto) & Mid(Key, Pos, Nro)
      Case 2: Cadena = LTrim(Cadena) & LTrim(NFactura) & Mid(Key, Pos, Nro)
      Case 3: Cadena = LTrim(Cadena) & LTrim(Nit) & Mid(Key, Pos, Nro)
      Case 4: Cadena = LTrim(Cadena) & LTrim(Fecha) & Mid(Key, Pos, Nro)
      Case 5: Cadena = LTrim(Cadena) & LTrim(Monto) & Mid(Key, Pos, Nro)
    End Select
    Pos = Pos + Nro
  Next i

  Cadena = AllegedRC4(Cadena, (Key & NroVer))


  SumTot = 0
  i = 0
  Do While i < Len(Cadena)
    i = i + 1
    SumTot = SumTot + Asc(Mid(Trim(Cadena), i, 1))
    sino = Mid(Trim(Cadena), i, 1)
  Loop



  For i = 1 To 5
    SumPar(i) = 0
    sino = ""

    j = i
    Do While j <= Len(Cadena)

      SumPar(i) = SumPar(i) + Asc(Mid(Cadena, j, 1))
      sino = Asc(Mid(Cadena, j, 1))
      Caracter = Mid(Cadena, j, 1)
      j = j + 5

    Loop

  Next i

  Suma = 0
  For i = 1 To 5
    SumPar(i) = Int((SumTot * SumPar(i)) / (Val(Mid(NroVer, i, 1)) + 1))
    Suma = Suma + SumPar(i)
  Next i
  Cadena = Base64(Str(Suma))

  Cadena = AllegedRC4(Cadena, (Key & NroVer))


  CodigoControl = ""
  i = 0
  j = 1

  Do While i < Len(Cadena)
    i = i + 1
    If i Mod 2 = 0 Then
      CodigoControl = CodigoControl & Mid(Cadena, j, 2) & "-"
      j = i + 1
    End If
  Loop

  CodigoControl = Mid(CodigoControl, 1, (Len(CodigoControl) - 1))
End Function
Public Function Redondear(dNumero As Double, iDecimales As Integer) As Double
    Dim lMultiplicador As Long
    Dim dRetorno As Double

    If iDecimales > 9 Then iDecimales = 9
    lMultiplicador = 10 ^ iDecimales
    dRetorno = CDbl(CLng(dNumero * lMultiplicador)) / lMultiplicador

    Redondear = dRetorno
End Function

Private Function Redondeo(ByVal Numero, ByVal Decimales)
      Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
End Function

Private Sub TxtMonto02_LostFocus()
    TxtMonto02D.Text = Round(CDbl(TxtMonto02.Text) / Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMonto02D_LostFocus()
    TxtMonto02.Text = Round(CDbl(TxtMonto02D.Text) * Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMonto1_LostFocus()
    TxtDscto1.Text = Round(CDbl(TxtMonto1.Text) / GlTipoCambioMercado, 2)
End Sub

Private Sub TxtMontoDol_Change()
    'TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(txt_tdc.Text)
End Sub

Private Sub TxtMontoDol_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtMontoDol_LostFocus()
    TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(Txt_tdc.Text)
End Sub



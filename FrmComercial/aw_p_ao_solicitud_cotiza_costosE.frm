VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_p_ao_solicitud_cotiza_costosE 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización Venta - Hoja de costos (Europa)"
   ClientHeight    =   8250
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12750
   ControlBox      =   0   'False
   Icon            =   "aw_p_ao_solicitud_cotiza_costosE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   12750
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   12435
      TabIndex        =   27
      Top             =   120
      Width           =   12495
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2040
         Picture         =   "aw_p_ao_solicitud_cotiza_costosE.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   108
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   480
         Picture         =   "aw_p_ao_solicitud_cotiza_costosE.frx":12EE
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   107
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HOJA DE COSTOS - EUROPA"
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
         Left            =   5565
         TabIndex        =   28
         Top             =   240
         Width           =   4425
      End
   End
   Begin VB.Frame Fra_datos99 
      BackColor       =   &H00C0C0C0&
      Height          =   6975
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   12495
      Begin VB.TextBox txt_tdc_me 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "cotiza_tdc_bol"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
         Height          =   285
         Left            =   5880
         TabIndex        =   2
         Text            =   "7.5"
         Top             =   1200
         Width           =   885
      End
      Begin VB.TextBox txt_montobase 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "costo_monto"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10965
         TabIndex        =   4
         Text            =   "0"
         Top             =   1200
         Width           =   1340
      End
      Begin VB.TextBox txt_tdc 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "cotiza_tdc_bol"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
         Height          =   285
         Left            =   7845
         TabIndex        =   3
         Text            =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox cmd_moneda 
         BackColor       =   &H0080FFFF&
         DataField       =   "tipo_moneda"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
         Height          =   315
         ItemData        =   "aw_p_ao_solicitud_cotiza_costosE.frx":1AC4
         Left            =   4080
         List            =   "aw_p_ao_solicitud_cotiza_costosE.frx":1AD7
         TabIndex        =   1
         Text            =   "EUR"
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cmd_dec 
         BackColor       =   &H0080FFFF&
         DataField       =   "cotiza_dec"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
         Height          =   315
         ItemData        =   "aw_p_ao_solicitud_cotiza_costosE.frx":1AF4
         Left            =   1720
         List            =   "aw_p_ao_solicitud_cotiza_costosE.frx":1B01
         TabIndex        =   0
         Text            =   "2"
         Top             =   1200
         Width           =   580
      End
      Begin VB.Frame FraModeloCostoA 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"aw_p_ao_solicitud_cotiza_costosE.frx":1B0E
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5265
         Left            =   120
         TabIndex        =   35
         Top             =   1560
         Width           =   12240
         Begin VB.TextBox txt_paradas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10800
            TabIndex        =   105
            Text            =   "0"
            Top             =   4800
            Width           =   765
         End
         Begin VB.TextBox Text2 
            DataField       =   "bien_codigo"
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            Height          =   315
            Left            =   8400
            TabIndex        =   104
            Text            =   "0"
            Top             =   4200
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox Txt_campo5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_nro_montador"
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7680
            TabIndex        =   103
            Text            =   "0"
            Top             =   4800
            Width           =   765
         End
         Begin VB.TextBox txt_base_imp_me 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            DataField       =   "cotiza_precio_base_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   5265
            Locked          =   -1  'True
            TabIndex        =   101
            Text            =   "0"
            Top             =   3960
            Width           =   1340
         End
         Begin VB.TextBox txt_gac_me 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_GAC_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   100
            Text            =   "0"
            Top             =   3560
            Width           =   1340
         End
         Begin VB.TextBox txt_base_imp_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            DataField       =   "cotiza_precio_base_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   98
            Text            =   "0"
            Top             =   3960
            Width           =   1340
         End
         Begin VB.TextBox txt_base_imp_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            DataField       =   "cotiza_precio_base_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   97
            Text            =   "0"
            Top             =   3960
            Width           =   1340
         End
         Begin VB.TextBox txt_gac_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_GAC_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   95
            Text            =   "0"
            Top             =   3560
            Width           =   1340
         End
         Begin VB.TextBox txt_gac_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_GAC_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2340
            TabIndex        =   11
            Text            =   "0"
            Top             =   3560
            Width           =   1340
         End
         Begin VB.TextBox txt_total_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00404080&
            DataField       =   "cotiza_precio_total_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "0"
            Top             =   4760
            Width           =   1340
         End
         Begin VB.TextBox txt_gastos_locales_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            DataField       =   "cotiza_gasto_local_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   89
            Text            =   "0"
            Top             =   4360
            Width           =   1340
         End
         Begin VB.TextBox txt_spread_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_spread_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2340
            TabIndex        =   8
            Text            =   "0"
            Top             =   1560
            Width           =   1340
         End
         Begin VB.TextBox txt_tacb_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_tacb_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2340
            TabIndex        =   7
            Text            =   "0"
            Top             =   1160
            Width           =   1340
         End
         Begin VB.TextBox txt_fletefrontera_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_flete_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2340
            TabIndex        =   10
            Text            =   "0"
            Top             =   2760
            Width           =   1340
         End
         Begin VB.TextBox txt_cif_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            DataField       =   "cotiza_precio_cif_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   2340
            TabIndex        =   88
            Text            =   "0"
            Top             =   3160
            Width           =   1340
         End
         Begin VB.TextBox txt_fob_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_fob_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   2340
            TabIndex        =   5
            Text            =   "0"
            Top             =   360
            Width           =   1340
         End
         Begin VB.TextBox txt_dcto_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_dcto_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2340
            TabIndex        =   6
            Text            =   "0"
            Top             =   760
            Width           =   1340
         End
         Begin VB.TextBox txt_seguro_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_seg_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2340
            TabIndex        =   9
            Text            =   "0"
            Top             =   1960
            Width           =   1340
         End
         Begin VB.TextBox txt_fob_seg_eu 
            Alignment       =   2  'Center
            BackColor       =   &H00004040&
            DataField       =   "cotiza_fob_seg_me"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "0"
            Top             =   2360
            Width           =   1340
         End
         Begin VB.TextBox txt_fob_seg_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00004040&
            DataField       =   "cotiza_fob_seg_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "0"
            Top             =   2360
            Width           =   1340
         End
         Begin VB.TextBox txt_fob_seg_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00004040&
            DataField       =   "cotiza_fob_seg_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   5265
            Locked          =   -1  'True
            TabIndex        =   63
            Text            =   "0"
            Top             =   2360
            Width           =   1340
         End
         Begin VB.TextBox txt_tac_billing_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_tac_billing_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0"
            Top             =   2715
            Width           =   1340
         End
         Begin VB.TextBox txt_tac_billing_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_tac_billing_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9285
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "0"
            Top             =   2760
            Width           =   1340
         End
         Begin VB.TextBox txt_cge_IVA_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_cge_IVA_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0"
            Top             =   2325
            Width           =   1340
         End
         Begin VB.TextBox txt_cge_IVA_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_cge_IVA_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9285
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "0"
            Top             =   2325
            Width           =   1340
         End
         Begin VB.TextBox txt_cge_IT_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_cge_IT_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0"
            Top             =   1920
            Width           =   1340
         End
         Begin VB.TextBox txt_cge_IT_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_cge_IT_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9285
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "0"
            Top             =   1920
            Width           =   1340
         End
         Begin VB.TextBox txt_local_IVA_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_local_IVA_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0"
            Top             =   760
            Width           =   1340
         End
         Begin VB.TextBox txt_local_IVA_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_local_IVA_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9285
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "0"
            Top             =   760
            Width           =   1340
         End
         Begin VB.TextBox txt_gastos_locales_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            DataField       =   "cotiza_gasto_local_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   5265
            Locked          =   -1  'True
            TabIndex        =   58
            Text            =   "0"
            Top             =   4360
            Width           =   1340
         End
         Begin VB.TextBox txt_gastos_locales_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            DataField       =   "cotiza_gasto_local_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "0"
            Top             =   4360
            Width           =   1340
         End
         Begin VB.TextBox txt_local_IT_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_local_IT_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   360
            Width           =   1340
         End
         Begin VB.TextBox txt_local_IT_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_saldo_local_IT_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9285
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "0"
            Top             =   360
            Width           =   1340
         End
         Begin VB.TextBox txt_seguro_me 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_seg_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5265
            TabIndex        =   55
            Text            =   "0"
            Top             =   1960
            Width           =   1340
         End
         Begin VB.TextBox txt_dcto_me 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_dcto_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5265
            TabIndex        =   54
            Text            =   "0"
            Top             =   760
            Width           =   1340
         End
         Begin VB.TextBox txt_fob_me 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_fob_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   5265
            TabIndex        =   53
            Text            =   "0"
            Top             =   360
            Width           =   1340
         End
         Begin VB.TextBox txt_cif_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            DataField       =   "cotiza_precio_cif_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "0"
            Top             =   3160
            Width           =   1340
         End
         Begin VB.TextBox txt_fletefrontera_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_flete_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "0"
            Top             =   2760
            Width           =   1340
         End
         Begin VB.TextBox txt_total_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00404080&
            DataField       =   "cotiza_precio_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "0"
            Top             =   4760
            Width           =   1340
         End
         Begin VB.TextBox txt_cif_me 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            DataField       =   "cotiza_precio_cif_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   5265
            TabIndex        =   49
            Text            =   "0"
            Top             =   3160
            Width           =   1340
         End
         Begin VB.TextBox txt_fletefrontera_me 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_flete_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5265
            TabIndex        =   48
            Text            =   "0"
            Top             =   2760
            Width           =   1340
         End
         Begin VB.TextBox txt_total_me 
            Alignment       =   2  'Center
            BackColor       =   &H00404080&
            DataField       =   "cotiza_precio_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   5265
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "0"
            Top             =   4760
            Width           =   1340
         End
         Begin VB.TextBox txt_seguro_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_seg_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "0"
            Top             =   1960
            Width           =   1340
         End
         Begin VB.TextBox txt_dcto_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_dcto_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "0"
            Top             =   760
            Width           =   1340
         End
         Begin VB.TextBox txt_fob_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_fob_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "0"
            Top             =   345
            Width           =   1340
         End
         Begin VB.TextBox txt_totalCli_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            DataField       =   "cotiza_precio_total_bs_cli"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   9285
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0"
            Top             =   1160
            Width           =   1340
         End
         Begin VB.TextBox txt_totalCli_me 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            DataField       =   "cotiza_precio_total_dol_cli"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0"
            Top             =   1160
            Width           =   1340
         End
         Begin VB.TextBox txt_totalCGE_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            DataField       =   "cotiza_precio_total_bs_cge"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   9285
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "0"
            Top             =   3120
            Width           =   1340
         End
         Begin VB.TextBox txt_totalCGE_me 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            DataField       =   "cotiza_precio_total_dol_cge"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
            Height          =   315
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0"
            Top             =   3120
            Width           =   1340
         End
         Begin VB.TextBox txt_tacb_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_tacb_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3795
            TabIndex        =   39
            Text            =   "0"
            Top             =   1160
            Width           =   1340
         End
         Begin VB.TextBox txt_spread_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_spread_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3795
            TabIndex        =   38
            Text            =   "0"
            Top             =   1560
            Width           =   1340
         End
         Begin VB.TextBox txt_tacb_me 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_tacb_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5265
            TabIndex        =   37
            Text            =   "0"
            Top             =   1160
            Width           =   1340
         End
         Begin VB.TextBox txt_spread_me 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cotiza_precio_spread_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5265
            TabIndex        =   36
            Text            =   "0"
            Top             =   1560
            Width           =   1340
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Número de Paradas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   10200
            TabIndex        =   106
            Top             =   4520
            Width           =   1830
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FF0000&
            X1              =   6765
            X2              =   12195
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad de Montadores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   7080
            TabIndex        =   102
            Top             =   4520
            Width           =   2220
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Base Imponible:"
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
            Left            =   840
            TabIndex        =   99
            Top             =   3960
            Width           =   1365
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "GAC:"
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
            Left            =   1800
            TabIndex        =   96
            Top             =   3585
            Width           =   465
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   6765
            X2              =   12240
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lbl_campo6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "FOB+SEG+Puerto+Emb.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   195
            Left            =   75
            TabIndex        =   81
            Top             =   2385
            Width           =   2100
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Tac Billing:"
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
            Left            =   7635
            TabIndex        =   80
            Top             =   2745
            Width           =   1575
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Importacion - IVA:"
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
            Left            =   7065
            TabIndex        =   79
            Top             =   2340
            Width           =   2145
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Importacion - IT:"
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
            Left            =   7185
            TabIndex        =   78
            Top             =   1935
            Width           =   2010
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Local - IVA:"
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
            Left            =   7635
            TabIndex        =   77
            Top             =   780
            Width           =   1590
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos Locales:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   780
            TabIndex        =   76
            Top             =   4380
            Width           =   1380
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Local - IT:"
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
            Left            =   7740
            TabIndex        =   75
            Top             =   375
            Width           =   1455
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FF0000&
            X1              =   6765
            X2              =   6765
            Y1              =   120
            Y2              =   5280
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento:"
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
            Left            =   1245
            TabIndex        =   74
            Top             =   780
            Width           =   1020
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Seguro Transporte(SEG):"
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
            Left            =   60
            TabIndex        =   73
            Top             =   1980
            Width           =   2280
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio FOB:"
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
            Left            =   1185
            TabIndex        =   72
            Top             =   375
            Width           =   1080
         End
         Begin VB.Label lbl_campo7 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Flete de Frontera:"
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
            Left            =   675
            TabIndex        =   71
            Top             =   2775
            Width           =   1575
         End
         Begin VB.Label lbl_campo8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio CIF:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1275
            TabIndex        =   70
            Top             =   3180
            Width           =   960
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal Funcionando:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   90
            TabIndex        =   69
            Top             =   4785
            Width           =   1995
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Importación Directa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   6945
            TabIndex        =   68
            Top             =   1185
            Width           =   2220
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Facturación Local:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   7065
            TabIndex        =   67
            Top             =   3135
            Width           =   2085
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Puerto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   195
            Left            =   1605
            TabIndex        =   66
            Top             =   1185
            Width           =   630
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Embalaje:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   195
            Left            =   1365
            TabIndex        =   65
            Top             =   1575
            Width           =   840
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Múltiplo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   7080
         TabIndex        =   94
         Top             =   1200
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFC0&
         X1              =   0
         X2              =   12480
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pais"
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
         Left            =   11520
         TabIndex        =   93
         Top             =   330
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Continente"
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
         Left            =   10080
         TabIndex        =   92
         Top             =   330
         Width           =   945
      End
      Begin VB.Label txt_pais 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "pais_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
         Height          =   300
         Left            =   11400
         TabIndex        =   91
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Moneda Base:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   9000
         TabIndex        =   86
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TDC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   5385
         TabIndex        =   85
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda Origen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   2640
         TabIndex        =   84
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Decimales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   240
         TabIndex        =   83
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label txt_conti 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "pais_continente"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
         Height          =   300
         Left            =   9960
         TabIndex        =   82
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_descripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   1800
         TabIndex        =   34
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Cotización"
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
         Left            =   6600
         TabIndex        =   33
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
         Height          =   300
         Left            =   6480
         TabIndex        =   32
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
         Height          =   300
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidad Ejecutora"
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
         Left            =   1800
         TabIndex        =   29
         Top             =   330
         Width           =   2040
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Trámite "
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
         Left            =   240
         TabIndex        =   26
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36NO-"
         DataField       =   "edif_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datosE"
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
         Height          =   300
         Left            =   8160
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Edificio"
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
         Left            =   8280
         TabIndex        =   25
         Top             =   330
         Width           =   1365
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   12750
      TabIndex        =   18
      Top             =   8250
      Width           =   12750
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   23
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2400
      Top             =   7680
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
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   8040
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
      Caption         =   "Ado_datos1"
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
      Left            =   2880
      Top             =   8040
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   5040
      Top             =   8040
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
      Caption         =   "Ado_datos9"
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
Attribute VB_Name = "aw_p_ao_solicitud_cotiza_costosE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos1 As New ADODB.Recordset
Attribute rs_datos1.VB_VarHelpID = -1
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
'BUSCADOR

'OTROS
Dim var_cod As String
Dim VAR_VAL As String

Dim VAR_1A, VAR_2A As Double
Dim VAR_3B, VAR_4B, VAR_5B, VAR_6B, VAR_7B As Double
Dim VAR_8C, VAR_9C, VAR_10C, VAR_11C, VAR_12C As Double
Dim VAR_13D, VAR_14D As Double
Dim totbs2, totdl2, totbs3, totdl3 As Double
Dim VAR_SUBD, VAR_SUBB As Double
Dim VAR_FLE_ME, VAR_NAC_ME, VAR_ALM_ME, VAR_AGE_ME, VAR_UTIL_ME As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error GoTo AddErr
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        mw_solicitud_cotiza_venta.Ado_datosE.Recordset.CancelUpdate
        Unload Me
    End If
      Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
'WWWWWWWWWWWWWWWWWWWWWWWW
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  VAR_CONTI = "EUROPA"
  Call valida_campos
  If VAR_VAL = "OK" Then
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  ", db, adOpenKeyset, adLockOptimistic
    'Set Ado_datos3.Recordset = rs_datos3
    If rs_datos10.RecordCount > 0 Then
       'sino = MsgBox("SI (Graba todos los Registros) - NO (Graba SOLO el Registro Activo) ... ", vbYesNo + vbQuestion, "Atención")
       'If sino = vbYes Then
'           'TODOS LOS REGISTROS
'           Set mw_solicitud_cotiza_venta.Ado_datosE.Recordset = rs_datos10
'           mw_solicitud_cotiza_venta.Ado_datos1E.Recordset.MoveFirst
'           While Not mw_solicitud_cotiza_venta.Ado_datos1E.Recordset.EOF
''         'MsgBox "codigo: " + Str(rs_datos10!cotiza_codigo)
'             'Set mw_solicitud_cotiza_venta.Ado_datos1E.Recordset = rs_datos10
'             Txt_Correl.Caption = mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_codigo
'             If Val(Txt_Correl.Caption) = 1 Then
'                'GUARDA EL PRIMER REGISTRO
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_dec = CDbl(cmd_dec.Text)
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!tipo_moneda = cmd_moneda.Text
'                 If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
'                    txt_tdc.Text = GlTipoCambioOficial
'                 End If
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_tdc_me = txt_tdc_me.Text
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_tdc_bol = txt_tdc.Text
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!costo_monto = txt_montobase.Text
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_fob_me = IIf(txt_fob_eu = "", "0", Round(CDbl(txt_fob_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_fob_bs = Round(CDbl(txt_fob_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_fob_dol = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_dcto_me = IIf(txt_dcto_eu = "", "0", Round(CDbl(txt_dcto_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_dcto_bs = Round(CDbl(txt_dcto_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_dcto_dol = Round(CDbl(txt_dcto_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_tacb_me = IIf(txt_tacb_eu = "", "0", Round(CDbl(txt_tacb_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_tacb_bs = Round(CDbl(txt_tacb_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_tacb_dol = Round(CDbl(txt_tacb_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_spread_me = IIf(txt_spread_eu = "", "0", Round(CDbl(txt_spread_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_spread_bs = Round(CDbl(txt_spread_me) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_spread_dol = Round(CDbl(txt_spread_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_seg_me = IIf(txt_seguro_eu = "", "0", Round(CDbl(txt_seguro_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_seg_bs = Round(CDbl(txt_seguro_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_seg_dol = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_fob_seg_me = IIf(txt_fob_me = "", "0", Round(CDbl(txt_fob_me), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_fob_seg_bs = Round(CDbl(txt_fob_me) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_fob_seg_dol = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_flete_me = IIf(txt_fletefrontera_eu = "", "0", Round(CDbl(txt_fletefrontera_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_flete_bs = Round(CDbl(txt_fletefrontera_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_flete_dol = Round(CDbl(txt_fletefrontera_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_cif_me = IIf(txt_cif_eu = "", "0", Round(CDbl(txt_cif_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_cif_bs = Round(CDbl(txt_cif_eu) * CDbl(txt_tdc_me), Val(cmd_dec)) '
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_cif_dol = Round(CDbl(txt_cif_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_GAC_me = IIf(txt_gac_eu = "", "0", Round(CDbl(txt_gac_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_GAC_bs = Round(CDbl(txt_gac_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_GAC_dol = Round(CDbl(txt_gac_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_base_me = IIf(txt_base_imp_eu = "", "0", Round(CDbl(txt_base_imp_eu), Val(cmd_dec)))
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_base_bs = Round(CDbl(txt_base_imp_eu) * CDbl(txt_tdc_me), Val(cmd_dec)) '
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_base_dol = Round(CDbl(txt_base_imp_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!fecha_registro = Date     'no cambia
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
'                 mw_solicitud_cotiza_venta.Ado_datos1E.Recordset.Update    'Batch 'adAffectAll
'                 db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'SI' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'             Else
'                'CLONA REGISTROS EN BASE AL PRIMER REGISTRO
'                Set rs_aux7 = New ADODB.Recordset
'                If rs_aux7.State = 1 Then rs_aux7.Close
'                rs_aux7.Open "ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = 1  ", db, adOpenStatic
'                'Set Ado_datos11.Recordset = rs_aux7
'                If rs_aux7.RecordCount > 0 Then
'                    'WWWWWWWWWWWWWWWWWWWWWW
'                    db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'SI' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_dec = " & rs_aux7!cotiza_dec & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set tipo_moneda= '" & rs_aux7!tipo_moneda & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_tdc_bol = " & rs_aux7!cotiza_tdc_bol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set costo_monto = " & rs_aux7!costo_monto & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_dol = " & rs_aux7!cotiza_precio_fob_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_bs = " & Round(CDbl(rs_aux7!cotiza_precio_fob_bs), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_dol = " & rs_aux7!cotiza_precio_dcto_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_bs = " & CDbl(rs_aux7!cotiza_precio_dcto_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_dol = " & rs_aux7!cotiza_precio_seg_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_bs = " & CDbl(rs_aux7!cotiza_precio_seg_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_dol = " & CDbl(rs_aux7!cotiza_fob_seg_dol) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_bs = " & CDbl(rs_aux7!cotiza_fob_seg_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_dol = " & rs_aux7!cotiza_precio_flete_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_bs = " & CDbl(rs_aux7!cotiza_precio_flete_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_tacb_dol = " & Round(rs_aux7!cotiza_precio_tacb_dol, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_tacb_bs = " & Round(rs_aux7!cotiza_precio_tacb_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_spread_dol  = " & Round(rs_aux7!cotiza_precio_spread_dol, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_spread_bs  = " & Round(rs_aux7!cotiza_precio_spread_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_dol = " & Round(CDbl(rs_aux7!cotiza_precio_cif_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_bs = " & Round(rs_aux7!cotiza_precio_cif_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_GAC_dol = " & Round(rs_aux7!cotiza_precio_GAC_dol, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_GAC_bs  = " & Round(rs_aux7!cotiza_precio_GAC_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_base_dol  = " & Round(rs_aux7!cotiza_precio_base_dol, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_base_bs  = " & Round(rs_aux7!cotiza_precio_base_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(rs_aux7!cotiza_precio_total_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol_cli), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(rs_aux7!cotiza_precio_total_bs_cli, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol_cge), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(rs_aux7!cotiza_precio_total_bs_cge, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(CDbl(rs_aux7!cotiza_gasto_local_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(rs_aux7!cotiza_gasto_local_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_local_IT_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & Round(rs_aux7!cotiza_saldo_local_IT_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_local_IVA_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & Round(rs_aux7!cotiza_saldo_local_IVA_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_cge_IT_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & Round(rs_aux7!cotiza_saldo_cge_IT_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_cge_IVA_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & Round(rs_aux7!cotiza_saldo_cge_IVA_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_tac_billing_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & Round(rs_aux7!cotiza_saldo_tac_billing_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
'
'                    'WWWWWWWWWWWWWWWWWWWWWW
'                End If
'             End If
'             'GRABA costo_monto DE TODOS LOS REGISTROS
'             Set rs_aux5 = New ADODB.Recordset
'             If rs_aux5.State = 1 Then rs_aux5.Close
'             rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   ", db, adOpenKeyset, adLockOptimistic
'             If rs_aux5.RecordCount = 0 Then
'                Call GRABA_COSTOS
'             Else
'                sino = MsgBox("La Hoja de Costos ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
'                If sino = vbYes Then
'                    'OJO BORRAR ao_solicitud_costos
'                    db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
'                    'corrprog = 0
'                    Call GRABA_COSTOS
'                Else
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_NAC_ME = rs_aux6!costo_monto2
'                        VAR_NAC = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_ALM_ME = rs_aux6!costo_monto2
'                        VAR_ALM = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_AGE_ME = rs_aux6!costo_monto2
'                        VAR_AGE = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_FLE_ME = IIf(IsNull(rs_aux6!costo_monto2), "0", rs_aux6!costo_monto2)
'                        VAR_FLE = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_UTIL_ME = IIf(IsNull(rs_aux6!costo_monto2), "0", rs_aux6!costo_monto2)
'                        VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                    End If
'                End If
'
'             End If
'             'GRABA TOTALES E IMPUESTOS
'             If mw_solicitud_cotiza_venta.Ado_datosE.Recordset!pais_continente = "EUROPA" Then
'                    Set rs_aux4 = New ADODB.Recordset
'                    If rs_aux4.State = 1 Then rs_aux4.Close
'                    rs_aux4.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl, sum(costo_monto2) as toteu  from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA'   ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux4.RecordCount > 0 Then
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_me = round(" & rs_aux4!toteu & " + cotiza_precio_base_me  - cotiza_precio_flete_me, " & CDbl(cmd_dec) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = round(cotiza_precio_total_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = round(cotiza_precio_total_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    Else
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_me = round(" & rs_aux4!toteu & " + cotiza_precio_base_me  - cotiza_precio_flete_me, " & CDbl(cmd_dec) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = round(cotiza_precio_total_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = round(cotiza_precio_total_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    End If
'                    'Importaion Cliente
'                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_me = " & rs_aux4!toteu & " - " & VAR_NAC_ME & " - " & VAR_ALM_ME & " - " & VAR_AGE_ME & " - " & VAR_FLE_ME & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_me = " & rs_aux4!toteu & " - " & VAR_NAC_ME & " - " & VAR_ALM_ME & " - " & VAR_AGE_ME & " - " & VAR_FLE_ME & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = round(cotiza_gasto_local_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = round(cotiza_gasto_local_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bs.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = cotiza_gasto_local_dol * " & CDbl(txt_local_IT_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = cotiza_gasto_local_dol * " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = cotiza_precio_total_dol + cotiza_saldo_local_IT_dol + cotiza_saldo_local_IVA_dol where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = cotiza_precio_total_dol_cli * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
'
'                    VAR_DOLCLI = rs_aux4!totdl + mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_cif_dol - mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_fob_dol - mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_seg_dol
'                    VAR_BSCLI = rs_aux4!totbs + mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_cif_bs - mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_fob_bs - mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_seg_bs
'                    'no suma Ado_datos1A.Recordset!cotiza_precio_total_dol JQA-2015
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
'                    'VAR_SUBD = mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_total_dol - mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_seg_dol    'sin SEGURO
'                    VAR_SUBD = mw_solicitud_cotiza_venta.Ado_datos1E.Recordset!cotiza_precio_total_dol         'mas SEGURO
'                    VAR_SUBB = VAR_SUBD * GlTipoCambioOficial
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IT_bs) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_base_dol * 0.1498) )-((" & CDbl(VAR_AGE) & " * 0.13))  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_cif_dol * 0.1498) * " & CDbl(dtc_desc16) & ")-((" & CDbl(VAR_AGE) & " * 0.13)* " & CDbl(dtc_desc16) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(IIf(txt_tac_billing_bs = "", "0", txt_tac_billing_bs)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = (" & VAR_SUBD & " * " & CDbl(IIf(txt_tac_billing_bs = "", "1", txt_tac_billing_bs)) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & VAR_SUBD & "  + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IT_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IVA_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_tac_billing_dol) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = ao_solicitud_cotiza_venta.cotiza_precio_total_dol_cge * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'
'             End If
'           mw_solicitud_cotiza_venta.Ado_datos1E.Recordset.MoveNext
'           Wend
       'Else
             '- SOLO EL REGISTRO ACTIVO
             'WWWWWWWWWWWWWWWWW
             'GUARDA EL PRIMER REGISTRO
             'WWWWWWWWWWWWWWWWWW
             mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_dec = CDbl(cmd_dec.Text)
             mw_solicitud_cotiza_venta.Ado_datosE.Recordset!tipo_moneda = cmd_moneda.Text
             If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
                Txt_tdc.Text = GlTipoCambioOficial
             End If
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_tdc_me = txt_tdc_me.Text
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_tdc_bol = GlTipoCambioOficial
            
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!costo_monto = txt_montobase.Text
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_me = IIf(txt_fob_eu = "", "0", Round(CDbl(txt_fob_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_bs = IIf(txt_fob_bs = "", "0", Round(CDbl(txt_fob_bs), Val(cmd_dec)))      'Round(CDbl(txt_fob_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_dol = IIf(txt_fob_me = "", "0", Round(CDbl(txt_fob_me), Val(cmd_dec)))     'Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
            
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_dcto_me = IIf(txt_dcto_eu = "", "0", Round(CDbl(txt_dcto_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_dcto_bs = IIf(txt_dcto_bs = "", "0", Round(CDbl(txt_dcto_bs), Val(cmd_dec)))       'Round(CDbl(txt_dcto_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_dcto_dol = IIf(txt_dcto_me = "", "0", Round(CDbl(txt_dcto_me), Val(cmd_dec)))      'Round(CDbl(txt_dcto_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
             
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_tacb_me = IIf(txt_tacb_eu = "", "0", Round(CDbl(txt_tacb_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_tacb_bs = IIf(txt_tacb_bs = "", "0", Round(CDbl(txt_tacb_bs), Val(cmd_dec)))       'Round(CDbl(txt_tacb_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_tacb_dol = IIf(txt_tacb_me = "", "0", Round(CDbl(txt_tacb_me), Val(cmd_dec)))      'Round(CDbl(txt_tacb_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
             
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_spread_me = IIf(txt_spread_eu = "", "0", Round(CDbl(txt_spread_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_spread_bs = IIf(txt_spread_bs = "", "0", Round(CDbl(txt_spread_bs), Val(cmd_dec)))     'Round(CDbl(txt_spread_me) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_spread_dol = IIf(txt_spread_me = "", "0", Round(CDbl(txt_spread_me), Val(cmd_dec)))    'Round(CDbl(txt_spread_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_seg_me = IIf(txt_seguro_eu = "", "0", Round(CDbl(txt_seguro_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_seg_bs = IIf(txt_seguro_bs = "", "0", Round(CDbl(txt_seguro_bs), Val(cmd_dec)))        'Round(CDbl(txt_seguro_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_seg_dol = IIf(txt_seguro_me = "", "0", Round(CDbl(txt_seguro_me), Val(cmd_dec)))       'Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_fob_seg_me = IIf(txt_fob_seg_eu = "", "0", Round(CDbl(txt_fob_seg_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_fob_seg_bs = IIf(txt_fob_seg_bs = "", "0", Round(CDbl(txt_fob_seg_bs), Val(cmd_dec)))     'Round(CDbl(txt_fob_me) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_fob_seg_dol = IIf(txt_fob_seg_dol = "", "0", Round(CDbl(txt_fob_seg_dol), Val(cmd_dec)))    'Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_flete_me = IIf(txt_fletefrontera_eu = "", "0", Round(CDbl(txt_fletefrontera_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_flete_bs = IIf(txt_fletefrontera_bs = "", "0", Round(CDbl(txt_fletefrontera_bs), Val(cmd_dec)))        'Round(CDbl(txt_fletefrontera_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_flete_dol = IIf(txt_fletefrontera_me = "", "0", Round(CDbl(txt_fletefrontera_me), Val(cmd_dec)))       'Round(CDbl(txt_fletefrontera_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_cif_me = IIf(txt_cif_eu = "", "0", Round(CDbl(txt_cif_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_cif_bs = IIf(txt_cif_bs = "", "0", Round(CDbl(txt_cif_bs), Val(cmd_dec)))      'Round(CDbl(txt_cif_eu) * CDbl(txt_tdc_me), Val(cmd_dec)) '
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_cif_dol = IIf(txt_cif_me = "", "0", Round(CDbl(txt_cif_me), Val(cmd_dec)))     'Round(CDbl(txt_cif_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_GAC_me = IIf(txt_gac_eu = "", "0", Round(CDbl(txt_gac_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_GAC_bs = IIf(txt_gac_bs = "", "0", Round(CDbl(txt_gac_bs), Val(cmd_dec)))      'Round(CDbl(txt_gac_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_GAC_dol = IIf(txt_gac_me = "", "0", Round(CDbl(txt_gac_me), Val(cmd_dec)))     'Round(CDbl(txt_gac_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_base_me = IIf(txt_base_imp_eu = "", "0", Round(CDbl(txt_base_imp_eu), Val(cmd_dec)))
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_base_bs = IIf(txt_base_imp_bs = "", "0", Round(CDbl(txt_base_imp_bs), Val(cmd_dec)))       'Round(CDbl(txt_base_imp_eu) * CDbl(txt_tdc_me), Val(cmd_dec)) '
            mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_base_dol = IIf(txt_base_imp_me = "", "0", Round(CDbl(txt_base_imp_me), Val(cmd_dec)))      'Round(CDbl(txt_base_imp_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

             mw_solicitud_cotiza_venta.Ado_datosE.Recordset!fecha_registro = Date     'no cambia
             mw_solicitud_cotiza_venta.Ado_datosE.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
             mw_solicitud_cotiza_venta.Ado_datosE.Recordset.Update    'Batch 'adAffectAll
             db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'NO' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
             
             'GRABA COSTOS
             Set rs_aux5 = New ADODB.Recordset
             If rs_aux5.State = 1 Then rs_aux5.Close
             rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   ", db, adOpenKeyset, adLockOptimistic
             If rs_aux5.RecordCount = 0 Then
                Call GRABA_COSTOS
             Else
                sino = MsgBox("La Hoja de Costos ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
                If sino = vbYes Then
                    'OJO BORRAR ao_solicitud_costos
                    db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
                    'corrprog = 0
                    Call GRABA_COSTOS
                Else
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        'VAR_NAC = rs_aux6!costo_monto_usd
                        VAR_NAC = rs_aux6!costo_monto2
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        'VAR_ALM = rs_aux6!costo_monto_usd
                        VAR_ALM = rs_aux6!costo_monto2
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        'VAR_AGE = rs_aux6!costo_monto_usd
                        VAR_AGE = rs_aux6!costo_monto2
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        'VAR_FLE = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
                        VAR_FLE = IIf(IsNull(rs_aux6!costo_monto2), "0", rs_aux6!costo_monto2)
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        'VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
                        VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto2), "0", rs_aux6!costo_monto2)
                    End If
                End If
        
             End If
             If mw_solicitud_cotiza_venta.Ado_datosE.Recordset!pais_continente = "EUROPA" And mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_codigo = Val(Txt_Correl.Caption) Then
                    Set rs_aux4 = New ADODB.Recordset
                    If rs_aux4.State = 1 Then rs_aux4.Close
                    rs_aux4.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl, sum(costo_monto2) as toteu from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' and cotiza_codigo = " & CDbl(Txt_Correl) & "  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux4.RecordCount > 0 Then
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_me = round(" & rs_aux4!toteu & " + cotiza_precio_base_me  - cotiza_precio_flete_me, " & CDbl(cmd_dec) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = round(cotiza_precio_total_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = round(cotiza_precio_total_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    Else
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_me = round(" & rs_aux4!toteu & " + cotiza_precio_base_me  - cotiza_precio_flete_me, " & CDbl(cmd_dec) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = round(cotiza_precio_total_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = round(cotiza_precio_total_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    End If
                    'Importacion Cliente
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_me = " & rs_aux4!toteu & " - " & VAR_NAC & " - " & VAR_ALM & " - " & VAR_AGE & " - " & VAR_FLE & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = round(cotiza_gasto_local_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = round(cotiza_gasto_local_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bs.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = round(cotiza_gasto_local_dol * " & CDbl(txt_local_IT_bs.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = round(cotiza_gasto_local_dol * " & CDbl(txt_local_IVA_bs.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = cotiza_precio_total_dol + cotiza_saldo_local_IT_dol + cotiza_saldo_local_IVA_dol where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = cotiza_precio_total_dol_cli * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
                    'Facturacion Local
                    'VAR_DOLCLI = rs_aux4!totdl + Ado_datos1A.Recordset!cotiza_precio_cif_dol - Ado_datos1A.Recordset!cotiza_precio_fob_dol - Ado_datos1A.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = rs_aux4!totbs + Ado_datos1A.Recordset!cotiza_precio_cif_bs - Ado_datos1A.Recordset!cotiza_precio_fob_bs - Ado_datos1A.Recordset!cotiza_precio_seg_bs
                    
                    VAR_BSCLI = rs_aux4!totbs - (VAR_FLE * CDbl(txt_tdc_me.Text))
                    VAR_DOLCLI = VAR_BSCLI / GlTipoCambioOficial
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
                    VAR_SUBD = IIf(IsNull(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_total_dol), 0, mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_total_dol) - mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_seg_dol
                    VAR_SUBB = VAR_SUBD * GlTipoCambioOficial
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IT_bs) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_cif_dol * 0.1498) )-((" & CDbl(VAR_AGE) & " * 0.13))  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_cif_dol * 0.1498) * " & CDbl(dtc_desc16) & ")-((" & CDbl(VAR_AGE) & " * 0.13)* " & CDbl(dtc_desc16) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(IIf(txt_tac_billing_bs = "", "0", txt_tac_billing_bs)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = (" & VAR_SUBD & " * " & CDbl(IIf(txt_tac_billing_bs = "", "1", txt_tac_billing_bs)) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & VAR_SUBD & "  + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IT_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IVA_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_tac_billing_dol) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = ao_solicitud_cotiza_venta.cotiza_precio_total_dol_cge * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
    
             End If
       '  End If

     End If
'     mw_solicitud_cotiza_venta.sstab1.Tab = 2
''     If Ado_datos1A.Recordset!pais_continente = "EUROPA" Then
'     If VAR_CONTI = "AMERICA" Then
'         mw_solicitud_cotiza_venta.sstab1.TabEnabled(0) = True
'     Else
'        mw_solicitud_cotiza_venta.sstab1.TabEnabled(0) = False
'     End If
'     If VAR_CONTI = "ASIA" Then
'        mw_solicitud_cotiza_venta.sstab1.TabEnabled(1) = True
'     Else
'        mw_solicitud_cotiza_venta.sstab1.TabEnabled(1) = False
'     End If
'     If VAR_CONTI = "EUROPA" Then
'        mw_solicitud_cotiza_venta.sstab1.TabEnabled(2) = True
'     Else
'        mw_solicitud_cotiza_venta.sstab1.TabEnabled(2) = False
'     End If
'     Call ABRIR_TABLA
''     rs_datosA.MoveLast
''     mbDataChanged = False
''        Fra_datos.Enabled = False
'        FraModeloCostoA.Enabled = False
'        Fra_datos2.Enabled = False
'        fraOpciones2A.Visible = True
'        fraOpciones1A.Visible = True
'        FraGrabarCancelarA.Visible = False
'        FrmABMDet.Visible = True
'        FraDet1E.Enabled = True
'        dg_datosA.Enabled = True
'        dg_datos1A.Enabled = True
'        VAR_SW = ""
''        SSTab1.Tab = 1
''        SSTab1.TabEnabled(0) = False
''        SSTab1.TabEnabled(1) = True
''        SSTab1.TabEnabled(2) = False
''     dtc_codigo9.Enabled = True
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Unload Me
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub GRABA_COSTOS()
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    VAR_CONTI = "EUROPA"
    If VAR_CONTI = "AMERICA" Then
        rs_datos3.Open "select * from ac_costos_comercializacion where costo_tipo= 'B' ", db, adOpenStatic
    End If
    If VAR_CONTI = "ASIA" Then
        rs_datos3.Open "select * from ac_costos_comercializacion where costo_tipoA= 'B' ", db, adOpenStatic
    End If
    If VAR_CONTI = "EUROPA" Then
        rs_datos3.Open "select * from ac_costos_comercializacion where costo_tipoE= 'B' ", db, adOpenStatic
    End If
    Set Ado_datos3.Recordset = rs_datos3
    If Ado_datos3.Recordset.RecordCount > 0 Then
        Ado_datos3.Recordset.MoveFirst
        While Not Ado_datos3.Recordset.EOF
            'codigo_costo
            'costo_descripcion
            'costo_monto
            'costo_porcentaje
            'costo_tipo
            Set rs_aux5 = New ADODB.Recordset
            If rs_aux5.State = 1 Then rs_aux5.Close
            rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " ", db, adOpenKeyset, adLockOptimistic      'AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "
            'If rs_aux5.RecordCount = 0 Then
                rs_aux5.AddNew
                rs_aux5!ges_gestion = Year(Date)
                rs_aux5!unidad_codigo = parametro           'Txt_campo1.Caption
                rs_aux5!solicitud_codigo = GlSolicitud      'Ado_datos.Recordset!solicitud_codigo
                rs_aux5!edif_codigo = GlEdificio            'Ado_datos.Recordset!edif_codigo
                rs_aux5!cotiza_codigo = Txt_Correl         'Ado_datos.Recordset!cotiza_codigo
                rs_aux5!pais_continente = VAR_CONTI
                rs_aux5!estado_codigo = "REG"
                rs_aux5!codigo_costo = Ado_datos3.Recordset!codigo_costo
                rs_aux5!costo_porcentaje = Ado_datos3.Recordset!costo_porcentaje
                rs_aux5!costo_monto = 0
                rs_aux5!costo_monto2 = 0
                rs_aux5!costo_monto_usd = 0
                If Ado_datos3.Recordset!costo_porcentaje > 0 Then
'                    If VAR_CONTI = "AMERICA" Then
'                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
'                            rs_aux5!costo_monto_usd = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto_usd * CDbl(GlTipoCambioOficial)), CDbl(cmd_dec))
'                        Else
'                            rs_aux5!costo_monto_usd = Round(CDbl(Ado_datos1.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto_usd * CDbl(GlTipoCambioOficial)), CDbl(cmd_dec))
'                        End If
'                    End If
'                    If VAR_CONTI = "ASIA" Then
'                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
'                            rs_aux5!costo_monto = Round(CDbl((mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_bs + mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_spread_bs) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto_usd = Round(CDbl((mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_dol + mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_spread_dol) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                        Else
'                            'rs_aux5!costo_monto = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_cif_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            'rs_aux5!costo_monto_usd = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_base_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto_usd = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_base_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                        End If
'                    End If
                    If VAR_CONTI = "EUROPA" Then
                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
                            rs_aux5!costo_monto2 = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_fob_seg_me * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto2) * CDbl(txt_tdc_me.Text), CDbl(cmd_dec))
                            rs_aux5!costo_monto_usd = Round(CDbl(rs_aux5!costo_monto) / CDbl(GlTipoCambioOficial), CDbl(cmd_dec))
                        Else
                            rs_aux5!costo_monto2 = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_base_me * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto2) * CDbl(txt_tdc_me.Text), CDbl(cmd_dec))
                            rs_aux5!costo_monto_usd = Round(CDbl(rs_aux5!costo_monto) / CDbl(GlTipoCambioOficial), CDbl(cmd_dec))
                        End If
                    End If
'                    rs_aux5!costo_monto2 = 0    'Round(CDbl(IIf(txt_total_bs1.Text = "", "0", txt_total_bs1.Text)), 2)
'                    rs_aux5!costo_monto_usd2 = 0    'Round(CDbl(txt_total_me1.Text), 2)
'                    rs_aux5!costo_monto3 = 0    'Round(CDbl(IIf(txt_dcto_bs1.Text = "", "0", txt_dcto_bs1.Text)), 2)
'                    rs_aux5!costo_monto_usd3 = 0    'Round(CDbl(txt_dcto_me1.Text), 2)
                Else
                    'abrir tabla costos_paradas
                    Set rs_datos9 = New ADODB.Recordset
                    If rs_datos9.State = 1 Then rs_datos9.Close
                    'rs_datos9.Open "SELECT * FROM ac_costos_paradas where trafico_num_paradas = " & Val(txt_paradas.Text) & " ", db, adOpenStatic
                    rs_datos9.Open "ac_costos_paradas where trafico_num_paradas = " & Val(txt_paradas.Text) & " ", db, adOpenStatic
                    Set Ado_datos9.Recordset = rs_datos9
                    If Ado_datos9.Recordset.RecordCount > 0 Then
                        If Ado_datos3.Recordset!codigo_costo = 9 Then
                            rs_aux5!costo_monto2 = Round(CDbl(rs_datos9!costo_instal_pintura_eu), CDbl(cmd_dec))
                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto2) * CDbl(txt_tdc_me.Text), CDbl(cmd_dec))
                            rs_aux5!costo_monto_usd = Round(CDbl(rs_aux5!costo_monto) / CDbl(GlTipoCambioOficial), CDbl(cmd_dec))
                        End If
                        If Ado_datos3.Recordset!codigo_costo = 11 Then
                            If VAR_CONTI = "AMERICA" Then
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs), 2) * CDbl(Txt_campo5.Text)
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd), 2) * CDbl(Txt_campo5.Text)
                            End If
                            If VAR_CONTI = "ASIA" Then
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs), 2) * CDbl(Txt_campo5A.Text)
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd), 2) * CDbl(Txt_campo5A.Text)
                            End If
                            If VAR_CONTI = "EUROPA" Then
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs), CDbl(cmd_dec)) * CDbl(Txt_campo5.Text)
                                rs_aux5!costo_monto2 = Round(CDbl(rs_aux5!costo_monto) / CDbl(txt_tdc_me.Text), CDbl(cmd_dec))
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_aux5!costo_monto) / CDbl(GlTipoCambioOficial), CDbl(cmd_dec))
                            End If
                        End If
                        If Ado_datos3.Recordset!codigo_costo = 12 Then
                            rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_ajuste_bs), CDbl(cmd_dec))
                            rs_aux5!costo_monto2 = Round(CDbl(rs_aux5!costo_monto) / CDbl(txt_tdc_me.Text), CDbl(cmd_dec))
                            rs_aux5!costo_monto_usd = Round(CDbl(rs_aux5!costo_monto) / CDbl(GlTipoCambioOficial), CDbl(cmd_dec))
                        End If
                    End If
                End If
                If Ado_datos3.Recordset!codigo_costo = 3 Then   'NACIONALIZACION
                    VAR_NAC = rs_aux5!costo_monto_usd
                    VAR_NAC_ME = rs_aux5!costo_monto2
                End If
                If Ado_datos3.Recordset!codigo_costo = 5 Then   'ALMACENAJE
                    VAR_ALM = rs_aux5!costo_monto_usd
                    VAR_ALM_ME = rs_aux5!costo_monto2
                End If
                If Ado_datos3.Recordset!codigo_costo = 6 Then   'COMISION AGENCIA
                    VAR_AGE = rs_aux5!costo_monto_usd
                    VAR_AGE_ME = rs_aux5!costo_monto2
                End If
                If Ado_datos3.Recordset!codigo_costo = 8 Then   'TOTAL FLETES
                    VAR_FLE = IIf(IsNull(rs_aux5!costo_monto_usd), "0", rs_aux5!costo_monto_usd)
                    VAR_FLE_ME = IIf(IsNull(rs_aux5!costo_monto2), "0", rs_aux5!costo_monto2)
                End If
                If VAR_CONTI = "EUROPA" Then
                    'VAR_DOLCLI = Ado_datos.Recordset!cotiza_precio_total_dol - Ado_datos.Recordset!cotiza_precio_fob_dol - Ado_datos.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = Ado_datos.Recordset!cotiza_precio_total_bs - Ado_datos.Recordset!cotiza_precio_fob_bs - Ado_datos.Recordset!cotiza_precio_seg_bs
                End If
                rs_aux5!costo_observaciones = Trim(Ado_datos3.Recordset!costo_descripcion)

                rs_aux5!fecha_registro = Date
                'aw_p_ao_negociacion_cabecera.Ado_detalle1.Recordset("hora_registro").Value = Date
                rs_aux5!usr_codigo = glusuario
                rs_aux5.Update
            'End If
            Ado_datos3.Recordset.MoveNext
        Wend
    End If
End Sub

Private Sub AcumulaMonto(ges, uni, cod1, cod2)

'
'  If rs_aux1.State = 1 Then rs_aux1.Close
  
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    'rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND cotiza_codigo = " & rs_datos!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic
    rs_aux4.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl, sum(costo_monto2) as toteu  from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2, db, adOpenKeyset, adLockOptimistic
    If rs_aux4.RecordCount > 0 Then
        'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & rs_aux4!totdl & " + ao_solicitud_cotiza_venta.cotiza_precio_total_dol_x + ao_solicitud_cotiza_venta.cotiza_precio_fob_dol_x - ao_solicitud_cotiza_venta.cotiza_precio_total_dol_h   where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & rs_aux4!totbs & "  + ao_solicitud_cotiza_venta.cotiza_precio_total_bs_x + ao_solicitud_cotiza_venta.cotiza_precio_fob_bs_x - ao_solicitud_cotiza_venta.cotiza_precio_total_bs_h where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & rs_aux4!totdl & " + ao_solicitud_cotiza_venta.cotiza_precio_total_dol_x  - ao_solicitud_cotiza_venta.cotiza_precio_total_dol_h   where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & rs_aux4!totbs & "  + ao_solicitud_cotiza_venta.cotiza_precio_total_bs_x - ao_solicitud_cotiza_venta.cotiza_precio_total_bs_h where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
    End If
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2, db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        'VAR_DOLCLI = rs_aux4!totdl + rs_aux1!cotiza_precio_total_dol_x - rs_aux1!cotiza_precio_fob_dol - rs_aux1!cotiza_precio_fob_dol_x
        'VAR_BSCLI = rs_aux4!totbs + rs_aux1!cotiza_precio_total_bs_x - rs_aux1!cotiza_precio_fob_bs - rs_aux1!cotiza_precio_fob_bs_x
        
        VAR_DOLCLI = rs_aux4!totdl + rs_aux1!cotiza_precio_total_dol_x - rs_aux1!cotiza_precio_fob_dol - rs_aux1!cotiza_precio_fob_dol_x
        VAR_BSCLI = rs_aux4!totbs + rs_aux1!cotiza_precio_total_bs_x - rs_aux1!cotiza_precio_fob_bs - rs_aux1!cotiza_precio_fob_bs_x
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_NAC = rs_aux2!costo_monto_usd
        End If
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_ALM = rs_aux2!costo_monto_usd
        End If
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_AGE = rs_aux2!costo_monto_usd
        End If
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_FLE = IIf(IsNull(rs_aux2!costo_monto_usd), "0", rs_aux2!costo_monto_usd)
        End If
        'VAR_SUBD = VAR_DOLCLI - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE       'rs_aux1!cotiza_precio_total_dol +
        'VAR_SUBB = VAR_SUBD * GlTipoCambioOficial
        VAR_SUBD = rs_aux4!totdl - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE       'rs_aux1!cotiza_precio_total_dol +
        VAR_SUBB = VAR_SUBD * GlTipoCambioOficial
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = cotiza_precio_total_dol  + (" & VAR_SUBD & " * 0.0309) + (" & VAR_SUBD & " * 0.1491) where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = cotiza_precio_total_dol_cli * " & GlTipoCambioOficial & " where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = cotiza_precio_total_dol  + ((cotiza_precio_total_dol - cotiza_precio_fob_dol_x) * 0.0416) + ((cotiza_precio_total_dol - cotiza_precio_fob_dol_x) * 0.16) - ((cotiza_precio_total_dol_x * 0.1498) * " & Val(mw_solicitud_cotiza_venta.dtc_desc15) & " - ((" & VAR_AGE & " * 0.13)* " & Val(mw_solicitud_cotiza_venta.dtc_desc15) & " ) ) + ((cotiza_precio_total_dol - cotiza_precio_fob_dol_x) * 0.0350) where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = cotiza_precio_total_dol_cge * " & GlTipoCambioOficial & " where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
    End If
End Sub

Private Sub valida_campos()
  If (txt_fob_me = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_dcto_me.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_tacb_me.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_spread_me.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_seguro_me = "") Then
    MsgBox "Debe registrar ... " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_fob_seg_dol.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_fletefrontera_me.Text = "" Then
    MsgBox "Debe registrar... " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_cif_me.Text = "" Then
    MsgBox "Debe registrar... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub dtc_desc1_LostFocus()
    Txt_campo3.Text = dtc_aux1.Text
    
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl, sum(costo_monto2) as toteu, sum(costo_monto2) as totbs2, sum(costo_monto_usd2) as totdl2, sum(costo_monto3) as totbs3, sum(costo_monto_usd3) as totdl3 from ao_solicitud_costos where ges_gestion = '" & Year(Date) & "' and unidad_codigo = '" & txt_campo1 & "' and solicitud_codigo = '" & txt_codigo & "' and edif_codigo = '" & Txt_campo2 & "' and cotiza_codigo = " & Txt_Correl & "  ", db, adOpenKeyset, adLockOptimistic
    
    Select Case dtc_codigo1.Text
        Case 1
            'SEGURO DE TRANSPORTE   0.0078
            Txt_monto1.Text = CDbl(txt_monto01) * CDbl(Txt_campo3)
            txt_monto3.Text = CDbl(txt_monto02) * CDbl(Txt_campo3)
            Txt_monto5.Text = CDbl(txt_monto03) * CDbl(Txt_campo3)
            
        Case 2
            'FLETE FRONTERA
            Txt_monto1.Text = Dtc_aux2.Text
            txt_monto3.Text = Dtc_aux2.Text
            Txt_monto5.Text = Dtc_aux2.Text
            
        Case 3
            'NACIONALIZACION 0.1498
            'sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl, sum(costo_monto2) as totbs2, sum (costo_monto_usd2) as totdl2, sum(costo_monto3) as totbs3, sum (costo_monto_usd3) as totdl3
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
                        
        Case 4
            'GAC 0.05
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 5
            'ALMACENAJE 0.007
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 6
            'COMISION AGENCIA       0.015
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 7
            'SPREAD GLOBAL  0.08
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 8
            'TOTAL FLETES
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 9
            'INSTALACION Y PINTURA
            'corregrirrrrrrrrrrrrrrrrrrrrrrrrrrrr
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 10
            'COSTOS DE OPERACION
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
'            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
'            txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 11
            'COSTOS DE INSTALACION INTERIOR
            'corregrirrrrrrrrrrrrrrrrrrrrrrrrrrrr
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 12
            'COSTOS DE AJUSTE INTERIOR
            'corregrirrrrrrrrrrrrrrrrrrrrrrrrrrrr
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 13
            'IMPREVISTOS COMISIONES
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 14
            'UTILIDAD 0.15
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 15
            'OTROS
    End Select
        
    If rs_aux1.State = 1 Then rs_aux1.Close
End Sub

Private Sub cmd_moneda_LostFocus()
    If cmd_moneda.Text = "EUR" Then
        txt_tdc_me.Text = GlTipoCambioEuro
        'CAMBIA DE FONDO
        txt_fob_me.backColor = &HC0C0C0
        txt_dcto_me.backColor = &HC0C0C0
        txt_tacb_me.backColor = &HC0C0C0
        txt_spread_me.backColor = &HC0C0C0
        txt_seguro_me.backColor = &HC0C0C0
        'txt_fob_seg_dol.backColor = &HC0C0C0
        txt_fletefrontera_me.backColor = &HC0C0C0
        'txt_cif_me.backColor = &HC0C0C0
        txt_gac_me.backColor = &HC0C0C0
        'txt_base_imp_me.backColor = &HC0C0C0
        
        txt_fob_eu.backColor = &HFFFFFF
        txt_dcto_eu.backColor = &HFFFFFF
        txt_tacb_eu.backColor = &HFFFFFF
        txt_spread_eu.backColor = &HFFFFFF
        txt_seguro_eu.backColor = &HFFFFFF
        'txt_fob_seg_eu.backColor = &HFFFFFF
        txt_fletefrontera_eu.backColor = &HFFFFFF
        'txt_cif_eu.backColor = &HFFFFFF
        txt_gac_eu.backColor = &HFFFFFF
        'txt_base_imp_eu.backColor = &HFFFFFF
    Else
        txt_tdc_me.Text = GlTipoCambioOficial
        'CAMBIA DE FONDO
        txt_fob_me.backColor = &HFFFFFF
        txt_dcto_me.backColor = &HFFFFFF
        txt_tacb_me.backColor = &HFFFFFF
        txt_spread_me.backColor = &HFFFFFF
        txt_seguro_me.backColor = &HFFFFFF
        'txt_fob_seg_dol.backColor = &HFFFFFF
        txt_fletefrontera_me.backColor = &HFFFFFF
        'txt_cif_me.backColor = &HFFFFFF
        txt_gac_me.backColor = &HFFFFFF
        'txt_base_imp_me.backColor = &HFFFFFF
        
        txt_fob_eu.backColor = &HC0C0C0
        txt_dcto_eu.backColor = &HC0C0C0
        txt_tacb_eu.backColor = &HC0C0C0
        txt_spread_eu.backColor = &HC0C0C0
        txt_seguro_eu.backColor = &HC0C0C0
        'txt_fob_seg_eu.backColor = &HC0C0C0
        txt_fletefrontera_eu.backColor = &HC0C0C0
        'txt_cif_eu.backColor = &HC0C0C0
        txt_gac_eu.backColor = &HC0C0C0
        'txt_base_imp_eu.backColor = &HC0C0C0
    End If
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    mbDataChanged = False
    If txt_tdc_me.Text = "" Or txt_tdc_me.Text = "0" Then
       txt_tdc_me.Text = GlTipoCambioEuro
    End If
    If txt_tacb_bs.Text = "" Or txt_tacb_bs.Text = "0" Then
       txt_tacb_bs.Text = "0.017"
    End If
    If txt_spread_bs.Text = "" Or txt_spread_bs.Text = "0" Then
       txt_spread_bs.Text = "0.021"
    End If
    If txt_local_IT_bs.Text = "" Or txt_local_IT_bs.Text = "0" Then
       txt_local_IT_bs.Text = "0.0309"
    End If
    If txt_local_IVA_bs.Text = "" Or txt_local_IVA_bs.Text = "0" Then
       txt_local_IVA_bs.Text = "0.1491"
    End If
    If txt_cge_IT_bs.Text = "" Or txt_cge_IT_bs.Text = "0" Then
       txt_cge_IT_bs = "0.0416"
    End If
    If txt_cge_IVA_bs.Text = "" Or txt_cge_IVA_bs.Text = "0" Then
       txt_cge_IVA_bs = "0.151"
    End If
    If txt_tac_billing_bs.Text = "" Or txt_tac_billing_bs.Text = "0" Then
       txt_tac_billing_bs = "0.035"
    End If
    If txt_gac_bs = "" Or txt_gac_bs = "0" Then
       txt_gac_bs = "0.05"
    End If
    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_me = " & rs_aux4!toteu & " - " &
    VAR_NAC = "0"
    VAR_ALM = "0"
    VAR_AGE = "0"
    VAR_FLE = "0"
End Sub

Private Sub Form_Load()
'    Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from ac_costos_comercializacion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'    'gc_pais
'    Set rs_datos7 = New ADODB.Recordset
'    If rs_datos7.State = 1 Then rs_datos7.Close
'    If mw_solicitud_cotiza_venta.sstab1.Tab = 0 Then
'        rs_datos7.Open "Select * from gc_pais where pais_continente = 'AMERICA' order by pais_descripcion", db, adOpenStatic
'    End If
'    If mw_solicitud_cotiza_venta.sstab1.Tab = 1 Then
'        rs_datos7.Open "Select * from gc_pais where pais_continente = 'EUROPA' order by pais_descripcion", db, adOpenStatic
'    End If
'    If mw_solicitud_cotiza_venta.sstab1.Tab = 2 Then
'        rs_datos7.Open "Select * from gc_pais where pais_continente = 'EUROPA' order by pais_descripcion", db, adOpenStatic
'    End If
'    Set Ado_datos7.Recordset = rs_datos7
''    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

'Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub txt_dcto_eu_LostFocus()
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
     End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_dcto_eu = "" Then
        txt_dcto_bs.Text = "0"
        txt_dcto_me.Text = "0"
        txt_dcto_eu.Text = "0"
    Else
        txt_dcto_bs.Text = Round(CDbl(txt_dcto_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
        txt_dcto_me.Text = Round(CDbl(txt_dcto_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
        
    End If
    txt_seguro_eu.Text = Round((CDbl(txt_fob_eu) - CDbl(txt_dcto_eu.Text) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_seguro_me.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    
    txt_fob_seg_eu = Round(CDbl(txt_seguro_eu) + CDbl(txt_fob_eu) - CDbl(txt_dcto_eu) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_fob_seg_dol = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
End Sub

Private Sub txt_dcto_me_LostFocus()
'    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
'        Txt_tdc.Text = GlTipoCambioOficial
'     End If
'     If txt_dcto_me = "" Then
'        txt_dcto_bs.Text = "0"
'     Else
'        txt_dcto_bs.Text = CDbl(txt_dcto_me) * CDbl(GlTipoCambioOficial)
'        txt_seguro_bs.Text = Round((CDbl(txt_fob_bs1A) - CDbl(txt_dcto_bs.Text)) * 0.0078, Val(cmd_dec)) '+ 1
'        txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text)) * 0.0078, Val(cmd_dec)) '+ 1
'        txt_tacb_me = Round(CDbl(txt_fob_me) * CDbl(txt_tacb_bs), Val(cmd_dec))
'        txt_spread_me = Round(CDbl(txt_fob_me) * CDbl(txt_spread_bs), Val(cmd_dec))
'     End If
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
     End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_dcto_me = "" Then
        txt_dcto_bs.Text = "0"
        txt_dcto_me.Text = "0"
        txt_dcto_eu.Text = "0"
    Else
        txt_dcto_bs.Text = Round(CDbl(txt_dcto_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_dcto_eu.Text = Round(CDbl(txt_dcto_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
        
    End If
    txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_tacb_me) + CDbl(txt_spread_me)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_seguro_eu.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    
    txt_fob_seg_dol = Round(CDbl(txt_seguro_me) + CDbl(txt_fob_me) - CDbl(txt_dcto_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_fob_seg_eu = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))

End Sub

Private Sub txt_fletefrontera_eu_LostFocus()
    If txt_fletefrontera_eu.Text = "" Then
        txt_fletefrontera_bs.Text = "0"  'GlTipoCambioOficial
        txt_fletefrontera_me.Text = "0"
        txt_fletefrontera_eu.Text = "0"
    Else
        txt_fletefrontera_bs.Text = Round(CDbl(txt_fletefrontera_eu) * CDbl(txt_tdc_me), Val(cmd_dec)) '
        txt_fletefrontera_me.Text = Round(CDbl(txt_fletefrontera_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
    End If
    txt_cif_eu.Text = Round(CDbl(txt_fob_seg_eu) + CDbl(txt_fletefrontera_eu.Text), Val(cmd_dec))
    txt_cif_bs.Text = Round(CDbl(txt_cif_eu) * CDbl(txt_tdc_me), Val(cmd_dec))   '
    txt_cif_me.Text = Round(CDbl(txt_cif_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))

    txt_gac_eu.Text = Round(CDbl(txt_cif_eu) * 0.05, Val(cmd_dec))  '+ 1
    txt_gac_bs.Text = Round(CDbl(txt_gac_eu) * CDbl(txt_tdc_me), Val(cmd_dec))  '+ 1
    txt_gac_me.Text = Round(CDbl(txt_gac_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    
    txt_base_imp_eu.Text = Round(CDbl(txt_cif_eu) + CDbl(txt_gac_eu), Val(cmd_dec))
    txt_base_imp_bs.Text = Round(CDbl(txt_base_imp_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_base_imp_me.Text = Round(CDbl(txt_base_imp_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
End Sub

Private Sub txt_fletefrontera_me_LostFocus()
'    If txt_fletefrontera_me.Text = "" Then
'        txt_fletefrontera_bs.Text = "0"  'GlTipoCambioOficial
'    End If
'    txt_fletefrontera_bs.Text = Round(CDbl(txt_fletefrontera_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) 'GlTipoCambioOficial
'    'txt_cif_me.Text = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_seguro_me.Text) + CDbl(txt_fletefrontera_me.Text) + CDbl(txt_tacb_me.Text) + CDbl(txt_spread_me.Text), Val(cmd_dec))   '+ 1
'    'txt_cif_bs.Text = Round(CDbl(txt_cif_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))   '+ 1
'    txt_cif_me.Text = Round(CDbl(txt_fob_seg_dol) + CDbl(txt_fletefrontera_me.Text), Val(cmd_dec))   '+ 1
'    txt_cif_bs.Text = Round(CDbl(txt_cif_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))   '+ 1
'    txt_GAC_dol.Text = Round(CDbl(txt_cif_me) * CDbl(txt_gac_bs), Val(cmd_dec))  '+ 1
'    txt_base_imp_dol.Text = Round(CDbl(txt_cif_me) + CDbl(txt_GAC_dol), Val(cmd_dec))
'    txt_base_imp_bs.Text = Round(CDbl(txt_base_imp_dol) * CDbl(GlTipoCambioOficial))
    If txt_fletefrontera_me.Text = "" Then
        txt_fletefrontera_bs.Text = "0"  'GlTipoCambioOficial
        txt_fletefrontera_me.Text = "0"
        txt_fletefrontera_eu.Text = "0"
    Else
        txt_fletefrontera_bs.Text = Round(CDbl(txt_fletefrontera_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
        txt_fletefrontera_eu.Text = Round(CDbl(txt_fletefrontera_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec)) '
    End If
    txt_cif_me.Text = Round(CDbl(txt_fob_seg_dol) + CDbl(txt_fletefrontera_me.Text), Val(cmd_dec))
    txt_cif_bs.Text = Round(CDbl(txt_cif_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))   '
    txt_cif_eu.Text = Round(CDbl(txt_cif_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))

    txt_gac_me.Text = Round(CDbl(txt_cif_me) * 0.05, Val(cmd_dec))  '+ 1
    txt_gac_bs.Text = Round(CDbl(txt_gac_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))  '+ 1
    txt_gac_eu.Text = Round(CDbl(txt_gac_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    
    txt_base_imp_me.Text = Round(CDbl(txt_cif_me) + CDbl(txt_gac_me), Val(cmd_dec))
    txt_base_imp_bs.Text = Round(CDbl(txt_base_imp_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_base_imp_eu.Text = Round(CDbl(txt_base_imp_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
End Sub

Private Sub txt_fob_eu_LostFocus()
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_tacb_eu = "" Then
        txt_tacb_eu = "0"
    End If
    If txt_spread_eu = "" Then
        txt_tacb_eu = "0"
    End If
    If txt_fob_eu = "" Then
        txt_fob_bs.Text = "0"
        txt_fob_me.Text = "0"
        txt_fob_eu.Text = "0"
    Else
        txt_fob_bs.Text = Round(CDbl(txt_fob_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
        txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    End If
    txt_dcto_eu.Text = Round(CDbl(txt_fob_eu) * 0.1, Val(cmd_dec))
    txt_dcto_bs.Text = Round(CDbl(txt_dcto_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_dcto_me.Text = Round(CDbl(txt_dcto_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    
    txt_seguro_eu.Text = Round((CDbl(txt_fob_eu) - CDbl(txt_dcto_eu.Text) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_seguro_me.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    
    txt_fob_seg_eu = Round(CDbl(txt_seguro_eu) + CDbl(txt_fob_eu) - CDbl(txt_dcto_eu) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_fob_seg_dol = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'    'wwwwwwwwwwwwwwwwwwww JQA-2015
'    If txt_tacb_bs.Text = "" Then
'            txt_tacb_bs.Text = "0.02"
'            txt_spread_bs.Text = "0.02"
'    End If
'    txt_tacb_me = Round(CDbl(txt_fob_me) * CDbl(txt_tacb_bs), Val(cmd_dec))
'    txt_spread_me = Round(CDbl(txt_fob_me) * CDbl(txt_spread_bs), Val(cmd_dec))
End Sub

Private Sub txt_fob_me_LostFocus()
'    If txt_fob_bs = "" Then
'        txt_fob_bs.Text = "0"
'        txt_fob_me.Text = "0"
'    Else
'        txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
'        txt_dcto_me.Text = Round(CDbl(txt_fob_me) * 0.1, Val(cmd_dec))
'        txt_dcto_bs.Text = Round(CDbl(txt_dcto_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'        txt_seguro_bs.Text = Round((CDbl(txt_fob_bs) - CDbl(txt_dcto_bs.Text)) * 0.0078, Val(cmd_dec)) '
'        txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text)) * 0.0078, Val(cmd_dec)) '
'    End If
'    'wwwwwwwwwwwwwwwwwwww JQA-2015
'    If txt_tacb_bs.Text = "" Then
'            txt_tacb_bs.Text = "0.02"
'            txt_spread_bs.Text = "0.02"
'    End If
'    txt_tacb_me = Round(CDbl(txt_fob_me) * CDbl(txt_tacb_bs), Val(cmd_dec))
'    txt_spread_me = Round(CDbl(txt_fob_me) * CDbl(txt_spread_bs), Val(cmd_dec))
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_tacb_eu = "" Then
        txt_tacb_eu = "0"
    End If
    If txt_spread_eu = "" Then
        txt_tacb_eu = "0"
    End If
    If txt_fob_me = "" Then
        txt_fob_bs.Text = "0"
        txt_fob_me.Text = "0"
        txt_fob_eu.Text = "0"
    Else
        txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_fob_eu.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    End If
    txt_dcto_me.Text = Round(CDbl(txt_fob_me) * 0.1, Val(cmd_dec))
    txt_dcto_bs.Text = Round(CDbl(txt_dcto_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_dcto_eu.Text = Round(CDbl(txt_dcto_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    
    txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_tacb_me) + CDbl(txt_spread_me)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_seguro_eu.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
        
    txt_fob_seg_dol = Round(CDbl(txt_seguro_me) + CDbl(txt_fob_me) - CDbl(txt_dcto_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_fob_seg_eu = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
End Sub

Private Sub txt_gac_eu_LostFocus()
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
     
    If txt_gac_eu.Text = "" Then
        txt_gac_bs.Text = "0"
        txt_gac_me.Text = "0"
        txt_gac_eu.Text = "0"
    Else
        txt_gac_bs.Text = Round(CDbl(txt_gac_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
        txt_gac_me.Text = Round(CDbl(txt_gac_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
    End If
    txt_base_imp_eu.Text = Round(CDbl(txt_cif_eu) + CDbl(txt_gac_eu), Val(cmd_dec))
    txt_base_imp_bs.Text = Round(CDbl(txt_base_imp_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_base_imp_me.Text = Round(CDbl(txt_base_imp_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
End Sub

Private Sub txt_gac_me_Change()
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
     
    If txt_gac_me.Text = "" Then
        txt_gac_bs.Text = "0"
        txt_gac_me.Text = "0"
        txt_gac_eu.Text = "0"
    Else
        txt_gac_bs.Text = Round(CDbl(txt_gac_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_gac_eu.Text = Round(CDbl(txt_gac_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec)) '
    End If
    txt_base_imp_me.Text = Round(CDbl(txt_cif_me) + CDbl(txt_gac_me), Val(cmd_dec))
    txt_base_imp_bs.Text = Round(CDbl(txt_base_imp_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_base_imp_eu.Text = Round(CDbl(txt_base_imp_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
End Sub



'Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub txt_montobase_LostFocus()
'    If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
'        txt_tdc.Text = GlTipoCambioOficial
'    End If
'    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
'        txt_tdc_me = GlTipoCambioEuro
'    End If
'    If txt_montobase.Text = "" Then
'        txt_montobase.Text = "0"
'    Else
'        txt_fob_eu.Text = Round(CDbl(txt_montobase) * CDbl(txt_tdc), Val(cmd_dec))
'        txt_fob_bs.Text = Round(CDbl(txt_fob_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
'        If CDbl(txt_fob_bs.Text) > 0 Then
'            txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial))
'        Else
'            txt_fob_me.Text = "0"
'        End If
'    End If
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Select Case cmd_moneda
            Case "BRL"
                Txt_tdc.Text = GlTipoCambioBrl
            Case "BOB"
                Txt_tdc.Text = GlTipoCambioOficial      'GlTipoCambioBrl
            Case "USD"
                Txt_tdc.Text = GlTipoCambioOficial      'GlTipoCambioMercado
            Case "UFV"
                Txt_tdc.Text = GlTipoCambioUfv
            Case "RMB", "CNY"
                Txt_tdc.Text = GlTipoCambioRmb
            Case "EUR"
                Txt_tdc.Text = GlTipoCambioEuro
            Case Else
                Txt_tdc.Text = GlTipoCambioOficial
        End Select
     End If
    If txt_montobase.Text = "" Then
        txt_montobase.Text = "0"
        'txt_montobase.Text = "0"
    Else
        Select Case cmd_moneda
            Case "BRL", "UFV", "RMB", "CNY"
                txt_fob_bs.Text = Round(CDbl(txt_montobase) * CDbl(Txt_tdc), Val(cmd_dec))
                txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
                txt_fob_eu.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
            Case "BOB"
                txt_fob_bs.Text = Round(CDbl(txt_montobase), Val(cmd_dec))
                txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
                txt_fob_eu.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
            Case "USD"
                txt_fob_me.Text = Round(CDbl(txt_montobase), Val(cmd_dec))
                txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(Txt_tdc), Val(cmd_dec))
                txt_fob_eu.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
            Case "EUR"
                txt_fob_eu.Text = Round(CDbl(txt_montobase), Val(cmd_dec))
                txt_fob_bs.Text = Round(CDbl(txt_fob_eu) * CDbl(GlTipoCambioEuro), Val(cmd_dec))
                txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
            Case Else
                txt_fob_me.Text = Round(CDbl(txt_montobase), Val(cmd_dec))
                txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(Txt_tdc), Val(cmd_dec))
                txt_fob_eu.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
        End Select
    End If
    If cmd_moneda = "EUR" Then
        txt_fob_me.Locked = True
        txt_dcto_me.Locked = True
        txt_tacb_me.Locked = True
        txt_spread_me.Locked = True
        txt_seguro_me.Locked = True
        txt_fob_seg_dol.Locked = True
        txt_fletefrontera_me.Locked = True
        txt_cif_me.Locked = True
        txt_gac_me.Locked = True
        txt_base_imp_me.Locked = True
        
        txt_fob_eu.Locked = False
        txt_dcto_eu.Locked = False
        txt_tacb_eu.Locked = False
        txt_spread_eu.Locked = False
        txt_seguro_eu.Locked = False
        txt_fob_seg_eu.Locked = False
        txt_fletefrontera_eu.Locked = False
        txt_cif_eu.Locked = False
        txt_gac_eu.Locked = False
        txt_base_imp_eu.Locked = False
    Else
        txt_fob_me.Locked = False
        txt_dcto_me.Locked = False
        txt_tacb_me.Locked = False
        txt_spread_me.Locked = False
        txt_seguro_me.Locked = False
        txt_fob_seg_dol.Locked = False
        txt_fletefrontera_me.Locked = False
        txt_cif_me.Locked = False
        txt_gac_me.Locked = False
        txt_base_imp_me.Locked = False
                
        txt_fob_eu.Locked = True
        txt_dcto_eu.Locked = True
        txt_tacb_eu.Locked = True
        txt_spread_eu.Locked = True
        txt_seguro_eu.Locked = True
        txt_fob_seg_eu.Locked = True
        txt_fletefrontera_eu.Locked = True
        txt_cif_eu.Locked = True
        txt_gac_eu.Locked = True
        txt_base_imp_eu.Locked = True
    End If
    
End Sub

Private Sub txt_seguro_eu_LostFocus()
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
     End If
     If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_seguro_eu = "" Then
        txt_seguro_bs.Text = "0"
        txt_seguro_me.Text = "0"
        txt_seguro_eu.Text = "0"
    Else
        txt_seguro_bs = Round(CDbl(txt_seguro_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
        txt_seguro_me = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    End If
    txt_fob_seg_eu = Round(CDbl(txt_seguro_eu) + CDbl(txt_fob_eu) - CDbl(txt_dcto_eu) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_fob_seg_dol = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
End Sub

Private Sub txt_seguro_me_LostFocus()
'    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
'        Txt_tdc.Text = GlTipoCambioOficial
'     End If
'     If txt_seguro_me = "" Then
'        txt_seguro_bs.Text = "0"
'     Else
'        txt_seguro_bs = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'        txt_fob_seg_dol = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me) + CDbl(txt_seguro_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
'        txt_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'     End If
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
     End If
     If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_seguro_me = "" Then
        txt_seguro_bs.Text = "0"
        txt_seguro_me.Text = "0"
        txt_seguro_eu.Text = "0"
    Else
        txt_seguro_bs = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_seguro_eu = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    End If
    txt_fob_seg_dol = Round(CDbl(txt_seguro_me) + CDbl(txt_fob_me) - CDbl(txt_dcto_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_fob_seg_eu = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))

End Sub

Private Sub txt_spread_eu_LostFocus()
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_spread_eu = "" Then
        txt_spread_bs.Text = "0"
        txt_spread_me.Text = "0"
        txt_spread_eu.Text = "0"
    Else
        txt_spread_bs = Round(CDbl(txt_spread_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
        txt_spread_me = Round(CDbl(txt_spread_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    End If
    txt_seguro_eu.Text = Round((CDbl(txt_fob_eu) - CDbl(txt_dcto_eu.Text) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_seguro_me.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    
    txt_fob_seg_eu = Round(CDbl(txt_seguro_eu) + CDbl(txt_fob_eu) - CDbl(txt_dcto_eu) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_fob_seg_dol = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
End Sub

Private Sub txt_spread_me_LostFocus()
'    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
'        Txt_tdc.Text = GlTipoCambioOficial
'     End If
'     If txt_spread_me = "" Then
'        txt_spread_me.Text = "0"
'     Else
'        'txt_spread_me = Round(CDbl(txt_fob_me) * CDbl(txt_spread_bs), Val(cmd_dec))
'        txt_spread_bs = Round(CDbl(txt_spread_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'        txt_fob_seg_dol = Round(CDbl(txt_seguro_bs) + CDbl(txt_fob_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
'        txt_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'     End If
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_spread_me = "" Then
        txt_spread_bs.Text = "0"
        txt_spread_me.Text = "0"
        txt_spread_eu.Text = "0"
    Else
        txt_spread_bs = Round(CDbl(txt_spread_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_spread_eu = Round(CDbl(txt_spread_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    End If
    txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_tacb_me) + CDbl(txt_spread_me)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_seguro_eu.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    
    txt_fob_seg_dol = Round(CDbl(txt_seguro_me) + CDbl(txt_fob_me) - CDbl(txt_dcto_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_fob_seg_eu = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
End Sub

Private Sub txt_tacb_eu_LostFocus()
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_tacb_eu = "" Then
        txt_tacb_bs.Text = "0"
        txt_tacb_me.Text = "0"
        txt_tacb_eu.Text = "0"
    Else
        txt_tacb_bs = Round(CDbl(txt_tacb_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
        txt_tacb_me = Round(CDbl(txt_tacb_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    End If
    txt_seguro_eu.Text = Round((CDbl(txt_fob_eu) - CDbl(txt_dcto_eu.Text) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_seguro_me.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
    
    txt_fob_seg_eu = Round(CDbl(txt_seguro_eu) + CDbl(txt_fob_eu) - CDbl(txt_dcto_eu) + CDbl(txt_tacb_eu) + CDbl(txt_spread_eu), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_eu) * CDbl(txt_tdc_me), Val(cmd_dec))
    txt_fob_seg_dol = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
End Sub

Private Sub txt_tacb_me_LostFocus()
'    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
'        Txt_tdc.Text = GlTipoCambioOficial
'     End If
'     If txt_tacb_me = "" Then
'        txt_tacb_me.Text = "0"
'     Else
'        'txt_tacb_me = Round(CDbl(txt_fob_me) * CDbl(txt_tacb_bs), Val(cmd_dec))
'        txt_tacb_bs = Round(CDbl(txt_tacb_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'        txt_fob_seg_dol = Round(CDbl(txt_seguro_bs) + CDbl(txt_fob_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
'        txt_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'     End If
    If Txt_tdc.Text = "0" Or Txt_tdc.Text = "" Then
        Txt_tdc.Text = GlTipoCambioOficial
    End If
    If txt_tdc_me.Text = 0 Or txt_tdc_me = "" Then
        txt_tdc_me = GlTipoCambioEuro
    End If
    If txt_tacb_me = "" Then
        txt_tacb_bs.Text = "0"
        txt_tacb_me.Text = "0"
        txt_tacb_eu.Text = "0"
    Else
        txt_tacb_bs = Round(CDbl(txt_tacb_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_tacb_eu = Round(CDbl(txt_tacb_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    End If
    txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_tacb_me) + CDbl(txt_spread_me)) * 0.0078, Val(cmd_dec)) '
    txt_seguro_bs.Text = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_seguro_eu.Text = Round(CDbl(txt_seguro_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
    
    txt_fob_seg_dol = Round(CDbl(txt_seguro_me) + CDbl(txt_fob_me) - CDbl(txt_dcto_me) + CDbl(txt_tacb_me) + CDbl(txt_spread_me), Val(cmd_dec))
    txt_fob_seg_bs = Round(CDbl(txt_fob_seg_ME) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
    txt_fob_seg_eu = Round(CDbl(txt_fob_seg_bs) / CDbl(GlTipoCambioEuro), Val(cmd_dec))
End Sub


VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form aw_p_ao_solicitud_cotiza_costos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización Venta - Hoja de costos (América)"
   ClientHeight    =   7275
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11175
   ControlBox      =   0   'False
   Icon            =   "aw_p_ao_solicitud_cotiza_costos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   10875
      TabIndex        =   23
      Top             =   120
      Width           =   10935
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1680
         Picture         =   "aw_p_ao_solicitud_cotiza_costos.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   78
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "aw_p_ao_solicitud_cotiza_costos.frx":12EE
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   77
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HOJA DE COSTOS - AMERICA"
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
         Left            =   4905
         TabIndex        =   24
         Top             =   240
         Width           =   4545
      End
   End
   Begin VB.Frame Fra_datos99 
      BackColor       =   &H00C0C0C0&
      Height          =   6135
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   10935
      Begin VB.TextBox txt_conti 
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   76
         Text            =   "AMERICA"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         DataField       =   "bien_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   9480
         TabIndex        =   75
         Text            =   "0"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txt_paradas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4440
         TabIndex        =   74
         Text            =   "0"
         Top             =   5400
         Width           =   765
      End
      Begin VB.TextBox Txt_campo5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "cotiza_nro_montador"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8760
         TabIndex        =   71
         Text            =   "0"
         Top             =   5520
         Width           =   765
      End
      Begin VB.ComboBox cmd_dec 
         BackColor       =   &H0080FFFF&
         DataField       =   "cotiza_dec"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         ItemData        =   "aw_p_ao_solicitud_cotiza_costos.frx":1AC4
         Left            =   1560
         List            =   "aw_p_ao_solicitud_cotiza_costos.frx":1AD1
         TabIndex        =   0
         Text            =   "2"
         Top             =   1080
         Width           =   580
      End
      Begin VB.ComboBox cmd_moneda 
         BackColor       =   &H0080FFFF&
         DataField       =   "tipo_moneda"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         ItemData        =   "aw_p_ao_solicitud_cotiza_costos.frx":1ADE
         Left            =   3990
         List            =   "aw_p_ao_solicitud_cotiza_costos.frx":1AEE
         TabIndex        =   1
         Text            =   "BRL"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt_tdc 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "cotiza_tdc_bol"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
         Height          =   285
         Left            =   6100
         TabIndex        =   3
         Text            =   "0"
         Top             =   1080
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
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9330
         TabIndex        =   2
         Text            =   "0"
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Frame FraModeloCostoA 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"aw_p_ao_solicitud_cotiza_costos.frx":1B06
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
         Height          =   3825
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   10680
         Begin VB.TextBox txt_fob_seg_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "0"
            Top             =   1560
            Width           =   1365
         End
         Begin VB.TextBox txt_fob_seg_dol 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "0"
            Top             =   1560
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   2745
            Width           =   1365
         End
         Begin VB.TextBox txt_tac_billing_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            DataField       =   "cotiza_saldo_tac_billing_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7725
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "0"
            Top             =   2745
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   2355
            Width           =   1365
         End
         Begin VB.TextBox txt_cge_IVA_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            DataField       =   "cotiza_saldo_cge_IVA_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7725
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "0"
            Top             =   2355
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0"
            Top             =   1965
            Width           =   1365
         End
         Begin VB.TextBox txt_cge_IT_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            DataField       =   "cotiza_saldo_cge_IT_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7725
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "0"
            Top             =   1965
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   750
            Width           =   1365
         End
         Begin VB.TextBox txt_local_IVA_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            DataField       =   "cotiza_saldo_local_IVA_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7725
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "0"
            Top             =   750
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "0"
            Top             =   2715
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "0"
            Top             =   2715
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0"
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txt_local_IT_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            DataField       =   "cotiza_saldo_local_IT_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7725
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "0"
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txt_seguro_me 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3765
            TabIndex        =   6
            Text            =   "0"
            Top             =   1140
            Width           =   1365
         End
         Begin VB.TextBox txt_dcto_me 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3765
            TabIndex        =   5
            Text            =   "0"
            Top             =   735
            Width           =   1365
         End
         Begin VB.TextBox txt_fob_me 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   3765
            TabIndex        =   4
            Text            =   "0"
            Top             =   345
            Width           =   1365
         End
         Begin VB.TextBox txt_cif_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00004040&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0"
            Top             =   2325
            Width           =   1365
         End
         Begin VB.TextBox txt_fletefrontera_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0"
            Top             =   1935
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "0"
            Top             =   3225
            Width           =   1365
         End
         Begin VB.TextBox txt_cif_me 
            Alignment       =   2  'Center
            BackColor       =   &H00004040&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   3765
            TabIndex        =   40
            Text            =   "0"
            Top             =   2325
            Width           =   1365
         End
         Begin VB.TextBox txt_fletefrontera_me 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3765
            TabIndex        =   7
            Text            =   "0"
            Top             =   1935
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "0"
            Top             =   3225
            Width           =   1365
         End
         Begin VB.TextBox txt_seguro_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            DataField       =   "cotiza_precio_seg_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "0"
            Top             =   1140
            Width           =   1365
         End
         Begin VB.TextBox txt_dcto_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "0"
            Top             =   735
            Width           =   1365
         End
         Begin VB.TextBox txt_fob_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "0"
            Top             =   345
            Width           =   1365
         End
         Begin VB.TextBox txt_totalCli_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   7725
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "0"
            Top             =   1230
            Width           =   1365
         End
         Begin VB.TextBox txt_totalCli_me 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "0"
            Top             =   1230
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   7725
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "0"
            Top             =   3200
            Width           =   1365
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
            DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "0"
            Top             =   3200
            Width           =   1365
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FF0000&
            X1              =   5325
            X2              =   10775
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FF0000&
            X1              =   5325
            X2              =   10700
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   0
            X2              =   5340
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   5325
            X2              =   10750
            Y1              =   1645
            Y2              =   1645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "FOB - DSCTO + SEG:"
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
            Left            =   240
            TabIndex        =   67
            Top             =   1575
            Width           =   1845
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
            Left            =   6075
            TabIndex        =   66
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Importación - IVA:"
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
            Left            =   5505
            TabIndex        =   65
            Top             =   2370
            Width           =   2145
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Importación - IT:"
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
            Left            =   5625
            TabIndex        =   64
            Top             =   1980
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
            Left            =   6075
            TabIndex        =   63
            Top             =   770
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
            TabIndex        =   62
            Top             =   2730
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
            Left            =   6180
            TabIndex        =   61
            Top             =   375
            Width           =   1455
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FF0000&
            X1              =   5325
            X2              =   5325
            Y1              =   120
            Y2              =   3720
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento en ME:"
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
            Left            =   645
            TabIndex        =   60
            Top             =   765
            Width           =   1635
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Seguro Transporte:"
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
            Left            =   540
            TabIndex        =   59
            Top             =   1155
            Width           =   1740
         End
         Begin VB.Label lbl_campo2 
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
            TabIndex        =   58
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lbl_campo3 
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
            TabIndex        =   57
            Top             =   1950
            Width           =   1575
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio CIF:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   240
            Left            =   1035
            TabIndex        =   56
            Top             =   2325
            Width           =   1155
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
            ForeColor       =   &H00C000C0&
            Height          =   195
            Left            =   210
            TabIndex        =   55
            Top             =   3240
            Width           =   1995
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Tot. Importación Directa:"
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
            Height          =   195
            Left            =   5490
            TabIndex        =   54
            Top             =   1245
            Width           =   2130
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
            ForeColor       =   &H00404080&
            Height          =   195
            Left            =   5505
            TabIndex        =   53
            Top             =   3210
            Width           =   2085
         End
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1920
         TabIndex        =   73
         Top             =   5400
         Width           =   2085
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Montadores"
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
         Left            =   6240
         TabIndex        =   72
         Top             =   5520
         Width           =   2430
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "País"
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
         TabIndex        =   70
         Top             =   330
         Width           =   405
      End
      Begin VB.Label txt_pais 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "pais_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
         Left            =   9840
         TabIndex        =   69
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"aw_p_ao_solicitud_cotiza_costos.frx":1B99
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
         TabIndex        =   68
         Top             =   1080
         Width           =   8910
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
         Left            =   5280
         TabIndex        =   26
         Top             =   360
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
         TabIndex        =   30
         Top             =   600
         Width           =   4935
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
         Left            =   6840
         TabIndex        =   29
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
         Left            =   6840
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
         Left            =   360
         TabIndex        =   27
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
         TabIndex        =   25
         Top             =   330
         Width           =   2160
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
         Left            =   360
         TabIndex        =   22
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "edif_codigo"
         DataSource      =   "mw_solicitud_cotiza_venta.Ado_datos"
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
         TabIndex        =   13
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
         TabIndex        =   21
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
      ScaleWidth      =   11175
      TabIndex        =   14
      Top             =   7275
      Width           =   11175
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   19
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   7320
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
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   2400
      Top             =   7320
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
      Caption         =   "Ado_datos21"
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
      Left            =   4680
      Top             =   7320
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
      ConnectStringType=   3
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   6960
      Top             =   7320
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
      ConnectStringType=   3
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
   Begin MSAdodcLib.Adodc Ado_datos61 
      Height          =   330
      Left            =   9240
      Top             =   7320
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos61"
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
      Left            =   120
      Top             =   7680
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
      Left            =   2400
      Top             =   7680
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
Attribute VB_Name = "aw_p_ao_solicitud_cotiza_costos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos1 As New ADODB.Recordset
Attribute rs_datos1.VB_VarHelpID = -1
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
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
Dim VAR_SEGURO As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
    On Error GoTo AddErr
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        'mw_solicitud_cotiza_venta.mw_solicitud_cotiza_venta.Ado_datos.Recordset.CancelUpdate
        'mw_solicitud_cotiza_venta.Ado_datos.Recordset.CancelUpdate
        mw_solicitud_cotiza_venta.Ado_datos.Recordset.CancelUpdate
        Unload Me
    End If
      Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    VAR_CONTI = "AMERICA"
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & Txt_Correl.Caption & "  ", db, adOpenKeyset, adLockOptimistic
    If rs_datos10.RecordCount > 0 Then
           '- SOLO EL REGISTRO ACTIVO
           VAR_PRDA = txt_paradas.Text
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_dec = IIf(cmd_dec.Text = "", "2", cmd_dec.Text)
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!tipo_moneda = IIf(cmd_moneda.Text = "", "BOB", cmd_moneda.Text)
             If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
                txt_tdc.Text = GlTipoCambioBrl
             End If
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_tdc_bol = txt_tdc.Text
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!costo_monto = IIf(txt_montobase.Text = "", "0", Round(txt_montobase.Text, Val(cmd_dec)))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_dol = IIf(txt_fob_me = "", "0", Round(txt_fob_me, Val(cmd_dec)))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_bs = Round(CDbl(txt_fob_me) * CDbl(txt_tdc.Text), Val(cmd_dec))   'Txt_campo6.Text
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_dcto_dol = IIf(txt_dcto_me = "", "0", Round(txt_dcto_me, Val(cmd_dec)))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_dcto_bs = Round(CDbl(txt_dcto_me) * CDbl(txt_tdc.Text), Val(cmd_dec))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_dol = IIf(txt_seguro_me = "", "0", Round(txt_seguro_me, Val(cmd_dec)))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_bs = Round(CDbl(txt_seguro_me) * CDbl(txt_tdc.Text), Val(cmd_dec))
    
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_fob_seg_dol = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me) + CDbl(txt_seguro_me), Val(cmd_dec))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(txt_tdc.Text), Val(cmd_dec))
    
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_flete_dol = IIf(txt_fletefrontera_me = "", "0", Round(txt_fletefrontera_me, Val(cmd_dec)))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_flete_bs = Round(CDbl(txt_fletefrontera_me) * CDbl(txt_tdc.Text), Val(cmd_dec))
             
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_cif_dol = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_seguro_me.Text) + CDbl(txt_fletefrontera_me.Text), Val(cmd_dec))
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_cif_bs = Round(CDbl(txt_cif_me) * CDbl(txt_tdc.Text), Val(cmd_dec))  '
    '         'rs_datos!Foto = Date
    '         'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
    '         'rs_datos!archivo_foto_cargado = "N"
    '         'hora_registro
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!fecha_registro = Date     'no cambia
             mw_solicitud_cotiza_venta.Ado_datos.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
             mw_solicitud_cotiza_venta.Ado_datos.Recordset.Update    'Batch 'adAffectAll
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
                    db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "       'AND pais_continente = '" & VAR_CONTI & "'
                    'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
                    'corrprog = 0
                    Call GRABA_COSTOS
                Else
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        VAR_NAC = rs_aux6!costo_monto_usd
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        VAR_ALM = rs_aux6!costo_monto_usd
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        VAR_AGE = rs_aux6!costo_monto_usd
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1' and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        VAR_FLE = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
                    End If
                    Set rs_aux6 = New ADODB.Recordset
                    If rs_aux6.State = 1 Then rs_aux6.Close
                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux6.RecordCount > 0 Then
                        VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
                    End If
                End If
        
             End If
             If mw_solicitud_cotiza_venta.Ado_datos.Recordset!pais_continente = "AMERICA" And mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_codigo = Val(Txt_Correl.Caption) Then
                    Set rs_aux4 = New ADODB.Recordset
                    If rs_aux4.State = 1 Then rs_aux4.Close
                    rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "'  AND cotiza_codigo = " & CDbl(Txt_Correl) & "   ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux4.RecordCount > 0 Then
                        SUBTOTD = Round(rs_aux4!totdl + mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_cif_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_flete_dol, Val(cmd_dec))
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(SUBTOTD, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(SUBTOTD * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    End If
                    'Importaion Cliente
                    VAR_LOCAL = Round(rs_aux4!totdl - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE, Val(cmd_dec))
                    'VAR_LOCAL = Round(rs_aux4!totdl, Val(cmd_dec))
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(VAR_LOCAL, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(VAR_LOCAL * CDbl(txt_tdc.Text), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bs.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IT_bs.Text), Val(cmd_dec)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    txt_local_IT_dol.Text = Round(VAR_LOCAL * CDbl(txt_local_IT_bs.Text), Val(cmd_dec))
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IVA_bs.Text), Val(cmd_dec)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
                    txt_local_IVA_dol = Round(VAR_LOCAL * CDbl(txt_local_IVA_bs.Text), Val(cmd_dec))
                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = ROUND(cotiza_precio_total_dol + cotiza_saldo_local_IT_dol + cotiza_saldo_local_IVA_dol, Val(cmd_dec)) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = ROUND(cotiza_precio_total_dol_cli * " & GlTipoCambioOficial & ", Val(cmd_dec)) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
                    VAR_DOLCLI2 = Round(SUBTOTD + CDbl(txt_local_IT_dol) + CDbl(txt_local_IVA_dol), Val(cmd_dec))
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(VAR_DOLCLI2, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(VAR_DOLCLI2 * CDbl(txt_tdc.Text), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
                    
                    VAR_DOLCLI = Round(rs_aux4!totdl + mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_cif_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_dol, Val(cmd_dec))
                    'VAR_BSCLI = rs_aux4!totbs + mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_cif_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_bs
                    VAR_BSCLI = Round(VAR_DOLCLI * GlTipoCambioOficial, Val(cmd_dec))
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "

                    'VAR_SUBD = IIf(IsNull(mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_dol), SUBTOTD, mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_dol) - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_dol
                    VAR_SUBD = Round(SUBTOTD - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_dol, Val(cmd_dec))
                    VAR_SUBB = Round(VAR_SUBD * GlTipoCambioOficial, Val(cmd_dec))
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(VAR_SUBD * CDbl(txt_cge_IT_bs), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    txt_cge_IT_dol = Round(VAR_SUBD * CDbl(txt_cge_IT_bs), Val(cmd_dec))
                    'IMPORTACION CGE
                    txt_cge_IVA_dol = Round((VAR_SUBD * CDbl(txt_cge_IVA_bs)) - ((mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_cif_dol * 0.1498)) - ((CDbl(VAR_AGE) * 0.13)), Val(cmd_dec))
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(txt_cge_IVA_dol, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_cif_dol * 0.1498) * " & CDbl(dtc_desc15) & ")-((" & CDbl(VAR_AGE) & " * 0.13)* " & CDbl(dtc_desc15) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(txt_tac_billing_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round((VAR_SUBD * CDbl(txt_tac_billing_bs)), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    txt_tac_billing_dol = Round((VAR_SUBD * CDbl(txt_tac_billing_bs)), Val(cmd_dec))
                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_SUBD + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IT_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IVA_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_tac_billing_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(ao_solicitud_cotiza_venta.cotiza_precio_total_dol_cge * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    VAR_DOLCGE = Round(VAR_SUBD + CDbl(txt_cge_IT_dol) + CDbl(txt_cge_IVA_dol) + CDbl(txt_tac_billing_dol), Val(cmd_dec))
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_DOLCGE, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(VAR_DOLCGE * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
             End If
'       End If
    End If

     Call ABRIR_TABLA
        VAR_SW = ""
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
            'VAR_VAL,
'        VAR_NO2 = VAR_NO2 + Val(dtc_desc11.Text) - 1
'        VAR_NO3 = "36NO-" + Trim(Str(VAR_NO2))
'        If rs_datos!h_nro_total_equipos > 1 Then
'            'If Right(VAR_NO3, 1) = 0 Then
'                rs_datos!unidad_codigo_ant = VAR_NO1 + "-" + Right(VAR_NO3, 2)
'            'Else
'            '    rs_datos!unidad_codigo_ant = VAR_NO1 + "/" + Right(VAR_NO3, 1)
'            'End If
'        Else
'            rs_datos!unidad_codigo_ant = Txt_campo11    'VAR_NO1
'        End If
        'WC2015
'        db.Execute "Update ao_solicitud Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & mw_solicitud_cotiza_venta.Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & mw_solicitud_cotiza_venta.Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & mw_solicitud_cotiza_venta.Ado_datos.Recordset!edif_codigo & "'  "
'        db.Execute "Update ao_solicitud_calculo_trafico Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & mw_solicitud_cotiza_venta.Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & mw_solicitud_cotiza_venta.Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & mw_solicitud_cotiza_venta.Ado_datos.Recordset!edif_codigo & "'  "
'        db.Execute "Update ao_negociacion_cabecera Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & mw_solicitud_cotiza_venta.Ado_datos.Recordset!unidad_codigo & "' and negocia_codigo = " & mw_solicitud_cotiza_venta.Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & mw_solicitud_cotiza_venta.Ado_datos.Recordset!edif_codigo & "'  "
'     Call GRABA_COSTOS
       
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
  End If
  Unload Me
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub GRABA_COSTOS()
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    VAR_CONTI = "AMERICA"
    If VAR_CONTI = "AMERICA" Then
        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipo= 'B' ", db, adOpenStatic
    End If
    If VAR_CONTI = "ASIA" Then
        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoA= 'B' ", db, adOpenStatic
    End If
    If VAR_CONTI = "EUROPA" Then
        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoE= 'B' ", db, adOpenStatic
    End If
    Set Ado_datos3.Recordset = rs_datos6
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
            rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and cotiza_codigo = " & CDbl(Txt_Correl) & " ", db, adOpenKeyset, adLockOptimistic      'AND cotiza_codigo = " & mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_codigo & "
            'If rs_aux5.RecordCount = 0 Then
                rs_aux5.AddNew
                rs_aux5!ges_gestion = Year(Date)
                rs_aux5!unidad_codigo = parametro           'Txt_campo1.Caption
                rs_aux5!solicitud_codigo = GlSolicitud      'mw_solicitud_cotiza_venta.Ado_datos.Recordset!solicitud_codigo
                rs_aux5!edif_codigo = GlEdificio            'mw_solicitud_cotiza_venta.Ado_datos.Recordset!edif_codigo
                rs_aux5!cotiza_codigo = Txt_Correl         'mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_codigo

                rs_aux5!pais_continente = VAR_CONTI
                rs_aux5!estado_codigo = "REG"
                rs_aux5!codigo_costo = Ado_datos3.Recordset!codigo_costo
                rs_aux5!costo_porcentaje = Ado_datos3.Recordset!costo_porcentaje
                If Ado_datos3.Recordset!costo_porcentaje > 0 Then
                    If VAR_CONTI = "AMERICA" Then
                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
                            rs_aux5!costo_monto_usd = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto_usd * CDbl(GlTipoCambioOficial)), CDbl(cmd_dec))
                        Else
                            rs_aux5!costo_monto_usd = Round(CDbl(mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto_usd * CDbl(GlTipoCambioOficial)), CDbl(cmd_dec))
                        End If
                    End If
                    If VAR_CONTI = "ASIA" Then
'                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
'                            If IsNull(Ado_datosA.Recordset!cotiza_precio_spread_bs) Then
'                                rs_aux5!costo_monto = Round(CDbl((Ado_datos1A.Recordset!cotiza_precio_fob_bs + Ado_datos1A.Recordset!cotiza_precio_spread_bs) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl((Ado_datos1A.Recordset!cotiza_precio_fob_dol + Ado_datos1A.Recordset!cotiza_precio_spread_dol) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            Else
'                                rs_aux5!costo_monto = Round(CDbl((Ado_datosA.Recordset!cotiza_precio_fob_bs + Ado_datosA.Recordset!cotiza_precio_spread_bs) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl((Ado_datosA.Recordset!cotiza_precio_fob_dol + Ado_datosA.Recordset!cotiza_precio_spread_dol) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            End If
'                        Else
'                            'rs_aux5!costo_monto = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_cif_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            'rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            If IsNull(Ado_datosA.Recordset!cotiza_precio_base_bs) Then
'                                rs_aux5!costo_monto = Round(CDbl(Ado_datos1A.Recordset!cotiza_precio_base_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl(Ado_datos1A.Recordset!cotiza_precio_base_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            Else
'                                rs_aux5!costo_monto = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_base_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_base_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            End If
'                        End If
                    End If
                    If VAR_CONTI = "EUROPA" Then
'                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
'                            rs_aux5!costo_monto = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_fob_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_fob_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                        Else
'                            rs_aux5!costo_monto = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_cif_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                        End If
                    End If
                    rs_aux5!costo_monto2 = 0    'Round(CDbl(IIf(txt_total_bs1.Text = "", "0", txt_total_bs1.Text)), 2)
                    rs_aux5!costo_monto_usd2 = 0    'Round(CDbl(txt_total_me1.Text), 2)
                    rs_aux5!costo_monto3 = 0    'Round(CDbl(IIf(txt_dcto_bs1.Text = "", "0", txt_dcto_bs1.Text)), 2)
                    rs_aux5!costo_monto_usd3 = 0    'Round(CDbl(txt_dcto_me1.Text), 2)
                Else
                    'abrir tabla costos_paradas
                    Set rs_datos9 = New ADODB.Recordset
                    If rs_datos9.State = 1 Then rs_datos9.Close
                    rs_datos9.Open "SELECT * FROM ac_costos_paradas where trafico_num_paradas = " & Val(txt_paradas.Text) & " ", db, adOpenStatic
                    Set Ado_datos9.Recordset = rs_datos9
                    If Ado_datos9.Recordset.RecordCount > 0 Then
                        If Ado_datos3.Recordset!codigo_costo = 9 Then
                            If VAR_CONTI = "AMERICA" Then
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_instal_pintura), CDbl(cmd_dec))
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_instal_pintura * GlTipoCambioOficial), CDbl(cmd_dec))
                            End If
                            If VAR_CONTI = "ASIA" Then
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_instal_pintura), CDbl(cmd_decA))
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_instal_pintura * GlTipoCambioOficial), CDbl(cmd_decA))
                            End If
                        End If
                        If Ado_datos3.Recordset!codigo_costo = 11 Then
                            If VAR_CONTI = "AMERICA" Then
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs) * CDbl(Txt_campo5.Text), CDbl(cmd_dec))
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd) * CDbl(Txt_campo5.Text), CDbl(cmd_dec))
                            End If
                            If VAR_CONTI = "ASIA" Then
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs) * CDbl(Txt_campo5A.Text), CDbl(cmd_decA))
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd) * CDbl(Txt_campo5A.Text), CDbl(cmd_decA))
                            End If
                            If VAR_CONTI = "EUROPA" Then
'                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs), 2) * CDbl(Txt_campo5E.Text)
'                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd), 2) * CDbl(Txt_campo5E.Text)
                            End If
                        End If
                        If Ado_datos3.Recordset!codigo_costo = 12 Then
                            If VAR_CONTI = "AMERICA" Then
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_ajuste_bs), CDbl(cmd_dec))
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_ajuste_usd), CDbl(cmd_dec))
                            End If
                            If VAR_CONTI = "ASIA" Then
                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_ajuste_bs), CDbl(cmd_decA))
                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_ajuste_usd), CDbl(cmd_decA))
                            End If
                        End If
                    End If
                End If
                If Ado_datos3.Recordset!codigo_costo = 3 Then   'NACIONALIZACION
                    VAR_NAC = rs_aux5!costo_monto_usd
                End If
                If Ado_datos3.Recordset!codigo_costo = 5 Then   'ALMACENAJE
                    VAR_ALM = rs_aux5!costo_monto_usd
                End If
                If Ado_datos3.Recordset!codigo_costo = 6 Then   'COMISION AGENCIA
                    VAR_AGE = rs_aux5!costo_monto_usd
                End If
                If Ado_datos3.Recordset!codigo_costo = 8 Then   'TOTAL FLETES
                    VAR_FLE = IIf(IsNull(rs_aux5!costo_monto_usd), "0", rs_aux5!costo_monto_usd)
                End If
                If VAR_CONTI = "AMERICA" Then
                    'VAR_DOLCLI = mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_bs
                End If
                If VAR_CONTI = "ASIA" Then
                    'VAR_DOLCLI = mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_bs
                End If
                If VAR_CONTI = "EUROPA" Then
                    'VAR_DOLCLI = mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_dol - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_total_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_fob_bs - mw_solicitud_cotiza_venta.Ado_datos.Recordset!cotiza_precio_seg_bs
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
'  If rs_aux1.State = 1 Then rs_aux1.Close
  
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    'rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND cotiza_codigo = " & rs_datos!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic
    rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2, db, adOpenKeyset, adLockOptimistic
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
'WWWWWWWWWWWWWWWWWWWWWW
  If (cmd_dec = "") Then
    MsgBox "Debe registrar el Número de Decimales (#Dec) ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (cmd_moneda.Text = "") Then
    MsgBox "Debe registrar la Moneda Origen ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_tdc = "") Then
    MsgBox "Debe registrar el Tipo de Cambio (TDC) ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_montobase = "") Then
    MsgBox "Debe registrar el Monto Base en la Moneda Origen ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If (txt_fob_me = "") Or (txt_fob_me = "0") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_seguro_me = "") Or (txt_seguro_me = "0") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_fletefrontera_me = "") Or (txt_fletefrontera_me = "0") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'WWWWWWWWWWWWWWWW
End Sub

Private Sub dtc_desc1_LostFocus()
    Txt_campo3.Text = dtc_aux1.Text
    
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl, sum(costo_monto2) as totbs2, sum(costo_monto_usd2) as totdl2, sum(costo_monto3) as totbs3, sum(costo_monto_usd3) as totdl3 from ao_solicitud_costos where ges_gestion = '" & Year(Date) & "' and unidad_codigo = '" & Txt_campo1 & "' and solicitud_codigo = '" & txt_codigo & "' and edif_codigo = '" & Txt_campo2 & "' and cotiza_codigo = " & Txt_Correl & "  ", db, adOpenKeyset, adLockOptimistic
    
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
    If cmd_moneda.Text = "BRL" Then
        txt_tdc.Text = GlTipoCambioBrl
    Else
        txt_tdc.Text = GlTipoCambioOficial
    End If
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    VAR_SEGURO = "0.0078"
    If txt_tdc.Text = "" Or txt_tdc.Text = "0" Then
       txt_tdc.Text = GlTipoCambioBrl
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
       txt_cge_IVA_bs = "0.16"  '"0.151"
    End If
    If txt_tac_billing_bs.Text = "" Or txt_tac_billing_bs.Text = "0" Then
       txt_tac_billing_bs = "0.035"
    End If
End Sub

Private Sub Form_Load()
'    Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
'    If txt_local_IT_bs.Text = "" Or txt_local_IT_bs.Text = "0" Then
'       txt_local_IT_bs.Text = "0.0309"
'    End If
'    If txt_local_IVA_bs.Text = "" Or txt_local_IVA_bs.Text = "0" Then
'       txt_local_IVA_bs.Text = "0.1491"
'    End If
'    If txt_cge_IT_bs.Text = "" Or txt_cge_IT_bs.Text = "0" Then
'       txt_cge_IT_bs = "0.0416"
'    End If
'    If txt_cge_IVA_bs.Text = "" Or txt_cge_IVA_bs.Text = "0" Then
'       txt_cge_IVA_bs = "0.151"
'    End If
'    If txt_tac_billing_bs.Text = "" Or txt_tac_billing_bs.Text = "0" Then
'       txt_tac_billing_bs = "0.035"
'    End If
'    If txt_gac_bs = "" Then
'       txt_gac_bs = "0.05"
'    End If
'                 If Ado_datosA.Recordset!pais_continente = "ASIA" And Val(txt_codigo1.Caption) = 1 Then
'                        'txt_local_IT_bsA.Text = "0.0309"
'                        'txt_local_IVA_bsA.Text = "0.1491"
'                        'txt_cge_IT_bsA = "0.0416"
'                        'txt_cge_IVA_bsA = "0.151"
'                        'txt_tac_billing_bsA = "0.035"
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from ac_costos_comercializacion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
    'Bien (Equipo)
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "Select * from ac_bienes where edif_codigo = '" & GlEdificio & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
'    dtc_desc21.BoundText = dtc_codigo21.BoundText
    'gc_pais
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    If mw_solicitud_cotiza_venta.SSTab1.Tab = 0 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'AMERICA' order by pais_descripcion", db, adOpenStatic
    End If
    If mw_solicitud_cotiza_venta.SSTab1.Tab = 1 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'ASIA' order by pais_descripcion", db, adOpenStatic
    End If
    If mw_solicitud_cotiza_venta.SSTab1.Tab = 2 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'EUROPA' order by pais_descripcion", db, adOpenStatic
    End If
    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'Tipo de Equipo
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from ac_bienes_equipo_tipos ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'Cuarto de Control
    Set rs_datos61 = New ADODB.Recordset
    If rs_datos61.State = 1 Then rs_datos61.Close
    rs_datos61.Open "Select * from ac_bienes_equipo_cuadro_ctrl ", db, adOpenStatic
    Set Ado_datos61.Recordset = rs_datos61
'    dtc_desc61.BoundText = dtc_codigo61.BoundText
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

Private Sub txt_dcto_me_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
End Sub

Private Sub txt_dcto_me_LostFocus()
    If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
        txt_tdc.Text = GlTipoCambioBrl
     End If
     If txt_dcto_me = "" Then
        txt_dcto_bs.Text = "0"
     Else
        txt_dcto_bs.Text = CDbl(txt_dcto_me) * CDbl(GlTipoCambioOficial)
        txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text)) * VAR_SEGURO, Val(cmd_dec)) '
        txt_seguro_bs.Text = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))  '
        txt_fob_seg_dol.Text = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_seguro_me), Val(cmd_dec))
        txt_fob_seg_bs.Text = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
     End If
End Sub

Private Sub txt_fletefrontera_me_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
End Sub

Private Sub txt_fletefrontera_me_LostFocus()
    If txt_fletefrontera_me.Text = "0" Or txt_fletefrontera_me.Text = "" Then
        txt_fletefrontera_bs.Text = "0"  'GlTipoCambioOficial
    Else
        txt_fletefrontera_bs.Text = Round(CDbl(txt_fletefrontera_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) 'GlTipoCambioOficial
        txt_cif_me.Text = Round(CDbl(txt_fob_seg_dol.Text) + CDbl(txt_fletefrontera_me.Text), Val(cmd_dec)) '+ 1
        'txt_cif_me.Text = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_seguro_me.Text) + CDbl(txt_fletefrontera_me.Text), Val(cmd_dec)) '+ 1
        txt_cif_bs.Text = Round(CDbl(txt_cif_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
    End If
End Sub

Private Sub txt_fob_me_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
End Sub

Private Sub txt_fob_me_LostFocus()
    If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
        txt_tdc.Text = GlTipoCambioBrl
    End If
    If txt_fob_me = "" Then
        txt_fob_bs.Text = "0"
        txt_fob_me.Text = "0"
    Else
        txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_dcto_me.Text = Round(CDbl(txt_fob_me) * 0.1, Val(cmd_dec))
        txt_dcto_bs.Text = Round(CDbl(txt_dcto_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        'txt_dcto_bs.Text = Round(CDbl(txt_fob_bs) * 0.1, Val(cmd_dec))
        txt_seguro_me.Text = Round((CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text)) * VAR_SEGURO, Val(cmd_dec)) '
        txt_seguro_bs.Text = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))  '
        txt_fob_seg_dol.Text = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_seguro_me), Val(cmd_dec))
        txt_fob_seg_bs.Text = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
    End If
End Sub

Private Sub txt_montobase_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
End Sub

Private Sub txt_montobase_LostFocus()
    If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
        Select Case cmd_moneda
            Case "BRL"
                txt_tdc.Text = GlTipoCambioBrl
            Case "BOB"
                txt_tdc.Text = GlTipoCambioOficial      'GlTipoCambioBrl
            Case "USD"
                txt_tdc.Text = GlTipoCambioOficial      'GlTipoCambioMercado
            Case "UFV"
                txt_tdc.Text = GlTipoCambioUfv
            Case "RMB", "CNY"
                txt_tdc.Text = GlTipoCambioRmb
            Case "EUR"
                txt_tdc.Text = GlTipoCambioEuro
            Case Else
                txt_tdc.Text = GlTipoCambioOficial
        End Select
     End If
    If txt_montobase.Text = "" Then
        txt_montobase.Text = "0"
    Else
        Select Case cmd_moneda
            Case "BRL", "UFV", "RMB", "CNY", "EUR"
                txt_fob_bs.Text = Round(CDbl(txt_montobase) * CDbl(txt_tdc), Val(cmd_dec))
                txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
            Case "BOB"
                txt_fob_bs.Text = Round(CDbl(txt_montobase), Val(cmd_dec))
                txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
            Case "USD"
                txt_fob_me.Text = Round(CDbl(txt_montobase), Val(cmd_dec))
                txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(txt_tdc), Val(cmd_dec))
            'Case UFV
            '    txt_fob_bs.Text = Round(CDbl(txt_montobase) * CDbl(txt_tdc), Val(cmd_dec))
            '    txt_fob_me.Text = Round(CDbl(txt_fob_bs) / CDbl(GlTipoCambioOficial), Val(cmd_dec))
            Case Else
                txt_fob_me.Text = Round(CDbl(txt_montobase), Val(cmd_dec))
                txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(txt_tdc), Val(cmd_dec))
        End Select
        'txt_fob_me.Text = Round(CDbl(txt_montobase) / CDbl(txt_tdc), Val(cmd_dec))
        'txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(GlTipoCambioOficial))
    End If
End Sub

Private Sub txt_seguro_me_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
End Sub

Private Sub txt_seguro_me_LostFocus()
    If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
        txt_tdc.Text = GlTipoCambioBrl
     End If
     If txt_seguro_me = "" Then
        txt_seguro_bs.Text = "0"
     Else
        txt_seguro_bs = Round(CDbl(txt_seguro_me) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
        txt_fob_seg_dol.Text = Round(CDbl(txt_fob_me) - CDbl(txt_dcto_me.Text) + CDbl(txt_seguro_me), Val(cmd_dec))
        txt_fob_seg_bs.Text = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
     End If
End Sub

'Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub txt_tdc_LostFocus()
    If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
        txt_tdc.Text = GlTipoCambioBrl
    End If
    If txt_montobase.Text = "" Or CDbl(txt_montobase.Text) < "0" Then
        txt_montobase.Text = "0"
        'txt_montobase.Text = "0"
    Else
        'if val(cmd_dec)= 0 then
        'txt_fob_me.Text = Round(CDbl(txt_montobase) / CDbl(txt_tdc), 2)
        txt_fob_me.Text = Round(CDbl(txt_montobase) / CDbl(txt_tdc), Val(cmd_dec))
        txt_fob_bs.Text = Round(CDbl(txt_fob_me) * CDbl(GlTipoCambioOficial))
    End If
End Sub

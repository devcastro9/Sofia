VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ro_Personal_Liquidacion 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administracion de Personal - Ficha Personal - Finiquitos, Quinquenios y otras Liquidaciones"
   ClientHeight    =   8730
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   9495
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ro_Personal_Liquidacion.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
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
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   9255
      Begin VB.ComboBox TxtGestion_ini 
         DataField       =   "ges_gestion_ini"
         DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
         Height          =   315
         ItemData        =   "ro_Personal_Liquidacion.frx":6C032
         Left            =   2760
         List            =   "ro_Personal_Liquidacion.frx":6C057
         TabIndex        =   80
         Text            =   "2015"
         Top             =   480
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPFechaLiq 
         DataField       =   "Fecha_Liquidacion"
         DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
         Height          =   255
         Left            =   5160
         TabIndex        =   64
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   96403457
         CurrentDate     =   42307
         MinDate         =   2
      End
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   5085
         MaxLength       =   80
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   2940
         MaxLength       =   80
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   1540
         MaxLength       =   80
         TabIndex        =   58
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "IV. TOTAL BENEFICIOS SOCIALES"
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
         Height          =   1360
         Left            =   120
         TabIndex        =   47
         Top             =   5880
         Width           =   9015
         Begin VB.CommandButton btn_total 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Calcular Total"
            Height          =   315
            Left            =   3000
            MaskColor       =   &H0000FF00&
            TabIndex        =   88
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox CmbChq_Trf 
            DataField       =   "Forma_pago"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   240
            TabIndex        =   57
            Text            =   "CHEQUE"
            Top             =   420
            Width           =   1980
         End
         Begin VB.TextBox TxtTotBenef 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Monto_Total"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
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
            Left            =   7200
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TxtNo_Chq 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Num_chq_cmpbte"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   3240
            MultiLine       =   -1  'True
            TabIndex        =   52
            Top             =   420
            Width           =   1335
         End
         Begin VB.TextBox TxtDeduccion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Deducciones"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TxtCta 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cta_codigo"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   5280
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   420
            Width           =   3375
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00000000&
            Caption         =   "Total Beneficios Sociales BOB"
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
            Index           =   27
            Left            =   4440
            TabIndex        =   55
            Top             =   1005
            Width           =   3000
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Deducciones"
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
            Index           =   26
            Left            =   240
            TabIndex        =   54
            Top             =   1005
            Width           =   1200
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Cuenta Bancaria"
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
            Index           =   25
            Left            =   5280
            TabIndex        =   53
            Top             =   165
            Width           =   1485
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Nro.Cheq./Cmpbte."
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
            Left            =   3120
            TabIndex        =   51
            Top             =   165
            Width           =   1710
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Forma de Pago"
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
            Left            =   240
            TabIndex        =   48
            Top             =   165
            Width           =   1410
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "III. TOTAL REMUNERACION PROMEDIO INDEMNIZABLE"
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
         Height          =   2175
         Left            =   120
         TabIndex        =   26
         Top             =   3600
         Width           =   9015
         Begin VB.TextBox txt_agui_d 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "estado_registro"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
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
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "AGUI D"
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txt_agui_m 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   86
            Text            =   "AGUI M"
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txt_dias_agui 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "dias_agui"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
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
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   85
            Text            =   "DIA"
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txt_meses_agui 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "meses_agui"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
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
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   84
            Text            =   "MES"
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtDesahucio 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Imdem_Año"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   3840
            MultiLine       =   -1  'True
            TabIndex        =   76
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox TxtOtros 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Otros_Pagos"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   7200
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   1725
            Width           =   1455
         End
         Begin VB.TextBox TxtPrima 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Prima_Legal"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   4680
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   1725
            Width           =   1455
         End
         Begin VB.ComboBox CmbDia 
            DataField       =   "dias"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            IntegralHeight  =   0   'False
            Left            =   6960
            TabIndex        =   40
            Text            =   "0"
            Top             =   765
            Width           =   900
         End
         Begin VB.ComboBox CmbMes 
            DataField       =   "meses"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            IntegralHeight  =   0   'False
            Left            =   4440
            TabIndex        =   39
            Text            =   "0"
            Top             =   765
            Width           =   900
         End
         Begin VB.TextBox TxtImdemDia 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Indem_dias"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   6960
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox TxtImdemMes 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Imdem_Mes"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   4440
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox TxtImdemAño 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Imdem_Año"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1140
            Width           =   1695
         End
         Begin VB.ComboBox CmbAño 
            DataField       =   "Años"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "ro_Personal_Liquidacion.frx":6C09D
            Left            =   1800
            List            =   "ro_Personal_Liquidacion.frx":6C09F
            TabIndex        =   32
            Text            =   "0"
            Top             =   765
            Width           =   900
         End
         Begin VB.TextBox TxtVacacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Aguin_Vacacion"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   1725
            Width           =   1455
         End
         Begin VB.TextBox TxtNavidad 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Aguin_Navidad"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1725
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Deshaucio 3 Meses por Retiro Forzoso:"
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
            Height          =   240
            Index           =   11
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   3540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Otros Pagos"
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
            Index           =   24
            Left            =   7200
            TabIndex        =   43
            Top             =   1485
            Width           =   1125
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Prima Legal"
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
            Index           =   23
            Left            =   4680
            TabIndex        =   42
            Top             =   1485
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Vacaciones"
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
            Index           =   22
            Left            =   2520
            TabIndex        =   41
            Top             =   1485
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   19
            Left            =   7920
            TabIndex        =   35
            Top             =   795
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Meses"
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
            Index           =   18
            Left            =   5400
            TabIndex        =   34
            Top             =   795
            Width           =   615
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Años"
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
            Index           =   17
            Left            =   2760
            TabIndex        =   33
            Top             =   795
            Width           =   465
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Aguinaldo"
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
            Left            =   240
            TabIndex        =   31
            Top             =   1485
            Width           =   915
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Imdemnización . ."
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
            TabIndex        =   30
            Top             =   1155
            Width           =   1530
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Tiempo Servicio"
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
            Index           =   16
            Left            =   240
            TabIndex        =   29
            Top             =   795
            Width           =   1485
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "II. LIQUIDACION PROMEDIO INDEMNIZABLE (3 Ultimos Meses)"
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
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   9015
         Begin VB.TextBox txt_promedio 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Left            =   7560
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox CmbMes3 
            DataField       =   "Mes_Utimo"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            ItemData        =   "ro_Personal_Liquidacion.frx":6C0A1
            Left            =   5760
            List            =   "ro_Personal_Liquidacion.frx":6C0C9
            TabIndex        =   78
            Text            =   "MARZO"
            Top             =   480
            Width           =   1785
         End
         Begin VB.ComboBox CmbMes2 
            DataField       =   "Mes_Penultimo"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            ItemData        =   "ro_Personal_Liquidacion.frx":6C132
            Left            =   3600
            List            =   "ro_Personal_Liquidacion.frx":6C15A
            TabIndex        =   77
            Text            =   "FEBRERO"
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtpago6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Utimo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   5760
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtpago5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Penul"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   3600
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtpago4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Antep"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   1485
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Txtpago3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Utimo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   5760
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   850
            Width           =   1695
         End
         Begin VB.TextBox TxtPago2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Penultimo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   3600
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   850
            Width           =   1695
         End
         Begin VB.ComboBox CmbMes1 
            DataField       =   "Mes_Antepenultimo"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            ItemData        =   "ro_Personal_Liquidacion.frx":6C1C3
            Left            =   1485
            List            =   "ro_Personal_Liquidacion.frx":6C1EB
            TabIndex        =   19
            Text            =   "ENERO"
            Top             =   500
            Width           =   1830
         End
         Begin VB.TextBox txtpago1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Antepenultimo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   285
            Left            =   1485
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   850
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Promedio"
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
            Height          =   240
            Index           =   31
            Left            =   7800
            TabIndex        =   83
            Top             =   600
            Width           =   885
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Mes Último"
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
            Height          =   240
            Index           =   28
            Left            =   5760
            TabIndex        =   67
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Mes Penúltimo"
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
            Height          =   240
            Index           =   20
            Left            =   3600
            TabIndex        =   66
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Mes Antepenúltimo"
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
            Height          =   240
            Index           =   15
            Left            =   1440
            TabIndex        =   65
            Top             =   240
            Width           =   1710
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Remuneracion"
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
            Index           =   21
            Left            =   120
            TabIndex        =   20
            Top             =   885
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Meses . . . . . . ."
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
            TabIndex        =   18
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Otros Pagos . ."
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
            Left            =   120
            TabIndex        =   17
            Top             =   1245
            Width           =   1305
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "I. DATOS GENERALES"
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
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   9015
         Begin MSDataListLib.DataCombo DTCFInicio 
            Bindings        =   "ro_Personal_Liquidacion.frx":6C254
            DataField       =   "fecha_inicio"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   240
            TabIndex        =   68
            Top             =   540
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   4210752
            ForeColor       =   16777152
            ListField       =   "fecha_inicio"
            BoundColumn     =   "fecha_ingreso"
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
         Begin MSDataListLib.DataCombo DTCFFin 
            Bindings        =   "ro_Personal_Liquidacion.frx":6C26F
            DataField       =   "fecha_fin"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   2025
            TabIndex        =   69
            Top             =   540
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   4210752
            ForeColor       =   16777152
            ListField       =   "fecha_fin"
            BoundColumn     =   "fecha_retiro"
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
         Begin MSDataListLib.DataCombo DtcRetiroDes 
            Bindings        =   "ro_Personal_Liquidacion.frx":6C28A
            DataField       =   "tipo_memo"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   3840
            TabIndex        =   8
            Top             =   540
            Width           =   4860
            _ExtentX        =   8573
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
         Begin MSDataListLib.DataCombo DtcRetiro 
            Bindings        =   "ro_Personal_Liquidacion.frx":6C2A3
            DataField       =   "tipo_memo"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   7200
            TabIndex        =   9
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSComCtl2.DTPicker DTPFInicio 
            DataField       =   "fecha_ingreso"
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Top             =   540
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   96403457
            CurrentDate     =   40471
         End
         Begin MSComCtl2.DTPicker DTPFFin 
            DataField       =   "fecha_retiro"
            Height          =   285
            Left            =   2025
            TabIndex        =   11
            Top             =   540
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   96403457
            CurrentDate     =   40471
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Fecha Retiro/Liq."
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
            Index           =   1
            Left            =   2055
            TabIndex        =   14
            Top             =   300
            Width           =   1530
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Motivo de Retiro o Liquidacion"
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
            Index           =   4
            Left            =   3840
            TabIndex        =   13
            Top             =   300
            Width           =   2715
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Fecha Ingreso"
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
            Left            =   255
            TabIndex        =   12
            Top             =   300
            Width           =   1290
         End
      End
      Begin VB.TextBox TxtAprob 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DataField       =   "estado_registro"
         DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
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
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "REG"
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
         Height          =   315
         ItemData        =   "ro_Personal_Liquidacion.frx":6C2BC
         Left            =   3960
         List            =   "ro_Personal_Liquidacion.frx":6C2E1
         TabIndex        =   0
         Text            =   "2015"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtLquida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DataField       =   "id_liquidacion"
         DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
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
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gestión Ini."
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
         Height          =   255
         Index           =   29
         Left            =   2760
         TabIndex        =   81
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Liquidación"
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
         Height          =   255
         Index           =   14
         Left            =   5160
         TabIndex        =   63
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gestión Fin."
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
         Height          =   255
         Index           =   13
         Left            =   3960
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nombre Archivo Digital"
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
         Index           =   5
         Left            =   6975
         TabIndex        =   61
         Top             =   285
         Width           =   2070
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado"
         DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
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
         Height          =   255
         Index           =   12
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblARCH 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         DataField       =   "ARCHIVO"
         DataSource      =   "Ado_Auxiliar"
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
         Left            =   6960
         TabIndex        =   5
         Top             =   555
         Width           =   2025
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Liquidación"
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
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "ro_Personal_Liquidacion.frx":6C327
      ScaleHeight     =   915
      ScaleWidth      =   9195
      TabIndex        =   70
      Top             =   240
      Width           =   9255
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "ro_Personal_Liquidacion.frx":D8359
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "ro_Personal_Liquidacion.frx":D8563
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cargar"
         Height          =   680
         Left            =   2160
         Picture         =   "ro_Personal_Liquidacion.frx":D876D
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Carga Contrato en PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver"
         Height          =   680
         Left            =   3000
         Picture         =   "ro_Personal_Liquidacion.frx":D8AF5
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Ver Contrato PDF"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIQUIDACIONES"
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
         Left            =   5235
         TabIndex        =   75
         Top             =   240
         Width           =   2445
      End
   End
   Begin MSAdodcLib.Adodc AdoBeneficiario 
      Height          =   330
      Left            =   0
      Top             =   8400
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
   Begin MSAdodcLib.Adodc AdoMotivos 
      Height          =   330
      Left            =   4320
      Top             =   8400
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
      Caption         =   "AdoMotivos"
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
   Begin MSAdodcLib.Adodc AdoContrato2 
      Height          =   330
      Left            =   6480
      Top             =   8400
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
      Caption         =   "AdoContrato2"
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
      Left            =   8640
      Top             =   8400
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
      Left            =   2160
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos_planilla 
      Height          =   330
      Left            =   0
      Top             =   8760
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
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      Index           =   30
      Left            =   0
      TabIndex        =   79
      Top             =   1200
      Width           =   45
   End
End
Attribute VB_Name = "ro_Personal_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_beneficiario As New ADODB.Recordset
Dim rs_Auxiliar As New ADODB.Recordset
Attribute rs_Auxiliar.VB_VarHelpID = -1
Dim rs_motivo As New ADODB.Recordset
Dim rs_Org As New ADODB.Recordset
Dim rs_Pry As New ADODB.Recordset
Dim rs_correlativo As New ADODB.Recordset
Dim rs_vacaciones_prog As ADODB.Recordset

Dim rs_planilla As ADODB.Recordset
Dim rs_planilla_det As ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset

Dim e As Long

Dim var_cod As Integer
Dim VAR_RETIRO As Integer
Dim VAR_DIA As Integer
Dim VAR_MES As Integer
Dim VAR_ANIO As Integer
Dim VAR_DIA2 As Integer
Dim VAR_MES2 As Integer
Dim VAR_ANIO2 As Integer
Dim VAR_DIA3 As Double
Dim VAR_MES3 As Double

Dim GESTION1 As Integer
Dim GESTION2 As Integer
Dim GESTION3 As Integer

Dim mes1 As Integer
Dim mes2 As Integer
Dim MES3 As Integer

Dim VAR_MES0 As Integer
Dim VAR_MES00 As Integer

Dim VAR_NMES1 As Integer
Dim VAR_NMES2 As Integer
Dim VAR_NMES3 As Integer

Dim DirLiq As String
Dim VAR_VAL As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean
Dim meses_vac, dias_vac As Integer
Dim total_d, total_m As Double



Public Function Calcular_tiempo_trabajado(FechaInicial As Date, FechaActual As Date, posicion As Integer) As String
' Dim Anios, Meses, Dias As Variant
   Dim ldtFecha1 As Date
        Dim ldtFecha2 As Date
        Dim ldtMesAnterior As Date
        Dim liDias As Integer
        Dim liMeses As Integer
        Dim liAños As Integer
        '
        ldtFecha1 = FechaActual
        ldtMesAnterior = DateAdd("M", -1, ldtFecha1)
        ldtFecha2 = FechaInicial
        '
        liAños = DatePart("yyyy", ldtFecha1) - DatePart("yyyy", ldtFecha2)
        liMeses = DatePart("m", ldtFecha1) - DatePart("m", ldtFecha2)
        liDias = (DatePart("d", ldtFecha1) - DatePart("d", ldtFecha2)) + 1
        Do While liDias < 0
            If liMeses = 0 Then
                liDias = liDias + DateTime.Day(DateSerial(Year(ldtFecha1), Month(ldtFecha1) + 1, 0))
            Else
                liDias = liDias + DateTime.Day(DateSerial(Year(ldtMesAnterior), Month(ldtMesAnterior) + 1, 0))
            End If
            liMeses = liMeses - 1
            ldtMesAnterior = DateAdd("M", -1, ldtMesAnterior)
        Loop
        If liMeses < 0 Then
            liMeses = liMeses + 12
            liAños = liAños - 1
        End If
 Select Case posicion
 Case 1
    CmbAño.Text = liAños
    CmbMes.Text = liMeses
    CmbDia.Text = liDias
 Case 2
 txt_meses_agui.Text = liMeses
 txt_dias_agui.Text = liDias
 TxtNavidad.Text = Val(txt_agui_m.Text) + Val(txt_agui_d.Text)
 End Select
End Function
Public Sub datod_plmilla()
Set rs_planilla = New ADODB.Recordset
If rs_planilla.State = 1 Then rs_planilla.Close
  rs_planilla.Open "select * from ro_personal_contratado WHERE beneficiario_codigo = '" & txtBenef & "'", db, adOpenKeyset, adLockOptimistic
  Set AdoBeneficiario.Recordset = rs_beneficiario.DataSource
'  DtcBenefDes.BoundText = DtcBenef.BoundText
Call Calcular_tiempo_trabajado(rs_planilla!fecha_ingreso, rs_planilla!fecha_expiracion, 1)
Call planilla_actual

TxtImdemAño.Text = Round(Val((txt_promedio)) * CmbAño.Text, 2)
TxtImdemMes.Text = Round((Val(txt_promedio) / 12) * CmbMes.Text, 2)
TxtImdemDia.Text = Round((Val(txt_promedio) / 360) * CmbDia.Text, 2)







End Sub
Private Sub planilla_actual()
Dim inicio_anio As Date
Dim dia As Integer
Set rs_planilla_det = New ADODB.Recordset
If rs_planilla_det.State = 1 Then rs_planilla_det.Close
rs_planilla_det.Open "select * from ro_pagos_cronograma_Detalle WHERE beneficiario_codigo = '" & txtBenef & "' and ges_gestion = '" & Year(DTCFFin.Text) & "' order by mes_grupo asc ", db, adOpenKeyset, adLockOptimistic
rs_planilla_det.MoveFirst
If rs_planilla_det!dias_trabajados = 30 Then
inicio_anio = "01/" & rs_planilla_det!mes_grupo & "/" & rs_planilla_det!mes_grupo
Else
dia = 30 - (rs_planilla_det!dias_trabajados - 1)
inicio_anio = "01/" & rs_planilla_det!mes_grupo & "/" & rs_planilla_det!mes_grupo
End If

Call Calcular_tiempo_trabajado(inicio_anio, DTCFFin.Text, 2)




End Sub


Private Sub calcular()
  VAR_MES0 = Month(DTPFechaLiq.Value)
    GESTION1 = Year(DTPFechaLiq.Value)
    Select Case VAR_MES0
        Case 1
            CmbMes3 = "DICIEMBRE"
            CmbMes2 = "NOVIEMBRE"
            CmbMes1 = "OCTUBRE"
            MES3 = 12
            mes2 = 11
            mes1 = 10
            GESTION3 = Year(DTPFechaLiq.Value) - 1
            GESTION2 = Year(DTPFechaLiq.Value) - 1
            GESTION1 = Year(DTPFechaLiq.Value) - 1
        Case 2
            CmbMes3 = "ENERO"
            CmbMes2 = "DICIEMBRE"
            CmbMes1 = "NOVIEMBRE"
             MES3 = 1
            mes2 = 12
            mes1 = 11
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value) - 1
            GESTION1 = Year(DTPFechaLiq.Value) - 1
        Case 3
            CmbMes3 = "FEBRERO"
            CmbMes2 = "ENERO"
            CmbMes1 = "DICIEMBRE"
             MES3 = 2
            mes2 = 1
            mes1 = 12
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value) - 1
        Case 4
            CmbMes3 = "MARZO"
            CmbMes2 = "FEBRERO"
            CmbMes1 = "ENERO"
             MES3 = 3
            mes2 = 2
            mes1 = 1
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 5
            CmbMes3 = "ABRIL"
            CmbMes2 = "MARZO"
            CmbMes1 = "FEBRERO"
            MES3 = 4
            mes2 = 3
            mes1 = 2
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 6
            CmbMes3 = "MAYO"
            CmbMes2 = "ABRIL"
            CmbMes1 = "MARZO"
             MES3 = 5
            mes2 = 4
            mes1 = 3
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 7
            CmbMes3 = "JUNIO"
            CmbMes2 = "MAYO"
            CmbMes1 = "ABRIL"
            MES3 = 6
            mes2 = 5
            mes1 = 4
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 8
            CmbMes3 = "JULIO"
            CmbMes2 = "JUNIO"
            CmbMes1 = "MAYO"
            MES3 = 7
            mes2 = 6
            mes1 = 5
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 9
            CmbMes3 = "AGOSTO"
            CmbMes2 = "JULIO"
            CmbMes1 = "JUNIO"
            MES3 = 8
            mes2 = 7
            mes1 = 6
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 10
            CmbMes3 = "SEPTIEMBRE"
            CmbMes2 = "AGOSTO"
            CmbMes1 = "JULIO"
            MES3 = 9
            mes2 = 8
            mes1 = 7
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 11
            CmbMes3 = "OCTUBRE"
            CmbMes2 = "SEPTIEMBRE"
            CmbMes1 = "AGOSTO"
            MES3 = 10
            mes2 = 9
            mes1 = 8
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
        Case 12
            CmbMes3 = "NOVIEMBRE"
            CmbMes2 = "OCTUBRE"
            CmbMes1 = "SEPTIEMBRE"
            MES3 = 11
            mes2 = 10
            mes1 = 9
            GESTION3 = Year(DTPFechaLiq.Value)
            GESTION2 = Year(DTPFechaLiq.Value)
            GESTION1 = Year(DTPFechaLiq.Value)
    End Select



  VAR_MES00 = VAR_MES0 - 1
    Set rs_aux1 = New ADODB.Recordset
    rs_aux1.Open "select * from ro_pagos_cronograma_Detalle where beneficiario_codigo = '" & txtBenef & "' and mes_grupo = " & MES3 & " AND ges_gestion = " & GESTION3 & "", db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        Txtpago3.Text = rs_aux1!total_ganado
    Else
        Txtpago3.Text = 0
    End If
    'MES PENULTIMO
    VAR_MES00 = VAR_MES0 - 2
    Set rs_aux1 = New ADODB.Recordset
    rs_aux1.Open "select * from ro_pagos_cronograma_Detalle where beneficiario_codigo = '" & txtBenef & "' and mes_grupo = " & mes2 & " AND ges_gestion = " & GESTION2 & "", db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        TxtPago2.Text = rs_aux1!total_ganado
    Else
        TxtPago2.Text = 0
    End If
    'MES ANTEPENULTIMO
    VAR_MES00 = VAR_MES0 - 3
    Set rs_aux1 = New ADODB.Recordset
    rs_aux1.Open "select * from ro_pagos_cronograma_Detalle where beneficiario_codigo = '" & txtBenef & "' and mes_grupo = " & mes1 & "  AND ges_gestion = " & GESTION1 & "", db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        txtpago1.Text = rs_aux1!total_ganado
    Else
        txtpago1.Text = 0
    End If
    
    If DtcRetiro.Text = "QUI" Then     'DtcRetiro
        TxtTotBenef.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * 5
        Frame4.Visible = False
    Else
        'If rw_ficha_rrhh.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
        Frame4.Visible = True
        If DtcRetiro.Text = "REF" Then
            VAR_RETIRO = 1
        Else
            VAR_RETIRO = 0
        End If
        txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
        VAR_DIA = Day(DTCFInicio)
        VAR_MES = Month(DTCFInicio)
        VAR_ANIO = Year(DTCFInicio)
        VAR_DIA2 = Day(DTPFechaLiq)
        VAR_MES2 = Month(DTPFechaLiq)
        VAR_ANIO2 = Year(DTPFechaLiq)
        If Val(TxtGestion.Text) > VAR_ANIO Then
            'MesesAguinaldo* ((S3UM/3)/12)+DíasAguinaldo* ((S3UM/3)/ 360)
            VAR_MES3 = VAR_MES2 * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 12
            VAR_DIA3 = VAR_DIA2 * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
            TxtNavidad.Text = VAR_MES3 + VAR_DIA3
        Else
            If VAR_DIA > VAR_DIA2 Then
                Select Case VAR_MES2
                  Case 1, 3, 5, 7, 8, 10, 12
                      VAR_NRODIAS = 31
                  Case 2
                      VAR_NRODIAS = 28
                  Case 4, 6, 9, 11
                      VAR_NRODIAS = 30
                End Select
                VAR_MES3 = (VAR_MES2 - VAR_MES - 1) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 12
                VAR_DIA3 = (VAR_DIA2 - VAR_DIA2 + VAR_NRODIAS) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
            Else
                VAR_MES3 = (VAR_MES2 - VAR_MES) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 12
                VAR_DIA3 = (VAR_DIA2 - VAR_DIA2) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
            End If
            TxtNavidad.Text = VAR_MES3 + VAR_DIA3
        End If
        'DíasVacaciónFiniquito*((S3UM/3)/ 360)
        Set rs_vacaciones_prog = New ADODB.Recordset
        If rs_vacaciones_prog.State = 1 Then rs_vacaciones_prog.Close
        rs_vacaciones_prog.Open "select sum(Dias_Pendientes) as vacas from ro_vacaciones_programadas where beneficiario_codigo = '" & txtBenef & "' ", db, adOpenKeyset, adLockOptimistic
        If rs_vacaciones_prog!vacas <> "NULL" Then
        
        If rs_vacaciones_prog!vacas <= 15 Then
        TxtVacacion.Text = Round(((Val(txt_promedio) / 2) / 360) * rs_vacaciones_prog!vacas, 2)
        End If
        
        If rs_vacaciones_prog!vacas > 15 And rs_vacaciones_prog!vacas <= 20 Then
        TxtVacacion.Text = Round(((Val(txt_promedio) / 1.5) / 360) * rs_vacaciones_prog!vacas, 2)
        End If
        
        If rs_vacaciones_prog!vacas > 20 And rs_vacaciones_prog!vacas <= 30 Then
        TxtVacacion.Text = Round(((Val(txt_promedio) / 1) / 360) * rs_vacaciones_prog!vacas, 2)
        End If
        If rs_vacaciones_prog!vacas > 30 Then
        TxtVacacion.Text = Round(((Val(txt_promedio) / 2) / 360) * rs_vacaciones_prog!vacas, 2)
'        dias_vac = rs_vacaciones_prog!vacas
'        meses_vac = 0
'         While dias_vac >= 30
'             meses_vac = meses_vac + 1
'             dias_vac = dias_vac - 30
'         Wend
'        End If
'
'        If dias_vac > 0 Then
'        total_d = Round((Val(txt_promedio) / 360) * dias_vac)
        
        End If
        TxtDeduccion = Round(Val(TxtVacacion.Text) * 0.13, 2)
        
'        If rs_vacaciones_prog!vacas > 30 Then
'
'            TxtVacacion.Text = rs_vacaciones_prog!Dias_Pendientes * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
'        Else
'            TxtVacacion.Text = "0"
'        End If
        

        End If
        'Set Ado_VacacionesProg.Recordset = rs_vacaciones_prog
        'Set DtgVacacionesProg.DataSource = Ado_VacacionesProg.Recordset
    End If
    
    
    
    
    
End Sub


Private Sub btn_total_Click()
TxtTotBenef.Text = Round(Val(txtpago4) + Val(txtpago5) + Val(txtpago6) + Val(txtDesahucio) + Val(TxtImdemAño) + Val(TxtImdemMes) + Val(TxtImdemDia) + Val(TxtNavidad) + Val(TxtVacacion) + Val(TxtPrima) + Val(TxtOtros) - Val(TxtDeduccion), 2)
End Sub

'Private Sub cmdAprueba_Click()
'  On Error GoTo UpdateErr
'   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'   If AdoLiquidacion.Recordset!estado_codigo = "NO" Then
'      If sino = vbYes Then
'         AdoLiquidacion.Recordset!estado_codigo = "SI"
'         AdoLiquidacion.Recordset!fecha_registro = Date
'         AdoLiquidacion.Recordset!usr_codigo = GlUsuario
'         AdoLiquidacion.Recordset.UpdateBatch adAffectAll
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
'        AdoLiquidacion.Recordset.CancelUpdate
'        If mvBookMark > 0 Then
'          AdoLiquidacion.Recordset.Bookmark = mvBookMark
'        Else
'          AdoLiquidacion.Recordset.MoveFirst
'        End If
'        mbDataChanged = False
''        Fra_ABM.Enabled = False
'        fraOpciones.Visible = True
'        FraGrabarCancelar.Visible = False
'        DtG_Auxiliar.Enabled = True
        Unload Me
    End If
End Sub

'Private Sub CmdDel_Click()
'  On Error GoTo UpdateErr
'   sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'   If AdoLiquidacion.Recordset!estado_codigo = "SI" Then
'      If sino = vbYes Then
'         AdoLiquidacion.Recordset!estado_codigo = "NL"
'         AdoLiquidacion.Recordset!fecha_registro = Date
'         AdoLiquidacion.Recordset!usr_codigo = GlUsuario
'         AdoLiquidacion.Recordset.UpdateBatch adAffectAll
'      End If
'   Else
'      MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'End Sub

'Private Sub cmdDesaprueba_Click()
'  On Error GoTo UpdateErr
'   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'   If rs_Auxiliar!estado_codigo = "SI" Then
'      If sino = vbYes Then
'         rs_Auxiliar!estado_codigo = "NO"
'         rs_Auxiliar!fecha_registro = Date
'         rs_Auxiliar!usr_codigo = GlUsuario
'         rs_Auxiliar.UpdateBatch adAffectAll
'      End If
'   Else
'        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'End Sub


Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If txtSW = "ADD" Then
    rw_ficha_rrhh.AdoLiquidacion.Recordset.AddNew
    rw_ficha_rrhh.AdoLiquidacion.Recordset!ges_gestion_ini = TxtGestion_ini.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!ges_gestion = TxtGestion.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!fecha_ingreso = DTPFInicio.Value
      rw_ficha_rrhh.AdoLiquidacion.Recordset!fecha_retiro = DTPFFin.Value
      rw_ficha_rrhh.AdoLiquidacion.Recordset!beneficiario_codigo = txtBenef.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!ges_gestion = glGestion
      rw_ficha_rrhh.AdoLiquidacion.Recordset!tipo_memo = DtcRetiro.Text
      
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ro_contratos_personas WHERE beneficiario_codigo = '" & txtBenef.Text & "'  ", db, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            rw_ficha_rrhh.AdoLiquidacion.Recordset!numero_consultoria = rs_correlativo.RecordCount
'            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
'            rs_correlativo.Update
'            rs_M1!Numero_FA = rs_correlativo!correlativo
      Else
            rw_ficha_rrhh.AdoLiquidacion.Recordset!numero_consultoria = 1
      End If
      rw_ficha_rrhh.AdoLiquidacion.Recordset!ARCHIVO = "Cargar_Archivo"
      rw_ficha_rrhh.AdoLiquidacion.Recordset!ARCHIVO_NOMB = Trim(TxtInicial.Text) & "_Finiquito_" & rw_ficha_rrhh.AdoLiquidacion.Recordset!numero_consultoria & ".pdf"
      TxtAprob.Text = "NO"
    End If
       rw_ficha_rrhh.AdoLiquidacion.Recordset!ges_gestion_ini = TxtGestion_ini.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!ges_gestion = TxtGestion.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!monto_mensual = Txtpago3.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Años = CmbAño.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!meses = CmbMes.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!DIAS = CmbDia.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Mes_Antepenultimo = CmbMes1.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Mes_Penultimo = CmbMes2.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Mes_Utimo = CmbMes3.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Pago_Antepenultimo = IIf((txtpago1.Text = ""), "0", txtpago1.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Pago_Penultimo = IIf((TxtPago2.Text = ""), "0", TxtPago2.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Pago_Utimo = IIf((Txtpago3.Text = ""), "0", Txtpago3.Text)
      If DtcRetiro.Text <> "QUI" Then
      rw_ficha_rrhh.AdoLiquidacion.Recordset!dias_agui = txt_dias_agui.Text
      rw_ficha_rrhh.AdoLiquidacion.Recordset!meses_agui = txt_meses_agui
      Else
      
      End If
'      If txtpago4.Text = "" Then
'      txtpago4.Text = "0"
'      End If
'
'      If txtpago5.Text = "" Then
'      txtpago5.Text = "0"
'      End If
'
'      If txtpago6.Text = "" Then
'      txtpago6.Text = "0"
'      End If
      
      rw_ficha_rrhh.AdoLiquidacion.Recordset!OtroPago_Antep = IIf((txtpago4.Text = ""), "0", txtpago4.Text)
'      If GlTipoCambioOficial > 0 Then
'        AdoLiquidacion.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      Else
'        GlTipoCambioOficial = 7.05
'        AdoLiquidacion.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      End If
      rw_ficha_rrhh.AdoLiquidacion.Recordset!OtroPago_Penul = IIf((txtpago5.Text = ""), "0", txtpago5.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!OtroPago_Utimo = IIf((txtpago6.Text = ""), "0", txtpago6.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Desah_3Meses = "0"
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Imdem_Año = IIf((TxtImdemAño.Text = ""), "0", TxtImdemAño.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Imdem_Mes = IIf((TxtImdemMes.Text = ""), "0", TxtImdemMes.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Indem_dias = IIf((TxtImdemDia.Text = ""), "0", TxtImdemDia.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Aguin_Navidad = IIf((TxtNavidad.Text = ""), "0", TxtNavidad.Text)
      
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Aguin_Vacacion = IIf((TxtVacacion.Text = ""), "0", TxtVacacion.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Prima_Legal = IIf((TxtPrima.Text = ""), "0", TxtPrima.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Otros_Pagos = IIf((TxtOtros.Text = ""), "0", TxtOtros.Text)
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Forma_pago = CmbChq_Trf
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Num_chq_cmpbte = TxtNo_Chq
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Cta_Codigo = TxtCta
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Deducciones = IIf((TxtDeduccion.Text = ""), "0", TxtDeduccion.Text)
      
      rw_ficha_rrhh.AdoLiquidacion.Recordset!monto_total = TxtTotBenef
      rw_ficha_rrhh.AdoLiquidacion.Recordset!Fecha_Liquidacion = DTPFechaLiq.Value
      rw_ficha_rrhh.AdoLiquidacion.Recordset!hora_registro = "8:00"
       rw_ficha_rrhh.AdoLiquidacion.Recordset!estado_codigo = "REG"
      rw_ficha_rrhh.AdoLiquidacion.Recordset!fecha_registro = Date
      rw_ficha_rrhh.AdoLiquidacion.Recordset!usr_codigo = glusuario
      rw_ficha_rrhh.AdoLiquidacion.Recordset.Update    'Batch adAffectAll
      
      mbDataChanged = False
    
'      Fra_ABM.Enabled = False
       
      Unload Me
  End If
   rw_ficha_rrhh.abrirtabla
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If txtBenef.Text = "" Then
    MsgBox "Debe registrar a la persona Beneficiaria ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If TxtTotBenef.Text = "" Then
    MsgBox "Debe registrar el Monto Total de la Liquidacion ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If DTPFInicio.Value > DTPFFin.Value Then
    MsgBox "La Fecha de Retiro NO puede ser Mayor a la de Ingreso  ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
'  If DTPFInicio.Value > DTPFFin.Value Then
'    MsgBox "La Fecha de Inicio NO puede ser Mayor a la de Finalizacion del Contrato ...", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If

End Sub

Private Sub CmbAño_LostFocus()
    '1.2. Indemnización:  S3UM/3 * AñosIndenización * (S3UM/3) / 12*MeseIndemnización + (S3UM/3) / 360* DíasIndemnización.
    TxtImdemAño.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * Val(CmbAño.Text)
End Sub

Private Sub CmbMes2_LostFocus()
    '1.2. Indemnización:  S3UM/3 * AñosIndenización * (S3UM/3) / 12*MeseIndemnización + (S3UM/3) / 360* DíasIndemnización.
    TxtImdemMes.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * Val(CmbMes.Text)
End Sub

Private Sub CmbMes3_Change()
If TxtImdemDia.Text <> "" Then
    '1.2. Indemnización:  S3UM/3 * AñosIndenización * (S3UM/3) / 12*MeseIndemnización + (S3UM/3) / 360* DíasIndemnización.
    TxtImdemDia.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * Val(CmbDia.Text)
End If
End Sub

'Private Sub CmdMod_Click()
'  On Error GoTo EditErr
'  If Ado_Auxiliar.Recordset!estado_codigo = "SI" Then
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

'Private Sub CmdSal_Click()
''  If glPersNew = "O" Then
''    frmmo_pacientes.Dtc_ocupac = AdoLiquidacion.Recordset!ocup_codigo
''    frmmo_pacientes.Dtc_OcupacDes = AdoLiquidacion.Recordset!ocup_descripcion
''  End If
''  glPersNew = "N"
'  Unload Me
'End Sub

Private Sub CmdVerDisco_Click()
  On Error GoTo Error_Sub
  
  If rw_ficha_rrhh.AdoLiquidacion.Recordset!ARCHIVO = "Cargar_Archivo" Then
     NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
     Frmexporta.DirDestino.Path = NombreCarpeta
     GlArch = "LQD"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
         DirLiq = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
      Else
         DirLiq = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = DirLiq
     Frmexporta.Show vbModal
  Else
'    MsgBox ""
     sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
     If sino = vbYes Then
        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        GlArch = "LQD"
        'If GlServidor <> GlMaquina Then      ' "SRVPRO" Then
        If GlServidor = "SRVPRO" Then
           DirLiq = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
        Else
           DirLiq = NombreCarpeta
        End If
        Frmexporta.DirDestino2.Path = DirLiq
        Frmexporta.Show vbModal
     End If
  End If

  Exit Sub
Error_Sub:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub DTCFFin_LostFocus()
    'DTPFFin.Value = DTCFFin.Text
     'DTPFFin.Value = IIf(DTPFFin.Text = "", Date, DTPFFin.Text)
End Sub

Private Sub DTCFInicio_LostFocus()
    DTPFInicio.Value = IIf(DTCFInicio.Text = "", Date, DTCFInicio.Text)
End Sub

Private Sub DtcRetiro_Click(Area As Integer)
    DtcRetiroDes.BoundText = DtcRetiro.BoundText
End Sub

Private Sub DtcRetiroDes_Click(Area As Integer)
    DtcRetiro.BoundText = DtcRetiroDes.BoundText
    If DtcRetiro.Text <> "QUI" Then
    btn_total.Visible = True
    TxtGestion_ini.Visible = False
    lblLabels(13).Caption = "Gestion"
    lblLabels(29).Caption = ""
    Else
    btn_total.Visible = False
    TxtGestion_ini.Visible = True
    lblLabels(13).Caption = "Gestion Fin"
    lblLabels(29).Caption = "Gestion Ini"
    End If
    TxtTotBenef.Text = "0"
    TxtDeduccion.Text = "0"
    If DtcRetiroDes.Text <> "" Then
    Call calcular
    Call datod_plmilla
    txt_promedio = Round((Val(txtpago1) + Val(TxtPago2) + Val(txtpago1)) / 3, 2)
    txtDesahucio.Text = Round(Val(txt_promedio.Text) * 3, 2)
    End If
End Sub

Private Sub DtcRetiroDes_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
'Txt01.Text = ""
'Txt01.Text = UCase(MonthName(Month(Date)))
KeyAscii = 0
End If
End Sub

Private Sub DtcRetiroDes_LostFocus()
'    'MES ULTIMO
'    VAR_MES00 = VAR_MES0 - 1
'    Set rs_aux1 = New ADODB.Recordset
'    rs_aux1.Open "select * from ro_pagos_cronograma_Detalle where beneficiario_codigo = '" & txtBenef & "' and mes_grupo = " & MES3 & " AND ges_gestion = " & GESTION3 & "", db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        Txtpago3.Text = rs_aux1!total_ganado
'    Else
'        Txtpago3.Text = 0
'    End If
'    'MES PENULTIMO
'    VAR_MES00 = VAR_MES0 - 2
'    Set rs_aux1 = New ADODB.Recordset
'    rs_aux1.Open "select * from ro_pagos_cronograma_Detalle where beneficiario_codigo = '" & txtBenef & "' and mes_grupo = " & mes2 & " AND ges_gestion = " & GESTION2 & "", db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        TxtPago2.Text = rs_aux1!total_ganado
'    Else
'        TxtPago2.Text = 0
'    End If
'    'MES ANTEPENULTIMO
'    VAR_MES00 = VAR_MES0 - 3
'    Set rs_aux1 = New ADODB.Recordset
'    rs_aux1.Open "select * from ro_pagos_cronograma_Detalle where beneficiario_codigo = '" & txtBenef & "' and mes_grupo = " & mes1 & "  AND ges_gestion = " & GESTION1 & "", db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        txtpago1.Text = rs_aux1!total_ganado
'    Else
'        txtpago1.Text = 0
'    End If
'
'    If DtcRetiro.Text = "QUI" Then     'DtcRetiro
'        TxtTotBenef.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * 5
'        Frame4.Visible = False
'    Else
'        'If rw_ficha_rrhh.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
'        Frame4.Visible = True
'        If DtcRetiro.Text = "REF" Then
'            VAR_RETIRO = 1
'        Else
'            VAR_RETIRO = 0
'        End If
'        txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
'        VAR_DIA = Day(DTCFInicio)
'        VAR_MES = Month(DTCFInicio)
'        VAR_ANIO = Year(DTCFInicio)
'        VAR_DIA2 = Day(DTPFechaLiq)
'        VAR_MES2 = Month(DTPFechaLiq)
'        VAR_ANIO2 = Year(DTPFechaLiq)
'        If Val(TxtGestion.Text) > VAR_ANIO Then
'            'MesesAguinaldo* ((S3UM/3)/12)+DíasAguinaldo* ((S3UM/3)/ 360)
'            VAR_MES3 = VAR_MES2 * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 12
'            VAR_DIA3 = VAR_DIA2 * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
'            TxtNavidad.Text = VAR_MES3 + VAR_DIA3
'        Else
'            If VAR_DIA > VAR_DIA2 Then
'                Select Case VAR_MES2
'                  Case 1, 3, 5, 7, 8, 10, 12
'                      VAR_NRODIAS = 31
'                  Case 2
'                      VAR_NRODIAS = 28
'                  Case 4, 6, 9, 11
'                      VAR_NRODIAS = 30
'                End Select
'                VAR_MES3 = (VAR_MES2 - VAR_MES - 1) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 12
'                VAR_DIA3 = (VAR_DIA2 - VAR_DIA2 + VAR_NRODIAS) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
'            Else
'                VAR_MES3 = (VAR_MES2 - VAR_MES) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 12
'                VAR_DIA3 = (VAR_DIA2 - VAR_DIA2) * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
'            End If
'            TxtNavidad.Text = VAR_MES3 + VAR_DIA3
'        End If
'        'DíasVacaciónFiniquito*((S3UM/3)/ 360)
'        Set rs_vacaciones_prog = New ADODB.Recordset
'        If rs_vacaciones_prog.State = 1 Then rs_vacaciones_prog.Close
'        rs_vacaciones_prog.Open "select * from ro_vacaciones_programadas where beneficiario_codigo = '" & txtBenef & "' ", db, adOpenKeyset, adLockOptimistic
'        If rs_vacaciones_prog.RecordCount > 0 Then
'            TxtVacacion.Text = rs_vacaciones_prog!Dias_Pendientes * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
'        Else
'            TxtVacacion.Text = "0"
'        End If
'        'Set Ado_VacacionesProg.Recordset = rs_vacaciones_prog
'        'Set DtgVacacionesProg.DataSource = Ado_VacacionesProg.Recordset
'    End If
    
    
End Sub

Private Sub DTPFechaLiq_Change()
 If DtcRetiroDes.Text <> "" Then
    Call calcular
    Call datod_plmilla
    End If
End Sub

Private Sub DTPFechaLiq_LostFocus()
'    VAR_MES0 = Month(DTPFechaLiq.Value)
'    GESTION1 = Year(DTPFechaLiq.Value)
'    Select Case VAR_MES0
'        Case 1
'            CmbMes3 = "DICIEMBRE"
'            CmbMes2 = "NOVIEMBRE"
'            CmbMes1 = "OCTUBRE"
'            MES3 = 12
'            mes2 = 11
'            mes1 = 10
'            GESTION3 = Year(DTPFechaLiq.Value) - 1
'            GESTION2 = Year(DTPFechaLiq.Value) - 1
'            GESTION1 = Year(DTPFechaLiq.Value) - 1
'        Case 2
'            CmbMes3 = "ENERO"
'            CmbMes2 = "DICIEMBRE"
'            CmbMes1 = "NOVIEMBRE"
'             MES3 = 1
'            mes2 = 12
'            mes1 = 11
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value) - 1
'            GESTION1 = Year(DTPFechaLiq.Value) - 1
'        Case 3
'            CmbMes3 = "FEBRERO"
'            CmbMes2 = "ENERO"
'            CmbMes1 = "DICIEMBRE"
'             MES3 = 2
'            mes2 = 1
'            mes1 = 12
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value) - 1
'        Case 4
'            CmbMes3 = "MARZO"
'            CmbMes2 = "FEBRERO"
'            CmbMes1 = "ENERO"
'             MES3 = 3
'            mes2 = 2
'            mes1 = 1
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 5
'            CmbMes3 = "ABRIL"
'            CmbMes2 = "MARZO"
'            CmbMes1 = "FEBRERO"
'            MES3 = 4
'            mes2 = 3
'            mes1 = 2
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 6
'            CmbMes3 = "MAYO"
'            CmbMes2 = "ABRIL"
'            CmbMes1 = "MARZO"
'             MES3 = 5
'            mes2 = 4
'            mes1 = 3
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 7
'            CmbMes3 = "JUNIO"
'            CmbMes2 = "MAYO"
'            CmbMes1 = "ABRIL"
'            MES3 = 6
'            mes2 = 5
'            mes1 = 4
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 8
'            CmbMes3 = "JULIO"
'            CmbMes2 = "JUNIO"
'            CmbMes1 = "MAYO"
'            MES3 = 7
'            mes2 = 6
'            mes1 = 5
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 9
'            CmbMes3 = "AGOSTO"
'            CmbMes2 = "JULIO"
'            CmbMes1 = "JUNIO"
'            MES3 = 8
'            mes2 = 7
'            mes1 = 6
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 10
'            CmbMes3 = "SEPTIEMBRE"
'            CmbMes2 = "AGOSTO"
'            CmbMes1 = "JULIO"
'            MES3 = 9
'            mes2 = 8
'            mes1 = 7
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 11
'            CmbMes3 = "OCTUBRE"
'            CmbMes2 = "SEPTIEMBRE"
'            CmbMes1 = "AGOSTO"
'            MES3 = 10
'            mes2 = 9
'            mes1 = 8
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'        Case 12
'            CmbMes3 = "NOVIEMBRE"
'            CmbMes2 = "OCTUBRE"
'            CmbMes1 = "SEPTIEMBRE"
'            MES3 = 11
'            mes2 = 10
'            mes1 = 9
'            GESTION3 = Year(DTPFechaLiq.Value)
'            GESTION2 = Year(DTPFechaLiq.Value)
'            GESTION1 = Year(DTPFechaLiq.Value)
'    End Select
End Sub

Private Sub Form_Activate()
 TxtGestion.Text = Year(Date)
 DTPFechaLiq.Value = Date
                   
 DTCFInicio.Text = rw_ficha_rrhh.Ado_datos.Recordset!fecha_ingreso
 DTCFFin.Text = IIf(IsNull(rw_ficha_rrhh.Ado_datos.Recordset!fecha_expiracion), Date, Trim(rw_ficha_rrhh.Ado_datos.Recordset!fecha_expiracion))

  Set rs_beneficiario = New ADODB.Recordset
  rs_beneficiario.Open "select * from gc_Beneficiario WHERE tipoben_codigo = '1' ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockOptimistic
  Set AdoBeneficiario.Recordset = rs_beneficiario.DataSource
'  DtcBenefDes.BoundText = DtcBenef.BoundText
  
  Set rs_motivo = New ADODB.Recordset
  rs_motivo.Open "select * from rc_tipo_memoranda WHERE uso = 'A'  ", db, adOpenKeyset, adLockOptimistic
  'rs_motivo.Open "select * from rc_motivo_proceso WHERE estado_codigo = 'APR' ", db, adOpenKeyset, adLockOptimistic
  Set AdoMotivos.Recordset = rs_motivo.DataSource
  DtcRetiroDes.BoundText = DtcRetiro.BoundText
  
  Set rs_Contato2 = New ADODB.Recordset
  rs_Contato2.Open "select * from ro_contratos_personas where beneficiario_codigo = '" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "' and estado_contrato ='APR' AND Estado_liquidacion= 'REG' ", db, adOpenKeyset, adLockOptimistic
  Set AdoContrato2.Recordset = rs_Contato2.DataSource
'  Dtc_descrip.BoundText = Dtc_codigo.BoundText
  mbDataChanged = False
  GlSW = "NADA"
End Sub

Private Sub Form_Load()

'  Call abrirtabla
  
'  Set rs_beneficiario = New ADODB.Recordset
'  rs_beneficiario.Open "select * from gc_Beneficiario WHERE tipoben_codigo='1' ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockOptimistic
'  Set AdoBeneficiario.Recordset = rs_beneficiario.DataSource
''  DtcBenefDes.BoundText = DtcBenef.BoundText

'  Set rs_motivo = New ADODB.Recordset
'  rs_motivo.Open "select * from rc_motivo_proceso WHERE estado_codigo = 'APR'  ", db, adOpenKeyset, adLockOptimistic
'  Set AdoMotivos.Recordset = rs_motivo.DataSource
''  DtcRetiroDes.BoundText = DtcRetiro.BoundText

'  Set rs_Contato2 = New ADODB.Recordset
'  rs_Contato2.Open "select * from ro_contratos_personas where beneficiario_codigo = '" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "' and estado_contrato ='SI' AND Estado_liquidacion= 'N' ", db, adOpenKeyset, adLockOptimistic
'  Set AdoContrato2.Recordset = rs_Contato2.DataSource
''  Dtc_descrip.BoundText = Dtc_codigo.BoundText
'
'  Set rs_Org = New ADODB.Recordset
'  rs_Org.Open "select * from fc_convenios  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoOrg.Recordset = rs_Org.DataSource
'  DtcOrgDes.BoundText = DtcOrg.BoundText
'
'  Set rs_Pry = New ADODB.Recordset
'  rs_Pry.Open "select * from fc_estructura_programatica  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoPry.Recordset = rs_Pry.DataSource
'  DtcPryDes.BoundText = DtcPry.BoundText

'  AdoLiquidacion.Recordset.AddNew
'  txtParam.Text = GlParametro
'  TxtForm.Text = GlForm
'  TxtCorrel.Text = GlCorrel

  mbDataChanged = False
'  Fra_ABM.Enabled = False
'  DtG_Auxiliar.Enabled = True
  GlSW = "NADA"
  
  If DtcRetiro.Text <> "QUI" Then
    btn_total.Visible = True
    TxtGestion_ini.Visible = False
    lblLabels(13).Caption = "Gestion"
    lblLabels(29).Caption = ""
    Else
    btn_total.Visible = False
    TxtGestion_ini.Visible = True
    lblLabels(13).Caption = "Gestion Fin"
    lblLabels(29).Caption = "Gestion Ini"
    End If
	Call SeguridadSet(Me)
End Sub

'Private Sub abrirtabla()
'  Set AdoLiquidacion.Recordset = New Recordset
'  If AdoLiquidacion.Recordset.State = 1 Then AdoLiquidacion.Recordset.Close
'  'queryinicial = "select * from rc_puesto_organizacional where param_codigo = '" & GlParametro & "' "
'  queryinicial = "select * from ro_liquidaciones "
'  AdoLiquidacion.Recordset.Open queryinicial, DB, adOpenKeyset, adLockOptimistic
'  AdoLiquidacion.Recordset.Sort = "beneficiario_codigo, fecha_ingreso"
'  Set Ado_Auxiliar.Recordset = AdoLiquidacion.Recordset.DataSource
'  Set DtG_Auxiliar.DataSource = Ado_Auxiliar.Recordset
'End Sub

Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
'    frmeo_Larvas_mosquitos.Fra_detalle.Visible = False
End Sub

'Private Sub Ado_Auxiliar_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'Muestra la posición de registro actual para este Recordset
'      Ado_Auxiliar.Caption = Ado_Auxiliar.Recordset.AbsolutePosition & " / " & Ado_Auxiliar.Recordset.RecordCount
'End Sub

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
'    'AdoLiquidacion.Recordset.MoveLast
'    AdoLiquidacion.Recordset.AddNew
'    'lblStatus.Caption = "Agregar registro"
'    Fra_ABM.Enabled = True
'    fraOpciones.Visible = False
'    FraGrabarCancelar.Visible = True
'    DtG_Auxiliar.Enabled = False
'    GlSW = "ADD"
''    AdoLiquidacion.Recordset.AddNew
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
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(rw_ficha_rrhh.AdoLiquidacion.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub



Private Sub txt_dias_agui_Change()
txt_agui_d.Text = Round((Val(txt_promedio) / 360) * Val(txt_dias_agui.Text), 2)
'TxtNavidad.Text = txt_agui_m.Text + txt_agui_d.Text
End Sub

Private Sub txt_meses_agui_Change()
txt_agui_m.Text = Round((Val(txt_promedio) / 12) * Val(txt_meses_agui.Text), 2)
'TxtNavidad.Text = txt_agui_m.Text + txt_agui_d.Text
End Sub

Private Sub txt_promedio_Change()

TxtImdemAño.Text = Round(Val((txt_promedio)) * CmbAño.Text, 2)
TxtImdemMes.Text = Round((Val(txt_promedio) / 12) * CmbMes.Text, 2)
TxtImdemDia.Text = Round((Val(txt_promedio) / 360) * CmbDia.Text, 2)

txt_agui_m.Text = Round((Val(txt_promedio) / 12) * Val(txt_meses_agui.Text), 2)
txt_agui_d.Text = Round((Val(txt_promedio) / 360) * Val(txt_dias_agui.Text), 2)

'Call planilla_actual
TxtNavidad.Text = Val(txt_agui_m.Text) + Val(txt_agui_d.Text)

End Sub

Private Sub txtpago1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub txtpago1_LostFocus()
'    If rw_ficha_rrhh.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
'        VAR_RETIRO = 1
'    Else
'        VAR_RETIRO = 0
'    End If
'    txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
'    If DtcRetiro.Text = "QUI" Then     'DtcRetiro
'        TxtTotBenef.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * 5
'    End If
End Sub

Private Sub TxtPago2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub TxtPago2_LostFocus()
'    If rw_ficha_rrhh.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
'        VAR_RETIRO = 1
'    Else
'        VAR_RETIRO = 0
'    End If
'    txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
'    If DtcRetiro.Text = "QUI" Then     'DtcRetiro
'        TxtTotBenef.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * 5
'    End If
End Sub

Private Sub Txtpago1_Change()
txt_promedio = Round((Val(txtpago1) + Val(TxtPago2) + Val(txtpago1)) / 3, 2)
End Sub
Private Sub Txtpago2_Change()
txt_promedio = Round((Val(txtpago1) + Val(TxtPago2) + Val(txtpago1)) / 3, 2)
End Sub
Private Sub Txtpago3_Change()
txt_promedio = Round((Val(txtpago1) + Val(TxtPago2) + Val(txtpago1)) / 3, 2)
End Sub

Private Sub Txtpago3_KeyUp(KeyCode As Integer, Shift As Integer)
txt_promedio = Round((Val(txtpago1) + Val(TxtPago2) + Val(Txtpago3)) / 3, 2)
End Sub
Private Sub Txtpago2_KeyUp(KeyCode As Integer, Shift As Integer)
txt_promedio = Round((Val(txtpago1) + Val(TxtPago2) + Val(Txtpago3)) / 3, 2)
End Sub
Private Sub txtpago1_KeyUp(KeyCode As Integer, Shift As Integer)
txt_promedio = Round((Val(txtpago1) + Val(TxtPago2) + Val(txtpago1)) / 3, 2)
End Sub


Private Sub Txtpago3_LostFocus()
'    If rw_ficha_rrhh.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
'        VAR_RETIRO = 1
'    Else
'        VAR_RETIRO = 0
'    End If
'    txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
'    If DtcRetiro.Text = "QUI" Then     'DtcRetiro
'        TxtTotBenef.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * 5
'    End If
End Sub

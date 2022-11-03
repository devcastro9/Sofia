VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ro_Personal_Planilla 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RRHH - Planillas - Seleccion de Personas "
   ClientHeight    =   8460
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   9450
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ro_Personal_Planilla.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "ro_Personal_Planilla.frx":6C032
      ScaleHeight     =   915
      ScaleWidth      =   9195
      TabIndex        =   70
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "ro_Personal_Planilla.frx":D8064
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
         Picture         =   "ro_Personal_Planilla.frx":D826E
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
         Picture         =   "ro_Personal_Planilla.frx":D8478
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
         Picture         =   "ro_Personal_Planilla.frx":D8800
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
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   9255
      Begin MSComCtl2.DTPicker DTPFechaLiq 
         DataField       =   "Fecha_Liquidacion"
         DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
         Height          =   285
         Left            =   5265
         TabIndex        =   64
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   91619329
         CurrentDate     =   42307
         MinDate         =   2
      End
      Begin VB.TextBox TxtInicial 
         Height          =   285
         Left            =   4725
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
         Begin VB.ComboBox CmbChq_Trf 
            DataField       =   "Forma_pago"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   240
            TabIndex        =   57
            Text            =   "CHEQUE"
            Top             =   540
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
            Left            =   6960
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   960
            Width           =   1695
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
            Top             =   540
            Width           =   1335
         End
         Begin VB.TextBox TxtDeduccion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Deducciones"
            DataSource      =   "frmBeneficiario_admin.AdoLiquidacion"
            Height          =   315
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   960
            Width           =   1455
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
            Top             =   540
            Width           =   3375
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
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
            Left            =   4080
            TabIndex        =   55
            Top             =   1005
            Width           =   2760
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
            Top             =   280
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
            Top             =   280
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
            Top             =   280
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
         Top             =   3720
         Width           =   9015
         Begin VB.TextBox txtDesahucio 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Imdem_Año"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   315
            Left            =   4800
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   1725
            Width           =   1455
         End
         Begin VB.ComboBox CmbDia 
            DataField       =   "dias"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   285
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1140
            Width           =   1695
         End
         Begin VB.ComboBox CmbAño 
            DataField       =   "Años"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "ro_Personal_Planilla.frx":D8B88
            Left            =   1800
            List            =   "ro_Personal_Planilla.frx":D8B8A
            TabIndex        =   32
            Text            =   "0"
            Top             =   765
            Width           =   900
         End
         Begin VB.TextBox TxtVacacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Aguin_Vacacion"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            Left            =   4800
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
         Top             =   2040
         Width           =   9015
         Begin VB.ComboBox CmbMes3 
            DataField       =   "Mes_Utimo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   315
            ItemData        =   "ro_Personal_Planilla.frx":D8B8C
            Left            =   6720
            List            =   "ro_Personal_Planilla.frx":D8BB4
            TabIndex        =   78
            Text            =   "MARZO"
            Top             =   480
            Width           =   1980
         End
         Begin VB.ComboBox CmbMes2 
            DataField       =   "Mes_Penultimo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   315
            ItemData        =   "ro_Personal_Planilla.frx":D8C1D
            Left            =   4200
            List            =   "ro_Personal_Planilla.frx":D8C45
            TabIndex        =   77
            Text            =   "FEBRERO"
            Top             =   480
            Width           =   1980
         End
         Begin VB.TextBox txtpago6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Utimo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   285
            Left            =   6720
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtpago5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Penul"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   285
            Left            =   4200
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtpago4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Antep"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   285
            Left            =   1600
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Txtpago3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Utimo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   285
            Left            =   6720
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   850
            Width           =   1695
         End
         Begin VB.TextBox TxtPago2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Penultimo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   285
            Left            =   4200
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   850
            Width           =   1695
         End
         Begin VB.ComboBox CmbMes1 
            DataField       =   "Mes_Antepenultimo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   315
            ItemData        =   "ro_Personal_Planilla.frx":D8CAE
            Left            =   1600
            List            =   "ro_Personal_Planilla.frx":D8CD6
            TabIndex        =   19
            Text            =   "ENERO"
            Top             =   500
            Width           =   1980
         End
         Begin VB.TextBox txtpago1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Pago_Antepenultimo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   285
            Left            =   1600
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   850
            Width           =   1695
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
            Left            =   6720
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
            Left            =   4200
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
            Left            =   1560
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
            Left            =   240
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
            Left            =   240
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
            Left            =   240
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
         Top             =   1035
         Width           =   9015
         Begin MSDataListLib.DataCombo DTCFInicio 
            Bindings        =   "ro_Personal_Planilla.frx":D8D3F
            DataField       =   "fecha_inicio"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   315
            Left            =   240
            TabIndex        =   68
            Top             =   540
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
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
            Bindings        =   "ro_Personal_Planilla.frx":D8D5A
            DataField       =   "fecha_fin"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
            Height          =   315
            Left            =   2025
            TabIndex        =   69
            Top             =   540
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
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
            DataField       =   "tipo_memo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            DataField       =   "tipo_memo"
            DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
            Format          =   91619329
            CurrentDate     =   40471
         End
         Begin MSComCtl2.DTPicker DTPFFin 
            DataField       =   "fecha_retiro"
            Height          =   285
            Left            =   2025
            TabIndex        =   11
            Top             =   540
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   91619329
            CurrentDate     =   40471
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Fecha Retiro"
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
            Left            =   2070
            TabIndex        =   14
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Motivo de Retiro"
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
            Width           =   1470
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
         DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
         Height          =   315
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "REG"
         Top             =   540
         Width           =   495
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
         Height          =   315
         ItemData        =   "ro_Personal_Planilla.frx":D8D75
         Left            =   3705
         List            =   "ro_Personal_Planilla.frx":D8D9A
         TabIndex        =   0
         Text            =   "2015"
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox TxtLquida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DataField       =   "id_liquidacion"
         DataSource      =   "frm_ro_LiquidaMain.AdoBeneficiario"
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
         Top             =   540
         Width           =   1335
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
         Height          =   240
         Index           =   14
         Left            =   5175
         TabIndex        =   63
         Top             =   285
         Width           =   1650
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gestión Liq."
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
         Index           =   13
         Left            =   3660
         TabIndex        =   62
         Top             =   285
         Width           =   1050
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
         Height          =   240
         Index           =   12
         Left            =   2360
         TabIndex        =   6
         Top             =   285
         Width           =   645
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
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   280
         Width           =   1410
      End
   End
   Begin MSAdodcLib.Adodc Ado_pagos_detalle 
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
      Caption         =   "Ado_pagos_detalle"
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
Attribute VB_Name = "ro_Personal_Planilla"
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
Dim rs_pagos_detalle As ADODB.Recordset
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

Dim DirLiq As String
Dim VAR_VAL As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean
Private Sub clculo_finiquito()

  Set rs_pagos_detalle = New ADODB.Recordset
  rs_pagos_detalle.Open "select * from ro_pagos_cronograma_Detalle WHERE tipoben_codigo = '1' ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockOptimistic
  Set Ado_pagos_detalle.Recordset = rs_pagos_detalle.DataSource

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
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!fecha_ingreso = DTPFInicio.Value
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!fecha_retiro = DTPFFin.Value
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!beneficiario_codigo = txtBenef.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!ges_gestion = glGestion
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!tipo_memo = DtcRetiro.Text
      
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ro_contratos_personas WHERE beneficiario_codigo = '" & txtBenef.Text & "'  ", db, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            frmBeneficiario_Admin.AdoLiquidacion.Recordset!numero_consultoria = rs_correlativo.RecordCount
'            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
'            rs_correlativo.Update
'            rs_M1!Numero_FA = rs_correlativo!correlativo
      Else
            frmBeneficiario_Admin.AdoLiquidacion.Recordset!numero_consultoria = 1
      End If
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!ARCHIVO = "Cargar_Archivo"
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!ARCHIVO_NOMB = Trim(TxtInicial.Text) & "_Finiquito_" & frmBeneficiario_Admin.AdoLiquidacion.Recordset!numero_consultoria & ".pdf"
      TxtAprob.Text = "NO"
    End If
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!monto_mensual = Txtpago3.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Años = CmbAño.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!meses = CmbMes.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!DIAS = CmbDia.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Mes_Antepenultimo = CmbMes1.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Mes_Penultimo = CmbMes2.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Mes_Utimo = CmbMes3.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Pago_Antepenultimo = txtpago1.Text
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Pago_Penultimo = TxtPago2
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Pago_Utimo = Txtpago3
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!OtroPago_Antep = txtpago4.Text
'      If GlTipoCambioOficial > 0 Then
'        AdoLiquidacion.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      Else
'        GlTipoCambioOficial = 7.05
'        AdoLiquidacion.Recordset!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
'      End If
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!OtroPago_Penul = txtpago5
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!OtroPago_Utimo = txtpago6
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Desah_3Meses = "0"
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Imdem_Año = TxtImdemAño
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Imdem_Mes = TxtImdemMes
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Indem_dias = IIf((TxtImdemDia = ""), "0", TxtImdemDia)
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Aguin_Navidad = TxtNavidad
      
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Aguin_Vacacion = TxtVacacion
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Prima_Legal = TxtPrima
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Otros_Pagos = TxtOtros
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Forma_pago = CmbChq_Trf
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Num_chq_cmpbte = TxtNo_Chq
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Cta_Codigo = TxtCta
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Deducciones = TxtDeduccion
      
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!monto_total = TxtTotBenef
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!Fecha_Liquidacion = DTPFechaLiq.Value
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!hora_registro = "8:00"
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!fecha_registro = Date
      frmBeneficiario_Admin.AdoLiquidacion.Recordset!usr_codigo = glusuario
      frmBeneficiario_Admin.AdoLiquidacion.Recordset.Update    'Batch adAffectAll
      
      mbDataChanged = False
    
'      Fra_ABM.Enabled = False
      Unload Me
  End If
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
    '1.2. Indemnización:  S3UM/3 * AñosIndenización * (S3UM/3) / 12*MeseIndemnización + (S3UM/3) / 360* DíasIndemnización.
    TxtImdemDia.Text = ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) * Val(CmbDia.Text)
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
  
  If frmBeneficiario_Admin.AdoLiquidacion.Recordset!ARCHIVO = "Cargar_Archivo" Then
     NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
     Frmexporta.DirDestino.Path = NombreCarpeta
     GlArch = "LQD"
      'If GlServidor <> GlMaquina Then      ' "-" Then
      If GlServidor = "SRVPRO" Then
         DirLiq = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
      Else
         DirLiq = NombreCarpeta
      End If
      Frmexporta.DirDestino2.Path = DirLiq
     Frmexporta.Show vbModal
  Else
'    MsgBox ""
     sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
     If sino = vbYes Then
        NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
        Frmexporta.DirDestino.Path = NombreCarpeta
        GlArch = "LQD"
        'If GlServidor <> GlMaquina Then      ' "SRVPRO" Then
        If GlServidor = "SRVPRO" Then
           DirLiq = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial) & "-" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\"
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
    DTPFFin.Value = DTCFFin.Text
End Sub

Private Sub DTCFInicio_LostFocus()
    DTPFInicio.Value = IIf(DTCFInicio.Text = "", Date, DTCFInicio.Text)
End Sub

Private Sub DtcRetiro_Click(Area As Integer)
    DtcRetiroDes.BoundText = DtcRetiro.BoundText
End Sub

Private Sub DtcRetiroDes_Click(Area As Integer)
    DtcRetiro.BoundText = DtcRetiroDes.BoundText
End Sub

Private Sub DtcRetiroDes_LostFocus()
    If frmBeneficiario_Admin.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
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
    rs_vacaciones_prog.Open "select * from ro_vacaciones_programadas where beneficiario_codigo = '" & txtBenef & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_vacaciones_prog.RecordCount > 0 Then
        TxtVacacion.Text = rs_vacaciones_prog!Dias_Pendientes * ((CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) / 3) / 360
    Else
        TxtVacacion.Text = "0"
    End If
    'Set Ado_VacacionesProg.Recordset = rs_vacaciones_prog
    'Set DtgVacacionesProg.DataSource = Ado_VacacionesProg.Recordset
    
    
    
End Sub

Private Sub Form_Activate()
  Set rs_beneficiario = New ADODB.Recordset
  rs_beneficiario.Open "select * from gc_Beneficiario WHERE tipoben_codigo = '1' ORDER BY beneficiario_denominacion ", db, adOpenKeyset, adLockOptimistic
  Set AdoBeneficiario.Recordset = rs_beneficiario.DataSource
'  DtcBenefDes.BoundText = DtcBenef.BoundText
  
  Set rs_motivo = New ADODB.Recordset
  rs_motivo.Open "select * from rc_tipo_memoranda WHERE uso = 'A'  ", db, adOpenKeyset, adLockOptimistic
  Set AdoMotivos.Recordset = rs_motivo.DataSource
  DtcRetiroDes.BoundText = DtcRetiro.BoundText
  
  Set rs_Contato2 = New ADODB.Recordset
  rs_Contato2.Open "select * from ro_contratos_personas where beneficiario_codigo = '" & frmBeneficiario_Admin.Adolista.Recordset!beneficiario_codigo & "' and estado_contrato ='APR' AND Estado_liquidacion= 'REG' ", db, adOpenKeyset, adLockOptimistic
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
'  rs_Contato2.Open "select * from ro_contratos_personas where beneficiario_codigo = '" & frmBeneficiario_Admin.adoLista.Recordset!beneficiario_codigo & "' and estado_contrato ='SI' AND Estado_liquidacion= 'N' ", db, adOpenKeyset, adLockOptimistic
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
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!beneficiario_codigo) & "\FINIQUITO\" & Trim(frmBeneficiario_Admin.AdoLiquidacion.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub

Private Sub txtpago1_LostFocus()
    If frmBeneficiario_Admin.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
        VAR_RETIRO = 1
    Else
        VAR_RETIRO = 0
    End If
    txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
End Sub

Private Sub TxtPago2_LostFocus()
    If frmBeneficiario_Admin.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
        VAR_RETIRO = 1
    Else
        VAR_RETIRO = 0
    End If
    txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
End Sub

Private Sub Txtpago3_LostFocus()
    If frmBeneficiario_Admin.AdoLiquidacion.Recordset!tipo_memo = "REF" Then
        VAR_RETIRO = 1
    Else
        VAR_RETIRO = 0
    End If
    txtDesahucio.Text = (CDbl(txtpago1.Text) + CDbl(TxtPago2.Text) + CDbl(Txtpago3.Text)) * VAR_RETIRO
End Sub

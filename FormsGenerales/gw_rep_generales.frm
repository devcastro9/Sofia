VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form gw_rep_generales 
   Caption         =   "Reportes de Uso General"
   ClientHeight    =   10620
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   16365
   Icon            =   "gw_rep_generales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   16365
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Elija..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   2460
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   600
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox cbo_mes_rep 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "gw_rep_generales.frx":0A02
         Left            =   120
         List            =   "gw_rep_generales.frx":0A2D
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox cmb_gestion_rep 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "gw_rep_generales.frx":0A9D
         Left            =   600
         List            =   "gw_rep_generales.frx":0AC2
         TabIndex        =   27
         Text            =   "2017"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label18 
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
         Left            =   480
         TabIndex        =   32
         Top             =   885
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "GESTION"
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
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   165
         Width           =   990
      End
   End
   Begin VB.Frame Fra_04 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2760
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   15975
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "COBRADORES"
         Height          =   255
         Left            =   13080
         TabIndex        =   25
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Image rep_cobr 
         Height          =   1455
         Left            =   12960
         Picture         =   "gw_rep_generales.frx":0B08
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "EJECUTIVO VENTAS"
         Height          =   255
         Left            =   10080
         TabIndex        =   24
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CLIENTES"
         Height          =   255
         Left            =   6960
         TabIndex        =   23
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "EQUIPOS"
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "EDIFICIOS"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Image rep_cli 
         Height          =   1455
         Left            =   6960
         Picture         =   "gw_rep_generales.frx":3347
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_resp 
         Height          =   1455
         Left            =   10080
         Picture         =   "gw_rep_generales.frx":5FFB
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_eqp 
         Height          =   1455
         Left            =   3720
         Picture         =   "gw_rep_generales.frx":7973
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_edif 
         Height          =   1455
         Left            =   360
         Picture         =   "gw_rep_generales.frx":AAD2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Fra_03 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   15975
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFC0&
         X1              =   9600
         X2              =   9600
         Y1              =   0
         Y2              =   4800
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CANTIDAD DE COBRANZAS"
         Height          =   255
         Left            =   11640
         TabIndex        =   20
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CANTIDAD DE FACTURAS"
         Height          =   255
         Left            =   13200
         TabIndex        =   19
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CANTIDAD DE CONTRATOS"
         Height          =   255
         Left            =   10080
         TabIndex        =   18
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTRATOS, FACTURACION Y COBRANZAS EN $"
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "FACTURACION Y COBRANZAS EN $"
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTRATOS Y FACTURACION EN $"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "COBRANZAS EN $"
         Height          =   255
         Left            =   6840
         TabIndex        =   14
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "FACTURACION EN $"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lbl_vta 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTRATOS EN $"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Image rep_fac2 
         Height          =   1335
         Left            =   13080
         Picture         =   "gw_rep_generales.frx":E447
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_vta_fac_cob 
         Height          =   1335
         Left            =   6840
         Picture         =   "gw_rep_generales.frx":1055A
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Image rep_cob 
         Height          =   1335
         Left            =   6840
         Picture         =   "gw_rep_generales.frx":122AB
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_vta2 
         Height          =   1335
         Left            =   10080
         Picture         =   "gw_rep_generales.frx":131E3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_cob2 
         Height          =   1335
         Left            =   11520
         Picture         =   "gw_rep_generales.frx":147DB
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Image rep_fac_cob 
         Height          =   1335
         Left            =   3720
         Picture         =   "gw_rep_generales.frx":169F2
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Image rep_fac 
         Height          =   1335
         Left            =   3720
         Picture         =   "gw_rep_generales.frx":190EF
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_vta 
         Height          =   1335
         Left            =   480
         Picture         =   "gw_rep_generales.frx":1A55D
         Stretch         =   -1  'True
         ToolTipText     =   "CONTRATOS EN $"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image rep_vta_fac 
         Height          =   1335
         Left            =   480
         Picture         =   "gw_rep_generales.frx":1B889
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2295
      End
   End
   Begin VB.Frame Fra_02 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label lbl_serv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por Servicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   300
         Left            =   345
         TabIndex        =   4
         Top             =   3945
         Width           =   1440
      End
      Begin VB.Label lbl_unid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por Unidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   300
         Left            =   345
         TabIndex        =   3
         Top             =   6045
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lbl_dpto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por Regional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   300
         Left            =   285
         TabIndex        =   2
         Top             =   2025
         Width           =   1560
      End
      Begin VB.Image opt_serv 
         Height          =   975
         Left            =   240
         Picture         =   "gw_rep_generales.frx":1DF8F
         Stretch         =   -1  'True
         Top             =   3600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image opt_unid 
         Height          =   1095
         Left            =   240
         Picture         =   "gw_rep_generales.frx":23F71
         Stretch         =   -1  'True
         Top             =   5640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image opt_dpto 
         Height          =   975
         Left            =   240
         Picture         =   "gw_rep_generales.frx":296E3
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport CryUnidad 
      Left            =   120
      Top             =   10080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Fra_01 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
      Begin VB.Label lbl_ges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por Gestión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   225
         TabIndex        =   8
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label lbl_mes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por Meses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   300
         Left            =   225
         TabIndex        =   7
         Top             =   4935
         Width           =   1275
      End
      Begin VB.Image opt_mes 
         Height          =   975
         Left            =   120
         Picture         =   "gw_rep_generales.frx":2ED5D
         Stretch         =   -1  'True
         Top             =   4605
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image opt_ges 
         Height          =   975
         Left            =   120
         Picture         =   "gw_rep_generales.frx":3453B
         Stretch         =   -1  'True
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.2. REPORTES VARIOS: de Uso General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   4560
      TabIndex        =   10
      Top             =   6480
      Width           =   5880
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0. REPORTES GENERALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   3465
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.1. REPORTES DE: Contratos de Ventas, Facturación y Cobranzas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   9420
   End
   Begin VB.Image Image8 
      Height          =   855
      Left            =   120
      Picture         =   "gw_rep_generales.frx":38765
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "gw_rep_generales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_GES, VAR_MES As String
Dim VAR_DPTO, VAR_UNID, VAR_SERV As String
Dim VAR_OPT1, VAR_OPT2 As Integer

Private Sub Form_Load()
    VAR_GES = "0"
    VAR_MES = "0"
    VAR_DPTO = "0"
    VAR_UNID = "0"
    VAR_SERV = "0"
    'VAR_OPT1 = 0
    VAR_OPT2 = 0
    Fra_02.Visible = True
    opt_mes.Visible = True
    opt_ges.Visible = False
    Frame2.Visible = True
    VAR_OPT1 = 2
	Call SeguridadSet(Me)
End Sub

Private Sub lbl_dpto_Click()
    Fra_03.Visible = True
    Fra_04.Visible = True
    opt_dpto.Visible = True
    opt_serv.Visible = False
    opt_unid.Visible = False
    'VAR_DPTO = "1"
    'VAR_UNID = "0"
    'VAR_SERV = "0"
    VAR_OPT2 = 1
End Sub

Private Sub lbl_ges_Click()
    Fra_02.Visible = True
    opt_mes.Visible = False
    opt_ges.Visible = True
    Frame2.Visible = False
    VAR_OPT1 = 1
    'VAR_GES = "1"
    'VAR_MES = "0"
End Sub

Private Sub lbl_mes_Click()
    Fra_02.Visible = True
    opt_mes.Visible = True
    opt_ges.Visible = False
    Frame2.Visible = True
    VAR_OPT1 = 2
    'VAR_MES = "1"
    'VAR_GES = "0"
End Sub

Private Sub lbl_serv_Click()
    Fra_03.Visible = True
    Fra_04.Visible = True
    opt_dpto.Visible = False
    opt_serv.Visible = True
    opt_unid.Visible = False
    VAR_OPT2 = 2
    'VAR_DPTO = "0"
    'VAR_UNID = "0"
    'VAR_SERV = "1"
End Sub

Private Sub lbl_unid_Click()
    Fra_03.Visible = True
    Fra_04.Visible = True
    opt_dpto.Visible = False
    opt_serv.Visible = False
    opt_unid.Visible = True
    VAR_OPT2 = 3
    'VAR_DPTO = "0"
    'VAR_UNID = "1"
    'VAR_SERV = "0"
End Sub

Private Sub rep_cob_Click()
'VENTAS EN $
    Select Case VAR_OPT1
        Case 1
            'POR GESTIONES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_depto_y_mes_Bs.rpt"
                    titulo2 = "DEPARTAMENTOS VS. MESES"
                    subtitulo2 = "VENTAS ACUMULADAS EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 2
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 2
                    'POR SERVICIOS
                Case 3
                    'POR UNIDADES
                Case Else
                    '**no identificado**"
            End Select
        Case 2
            'POR MESES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS 'gr_ventas_por_depto_y_mes_Fac_Bs.rpt
                    'CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_depto_y_mes_Fac_Bs.rpt"
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Cobro_por_depto_y_mes_Bs.rpt"
                    titulo2 = "DEPARTAMENTOS VS. MESES"
                    subtitulo2 = "COBRANZA ACUMULADA EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 3
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 2
                    'POR SERVICIOS
                    'CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes_Bs.rpt"
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Cobro_por_proceso_y_mes_Bs.rpt"
                    titulo2 = "SERVICIOS VS. MESES"
                    subtitulo2 = "COBRANZA ACUMULADA EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 3
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 3
                    'POR UNIDADES
                    'CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes_Bs.rpt"
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Cobro_por_unidad_y_mes_Bs.rpt"
                    titulo2 = "UNIDADES VS. MESES"
                    subtitulo2 = "COBRANZA ACUMULADA EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 3
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case Else
                    '**no identificado**"
            End Select
    End Select
    'VAR_OPT1 = "1"
    'VAR_OPT2 = 0

End Sub

Private Sub rep_edif_Click()
    'POR EDIFICIOS
    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_edif.rpt"
    titulo2 = "DEPARTAMENTOS VS. MESES"
    subtitulo2 = "EDIFICIOS (Cantidad)"
    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
    CryUnidad.StoredProcParam(1) = 1
    iResult = CryUnidad.PrintReport
    If iResult <> 0 Then
        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
End Sub

Private Sub rep_eqp_Click()
    'POR EDIFICIOS
    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_eqp.rpt"
    titulo2 = "DEPARTAMENTOS VS. MESES"
    subtitulo2 = "EQUIPOS (Cantidad)"
    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
    CryUnidad.StoredProcParam(1) = 1
    iResult = CryUnidad.PrintReport
    If iResult <> 0 Then
        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
End Sub

Private Sub rep_fac_Click()
'VENTAS EN $
    Select Case VAR_OPT1
        Case 1
            'POR GESTIONES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_depto_y_mes_Bs.rpt"
                    titulo2 = "DEPARTAMENTOS VS. MESES"
                    subtitulo2 = "VENTAS ACUMULADAS EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 2
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 2
                    'POR SERVICIOS
                Case 3
                    'POR UNIDADES
                Case Else
                    '**no identificado**"
            End Select
        Case 2
            'POR MESES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS 'gr_ventas_por_depto_y_mes_Fac_Bs.rpt
                    'CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_depto_y_mes_Fac_Bs.rpt"
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Fac_por_depto_y_mes_Bs.rpt"
                    titulo2 = "DEPARTAMENTOS VS. MESES"
                    subtitulo2 = "FACTURACION ACUMULADA EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 2
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 2
                    'POR SERVICIOS
                    'CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes_Bs.rpt"
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Fac_por_proceso_y_mes_Bs.rpt"
                    titulo2 = "SERVICIOS VS. MESES"
                    subtitulo2 = "FACTURACION ACUMULADA EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 2
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 3
                    'POR UNIDADES
                    'CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes_Bs.rpt"
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Fac_por_unidad_y_mes_Bs.rpt"
                    titulo2 = "UNIDADES VS. MESES"
                    subtitulo2 = "FACTURACION ACUMULADA EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 2
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case Else
                    '**no identificado**"
            End Select
    End Select
    'VAR_OPT1 = "1"
    'VAR_OPT2 = 0

End Sub

Private Sub rep_fac2_Click()
'CANTIDAD DE CONTRATOS
    Select Case VAR_OPT1
        Case 1
            'POR GESTIONES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
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
                Case 2
                    'POR SERVICIOS
                Case 3
                    'POR UNIDADES
                Case Else
                    '**no identificado**"
            End Select
        Case 2
            'POR MESES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Fac_por_depto_y_mes.rpt"
                    titulo2 = "DEPARTAMENTOS VS. MESES"
                    subtitulo2 = "CANTIDAD ACUMULADA FACTURAS"
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 2
                    'POR SERVICIOS
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Fac_por_proceso_y_mes.rpt"
                    titulo2 = "SERVICIOS VS. MESES"
                    subtitulo2 = "CANTIDAD ACUMULADA FACTURAS"
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 3
                    'POR UNIDADES
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_Fac_por_unidad_y_mes_Bs.rpt"
                    titulo2 = "UNIDADES VS. MESES"
                    subtitulo2 = "CANTIDAD ACUMULADA FACTURAS"
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case Else
                    '**no identificado**"
            End Select
    End Select
    'VAR_OPT1 = "1"
    'VAR_OPT2 = 0

End Sub

Private Sub rep_vta_Click()
'VENTAS EN $        VAR_OPT2
    Select Case VAR_OPT1
        Case 1
            'POR GESTIONES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
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
                Case 2
                    'POR SERVICIOS
                Case 3
                    'POR UNIDADES
                Case Else
                    '**no identificado**"
            End Select
        Case 2
            'POR MESES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
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
                Case 2
                    'POR SERVICIOS
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes_Bs.rpt"
                    titulo2 = "SERVICIOS VS. MESES"
                    subtitulo2 = "VENTAS ACUMULADAS EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 3
                    'POR UNIDADES
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_unidad_y_mes_Bs.rpt"
                    titulo2 = "UNIDADES VS. MESES"
                    subtitulo2 = "VENTAS ACUMULADAS EN Bs."
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case Else
                    '**no identificado**"
            End Select
    End Select
    'VAR_OPT1 = "1"
    'VAR_OPT2 = 0
End Sub

Private Sub rep_vta2_Click()
'CANTIDAD DE CONTRATOS
    Select Case VAR_OPT1
        Case 1
            'POR GESTIONES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
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
                Case 2
                    'POR SERVICIOS
                Case 3
                    'POR UNIDADES
                Case Else
                    '**no identificado**"
            End Select
        Case 2
            'POR MESES
            Select Case VAR_OPT2
                Case 1
                    'POR DEPTOS
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_depto_y_mes.rpt"
                    titulo2 = "DEPARTAMENTOS VS. MESES"
                    subtitulo2 = "CANTIDAD ACUMULADA CONTRATOS"
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 2
                    'POR SERVICIOS
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes.rpt"
                    titulo2 = "SERVICIOS VS. MESES"
                    subtitulo2 = "CANTIDAD ACUMULADA CONTRATOS"
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case 3
                    'POR UNIDADES
                    CryUnidad.ReportFileName = App.Path & "\Reportes\Gerenciales\gr_ventas_por_proceso_y_mes_Bs.rpt"
                    titulo2 = "UNIDADES VS. MESES"
                    subtitulo2 = "CANTIDAD ACUMULADA CONTRATOS"
                    CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
                    CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
                    CryUnidad.StoredProcParam(0) = cmb_gestion_rep.Text
                    CryUnidad.StoredProcParam(1) = 1
                    iResult = CryUnidad.PrintReport
                    If iResult <> 0 Then
                        MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
                    End If
                Case Else
                    '**no identificado**"
            End Select
    End Select
    'VAR_OPT1 = "1"
    'VAR_OPT2 = 0

End Sub

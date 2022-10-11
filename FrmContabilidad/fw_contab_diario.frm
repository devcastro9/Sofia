VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_contab_diario 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos Financieros - Contabilidad - Registro Diario"
   ClientHeight    =   10170
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   15120
   Icon            =   "fw_contab_diario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13080
   ScaleWidth      =   21510
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   1935
      Left            =   9120
      TabIndex        =   165
      Top             =   4200
      Visible         =   0   'False
      Width           =   6135
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   166
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_contab_diario.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   170
            ToolTipText     =   "cOMPROBANTE cONTABLE"
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4560
            Picture         =   "fw_contab_diario.frx":12CF
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   167
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
      End
      Begin MSComCtl2.DTPicker DTP_Finicio 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   840
         TabIndex        =   168
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   85065729
         CurrentDate     =   42880
      End
      Begin MSComCtl2.DTPicker DTP_Ffin 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   3600
         TabIndex        =   169
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   85065729
         CurrentDate     =   42891
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE FIN"
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
         Left            =   3600
         TabIndex        =   172
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE INICIO"
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
         Left            =   840
         TabIndex        =   171
         Top             =   1080
         Width           =   1620
      End
   End
   Begin VB.PictureBox FrmABMDet1 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   3885
      Left            =   120
      ScaleHeight     =   3825
      ScaleWidth      =   1635
      TabIndex        =   103
      Top             =   4680
      Width           =   1695
      Begin VB.CommandButton BtnImprimir2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   120
         Picture         =   "fw_contab_diario.frx":1A91
         Style           =   1  'Graphical
         TabIndex        =   158
         ToolTipText     =   "Mayor HABER"
         Top             =   2520
         Width           =   1365
      End
      Begin VB.CommandButton BtnAñadir1 
         BackColor       =   &H80000015&
         Height          =   525
         Left            =   120
         Picture         =   "fw_contab_diario.frx":235E
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Adiciona Detalle de Registro"
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton BtnModificar1 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   120
         Picture         =   "fw_contab_diario.frx":2B1D
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Modifica Detalle de Registro"
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton BtnEliminar1 
         BackColor       =   &H80000015&
         Height          =   525
         Left            =   120
         Picture         =   "fw_contab_diario.frx":3432
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Anula  Detalle de Registro"
         Top             =   1320
         Width           =   1365
      End
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   120
         Picture         =   "fw_contab_diario.frx":3B7E
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Mayor DEBE"
         Top             =   1920
         Width           =   1365
      End
   End
   Begin VB.Frame Fram_AsientoD 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7785
      Left            =   3960
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   14895
      Begin VB.Frame Fram_AsientoH 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3345
         Left            =   0
         TabIndex        =   109
         Top             =   3720
         Visible         =   0   'False
         Width           =   14775
         Begin VB.Frame Fra_ABM1 
            BackColor       =   &H00C0C0C0&
            Height          =   975
            Left            =   120
            TabIndex        =   123
            Top             =   420
            Width           =   14535
            Begin MSDataListLib.DataCombo h_cuenta_des 
               Bindings        =   "fw_contab_diario.frx":444B
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   1800
               TabIndex        =   157
               Top             =   600
               Width           =   12135
               _ExtentX        =   21405
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "NombreCtaAux"
               BoundColumn     =   "correl"
               Text            =   "Todos"
            End
            Begin VB.TextBox Text23 
               BackColor       =   &H80000011&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   300
               Left            =   4320
               TabIndex        =   153
               Top             =   600
               Width           =   350
            End
            Begin VB.TextBox Text15 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   290
               Left            =   690
               TabIndex        =   145
               Top             =   620
               Visible         =   0   'False
               Width           =   300
            End
            Begin MSDataListLib.DataCombo H_Nombre_cmb 
               Bindings        =   "fw_contab_diario.frx":4464
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   5400
               TabIndex        =   38
               Top             =   600
               Width           =   8295
               _ExtentX        =   14631
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "NombreCta"
               BoundColumn     =   "correl"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo H_Subcta2_cmb 
               Bindings        =   "fw_contab_diario.frx":447D
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   3600
               TabIndex        =   124
               Top             =   600
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   -2147483631
               ForeColor       =   16777215
               ListField       =   "subcta2"
               BoundColumn     =   "correl"
               Text            =   "0000"
            End
            Begin MSDataListLib.DataCombo H_Subcta1_cmb 
               Bindings        =   "fw_contab_diario.frx":4496
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   2640
               TabIndex        =   125
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   -2147483631
               ForeColor       =   16777215
               ListField       =   "subcta1"
               BoundColumn     =   "correl"
               Text            =   "0000"
            End
            Begin MSDataListLib.DataCombo H_Cuenta_cmb 
               Bindings        =   "fw_contab_diario.frx":44AF
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   1800
               TabIndex        =   37
               Top             =   600
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "Cuenta"
               BoundColumn     =   "correl"
               Text            =   "0000"
            End
            Begin MSDataListLib.DataCombo H_Correl_cmb 
               Bindings        =   "fw_contab_diario.frx":44C8
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   120
               TabIndex        =   126
               Top             =   600
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "correl"
               BoundColumn     =   "correl"
               Text            =   "0000"
            End
            Begin VB.Label lbl_cuenta1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cuenta"
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
               TabIndex        =   131
               Top             =   240
               Width           =   630
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Sub Cta 1"
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
               Left            =   2640
               TabIndex        =   130
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Sub Cta 2"
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
               Left            =   3600
               TabIndex        =   129
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Correlativo"
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
               Left            =   120
               TabIndex        =   128
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblNombre1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nombre de la Cuenta"
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
               Left            =   5400
               TabIndex        =   127
               Top             =   240
               Width           =   1905
            End
         End
         Begin VB.Frame Fra_Aux1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipos de Auxiliares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1815
            Left            =   120
            TabIndex        =   110
            Top             =   1440
            Width           =   14535
            Begin VB.TextBox Text21 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   280
               Left            =   2420
               TabIndex        =   151
               Top             =   980
               Width           =   350
            End
            Begin VB.TextBox Text20 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   290
               Left            =   1080
               TabIndex        =   150
               Top             =   610
               Width           =   270
            End
            Begin VB.TextBox Text19 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   290
               Left            =   5880
               TabIndex        =   149
               Top             =   610
               Width           =   270
            End
            Begin VB.TextBox Text18 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   280
               Left            =   7220
               TabIndex        =   148
               Top             =   980
               Width           =   350
            End
            Begin VB.TextBox Text17 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   290
               Left            =   10800
               TabIndex        =   147
               Top             =   610
               Width           =   270
            End
            Begin VB.TextBox Text16 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   280
               Left            =   12140
               TabIndex        =   146
               Top             =   980
               Width           =   350
            End
            Begin VB.PictureBox Buscar4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   2880
               Picture         =   "fw_contab_diario.frx":44E1
               ScaleHeight     =   615
               ScaleWidth      =   1215
               TabIndex        =   78
               Top             =   720
               Width           =   1215
            End
            Begin VB.PictureBox Buscar5 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   7800
               Picture         =   "fw_contab_diario.frx":4C96
               ScaleHeight     =   615
               ScaleWidth      =   1215
               TabIndex        =   80
               Top             =   720
               Width           =   1215
            End
            Begin VB.PictureBox Buscar6 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   12600
               Picture         =   "fw_contab_diario.frx":544B
               ScaleHeight     =   615
               ScaleWidth      =   1215
               TabIndex        =   82
               Top             =   720
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo H_Cta_Aux1_cmb 
               Bindings        =   "fw_contab_diario.frx":5C00
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   120
               TabIndex        =   111
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "Aux1"
               BoundColumn     =   "correl"
               Text            =   "0000"
            End
            Begin MSDataListLib.DataCombo H_Cta_Aux2_cmb 
               Bindings        =   "fw_contab_diario.frx":5C19
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   4920
               TabIndex        =   112
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "Aux2"
               BoundColumn     =   "correl"
               Text            =   "0000"
            End
            Begin MSDataListLib.DataCombo H_Cta_Aux3_cmb 
               Bindings        =   "fw_contab_diario.frx":5C32
               DataField       =   "H_Correl"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   9840
               TabIndex        =   113
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "Aux3"
               BoundColumn     =   "correl"
               Text            =   "0000"
            End
            Begin MSDataListLib.DataCombo dtc_desc11 
               Bindings        =   "fw_contab_diario.frx":5C4B
               DataField       =   "H_Cta_Aux1"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   120
               TabIndex        =   39
               Top             =   1320
               Visible         =   0   'False
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "desc1"
               BoundColumn     =   "codigo1"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc12 
               Bindings        =   "fw_contab_diario.frx":5C65
               DataField       =   "H_Cta_Aux2"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   4920
               TabIndex        =   40
               Top             =   1320
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "desc2"
               BoundColumn     =   "codigo2"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc13 
               Bindings        =   "fw_contab_diario.frx":5C7F
               DataField       =   "H_Cta_Aux3"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   9840
               TabIndex        =   41
               Top             =   1320
               Visible         =   0   'False
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "desc3"
               BoundColumn     =   "codigo3"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo11 
               Bindings        =   "fw_contab_diario.frx":5C99
               DataField       =   "H_Cta_Aux1"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   120
               TabIndex        =   79
               Top             =   960
               Visible         =   0   'False
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "codigo1"
               BoundColumn     =   "codigo1"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo12 
               Bindings        =   "fw_contab_diario.frx":5CB3
               DataField       =   "H_Cta_Aux2"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   4920
               TabIndex        =   81
               Top             =   960
               Visible         =   0   'False
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "codigo2"
               BoundColumn     =   "codigo2"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo13 
               Bindings        =   "fw_contab_diario.frx":5CCD
               DataField       =   "H_Cta_Aux3"
               DataSource      =   "Ado_detalle1"
               Height          =   315
               Left            =   9840
               TabIndex        =   83
               Top             =   960
               Visible         =   0   'False
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               BackColor       =   12632256
               ForeColor       =   0
               ListField       =   "codigo3"
               BoundColumn     =   "codigo3"
               Text            =   "Todos"
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Auxiliar - 3"
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
               Left            =   11520
               TabIndex        =   122
               Top             =   600
               Width           =   900
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Auxiliar - 2"
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
               Left            =   6600
               TabIndex        =   121
               Top             =   600
               Width           =   900
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Auxiliar - 1"
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
               Index           =   0
               Left            =   1800
               TabIndex        =   120
               Top             =   600
               Width           =   900
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               DataField       =   "Denominacion_Aux1"
               DataSource      =   "Ado_datos"
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
               Height          =   255
               Left            =   120
               TabIndex        =   119
               Top             =   1320
               Width           =   3975
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               DataField       =   "Cod_Aux2"
               DataSource      =   "Ado_datos"
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
               Height          =   255
               Left            =   4920
               TabIndex        =   118
               Top             =   960
               Width           =   2655
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               DataField       =   "Denominacion_Aux2"
               DataSource      =   "Ado_datos"
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
               Height          =   255
               Left            =   4920
               TabIndex        =   117
               Top             =   1320
               Width           =   4095
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               DataField       =   "Cod_Aux3"
               DataSource      =   "Ado_datos"
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
               Height          =   255
               Left            =   9840
               TabIndex        =   116
               Top             =   960
               Width           =   2655
            End
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               DataField       =   "Denominacion_Aux3"
               DataSource      =   "Ado_datos"
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
               Height          =   255
               Left            =   9840
               TabIndex        =   115
               Top             =   1320
               Width           =   3975
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               DataField       =   "Cod_Aux1"
               DataSource      =   "Ado_datos"
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
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   960
               Width           =   2655
            End
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Caption         =   "HABER"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   132
            Top             =   120
            Width           =   14535
         End
      End
      Begin VB.ComboBox cmb_moneda 
         DataField       =   "tipo_moneda"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         ItemData        =   "fw_contab_diario.frx":5CE7
         Left            =   1320
         List            =   "fw_contab_diario.frx":5CF1
         TabIndex        =   33
         Text            =   "BOB"
         Top             =   3240
         Width           =   1515
      End
      Begin VB.TextBox D_Cambio_cmb 
         DataField       =   "D_Cambio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   5400
         TabIndex        =   34
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox D_MontoDl_cmb 
         DataField       =   "D_MontoDl"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   12000
         TabIndex        =   36
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox D_MontoBs_cmb 
         DataField       =   "D_MontoBs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   8640
         TabIndex        =   35
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   14640
         TabIndex        =   96
         Top             =   6960
         Width           =   14640
         Begin VB.PictureBox BtnCancelar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7155
            Picture         =   "fw_contab_diario.frx":5CFF
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   98
            Top             =   0
            Width           =   1455
         End
         Begin VB.PictureBox BtnGrabar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5880
            Picture         =   "fw_contab_diario.frx":65EB
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   97
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   13680
            TabIndex        =   99
            Top             =   195
            Width           =   75
         End
      End
      Begin VB.Frame Fra_Aux 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipos de Auxiliares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1695
         Left            =   120
         TabIndex        =   73
         Top             =   1440
         Width           =   14655
         Begin VB.TextBox Text14 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   5880
            TabIndex        =   144
            Top             =   610
            Width           =   270
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   280
            Left            =   7220
            TabIndex        =   143
            Top             =   980
            Width           =   345
         End
         Begin VB.TextBox Text12 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   10800
            TabIndex        =   142
            Top             =   610
            Width           =   270
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   280
            Left            =   12140
            TabIndex        =   141
            Top             =   980
            Width           =   345
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   1080
            TabIndex        =   140
            Top             =   610
            Width           =   270
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   280
            Left            =   2410
            TabIndex        =   139
            Top             =   980
            Width           =   345
         End
         Begin VB.PictureBox Buscar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "fw_contab_diario.frx":6DC1
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   65
            Top             =   700
            Width           =   1215
         End
         Begin VB.PictureBox Buscar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7800
            Picture         =   "fw_contab_diario.frx":7576
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   74
            Top             =   700
            Width           =   1215
         End
         Begin VB.PictureBox Buscar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   12600
            Picture         =   "fw_contab_diario.frx":7D2B
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   76
            Top             =   700
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo D_Cta_Aux1_cmb 
            Bindings        =   "fw_contab_diario.frx":84E0
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "Aux1"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo D_Cta_Aux2_cmb 
            Bindings        =   "fw_contab_diario.frx":84F9
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   4920
            TabIndex        =   85
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "Aux2"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo D_Cta_Aux3_cmb 
            Bindings        =   "fw_contab_diario.frx":8512
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   9840
            TabIndex        =   86
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "Aux3"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "fw_contab_diario.frx":852B
            DataField       =   "D_Cta_Aux1"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   1320
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "desc1"
            BoundColumn     =   "codigo1"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc9 
            Bindings        =   "fw_contab_diario.frx":8544
            DataField       =   "D_Cta_Aux2"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   4920
            TabIndex        =   31
            Top             =   1320
            Visible         =   0   'False
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "desc2"
            BoundColumn     =   "codigo2"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc10 
            Bindings        =   "fw_contab_diario.frx":855D
            DataField       =   "D_Des_Aux3"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   9840
            TabIndex        =   32
            Top             =   1320
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "desc3"
            BoundColumn     =   "codigo3"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "fw_contab_diario.frx":8577
            DataField       =   "D_Cta_Aux1"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "codigo1"
            BoundColumn     =   "codigo1"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo9 
            Bindings        =   "fw_contab_diario.frx":8590
            DataField       =   "D_Cta_Aux2"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   4920
            TabIndex        =   75
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "codigo2"
            BoundColumn     =   "codigo2"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo10 
            Bindings        =   "fw_contab_diario.frx":85A9
            DataField       =   "D_Cta_Aux3"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   9840
            TabIndex        =   77
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "codigo3"
            BoundColumn     =   "codigo3"
            Text            =   "Todos"
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 3"
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
            Left            =   11520
            TabIndex        =   95
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 2"
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
            Left            =   6600
            TabIndex        =   94
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 1"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   93
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Txt_campo6 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "Denominacion_Aux1"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Label Txt_campo7 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "Cod_Aux2"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   4920
            TabIndex        =   91
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Txt_campo8 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "Denominacion_Aux2"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   4920
            TabIndex        =   90
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label Txt_campo9 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "Cod_Aux3"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   9840
            TabIndex        =   89
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Txt_campo10 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "Denominacion_Aux3"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   9840
            TabIndex        =   88
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Label Txt_campo5 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "Cod_Aux1"
            DataSource      =   "Ado_datos"
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
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   960
            Width           =   2655
         End
      End
      Begin VB.Frame Fra_ABM2 
         BackColor       =   &H00C0C0C0&
         Height          =   975
         Left            =   120
         TabIndex        =   64
         Top             =   420
         Width           =   14655
         Begin MSDataListLib.DataCombo d_cuenta_des 
            Bindings        =   "fw_contab_diario.frx":85C3
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   1680
            TabIndex        =   156
            Top             =   600
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "NombreCtaAux"
            BoundColumn     =   "correl"
            Text            =   "Todos"
         End
         Begin VB.TextBox Text22 
            BackColor       =   &H80000011&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   300
            Left            =   4320
            TabIndex        =   152
            Top             =   600
            Width           =   350
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   730
            TabIndex        =   138
            Top             =   610
            Visible         =   0   'False
            Width           =   250
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "estado_codigo"
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
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   136
            Text            =   "REG"
            Top             =   180
            Width           =   1125
         End
         Begin MSDataListLib.DataCombo D_Nombre_cmb 
            Bindings        =   "fw_contab_diario.frx":85DC
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   5400
            TabIndex        =   29
            Top             =   600
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "NombreCta"
            BoundColumn     =   "correl"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo D_Subcta2_cmb 
            Bindings        =   "fw_contab_diario.frx":85F5
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   3600
            TabIndex        =   67
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "subcta2"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo D_Subcta1_cmb 
            Bindings        =   "fw_contab_diario.frx":860E
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   2640
            TabIndex        =   68
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483631
            ForeColor       =   16777215
            ListField       =   "subcta1"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo D_Correl_cmb 
            Bindings        =   "fw_contab_diario.frx":8627
            DataField       =   "D_Correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   120
            TabIndex        =   69
            Top             =   600
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "correl"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo D_Cuenta_cmb 
            Bindings        =   "fw_contab_diario.frx":8640
            DataField       =   "D_correl"
            DataSource      =   "Ado_detalle1"
            Height          =   315
            Left            =   1680
            TabIndex        =   28
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "Cuenta"
            BoundColumn     =   "correl"
            Text            =   "0000"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estado"
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
            Left            =   12720
            TabIndex        =   135
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl_cuenta 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cuenta "
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
            Left            =   1680
            TabIndex        =   72
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Correlativo"
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
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblNombreC 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nombre de la Cuenta"
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
            Left            =   3120
            TabIndex        =   70
            Top             =   240
            Width           =   1905
         End
      End
      Begin VB.Label lblMonto 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Bs.:"
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
         Left            =   7560
         TabIndex        =   108
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cambio:"
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
         Left            =   3720
         TabIndex        =   102
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblMontoDls 
         BackStyle       =   0  'Transparent
         Caption         =   "MontoDls.:"
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
         Left            =   10920
         TabIndex        =   101
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblMoneda 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
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
         TabIndex        =   100
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "DEBE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   14655
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   53
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "fw_contab_diario.frx":8659
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   62
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
         Left            =   5520
         Picture         =   "fw_contab_diario.frx":8E1B
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   61
         ToolTipText     =   "cOMPROBANTE cONTABLE"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "fw_contab_diario.frx":96E8
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   60
         ToolTipText     =   "Buscar Registros"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "fw_contab_diario.frx":9E9D
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   59
         ToolTipText     =   "Aprueba el Registro Seleccionado"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "fw_contab_diario.frx":A6D0
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   58
         ToolTipText     =   "Anula el Registro Seleccionado"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1305
         Picture         =   "fw_contab_diario.frx":AE1C
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   57
         ToolTipText     =   "Modifica el Registro Seleccionado"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "fw_contab_diario.frx":B731
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   56
         ToolTipText     =   "Adiciona un Nuevo Registro"
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "fw_contab_diario.frx":BEF0
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "fw_contab_diario.frx":C332
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROYECTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   12990
         TabIndex        =   63
         Top             =   195
         Width           =   1545
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   676
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   20280
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "fw_contab_diario.frx":C53C
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   51
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "fw_contab_diario.frx":CD12
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   50
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROYECTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   12945
         TabIndex        =   52
         Top             =   195
         Width           =   1545
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "HABER (CREDITO)"
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   1860
      TabIndex        =   47
      Top             =   6600
      Width           =   17100
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "fw_contab_diario.frx":D5FE
         Height          =   1575
         Left            =   150
         TabIndex        =   48
         Top             =   240
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   2778
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "Cod_Comp_Detalle"
            Caption         =   "Cod.Det"
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
            DataField       =   "H_Cuenta"
            Caption         =   "H_Cuenta"
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
            DataField       =   "H_Nombre"
            Caption         =   "H_Nombre"
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
            DataField       =   "H_SubCta1"
            Caption         =   "H_SubCta1"
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
            DataField       =   "H_SubCta2"
            Caption         =   "H_SubCta2"
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
            DataField       =   "H_Cta_Aux1"
            Caption         =   "H_Cod_Aux1"
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
         BeginProperty Column06 
            DataField       =   "H_Cta_Aux2"
            Caption         =   "H_Cod_Aux2"
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
            DataField       =   "H_Cta_Aux3"
            Caption         =   "H_Cod_Aux3"
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
            DataField       =   "H_MontoBs"
            Caption         =   "H_MontoBs"
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
            DataField       =   "H_MontoDl"
            Caption         =   "H_MontoDl"
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
         BeginProperty Column10 
            DataField       =   "H_Cambio"
            Caption         =   "H_Cambio"
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
         BeginProperty Column11 
            DataField       =   "H_Des_Aux1"
            Caption         =   "H_Des_Aux1"
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
            DataField       =   "H_Des_Aux2"
            Caption         =   "H_Des_Aux2"
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
            DataField       =   "H_Des_Aux3"
            Caption         =   "H_Des_Aux3"
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
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   4635.213
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1874.835
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1980.284
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DEBE"
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   1860
      TabIndex        =   45
      Top             =   4560
      Width           =   17100
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "fw_contab_diario.frx":D619
         Height          =   1575
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   2778
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "Cod_Comp_Detalle"
            Caption         =   "Cod.Det"
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
            DataField       =   "D_Cuenta"
            Caption         =   "D_Cuenta"
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
            DataField       =   "D_Nombre"
            Caption         =   "D_Nombre_Cuenta"
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
            DataField       =   "D_Subcta1"
            Caption         =   "D_Subcta1"
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
            DataField       =   "D_SubCta2"
            Caption         =   "D_SubCta2"
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
            DataField       =   "D_Cta_Aux1"
            Caption         =   "D_Cod_Aux1"
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
         BeginProperty Column06 
            DataField       =   "D_Cta_Aux2"
            Caption         =   "D_Cod_Aux2"
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
            DataField       =   "D_Cta_Aux3"
            Caption         =   "D_Cod_Aux3"
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
            DataField       =   "D_MontoBs"
            Caption         =   "D_MontoBs"
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
            DataField       =   "D_MontoDl"
            Caption         =   "D_MontoDl"
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
         BeginProperty Column10 
            DataField       =   "D_Cambio"
            Caption         =   "D_Cambio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
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
         BeginProperty Column12 
            DataField       =   "D_Des_Aux1"
            Caption         =   "D_Des_Aux1"
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
            DataField       =   "D_Des_Aux2"
            Caption         =   "D_Des_Aux2"
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
            DataField       =   "D_Des_Aux3"
            Caption         =   "D_Des_Aux3"
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
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   4710.047
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   2025.071
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1995.024
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   3915
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   6375
      Begin VB.OptionButton OptTodos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   3615
         Width           =   1095
      End
      Begin VB.OptionButton OptSinAprobar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pendientes"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   3615
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "fw_contab_diario.frx":D634
         Height          =   3315
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   5847
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "doc_codigo"
            Caption         =   "Tipo"
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
            DataField       =   "Fecha_transacion"
            Caption         =   "Fecha"
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
            DataField       =   "doc_numero"
            Caption         =   "Comprobante"
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
         BeginProperty Column04 
            DataField       =   "Glosa"
            Caption         =   "Glosa"
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
         BeginProperty Column05 
            DataField       =   "unidad_codigo"
            Caption         =   "U.Ejecutora"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Trámite"
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
            DataField       =   "Cod_Comp"
            Caption         =   "Correlativo"
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
         BeginProperty Column08 
            DataField       =   "Tipo_Comp"
            Caption         =   "Tipo.Transaccion."
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
            DataField       =   "cod_trans"
            Caption         =   "Corelativo.Origen"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cod.Beneficiario"
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
            DataField       =   "mes_trasaccion"
            Caption         =   "MES"
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
            DataField       =   "ges_gestion"
            Caption         =   "GESTION"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2954.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   60
         Top             =   3555
         Width           =   6240
         _ExtentX        =   11007
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
         BackColor       =   16777215
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
   End
   Begin MSAdodcLib.Adodc AdoConvenio 
      Height          =   330
      Left            =   1200
      Top             =   9720
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
      Caption         =   "AdoConvenio"
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
   Begin Crystal.CrystalReport CryRepGrid 
      Left            =   240
      Top             =   9720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryComp_Manual 
      Left            =   720
      Top             =   9720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc AdodCtaBancaria 
      Height          =   330
      Left            =   2400
      Top             =   9360
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
      Caption         =   "AdodCtaBancaria"
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
   Begin MSAdodcLib.Adodc Adodcbeneficiario 
      Height          =   330
      Left            =   11640
      Top             =   9360
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Adodcbeneficiario"
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
   Begin MSAdodcLib.Adodc Adodcdocumento 
      Height          =   330
      Left            =   9360
      Top             =   9360
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Adodcdocumento"
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
   Begin VB.Frame FraGlobal 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   3915
      Left            =   6560
      TabIndex        =   0
      Top             =   660
      Width           =   12405
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   9580
         TabIndex        =   164
         Top             =   380
         Width           =   300
      End
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "fw_contab_diario.frx":D64C
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7320
         TabIndex        =   42
         Top             =   2520
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "etapa_descripcion"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "fw_contab_diario.frx":D665
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   23
         Top             =   2520
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox txt_mes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "mes_trasaccion"
         DataSource      =   "Ado_datos"
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
         Left            =   2240
         Locked          =   -1  'True
         TabIndex        =   162
         Top             =   360
         Width           =   2085
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   9075
         TabIndex        =   155
         Top             =   975
         Width           =   320
      End
      Begin VB.TextBox txtcodsolicitud 
         DataField       =   "solicitud_codigo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   10920
         TabIndex        =   154
         Text            =   "-"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "estado_codigo"
         DataSource      =   "Ado_datos"
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   134
         Text            =   "REG"
         Top             =   1560
         Width           =   1245
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "fw_contab_diario.frx":D67E
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   11880
         TabIndex        =   44
         Top             =   2520
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "cod_trans"
         DataSource      =   "Ado_datos"
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
         Left            =   7155
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox txt_ges 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ges_gestion"
         DataSource      =   "Ado_datos"
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
         Left            =   860
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox TxtComprobante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "Cod_Comp"
         DataSource      =   "Ado_datos"
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox Txt_glosa 
         DataField       =   "Glosa"
         DataSource      =   "Ado_datos"
         Height          =   510
         Left            =   795
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   3120
         Width           =   11460
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "fw_contab_diario.frx":D697
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   960
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "fw_contab_diario.frx":D6B0
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboNomTipo 
         Bindings        =   "fw_contab_diario.frx":D6C9
         DataField       =   "Tipo_Comp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7320
         TabIndex        =   24
         Top             =   2820
         Visible         =   0   'False
         Width           =   4970
         _ExtentX        =   8758
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Denominacion_Tipo"
         BoundColumn     =   "Codigo_Tipo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo CboTipo 
         Bindings        =   "fw_contab_diario.frx":D6E2
         DataField       =   "Tipo_Comp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6720
         TabIndex        =   25
         Top             =   2520
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ListField       =   "Codigo_Tipo"
         BoundColumn     =   "Codigo_Tipo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "fw_contab_diario.frx":D6FB
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   11400
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "etapa_codigo"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha_trans 
         DataField       =   "Fecha_transacion"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   10560
         TabIndex        =   27
         Top             =   1560
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85065729
         CurrentDate     =   41678
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_desc14 
         Bindings        =   "fw_contab_diario.frx":D714
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5760
         TabIndex        =   159
         Top             =   360
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo txt_codigo1 
         Bindings        =   "fw_contab_diario.frx":D72E
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5040
         TabIndex        =   160
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   ""
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo:"
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
         Left            =   2400
         TabIndex        =   163
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Mes:"
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
         Left            =   1800
         TabIndex        =   161
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lbl_etapa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Etapa Proceso ISO"
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
         Left            =   7320
         TabIndex        =   137
         Top             =   2235
         Width           =   1725
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Left            =   120
         TabIndex        =   133
         Top             =   1560
         Width           =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   0
         X2              =   12360
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Correlativo Origen:"
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
         Left            =   5040
         TabIndex        =   22
         Top             =   1560
         Width           =   2085
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "doc_numero"
         DataSource      =   "Ado_datos"
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
         Height          =   300
         Left            =   11385
         TabIndex        =   20
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lbl_tramite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Trámite:"
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
         Left            =   9720
         TabIndex        =   17
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidad Ejecutora:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lbl_tipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   4560
         TabIndex        =   6
         Top             =   405
         Width           =   465
      End
      Begin VB.Label Label_Fecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro:"
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
         Left            =   8835
         TabIndex        =   5
         Top             =   1560
         Width           =   1710
      End
      Begin VB.Label lbl_glosa 
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   3255
         Width           =   630
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lbl_beneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiario:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   2220
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprobante:"
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
         Left            =   10080
         TabIndex        =   1
         Top             =   360
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2400
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   4680
      Top             =   9000
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   7080
      Top             =   9000
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
      Left            =   9360
      Top             =   9000
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
      Left            =   11640
      Top             =   9000
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   120
      Top             =   9360
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "AdoCaja"
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   4680
      Top             =   9360
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
      Caption         =   "Ado_detalle1"
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   7080
      Top             =   9360
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
      Caption         =   "Ado_detalle2"
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
      Left            =   3480
      Top             =   9720
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   5880
      Top             =   9720
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   8280
      Top             =   9720
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   10800
      Top             =   9720
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
      Caption         =   "Ado_datos10"
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
      Left            =   1080
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_datos12 
      Height          =   330
      Left            =   3480
      Top             =   10200
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
      Caption         =   "Ado_datos12"
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
      Left            =   5880
      Top             =   10200
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
   Begin Crystal.CrystalReport CryDebe 
      Left            =   240
      Top             =   10200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryHaber 
      Left            =   720
      Top             =   10200
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
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos14 
      Height          =   330
      Left            =   120
      Top             =   8640
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
      Caption         =   "Ado_datos14"
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
   Begin VB.Menu mnumenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAnulacion 
         Caption         =   "Anulación"
      End
      Begin VB.Menu mnuReversion 
         Caption         =   "Reversión"
      End
      Begin VB.Menu mnuDevolucion 
         Caption         =   "Devolución"
      End
   End
End
Attribute VB_Name = "fw_contab_diario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---variables para determinar el estado del comprobante contable en la tabla pagos
Public estadoconta As String
Public estadopago As String
'---
Dim MontoAnterior As Double
Dim Gdenomcaja As String
'--
Public salir As Integer
'---
Public num_comprobante As Integer ' vaiable donde se almacena nùmero de comprobante
Public MovCuenta As String  'variable para el tipo de cuenta ("T" título, "D" detalle

'********RECORDSETS
Dim rs_datos As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

Dim rs_aux1 As ADODB.Recordset
'Dim rsdocumento As ADODB.Recordset
'Dim rsOrganismo As ADODB.Recordset
'Dim rsbenef_traspaso As ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset

Dim rstipocomp As ADODB.Recordset
''Dim rscaja As ADODB.Recordset
'Dim rspco As ADODB.Recordset  'Movimientos de PCO

Dim rs_detalle1 As New ADODB.Recordset
Dim rs_detalle2 As New ADODB.Recordset

'----
Dim adiciona As String
Dim VAR_BUS As String

'Public CAMcorrel As String
'Dim lcta As String
'---
'*******************
'Dim daux1 As String
'Dim daux2 As String
'Dim daux3 As String
'Dim haux1 As String
'Dim haux2 As String
'Dim haux3 As String
'Dim dctalarga As String
'Dim dctaaux2 As String
'Dim dctaaux3 As String
'Dim hctalarga As String
'Dim hctaaux2 As String
'Dim hctaaux3 As String
'----------
'Dim DebeAuxiliar As String
'Dim haberAuxiliar As String
Dim VAR_CITE As String
'****
'Dim aprobacion() As Integer
'Dim CTipoC As Double  'tipo de cambio
'Dim CFecha  As Date   'fecha actual
'Dim CmonedaBs As String
'Dim CmonedaSus As String
'Dim Ctipomoneda As String
Dim cmodificar As String

Dim VAR_TABLA, VAR_CODIGO, VAR_DES As String
Dim VAR_SW As String
Dim VAR_SW2 As String
Dim VAR_VAL As String
Dim VAR_SUB1 As String
Dim VAR_AUX1, VAR_AUX2, VAR_AUX3 As String
Dim VAR_TIT As String
Dim VAR_CTA, VAR_SUB2 As String

Dim Monto As Double

'Dim cmoney As String  ''Bs' para Bs y 'Sus' para sus
'Public Cdenominacion As String
'Public cdenomctabancaria As String
'Public denomorgan As String
'Public orgo As String
'Public sw1 As Integer
'Public sw2 As Integer

'Para B{usqueda
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'Private Sub cboDCodOrg_Click()
'  rsorganismo.Filter = adFilterNone
'  rsorganismo.Filter = "org_codigo='" & Trim(Me.cboDCodOrg) & "'"
'  If rsorganismo.RecordCount <> 0 Then
'    Me.cboDDenomOrg.Text = Trim(rsorganismo!descripcion)
'  Else
'    Exit Sub
'  End If
'  dctalarga = Trim(cboDCodOrg.Text)
'
'End Sub
'
'Private Sub CboDCta_Click()
'
'  Me.CboDSubcta1.Clear
'  Me.CboDSubcta2.Clear
'  rsplanctas.MoveFirst
'  rsplanctas.Find "cuenta=" & "'" & Trim(CboDCta.Text) & "'"
'  If rscuentas.State = adStateOpen Then rscuentas.Close
'  rscuentas.Open "SELECT SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.CboDCta.Text) & "')", db, adOpenKeyset, adLockReadOnly
'  'MsgBox rscuentas.RecordCount
'  Do While Not rscuentas.EOF
'    Me.CboDSubcta1.AddItem rscuentas!subcta1
'    rscuentas.MoveNext
'  Loop
'  If rscuentas.RecordCount = 0 Then
'  Me.CboDSubcta1.AddItem "00"
'  End If
'  'Me.CboDSubcta1.Text = Me.CboDSubcta1.List(0)
'End Sub
'
'Private Sub CboDCta_KeyPress(KeyAscii As Integer)
'  'KeyAscii = 0
'End Sub
'
'Private Sub cboDctaaux1_Click()
'    'On Error GoTo error6
'    'rscta_corrienteDebe.MoveFirst
'    rscta_corrienteDebe.Filter = adFilterNone
'    'rscta_corrienteDebe.Find "cta_codigo='" & Trim(Me.cboDctaaux1) & "'"
'    rscta_corrienteDebe.Filter = "cta_codigo='" & Trim(Me.cboDctaaux1) & "'"
'    If rscta_corrienteDebe.RecordCount <> 0 Then
'      Me.cboDctanomaux1.Text = Trim(rscta_corrienteDebe!cta_descripcion)
'    Else
'      Exit Sub
'    End If
'    dctalarga = Trim(cboDctaaux1)
'    Exit Sub
'error6:
'    If Err.Number = 28 Then
'        Exit Sub
'    End If
'End Sub
'
'Private Sub CboDCtaCAM_Click()
''comprobante contable  de diferencias cambiarias
'  Me.CboDSub1CAM.Clear
'  Me.CboDSub2CAM.Clear
'  rsplanctas.MoveFirst
'  rsplanctas.Find "cuenta=" & "'" & Trim(Me.CboDCtaCAM.Text) & "'"
'  If rscuentas.State = adStateOpen Then rscuentas.Close
'  rscuentas.Open "SELECT SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.CboDCtaCAM.Text) & "')", db, adOpenKeyset, adLockReadOnly
'  'MsgBox rscuentas.RecordCount
'  Do While Not rscuentas.EOF
'    Me.CboDSub1CAM.AddItem rscuentas!subcta1
'    rscuentas.MoveNext
'  Loop
'  If Me.CboDCtaCAM.Text = "1111" Then
'      Me.CboDSub1CAM.Clear
'      Me.CboDSub1CAM.AddItem "02"
'  End If
'  If rscuentas.RecordCount = 0 Then
'  Me.CboDSub1CAM.AddItem "00"
'  End If
'  Select Case Trim(CboDCtaCAM.Text)
'    Case "1111"
'      CboHCtaCAM.Clear
'      CboHCtaCAM.AddItem "5174"
'      'CboHCtaCAM.Text = CboHCtaCAM.List(0)
'      'CboHCtaCAM.Locked = True
'    Case "6141"
'      CboHCtaCAM.Clear
'      CboHCtaCAM.AddItem "1111"
'      'CboHCtaCAM.Text = CboHCtaCAM.List(0)
'      'CboHCtaCAM.Locked = True
'  End Select
'  'CboDSub1CAM.Text = CboDSub1CAM.List(0)
'End Sub
'Private Sub cboDctanomaux1_Click()
'    On Error GoTo err1
'    rscta_corrienteDebe.MoveFirst
'    rscta_corrienteDebe.Find "cta_descripcion='" & Trim(Me.cboDctanomaux1) & "'"
'    cboDctaaux1.Text = rscta_corrienteDebe!Cta_Codigo
'    dctalarga = Trim(cboDctaaux1)
'err1:
'    If Err.Number = 28 Then
'    Exit Sub
'    End If
'End Sub
'
'Private Sub cboDDenomOrg_Click()
'On Error GoTo err1
'    rsorganismo.Filter = adFilterNone
'    rsorganismo.MoveFirst
'    rsorganismo.Find "descripcion='" & Trim(cboDDenomOrg) & "'"
'    cboDCodOrg = rsorganismo!org_codigo
'    dctalarga = Trim(cboDCodOrg)
'err1:
'    If Err.Number = 28 Then
'    Exit Sub
'    End If
'End Sub
'
'Private Sub CboDSub1CAM_Click()
' Dim i As Integer
' On Error GoTo Laberror1
'    Me.CboDSub2CAM.Clear
'      If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
'      rssubcuenta.Open "SELECT SubCta2,Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.CboDCtaCAM.Text) & "') AND (SubCta1 ='" & Trim(Me.CboDSub1CAM.Text) & "')", db, adOpenKeyset, adLockReadOnly
'      If rssubcuenta.RecordCount = 0 Then
'        Me.CboDSub2CAM.AddItem "00"
'        'Me.CboDSubcta2.Text = "00"
'      Else
'        rssubcuenta.MoveFirst
'        Do While Not rssubcuenta.EOF
'           Me.CboDSub2CAM.AddItem rssubcuenta!subcta2
'           rssubcuenta.MoveNext
'        Loop
'      End If
'      If Me.CboDCtaCAM.Text = "1111" Then
'        For i = 0 To Me.CboDSub2CAM.ListCount
'          If Me.CboDSub2CAM.List(i) = "00" Then
'             Me.CboDSub2CAM.RemoveItem (i)
'          End If
'        Next
'      End If
'   ' Me.CboDSubcta2.Text = Me.CboDSubcta2.List(0)
'   'CboDSub2CAM.Text = CboDSub2CAM.List(0)
'Exit Sub
'Laberror1:
'If Err.Number = 3021 Then
' MsgBox "Elija una cuenta", vbExclamation + vbDefaultButton1
' Me.CboDCtaCAM.SetFocus
' 'Me.CboDCta.SetFocus
'End If
'End Sub
'
'Private Sub CboDSub2CAM_Change()
'Dim sql_cuenta As String
''    Call Titulo(Me.CboDCtaCAM, Me.CboDSub1CAM, Me.CboDSub2CAM)
''    If lcta = "N" Then
'        Exit Sub
''    End If
''    If lcta = "S" Then
'        If MovCuenta = "T" Then
'            Exit Sub
''            Me.CboDCta.SetFocus
'        End If
'        If MovCuenta = "D" Then
'            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
            'sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(Me.CboDCta) & "' and subcta1='" & Trim(Me.CboDSubcta1) & "' and subcta2='" & Trim(Me.CboDSubcta2) & "'"
'            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(CboDCtaCAM) & "' and subcta1='" & Trim(CboDSub1CAM) & "' and subcta2='" & Trim(Me.CboDSub2CAM) & "'"
'            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
'            daux1 = Trim(rsPlan_cuentas!aux1)
'            daux2 = Trim(rsPlan_cuentas!AUX2)
'            daux3 = Trim(rsPlan_cuentas!aux3)
            '---habilitacion de auxiliares---
'            If rsPlan_cuentas!aux1 <> "00" Then
'              SSTabDebe.TabEnabled(0) = True
'            Else
''              SSTabDebe.TabEnabled(0) = False
'            End If
'            If rsPlan_cuentas!AUX2 <> "00" Then
'              SSTabDebe.TabEnabled(1) = True
'            Else
'              SSTabDebe.TabEnabled(1) = False
'            End If
'            If rsPlan_cuentas!aux3 <> "00" Then
'                SSTabDebe.TabEnabled(2) = True
'            Else
'              SSTabDebe.TabEnabled(2) = False
'            End If
'            auxDebe daux1
'            auxDebe daux2
'            auxDebe daux3
'            SSTabDebe_Click (0)
'        End If
'    End If
'End Sub

'Private Sub CboDSub2CAM_Click()
''*******
'    Dim sql_cuenta As String
'    CboDCta.Text = ""
'
'    Call Titulo(Me.CboDCtaCAM, Me.CboDSub1CAM, Me.CboDSub2CAM)
'    If lcta = "N" Then
'        Exit Sub
'    End If
'    If lcta = "S" Then
'        If MovCuenta = "T" Then
'            Exit Sub
'            Me.CboDCta.SetFocus
'        End If
'        If MovCuenta = "D" Then
'            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
'            'sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(Me.CboDCta) & "' and subcta1='" & Trim(Me.CboDSubcta1) & "' and subcta2='" & Trim(Me.CboDSubcta2) & "'"
'            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(CboDCtaCAM) & "' and subcta1='" & Trim(CboDSub1CAM) & "' and subcta2='" & Trim(Me.CboDSub2CAM) & "'"
'            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
'            daux1 = Trim(rsPlan_cuentas!aux1)
'            daux2 = Trim(rsPlan_cuentas!AUX2)
'            daux3 = Trim(rsPlan_cuentas!aux3)
'            '---habilitacion de auxiliares---
'            If rsPlan_cuentas!aux1 <> "00" Then
''              SSTabDebe.TabEnabled(0) = True
'            Else
''              SSTabDebe.TabEnabled(0) = False
'            End If
'            If rsPlan_cuentas!AUX2 <> "00" Then
''              SSTabDebe.TabEnabled(1) = True
'            Else
''              SSTabDebe.TabEnabled(1) = False
'            End If
'            If rsPlan_cuentas!aux3 <> "00" Then
''                SSTabDebe.TabEnabled(2) = True
'            Else
''              SSTabDebe.TabEnabled(2) = False
'            End If
'            auxDebe daux1
'            auxDebe daux2
'            auxDebe daux3
''            SSTabDebe_Click (0)
'        End If
'    End If
''    If lcta = "N" Then
''        Exit Sub
''    End If
''    If lcta = "S" Then
''        If MovCuenta = "T" Then
''            Exit Sub
''            'Me.CboDCtaCAM.SetFocus
''            'Me.CboDCta.SetFocus
''        End If
''        If MovCuenta = "D" Then
''            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
''            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(CboDCtaCAM) & "' and subcta1='" & Trim(CboDSub1CAM) & "' and subcta2='" & Trim(Me.CboDSub2CAM) & "'"
''            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
''            daux1 = Trim(rsPlan_cuentas!aux1)
''            daux2 = Trim(rsPlan_cuentas!aux2)
''            daux3 = Trim(rsPlan_cuentas!aux3)
''            Select Case rsPlan_cuentas!aux1
''                Dim sql1 As String
''                Case "00" ' no se introduce nada
''                    frameDOrganismos.Visible = False
''                    frameDaux00.Visible = True
''                    frameDCtaBancaria.Visible = False
''                    Me.FrameDBeneficiario.Visible = False
''                    dctalarga = ""
''                Case "01" ' se introduce un beneficiario
''                    frameDOrganismos.Visible = False
''                    frameDaux00.Visible = False
''                    frameDCtaBancaria.Visible = False
''                    Me.FrameDBeneficiario.Visible = True
''                    Me.lblDBenefaux1 = Trim(Me.DtCDcodbenef.Text)
''                    Me.lblDnomBenefaux1 = Trim(Me.dtc_desc4.Text)
''                    dctalarga = Trim(Me.DtCDcodbenef.Text)
''                Case "02" 'se introduce una cuenta bancaria
''                    frameDOrganismos.Visible = False
''                    frameDaux00.Visible = False
''                    Me.FrameDBeneficiario.Visible = False
''                    frameDCtaBancaria.Visible = True
''                    If Trim(CboDCtaCAM) = "1111" And Trim(CboDSub1CAM) = "02" Then
''                        Select Case Me.CboDSub2CAM
''                            Case "01"
''                                sql1 = "SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                                    "where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
''                            Case "02"
''                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                                    "where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
''                            Case "03"
''                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                                    "where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
''                        End Select
''                        Me.cboDctaaux1.Clear
''                        Me.cboDctanomaux1.Clear
''                        Set rscta_corrienteDebe = New ADODB.Recordset
''                        rscta_corrienteDebe.Filter = adFilterNone
''                        If rscta_corrienteDebe.State = 1 Then rscta_corrienteDebe.Close
''                        rscta_corrienteDebe.CursorLocation = adUseClient
''                        rscta_corrienteDebe.Open sql1, db, adOpenForwardOnly, adLockReadOnly
''                        If rscta_corrienteDebe.RecordCount <> 0 Then
''                            rscta_corrienteDebe.MoveFirst
''                            Do While Not rscta_corrienteDebe.EOF
''                                cboDctaaux1.AddItem rscta_corrienteDebe!cta_codigo
''                                cboDctanomaux1.AddItem rscta_corrienteDebe!cta_descripcion
''                                rscta_corrienteDebe.MoveNext
''                            Loop
''                        End If
''                    End If
''                Case "08"
''                    frameDaux00.Visible = False
''                    Me.FrameDBeneficiario.Visible = False
''                    frameDCtaBancaria.Visible = False
''                    frameDOrganismos.Enabled = True
''                    frameDOrganismos.Visible = True
''                    If rsorganismo.State = 1 Then rsorganismo.Close
''                    rsorganismo.CursorLocation = adUseClient
''                    rsorganismo.Filter = adFilterNone
''                    rsorganismo.Open "SELECT Org_codigo,(Org_descripcion) AS descripcion" & _
''                                      " From fc_organismo_financiamiento order by org_codigo", db, adOpenKeyset, adLockReadOnly
''                    cboDCodOrg.Clear
''                    cboDDenomOrg.Clear
''                    If rsorganismo.RecordCount <> 0 Then
''                      rsorganismo.MoveFirst
''                      Do While Not rsorganismo.EOF
''                          cboDCodOrg.AddItem rsorganismo!org_codigo
''                          cboDDenomOrg.AddItem rsorganismo!descripcion
''                          rsorganismo.MoveNext
''                      Loop
''                    End If
''                Case Else ' no se ha definido todavia
''                    frameDaux00.Visible = True
''                    frameDCtaBancaria.Visible = False
''                    Me.FrameDBeneficiario.Visible = False
''                    dctalarga = ""
''            End Select
''        End If
''    End If
'End Sub
'
'Private Sub CboDSubcta1_Click()
'    On Error GoTo Laberror1
'    Me.CboDSubcta2.Clear
'      If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
'      rssubcuenta.Open "SELECT SubCta2,Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.CboDCta.Text) & "') AND (SubCta1 ='" & Trim(Me.CboDSubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
'      If rssubcuenta.RecordCount = 0 Then
'        Me.CboDSubcta2.AddItem "00"
'      Else
'        rssubcuenta.MoveFirst
'        Do While Not rssubcuenta.EOF
'           Me.CboDSubcta2.AddItem rssubcuenta!subcta2
'           rssubcuenta.MoveNext
'        Loop
'      End If
'   ' Me.CboDSubcta2.Text = Me.CboDSubcta2.List(0)
'Exit Sub
'Laberror1:
'If Err.Number = 3021 Then
' MsgBox "Elija una cuenta", vbExclamation + vbDefaultButton1
' Me.CboDCta.SetFocus
'End If
'End Sub
'
'Private Sub CboDSubcta1_KeyPress(KeyAscii As Integer)
''  KeyAscii = 0
'End Sub

Private Sub BtnAñadir1_Click()
   On Error GoTo AddErr

     Call ABRIR_DEBE
     
     Ado_detalle1.Recordset.AddNew
     VAR_SW2 = "ADD"
     
     Fra_ABM2.Visible = True
     Fram_AsientoH.Visible = False
     Fram_AsientoH.Enabled = False
     Fram_AsientoD.Visible = True
     Picture2.Visible = True
     FrmABMDet1.Visible = False
     FraDet1.Enabled = False
     FraDet2.Enabled = False
     FraNavega.Enabled = False
     fraOpciones.Enabled = False
     FraGlobal.Enabled = False
     BtnCancelar1.Enabled = True
     
     D_Cambio_cmb.Text = GlTipoCambioOficial
     'D_Nombre_cmb.SetFocus
'VAR_SW = ""
  Exit Sub
AddErr:
  MsgBox Err.Description
  
End Sub

Private Sub BtnAprobar_Click()
 On Error GoTo UpdateErr
   If Ado_datos.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datos.Recordset!estado_codigo = "APR"
         Ado_datos.Recordset!Fecha_transacion = Date
        ' Ado_datos.Recordset!usr_codigo = glusuario
         Ado_datos.Recordset.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    VAR_BUS = 1
    buscados = 1
    Call opttodos_Click
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
    
End Sub

Private Sub BtnCancelar1_Click()
'FrmABMDet2.Visible = True
On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
            Fra_ABM2.Visible = False
            FrmABMDet1.Visible = True
            Fram_AsientoD.Visible = False
            Fram_AsientoH.Visible = False
            Fram_AsientoH.Enabled = True
            
             FraDet1.Enabled = True
             FraDet2.Enabled = True
            
            fraOpciones.Enabled = True
            FraNavega.Enabled = True
            FraGlobal.Enabled = True
            BtnCancelar1.Enabled = False
            VAR_SW = ""
   End If
     Exit Sub
'Call ABRIR_DEBE
End Sub

Private Sub BtnEliminar1_Click()
On Error GoTo UpdateErr
   'If ExisteReg(rs_datos!descargo_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención":
   If Ado_detalle1.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
'         Ado_datos.Recordset!estado_codigo = "ANL"
'         Ado_datos.Recordset!Fecha_transacion = Date
'         Ado_datos.Recordset!usr_codigo = glusuario
'         Ado_datos.Recordset.UpdateBatch adAffectAll
          Ado_detalle1.Recordset.Delete
          Call ABRIR_HABER
          Call ABRIR_DEBE
          
      End If
      
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
    MsgBox Err.Description
    
End Sub

Private Sub BtnGrabar1_Click()
On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos1
  If VAR_VAL = "OK" Then
    If VAR_SW2 = "ADD" Then

        db.Execute " INSERT INTO co_diario (Cod_Comp,            Cod_Comp_Detalle,               estado_codigo,         D_Cuenta,                    D_Nombre,                    D_Subcta1,                       D_SubCta2,                  D_Aux1,                        D_Aux2,                       D_Aux3,                D_Cta_Aux1,                      D_Cta_Aux2,             D_Cta_Aux3,               D_Des_Aux1,              D_Des_Aux2,                   D_Des_Aux3,            D_MontoBs,                          D_MontoDl,                       D_Cambio,                    D_Correl,          tipo_moneda, " & _
                                           " H_Cuenta,                  H_Nombre,                 H_Subcta1,                   H_SubCta2,                     H_Aux1,                    H_Aux2,                          H_Aux3,                     H_Cta_Aux1,                  H_Cta_Aux2,                   H_Cta_Aux3,                H_Des_Aux1,              H_Des_Aux2,               H_Des_Aux3,                    H_MontoBs,                      H_MontoDl,                    H_Cambio,              H_Correl,                Usr_codigo,                    Fecha_registro) " & _
                        " VALUES (" & TxtComprobante.Text & "," & Ado_detalle1.Recordset.RecordCount & ",'REG','" & D_Cuenta_cmb.Text & "','" & D_Nombre_cmb.Text & "','" & D_Subcta1_cmb.Text & "','" & D_Subcta2_cmb.Text & "','" & D_Cta_Aux1_cmb.Text & "','" & D_Cta_Aux2_cmb.Text & "','" & D_Cta_Aux3_cmb.Text & "','" & dtc_codigo8.Text & "','" & dtc_codigo9.Text & "','" & dtc_codigo10.Text & "', '" & dtc_desc8.Text & "','" & dtc_desc9.Text & "','" & dtc_desc10.Text & "'," & CDbl(D_MontoBs_cmb.Text) & "," & CDbl(D_MontoDl_cmb.Text) & "," & D_Cambio_cmb.Text & "," & D_Correl_cmb.Text & ",'" & cmb_moneda.Text & "', " & _
                                    " '" & H_Cuenta_cmb.Text & "','" & H_Nombre_cmb.Text & "','" & H_Subcta1_cmb.Text & "','" & H_Subcta2_cmb.Text & "','" & H_Cta_Aux1_cmb.Text & "','" & H_Cta_Aux2_cmb.Text & "','" & H_Cta_Aux3_cmb.Text & "','" & dtc_codigo11.Text & "','" & dtc_codigo12.Text & "','" & dtc_codigo13.Text & "', '" & dtc_desc11.Text & "','" & dtc_desc12.Text & "','" & dtc_desc13.Text & "'," & CDbl(D_MontoBs_cmb.Text) & "," & CDbl(D_MontoDl_cmb.Text) & "," & D_Cambio_cmb.Text & "," & H_Correl_cmb.Text & ",'" & Trim(glusuario) & "','" & Format(Date, "dd/mm/yyyy") & "' )"
        
        ',                   Hora_registro
        ','" & Format(Time, "hh:mm:ss") & "'
        
        'db.Execute sql_adicionM
    End If
    
    If VAR_SW2 = "MOD" Then
             db.Execute " UPDATE co_diario set D_Cta_Aux1='" & dtc_codigo8.Text & "',D_Des_Aux1='" & dtc_desc8.Text & "',D_Cta_Aux2='" & dtc_codigo9.Text & "',D_Des_Aux2='" & dtc_desc9.Text & "',D_Cta_Aux3='" & dtc_codigo10.Text & "',D_Des_Aux3='" & dtc_desc10.Text & "',D_MontoBs=" & CDbl(D_MontoBs_cmb) & ",D_MontoDl=" & CDbl(D_MontoDl_cmb) & ",D_Cambio=" & D_Cambio_cmb & " , " & _
                                               " H_Cta_Aux1='" & dtc_codigo11.Text & "',H_Des_Aux1='" & dtc_desc11.Text & "',H_Cta_Aux2='" & dtc_codigo12.Text & "',H_Des_Aux2='" & dtc_desc12.Text & "',H_Cta_Aux3='" & dtc_codigo13.Text & "',H_Des_Aux3='" & dtc_desc13.Text & "',H_MontoBs=" & CDbl(D_MontoBs_cmb) & ",H_MontoDl=" & CDbl(D_MontoDl_cmb) & " WHERE co_diario.Cod_Comp= " & Ado_detalle1.Recordset!Cod_Comp & " AND co_diario.Cod_Comp_Detalle= " & Ado_detalle1.Recordset!Cod_Comp_Detalle & " "
    End If
        db.Execute "UPDATE co_diario SET NOMCTADEBE = (SELECT CC_Plan_Cuentas.NombreCta From CC_Plan_Cuentas Where CC_Plan_Cuentas.Cuenta =  '" & D_Cuenta_cmb.Text & "' and CC_Plan_Cuentas.nivel = '4')"
        db.Execute "UPDATE co_diario SET NOMCTAHABER = (SELECT CC_Plan_Cuentas.NombreCta From CC_Plan_Cuentas Where CC_Plan_Cuentas.Cuenta =  '" & H_Cuenta_cmb.Text & "' and CC_Plan_Cuentas.nivel = '4')"
  
       'adiciona = ""
     
        FrmABMDet1.Visible = True
        Fram_AsientoD.Visible = False
        Fra_Aux.Enabled = False
        dtc_codigo8.Visible = False
        dtc_desc8.Visible = False
        dtc_codigo9.Visible = False
        dtc_desc9.Visible = False
        dtc_codigo10.Visible = False
        dtc_desc10.Visible = False
        
'        Buscar1.Visible = False
'        Buscar2.Visible = False
'        Buscar3.Visible = False
        
        rs_datos.Update
        rs_datos.MoveLast
        mbDataChanged = False
    
         FraNavega.Enabled = True
         FrmABMDet1.Visible = True
         FraDet1.Enabled = True
         FraDet2.Enabled = True
         FraGlobal.Enabled = True
         fraOpciones.Enabled = True
         BtnCancelar1.Enabled = False
         Fra_ABM2.Visible = False
         
         dg_datos.Enabled = True
      VAR_SW = " "
         'Ado_detalle1.Recordset.Requery
    End If
        'Call ABRIR_DEBE
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos1()
  If D_Cuenta_cmb.Text = "" Then
    MsgBox "Debe registrar la " + lbl_cuenta.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If D_Nombre_cmb.Text = "" Then
    MsgBox "Debe registrar: " + lblNombreC.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If cmb_moneda.Text = "" Then
    MsgBox "Debe registrar: " + lblMoneda.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If D_Cambio_cmb.Text = "" Then
    MsgBox "Debe registrar: " + lblTipo.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
    If D_MontoBs_cmb.Text = "" Then
    MsgBox "Debe registrar: " + lblMonto.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
    If D_MontoDl_cmb.Text = "" Then
    MsgBox "Debe registrar: " + lblMontoDls.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
    If H_Cuenta_cmb.Text = "" Then
    MsgBox "Debe registrar: " + lbl_cuenta1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
      If H_Nombre_cmb.Text = "" Then
    MsgBox "Debe registrar: " + lblNombre1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir1_Click()
  If Ado_detalle1.Recordset!d_cuenta = "" Then
    MsgBox "Debe Seleccionar una Cuenta para obtener un resultado, Vuelva a intentar ...", vbCritical + vbExclamation, "Validación de datos"
    Exit Sub
  Else
'valida_campos
'If (Ado_detalle1.Recordset.RecordCount > 0) Then
'And (Ado_datos8.Recordset.RecordCount > 0)
'And (Ado_datos9.Recordset.RecordCount > 0) Or (Ado_datos10.Recordset.RecordCount > 0)
        Dim iResult As Integer
        CryDebe.ReportFileName = App.Path & "\REPORTES\Contabilidad\cr_mayor_auxiliar_bs.rpt"
        CryDebe.WindowShowPrintSetupBtn = True
        CryDebe.WindowShowRefreshBtn = True
        'If stp_fecha_inicio.Value = "" Then     '@FFInicio varchar(10),
        '     CryDebe.StoredProcParam(0) = "%"
        ' Else
            CryDebe.StoredProcParam(0) = Format("01/01/2015", "dd/mm/yyyy")
        'End If
        'If stp_fecha_final.Value = "" Then      '@FFFinal varchar(10) ,
        '     CryDebe.StoredProcParam(1) = "%"
        ' Else
            CryDebe.StoredProcParam(1) = Format("30/06/2015", "dd/mm/yyyy")
        'End If
        If Ado_detalle1.Recordset!d_cuenta = "" Then             '@cuenta  varchar  (5) ,
            CryDebe.StoredProcParam(2) = "%"
         Else
            CryDebe.StoredProcParam(2) = Trim(Ado_detalle1.Recordset!d_cuenta)
        End If
        
        If Ado_detalle1.Recordset!d_subcta1 = "" Then           '@subcta1 varchar (3) ,
             CryDebe.StoredProcParam(3) = "%"
         Else
            CryDebe.StoredProcParam(3) = Trim(Ado_detalle1.Recordset!d_subcta1)
        End If
        
        If Ado_detalle1.Recordset!d_subcta2 = "" Then           '@subcta2 varchar (3) ,
             CryDebe.StoredProcParam(4) = "%"
         Else
            CryDebe.StoredProcParam(4) = Trim(Ado_detalle1.Recordset!d_subcta2)
        End If
'
        If Ado_detalle1.Recordset!D_Cta_Aux1 = "" Then            '@busca1 varchar(40),
            CryDebe.StoredProcParam(5) = "%"
         Else
            CryDebe.StoredProcParam(5) = Trim(Ado_detalle1.Recordset!D_Cta_Aux1)
        End If
        If Ado_detalle1.Recordset!D_Cta_Aux2 = "" Then           '@busca2 varchar(40),
            CryDebe.StoredProcParam(6) = "%"
         Else
            CryDebe.StoredProcParam(6) = Trim(Ado_detalle1.Recordset!D_Cta_Aux2)
        End If
        If Ado_detalle1.Recordset!D_Cta_Aux3 = "" Then          '@busca3 varchar(40),
            CryDebe.StoredProcParam(7) = "%"
         Else
            CryDebe.StoredProcParam(7) = Trim(Ado_detalle1.Recordset!D_Cta_Aux3)
        End If
        If Ado_detalle1.Recordset!d_Aux1 = "" Then              '@aux1 varchar(3),
             CryDebe.StoredProcParam(8) = "%"
         Else
            CryDebe.StoredProcParam(8) = Trim(Ado_detalle1.Recordset!d_Aux1)
        End If
        If Ado_detalle1.Recordset!d_Aux2 = "" Then              '@aux2 varchar(3),
             CryDebe.StoredProcParam(9) = "%"
         Else
            CryDebe.StoredProcParam(9) = Trim(Ado_detalle1.Recordset!d_Aux2)
        End If
        If Ado_detalle1.Recordset!d_Aux3 = "" Then              '@aux3 varchar(3)
            CryDebe.StoredProcParam(10) = "%"
         Else
            CryDebe.StoredProcParam(10) = Trim(Ado_detalle1.Recordset!d_Aux3)
        End If

        CryDebe.Formulas(2) = "nomaux1 = '" & Trim(Ado_detalle1.Recordset!D_Des_Aux1) & "' "
        CryDebe.Formulas(3) = "nomaux2 = '" & Trim(Ado_detalle1.Recordset!D_Des_Aux2) & "' "
        CryDebe.Formulas(4) = "nomaux3 = '" & Trim(Ado_detalle1.Recordset!D_Des_Aux3) & "' "
        CryDebe.Formulas(5) = "nomcta = '" & Trim(Ado_detalle1.Recordset!D_Nombre) & "' "
         
'         .Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
'         .Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"

'            If dtc_desc8.Text = "" Then
'         CryDebe.StoredProcParam(5) = "%"
'         Else
'          CryDebe.StoredProcParam(5) = dtc_desc8.Text
'        End If
'            If dtc_desc9.Text = "" Then
'         CryDebe.StoredProcParam(8) = "%"
'         Else
'          CryDebe.StoredProcParam(8) = dtc_desc9.Text
'        End If

'        If dtc_desc10.Text = "" Then
'         CryDebe.StoredProcParam(11) = "%"
'         Else
'          CryDebe.StoredProcParam(11) = dtc_desc10.Text
'        End If
        '
        '
        'nomaux2
        'nomaux3
        
        iResult = CryDebe.PrintReport
        If iResult <> 0 Then MsgBox CryDebe.LastErrorNumber & " : " & CryDebe.LastErrorString, vbCritical, "Error de impresión"
        CryDebe.WindowState = crptMaximized
 ' Else
 '   MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnImprimir2_Click()
  If Ado_detalle1.Recordset!d_cuenta = "" Then
    MsgBox "Debe Seleccionar una Cuenta para obtener un resultado, Vuelva a intentar ...", vbCritical + vbExclamation, "Validación de datos"
    Exit Sub
  Else
'valida_campos
'If (Ado_detalle1.Recordset.RecordCount > 0) Then
'And (Ado_datos8.Recordset.RecordCount > 0)
'And (Ado_datos9.Recordset.RecordCount > 0) Or (Ado_datos10.Recordset.RecordCount > 0)
        Dim iResult As Integer
        CryHaber.ReportFileName = App.Path & "\REPORTES\Contabilidad\cr_mayor_auxiliar_bs.rpt"
        CryHaber.WindowShowPrintSetupBtn = True
        CryHaber.WindowShowRefreshBtn = True
        'If stp_fecha_inicio.Value = "" Then     '@FFInicio varchar(10),
        '     CryHaber.StoredProcParam(0) = "%"
        ' Else
            CryHaber.StoredProcParam(0) = Format("01/01/2015", "dd/mm/yyyy")
        'End If
        'If stp_fecha_final.Value = "" Then      '@FFFinal varchar(10) ,
        '     CryHaber.StoredProcParam(1) = "%"
        ' Else
            CryHaber.StoredProcParam(1) = Format("30/06/2015", "dd/mm/yyyy")
        'End If
        If Ado_detalle1.Recordset!h_cuenta = "" Then             '@cuenta  varchar  (5) ,
            CryHaber.StoredProcParam(2) = "%"
         Else
            CryHaber.StoredProcParam(2) = Trim(Ado_detalle1.Recordset!h_cuenta)
        End If
        
        If Ado_detalle1.Recordset!h_subcta1 = "" Then           '@subcta1 varchar (3) ,
             CryHaber.StoredProcParam(3) = "%"
         Else
            CryHaber.StoredProcParam(3) = Trim(Ado_detalle1.Recordset!h_subcta1)
        End If
        
        If Ado_detalle1.Recordset!h_subcta2 = "" Then           '@subcta2 varchar (3) ,
             CryHaber.StoredProcParam(4) = "%"
         Else
            CryHaber.StoredProcParam(4) = Trim(Ado_detalle1.Recordset!h_subcta2)
        End If
'
        If Ado_detalle1.Recordset!H_Cta_Aux1 = "" Then            '@busca1 varchar(40),
            CryHaber.StoredProcParam(5) = "%"
         Else
            CryHaber.StoredProcParam(5) = Trim(Ado_detalle1.Recordset!H_Cta_Aux1)
        End If
        If Ado_detalle1.Recordset!H_Cta_Aux2 = "" Then           '@busca2 varchar(40),
            CryHaber.StoredProcParam(6) = "%"
         Else
            CryHaber.StoredProcParam(6) = Trim(Ado_detalle1.Recordset!H_Cta_Aux2)
        End If
        If Ado_detalle1.Recordset!H_Cta_Aux3 = "" Then          '@busca3 varchar(40),
            CryHaber.StoredProcParam(7) = "%"
         Else
            CryHaber.StoredProcParam(7) = Trim(Ado_detalle1.Recordset!H_Cta_Aux3)
        End If
        If Ado_detalle1.Recordset!h_Aux1 = "" Then              '@aux1 varchar(3),
             CryHaber.StoredProcParam(8) = "%"
         Else
            CryHaber.StoredProcParam(8) = Trim(Ado_detalle1.Recordset!h_Aux1)
        End If
        If Ado_detalle1.Recordset!h_Aux2 = "" Then              '@aux2 varchar(3),
             CryHaber.StoredProcParam(9) = "%"
         Else
            CryHaber.StoredProcParam(9) = Trim(Ado_detalle1.Recordset!h_Aux2)
        End If
        If Ado_detalle1.Recordset!h_Aux3 = "" Then              '@aux3 varchar(3)
            CryHaber.StoredProcParam(10) = "%"
         Else
            CryHaber.StoredProcParam(10) = Trim(Ado_detalle1.Recordset!h_Aux3)
        End If

        CryHaber.Formulas(2) = "nomaux1 = '" & Trim(Ado_detalle1.Recordset!H_Des_Aux1) & "' "
        CryHaber.Formulas(3) = "nomaux2 = '" & Trim(Ado_detalle1.Recordset!H_Des_Aux2) & "' "
        CryHaber.Formulas(4) = "nomaux3 = '" & Trim(Ado_detalle1.Recordset!H_Des_Aux3) & "' "
        CryHaber.Formulas(5) = "nomcta = '" & Trim(Ado_detalle1.Recordset!H_Nombre) & "' "
         
'         .Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
'         .Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"

'            If dtc_desc8.Text = "" Then
'         CryHaber.StoredProcParam(5) = "%"
'         Else
'          CryHaber.StoredProcParam(5) = dtc_desc8.Text
'        End If
'            If dtc_desc9.Text = "" Then
'         CryHaber.StoredProcParam(8) = "%"
'         Else
'          CryHaber.StoredProcParam(8) = dtc_desc9.Text
'        End If

'        If dtc_desc10.Text = "" Then
'         CryHaber.StoredProcParam(11) = "%"
'         Else
'          CryHaber.StoredProcParam(11) = dtc_desc10.Text
'        End If
        '
        '
        'nomaux2
        'nomaux3
        
        iResult = CryHaber.PrintReport
        If iResult <> 0 Then MsgBox CryHaber.LastErrorNumber & " : " & CryHaber.LastErrorString, vbCritical, "Error de impresión"
        CryHaber.WindowState = crptMaximized
 ' Else
 '   MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnModificar1_Click()
 On Error GoTo EditErr

If Ado_datos.Recordset!estado_codigo = "REG" Then
            VAR_SW2 = "MOD"
            cmodificar = "M"
            adiciona = "M"
            Me.Fra_ABM2.Enabled = False
            Me.Fra_ABM1.Enabled = False
            Me.FraNavega.Enabled = True
            Me.Fra_Aux.Enabled = True
            Me.Fra_ABM2.Visible = True
            Me.Fram_AsientoD.Enabled = True
            Me.Fram_AsientoD.Visible = True
           ' Me.dg_datos.Enabled = False
'            Me.fraOpciones.Visible = False
'            Me.FraGrabarCancelar.Visible = True
            Me.fraOpciones.Enabled = False
            Me.FraGrabarCancelar.Enabled = False
            Me.BtnCancelar1.Enabled = True

    Else
            MsgBox "No se puede MODIFICAR un registro APROBADO o Errado ...", vbExclamation, "Validación de Registro"
End If
Exit Sub
EditErr:
MsgBox Err.Description
End Sub

Private Sub Buscar1_Click()
VAR_AUX1 = D_Cta_Aux1_cmb
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And D_Cta_Aux1_cmb = "00" Then
        dtc_codigo8.Text = "0"
        dtc_desc8.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
        dtc_codigo8.Visible = True
        dtc_desc8.Visible = True
        Set rs_datos8 = New ADODB.Recordset
        If rs_datos8.State = 1 Then rs_datos8.Close
            rs_datos8.Open "Select " + VAR_CODIGO + " as codigo1 , " + VAR_DES + " as desc1 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos8.Recordset = rs_datos8
            dtc_desc8.BoundText = dtc_codigo8.BoundText
    End If
End Sub

Private Sub Buscar2_Click()
VAR_AUX1 = D_Cta_Aux2_cmb
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And D_Cta_Aux2_cmb = "00" Then
        dtc_codigo9.Text = "0"
        dtc_desc9.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
        
    Else
  dtc_codigo9.Visible = True
 dtc_desc9.Visible = True
        Set rs_datos9 = New ADODB.Recordset
        If rs_datos9.State = 1 Then rs_datos9.Close
            rs_datos9.Open "Select " + VAR_CODIGO + " as codigo2 , " + VAR_DES + " as desc2 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos9.Recordset = rs_datos9
            dtc_desc9.BoundText = dtc_codigo9.BoundText
    End If
End Sub

Private Sub Buscar3_Click()
VAR_AUX1 = D_Cta_Aux3_cmb
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And D_Cta_Aux3_cmb = "00" Then
        dtc_codigo10.Text = "0"
        dtc_desc10.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
  dtc_codigo10.Visible = True
 dtc_desc10.Visible = True
        Set rs_datos10 = New ADODB.Recordset
        If rs_datos10.State = 1 Then rs_datos10.Close
            rs_datos10.Open "Select " + VAR_CODIGO + " as codigo3 , " + VAR_DES + " as desc3 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos10.Recordset = rs_datos10
            dtc_desc10.BoundText = dtc_codigo10.BoundText
    End If
End Sub

'Private Sub CboDSubcta2_Click()
'    Dim sql_cuenta As String
'    CboDCtaCAM.Text = ""
'    Call Titulo(CboDCta, CboDSubcta1, CboDSubcta2)
''    If lcta = "N" Then
'        Exit Sub
'    End If
'    If lcta = "S" Then
'        If MovCuenta = "T" Then
'            Exit Sub
'            Me.CboDCta.SetFocus
'        End If
'        If MovCuenta = "D" Then
'            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
'            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(Me.CboDCta) & "' and subcta1='" & Trim(Me.CboDSubcta1) & "' and subcta2='" & Trim(Me.CboDSubcta2) & "'"
'            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
'            daux1 = Trim(rsPlan_cuentas!aux1)
'            daux2 = Trim(rsPlan_cuentas!AUX2)
'            daux3 = Trim(rsPlan_cuentas!aux3)
'            '---habilitacion de auxiliares---
'            If rsPlan_cuentas!aux1 <> "00" Then
''              SSTabDebe.TabEnabled(0) = True
'            Else
''              SSTabDebe.TabEnabled(0) = False
'            End If
'            If rsPlan_cuentas!AUX2 <> "00" Then
''              SSTabDebe.TabEnabled(1) = True
'            Else
''              SSTabDebe.TabEnabled(1) = False
'            End If
'            If rsPlan_cuentas!aux3 <> "00" Then
''                SSTabDebe.TabEnabled(2) = True
'            Else
''              SSTabDebe.TabEnabled(2) = False
'            End If
'            auxDebe daux1
'            auxDebe daux2
'            auxDebe daux3
''            SSTabDebe_Click (0)
'        End If
'    End If
'
'End Sub
'
'Private Sub CboDSubcta2_KeyPress(KeyAscii As Integer)
''  KeyAscii = 0
'End Sub
'
'Private Sub cboHCodOrg_Click()
'  On Error GoTo err3
'  rsorganismo.Filter = adFilterNone
'  rsorganismo.Filter = "org_codigo='" & Trim(Me.cboHCodOrg) & "'"
'  If rsorganismo.RecordCount <> 0 Then
'    Me.cboHDenomOrg.Text = Trim(rsorganismo!descripcion)
'  Else
'    Exit Sub
'  End If
'  hctalarga = Trim(cboHCodOrg.Text)
'err3:
'  If Err.Number = 28 Then
'    Exit Sub
'  End If
'End Sub
'
'Private Sub CboHcta_Click()
' Me.CbohSubcta1.Clear
'  Me.CbohSubcta2.Clear
'  rsplanctas.MoveFirst
'  rsplanctas.Find "cuenta=" & "'" & Trim(CboHcta.Text) & "'"
'  If rscuentas.State = adStateOpen Then rscuentas.Close
'  rscuentas.Open "SELECT SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.CboHcta.Text) & "')", db, adOpenKeyset, adLockReadOnly
'  Do While Not rscuentas.EOF
'    Me.CbohSubcta1.AddItem rscuentas!subcta1
'    rscuentas.MoveNext
'  Loop
'  If rscuentas.RecordCount = 0 Then
'  Me.CbohSubcta1.AddItem "00"
'  End If
'End Sub
'Private Sub cboHctaaux1_Click()
'    rscta_corrienteHaber.Filter = adFilterNone
''    If CboTipo = "CAM" And frameDOrganismos.Visible = True Then
''      rscta_corrienteHaber.Filter = "org_codigo='" & Trim(cboDCodOrg) & "'"
''    End If
'    rscta_corrienteHaber.Filter = "cta_codigo='" & Trim(Me.cboHctaaux1) & "'"
'    If rscta_corrienteHaber.RecordCount <> 0 Then
'      Me.cboHctanomaux1.Text = Trim(rscta_corrienteHaber!cta_descripcion)
'    Else
'      Exit Sub
'    End If
'    hctalarga = Trim(cboHctaaux1)
'End Sub
'
'
'
'Private Sub CboHCtaCAM_Click()
' Me.CboHSub1CAM.Clear
'  Me.CboHSub2CAM.Clear
'  rsplanctas.MoveFirst
'  rsplanctas.Find "cuenta=" & "'" & Trim(CboHCtaCAM.Text) & "'"
'  If rscuentas.State = adStateOpen Then rscuentas.Close
'  rscuentas.Open "SELECT SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.CboHCtaCAM.Text) & "')", db, adOpenKeyset, adLockReadOnly
'  Do While Not rscuentas.EOF
'    Me.CboHSub1CAM.AddItem rscuentas!subcta1
'    rscuentas.MoveNext
'  Loop
'   If Me.CboHCtaCAM.Text = "1111" Then
'      Me.CboHSub1CAM.Clear
'      Me.CboHSub1CAM.AddItem "02"
'  End If
'  If rscuentas.RecordCount = 0 Then
'    Me.CboHSub1CAM.AddItem "00"
'  End If
'  'Me.CboHSub1CAM.Text = Me.CboHSub1CAM.List(0)
'End Sub
'
'Private Sub cboHctanomaux1_Click()
'  rscta_corrienteHaber.MoveFirst
'    rscta_corrienteHaber.Find "cta_descripcion='" & Trim(Me.cboHctanomaux1) & "'"
'    cboHctaaux1.Text = rscta_corrienteHaber!Cta_Codigo
'    hctalarga = Trim(cboHctaaux1)
'End Sub
'Private Sub cboHDenomOrg_Click()
'On Error GoTo err1
'    rsorganismo.Filter = adFilterNone
'    rsorganismo.MoveFirst
'    rsorganismo.Find "descripcion='" & Trim(cboHDenomOrg) & "'"
'    cboHCodOrg = rsorganismo!org_codigo
'    dctalarga = Trim(cboHCodOrg)
'err1:
'    If Err.Number = 28 Then
'    Exit Sub
'    End If
'End Sub
'
'Private Sub CboHSub1CAM_Click()
' On Error GoTo Laberror1
'  Me.CboHSub2CAM.Clear
'  If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
'  rssubcuenta.Open "SELECT SubCta2,Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(CboHCtaCAM.Text) & "') AND (SubCta1 ='" & Trim(Me.CboHSub1CAM.Text) & "')", db, adOpenKeyset, adLockReadOnly
'    If rssubcuenta.RecordCount = 0 Then
'      CboHSub2CAM.AddItem "00"
'    Else
'      rssubcuenta.MoveFirst
'      Do While Not rssubcuenta.EOF
'        Me.CboHSub2CAM.AddItem rssubcuenta!subcta2
'        rssubcuenta.MoveNext
'      Loop
'    End If
'      If Me.CboHCtaCAM.Text = "1111" Then
'        For i = 0 To Me.CboHSub2CAM.ListCount
'          If Me.CboHSub2CAM.List(i) = "00" Then
'             Me.CboHSub2CAM.RemoveItem (i)
'          End If
'        Next
'      End If
'      'CboHSub2CAM.Text = CboHSub2CAM.List(0)
'Exit Sub
'Laberror1:
'If Err.Number = 3021 Then
' MsgBox "Elija una cuenta", vbExclamation + vbDefaultButton1
' 'Me.CboHcta.SetFocus
'End If
'End Sub

'Private Sub CboHSub2CAM_Change()
' Dim sql_cuenta As String
''    Call Titulo(Trim(Me.CboHCtaCAM), Trim(Me.CboHSub1CAM), Trim(CboHSub2CAM))
''    If lcta = "N" Then
'        Exit Sub
''    End If
''    If lcta = "S" Then
'        If MovCuenta = "T" Then
'            Exit Sub
''            Me.CboHCtaCAM.SetFocus
'        End If
'        If MovCuenta = "D" Then
''            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
''            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(Me.CboHCtaCAM) & "' and subcta1='" & Trim(CboHSub1CAM) & "' and subcta2='" & Trim(Me.CboHSub2CAM) & "'"
''            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
''            haux1 = Trim(rsPlan_cuentas!aux1)
''            haux2 = Trim(rsPlan_cuentas!AUX2)
''            haux3 = Trim(rsPlan_cuentas!aux3)
''            If rsPlan_cuentas!aux1 <> "00" Then
''              SSTabHaber.TabEnabled(0) = True
'            Else
''              SSTabHaber.TabEnabled(0) = False
'            End If
''            If rsPlan_cuentas!AUX2 <> "00" Then
''              SSTabHaber.TabEnabled(1) = True
'            Else
''              SSTabHaber.TabEnabled(1) = False
'            End If
'            If rsPlan_cuentas!aux3 <> "00" Then
''                SSTabHaber.TabEnabled(2) = True
'            Else
''              SSTabHaber.TabEnabled(2) = False
'            End If
''            Auxhaber haux1
''            Auxhaber haux2
''            Auxhaber haux3
''            SSTabHaber_Click (0)
'        End If
'    End If
'End Sub

'Private Sub CboHSub2CAM_Click()
'   Dim sql_cuenta As String
'   CboHcta.Text = ""
'    Call Titulo(Trim(Me.CboHCtaCAM), Trim(Me.CboHSub1CAM), Trim(CboHSub2CAM))
'    If lcta = "N" Then
'        Exit Sub
'    End If
'    If lcta = "S" Then
'        If MovCuenta = "T" Then
'            Exit Sub
'            Me.CboHCtaCAM.SetFocus
'        End If
'        If MovCuenta = "D" Then
'            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
'            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(Me.CboHCtaCAM) & "' and subcta1='" & Trim(CboHSub1CAM) & "' and subcta2='" & Trim(Me.CboHSub2CAM) & "'"
'            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
'            haux1 = Trim(rsPlan_cuentas!aux1)
'            haux2 = Trim(rsPlan_cuentas!AUX2)
'            haux3 = Trim(rsPlan_cuentas!aux3)
'            If rsPlan_cuentas!aux1 <> "00" Then
''              SSTabHaber.TabEnabled(0) = True
'            Else
''              SSTabHaber.TabEnabled(0) = False
'            End If
'            If rsPlan_cuentas!AUX2 <> "00" Then
''              SSTabHaber.TabEnabled(1) = True
'            Else
''              SSTabHaber.TabEnabled(1) = False
'            End If
'            If rsPlan_cuentas!aux3 <> "00" Then
''                SSTabHaber.TabEnabled(2) = True
'            Else
''              SSTabHaber.TabEnabled(2) = False
'            End If
'            Auxhaber haux1
'            Auxhaber haux2
'            Auxhaber haux3
''            SSTabHaber_Click (0)
'        End If
'    End If
''            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
''            haux1 = Trim(rsPlan_cuentas!aux1)
''            haux2 = Trim(rsPlan_cuentas!aux2)
''            haux3 = Trim(rsPlan_cuentas!aux3)
''            Select Case rsPlan_cuentas!aux1
''                Case "00" ' no se introduce nada
''                    Me.frameHOrganismos.Visible = False
''                    frameHAux00.Visible = True
''                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
''                    hctalarga = ""
''                Case "01" ' se introduce un beneficiario
''                    Me.frameHOrganismos.Visible = False
''                    frameHAux00.Visible = False
''                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = True
''                    Me.lblHBenefaux1 = Trim(Me.DtCHcodbenef.Text)
''                    Me.lblHnomBenefaux1 = Trim(Me.dtc_desc4.Text)
''                    hctalarga = Trim(Me.DtCHcodbenef.Text)
''                 Case "02" 'se introduce una cuenta bancaria
''                    frameHAux00.Visible = False
''                    frameHCtaBancaria.Visible = True
''                    Me.FrameHBeneficiario.Visible = False
''                    Me.frameHOrganismos.Visible = False
''                    If Trim(CboHCtaCAM) = "1111" And Trim(CboHSub1CAM) = "02" Then
''                        Select Case Me.CboHSub2CAM
''                            Case "01"
''                                sql1 = "SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                                    "where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
''                            Case "02"
''                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                                    "where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
''                            Case "03"
''                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                                    "where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
''                        End Select
''                        Me.cboHctaaux1.Clear
''                        Me.cboHctanomaux1.Clear
''                        If rscta_corrienteHaber.State = 1 Then rscta_corrienteHaber.Close
''                        Set rscta_corrienteHaber = New ADODB.Recordset
''                        rscta_corrienteHaber.Filter = adFilterNone
''                        rscta_corrienteHaber.CursorLocation = adUseClient
''                        rscta_corrienteHaber.Open sql1, db, adOpenForwardOnly, adLockReadOnly
''                        If rscta_corrienteHaber.RecordCount <> 0 Then
''                            rscta_corrienteHaber.MoveFirst
''                            Do While Not rscta_corrienteHaber.EOF
''                                cboHctaaux1.AddItem rscta_corrienteHaber!cta_codigo
''                                cboHctanomaux1.AddItem rscta_corrienteHaber!cta_descripcion
''                                rscta_corrienteHaber.MoveNext
''                            Loop
''                        End If
''                    End If
''                Case "08"
''                    frameHAux00.Visible = False
''                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
''                    Me.frameHOrganismos.Visible = True
''                    Me.frameHOrganismos.Enabled = True
''                    If rsorganismo.State = 1 Then rsorganismo.Close
''                    rsorganismo.CursorLocation = adUseClient
''                    rsorganismo.Filter = adFilterNone
''                    rsorganismo.Open "SELECT Org_codigo,(Org_descripcion) AS descripcion" & _
''                                      " From fc_organismo_financiamiento order by org_codigo", db, adOpenKeyset, adLockReadOnly
''                    cboHCodOrg.Clear
''                    cboHDenomOrg.Clear
''                    If rsorganismo.RecordCount <> 0 Then
''                      rsorganismo.MoveFirst
''                      Do While Not rsorganismo.EOF
''                          cboHCodOrg.AddItem rsorganismo!org_codigo
''                          cboHDenomOrg.AddItem rsorganismo!descripcion
''                          rsorganismo.MoveNext
''                      Loop
''                    End If
''                Case Else ' no se ha definido todavia
''                    frameHAux00.Visible = True
''                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
''                    Me.frameHOrganismos.Visible = False
''                    hctalarga = ""
''            End Select
''        End If
''    End If
'End Sub
'
'Private Sub CbohSubcta1_Click()
'  On Error GoTo Laberror1
'  Me.CbohSubcta2.Clear
'  If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
'  rssubcuenta.Open "SELECT SubCta2,Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.CboHcta.Text) & "') AND (SubCta1 ='" & Trim(Me.CbohSubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
'    If rssubcuenta.RecordCount = 0 Then
'      Me.CbohSubcta2.AddItem "00"
'    Else
'      rssubcuenta.MoveFirst
'      Do While Not rssubcuenta.EOF
'        Me.CbohSubcta2.AddItem rssubcuenta!subcta2
'        rssubcuenta.MoveNext
'      Loop
'    End If
'Exit Sub
'Laberror1:
'If Err.Number = 3021 Then
' MsgBox "Elija una cuenta", vbExclamation + vbDefaultButton1
' Me.CboHcta.SetFocus
'End If
'End Sub

'Private Sub CbohSubcta2_Change()
'   Dim sql_cuenta As String
''    Call Titulo(Trim(Me.CboHcta), Trim(Me.CbohSubcta1), Trim(CbohSubcta2))
''    If lcta = "N" Then
'        Exit Sub
'    End If
''    If lcta = "S" Then
'        If MovCuenta = "T" Or MovCuenta = "S" Then
'            Exit Sub
''            Me.CboHcta.SetFocus
'        End If
'        If MovCuenta = "D" Then
''            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
''            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(Me.CboHcta) & "' and subcta1='" & Trim(Me.CbohSubcta1) & "' and subcta2='" & Trim(Me.CbohSubcta2) & "'"
''            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
''            haux1 = Trim(rsPlan_cuentas!aux1)
''            haux2 = Trim(rsPlan_cuentas!AUX2)
''            haux3 = Trim(rsPlan_cuentas!aux3)
''            If rsPlan_cuentas!aux1 <> "00" Then
''              SSTabHaber.TabEnabled(0) = True
'            Else
''              SSTabHaber.TabEnabled(0) = False
'            End If
''            If rsPlan_cuentas!AUX2 <> "00" Then
''              SSTabHaber.TabEnabled(1) = True
'            Else
''              SSTabHaber.TabEnabled(1) = False
'            End If
''            If rsPlan_cuentas!aux3 <> "00" Then
''                SSTabHaber.TabEnabled(2) = True
'            Else
''              SSTabHaber.TabEnabled(2) = False
'            End If
''            Auxhaber haux1
''            Auxhaber haux2
'            Auxhaber haux3
''            SSTabHaber_Click (0)
'        End If
'    End If
'End Sub

'Private Sub CbohSubcta2_Click()
'  Dim sql_cuenta As String
'  CboHCtaCAM.Text = ""
'    Call Titulo(Trim(Me.CboHcta), Trim(Me.CbohSubcta1), Trim(CbohSubcta2))
'    If lcta = "N" Then
'        Exit Sub
'    End If
'    If lcta = "S" Then
'        If MovCuenta = "T" Then
'            Exit Sub
'            Me.CboHcta.SetFocus
'        End If
'        If MovCuenta = "D" Then
'            If rsPlan_cuentas.State = 1 Then rsPlan_cuentas.Close
'            sql_cuenta = "select aux1,aux2,aux3 from cc_Plan_cuentas where cuenta='" & Trim(Me.CboHcta) & "' and subcta1='" & Trim(Me.CbohSubcta1) & "' and subcta2='" & Trim(Me.CbohSubcta2) & "'"
'            rsPlan_cuentas.Open sql_cuenta, db, adOpenForwardOnly, adLockReadOnly
'            haux1 = Trim(rsPlan_cuentas!aux1)
'            haux2 = Trim(rsPlan_cuentas!AUX2)
'            haux3 = Trim(rsPlan_cuentas!aux3)
'            If rsPlan_cuentas!aux1 <> "00" Then
''              SSTabHaber.TabEnabled(0) = True
'            Else
''              SSTabHaber.TabEnabled(0) = False
'            End If
'            If rsPlan_cuentas!AUX2 <> "00" Then
''              SSTabHaber.TabEnabled(1) = True
'            Else
''              SSTabHaber.TabEnabled(1) = False
'            End If
'            If rsPlan_cuentas!aux3 <> "00" Then
''                SSTabHaber.TabEnabled(2) = True
'            Else
''              SSTabHaber.TabEnabled(2) = False
'            End If
'            Auxhaber haux1
'            Auxhaber haux2
'            Auxhaber haux3
''            SSTabHaber_Click (0)
'        End If
'    End If
'End Sub
'
''Private Sub cboNomTipo_Change()
''rstipocomp.Filter = adFilterNone
''    rstipocomp.Filter = "Denominacion_Tipo='" & Trim(CboTipo.Text) & "'"
''    If rstipocomp.RecordCount <> 0 Then
''        CboTipo.Text = Trim(rstipocomp!Codigo_Tipo)
''    End If
''End Sub
'
'Private Sub cboNomTipo_Click()
'rstipocomp.Filter = adFilterNone
'    rstipocomp.Filter = "Denominacion_Tipo='" & Trim(cboNomTipo.Text) & "'"
'    If rstipocomp.RecordCount <> 0 Then
'        CboTipo.Text = Trim(rstipocomp!Codigo_tipo)
'    End If
'End Sub
'
'Private Sub CboTipo_Change()
'  rstipocomp.Filter = adFilterNone
'    rstipocomp.Filter = "Codigo_Tipo='" & Trim(CboTipo.Text) & "'"
'    If rstipocomp.RecordCount <> 0 Then
'        cboNomTipo.Text = rstipocomp!Denominacion_Tipo
'    End If
'End Sub
'
''Private Sub CboTipo_Change()
''    rstipocomp.Filter = adFilterNone
''    rstipocomp.Filter = "Codigo_Tipo='" & Trim(CboTipo.Text) & "'"
''    If rstipocomp.RecordCount <> 0 Then
''        cboNomTipo.Text = rstipocomp!Denominacion_Tipo
''    End If
''End Sub
'
''Private Sub CboTipo_Click()
''Select Case Trim(CboTipo.Text)
''    Case "PCO"
''        Me.DTPCAM.Visible = False
''        Me.txt_fecha.Visible = True
''        Me.txtcodsolicitud.Visible = False
''        Label26.Visible = False 'codigo solicitud
''        Me.dtc_codigo4.Text = "-"
''        Me.lblDTC.Visible = True
''        lblHTC.Visible = True
''        lblHTIPOCAM.Visible = True
''        lblDTIPOCAM.Visible = True
''        lblDMonSus.Visible = True
''        lblHMONSUS.Visible = True
''        TxtDSus.Visible = True
''        txtHsus.Visible = True
''        Me.lblDTC.Visible = True
''        Me.lblDTC.Locked = False
''        Me.lblDTC = CTipoC
''        Me.CboDCtaCAM.Visible = False
''        Me.CboDSub1CAM.Visible = False
''        Me.CboDSub2CAM.Visible = False
''        Me.CboHCtaCAM.Visible = False
''        Me.CboHSub1CAM.Visible = False
''        Me.CboHSub2CAM.Visible = False
''        Me.frame_moneda.Enabled = True
''        CboDCta.Visible = True
''        CboDSubcta1.Visible = True
''        CboDSubcta2.Visible = True
''        CboHcta.Visible = True
''        CbohSubcta1.Visible = True
''        CbohSubcta2.Visible = True
''    Case "PCE"
''        Me.DTPCAM.Visible = False
''        Me.txt_fecha.Visible = True
''        Me.txtcodsolicitud.Visible = True
''        Label26.Visible = True
''        Me.lblDTC.Visible = True
''        lblHTC.Visible = True
''        lblHTIPOCAM.Visible = True
''        lblDTIPOCAM.Visible = True
''        lblDMonSus.Visible = True
''        lblHMONSUS.Visible = True
''        TxtDSus.Visible = True
''        txtHsus.Visible = True
''        Me.lblDTC.Visible = True
''        Me.lblDTC.Locked = True
''        Me.lblDTC = CTipoC
''        Me.CboDCtaCAM.Visible = False
''        Me.CboDSub1CAM.Visible = False
''        Me.CboDSub2CAM.Visible = False
''        Me.CboHCtaCAM.Visible = False
''        Me.CboHSub1CAM.Visible = False
''        Me.CboHSub2CAM.Visible = False
''        CboDCta.Visible = True
''        CboDSubcta1.Visible = True
''        CboDSubcta2.Visible = True
''        CboHcta.Visible = True
''        CbohSubcta1.Visible = True
''        CbohSubcta2.Visible = True
''        Me.frame_moneda.Enabled = True
''    Case "CAM"
''        Me.DTPCAM.Visible = True
''        Me.txt_fecha.Visible = False
''        Me.txtcodsolicitud.Visible = False
''        Label26.Visible = False 'codigo solicitud
''        Me.dtc_codigo4.Text = "-"
''        Me.lblDTC = "0.0"
''        lblHTC = "0.0"
''        Me.lblDTC.Visible = False
''        lblHTC.Visible = False
''        lblHTIPOCAM.Visible = False
''        lblDTIPOCAM.Visible = False
''        lblDMonSus.Visible = False
''        lblHMONSUS.Visible = False
''        Me.txtHsus.Visible = False
''        Me.TxtDSus.Visible = False
''        Me.TxtDSus = "0.0"
''        Me.txtHsus = "0.0"
''        CboDCta.Visible = False
''        CboDSubcta1.Visible = False
''        CboDSubcta2.Visible = False
''        CboHcta.Visible = False
''        CbohSubcta1.Visible = False
''        CbohSubcta2.Visible = False
''        Me.CboDCtaCAM.Visible = True
''        Me.CboDSub1CAM.Visible = True
''        Me.CboDSub2CAM.Visible = True
''        Me.CboHCtaCAM.Visible = True
''        Me.CboHSub1CAM.Visible = True
''        Me.CboHSub2CAM.Visible = True
''        Me.frame_moneda.Enabled = False
''        Me.optbolivianos = True
''End Select
'' ' Dim rsbustipo As ADODB.Recordset
'' ' Set rsbustipo = New ADODB.Recordset
''
''  rstipocomp.Filter = adFilterNone
''    rstipocomp.Filter = "Codigo_Tipo='" & Trim(CboTipo.Text) & "'"
''    If rstipocomp.RecordCount <> 0 Then
''        cboNomTipo.Text = rstipocomp!Denominacion_Tipo
''    End If
''End Sub

'Private Sub CboTipo_Click()
'  Select Case Trim(CboTipo.Text)
'    Case "PCO"
'      ' TxtDBs.Enabled = True
'      '  TxtDSus.Enabled = True
'        Me.frameCAM.Visible = False
'        Me.DTPCAM.Visible = False
'        Me.txt_fecha.Visible = True
'        Me.txtcodsolicitud.Visible = False
'        Label26.Visible = False 'codigo solicitud
'       If adiciona = "S" Then
'        Me.dtc_codigo4.Text = "-"
'       End If
'        Me.lblDTC.Visible = True
'        lblHTC.Visible = True
'        lblHTIPOCAM.Visible = True
'        lblDTIPOCAM.Visible = True
'        lblDMonSus.Visible = True
'        lblHMONSUS.Visible = True
'        TxtDSus.Visible = True
'        txtHsus.Visible = True
'        Me.lblDTC.Visible = True
'        Me.lblDTC.Locked = False
'        '--
'        dtc_codigo4.Visible = True
'        dtc_desc4.Visible = True
'        DtCHDescripbenef.Visible = True
'        DtCHcodbenef.Visible = True
'        lblDBenefaux1.Visible = False
'        lblDnomBenefaux1.Visible = False
'        lblHBenefaux1.Visible = fALS
'        lblHnomBenefaux1.Visible = False
'        '----
'      If adiciona = "S" Then
'        Me.lblDTC = CTipoC
'        lblDTC_Change
'      End If
'
'        Me.CboDCtaCAM.Visible = False
'        Me.CboDSub1CAM.Visible = False
'        Me.CboDSub2CAM.Visible = False
'        Me.CboHCtaCAM.Visible = False
'        Me.CboHSub1CAM.Visible = False
'        Me.CboHSub2CAM.Visible = False
'        Me.frame_moneda.Enabled = True
'        CboDCta.Visible = True
'        CboDSubcta1.Visible = True
'        CboDSubcta2.Visible = True
'        CboHcta.Visible = True
'        CbohSubcta1.Visible = True
'        CbohSubcta2.Visible = True
'        optbolivianos_Click
'        TxtDBs = ""
'        TxtDSus = ""
'    Case "PCE"
'      '  TxtDBs.Enabled = True
'      '  TxtDSus.Enabled = True
'        Me.frameCAM.Visible = False
'        Me.DTPCAM.Visible = False
'        Me.txt_fecha.Visible = True
'        Me.txtcodsolicitud.Visible = True
'        Label26.Visible = True
'        Me.lblDTC.Visible = True
'        lblHTC.Visible = True
'        Me.lblDTC.Locked = True
'        '----------
'        dtc_codigo4.Visible = False
'        dtc_desc4.Visible = False
'        DtCHDescripbenef.Visible = False
'        DtCHcodbenef.Visible = False
'        lblDBenefaux1.Visible = True
'        lblDnomBenefaux1.Visible = True
'        lblHBenefaux1.Visible = True
'        lblHnomBenefaux1.Visible = True
'        '-----
'        'Me.lblDTC = CTipoC
'        If adiciona = "S" Then
'          Me.lblDTC = CTipoC
'          lblDTC_Change
'        End If
'        lblHTIPOCAM.Visible = True
'        lblDTIPOCAM.Visible = True
'        lblDMonSus.Visible = True
'        lblHMONSUS.Visible = True
'        TxtDSus.Visible = True
'        txtHsus.Visible = True
'        Me.lblDTC.Visible = True
'        Me.lblDTC.Locked = True
'        '---
'        lblDBenefaux1.Visible = True
'        lblDnomBenefaux1.Visible = True
'        '---
'        Me.CboDCtaCAM.Visible = False
'        Me.CboDSub1CAM.Visible = False
'        Me.CboDSub2CAM.Visible = False
'        Me.CboHCtaCAM.Visible = False
'        Me.CboHSub1CAM.Visible = False
'        Me.CboHSub2CAM.Visible = False
'        CboDCta.Visible = True
'        CboDSubcta1.Visible = True
'        CboDSubcta2.Visible = True
'        CboHcta.Visible = True
'        CbohSubcta1.Visible = True
'        CbohSubcta2.Visible = True
'        Me.frame_moneda.Enabled = True
'        'TxtDBs = ""
'        'TxtDSus = ""
'        optbolivianos_Click
'    Case "CAM"
'       ' TxtDBs.Enabled = True
'       ' TxtDSus.Enabled = True
'        If adiciona = "S" Then
'          Me.frameCAM.Visible = True
'        Else
'          Me.frameCAM.Visible = False
'        End If
'        Me.optCAMNo.Value = False
'        Me.optCAMSi.Value = False
'        Me.DTPCAM.Visible = True
'        Me.txt_fecha.Visible = False
'        Me.txtcodsolicitud.Visible = False
'        Label26.Visible = False 'codigo solicitud
'        Me.dtc_codigo4.Text = "-"
'        Me.lblDTC = "1.0"
'        lblHTC = "1.0"
'        '----
'        dtc_codigo4.Visible = False
'        dtc_desc4.Visible = False
'        DtCHDescripbenef.Visible = False
'        DtCHcodbenef.Visible = False
'        lblDBenefaux1.Visible = True
'        lblDnomBenefaux1.Visible = True
'        lblHBenefaux1.Visible = True
'        lblHnomBenefaux1.Visible = True
'        '----
'        Me.lblDTC.Visible = False
'        Me.lblDTC.Locked = True
'        lblHTC.Visible = False
'        lblHTIPOCAM.Visible = False
'        lblDTIPOCAM.Visible = False
'        'lblDMonSus.Visible = False
'        'lblHMONSUS.Visible = False
'        'Me.txtHsus.Visible = False
'        'Me.TxtDSus.Visible = False
'        'Me.TxtDSus = "0.0"
'        'Me.txtHsus = "0.0"
'        CboDCta.Visible = False
'        CboDSubcta1.Visible = False
'        CboDSubcta2.Visible = False
'        CboHcta.Visible = False
'        CbohSubcta1.Visible = False
'        CbohSubcta2.Visible = False
'        Me.CboDCtaCAM.Visible = True
'        Me.CboDSub1CAM.Visible = True
'        Me.CboDSub2CAM.Visible = True
'        Me.CboHCtaCAM.Visible = True
'        Me.CboHSub1CAM.Visible = True
'        Me.CboHSub2CAM.Visible = True
'
'        'Me.frame_moneda.Enabled = False
'        'Me.optbolivianos = True
'        optbolivianos_Click
'  End Select
'  ' Dim rsbustipo As ADODB.Recordset
'  ' Set rsbustipo = New ADODB.Recordset
'
'  rstipocomp.Filter = adFilterNone
'    rstipocomp.Filter = "Codigo_Tipo='" & Trim(CboTipo.Text) & "'"
'    If rstipocomp.RecordCount <> 0 Then
'        cboNomTipo.Text = rstipocomp!Denominacion_Tipo
'    End If
'End Sub


'Private Sub CboTipo_KeyPress(KeyAscii As Integer)
' KeyAscii = 0
'End Sub
'
''Private Sub Cmbo_Atributo_Click()
''    If Me.Cmbo_Atributo.Text = "status" Then
''        Me.Cbostatus.Visible = True
''        Text_Valor.Visible = False
''    Else
''        Me.Cbostatus.Visible = False
''        Text_Valor.Visible = True
''    End If
''End Sub
'
'Private Sub cmd_aprob_aceptar_Click()
'
'Dim codigo_pago As Integer
'Dim aprobindiv As Integer
'Dim aprobcjto As Integer
'Dim rsctabancariaDebe As ADODB.Recordset
'Set rsctabancariaDebe = New ADODB.Recordset
'Dim rsctabancariaHaber As ADODB.Recordset
'Set rsctabancariaHaber = New ADODB.Recordset
'Dim rsctabanc As ADODB.Recordset
'Set rsctabanc = New ADODB.Recordset
'Set rspco = New ADODB.Recordset
'
'If optconjunto.Value = True Then
'    If (Me.cboaprob_inicio.Text = "" Or Me.cboaprob_inicio.ListIndex = -1) Or (Me.cbo_aprob_final.Text = "" Or Me.cbo_aprob_final.ListIndex = -1) Then
'        MsgBox "Elija los comprobantes a aprobar", vbExclamation + vbDefaultButton1, "APROBACION"
'        Exit Sub
'    End If
'End If
'If optindividual.Value = True Then
'    If Me.cboaprob_inicio.Text = "" Or cboaprob_inicio.ListIndex = -1 Then
'          MsgBox "Elija el comprobante a aprobar", vbExclamation + vbDefaultButton1, "APROBACION"
'          Exit Sub
'    End If
'End If
'Set rspago = New ADODB.Recordset
'Set rspago_detalle = New ADODB.Recordset
'If sw1 = 1 Then  'aprobacion individual
'        '********CAMBIO DE STATUS A APROBADO
'  aprobindiv = MsgBox("Está seguro de aprobar el comprobante: " & Trim(Me.cboaprob_inicio.Text), vbQuestion + vbYesNo)
'  If aprobindiv = 6 Then
'    db.BeginTrans
'    Set rs_datos_M = New ADODB.Recordset
'    If rs_datos_M.State = 1 Then rs_datos_M.Close
'    rs_datos_M.Open "select * from Co_Comprobante_M where cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockOptimistic
'    rs_datos_M.MoveFirst
'    If rs_datos_M!Status = "N" Then
'        rs_datos_M!Status = "S"
'        'rs_datos_M!Fecha_transacion = CDate(Format(Date, "dd/mm/yyyy"))
'        'rs_datos_M!Cod_Trans = Trim(Me.cboaprob_inicio.Text)
'        codigo_pago = Val(rs_datos_M!Cod_Comp)
'        rs_datos_M.Update
'    If rs_datos_M!tipo_comp = "CAM" Then
'            MsgBox "Aprobación con éxito", vbInformation + vbDefaultButton1, "Atencion"
'    End If
'    If rs_datos_M!tipo_comp = "RVT" Then
'          Dim rspag1 As ADODB.Recordset
'          Set rspag1 = New ADODB.Recordset
'          If rspag1.State = 1 Then rspag1.Close
'          rspag1.Open "select * from pagos where codigo_pago=" & Val(rs_datos_M!cod_trans) & " and  org_codigo='" & rs_datos_M!org_codigo & "'", db, adOpenKeyset, adLockOptimistic
'          If rspag1.RecordCount <> 0 Then
'            rspag1!nro_comprobante_anterior = rs_aux1!cod_trans
'            rspag1!tipo_formulario = "RVT"
'            rspag1!estado_contabilidad = "R"
'            rspag1!estado_aprobacion = "N"
'            rspag1!usr_usuario = GlUsuario
'            rspag1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'            rspag1!hora_registro = Format(Time, "hh:mm:ss")
'            rspag1.Update
'            MsgBox "Aprobación con éxito", vbInformation + vbDefaultButton1, "Atencion"
'          End If
'    End If
'        If rs_datos_M!tipo_comp = "ANL" Or rs_datos_M!tipo_comp = "DVL" Then
'          '****revisar g--!!!!!!!!!!!
'          Dim rsp As ADODB.Recordset
'          Dim rspadeta As ADODB.Recordset
'          Set rsp = New ADODB.Recordset
'          Set rspadeta = New ADODB.Recordset
'          If rsp.State = 1 Then rspa.Close
'          rsp.Open "select * from pagos where codigo_pago=" & rs_aux1!cod_trans & " and  org_codigo='" & rs_aux1!org_codigo & "'", db, adOpenKeyset, adLockOptimistic
'          If rsp.RecordCount <> 0 Then
'            rsp!nro_comprobante_anterior = rs_aux1!cod_trans
'            rsp!tipo_formulario = "ANL"
'            rsp!estado_pagado = "L"
'            rsp!estado_aprobacion = "N"
'            rsp!usr_usuario = GlUsuario
'            rsp!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'            rsp!hora_registro = Format(Time, "hh:mm:ss")
'            rsp.Update
'          If rspadeta.State = 1 Then rspadeta.Close
'          rspadeta.Open "select * from pago_detalle where codigo_pago=" & rs_aux1!cod_trans & " and org_codigo='" & rs_aux1!org_codigo & "'", db, adOpenKeyset, adLockOptimistic
'          If rspadeta.RecordCount <> o Then
'            rspadeta!estado_aprobacion = "N"
'            rspadeta!usr_usuario = GlUsuario
'            rspadeta!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'            rspadeta!hora_registro = Format(Time, "hh:mm:ss")
'            rspadeta.Update
'          End If
'          End If
'          Set rsdiario = New ADODB.Recordset
'          If rsdiario.State = 1 Then rsdiario.Close
'          rsdiario.CursorLocation = adUseClient
'          rsdiario.Open "select D_Cta_Aux1,d_montoBs from co_diario where cod_comp=" & Val(cboaprob_inicio), db, adOpenKeyset, adLockReadOnly
'          If rsdiario.RecordCount <> 0 Then
'              If rsctabanc.State = 1 Then rsctabancaria.Close
'              rsctabanc.CursorLocation = adUseClient
'              rsctabanc.Open "SELECT Cta_Codigo,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & Trim(rsdiario!D_Cta_Aux1) & "'", db, adOpenKeyset, adLockOptimistic
'              If rsctabanc.RecordCount <> 0 Then
'                rsctabanc!cta_acum_anl = IIf(IsNull(rsctabanc!cta_acum_anl), 0, rsctabanc!cta_acum_anl) + IIf(IsNull(rsdiario!d_montoBs), 0, rsdiario!d_montoBs)
'              End If
'              rsctabanc.Update
'          End If
'        End If
'        If rs_datos_M!tipo_comp = "ANC" Then
'            If rsdiario.State = 1 Then rsdiario.Close
'            rsdiario.Open "SELECT D_Cta_Aux1,H_Cta_Aux1,D_MontoBs FROM CO_Diario " & _
'                 "WHERE Cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockReadOnly
'           If rsdiario.RecordCount <> 0 Then
'
'            '****cta del Debe
'            ctacodigoDebe = rsdiario!H_Cta_Aux1
'            ctacodigoHaber = rsdiario!D_Cta_Aux1
'            If rsctabancariaDebe.State = 1 Then rsctabancariaDebe.Close
'            rsctabancariaDebe.CursorLocation = adUseClient
'            rsctabancariaDebe.Open "SELECT Cta_Codigo,Cta_Anl_TRP,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & ctacodigoDebe & "'", db, adOpenKeyset, adLockOptimistic
'            If rsctabancariaDebe.RecordCount <> 0 Then
'                 rsctabancariaDebe!cta_anl_TRP = IIf(IsNull(rsctabancariaDebe!cta_anl_TRP), 0, rsctabancariaDebe!cta_anl_TRP) + rsdiario!d_montoBs
'              rsctabancariaDebe.Update
'            End If
'            '****cta del haber
'            If rsctabancariaHaber.State = 1 Then rsctabancariaHaber.Close
'            rsctabancariaHaber.CursorLocation = adUseClient
'            rsctabancariaHaber.Open "SELECT Cta_Codigo,Cta_Anl_TRP,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & ctacodigoHaber & "'", db, adOpenKeyset, adLockOptimistic
'            If rsctabancariaHaber.RecordCount <> 0 Then
'              rsctabancariaHaber!cta_acum_anl = rsctabancariaHaber!cta_acum_anl + rsdiario!d_montoBs
'              rsctabancariaHaber.Update
'            End If
'            'Exit Sub
'           End If
'        End If
'
'        If rs_datos_M!tipo_comp = "PCE" Then
'            Set rsdiario = New ADODB.Recordset
'            If rsdiario.State = 1 Then rsdiario.Close
'            rsdiario.Open "SELECT * FROM CO_Diario " & _
'                 "WHERE Cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockReadOnly
'            Set rspago = New ADODB.Recordset
'            Set rspago_detalle = New ADODB.Recordset
'            If rspago.State = 1 Then rspago.Close
'            rspago.CursorLocation = adUseClient
'            rspago.Open "SELECT * FROM pagos WHERE (org_codigo = '999')  and codigo_pago=" & codigo_pago, db, adOpenKeyset, adLockOptimistic
'            '*********ADICION A LA TABLA PAGO
'            If rspago.RecordCount = 0 Then
'                rspago.AddNew
'            End If
'            rspago!ges_gestion = IIf(IsNull(Trim(rs_datos_M!ges_gestion)), "", Trim(rs_datos_M!ges_gestion))
'            rspago!org_codigo = "999"
'            rspago!codigo_pago = IIf(IsNull(rs_datos_M!Cod_Comp), "", Trim(rs_datos_M!Cod_Comp))
'            rspago!tipo_comp = IIf(IsNull(rs_datos_M!tipo_comp), "", Trim(rs_datos_M!tipo_comp))
'            rspago!Codigo_orden = IIf(IsNull(rs_datos_M!num_respaldo), "", Trim(rs_datos_M!num_respaldo))
'            rspago!codigo_documento = IIf(IsNull(rs_datos_M!codigo_documento), "", Trim(rs_datos_M!codigo_documento))
'            rspago!fecha_egreso = (Format(rs_datos_M!Fecha_transacion, "dd/mm/yyyy"))
'            rspago!monto_Bolivianos = Val(rsdiario!d_montoBs)
'            rspago!monto_dolares = Val(rsdiario!d_montoDl)
'            rspago!liquido_pagar = Val(rsdiario!d_montoBs)
'            rspago!estado_aprobacion = "N"
'            rspago!estado_contabilidad = "P"
'            'rspago!estado_devengado = "S"
'            rspago!estado_pagado = "N"
'            rspago!justificacion = IIf(IsNull(rs_datos_M!glosa), "", Trim(CStr(rs_datos_M!glosa)))
'            rspago!usr_usuario = GlUsuario  'IIf(IsNull(rs_datos_M!usr_usuario), "", Trim(rs_datos_M!usr_usuario))
'            rspago!fecha_aprueba = CDate(Format(CFecha, "dd/mm/yyyy"))
'            rspago!hora_aprueba = (Format(Time, "hh:mm:ss"))
'            rspago!fecha_registro = CDate(Format(CFecha, "dd/mm/yyyy"))
'            rspago!hora_registro = (Format(Time, "hh:mm:ss"))
'            rspago!codigo_solicitud = IIf(IsNull(rs_datos_M!codigo_solicitud), "", Trim(rs_datos_M!codigo_solicitud))
'            '********ADICION A LA TABLA PAGO DETALLE
'            If rspago_detalle.State = 1 Then rspago_detalle.Close
'            rspago_detalle.CursorLocation = adUseClient
'            rspago_detalle.Open "SELECT * FROM pago_detalle WHERE  (org_codigo = '999')  and codigo_pago=" & codigo_pago, db, adOpenKeyset, adLockOptimistic
'            If rspago_detalle.RecordCount = 0 Then
'            rspago_detalle.AddNew
'            End If
'            'rspago_detalle.AddNew
'            rspago_detalle!ges_gestion = IIf(IsNull(Trim(rs_datos_M!ges_gestion)), "", Trim(rs_datos_M!ges_gestion))
'            rspago_detalle!org_codigo = "999"
'            rspago_detalle!codigo_pago = Val(Trim(rs_datos_M!Cod_Comp))
'            rspago_detalle!codigo_pago_detalle = "1"
'            rspago_detalle!beneficiario_codigo = IIf(IsNull(rs_datos_M!beneficiario_codigo), "", Trim(rs_datos_M!beneficiario_codigo))
'            rspago_detalle!tipo_cambio = Val(rsdiario!d_Cambio)
'            rspago_detalle!monto_total = Val(rsdiario!d_montoBs)
'            rspago_detalle!departamento = "La Paz"
'            rspago_detalle!honorarios = "N"
'            rspago_detalle!tipo_cambio = Val(rsdiario!d_Cambio)
'            rspago_detalle!estado_aprobacion = "N"
'            rspago_detalle!monto_Bolivianos = Val(rsdiario!d_montoBs)
'            rspago_detalle!monto_dolares = Val(rsdiario!d_montoDl)
'            rspago_detalle!fecha_pago = CDate(Format(CFecha, "dd/mm/yyyy"))
'            rspago_detalle!usr_usuario = GlUsuario 'IIf(IsNull(rs_datos_M!usr_usuario), "", Trim(rs_datos_M!usr_usuario))
'            rspago_detalle!fecha_registro = Format(CFecha, "dd/mm/yyyy")
'            rspago_detalle!hora_registro = Format(Time, "hh:mm:ss")
'            rspago.Update
'            rspago_detalle.Update
'            'db.CommitTrans
'            MsgBox "Aprobación con éxito", vbInformation + vbDefaultButton1, "Atencion"
'        End If
'            '*****TIPO COMPROBANTE PCO
'
'        If rs_datos_M!tipo_comp = "PCO" Then
'         '*****CREAR DOS REGISTROS PCO
'            Set rsdiario = New ADODB.Recordset
'            If rsdiario.State = 1 Then rsdiario.Close
'            rsdiario.Open "SELECT * FROM CO_Diario " & _
'                 "WHERE Cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockReadOnly
'
''g-
'            If (rsdiario!d_cuenta = "1121" And rsdiario!d_subcta1 = "02") And (rsdiario!h_cuenta = "2116" And rsdiario!h_subcta1 = "04") And (rsdiario!tipo_comp = "PCO") Or ((rsdiario!d_cuenta = "2116" And rsdiario!d_subcta1 = "04") And (rsdiario!h_cuenta = "1121" And rsdiario!h_subcta1 = "02") And (rsdiario!tipo_comp = "PCO")) Then
'              Dim sqlx As String
'              sqlx = "update co_diario set H_Cta_Aux3 = D_Cta_Aux2 , d_ctaaux3 = H_Cta_Aux2 WHERE COD_COMP =" & Val(Trim(Me.cboaprob_inicio.Text))
'              db.Execute sqlx
'            End If
''g-
'
'            If (rsdiario!d_cuenta = "1111" And rsdiario!d_subcta1 = "02") And (rsdiario!h_cuenta = "1111" And rsdiario!h_subcta1 = "02") Then
'                Call PCO(Trim(rsdiario!d_cuenta), "D", Val(rs_datos_M!Cod_Comp))
'                Call PCO(Trim(rsdiario!h_cuenta), "H", Val(rs_datos_M!Cod_Comp))
'            Else
'                If (rsdiario!d_cuenta = "1111" And rsdiario!d_subcta1 = "02") Or (rsdiario!h_cuenta = "1111" And rsdiario!h_subcta1 = "02") Then
'                    If (rsdiario!d_cuenta = "1111" And rsdiario!d_subcta1 = "02") Then
'                        Call PCO(Trim(rsdiario!d_cuenta), "D", Val(rs_datos_M!Cod_Comp))
'                    End If
'                    If (rsdiario!h_cuenta = "1111" And rsdiario!h_subcta1 = "02") Then
'                        Call PCO(Trim(rsdiario!h_cuenta), "H", Val(rs_datos_M!Cod_Comp))
'                    End If
'                End If
'            End If
'          MsgBox "Aprobación con éxito", vbInformation + vbDefaultButton1, "Atencion"
'        End If
'
'    Else '*******estado comprobante
'        MsgBox "El comprobante " & Trim(Me.cboaprob_inicio) & " ya está aprobado", vbExclamation + vbDefaultButton1
'        Me.cboaprob_inicio.SetFocus
'        Exit Sub
'    End If
'  Else
'   Exit Sub
'  End If
'  db.CommitTrans
'Else '***del swich
'    If sw1 = 0 And (Val(Trim(Me.cboaprob_inicio.Text)) < Val(Trim(Me.cbo_aprob_final.Text))) Then
'
'        Set rs_datos_M = New ADODB.Recordset
'        If rs_datos_M.State = 1 Then rs_datos_M.Close
'        rs_datos_M.Open " Select * from co_comprobante_M where cod_comp between " & Val(Me.cboaprob_inicio.Text) & " and " & Val(Me.cbo_aprob_final.Text) & " and status='N'", db, adOpenKeyset, adLockOptimistic
'        rs_datos_M.Sort = "cod_comp"
'        Me.lstcomprobantes.Clear
'        Do While Not rs_datos_M.EOF
'            Me.lstcomprobantes.AddItem Str(rs_datos_M!Cod_Comp) + " " + rs_datos_M!tipo_comp
'            rs_datos_M.MoveNext
'        Loop
'        Me.Framecomprobantes.Visible = True
'        Me.Framecomprobantes.Enabled = True
'        aprobcjto = MsgBox("Está seguro ???", vbQuestion + vbYesNo)
'        If aprobcjto = 6 Then
'            db.BeginTrans
'        'MsgBox rs_datos_M.RecordCount
'            Set rsdiario = New ADODB.Recordset
'            If rsdiario.State = 1 Then rsdiario.Close
'            rsdiario.Open " select * from Co_Diario where cod_comp between " & Val(Me.cboaprob_inicio.Text) & " and " & Val(Me.cbo_aprob_final.Text), db, adOpenKeyset, adLockReadOnly
'            rs_datos_M.MoveFirst
'            For i = Val(Trim(Me.cboaprob_inicio)) To Val(Trim(Me.cbo_aprob_final))
'
'                rs_datos_M.Filter = adFilterNone
'                rs_datos_M.Filter = "cod_comp=" & i
'                'MsgBox rs_datos_M.RecordCount
'              '********CAMBIO DE STATUS A APROBADO
'                'rs_datos_M.MoveFirst
'                If rs_datos_M.RecordCount <> 0 Then
'                  If rs_datos_M!Status = "N" Then
'                    rs_datos_M!Status = "S"
'                    'rs_datos_M!Fecha_transacion = CDate(Format(CFecha, "dd/mm/yyyy"))
'                    'rs_datos_M!Cod_Trans = Trim(Me.cboaprob_inicio.Text)
'                    codigo_pago = rs_datos_M!Cod_Comp
'                    rs_datos_M.Update
'                    rsdiario.MoveFirst
'                    'rsdiario.Filter = adFilterNone
'                    'rsdiario.Filter = "cod_comp=" & i
'                        'rsdiario.Find "cod_comp=" & i
'                        'Set rspago = New ADODB.Recordset
'                    rs_datos_M.Filter = adFilterNone
'                    rs_datos_M.Filter = "cod_comp=" & i
'                    '********
'                    If rs_datos_M!tipo_comp = "ANC" Then
'            If rsdiario.State = 1 Then rsdiario.Close
'            rsdiario.Open "SELECT D_Cta_Aux1,H_Cta_Aux1,D_MontoBs FROM CO_Diario " & _
'                 "WHERE Cod_comp=" & i, db, adOpenKeyset, adLockReadOnly
'           If rsdiario.RecordCount <> 0 Then
'
'            '****cta del Debe
'            ctacodigoDebe = rsdiario!H_Cta_Aux1
'            ctacodigoHaber = rsdiario!D_Cta_Aux1
'            If rsctabancariaDebe.State = 1 Then rsctabancariaDebe.Close
'            rsctabancariaDebe.CursorLocation = adUseClient
'            rsctabancariaDebe.Open "SELECT Cta_Codigo,Cta_Anl_TRP,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & ctacodigoDebe & "'", db, adOpenKeyset, adLockOptimistic
'            If rsctabancariaDebe.RecordCount <> 0 Then
'                 rsctabancariaDebe!cta_anl_TRP = IIf(IsNull(rsctabancariaDebe!cta_anl_TRP), 0, rsctabancariaDebe!cta_anl_TRP) + rsdiario!d_montoBs
'              rsctabancariaDebe.Update
'            End If
'            '****cta del haber
'            If rsctabancariaHaber.State = 1 Then rsctabancariaHaber.Close
'            rsctabancariaHaber.CursorLocation = adUseClient
'            rsctabancariaHaber.Open "SELECT Cta_Codigo,Cta_Anl_TRP,CTA_ACUM_ANL from fc_cuenta_bancaria where cta_codigo='" & ctacodigoHaber & "'", db, adOpenKeyset, adLockOptimistic
'            If rsctabancariaHaber.RecordCount <> 0 Then
'              rsctabancariaHaber!cta_acum_anl = rsctabancariaHaber!cta_acum_anl + rsdiario!d_montoBs
'              rsctabancariaHaber.Update
'            End If
'            'Exit Sub
'           End If
'        End If
'        '****
'
'
'
'                    If rs_datos_M!tipo_comp = "PCE" Then
'                        rsdiario.Filter = adFilterNone
'                        rsdiario.Filter = "cod_comp=" & i
'                        If rspago.State = 1 Then rspago.Close
'                        rspago.CursorLocation = adUseClient
'                        rspago.Open "SELECT * FROM pagos where (org_codigo = '999') and codigo_pago=" & codigo_pago, db, adOpenKeyset, adLockOptimistic
'                        'Set rspago_detalle = New ADODB.Recordset
'                      '*********ADICION A LA TABLA PAGO
'                        If rspago.RecordCount = 0 Then
'                            rspago.AddNew
'                        End If
'                        rspago!ges_gestion = IIf(IsNull(rs_datos_M!ges_gestion), "", Trim(rs_datos_M!ges_gestion))
'                        rspago!org_codigo = "999"
'                        rspago!codigo_pago = IIf(IsNull(rs_datos_M!Cod_Comp), 0, rs_datos_M!Cod_Comp)
'                        '.rspago!nro_comprobante_anterior = .rs_datos!Cod_Comp
'                        rspago!tipo_comp = "PCE"
'                        rspago!Codigo_orden = IIf(IsNull(rs_datos_M!num_respaldo), "", Trim(rs_datos_M!num_respaldo))
'                        rspago!codigo_documento = IIf(IsNull(rs_datos_M!codigo_documento), "", Trim(rs_datos_M!codigo_documento))
'                        rspago!fecha_egreso = (Format(rs_datos_M!Fecha_transacion, "dd/mm/yyyy"))
'                        rspago!monto_Bolivianos = Val(rsdiario!d_montoBs)
'                        rspago!monto_dolares = Val(rsdiario!d_montoDl)
'                        rspago!liquido_pagar = Val(rsdiario!d_montoBs)
'                        'celia rspago!estado_aprobacion = "N" o "A"
'                        rspago!estado_aprobacion = "N"
'                        rspago!estado_contabilidad = "P"
'                        'Rspago!estado_devengado = "S"
'                        rspago!estado_pagado = "N"
'                        rspago!justificacion = IIf(IsNull(rs_datos_M!glosa), "", Trim(rs_datos_M!glosa))
'                        rspago!usr_usuario = IIf(IsNull(rs_datos_M!usr_usuario), "", Trim(rs_datos_M!usr_usuario))
'                        rspago!fecha_aprueba = Format(CFecha, "dd/mm/yyyy")
'                        rspago!hora_aprueba = (Format(Time, "hh:mm:ss"))
'                        rspago!fecha_registro = Format(CFecha, "dd/mm/yyyy")
'                        rspago!hora_registro = (Format(Time, "hh:mm:ss"))
'                        '********ADICION A LA TABLA PAGO DETALLE
'                        If rspago_detalle.State = 1 Then rspago_detalle.Close
'                        rspago_detalle.CursorLocation = adUseClient
'                        rspago_detalle.Open "SELECT * FROM pago_detalle where (org_codigo = '999')  and codigo_pago=" & codigo_pago, db, adOpenKeyset, adLockOptimistic
'                        If rspago_detalle.RecordCount = 0 Then
'                           rspago_detalle.AddNew
'                        End If
'                        rspago_detalle!ges_gestion = IIf(IsNull(rs_datos_M!ges_gestion), "", Trim(rs_datos_M!ges_gestion))
'                        rspago_detalle!org_codigo = "999"
'                        rspago_detalle!codigo_pago = IIf(IsNull(rs_datos_M!Cod_Comp), 0, rs_datos_M!Cod_Comp)
'                        rspago_detalle!codigo_pago_detalle = "1"
'                        rspago_detalle!beneficiario_codigo = IIf(IsNull(rs_datos_M!beneficiario_codigo), "", Trim(rs_datos_M!beneficiario_codigo))
'                        rspago_detalle!tipo_cambio = Val(rsdiario!d_Cambio)
'                        rspago_detalle!monto_total = Val(rsdiario!d_montoBs)
'                        rspago_detalle!departamento = "La Paz"
'                        rspago_detalle!honorarios = "N"
'                        ''''''''''''
'                        rspago_detalle!tipo_cambio = Val(rsdiario!d_Cambio)
'                        rspago_detalle!estado_aprobacion = "N"
'                        rspago_detalle!monto_Bolivianos = Val(rsdiario!d_montoBs)
'                        rspago_detalle!monto_dolares = Val(rsdiario!d_montoDl)
'                        rspago_detalle!fecha_pago = Format(CFecha, "dd/mm/yyyy")
'                        rspago_detalle!usr_usuario = IIf(IsNull(rs_datos_M!usr_usuario), "", Trim(rs_datos_M!usr_usuario))
'                        rspago_detalle!fecha_registro = Format(CFecha, "dd/mm/yyyy")
'                        rspago_detalle!hora_registro = Format(Time, "hh:mm:ss")
'                        rspago.Update
'                        rspago_detalle.Update
'                    End If
'                    '****TIPÖ COMPROBANTE PCO
'                    If rs_datos_M!tipo_comp = "PCO" Then
'                      If (rsdiario!d_cuenta = "1121" And rsdiario!d_subcta1 = "02") And (rsdiario!h_cuenta = "2116" And rsdiario!h_subcta1 = "04") And (rsdiario!tipo_comp = "PCO") Or ((rsdiario!d_cuenta = "2116" And rsdiario!d_subcta1 = "04") And (rsdiario!h_cuenta = "1121" And rsdiario!h_subcta1 = "02") And (rsdiario!tipo_comp = "PCO")) Then
'                        Dim sqlx1 As String
'                        sqlx1 = "update co_diario set H_Cta_Aux3 = D_Cta_Aux2 , d_ctaaux3 = H_Cta_Aux2 WHERE COD_COMP =" & Val(Trim(i))
'                        db.Execute sqlx1
'                      End If
'                    '*****CREAR DOS REGISTROS PCO
'                        Set rsdiario = New ADODB.Recordset
'                        If rsdiario.State = 1 Then rsdiario.Close
'                        rsdiario.Open "SELECT * FROM CO_Diario " & _
'                                "WHERE Cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockReadOnly
'                        If (rsdiario!d_cuenta = "1111" And rsdiario!d_subcta1 = "02") And (rsdiario!h_cuenta = "1111" And rsdiario!h_subcta1 = "02") Then
'                            Call PCO(Trim(rsdiario!d_cuenta), "D", Val(rs_datos_M!Cod_Comp))
'                            Call PCO(Trim(rsdiario!h_cuenta), "H", Val(rs_datos_M!Cod_Comp))
'                        Else
'                            If (rsdiario!d_cuenta = "1111" And rsdiario!d_subcta1 = "02") Or (rsdiario!h_cuenta = "1111" And rsdiario!h_subcta1 = "02") Then
'                                If (rsdiario!d_cuenta = "1111" And rsdiario!d_subcta1 = "02") Then
'                                    Call PCO(Trim(rsdiario!d_cuenta), "D", Val(rs_datos_M!Cod_Comp))
'                                End If
'                                If (rsdiario!h_cuenta = "1111" And rsdiario!h_subcta1 = "02") Then
'                                    Call PCO(Trim(rsdiario!h_cuenta), "H", Val(rs_datos_M!Cod_Comp))
'                                End If
'                            End If
'                       End If
'                    End If
''          MsgBox "Aprobación con éxito", vbInformation + vbDefaultButton1, "Atencion"
'
'          Else '******* si esta aprobado
'                   MsgBox " El comprobante " & i & "ya está aprobado", vbExclamation + vbDefaultButton1
'                End If
'        End If
'        Next
'        db.CommitTrans
'        MsgBox "Aprobación con éxito", vbInformation + vbDefaultButton1
'        Framecomprobantes.Visible = False
'  Else
'        Me.Framecomprobantes.Visible = False
'        Exit Sub
'  End If
'Else
'    MsgBox "Introduzca nuevamente el rango", vbExclamation + vbDefaultButton1, "Atencion"
'    Exit Sub
'End If
'End If ' del sw
'
'        Me.fraOpciones.Enabled = True
'        Me.cbo_aprob_final.Clear
'        Me.cboaprob_inicio.Clear
'        rs_datos.Requery
'        'MsgBox queryinicial
'        rs_datos.Filter = adFilterNone
'        rs_datos.Filter = "status='N'"
'        Set Me.dg_datos.DataSource = Nothing
'          If rs_datos.RecordCount <> 0 Then
'          Do While Not rs_datos.EOF
'            Me.cboaprob_inicio.AddItem Trim(rs_datos!Cod_Comp)
'            Me.cbo_aprob_final.AddItem Trim(rs_datos!Cod_Comp)
'            'g-
''            If rs_datos!Cod_Comp <> "PCE" Then MsgBox rs_datos!Cod_Comp
'            rs_datos.MoveNext
'          Loop
'        End If
'          'rs_datos.Filter = adFilterNone
'        'Set Me.dg_datos.DataSource = rs_datos
'
'End Sub
'
'Private Sub cmd_aprob_cancel_Click()
'    Me.fraOpciones.Enabled = True
'    Me.FraNavega.Enabled = True
'    Me.Frame_aprobacion.Visible = False
'    rs_datos.Requery
'    rs_datos.Filter = adFilterNone
'    Set Me.dg_datos.DataSource = rs_datos
'End Sub
'
'Private Sub BtnAprobar_Click()
''Me.Cmbo_Atributo = Clear
''With dtetraspasos
''If .rs_datos.State = 1 Then .rs_datos.Close
'    Me.fraOpciones.Enabled = False
'    Me.FraNavega.Enabled = False
'    Me.cbo_aprob_final.Clear
'    Me.cboaprob_inicio.Clear
'    rs_datos.Filter = adFilterNone
'    rs_datos.Filter = "status ='N'"
''.rs_datos.Open
'    Set Me.dg_datos.DataSource = Nothing
'    If rs_datos.RecordCount <> 0 Then
'     'rs_datos.MoveFirst
'        'For i = 0 To rs_datos.RecordCount
'        Do While Not rs_datos.EOF
'          Select Case rs_datos!tipo_comp
'            Case "PCE", "PCO", "CAM", "RVT"
'              Me.cboaprob_inicio.AddItem rs_datos!Cod_Comp
'              Me.cbo_aprob_final.AddItem rs_datos!Cod_Comp
'        '     aprobacion(i) = rs_datos!Cod_Comp
'          End Select
'            rs_datos.MoveNext
'        'Next
'        Loop
'        cmd_aprob_aceptar_Click        'Me.Frame_aprobacion.Visible = True
'    Else
'        MsgBox "No existen comprobantes para aprobar", vbExclamation + vbDefaultButton1
'    End If
'End Sub
'
'Private Sub Cmd_BSalir_Click()
'    Me.fraOpciones.Enabled = True
'    Me.FraNavega.Enabled = True
'    Set Me.dg_datos.DataSource = rs_datos
'    Me.dg_datos.Refresh
'  '  Me.Fra_Busqueda.Visible = False
'    Me.OptTodos.Value = False
'   Me.OptSinAprobar.Value = False
'End Sub
'
''Private Sub Cmd_Cancelar_Click()
'''With dtetraspasos
''Me.FraGlobal.Enabled = False
''Me.Fram_AsientoD.Enabled = False
''Me.Fram_AsientoH.Enabled = False
''   rs_datos.Filter = adFilterNone
''Set Me.dg_datos.DataSource = rs_datos
''  Me.dg_datos.Refresh
'''End With
''  Call limpiar
''  Me.Cmd_GrabaM.Enabled = False
''  Me.BtnSalir.Enabled = True
''  Me.Cmd_Modificar.Enabled = True
''  Me.BtnAñadir.Enabled = True
''  Me.BtnAprobar.Enabled = True
''  Me.BtnBuscar.Enabled = True
''  Me.BtnDesAprobar.Enabled = True
''  Me.Cmd_Eligir.Enabled = True
''  Me.BtnImprimir.Enabled = True
''  Me.dg_datos.Enabled = True
''  Me.frame_moneda.Visible = False
''  'Me.FraGlobal.Enabled = True
''  'Me.Fram_AsientoD.Enabled = True
''  'Me.Fram_AsientoH.Enabled = True
''
''End Sub
'
'Private Sub BtnDesAprobar_Click()
'    cmodificar = "C"
'    BtnGrabar_Click
'    frame_moneda.Enabled = True
'End Sub
''Private Sub cmd_Ejecutar_Click()
''   opttodos_Click
''   rs_datos.Filter = adFilterNone
''   Select Case Cmbo_Atributo.Text
''     Case "Cod_Comp"
''            Select Case Me.Cmbo_Operador.Text
''                Case "="
''                    rs_datos.Filter = "cod_comp =" & Val(Me.Text_Valor)
''                Case ">"
''                    rs_datos.Filter = "cod_comp >" & Val(Me.Text_Valor)
''                Case "<"
''                    rs_datos.Filter = "cod_comp <" & Val(Me.Text_Valor)
''                Case "<="
''                    rs_datos.Filter = "cod_comp <=" & Val(Me.Text_Valor)
''                Case ">="
''                    rs_datos.Filter = "cod_comp >=" & Val(Me.Text_Valor)
''             End Select
''         'Set Me.dg_datos.DataSource = rs_datos
''     Case "beneficiario_codigo"
''        Select Case Me.Cmbo_Operador.Text
''            Case "="
''              rs_datos.Filter = "beneficiario_codigo=" & Trim(Me.Text_Valor)
''            Case ">", "<", "<=", ">="
''              rs_datos.Filter = "beneficiario_codigo >" & Trim(Me.Text_Valor)
''        End Select
''        'Set Me.dg_datos.DataSource = rs_datos
''    Case "cod_trans"
''        Select Case Me.Cmbo_Operador.Text
''            Case "="
''                rs_datos.Filter = "cod_trans =" & Val(Me.Text_Valor)
''            Case ">"
''                rs_datos.Filter = "cod_trans  >" & Val(Me.Text_Valor)
''            Case "<"
''                rs_datos.Filter = "cod_trans  <" & Val(Me.Text_Valor)
''            Case "<="
''                rs_datos.Filter = "cod_trans  <=" & Val(Me.Text_Valor)
''            Case ">="
''                rs_datos.Filter = "cod_trans  >=" & Val(Me.Text_Valor)
''        End Select
''        'Set Me.dg_datos.DataSource = rs_datos
''    Case "org_codigo"
''        Select Case Me.Cmbo_Operador.Text
''            Case "="
''                rs_datos.Filter = "org_codigo='" & Trim(Me.Text_Valor) & "'"
''            Case Else
''                rs_datos.Filter = "org_codigo='" & Trim(Me.Text_Valor) & "'"
''        End Select
''        'Set Me.dg_datos.DataSource = rs_datos
'' Case "tipo_comp"
''        Select Case Me.Cmbo_Operador.Text
''            Case "="
''                rs_datos.Filter = "tipo_comp='" & Trim(Me.Text_Valor) & "'"
''            Case Else
''                rs_datos.Filter = "tipo_comp='" & Trim(Me.Text_Valor) & "'"
''        End Select
''        'Set Me.dg_datos.DataSource = rs_datos
'' Case "status"
''        Select Case Me.Cmbo_Operador.Text
''            Case "="
''                rs_datos.Filter = "status='" & Trim(Me.Cbostatus) & "'"
''            Case Else
''                rs_datos.Filter = "status='" & Trim(Me.Text_Valor) & "'"
''        End Select
'' End Select
''
''If rs_datos.RecordCount = 0 Then
''  MsgBox "No existe ese registro", vbExclamation, "Atencion"
''  rs_datos.Filter = adFilterNone
''  Set Me.dg_datos.DataSource = rs_datos
''  Me.dg_datos.Refresh
''  Me.fraOpciones.Enabled = False
''  Me.FraNavega.Enabled = False
''End If
''    Set Me.dg_datos.DataSource = rs_datos
''    Me.dg_datos.Refresh
''    rs_datos.MoveFirst
''    dg_datos_Click
''End Sub

Private Sub BtnImprimir_Click()
    
    Dim recsetaux As ADODB.Recordset
    Dim literales As String
    Dim decimal2 As String
    Dim LiteralCry As String
    Monto = 0
'    db.Execute "UPDATE co_diario SET NOMCTADEBE = (SELECT CC_Plan_Cuentas.NombreCta From CC_Plan_Cuentas Where CC_Plan_Cuentas.Cuenta =  co_diario.d_Cuenta and CC_Plan_Cuentas.SubCta1 = co_diario.d_Subcta1 and CC_Plan_Cuentas.SubCta2 = '00')"
    
    Set recsetaux = New ADODB.Recordset
    If rs_datos.RecordCount <> 0 Then
          If recsetaux.State = 1 Then recsetaux.Close
          recsetaux.Open "SELECT DISTINCT Co_Comprobante_M.Cod_Comp," & _
                       "Co_Comprobante_M.Tipo_Comp,CO_Diario.D_MontoBs " & _
                       "FROM Co_Comprobante_M INNER JOIN CO_Diario ON " & _
                       "Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp " & _
                       "WHERE Co_Comprobante_M.Cod_Comp = " & Val(rs_datos!Cod_Comp), db, adOpenForwardOnly, adLockReadOnly

        If recsetaux.RecordCount <> 0 Then
            Set rs_aux1 = New ADODB.Recordset
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select sum(d_montoBs) as totbs, sum(D_MontoDl) as totdl from co_diario where Cod_Comp = " & Ado_datos.Recordset!Cod_Comp & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux1.RecordCount > 0 Then
                LiteralCry = Str(rs_aux1!totbs)
                literales = Literal(Str(rs_aux1!totbs)) + " Bolivianos"
                db.Execute "Update Co_Comprobante_M Set literal = '" & literales & "'  Where Cod_Comp = " & Ado_datos.Recordset!Cod_Comp & "  "
            Else
                literales = "CERO 00/100 Bolivianos"
            End If

            Do While Not recsetaux.EOF
            LiteralCry = Str(Int(recsetaux!d_montoBs))
                Monto = CDbl(Monto) + recsetaux!d_montoBs
                recsetaux.MoveNext
            Loop
            LiteralCry = Str(Int(Monto))
            recsetaux.MoveFirst
            decimal2 = Str(Round((recsetaux!d_montoBs - Val(LiteralCry)), 2))
            If Monto <> 0 Then
                literales = Literal(Str(Monto)) + " Bolivianos"

            Else
                literales = "CERO 00/100 Bolivianos"
            End If
            Dim iResult As Integer
            CryComp_Manual.Destination = crptToWindow
            CryComp_Manual.WindowState = crptMaximized
            CryComp_Manual.WindowShowPrintSetupBtn = True
            CryComp_Manual.WindowShowRefreshBtn = True
            If GlBaseDatos = "ADMIN_EMPRESA" Then
                CryComp_Manual.ReportFileName = App.Path & "\reportes\Contabilidad\cr_registro_diario.rpt"
            Else
                CryComp_Manual.ReportFileName = App.Path & "\reportes\Contabilidad\cr_registro_diario_prueba.rpt"
            End If
            CryComp_Manual.StoredProcParam(0) = recsetaux!Cod_Comp
            CryComp_Manual.StoredProcParam(1) = recsetaux!tipo_comp
            'CryComp_Manual.StoredProcParam(2) = "g--"
            CryComp_Manual.StoredProcParam(2) = literales
            VAR_TIT = "MODULO CONTABILIDAD"
            CryComp_Manual.Formulas(0) = "titulo = '" & dtc_desc14.Text & "' "
            CryComp_Manual.Formulas(1) = "titulo1 = '" & VAR_TIT & "' "
            '
            iResult = CryComp_Manual.PrintReport
            If iResult <> 0 Then
                   MsgBox CryComp_Manual.LastErrorNumber & " : " & CryComp_Manual.LastErrorString, vbExclamation + vbOKOnly, "Error..."
            End If
       End If
    Else

       Exit Sub
    End If
End Sub

''Private Sub cmdanterior_Click()
''If rs_datos.RecordCount = 0 Then
''  Exit Sub
''End If
''    rs_datos.MovePrevious
''
''If rs_datos.BOF Then
''    rs_datos.MoveFirst
''    dg_datos_Click
''Else
'''    rs_datos.MovePrevious
''    dg_datos_Click
''End If
''End Sub
'
Private Sub BtnEliminar_Click()
On Error GoTo UpdateErr
   If ExisteReg(rs_datos!Cod_Comp) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención":
   If Ado_datos.Recordset!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datos.Recordset!estado_codigo = "ANL"
         Ado_datos.Recordset!Fecha_transacion = Date
         Ado_datos.Recordset!usr_codigo = glusuario
         Ado_datos.Recordset.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub
Private Function ExisteReg(cuenta2 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM co_diario WHERE Cod_Comp = '" & cuenta2 & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

'Private Sub BtnEliminar_Click()
'Dim opt As Integer
'Dim rsanular As ADODB.Recordset
'Set rsanular = New ADODB.Recordset
'rsanular.Open "select status from co_comprobante_M  where cod_comp= " & Val(rs_datos!Cod_Comp), db, adOpenKeyset, adLockOptimistic
'opt = MsgBox("Está seguro de anular el comprobante " & Trim(rs_datos!Cod_Comp) & "  " & Trim(rs_datos!tipo_comp), vbExclamation + vbYesNo)
'If opt = vbYes Then
'    'If rsanular.RecordCount <> 0 Then
'     '   rsanular!Status = "E"
'     '   rsanular.Update
'        db.Execute "update co_comprobante_M set status='E' where cod_comp=" & Val(rs_datos!Cod_Comp)
'        rs_datos.Requery
'        Set Me.dg_datos.DataSource = rs_datos
'    'End If
'Else
'    rsanular.Close
'    Exit Sub
'End If
'End Sub

''Private Sub BtnEliminar_Click()
''Dim opt As Integer
''Dim rsanular As ADODB.Recordset
''Set rsanular = New ADODB.Recordset
''rsanular.Open "select status from co_comprobante_M  where cod_comp= " & Val(rs_datos!Cod_Comp), db, adOpenKeyset, adLockOptimistic
''opt = MsgBox("Está seguro de anular el comprobante " & Trim(rs_datos!Cod_Comp) & "  " & Trim(rs_datos!tipo_comp), vbExclamation + vbYesNo)
''If opt = vbYes Then
''    If rsanular.RecordCount <> 0 Then
''        rsanular!Status = "E"
''        rsanular.Update
''        rs_datos.Requery
''        Set Me.dg_datos.DataSource = rs_datos
''    End If
''Else
''    rsanular.Close
''    Exit Sub
''End If
''End Sub

''Private Sub cmdfinal_Click()
''If rs_datos.RecordCount = 0 Then
''  Exit Sub
''End If
''If rs_datos.EOF Then
''    rs_datos.MovePrevious
''    dg_datos_Click
''Else
''    rs_datos.MoveLast
''    dg_datos_Click
''End If
''End Sub

Private Sub BtnModificar_Click()
    On Error GoTo EditErr
    VAR_SW = "MOD"
    VAR_BUS = 1
    cmodificar = "M"
    adiciona = "M"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
            Me.FrmABMDet1.Visible = False
            Me.FraDet1.Visible = False
            Me.FraDet2.Visible = False
            Me.FraNavega.Enabled = False
            Me.FraGlobal.Enabled = True
            Me.dg_datos.Enabled = False
            Me.fraOpciones.Visible = False
            Me.FraGrabarCancelar.Visible = True
            dtc_desc14.Locked = True
            Text7.Visible = True
            'DTPfecha_trans.Enabled = False
    Else
            MsgBox "No se puede MODIFICAR un registro APROBADO o Errado ...", vbExclamation, "Validación de Registro"
    End If
'    CboTipo.Enabled = False
'    cboNomTipo.Enabled = False
   Exit Sub
EditErr:
MsgBox Err.Description
End Sub

Private Sub BtnAñadir_Click()
On Error GoTo AddErr
    VAR_SW = "ADD"
    VAR_BUS = 1
    Call limpiar
    Call OptSinAprobar_Click
    rs_datos.AddNew
    '********
    Me.FraGrabarCancelar.Visible = True
    Me.fraOpciones.Visible = False
    Me.FraGlobal.Enabled = True
    Me.FraNavega.Enabled = False
    Me.dg_datos.Enabled = False
    Me.FrmABMDet1.Visible = False
    Me.FraDet1.Visible = False
    Me.FraDet2.Visible = False
    dtc_desc14.Locked = False
    Text7.Visible = False
    'DTPfecha_trans.Enabled = True
    cmodificar = "N"   'comprobante nuevo
    adiciona = "S"
    dtc_codigo1.Text = "DCONT"
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    CboTipo.Text = "PCO"
    cboNomTipo.BoundText = CboTipo.BoundText
    Txt_glosa.SetFocus
'    'gc_proceso_nivel3
'    Set rs_datos7 = New ADODB.Recordset
'    If rs_datos7.State = 1 Then rs_datos7.Close
'    rs_datos7.Open "Select * from gc_proceso_nivel3 WHERE (proceso_codigo = 'FIN') order by etapa_descripcion", db, adOpenStatic
'    'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
'    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'    'CboTipo.Enabled = True
'    'cboNomTipo.Enabled = True
Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnCancelar_Click()
On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelBatch

        Call OptSinAprobar_Click

'        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        FraGlobal.Enabled = False
        FraNavega.Enabled = True
        FrmABMDet1.Visible = True
        FraDet1.Visible = True
        FraDet2.Visible = True
        dg_datos.Enabled = True
        
        FraDet1.Visible = False
        FraDet2.Visible = False
        
        VAR_SW = ""
    End If
     Exit Sub
End Sub

Private Sub BtnGrabar_Click()
'On Error GoTo err3
'GRABAR MEVM INICIO
 On Error GoTo UpdateErr
VAR_VAL = "OK"
Call valida_campos
If VAR_VAL = "OK" Then
        If adiciona = "S" Then
          num_comprobante = 0
          Call genera_codigo
          'R-112, R-110, R-111
          Set rs_aux2 = New ADODB.Recordset
          SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & txt_codigo1 & "' "             'R-112' "
          rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
          If rs_aux2.RecordCount > 0 Then
                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                'txt_campo1.Caption = rs_aux2!correl_doc
                rs_aux2.Update
          End If
                                                                                                                         
          If num_comprobante < 10 Then
                VAR_CITE = dtc_codigo1 + "-00000" + Trim(num_comprobante)
             End If
             If num_comprobante > 9 And var_cod < 100 Then
               VAR_CITE = dtc_codigo1 + "-0000" + Trim(num_comprobante)
             End If
             If num_comprobante > 99 And var_cod < 1000 Then
                VAR_CITE = dtc_codigo1 + "-000" + Trim(num_comprobante)
             End If
             If num_comprobante > 999 And var_cod < 10000 Then
                VAR_CITE = dtc_codigo1 + "-00" + Trim(num_comprobante)
             End If
             If num_comprobante > 9999 And var_cod < 100000 Then
                VAR_CITE = dtc_codigo1 + "-0" + Trim(num_comprobante)
             End If
             If num_comprobante > 99999 Then
               VAR_CITE = dtc_codigo1 + "-" + Trim(num_comprobante)
             End If
             sql_adicionM = "insert into Co_Comprobante_M (Cod_Comp,                tipo_comp,                    cod_trans,       org_codigo ,   ges_gestion ,          Fecha_transacion        ,             beneficiario_codigo,              Glosa              ,           unidad_codigo    ,              solicitud_codigo,               tipo_moneda ,       unidad_codigo_ant,          proceso_codigo,                 subproceso_codigo ,                 etapa_codigo,      clasif_codigo,   doc_codigo ,            doc_numero,             pro_codigo_det,  literal, estado_codigo  ,       usr_codigo,       fecha_registro,          Hora_Aprueba,  mes_trasaccion) " & _
                                           "values (" & num_comprobante & ",'" & Trim(Me.CboTipo) & "', " & num_comprobante & ",  '999',  '" & Year(DTPfecha_trans.Value) & "',   '" & DTPfecha_trans.Value & "', '" & dtc_codigo4.Text & "', '" & Trim(Me.Txt_glosa) & "',    '" & dtc_codigo1.Text & "' ,      " & txtcodsolicitud.Text & " ,'" & cmb_moneda & "' ,'" & VAR_CITE & "', '" & Left(dtc_codigo7.Text, 3) & "', '" & Left(dtc_codigo7.Text, 6) & "', '" & dtc_codigo7.Text & "', 'ADM', '" & txt_codigo1.Text & "',  " & txt_campo1.Caption & " , '20101-0', '-'  ,  'REG' ,     '" & Trim(glusuario) & "' ,'" & Date & "','" & Format(Time, "hh:mm:ss") & "', '" & txt_mes.Text & "' )"
             db.Execute sql_adicionM
        End If
        If adiciona = "M" Then
            If UCase(MonthName(Month(DTPfecha_trans.Value))) = txt_mes.Text Then      ' Cambiar al mismo formato !!!!
                db.Execute " UPDATE Co_Comprobante_M set beneficiario_codigo='" & dtc_codigo4 & "',Glosa='" & Txt_glosa & "',Fecha_transacion='" & DTPfecha_trans & "',solicitud_codigo=" & txtcodsolicitud & ",Tipo_Comp='" & CboTipo & "',etapa_codigo='" & dtc_codigo7 & "'  WHERE Cod_Comp=" & TxtComprobante & " "
            Else
                MsgBox "No se puede modificar una fecha correspondiente a otro MES, Intente Nuevamente ...", vbExclamation + vbDefaultButton1
                Exit Sub
            End If
        End If
'   adiciona = ""
           Call OptSinAprobar_Click
            'rs_datos.Update
           ' rs_datos.MoveLast
            Me.FraNavega.Enabled = True
            Me.FrmABMDet1.Visible = True
            Me.FraDet1.Visible = True
            Me.FraDet2.Visible = True
            FraGlobal.Enabled = False
            dg_datos.Enabled = True
            FraDet1.Visible = False
            FraDet2.Visible = False
            Me.fraOpciones.Visible = True
            VAR_SW = ""
            dtc_desc14.Locked = False
End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
'GRABAR MEVM FIN




'  Me.frameCAM.Visible = False
'  Dim sql_adicionM As String
'  Dim sql_adicionD As String
'  Dim rsbef As ADODB.Recordset
'  Set rsbef = New ADODB.Recordset
'  Dim rsbef1 As ADODB.Recordset
'  Set rsbef1 = New ADODB.Recordset
'  If rsbef.State = 1 Then rsbef.Close
'  rsbef.CursorLocation = adUseClient
'  rsbef.Open "SELECT beneficiario_codigo, beneficiario_denominacion From fc_beneficiario " & _
'            " where beneficiario_codigo='" & Trim(Me.dtc_codigo4.Text) & "'", db, adOpenKeyset, adLockReadOnly
'  If rsbef.RecordCount = 0 Then
'    MsgBox "El beneficiario no existe. Seleccione un beneficiario", vbExclamation + vbDefaultButton1
'    'Me.dtc_codigo4.SetFocus
'    Exit Sub
'  End If
'  If rsbef1.State = 1 Then rsbef1.Close
'  rsbef1.CursorLocation = adUseClient
'  rsbef1.Open "SELECT beneficiario_codigo, beneficiario_denominacion From fc_beneficiario " & _
'             " where beneficiario_denominacion='" & Trim(Me.dtc_desc4.Text) & "'", db, adOpenKeyset, adLockReadOnly
'  If rsbef1.RecordCount = 0 Then
'     MsgBox "El beneficiario no existe. Seleccione un beneficiario", vbExclamation + vbDefaultButton1
'     'Me.dtc_desc4.SetFocus
'     Exit Sub
'  End If
'   ' If cmodificar = "N" Then
'   '****VALIDACION DE CAMPOS VACIOS GENERALES
'        If Len(Trim(CboTipo.Text)) = 0 Then
'          MsgBox "Elija el tipo de comprobante", vbExclamation + vbDefaultButton1
'          'CboTipo.SetFocus
'          Exit Sub
'        End If
''        If Len(Trim(txt_codigo1.Text)) = 0 Then
''              MsgBox "Elija el tipo de documento de respaldo", vbExclamation + vbDefaultButton1
''              'dtcbodocumento1.SetFocus
''              Exit Sub
''        End If
'        If Len(Trim(Me.txt_campo1)) = 0 Then
'          MsgBox "Coloque el número de respaldo", vbExclamation + vbDefaultButton1
'          'Me.txt_campo1.SetFocus
'          Exit Sub
'        End If
'        If Me.CboTipo = "PCE" And cmodificar = "N" Then
'            If Len(Trim(Me.txtcodsolicitud)) = 0 Then
'                MsgBox "Coloque el número de solicitud", vbExclamation + vbDefaultButton1
'                'txtcodsolicitud.SetFocus
'                Exit Sub
'            End If
'        End If
'        If Len(Trim(Me.dtc_codigo4)) = 0 Or Len(Trim(Me.dtc_desc4)) = 0 Then
'          MsgBox "Elija un beneficiario", vbExclamation + vbDefaultButton1
'          'dtc_codigo4.SetFocus
'          Exit Sub
'        End If
'        'If Len(Trim(Me.dtc_desc4)) = 0 Then
'        '  MsgBox "Elija un beneficiario", vbExclamation + vbDefaultButton1
'          'dtc_desc4.SetFocus
'        '  Exit Sub
'        'End If
'        If Len(Trim(Me.Txt_glosa)) = 0 Then
'          MsgBox "Escriba la glosa", vbExclamation + vbDefaultButton1
'          'Txt_glosa.SetFocus
'          Exit Sub
'        End If
'    'VALIDACION PARA COMPROBANTES DIFERENTES DE CAM
'    If CboTipo.Text <> "CAM" Then
'        If Len(Trim(CboDCta.Text)) = 0 Then
'           MsgBox "Elija la cuenta Debe", vbExclamation + vbDefaultButton1
'           'CboDCta.SetFocus
'           Exit Sub
'        End If
'        If Len(Trim(CboDSubcta1.Text)) = 0 Then
'              MsgBox "Elija la subcuenta Debe", vbExclamation + vbDefaultButton1
'              'CboDSubcta1.SetFocus
'              Exit Sub
'        End If
'        If Len(Trim(CboDSubcta2.Text)) = 0 Then
'              MsgBox "Elija la subcuenta Debe", vbExclamation + vbDefaultButton1
'              'CboDSubcta2.SetFocus
'              Exit Sub
'        End If
''        If Len(Trim(Me.TxtDSus)) = 0 Then
''          MsgBox "Escriba un monto en el Debe", vbExclamation + vbDefaultButton1
''          ' TxtDSus.SetFocus
''          Exit Sub
''        End If
'        If Len(Trim(CboHcta.Text)) = 0 Then
'              MsgBox "Elija la cuenta Haber", vbExclamation + vbDefaultButton1
'              'CboHcta.SetFocus
'              Exit Sub
'        End If
'        If Len(Trim(CbohSubcta1.Text)) = 0 Then
'              MsgBox "Elija la subcuenta Haber", vbExclamation + vbDefaultButton1
'              'CbohSubcta1.SetFocus
'              Exit Sub
'        End If
'        If Len(Trim(CbohSubcta2.Text)) = 0 Then
'              MsgBox "Elija la subcuenta Haber", vbExclamation + vbDefaultButton1
'              'CbohSubcta2.SetFocus
'              Exit Sub
'        End If
'    '---
''        Call Titulo(Me.CboDCta, Me.CboDSubcta1, Me.CboDSubcta2)
'        Select Case lcta
'         Case "N"
'            Exit Sub
'         Case "S"
'            If MovCuenta = "T" Then Exit Sub
'        End Select
'    '---
'        Call Titulo(Me.CboHcta, Me.CbohSubcta1, Me.CbohSubcta2)
'        Select Case lcta
'         Case "N"
'            Exit Sub
'         Case "S"
'            If MovCuenta = "T" Then Exit Sub
'        End Select
'      '-----
''        If Len(Trim(Me.TxtDBs)) = 0 Then
''          MsgBox "Escriba un monto en el Debe", vbExclamation + vbDefaultButton1
''          'Me.TxtDBs.SetFocus
''          Exit Sub
''        End If
''        If Me.frameDCtaBancaria.Visible = True And CboTipo <> "CAM" Then
''          'If Me.CboTipo <> "CAM" Then
''            If Me.CboDCta.Text = Empty Or Me.cboDctaaux1.Text = Empty Then
''                MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1
''                Exit Sub
''            End If
''         ' End If
''        End If
'    End If
'    'VALIDACION PARA COMPROBANTES DE TIPO CAM
'    If Me.CboTipo = "CAM" Then
''      If Me.CboDCtaCAM.Text = "1111" Then
''            If Me.CboDCtaCAM.Text = Empty Or Me.cboDctaaux1.Text = Empty Then
''                MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1
''                Exit Sub
''            End If
''      End If
'      '--------- g-- CAMBIO PARA CAMBIAR DE AUXILIAR A LAS CUENTAS 6141 Y 5174
''      If CboDCtaCAM = "6141" Then
''          If Me.cboDCodOrg = Empty Then
''            MsgBox "Seleccione un organismo ", vbExclamation + vbDefaultButton1
''            Exit Sub
''          End If
''      End If
'    'End If
'    If Me.frameHCtaBancaria.Visible = True Then
'        If Me.CboTipo = "CAM" Then
'           If Me.CboHCtaCAM.Text = "1111" Then
'              If Me.CboHCtaCAM.Text = Empty Or Me.cboHctaaux1.Text = Empty Then
'                MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1
'                Exit Sub
'              End If
'            End If
'        End If
'    End If
'
'    'End If
'    '******
'    If Trim(CboTipo.Text) = "PCE" Then
'         permitectas Trim(CboDCta), Trim(CboDSubcta1.Text), Trim(CboTipo.Text)
'         If permite = 1 Then Exit Sub
'         permitectas Trim(CboHcta), Trim(CbohSubcta1.Text), Trim(CboTipo.Text)
'         If permite = 1 Then Exit Sub
'    End If
'    If Trim(CboTipo.Text) = "PCO" Or Trim(CboTipo.Text) = "CAM" Then
'          Me.txtcodsolicitud = "-"
'    End If
'
'    '-----
'    '----
''    If SSTabDebe.TabEnabled(0) = True Then
''    Else
''      dctalarga = ""
''    End If
''    If SSTabDebe.TabEnabled(1) = True Then
''
''    Else
''      dctaaux2 = ""
''    End If
''    If SSTabDebe.TabEnabled(2) = True Then
''    Else
''     dctaaux3 = ""
''    End If
''
''    If SSTabHaber.TabEnabled(0) = True Then
''    Else
''      hctalarga = ""
''    End If
''    If SSTabHaber.TabEnabled(1) = True Then
''    Else
''      hctaaux2 = ""
''    End If
''    If SSTabHaber.TabEnabled(2) = True Then
''    Else
''      hctaaux3 = ""
''    End If
'    '---verificar llenado de convenios
'    'If TDBFrameDConvenio.Visible = True Then
'    '---nuevo por adicion de unidades educativas
'    If daux1 = "10" Or daux2 = "10" Or daux3 = "10" Then
'       Dim rsedu1 As ADODB.Recordset
'       Set rsedu1 = New ADODB.Recordset
'       rsedu1.CursorLocation = adUseClient
'       rsedu1.Open "SELECT codigo, denominacion From fc_unidad_educativa WHERE codigo = '" & Trim(dtcDIdCaja.Text) & "'", db, adOpenKeyset, adLockReadOnly
'       If rsedu1.RecordCount = 0 Then
'            MsgBox "Seleccione una Unidad Educativa válida!!!!", vbExclamation + vbDefaultButton1
'            Exit Sub
'       End If
'    End If
'
'    If haux1 = "10" Or haux2 = "10" Or haux3 = "10" Then
'       Dim rsedu As ADODB.Recordset
'       Set rsedu = New ADODB.Recordset
'       rsedu.CursorLocation = adUseClient
'       rsedu.Open "SELECT codigo, denominacion From fc_unidad_educativa WHERE codigo = '" & Trim(DTCHidcaja.Text) & "'", db, adOpenKeyset, adLockReadOnly
'       If rsedu.RecordCount = 0 Then
'            MsgBox "Seleccione una Unidad Educativa válida!!!!", vbExclamation + vbDefaultButton1
'            Exit Sub
'       End If
'    End If
'
'    '----
'    If daux1 = "09" Or daux2 = "09" Or daux3 = "09" Then
'      If Trim(DtCDIdConvenio.Text) = "" Then
'            MsgBox "Seleccione un Convenio en el Debe", vbExclamation + vbDefaultButton1
'            Exit Sub
'      End If
'    End If
'
'    'If TDBFrameHConvenio.Visible = True Then
'    If haux1 = "09" Or haux2 = "09" Or haux3 = "09" Then
'      If Trim(DtCHIdConvenio.Text) = "" Then
'            MsgBox "Seleccione un Convenio en el Haber", vbExclamation + vbDefaultButton1
'            Exit Sub
'      End If
'    End If
'    '---
''    frameactivoDebe
'    If salir = 1 Then
'      Exit Sub
'    End If
''    frameactivoHaber
'    If salir = 1 Then
'      Exit Sub
'    End If
''    MsgBox "dctalargA:    " & dctalarga
''    MsgBox "DCUENTA2:     " & dctaaux2
''    MsgBox "DCUENTA3:     " & dctaaux3
''    frameactivoHaber
''    MsgBox "hctalargA:    " & hctalarga
''    MsgBox "hCUENTA2:     " & hctaaux2
''    MsgBox "hCUENTA3:     " & hctaaux3
'    db.BeginTrans
'    Select Case cmodificar
'    Case "N", "C"
'    '    db.BeginTrans 'inicio de la transaccion
'        '****ADICION ALCOMPROBANTE_M
'        'Call genera_codigo
'        '****ADICION ALCOMPROBANTE_M
'        If Me.CboTipo = "CAM" Then
'            Select Case CAMcorrel
'              Case "NOR"
''                Call genera_codigo
'              Case "CAM"
''                genera_CorrelCAM Me.DTPCAM.Value
'            End Select
'        Else
''          Call genera_codigo
'        End If
'
'        '********ADICION AL DIARIO
''      If Trim(CboTipo.Text) = "PCO" Or Trim(CboTipo.Text) = "PCE" Then
''        sql_adicionM = "insert into Co_Comprobante_M (cod_comp,tipo_comp," & _
''                    "cod_trans,cod_trans_detalle,org_codigo," & _
''                    "ges_gestion,num_respaldo,Fecha_transacion,beneficiario_codigo," & _
''                    "codigo_documento,glosa,status,usr_usuario,fecha_registro," & _
''                    "hora_registro,tipo_moneda,codigo_solicitud)" & _
''                    "values (" & Trim(Str(num_comprobante)) & ",'" & Trim(Me.CboTipo) & "'," & _
''                    "'-','1','999','" & Trim(Me.txt_ges) & "','" & Trim(Me.txt_campo1) & "','" & _
''                    CDate(Format(CFecha, "dd/mm/yyyy")) & "','" & Trim(Me.dtc_codigo4.Text) & _
''                    "','" & Trim(Me.txt_codigo1.Text) & "','" & Trim(Me.Txt_glosa) & "'," & _
''                    "'N','" & Trim(glusuario) & "','" & CDate(Format(CFecha, "dd/mm/yyyy")) & _
''                    "','" & Format(Time, "hh:mm:ss") & "','" & Trim(Ctipomoneda) & "','" & Trim(Me.txtcodsolicitud) & " ')"
''
''        sql_adicionD = "insert into Co_Diario (cod_comp,tipo_comp,cod_comp_c,d_cuenta,d_subcta1,d_subcta2,d_aux1," & _
''            "d_aux2,d_aux3,D_Cta_Aux1,D_Cta_Aux2,d_ctaAux3,d_montoBs,d_montoDl,d_Cambio," & _
''            "h_cuenta,h_subcta1,h_subcta2,h_aux1,h_aux2,h_aux3,H_Cta_Aux1," & _
''            "H_Cta_Aux2,H_Cta_Aux3,h_montoBs,h_montoDl,h_Cambio,usr_usuario,fecha_registro,hora_registro) " & _
''            "values (" & Trim(Str(num_comprobante)) & ",'" & Trim(Me.CboTipo) & "',0,'" & _
''            Trim(Me.CboDCta) & "','" & Trim(Me.CboDSubcta1) & "','" & Trim(Me.CboDSubcta2) & "','" & _
''            daux1 & "','" & daux2 & "','" & daux3 & "','" & dctalarga & "','" & dctaaux2 & "','" & _
''            dctaaux3 & "'," & Val(TxtDBs) & "," & _
''            Val(TxtDSus) & "," & Val(lblDTC) & ",'" & Trim(Me.CboHcta) & "','" & Trim(Me.CbohSubcta1) & "','" & _
''            Trim(Me.CbohSubcta2) & "','" & haux1 & "','" & haux2 & "','" & haux3 & "','" & hctalarga & "','" & _
''            hctaaux2 & "','" & hctaaux3 & "'," & _
''            Val(txtHBs) & "," & Val(txtHsus) & "," & Val(lblDTC) & ",'" & glusuario & "','" & _
''            CDate(Format(CFecha, "dd/mm/yyyy")) & "','" & Format(Time, "hh:mm:ss") & "')"
''      End If
'      If Trim(CboTipo.Text) = "CAM" Then
'        If optdolares.Value = True Then
''          Me.TxtDBs = "0.0"
'          Me.txtHBs = "0.0"
'        End If
''        If optbolivianos.Value = True Then
''          Me.TxtDSus = "0.0"
''          Me.txtHsus = "0.0"
''        End If
''        sql_adicionM = "insert into Co_Comprobante_M (cod_comp,tipo_comp," & _
''                    "cod_trans,cod_trans_detalle,org_codigo," & _
''                    "ges_gestion,num_respaldo,Fecha_transacion,beneficiario_codigo," & _
''                    "codigo_documento,glosa,status,usr_usuario,fecha_registro," & _
''                    "hora_registro,tipo_moneda,codigo_solicitud)" & _
''                    "values (" & Trim(Str(num_comprobante)) & ",'" & Trim(Me.CboTipo) & "'," & _
''                    "'-','1','999','" & Trim(Me.txt_ges) & "','" & Trim(Me.txt_campo1) & "','" & _
''                    CDate(Format(Me.DTPCAM.Value, "dd/mm/yyyy")) & "','" & Trim(Me.dtc_codigo4.Text) & _
''                    "','" & Trim(Me.txt_codigo1.Text) & "','" & Trim(Me.Txt_glosa) & "'," & _
''                    "'N','" & Trim(glusuario) & "','" & CDate(Format(CFecha, "dd/mm/yyyy")) & _
''                    "','" & Format(Time, "hh:mm:ss") & "','" & Trim(Ctipomoneda) & "','" & Trim(Me.txtcodsolicitud) & " ')"
''
''        sql_adicionD = "insert into Co_Diario (cod_comp,tipo_comp,cod_comp_c,d_cuenta,d_subcta1,d_subcta2,d_aux1," & _
''            "d_aux2,d_aux3,D_Cta_Aux1,d_montoBs,d_montoDl,d_Cambio," & _
''            "h_cuenta,h_subcta1,h_subcta2,h_aux1,h_aux2,h_aux3,H_Cta_Aux1," & _
''            "h_montoBs,h_montoDl,h_Cambio,usr_usuario,fecha_registro,hora_registro) " & _
''            "values (" & Trim(Str(num_comprobante)) & ",'" & Trim(Me.CboTipo) & "',0,'" & _
''            Trim(Me.CboDCtaCAM) & "','" & Trim(Me.CboDSub1CAM) & "','" & Trim(Me.CboDSub2CAM) & "','" & _
''            daux1 & "','" & daux2 & "','" & daux3 & "','" & dctalarga & "'," & Val(TxtDBs) & "," & _
''            Val(TxtDSus) & "," & Val(lblDTC) & ",'" & Trim(Me.CboHCtaCAM) & "','" & Trim(Me.CboHSub1CAM) & "','" & _
''            Trim(Me.CboHSub2CAM) & "','" & haux1 & "','" & haux2 & "','" & haux3 & "','" & hctalarga & "'," & _
''            Val(txtHBs) & "," & Val(txtHsus) & "," & Val(lblDTC) & ",'" & glusuario & "','" & _
''            CDate(Format(CFecha, "dd/mm/yyyy")) & "','" & Format(Time, "hh:mm:ss") & "')"
''      End If
'        db.Execute sql_adicionM
'        db.Execute sql_adicionD
'
'      '  db.CommitTrans
'        If cmodificar = "C" Then
'          MsgBox "Copio el comprobante " & num_comprobante & "  " & Trim(CboTipo.Text), vbInformation + vbDefaultButton1, "Atencion"
'          frame_moneda.Enabled = True
'          'cmodificar = "M"
'        Else
'          MsgBox "Registro el comprobante " & num_comprobante & "  " & Trim(CboTipo.Text), vbInformation + vbDefaultButton1, "Atencion"
'        End If
'        Me.TxtComprobante = num_comprobante
'        rs_datos.Requery
'        rs_datos.Find "cod_comp=" & num_comprobante, , adSearchForward, 1
''      Case "M"
'     '       db.BeginTrans 'inicio de la transaccion
'            '****ADICION ALCOMPROBANTE_M
'            'Call genera_codigo
'          Select Case CboTipo
'           Case "ANL", "DVL", "RVT"
''               rs_datos.Requery
''               ModifAsientos Me.Txt_glosa, Val(Me.TxtDBs), Val(Me.TxtDSus)
'               rs_datos.Requery
'               MsgBox "Comprobante modificado", vbInformation + vbDefaultButton1
'           Case Else
'
'               Numero = Val(Trim(Me.TxtComprobante))
'               Dim rsmodificaM As ADODB.Recordset
'               Set rsmodificaM = New ADODB.Recordset
'               Dim rsmodificaD As ADODB.Recordset
'               Set rsmodificaD = New ADODB.Recordset
'               If rsmodificaM.State = 1 Then rsmodificaM.Close
'               rsmodificaM.Open "select * from Co_comprobante_M where cod_comp=" & Val(Trim(Me.TxtComprobante)), db, adOpenKeyset, adLockOptimistic
'               If rsmodificaD.State = 1 Then rsmodificaD.Close
'               rsmodificaD.Open "select * from CO_diario where cod_comp=" & Val(Trim(Me.TxtComprobante)), db, adOpenKeyset, adLockOptimistic
''               If rsmodificaM.RecordCount <> 0 And rsmodificaD.RecordCount <> 0 Then
''                   rsmodificaM!num_respaldo = Trim(Me.txt_campo1)
''                   'rsmodificaM!Fecha_transacion = CDate(Format(CFecha, "dd/mm/yyyy"))
''                   rsmodificaM!beneficiario_codigo = Trim(Me.dtc_codigo4.Text)
''                   rsmodificaM!codigo_documento = Trim(Me.txt_codigo1.Text)
''                   rsmodificaM!glosa = Trim(Me.Txt_glosa)
''                   rsmodificaM!usr_usuario = Trim(glusuario)
''                   rsmodificaM!fecha_registro = CDate(Format(CFecha, "dd/mm/yyyy"))
''                   rsmodificaM!hora_registro = Format(Time, "hh:mm:ss")
''                   rsmodificaM!tipo_moneda = Trim(Ctipomoneda)
''                   rsmodificaM!codigo_solicitud = Trim(Me.txtcodsolicitud)
'                   '********ADICION AL DIARIO
'                 Select Case Trim(CboTipo)
'                  Case "PCO", "PCE", "ANL", "DVL", "RVT"
'                 'If Trim(CboTipo) = "PCO" Or Trim(CboTipo) = "PCE" Or "ANL" Or "DVL" Or "RVT" Then
''                    rsmodificaD!D_Cuenta = Trim(Me.CboDCta)
''                    rsmodificaD!D_Subcta1 = Trim(Me.CboDSubcta1)
''                    rsmodificaD!D_Subcta2 = Trim(Me.CboDSubcta2)
'                    rsmodificaD!h_cuenta = Trim(Me.CboHcta)
'                    rsmodificaD!h_subcta1 = Trim(Me.CbohSubcta1)
'                    rsmodificaD!h_subcta2 = Trim(Me.CbohSubcta2)
'                    rsmodificaM!Fecha_transacion = CDate(Format(CFecha, "dd/mm/yyyy"))
''                    CboDSubcta2_Click
''                    CbohSubcta2_Click
'                  Case "CAM"
''                    If optdolares.Value = True Then
''                        Me.TxtDBs = "0.0"
''                        Me.txtHBs = "0.0"
''                    End If
''                    If optbolivianos.Value = True Then
''                        Me.TxtDSus = "0.0"
''                        Me.txtHsus = "0.0"
''                    End If
''                    rsmodificaD!D_Cuenta = Trim(Me.CboDCtaCAM)
''                    rsmodificaD!D_Subcta1 = Trim(Me.CboDSub1CAM)
''                    rsmodificaD!D_Subcta2 = Trim(Me.CboDSub2CAM)
'                    rsmodificaD!h_cuenta = Trim(Me.CboHCtaCAM)
'                    rsmodificaD!h_subcta1 = Trim(Me.CboHSub1CAM)
'                    rsmodificaD!h_subcta2 = Trim(Me.CboHSub2CAM)
'                    rsmodificaM!Fecha_transacion = CDate(Format(DTPCAM.Value, "dd/mm/yyyy"))
''                    CboDSub2CAM_Click
''                    CboHSub2CAM_Click
'                 End Select
'                    rsmodificaD!d_Aux1 = Trim(daux1)
'                    rsmodificaD!d_Aux2 = Trim(daux2)
'                    rsmodificaD!d_Aux3 = Trim(daux3)
'                    rsmodificaD!D_Cta_Aux1 = Trim(dctalarga)
'                    rsmodificaD!D_Cta_Aux2 = dctaaux2
'                    rsmodificaD!d_CtaAux3 = dctaaux3
'                    rsmodificaD!H_Cta_Aux2 = hctaaux2
'                    rsmodificaD!H_Cta_Aux3 = hctaaux3
'                    rsmodificaD!d_montoBs = Val(TxtDBs)
'                    rsmodificaD!d_montoDl = Val(TxtDSus)
''                    rsmodificaD!d_Cambio = Val(Me.lblDTC)
'                    rsmodificaD!h_Aux1 = Trim(haux1)
'                    rsmodificaD!h_Aux2 = Trim(haux2)
'                    rsmodificaD!h_Aux3 = Trim(haux3)
'                    rsmodificaD!H_Cta_Aux1 = Trim(hctalarga)
'                    rsmodificaD!h_montoBs = Val(txtHBs)
'                    rsmodificaD!h_montoDl = Val(txtHsus)
'                    rsmodificaD!h_Cambio = Val(Me.lblHTC)
'                    rsmodificaD!usr_usuario = glusuario
'                    rsmodificaD!fecha_registro = CDate(Format(CFecha, "dd/mm/yyyy"))
'                    rsmodificaD!hora_registro = Format(Time, "hh:mm:ss")
'                    rsmodificaM.Update
'                    rsmodificaD.Update
'              ' End If
'            '   db.CommitTrans
'               rs_datos.Requery
'               rs_datos.Find "Cod_Comp =" & Numero
'               MsgBox "Comprobante modificado", vbInformation + vbDefaultButton1
'           End Select
'       ' End Select
''        db.CommitTrans
'        'rs_datos.Sort = "cod_comp"
'        Set Me.dg_datos.DataSource = rs_datos
'        'rs_datos.Find "cod_comp=" & num_comprobante, , adSearchForward, 1
'        If cmodificar = "C" Then
'            Me.FraGrabarCancelar.Visible = True
'            Me.fraOpciones.Visible = False
'            'Me.fraOpciones.Visible = False
'            'Me.Fram_AsientoD.Enabled = True
'            'Me.Fram_AsientoH.Enabled = True
'            TDBFrameDebeCta.Enabled = True
'            TDBFrameDebe.Enabled = True
'            TDBFrameHaber.Enabled = True
'            TDBFrameHaberCta.Enabled = True
'            Me.FraGlobal.Enabled = True
'            Me.FraNavega.Enabled = False
'            Me.frame_moneda.Visible = True
'            Me.frame_moneda.Enabled = True
'            cmodificar = "M"
'        Else
''            Me.sstab1.Tab = 0
'            Me.FraGrabarCancelar.Visible = False
'            Me.fraOpciones.Visible = True
'            Me.frame_moneda.Enabled = False
'            'Me.FraGrabarCancelar.Visible = False
'            Me.fraOpciones.Visible = True
'            'Me.Fram_AsientoD.Enabled = False
'            'Me.Fram_AsientoH.Enabled = False
'            TDBFrameDebeCta.Enabled = False
'            TDBFrameDebe.Enabled = False
'            TDBFrameHaber.Enabled = False
'            TDBFrameHaberCta.Enabled = False
'            Me.FraGlobal.Enabled = False
'            Me.FraNavega.Enabled = True
'        End If
''        Me.lblDTC.Locked = True
'        Me.dg_datos.Enabled = True
'        'If cmodificar <> "C" Then
'        '  rs_datos.MoveLast
'        '  dg_datos_Click
'        'End If
'        'If cmodificar <> "C" Then
'        ' rs_datos.Find "cod_comp=" & num_comprobante, , adSearchForward, 1
'        'End If
'        db.CommitTrans
'       ' tipocompllena rs_datos!tipo_comp 'para llenar el combo de tipo de comprobantes
''        Framebotones.Enabled = True
'        frame_moneda.Enabled = False
'Exit Sub
'err3:
'    db.RollbackTrans
'    MsgBox "Error al actualizar los datos"
'    Exit Sub
'
'End If

End Sub
Private Sub valida_campos()
  If dtc_codigo4.Text = "" Then
    MsgBox "Debe registrar el " + lbl_beneficiario.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_glosa.Text = "" Then
    MsgBox "Debe registrar la " + lbl_glosa.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If CboTipo.Text = "" Then
    MsgBox "Debe registrar la " + lbl_tipo.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo7.Text = "" Then
    MsgBox "Debe registrar: " + lbl_etapa.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
    If txtcodsolicitud.Text = "" Then
    MsgBox "Debe registrar: " + lbl_tramite.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
End Sub

''Private Sub cmdimprime_grid_Click()
''Dim i As Integer
''Set rsbenef = New ADODB.Recordset
''Set rsimprgrid = New ADODB.Recordset
''db.Execute " truncate table impresion_grid"
''
''If rsimprgrid.State = 1 Then rsimprgrid.Close
''    rsimprgrid.Open " select * from impresion_grid", db, adOpenKeyset, adLockOptimistic
'''MsgBox rsimprgrid.RecordCount
''    'AdodcAprob.Recordset.MoveFirst
''If rs_datos.RecordCount > 0 Then
''rs_datos.MoveFirst
''Do While Not rs_datos.EOF
''  rsimprgrid.AddNew
''  rsimprgrid!Cod_Comp = rs_datos!Cod_Comp
''  rsimprgrid!tipo_comp = rs_datos!tipo_comp
''  rsimprgrid!beneficiario_codigo = rs_datos!beneficiario_codigo
''  rsimprgrid!cod_trans = rs_datos!cod_trans
''  rsimprgrid!org_codigo = rs_datos!org_codigo
''  rsimprgrid!Status = rs_datos!Status
''  If rsbenef.State = 1 Then rsbenef.Close
''    rsbenef.Open "select beneficiario_denominacion,beneficiario_codigo from fc_beneficiario where beneficiario_codigo = '" & rs_datos!beneficiario_codigo & "'", db, adOpenKeyset, adLockReadOnly
''  If rsbenef.RecordCount <> 0 Then
''    rsimprgrid!denom_beneficiario = rsbenef!beneficiario_denominacion
''  Else
''    rsimprgrid!denom_beneficiario = " "
''  End If
''  rsimprgrid.Update
''  rs_datos.MoveNext
''Loop
''CryRepGrid.Destination = crptToWindow
''CryRepGrid.WindowShowPrintSetupBtn = True
''CryRepGrid.WindowShowRefreshBtn = True
''CryRepGrid.WindowState = crptMaximized
''CryRepGrid.ReportFileName = App.Path & "\FormsContabilidad\reportes\CryRepGrid.rpt"
''i = CryRepGrid.PrintReport
''   If i <> 0 Then
''               MsgBox CryRepGrid.LastErrorNumber & " : " & CryRepGrid.LastErrorString, vbExclamation + vbOKOnly, "Error..."
''   End If
''rs_datos.MoveFirst
''dg_datos_Click
'''frmrepgrid.Show
'''rs_datos.MoveFirst
''End If
''End Sub

''Private Sub cmdPrimero_Click()
''If rs_datos.RecordCount = 0 Then
''  Exit Sub
''End If
''rs_datos.MoveFirst
''
''If rs_datos.BOF Then
''    rs_datos.MoveFirst
''    dg_datos_Click
''Else
''    dg_datos_Click
''End If
''End Sub

Private Sub BtnSalir_Click()
  Set Me.dg_datos.DataSource = Nothing
  Unload Me
End Sub

''Private Sub Cmdatras_Click()
''If rs_datos.BOF Then
''    rs_datos.MoveNext
''    dg_datos_Click
''  Else
''    rs_datos.MovePrevious
''    dg_datos_Click
''  End If
''End Sub
'
''Private Sub Cmdsgte_Click()
''If rs_datos.RecordCount = 0 Then
''  Exit Sub
''End If
''If rs_datos.EOF Then
''    rs_datos.MovePrevious
''    dg_datos_Click
''  Else
''    rs_datos.MoveNext
''    dg_datos_Click
''  End If
''End Sub
'
''Private Sub Cmdinicio_Click()
''  rs_datos.MoveFirst
''End Sub
'
''Private Sub Cmdfin_Click()
''  rs_datos.MoveLast
''End Sub

''Private Sub cmdsiguiente_Click()
''If rs_datos.RecordCount = 0 Then
''  Exit Sub
''End If
''rs_datos.MoveNext
''If rs_datos.EOF Then
''    rs_datos.MoveLast
''    dg_datos_Click
''Else
''    dg_datos_Click
''End If
''End Sub

'Private Sub DtCDcodbenef_Change()
'     Me.dtc_desc4.BoundText = Trim(Me.DtCDcodbenef.BoundText)
'     Select Case cmodificar
'        Case "M", "N"
'            Me.lblDBenefaux1 = DtCDcodbenef.Text
'            'Call buscabenef(Trim(DtCDcodbenef.Text))
'            'Me.lblDnomBenefaux1 = Cdenominacion
'            Me.lblDnomBenefaux1 = DtCDDescripbenef.Text
'            Me.lblHBenefaux1 = DtCHcodbenef.Text
'            Me.lblHnomBenefaux1 = DtCHDescripbenef.Text
'     End Select
'     If CboTipo.Text = "PCO" Then
'     DtCDcodbenef.Text = dtc_codigo4.Text
'     DtCDcodbenef_Click (1)
'     DtCHcodbenef.Text = dtc_codigo4.Text
'     DtCHcodbenef_Click (1)
'     End If
'End Sub
'Private Sub D1documento_Change()
'    'Me.D2descripcion.BoundText = Me.D1documento.BoundText
'End Sub
'
'Private Sub dtc_codigo4_LostFocus()
'Dim rsbef As ADODB.Recordset
'  Set rsbef = New ADODB.Recordset
'  rsbef.CursorLocation = adUseClient
'  rsbef.Open "SELECT beneficiario_codigo, beneficiario_denominacion From fc_beneficiario " & _
'            " where beneficiario_codigo='" & Trim(Me.dtc_codigo4.Text) & "'", db, adOpenKeyset, adLockReadOnly
'  If rsbef.RecordCount = 0 Then
'    MsgBox "El beneficiario no existe. Seleccione un beneficiario", vbExclamation + vbDefaultButton1
'    'Me.dtc_codigo4.SetFocus
'    Exit Sub
'  End If
'End Sub
'
'Private Sub D2descripcion_Change()
'    'Me.D1documento.Text = Me.D2descripcion.BoundText
'End Sub
'
'Private Sub D2descripcion_Click(Area As Integer)
'    'Me.D1documento.Text = Me.D2descripcion.BoundText
'End Sub
'
'Private Sub dtc_desc4_LostFocus()
'    Dim rsbef As ADODB.Recordset
'    Set rsbef = New ADODB.Recordset
'    rsbef.CursorLocation = adUseClient
'    rsbef.Open "SELECT beneficiario_codigo, beneficiario_denominacion From fc_beneficiario " & _
'                " where beneficiario_denominacion='" & Trim(Me.dtc_desc4.Text) & "'", db, adOpenKeyset, adLockReadOnly

Private Sub ABRIR_DEBE()
'DEBE
        Set rs_detalle1 = New ADODB.Recordset
        If rs_detalle1.State = 1 Then rs_detalle1.Close
        If VAR_SW = "CNL" Then
            Set dg_det1.DataSource = rsNada
        Else
        
            If VAR_SW = "ADD" Then
                rs_detalle1.Open "Select * from cv_diario_debe order by Cod_Comp_Detalle", db, adOpenKeyset, adLockOptimistic
            Else
                rs_detalle1.Open "Select * from cv_diario_debe where Cod_Comp = " & Ado_datos.Recordset!Cod_Comp & " order by Cod_Comp_Detalle", db, adOpenKeyset, adLockOptimistic
            End If
            
            Set Ado_detalle1.Recordset = rs_detalle1
            If Ado_detalle1.Recordset.RecordCount > 0 Then
                dg_det1.Visible = True
                Set dg_det1.DataSource = rs_detalle1
            Else
                dg_det1.Visible = False
                Set dg_det1.DataSource = rsNada
        
            End If
        End If
End Sub
Private Sub ABRIR_HABER()
'HABER
        Set rs_detalle2 = New ADODB.Recordset
        If rs_detalle2.State = 1 Then rs_detalle2.Close
        If VAR_SW = "CNL" Then
            Set dg_det2.DataSource = rsNada
        Else
        
            If VAR_SW = "ADD" Then
               rs_detalle2.Open "Select * from cv_diario_haber order by Cod_Comp_Detalle", db, adOpenKeyset, adLockOptimistic
            Else
                 
                rs_detalle2.Open "Select * from cv_diario_haber where Cod_Comp = " & Ado_datos.Recordset!Cod_Comp & " order by Cod_Comp_Detalle", db, adOpenKeyset, adLockOptimistic
            End If
            Set Ado_detalle2.Recordset = rs_detalle2
            If Ado_detalle2.Recordset.RecordCount > 0 Then
                dg_det2.Visible = True
                Set dg_det2.DataSource = rs_detalle2
              Else
                dg_det2.Visible = False
                Set dg_det2.DataSource = rsNada
            End If
        End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'error 6160 de acceso de datos
    'On Error GoTo error4
        'VAR_BUS = 0
    If (Ado_datos.Recordset.RecordCount > 0) Then 'Or (Ado_datos.Recordset.EOF) Or (Ado_datos.Recordset.BOF) Then
        If buscados = 0 Then
            OptSinAprobar.Visible = True
            OptTodos.Visible = True
        Else
            OptSinAprobar.Visible = False
            OptTodos.Visible = False
        End If
        If VAR_BUS = 0 Then
         'If VAR_SW <> "ADD" Then
              Call ABRIR_DEBE
              Call ABRIR_HABER
        ' End If
              VAR_BUS = 1
        End If
        If VAR_SW = "ADD" Or VAR_SW = "MOD" Then
            'gc_proceso_nivel3
            Set rs_datos7 = New ADODB.Recordset
            If rs_datos7.State = 1 Then rs_datos7.Close
            rs_datos7.Open "Select * from gc_proceso_nivel3 WHERE (proceso_codigo = 'FIN' AND subproceso_codigo <> 'FIN-02') ", db, adOpenStatic
            'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
            Set Ado_datos7.Recordset = rs_datos7
            dtc_desc7.BoundText = dtc_codigo7.BoundText
        Else
            Set rs_datos7 = New ADODB.Recordset
            If rs_datos7.State = 1 Then rs_datos7.Close
            rs_datos7.Open "Select * from gc_proceso_nivel3 WHERE (proceso_codigo = 'FIN') ", db, adOpenStatic
            'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
            Set Ado_datos7.Recordset = rs_datos7
            dtc_desc7.BoundText = dtc_codigo7.BoundText
            'Call ABRIR_DEBE
            'Call ABRIR_HABER
        End If
    Else
        ' ocultar grid
        MsgBox "Resultado sin datos", vbExclamation + vbDefaultButton1
        Exit Sub
    End If
   ' adiciona = "N"
    FraDet2.Visible = True
    FraDet1.Visible = True
    FrmABMDet1.Visible = True


'                Case "DAC", "PAC", "PCC", "ANL", "DVL", "RVT", "TRP", "PCO"
'                  mnuAnulacion.Enabled = False
'                  mnuDevolucion.Enabled = False
'                  mnuReversion.Enabled = False
'                Case "PCE"
'                  Dim rsestado As ADODB.Recordset
'                  Set rsestado = New ADODB.Recordset
'                  rsestado.CursorLocation = adUseClient
'                  rsestado.Open "select estado_pagado,estado_contabilidad from pagos where  codigo_pago=" & Val(rs_aux1!Cod_Comp) & " and org_codigo='" & _
'                                rs_aux1!org_codigo & "' and ges_gestion='" & rs_aux1!ges_gestion & "'", db, adOpenKeyset, adLockReadOnly
'                  If rsestado.RecordCount <> 0 Then
'                    If rsestado!estado_pagado = "S" Then
'                      mnuAnulacion.Enabled = True
'                      mnuDevolucion.Enabled = True
'                      mnuReversion.Enabled = False
'                    Else
'                        If rsestado!estado_contabilidad = "P" Then
'                           mnuAnulacion.Enabled = False
'                           mnuDevolucion.Enabled = False
'                           mnuReversion.Enabled = True
'                        Else
'                           mnuAnulacion.Enabled = False
'                           mnuDevolucion.Enabled = False
'                           mnuReversion.Enabled = False
'                        End If
'                    End If
'                  Else
'                      mnuAnulacion.Enabled = False
'                      mnuDevolucion.Enabled = False
'                      mnuReversion.Enabled = True
'                  End If
'                End Select
      
'    Else
'        MsgBox "Comprobantes sin datos", vbExclamation + vbDefaultButton1
'    End If
'error4:
'    If Err.Number = 383 Then
'        MsgBox "Comprobante con datos incorrectos", vbExclamation + vbDefaultButton1
  

End Sub

Private Sub Buscar4_Click()
VAR_AUX1 = H_Cta_Aux1_cmb
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And H_Cta_Aux1_cmb = "00" Then
        dtc_codigo11.Text = "0"
        dtc_desc11.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
         dtc_codigo11.Visible = True
        dtc_desc11.Visible = True
        Set rs_datos11 = New ADODB.Recordset
        If rs_datos11.State = 1 Then rs_datos11.Close
            rs_datos11.Open "Select " + VAR_CODIGO + " as codigo1 , " + VAR_DES + " as desc1 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos11.Recordset = rs_datos11
            dtc_desc11.BoundText = dtc_codigo11.BoundText
    End If
End Sub

Private Sub Buscar5_Click()
VAR_AUX1 = H_Cta_Aux2_cmb
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And H_Cta_Aux2_cmb = "00" Then
        dtc_codigo12.Text = "0"
        dtc_desc12.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
         dtc_codigo12.Visible = True
        dtc_desc12.Visible = True
        Set rs_datos12 = New ADODB.Recordset
        If rs_datos12.State = 1 Then rs_datos12.Close
            rs_datos12.Open "Select " + VAR_CODIGO + " as codigo2 , " + VAR_DES + " as desc2 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos12.Recordset = rs_datos12
            Set Ado_datos12.Recordset = rs_datos12
            dtc_desc12.BoundText = dtc_codigo12.BoundText
            
    End If
End Sub

Private Sub Buscar6_Click()
VAR_AUX1 = H_Cta_Aux3_cmb
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And H_Cta_Aux3_cmb = "00" Then
        dtc_codigo13.Text = "0"
        dtc_desc13.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
         dtc_codigo13.Visible = True
        dtc_desc13.Visible = True
        Set rs_datos13 = New ADODB.Recordset
        If rs_datos13.State = 1 Then rs_datos13.Close
            rs_datos13.Open "Select " + VAR_CODIGO + " as codigo3 , " + VAR_DES + " as desc3 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos13.Recordset = rs_datos13
            dtc_desc13.BoundText = dtc_codigo13.BoundText
    End If
End Sub

Private Sub cboNomTipo_Click(Area As Integer)
    CboTipo.BoundText = cboNomTipo.BoundText
'  Select Case Trim(CboTipo.Text)
'    Case "PCO"
'      ' TxtDBs.Enabled = True
'      '  TxtDSus.Enabled = True
'        Me.frameCAM.Visible = False
'        Me.DTPCAM.Visible = False
'        Me.txt_fecha.Visible = True
'        Me.txtcodsolicitud.Visible = False
'        Label26.Visible = False 'codigo solicitud
'       If adiciona = "S" Then
'        Me.dtc_codigo4.Text = "-"
'       End If
''        Me.lblDTC.Visible = True
'        lblHTC.Visible = True
'        lblHTIPOCAM.Visible = True
'        lblDTIPOCAM.Visible = True
'        lblDMonSus.Visible = True
'        lblHMONSUS.Visible = True
'        TxtDSus.Visible = True
'        txtHsus.Visible = True
''        Me.lblDTC.Visible = True
''        Me.lblDTC.Locked = False
'        '--
'        DtCDcodbenef.Visible = True
'        DtCDDescripbenef.Visible = True
'        DtCHDescripbenef.Visible = True
'        DtCHcodbenef.Visible = True
'        lblDBenefaux1.Visible = False
'        lblDnomBenefaux1.Visible = False
'        lblHBenefaux1.Visible = fALS
'        lblHnomBenefaux1.Visible = False
'        '----
'      If adiciona = "S" Then
''        Me.lblDTC = CTipoC
'        lblDTC_Change
'      End If
'
''        Me.CboDCtaCAM.Visible = False
''        Me.CboDSub1CAM.Visible = False
''        Me.CboDSub2CAM.Visible = False
'        Me.CboHCtaCAM.Visible = False
'        Me.CboHSub1CAM.Visible = False
'        Me.CboHSub2CAM.Visible = False
'        Me.frame_moneda.Enabled = True
'        CboDCta.Visible = True
'        CboDSubcta1.Visible = True
'        CboDSubcta2.Visible = True
'        CboHcta.Visible = True
'        CbohSubcta1.Visible = True
'        CbohSubcta2.Visible = True
'        optbolivianos_Click
'        TxtDBs = ""
'        TxtDSus = ""
'    Case "PCE"
'      '  TxtDBs.Enabled = True
'      '  TxtDSus.Enabled = True
'        Me.frameCAM.Visible = False
'        Me.DTPCAM.Visible = False
'        Me.txt_fecha.Visible = True
'        Me.txtcodsolicitud.Visible = True
'        Label26.Visible = True
''        Me.lblDTC.Visible = True
'        lblHTC.Visible = True
''        Me.lblDTC.Locked = True
'        '----------
'        DtCDcodbenef.Visible = False
'        DtCDDescripbenef.Visible = False
'        DtCHDescripbenef.Visible = False
'        DtCHcodbenef.Visible = False
'        lblDBenefaux1.Visible = True
'        lblDnomBenefaux1.Visible = True
'        lblHBenefaux1.Visible = True
'        lblHnomBenefaux1.Visible = True
'        '-----
'        'Me.lblDTC = CTipoC
'        If adiciona = "S" Then
''          Me.lblDTC = CTipoC
'          lblDTC_Change
'        End If
'        lblHTIPOCAM.Visible = True
'        lblDTIPOCAM.Visible = True
'        lblDMonSus.Visible = True
'        lblHMONSUS.Visible = True
'        TxtDSus.Visible = True
'        txtHsus.Visible = True
''        Me.lblDTC.Visible = True
''        Me.lblDTC.Locked = True
'        '---
'        lblDBenefaux1.Visible = True
'        lblDnomBenefaux1.Visible = True
'        '---
''        Me.CboDCtaCAM.Visible = False
''        Me.CboDSub1CAM.Visible = False
''        Me.CboDSub2CAM.Visible = False
'        Me.CboHCtaCAM.Visible = False
'        Me.CboHSub1CAM.Visible = False
'        Me.CboHSub2CAM.Visible = False
'        CboDCta.Visible = True
'        CboDSubcta1.Visible = True
'        CboDSubcta2.Visible = True
'        CboHcta.Visible = True
'        CbohSubcta1.Visible = True
'        CbohSubcta2.Visible = True
'        Me.frame_moneda.Enabled = True
'        'TxtDBs = ""
'        'TxtDSus = ""
'        optbolivianos_Click
'    Case "CAM"
'       ' TxtDBs.Enabled = True
'       ' TxtDSus.Enabled = True
'        If adiciona = "S" Then
'          Me.frameCAM.Visible = True
'        Else
'          Me.frameCAM.Visible = False
'        End If
'        Me.optCAMNo.Value = False
'        Me.optCAMSi.Value = False
'        Me.DTPCAM.Visible = True
'        Me.txt_fecha.Visible = False
'        Me.txtcodsolicitud.Visible = False
'        Label26.Visible = False 'codigo solicitud
'        Me.dtc_codigo4.Text = "-"
''        Me.lblDTC = "1.0"
'        lblHTC = "1.0"
'        '----
'        DtCDcodbenef.Visible = False
'        DtCDDescripbenef.Visible = False
'        DtCHDescripbenef.Visible = False
'        DtCHcodbenef.Visible = False
'        lblDBenefaux1.Visible = True
'        lblDnomBenefaux1.Visible = True
'        lblHBenefaux1.Visible = True
'        lblHnomBenefaux1.Visible = True
'        '----
''        Me.lblDTC.Visible = False
''        Me.lblDTC.Locked = True
'        lblHTC.Visible = False
'        lblHTIPOCAM.Visible = False
'        lblDTIPOCAM.Visible = False
'        'lblDMonSus.Visible = False
'        'lblHMONSUS.Visible = False
'        'Me.txtHsus.Visible = False
'        'Me.TxtDSus.Visible = False
'        'Me.TxtDSus = "0.0"
'        'Me.txtHsus = "0.0"
'        CboDCta.Visible = False
'        CboDSubcta1.Visible = False
'        CboDSubcta2.Visible = False
'        CboHcta.Visible = False
'        CbohSubcta1.Visible = False
'        CbohSubcta2.Visible = False
''        Me.CboDCtaCAM.Visible = True
''        Me.CboDSub1CAM.Visible = True
''        Me.CboDSub2CAM.Visible = True
'        Me.CboHCtaCAM.Visible = True
'        Me.CboHSub1CAM.Visible = True
'        Me.CboHSub2CAM.Visible = True
'
'        'Me.frame_moneda.Enabled = False
'        'Me.optbolivianos = True
'        optbolivianos_Click
'  End Select
'  ' Dim rsbustipo As ADODB.Recordset
'  ' Set rsbustipo = New ADODB.Recordset
'
'  rstipocomp.Filter = adFilterNone
'    rstipocomp.Filter = "Codigo_Tipo='" & Trim(CboTipo.Text) & "'"
'    If rstipocomp.RecordCount <> 0 Then
'        cboNomTipo.Text = rstipocomp!Denominacion_Tipo
'    End If

End Sub

Private Sub CboTipo_Click(Area As Integer)
    cboNomTipo.BoundText = CboTipo.BoundText
End Sub

Private Sub cmb_moneda_Click()
If cmb_moneda = "USD" Then
        D_MontoDl_cmb.Enabled = True
        D_MontoBs_cmb.Enabled = False
        D_MontoBs_cmb.Visible = True
        D_MontoDl_cmb.Visible = True
        D_Cambio_cmb.Visible = True
        Fram_AsientoH.Visible = True
'        Fram_AsientoH.Enabled = True
      

    Else
        D_MontoBs_cmb.Enabled = True
        D_MontoDl_cmb.Enabled = False
        D_MontoBs_cmb.Visible = True
        D_MontoDl_cmb.Visible = True
        D_Cambio_cmb.Visible = True
        Fram_AsientoH.Visible = True
'       Fram_AsientoH.Enabled = True

     End If
        Fram_AsientoH.Enabled = True

End Sub

Private Sub cmb_moneda_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub


Private Sub D_Cambio_cmb_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

'Private Sub D_Cambio_cmb_Change()
'Dim GlTipoCambioOficial As Currency
'End Sub

Private Sub D_Correl_cmb_Click(Area As Integer)
  D_Cuenta_cmb.BoundText = D_Correl_cmb.BoundText
  D_Nombre_cmb.BoundText = D_Correl_cmb.BoundText
  D_Subcta1_cmb.BoundText = D_Correl_cmb.BoundText
  D_Subcta2_cmb.BoundText = D_Correl_cmb.BoundText
  D_Cta_Aux1_cmb.BoundText = D_Correl_cmb.BoundText
  D_Cta_Aux2_cmb.BoundText = D_Correl_cmb.BoundText
  D_Cta_Aux3_cmb.BoundText = D_Correl_cmb.BoundText
End Sub

Private Sub D_Cta_Aux1_cmb_Click(Area As Integer)
  D_Correl_cmb.BoundText = D_Cta_Aux1_cmb.BoundText
  D_Cuenta_cmb.BoundText = D_Cta_Aux1_cmb.BoundText
  D_Nombre_cmb.BoundText = D_Cta_Aux1_cmb.BoundText
  D_Subcta1_cmb.BoundText = D_Cta_Aux1_cmb.BoundText
  D_Subcta2_cmb.BoundText = D_Cta_Aux1_cmb.BoundText
  D_Cta_Aux2_cmb.BoundText = D_Cta_Aux1_cmb.BoundText
  D_Cta_Aux3_cmb.BoundText = D_Cta_Aux1_cmb.BoundText
End Sub

Private Sub D_Cta_Aux2_cmb_Click(Area As Integer)
  D_Correl_cmb.BoundText = D_Cta_Aux2_cmb.BoundText
  D_Cuenta_cmb.BoundText = D_Cta_Aux2_cmb.BoundText
  D_Nombre_cmb.BoundText = D_Cta_Aux2_cmb.BoundText
  D_Subcta1_cmb.BoundText = D_Cta_Aux2_cmb.BoundText
  D_Subcta2_cmb.BoundText = D_Cta_Aux2_cmb.BoundText
  D_Cta_Aux1_cmb.BoundText = D_Cta_Aux2_cmb.BoundText
  D_Cta_Aux3_cmb.BoundText = D_Cta_Aux2_cmb.BoundText
End Sub

Private Sub D_Cta_Aux3_cmb_Click(Area As Integer)
  D_Correl_cmb.BoundText = D_Cta_Aux3_cmb.BoundText
  D_Cuenta_cmb.BoundText = D_Cta_Aux3_cmb.BoundText
  D_Nombre_cmb.BoundText = D_Cta_Aux3_cmb.BoundText
  D_Subcta1_cmb.BoundText = D_Cta_Aux3_cmb.BoundText
  D_Subcta2_cmb.BoundText = D_Cta_Aux3_cmb.BoundText
  D_Cta_Aux1_cmb.BoundText = D_Cta_Aux3_cmb.BoundText
  D_Cta_Aux2_cmb.BoundText = D_Cta_Aux3_cmb.BoundText
End Sub

Private Sub D_Cuenta_cmb_Click(Area As Integer)
  D_Correl_cmb.BoundText = D_Cuenta_cmb.BoundText
  D_Nombre_cmb.BoundText = D_Cuenta_cmb.BoundText
  D_Subcta1_cmb.BoundText = D_Cuenta_cmb.BoundText
  D_Subcta2_cmb.BoundText = D_Cuenta_cmb.BoundText
  D_Cta_Aux1_cmb.BoundText = D_Cuenta_cmb.BoundText
  D_Cta_Aux2_cmb.BoundText = D_Cuenta_cmb.BoundText
  D_Cta_Aux3_cmb.BoundText = D_Cuenta_cmb.BoundText
End Sub

Private Sub D_MontoBs_cmb_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46 Then
    Exit Sub
    Else
    KeyAscii = 0
    End If
End Sub

Private Sub D_MontoBs_cmb_LostFocus()
    If cmb_moneda = "BOB" Then
       
        D_MontoDl_cmb.Text = Round(CDbl(IIf(D_MontoBs_cmb.Text = "", "0", D_MontoBs_cmb.Text)) / CDbl(D_Cambio_cmb), 2)
    
    Else
         D_MontoBs_cmb.Text = Round(CDbl(IIf(D_MontoDl_cmb.Text = "", "0", D_MontoDl_cmb.Text)) * CDbl(D_Cambio_cmb), 2)
    
     End If
End Sub

Private Sub D_MontoDl_cmb_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46 Then
    Exit Sub
    Else
    KeyAscii = 0
    End If
End Sub

Private Sub D_MontoDl_cmb_LostFocus()
If cmb_moneda = "USD" Then
       
        D_MontoBs_cmb.Text = Round(CDbl(IIf(D_MontoDl_cmb.Text = "", "0", D_MontoDl_cmb.Text)) * CDbl(D_Cambio_cmb), 2)
    
    Else
         D_MontoDl_cmb.Text = Round(CDbl(IIf(D_MontoBs_cmb.Text = "", "0", D_MontoBs_cmb.Text)) / CDbl(D_Cambio_cmb), 2)
    
     End If
End Sub

Private Sub D_Nombre_cmb_Click(Area As Integer)
  D_Correl_cmb.BoundText = D_Nombre_cmb.BoundText
  D_Cuenta_cmb.BoundText = D_Nombre_cmb.BoundText
  D_Subcta1_cmb.BoundText = D_Nombre_cmb.BoundText
  D_Subcta2_cmb.BoundText = D_Nombre_cmb.BoundText
  D_Cta_Aux1_cmb.BoundText = D_Nombre_cmb.BoundText
  D_Cta_Aux2_cmb.BoundText = D_Nombre_cmb.BoundText
  D_Cta_Aux3_cmb.BoundText = D_Nombre_cmb.BoundText
End Sub

Private Sub D_Subcta1_cmb_Click(Area As Integer)
  D_Correl_cmb.BoundText = D_Subcta1_cmb.BoundText
  D_Cuenta_cmb.BoundText = D_Subcta1_cmb.BoundText
  D_Nombre_cmb.BoundText = D_Subcta1_cmb.BoundText
  D_Subcta2_cmb.BoundText = D_Subcta1_cmb.BoundText
  D_Cta_Aux1_cmb.BoundText = D_Subcta1_cmb.BoundText
  D_Cta_Aux2_cmb.BoundText = D_Subcta1_cmb.BoundText
  D_Cta_Aux3_cmb.BoundText = D_Subcta1_cmb.BoundText
End Sub

Private Sub D_Subcta2_cmb_Click(Area As Integer)
  D_Correl_cmb.BoundText = D_Subcta2_cmb.BoundText
  D_Cuenta_cmb.BoundText = D_Subcta2_cmb.BoundText
  D_Nombre_cmb.BoundText = D_Subcta2_cmb.BoundText
  D_Subcta1_cmb.BoundText = D_Subcta2_cmb.BoundText
  D_Cta_Aux1_cmb.BoundText = D_Subcta2_cmb.BoundText
  D_Cta_Aux2_cmb.BoundText = D_Subcta2_cmb.BoundText
  D_Cta_Aux3_cmb.BoundText = D_Subcta2_cmb.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo12_Click(Area As Integer)
    dtc_desc12.BoundText = dtc_codigo12.BoundText
End Sub

Private Sub dtc_codigo13_Click(Area As Integer)
    dtc_desc13.BoundText = dtc_codigo13.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc12_Click(Area As Integer)
    dtc_codigo12.BoundText = dtc_desc12.BoundText
End Sub

Private Sub dtc_desc13_Click(Area As Integer)
    dtc_codigo13.BoundText = dtc_desc13.BoundText
End Sub

Private Sub dtc_desc14_Click(Area As Integer)
    txt_codigo1.BoundText = dtc_desc14.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

'    If rsbef.RecordCount = 0 Then
'        MsgBox "El beneficiario no existe. Seleccione un beneficiario", vbExclamation + vbDefaultButton1
'        'Me.dtc_desc4.SetFocus
'        Exit Sub
'    End If
'End Sub
'
'Private Sub dtcbodocumento1_Change()
'   dtcbodocumento2.BoundText = dtcbodocumento1.BoundText
'End Sub
'
'Private Sub dtcbodocumento1_Click(Area As Integer)
'    dtcbodocumento2.BoundText = dtcbodocumento1.BoundText
'End Sub
'
'Private Sub dtcbodocumento2_Change()
' dtcbodocumento1.BoundText = dtcbodocumento2.BoundText
'End Sub
'
'Private Sub dtcbodocumento2_Click(Area As Integer)
'    dtcbodocumento1.BoundText = dtcbodocumento2.BoundText
'End Sub

'Private Sub DtCDcodbenef_Change()
'If CboTipo = "PCO" Then
'  DtCHcodbenef.Text = dtc_codigo4.Text
'  DtCHcodbenef_Click (1)
'End If
'End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
  dtc_desc4.BoundText = dtc_codigo4.BoundText
'Me.dtc_codigo4.BoundText = Me.dtc_desc4.BoundText
End Sub

'Private Sub DTCDDesCaja_Click(Area As Integer)
' dtcDIdCaja.Text = DTCDDesCaja.BoundText
''  dtcDIdCaja.Text = Trim(DTCDDesCaja.BoundText)
'End Sub

'Private Sub DtCDDesConvenio_Change()
'  DtCDIdConvenio.BoundText = DtCDDesConvenio.BoundText
'End Sub

Private Sub DtCDIDCaja_Click(Area As Integer)
  DTCDDesCaja.Text = dtcDIdCaja.BoundText
  'DTCDDesCaja.Text = Trim(dtcDIdCaja.BoundText)
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
dtc_codigo8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
dtc_codigo9.BoundText = dtc_desc9.BoundText
End Sub


Private Sub DtCHcodbenef_Click(Area As Integer)
  DtCHDescripbenef.BoundText = DtCHcodbenef.BoundText
End Sub

'Private Sub DTCHDesCaja_Click(Area As Integer)
'DTCHidcaja.BoundText = DTCHDesCaja.BoundText
''  DTCHidcaja.BoundText = DTCHDesCaja.BoundText
'End Sub
'
'Private Sub DtCHDesConvenio_Change()
'  DtCHIdConvenio.BoundText = DtCHDesConvenio.BoundText
'End Sub

Private Sub DtCHDescripbenef_Click(Area As Integer)
  DtCHcodbenef.BoundText = DtCHDescripbenef.BoundText
End Sub

Private Sub DtCDIdConvenio_Change()
 DtCDDesConvenio.BoundText = DtCDIdConvenio.BoundText
'dctalarga = Trim(DtCDIdConvenio.Text)
End Sub

'Private Sub DtCIdConvenio_Click(Area As Integer)
'  DtCDDesConvenio.BoundText = DtCDIdConvenio.BoundText
'  dctalarga = Trim(DtCDIdConvenio.Text)
'End Sub

Private Sub DtCHIdCaja_Click(Area As Integer)
  'DTCHDesCaja.BoundText = DTCHidcaja.BoundText
  DTCHDesCaja.Text = Trim(DTCHidcaja.BoundText)
End Sub

Private Sub DtCHIdConvenio_Change()
  DtCHDesConvenio.BoundText = DtCHIdConvenio.BoundText
'  hctalarga = Trim(DtCHIdConvenio.Text)
End Sub

'Private Sub DtCHIdConvenio_Click(Area As Integer)
'  DtCHDesConvenio.BoundText = DtCHIdConvenio.BoundText
'  hctalarga = Trim(DtCHIdConvenio.Text)
'End Sub
'
Private Sub dg_datos_Click()
    VAR_BUS = 0
''error 6160 de acceso de datos
'    'On Error GoTo error4
'    Fram_AsientoD.Enabled = True
'    Fram_AsientoH.Enabled = True
'    'TDBFrameDebe.Enabled = False
'    'TDBFrameDebeCta.Enabled = False
'    If (rs_datos.RecordCount = 0) Or (rs_datos.EOF) Or (rs_datos.BOF) Then
'      Exit Sub
'    End If
'    Call limpiar
''    If rs_datos.EOF = True And rs_datos.BOF = True Then
' '       Exit Sub
'  '  End If
'    Me.TxtComprobante = rs_datos!Cod_Comp 'Me.dg_datos.Columns(0).Value
'    adiciona = "N"
'    'Me.BtnModificar.Enabled = True
'    Set rs_aux1 = New ADODB.Recordset
'    If rs_aux1.State = 1 Then rs_datos.Close
'    rs_aux1.Open "SELECT Co_Comprobante_M.*," & _
'            "CO_Diario.* " & _
'            " FROM Co_Comprobante_M INNER JOIN " & _
'            "CO_Diario ON  Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp AND " & _
'            " Co_Comprobante_M.Tipo_Comp = CO_Diario.Tipo_Comp where " & _
'            " co_comprobante_M.cod_comp=" & Val(rs_datos!Cod_Comp) & _
'            " and Co_Comprobante_M.Tipo_Comp='" & Trim(rs_datos!tipo_comp) & "'", db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount <> 0 Then
'        Me.CboTipo = rs_aux1!tipo_comp
'        'CboTipo_Click
'        Me.txt_ges = rs_aux1!ges_gestion
'        Me.txtcodsolicitud = IIf(IsNull(rs_aux1!codigo_solicitud), "", rs_aux1!codigo_solicitud)
'        'Me.txt_fecha = IIf(IsNull(rs_aux1!Fecha_transacion), "", Format(rs_aux1!Fecha_transacion, "dd/mm/yyyy"))
'        Me.txt_codigo1.Text = rs_aux1!codigo_documento
'        Me.txt_campo1 = IIf(IsNull(rs_aux1!num_respaldo), "", rs_aux1!num_respaldo)
'        Me.dtc_codigo4.Text = IIf(IsNull(rs_aux1!beneficiario_codigo), "-", rs_aux1!beneficiario_codigo)
'        Me.Txt_glosa = IIf(IsNull(rs_aux1!glosa), "", rs_aux1!glosa)
'        'On Error Resume Next
'        '*****tipo de comprobante
'         If Trim(rs_aux1!tipo_comp) = "CAM" Then
'            Me.DTPCAM.Visible = True
'            Me.txt_fecha.Visible = False
'            Me.DTPCAM.Value = IIf(IsNull(rs_aux1!Fecha_transacion), Date, Format(rs_aux1!Fecha_transacion, "dd/mm/yyyy"))
'            Me.lblDTC.Visible = False
'            lblHTC.Visible = False
'            lblHTIPOCAM.Visible = False
'            lblDTIPOCAM.Visible = False
'            lblDMonSus.Visible = False
'            lblHMONSUS.Visible = False
'            Me.txtHsus.Visible = False
'            Me.TxtDSus.Visible = False
'            Me.CboDCta.Visible = False
'            Me.CboDSubcta1.Visible = False
'            Me.CboDSubcta2.Visible = False
'            Me.CboHcta.Visible = False
'            Me.CbohSubcta1.Visible = False
'            Me.CbohSubcta2.Visible = False
'            Me.CboDCtaCAM.Visible = True
'            Me.CboDSub1CAM.Visible = True
'            Me.CboDSub2CAM.Visible = True
'            Me.CboHCtaCAM.Visible = True
'            Me.CboHSub1CAM.Visible = True
'            Me.CboHSub2CAM.Visible = True
'            Me.CboHCtaCAM = IIf(IsNull(rs_aux1!h_cuenta), "", rs_aux1!h_cuenta)
'            Me.CboHSub1CAM = IIf(IsNull(rs_aux1!h_subcta1), "", rs_aux1!h_subcta1)
'            Me.CboHSub2CAM = IIf(IsNull(rs_aux1!h_subcta2), "", rs_aux1!h_subcta2)
'            CboHSub2CAM_Change
'            Me.CboDCtaCAM = IIf(IsNull(rs_aux1!d_cuenta), "", rs_aux1!d_cuenta)
'            Me.CboDSub1CAM = IIf(IsNull(rs_aux1!d_subcta1), "", rs_aux1!d_subcta1)
'            Me.CboDSub2CAM = IIf(IsNull(rs_aux1!d_subcta2), "", rs_aux1!d_subcta2)
'            CboDSub2CAM_Change
'         Else
'            Me.DTPCAM.Visible = False
'            Me.txt_fecha.Visible = True
'            Me.txt_fecha = IIf(IsNull(rs_aux1!Fecha_transacion), "", Format(rs_aux1!Fecha_transacion, "dd/mm/yyyy"))
'            Me.lblDTC.Visible = True
'            lblHTC.Visible = True
'            lblHTIPOCAM.Visible = True
'            lblDTIPOCAM.Visible = True
'            lblDMonSus.Visible = True
'            lblHMONSUS.Visible = True
'            TxtDSus.Visible = True
'            txtHsus.Visible = True
'            Me.lblDTC.Visible = True
'            Me.CboDCta.Visible = True
'            Me.CboDSubcta1.Visible = True
'            Me.CboDSubcta2.Visible = True
'            Me.CboHcta.Visible = True
'            Me.CbohSubcta1.Visible = True
'            Me.CbohSubcta2.Visible = True
'            Me.CboDCtaCAM.Visible = False
'            Me.CboDSub1CAM.Visible = False
'            Me.CboDSub2CAM.Visible = False
'            Me.CboHCtaCAM.Visible = False
'            Me.CboHSub1CAM.Visible = False
'            Me.CboHSub2CAM.Visible = False
'            Me.CboHcta = IIf(IsNull(rs_aux1!h_cuenta), "", rs_aux1!h_cuenta)
'            Me.CbohSubcta1 = IIf(IsNull(rs_aux1!h_subcta1), "", rs_aux1!h_subcta1)
'            Me.CbohSubcta2 = IIf(IsNull(rs_aux1!h_subcta2), "", rs_aux1!h_subcta2)
'            CbohSubcta2_Change
'            activdatosHaber
'            Me.CboDCta = IIf(IsNull(rs_aux1!d_cuenta), "", rs_aux1!d_cuenta)
'            Me.CboDSubcta1 = IIf(IsNull(rs_aux1!d_subcta1), "", rs_aux1!d_subcta1)
'            Me.CboDSubcta2 = IIf(IsNull(rs_aux1!d_subcta2), "", rs_aux1!d_subcta2)
'            CboDSubcta2_Change
'            activdatosdebe
'         End If
'
'        Me.lblHTC = IIf(IsNull(rs_aux1!h_Cambio), "1", Val(rs_aux1!h_Cambio))
'        If Val(Trim(lblHTC)) = 0 Then
'            lblDTC = "1"
'        End If
'        Me.txtHBs = IIf(IsNull(rs_aux1!d_montoBs), "", Val(rs_aux1!d_montoBs))
'        Me.txtHsus = IIf(IsNull(rs_aux1!h_montoDl), "", Val(rs_aux1!h_montoDl))
'        '-----'
'        If IIf(IsNull(rs_aux1!h_Aux1), "", rs_aux1!h_Aux1) <> "00" Then
'          DatosHaber IIf(IsNull(rs_aux1!h_Aux1), "", rs_aux1!h_Aux1), IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
''          SSTabHaber.TabEnabled(0) = True
'        End If
'        If IIf(IsNull(rs_aux1!h_Aux2), "", rs_aux1!h_Aux2) <> "00" Then
'          DatosHaber IIf(IsNull(rs_aux1!h_Aux2), "", rs_aux1!h_Aux2), IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
''          SSTabHaber.TabEnabled(1) = True
'        End If
'        If IIf(IsNull(rs_aux1!h_Aux3), "", rs_aux1!h_Aux3) <> "00" Then
'          DatosHaber IIf(IsNull(rs_aux1!h_Aux3), "", rs_aux1!h_Aux3), IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
''          SSTabHaber.TabEnabled(0) = True
'        End If
'        '-----'
'        If IIf(IsNull(rs_aux1!d_Aux1), "", rs_aux1!d_Aux1) <> "00" Then
'          DatosDebe IIf(IsNull(rs_aux1!d_Aux1), "", rs_aux1!d_Aux1), IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
''          SSTabDebe.TabEnabled(0) = True
'        End If
'        If IIf(IsNull(rs_aux1!d_Aux2), "", rs_aux1!d_Aux2) <> "00" Then
'          DatosDebe IIf(IsNull(rs_aux1!d_Aux2), "", rs_aux1!d_Aux2), IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
''          SSTabDebe.TabEnabled(1) = True
'        End If
'       If IIf(IsNull(rs_aux1!d_Aux3), "", rs_aux1!d_Aux3) <> "00" Then
'          DatosDebe IIf(IsNull(rs_aux1!d_Aux3), "", rs_aux1!d_Aux3), IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
''          SSTabDebe.TabEnabled(2) = True
'        End If
'        '-----
''        Select Case IIf(IsNull(rs_aux1!h_Aux1), "", rs_aux1!h_Aux1)
''            Case "00"
''                Me.FrameHBeneficiario.Visible = False
''                Me.frameHCtaBancaria.Visible = False
''                Me.frameHAux00.Visible = True
''                Me.frameHOrganismos.Visible = False
''            Case "01"
''                Me.frameHOrganismos.Visible = False
''                Me.FrameHBeneficiario.Visible = True
''                Me.frameHCtaBancaria.Visible = False
''                Me.frameHAux00.Visible = False
''                Me.lblHBenefaux1 = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
''                Call buscabenef(IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1))
''                hctalarga = Me.lblHBenefaux1
''                Me.lblHnomBenefaux1 = Trim(Cdenominacion)
''            '**buscar nombre beneficiario
''            Case "02"
''                Me.frameHOrganismos.Visible = False
''                Me.FrameHBeneficiario.Visible = False
''                Me.frameHAux00.Visible = False
''                Me.frameHCtaBancaria.Visible = True
''                Me.cboHctaaux1 = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
''                Call buscactabancaria(Trim(rs_aux1!H_Cta_Aux1))
''                Me.cboHctanomaux1 = cdenomctabancaria
''                hctalarga = Me.cboHctaaux1
''            Case "08"
''                Me.FrameHBeneficiario.Visible = False
''                Me.frameHAux00.Visible = False
''                Me.frameHCtaBancaria.Visible = False
''                frameHOrganismos.Visible = True
''                Me.cboHCodOrg = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
''                ''Call buscactabancaria(Trim(rs_aux1!H_Cta_Aux1))
''                Call buscaorganismo(Trim(cboHCodOrg.Text))
''                hctalarga = Me.cboHCodOrg
''                Me.cboHDenomOrg = Me.denomorgan
''            '***buscar nombre de la cuenta
''            Case Else
''                Me.FrameHBeneficiario.Visible = False
''                Me.frameHCtaBancaria.Visible = False
''                Me.frameHAux00.Visible = True
''                Me.frameHOrganismos.Visible = False
''                hctalarga = ""
''        End Select
'
'        '-----
'       ' Me.cboh_aux1_denom.Text = rs_aux1!H_Des_Larga
'        Me.lblDTC = IIf(IsNull(rs_aux1!d_Cambio), "1", rs_aux1!d_Cambio)
'        If Val(Trim(lblDTC)) = 0 Then
'            lblDTC = "1"
'        End If
'        Me.TxtDBs = IIf(IsNull(rs_aux1!d_montoBs), "", Val(rs_aux1!d_montoBs))
'        Me.TxtDSus = IIf(IsNull(rs_aux1!d_montoDl), "", Val(rs_aux1!d_montoDl))
''        Select Case IIf(IsNull(rs_aux1!d_Aux1), "", rs_aux1!d_Aux1)
''        Case "00"
''            Me.FrameDBeneficiario.Visible = False
''            Me.frameDCtaBancaria.Visible = False
''            Me.frameDOrganismos.Visible = False
''            Me.frameDaux00.Visible = True
''            dctalarga = ""
''        Case "01"
''            Me.frameDOrganismos.Visible = False
''            Me.frameDCtaBancaria.Visible = False
''            Me.frameDaux00.Visible = False
''            Me.FrameDBeneficiario.Visible = True
''            Me.lblDBenefaux1 = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
''            Call buscabenef(rs_aux1!D_Cta_Aux1)
''            Me.lblDnomBenefaux1 = Trim(Cdenominacion)
''            dctalarga = Me.lblDBenefaux1
''        Case "02"
''            Me.frameDOrganismos.Visible = False
''            Me.frameDaux00.Visible = False
''            Me.FrameDBeneficiario.Visible = False
''            Me.frameDCtaBancaria.Visible = True
''            Me.cboDctaaux1 = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
''            Call buscactabancaria(Trim(rs_aux1!D_Cta_Aux1))
''            Me.cboDctanomaux1 = cdenomctabancaria
''            dctalarga = Me.cboDctaaux1
''        Case "08"
''            Me.frameDaux00.Visible = False
''            Me.FrameDBeneficiario.Visible = False
''            Me.frameDCtaBancaria.Visible = True
''            frameDOrganismos.Visible = True
''            Me.cboDCodOrg = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
''            ''Call buscactabancaria(Trim(rs_aux1!H_Cta_Aux1))
''            Call buscaorganismo(Trim(cboDCodOrg.Text))
''            Me.cboDDenomOrg = Me.denomorgan
''            dctalarga = Me.cboDCodOrg
''        Case Else
''            Me.FrameDBeneficiario.Visible = False
''            Me.frameDCtaBancaria.Visible = False
''            Me.frameDaux00.Visible = True
''            Me.frameDOrganismos.Visible = False
''            dctalarga = ""
''        End Select
'    'Tipo de moneda
'        Select Case IIf(IsNull(rs_aux1!tipo_moneda), " ", rs_aux1!tipo_moneda)
'            Case "Bs"
'                Me.optbolivianos.Value = True
'                optbolivianos_Click
'            Case "$US"
'                Me.optdolares.Value = True
'                optdolares_Click
'            Case " ", ""  'las transacciones anteriores se realizaran  por defecto en Bolivianos
'                Me.optbolivianos.Value = True
'                optbolivianos_Click
'        End Select
'    'Me.cbod_aux1_denom.Text = rs_aux1!D_Des_Larga
'        If rs_aux1!Status = "S" Or rs_aux1!Status = "A" Then
'              Me.BtnModificar.Enabled = False
'              Me.BtnEliminar.Enabled = False
'              'Me.BtnDesAprobar.Enabled = False
'              Select Case rs_aux1!tipo_comp
'                Case "DAC", "PAC", "PCC", "ANL", "DVL", "RVT", "TRP", "PCO"
'                  mnuAnulacion.Enabled = False
'                  mnuDevolucion.Enabled = False
'                  mnuReversion.Enabled = False
'                Case "PCE"
'                  Dim rsestado As ADODB.Recordset
'                  Set rsestado = New ADODB.Recordset
'                  rsestado.CursorLocation = adUseClient
'                  rsestado.Open "select estado_pagado,estado_contabilidad from pagos where  codigo_pago=" & Val(rs_aux1!Cod_Comp) & " and org_codigo='" & _
'                                rs_aux1!org_codigo & "' and ges_gestion='" & rs_aux1!ges_gestion & "'", db, adOpenKeyset, adLockReadOnly
'                  If rsestado.RecordCount <> 0 Then
'                    If rsestado!estado_pagado = "S" Then
'                      mnuAnulacion.Enabled = True
'                      mnuDevolucion.Enabled = True
'                      mnuReversion.Enabled = False
'                    Else
'                        If rsestado!estado_contabilidad = "P" Then
'                           mnuAnulacion.Enabled = False
'                           mnuDevolucion.Enabled = False
'                           mnuReversion.Enabled = True
'                        Else
'                           mnuAnulacion.Enabled = False
'                           mnuDevolucion.Enabled = False
'                           mnuReversion.Enabled = False
'                        End If
'                    End If
'                  Else
'                      mnuAnulacion.Enabled = False
'                      mnuDevolucion.Enabled = False
'                      mnuReversion.Enabled = True
'                  End If
'                End Select
'        End If
'        Select Case rs_aux1!tipo_comp
'          'Case "PAC", "DAC", "ANL", "DVL", "RVT", "CAD", "CAR", "PCC"
'          Case "PCE", "PCO"
'            BtnDesAprobar.Enabled = True
'          Case Else
'            BtnDesAprobar.Enabled = False
'        End Select
'        If rs_aux1!Status = "N" Then
'              Me.BtnModificar.Enabled = True
'              'Me.BtnDesAprobar.Enabled = True
'              Me.BtnEliminar.Enabled = True
'              mnuAnulacion.Enabled = False
'              mnuDevolucion.Enabled = False
'              mnuReversion.Enabled = False
'        End If
''      SSTabDebe_Click (0)
''      SSTabHaber_Click (0)
'    Else
'        MsgBox "Comprobantes sin datos", vbExclamation + vbDefaultButton1
'    End If
'error4:
'    If Err.Number = 383 Then
'        MsgBox "Comprobante con datos incorrectos", vbExclamation + vbDefaultButton1
'    End If
End Sub

'Private Sub dg_datos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' If Button = vbRightButton Then Me.PopupMenu mnumenu
'End Sub
'
'
'Private Sub dg_datos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'  dg_datos_Click
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
'
'End Sub

Private Sub DTPfecha_trans_LostFocus()
    If adiciona = "S" Then
        txt_ges.Text = Year(DTPfecha_trans.Value)
        txt_mes.Text = UCase(MonthName(Month(DTPfecha_trans.Value)))
    End If
End Sub

Private Sub Form_Load()
    adiciona = ""
    'LblUsuario.Caption = Trim(GlUsuario)
'    DTPCAM.Value = CFecha
'    DTPCAM.MaxDate = Date
'    DTPCAM.Visible = False
'    Me.sstab1.Tab = 0
'    Me.frame_moneda.Visible = True
    VAR_SW = ""
    VAR_BUS = 0
    buscados = 0
    Me.FraGrabarCancelar.Visible = False
    Me.fraOpciones.Visible = True
    Me.FraGlobal.Enabled = False
    'Me.Fram_AsientoD.Enabled = False
   ' TDBFrameDebeCta.Enabled = False
    'TDBFrameDebe.Enabled = False
'    TDBFrameHaber.Enabled = False
'    TDBFrameHaberCta.Enabled = False
    'Me.Fram_AsientoH.Enabled = False

    'Me.Cmd_GrabaM.Enabled = False
    'me.frame
    Set rs_datos_M = New ADODB.Recordset
'    Set rsdiario = New ADODB.Recordset
'    Set rsPlan_cuentas = New ADODB.Recordset
'    Set rsplanctas = New ADODB.Recordset
    Set rscuentas = New ADODB.Recordset
    Set rssubcuenta = New ADODB.Recordset
    Set rsmoneda = New ADODB.Recordset
'    Set rsOrganismo = New ADODB.Recordset
    '*************recordset para el grid inicial
    Call OptSinAprobar_Click
    'Set rs_datos = New ADODB.Recordset
    'If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "Select * " & _
    '               "from CO_comprobante_M where estado_codigo='REG' "
    'rs_datos.Open queryinicial, db, adOpenKeyset, adLockReadOnly
    ''rs_datos.Sort = "cod_comp ASC"
    'Set Me.dg_datos.DataSource = rs_datos
    Call ABRIR_TABLAS_AUX
'    Me.frame_moneda.Enabled = False
    'Me.DTPCAM.Enabled = False
    'Me.DTPCAM.Value = CFecha

    OptSinAprobar.Value = True
    OptSinAprobar_Click
End Sub

Private Sub ABRIR_TABLAS_AUX()
    ' UNIDAD EJECUTORA
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    'rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    ' TIPOS DE COMPROBANTES (CONTROL SOFIA)
    Set rstipocomp = New ADODB.Recordset
    If rstipocomp.State = 1 Then rstipocomp.Close
    rstipocomp.Open "Select * from gc_tipo_comprobante where estado_codigo='APR' order by Denominacion_Tipo", db, adOpenStatic
    Set Ado_datos2.Recordset = rstipocomp
    cboNomTipo.BoundText = CboTipo.BoundText
    
    ' TIPOS DE COMPROBANTES CONTABLES ISO
    Set rs_datos14 = New ADODB.Recordset
    If rs_datos14.State = 1 Then rs_datos14.Close
    rs_datos14.Open "Select * from gc_documentos_respaldo where doc_codigo = 'R-110' OR doc_codigo = 'R-111' OR doc_codigo = 'R-112' order by doc_descripcion", db, adOpenStatic
    Set Ado_datos14.Recordset = rs_datos14
    dtc_desc14.BoundText = txt_codigo1.BoundText
    
    'where (tipoben_codigo < 20 and tipoben_codigo <> 1)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_beneficiario where estado_codigo='APR' order by beneficiario_denominacion", db, adOpenStatic
    'rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    
    '******se carga de los COMBO CUENTAS  -------------   and estado_codigo='APR'
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "SELECT Cuenta +'-'+SubCta1+'-'+SubCta2+'- '+ltrim(NombreCta) as NombreCtaAux,* FROM CC_Plan_Cuentas WHERE Nivel = '5' ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos3.Recordset = rs_datos3
    
    D_Cuenta_cmb.BoundText = D_Correl_cmb.BoundText
    D_Nombre_cmb.BoundText = D_Correl_cmb.BoundText
    D_Subcta1_cmb.BoundText = D_Correl_cmb.BoundText
    D_Subcta2_cmb.BoundText = D_Correl_cmb.BoundText
    D_Cta_Aux1_cmb.BoundText = D_Correl_cmb.BoundText
    D_Cta_Aux2_cmb.BoundText = D_Correl_cmb.BoundText
    D_Cta_Aux3_cmb.BoundText = D_Correl_cmb.BoundText
    
    H_Cuenta_cmb.BoundText = H_Correl_cmb.BoundText
    H_Nombre_cmb.BoundText = H_Correl_cmb.BoundText
    H_Subcta1_cmb.BoundText = H_Correl_cmb.BoundText
    H_Subcta2_cmb.BoundText = H_Correl_cmb.BoundText
    H_Cta_Aux1_cmb.BoundText = H_Correl_cmb.BoundText
    H_Cta_Aux2_cmb.BoundText = H_Correl_cmb.BoundText
    H_Cta_Aux3_cmb.BoundText = H_Correl_cmb.BoundText
    
'    Do While Not Ado_datos3.EOF
'        Me.CboHcta.AddItem Ado_datos3!cuenta
''        Me.CboDCta.AddItem rsplanctas!cuenta
'        Ado_datos3.MoveNext
'    Loop
    
    '**********recordset para cuentas bancarias
'    Set rscta_corrienteDebe = New ADODB.Recordset
'    Set rscta_corrienteHaber = New ADODB.Recordset
'    Set rscta_corriente = New ADODB.Recordset
'    If rscta_corriente.State = 1 Then rscta_corriente.Close
'    rscta_corriente.Open "SELECT fc_cuenta_bancaria.Cta_codigo,fc_cuenta_bancaria.cta_descripcion FROM fc_cuenta_bancaria " & _
'                        "order by cta_codigo", db, adOpenForwardOnly, adLockReadOnly
    'Me.OptSinAprobar.Value = True
    '*****se carga los combos para el comprobante  de diferencias cambiarias
'    Me.CboDCtaCAM.AddItem "1111"
    'Me.CboDCtaCAM.AddItem = "5174"
'    Me.CboDCtaCAM.AddItem "6141"
   ' CboDCtaCAM.Text = CboDCtaCAM.List(0)
    
    '******tipo de cambio
'    Set rstipocambio = New ADODB.Recordset
'    sql_TC = "select fecha_cambio, cambio_oficial_compra  from gc_tipo_cambio  where fecha_cambio = (select max(fecha_cambio) as expr1 from gc_tipo_cambio)"
'    rstipocambio.Open sql_TC, db, adOpenKeyset, adLockReadOnly
'    CTipoC = rstipocambio!cambio_oficial_compra
'    CFecha = rstipocambio!fecha_cambio
'    '*****tipo de moneda
'    If rsmoneda.State = 1 Then rsmoneda.Close
'    rsmoneda.Open "select * from gc_tipo_moneda", db, adOpenForwardOnly, adLockReadOnly
'    If rsmoneda.RecordCount <> 0 Then
'        rsmoneda.MoveFirst
'        rsmoneda.Find "pais_codigo='BOL'"  'moneda de Bolivia
'        CmonedaBs = rsmoneda!tipo_moneda
'        rsmoneda.MoveFirst
'        rsmoneda.Find "pais_codigo='USA'"
'        CmonedaSus = rsmoneda!tipo_moneda  'moneda americana
'    Else
'        MsgBox "Revise los datos de monedas", vbExclamation + vbDefaultButton1
'    End If
    
    '*******
    ' PROCEDIMIENTOS ADMINISTRATIVOS
    'gc_proceso_nivel1
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
    'rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
   ' dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    'gc_proceso_nivel2
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    'gc_proceso_nivel3
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from gc_proceso_nivel3 WHERE (proceso_codigo = 'FIN') order by etapa_descripcion", db, adOpenStatic
    'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText

    'jqa REVISAR wwwwwwwwWWWWWWWWWWWWWWWWWWWW

'    Set DtCHcodbenef.DataSource = Me.Ado_datos4.Recordset
'    DtCHcodbenef.ListField = "beneficiario_codigo"
'    dtc_codigo4.DataField = "beneficiario_codigo"
'    DtCHcodbenef.BoundColumn = "beneficiario_codigo"
'    Set DtCHcodbenef.RowSource = Me.Ado_datos4.Recordset
'
'
'    Set DtCHDescripbenef.DataSource = Me.Ado_datos4.Recordset
'    DtCHDescripbenef.ListField = "beneficiario_denominacion"
'    DtCHDescripbenef.BoundColumn = "beneficiario_codigo"
'    DtCHDescripbenef.DataField = "beneficiario_codigo"
'    Set DtCHDescripbenef.RowSource = Me.Ado_datos4.Recordset

'    '--------recordset para las cajas
'    Set rscaja = New ADODB.Recordset
'    With rscaja
'      If .State = 1 Then .Close
'      .CursorLocation = adUseClient
'     ' sqlc = "SELECT codigo_caja, denominacion_caja " & _
'     '         "From cc_cajas order by denominacion_caja"
'     sqlc = "SELECT codigo as codigo_caja , denominacion as denominacion_caja From fc_unidad_educativa"
'
'      .Open sqlc, db, adOpenKeyset, adLockReadOnly
'      Set AdoCaja.Recordset = rscaja
''======
'      If Not rscaja.BOF Then 'g-
'        .MoveFirst
'        DTCHidcaja.Text = !codigo_caja
'        DtCHIdCaja_Click 0
'        dtcDIdCaja.Text = !codigo_caja
'        DtCDIDCaja_Click 0
'      End If 'g-
''=======
'
''      DTCHidcaja.Text = !codigo_caja
''      DtCHIdCaja_Click 0
''      dtcDIdCaja.Text = !codigo_caja
'    End With
    

End Sub

Private Sub ABRIR_AUX_TABLA()
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from cc_tipo_auxiliar where aux = '" & VAR_AUX1 & "' order by aux ", db, adOpenStatic
    If rs_datos7.RecordCount > 0 Then
        VAR_TABLA = rs_datos7!NombreTabla
        VAR_CODIGO = rs_datos7!nombre_codigo
        VAR_DES = rs_datos7!nombre_descripcion
    Else
        VAR_TABLA = "NN"
        VAR_CODIGO = "NN"
        VAR_DES = "NN"
    End If
'    Set Ado_datos5.Recordset = rs_datos5
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub ABRIR_AUX1()
'    Set rs_datos11 = New ADODB.Recordset
'    If rs_datos11.State = 1 Then rs_datos11.Close
'    rs_datos11.Open "Select * from cc_tipo_auxiliar where aux = '" & VAR_AUX1 & "' order by aux ", db, adOpenStatic
'    If rs_datos11.RecordCount > 0 Then
'        VAR_TABLA = rs_datos11!NombreTabla
'        VAR_CODIGO = rs_datos11!nombre_codigo
'        VAR_DES = rs_datos11!nombre_descripcion
'    Else
'        VAR_TABLA = "NN"
'        VAR_CODIGO = "NN"
'        VAR_DES = "NN"
'    End If
''    Set Ado_datos5.Recordset = rs_datos5
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub ABRIR_AUX2()
'    Set rs_datos12 = New ADODB.Recordset
'    If rs_datos12.State = 1 Then rs_datos12.Close
'
'    rs_datos12.Open "Select * from cc_tipo_auxiliar where aux = '" & VAR_AUX2 & "' order by aux ", db, adOpenStatic
'    If rs_datos12.RecordCount > 0 Then
'        VAR_TABLA = rs_datos12!NombreTabla
'        VAR_CODIGO = rs_datos12!nombre_codigo
'        VAR_DES = rs_datos12!nombre_descripcion
'    Else
'        VAR_TABLA = "NN"
'        VAR_CODIGO = "NN"
'        VAR_DES = "NN"
'    End If
End Sub

Private Sub ABRIR_AUX3()
'    Set rs_datos13 = New ADODB.Recordset
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    rs_datos13.Open "Select * from cc_tipo_auxiliar where aux = '" & VAR_AUX3 & "' order by aux ", db, adOpenStatic
'    If rs_datos13.RecordCount > 0 Then
'        VAR_TABLA = rs_datos13!NombreTabla
'        VAR_CODIGO = rs_datos13!nombre_codigo
'        VAR_DES = rs_datos13!nombre_descripcion
'    Else
'        VAR_TABLA = "NN"
'        VAR_CODIGO = "NN"
'        VAR_DES = "NN"
'    End If
''    Set Ado_datos7.Recordset = rs_datos7
''    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set ClBuscaGrid = Nothing
End Sub

Private Sub lblDTC_Change()
 If Val(lblDTC.Text) <= 0 Then
    MsgBox "El tipo de cambio debe ser mayor a Cero", vbExclamation + vbDefaultButton1, "TIPO DE CAMBIO"
    Exit Sub
  End If
  If Trim(CboTipo.Text) = "PCO" Then
    If optbolivianos.Value = True Then
      TxtDSus = Round(Val(TxtDBs) / Val(lblDTC.Text), 2)
      txtHsus = TxtDSus
    End If
    If optdolares.Value = True Then
      TxtDBs = Round(Val(TxtDSus) * Val(lblDTC.Text), 2)
      txtHBs = TxtDBs
    End If
  End If
'Me.lblHTC = Trim(lblDTC.Text)
End Sub

'Private Sub lblDTC_Click()
'  If Val(lblDTC.Text) = 0 Then
'    MsgBox "El tipo de cambio debe ser mayor a Cero", vbExclamation + vbDefaultButton1, "TIPO DE CAMBIO"
'    Exit Sub
'  End If
'  If Trim(CboTipo.Text) = "PCO" Then
'    If optbolivianos.Value = True Then
'      TxtDSus = Round(Val(TxtDBs) / Val(lblDTC.Text), 2)
'      txtHsus = TxtDSus
'    End If
'    If optdolares.Value = True Then
'      TxtDBs = Round(Val(TxtDSus) * Val(lblDTC.Text), 2)
'      txtHBs = TxtDBs
'    End If
'  End If
'End Sub
'
'Private Sub mnuanulacion_Click()
'    buscacomprobante rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion, "ANL"
'    If existecomp <> 0 Then
'      MsgBox "El comprobante de anulación ya existe", vbExclamation + vbDefaultButton1
'      Exit Sub
'    Else
'      buscacomprobante rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion, "DVL"
'        If existecomp <> 0 Then
'          MsgBox "Existe un comprobante de devolución", vbExclamation + vbDefaultButton1
'          Exit Sub
'        End If
'    End If
'    Dim Opt1 As Integer
'    Opt1 = MsgBox("Está seguro de anular este comprobante??", vbQuestion + vbYesNo, "ANULACION")
'    If Opt1 = vbYes Then
'      Anulacion999 rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion
''g-
'      queryinicial = "Select cod_comp,tipo_comp,cod_trans,org_codigo,beneficiario_codigo,Num_Respaldo,status " & _
'                    ",codigo_documento,codigo_unidad, codigo_solicitud " & _
'                   "from CO_comprobante_M where status='N'"
'      If rs_datos.State = 1 Then rs_datos.Close
'      rs_datos.Open queryinicial, db, adOpenKeyset, adLockReadOnly
''g-
'      OptSinAprobar.Value = True
'      rs_datos.Requery
'
'      Set dg_datos.DataSource = rs_datos
'      If regANL999 = "1" Then
'        MsgBox "Anulación con éxito...Comprobante: " & Str(numANL999) & " ANL", vbInformation + vbDefaultButton1, "ANULACION"
'        If Not (rs_datos.EOF) Then rs_datos.MoveLast
'        rs_datos.Find "cod_comp=" & numANL999, , adSearchBackward
'        dg_datos_Click
'        Call DESHABILITA
'        'Call modificar
'        'Exit Sub
'      Else
'        MsgBox "Problemas en la Anulación", vbInformation + vbDefaultButton1, "ANULACION"
'        Exit Sub '****debe volver a intentar la  reversión
'      End If
'    Else
'      Exit Sub
'    End If
'End Sub
'Private Sub mnuDevolucion_Click()
'  buscacomprobante rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion, "DVL"
'    If existecomp <> 0 Then
'      MsgBox "El comprobante de devolución ya existe", vbExclamation + vbDefaultButton1
'      Exit Sub
'    Else
'      buscacomprobante rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion, "ANL"
'        If existecomp <> 0 Then
'          MsgBox "Existe un comprobante de Anulación", vbExclamation + vbDefaultButton1
'          Exit Sub
'        End If
'    End If
'  Dim Opt2 As Integer
'          Opt2 = MsgBox("Está seguro de la Devolución del comprobante  " & rs_aux1!Cod_Comp & " " & rs_aux1!org_codigo & "  ??", vbQuestion + vbYesNo, "DEVOLUCION")
'          If Opt2 = vbYes Then
'            DEVOLUCION999 rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion
'            'g-
'            queryinicial = "Select cod_comp,tipo_comp,cod_trans,org_codigo,beneficiario_codigo,Num_Respaldo,status " & _
'                          ",codigo_documento,codigo_unidad, codigo_solicitud " & _
'                         "from CO_comprobante_M where status='N'"
'            If rs_datos.State = 1 Then rs_datos.Close
'            rs_datos.Open queryinicial, db, adOpenKeyset, adLockReadOnly
'            'g-
'            OptSinAprobar.Value = True
'            rs_datos.Requery
'            Set dg_datos.DataSource = rs_datos
'            If regDEV999 = "1" Then
'              MsgBox "Devolución con éxito... Comprobante: " & Str(numDEV999) & "  DVL", vbInformation + vbDefaultButton1, "DEVOLUCION"
'              'g-
'              If Not (rs_datos.EOF) Then rs_datos.MoveLast
'              rs_datos.Find "cod_comp=" & numDEV999, , adSearchBackward 'g-
'              dg_datos_Click
'              Call DESHABILITA
'            Else
'              MsgBox "Problemas en la Devolución", vbInformation + vbDefaultButton1, "DEVOLUCION"
'              Exit Sub '****debe volver a intentar la  reversión
'            End If
'          Else
'            Exit Sub
'          End If
'End Sub
'Private Sub mnuReversion_Click()
'  Dim Opt3 As Integer
'  buscacomprobante rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion, "RVT"
'  If existecomp <> 0 Then
'     MsgBox "El comprobante de Reversión ya existe", vbExclamation + vbDefaultButton1, "REVERSION"
'     Exit Sub
'  End If
'  Opt3 = MsgBox("Está seguro de la Reversión del comprobante  " & rs_aux1!Cod_Comp & "  " & rs_aux1!org_codigo & "  ??", vbQuestion + vbYesNo, "ANULACION")
'  If Opt3 = vbYes Then
'    Reversion999 rs_aux1!Cod_Comp, rs_aux1!org_codigo, rs_aux1!ges_gestion
'  'g-
'      queryinicial = "Select cod_comp,tipo_comp,cod_trans,org_codigo,beneficiario_codigo,Num_Respaldo,status " & _
'                    ",codigo_documento,codigo_unidad, codigo_solicitud " & _
'                   "from CO_comprobante_M where status='N'"
'      If rs_datos.State = 1 Then rs_datos.Close
'      rs_datos.Open queryinicial, db, adOpenKeyset, adLockReadOnly
'  'g-
'    OptSinAprobar.Value = True
'    rs_datos.Requery
'    Set dg_datos.DataSource = rs_datos
'    If regRVT999 = "1" Then
'      MsgBox "Reversión con éxito!!. Comprobante : " & Str(numRVT999) & " RVT", vbInformation + vbDefaultButton1, "REVERSION"
'      If Not (rs_datos.EOF) Then rs_datos.MoveLast
'      rs_datos.Find "cod_comp=" & numRVT999, , adSearchBackward
'      dg_datos_Click
'      Call DESHABILITA
'    Else
'      MsgBox "Problemas en la reversión", vbInformation + vbDefaultButton1, "REVERSION"
'      Exit Sub '****debe volver a intentar la  reversión
'    End If
'  Else
'    Exit Sub
'  End If
'End Sub





Private Sub H_Correl_cmb_Click(Area As Integer)
  H_Cuenta_cmb.BoundText = H_Correl_cmb.BoundText
  H_Nombre_cmb.BoundText = H_Correl_cmb.BoundText
  H_Subcta1_cmb.BoundText = H_Correl_cmb.BoundText
  H_Subcta2_cmb.BoundText = H_Correl_cmb.BoundText
  H_Cta_Aux1_cmb.BoundText = H_Correl_cmb.BoundText
  H_Cta_Aux2_cmb.BoundText = H_Correl_cmb.BoundText
  H_Cta_Aux3_cmb.BoundText = H_Correl_cmb.BoundText
End Sub

Private Sub H_Cta_Aux1_cmb_Click(Area As Integer)
  H_Correl_cmb.BoundText = H_Cta_Aux1_cmb.BoundText
  H_Cuenta_cmb.BoundText = H_Cta_Aux1_cmb.BoundText
  H_Nombre_cmb.BoundText = H_Cta_Aux1_cmb.BoundText
  H_Subcta1_cmb.BoundText = H_Cta_Aux1_cmb.BoundText
  H_Subcta2_cmb.BoundText = H_Cta_Aux1_cmb.BoundText
  H_Cta_Aux2_cmb.BoundText = H_Cta_Aux1_cmb.BoundText
  H_Cta_Aux3_cmb.BoundText = H_Cta_Aux1_cmb.BoundText
End Sub

Private Sub H_Cta_Aux2_cmb_Click(Area As Integer)
  H_Correl_cmb.BoundText = H_Cta_Aux2_cmb.BoundText
  H_Cuenta_cmb.BoundText = H_Cta_Aux2_cmb.BoundText
  H_Nombre_cmb.BoundText = H_Cta_Aux2_cmb.BoundText
  H_Subcta1_cmb.BoundText = H_Cta_Aux2_cmb.BoundText
  H_Subcta2_cmb.BoundText = H_Cta_Aux2_cmb.BoundText
  H_Cta_Aux1_cmb.BoundText = H_Cta_Aux2_cmb.BoundText
  H_Cta_Aux3_cmb.BoundText = H_Cta_Aux2_cmb.BoundText
End Sub

Private Sub H_Cta_Aux3_cmb_Click(Area As Integer)
  H_Correl_cmb.BoundText = H_Cta_Aux3_cmb.BoundText
  H_Cuenta_cmb.BoundText = H_Cta_Aux3_cmb.BoundText
  H_Nombre_cmb.BoundText = H_Cta_Aux3_cmb.BoundText
  H_Subcta1_cmb.BoundText = H_Cta_Aux3_cmb.BoundText
  H_Subcta2_cmb.BoundText = H_Cta_Aux3_cmb.BoundText
  H_Cta_Aux1_cmb.BoundText = H_Cta_Aux3_cmb.BoundText
  H_Cta_Aux2_cmb.BoundText = H_Cta_Aux3_cmb.BoundText
End Sub

Private Sub H_Cuenta_cmb_Click(Area As Integer)
  H_Correl_cmb.BoundText = H_Cuenta_cmb.BoundText
  H_Nombre_cmb.BoundText = H_Cuenta_cmb.BoundText
  H_Subcta1_cmb.BoundText = H_Cuenta_cmb.BoundText
  H_Subcta2_cmb.BoundText = H_Cuenta_cmb.BoundText
  H_Cta_Aux1_cmb.BoundText = H_Cuenta_cmb.BoundText
  H_Cta_Aux2_cmb.BoundText = H_Cuenta_cmb.BoundText
  H_Cta_Aux3_cmb.BoundText = H_Cuenta_cmb.BoundText
End Sub

Private Sub H_Nombre_cmb_Click(Area As Integer)
  H_Correl_cmb.BoundText = H_Nombre_cmb.BoundText
  H_Cuenta_cmb.BoundText = H_Nombre_cmb.BoundText
  H_Subcta1_cmb.BoundText = H_Nombre_cmb.BoundText
  H_Subcta2_cmb.BoundText = H_Nombre_cmb.BoundText
  H_Cta_Aux1_cmb.BoundText = H_Nombre_cmb.BoundText
  H_Cta_Aux2_cmb.BoundText = H_Nombre_cmb.BoundText
  H_Cta_Aux3_cmb.BoundText = H_Nombre_cmb.BoundText
End Sub

Private Sub H_Subcta1_cmb_Click(Area As Integer)
  H_Correl_cmb.BoundText = H_Subcta1_cmb.BoundText
  H_Cuenta_cmb.BoundText = H_Subcta1_cmb.BoundText
  H_Nombre_cmb.BoundText = H_Subcta1_cmb.BoundText
  H_Subcta2_cmb.BoundText = H_Subcta1_cmb.BoundText
  H_Cta_Aux1_cmb.BoundText = H_Subcta1_cmb.BoundText
  H_Cta_Aux2_cmb.BoundText = H_Subcta1_cmb.BoundText
  H_Cta_Aux3_cmb.BoundText = H_Subcta1_cmb.BoundText
End Sub

Private Sub H_Subcta2_cmb_Click(Area As Integer)
  H_Correl_cmb.BoundText = H_Subcta2_cmb.BoundText
  H_Cuenta_cmb.BoundText = H_Subcta2_cmb.BoundText
  H_Nombre_cmb.BoundText = H_Subcta2_cmb.BoundText
  H_Subcta1_cmb.BoundText = H_Subcta2_cmb.BoundText
  H_Cta_Aux1_cmb.BoundText = H_Subcta2_cmb.BoundText
  H_Cta_Aux2_cmb.BoundText = H_Subcta2_cmb.BoundText
  H_Cta_Aux3_cmb.BoundText = H_Subcta2_cmb.BoundText
End Sub

Private Sub optbolivianos_Click()
' If adiciona = "S" Then
'    If Me.optbolivianos.Value = True Then
'       ' Me.TxtDSus.Enabled = False
'        'Me.TxtDSus.BackColor = &HE0E0E0
''        Me.TxtDBs.Enabled = True
'        'Me.TxtDBs.BackColor = &HFFFFFF
'        Ctipomoneda = CmonedaBs
'        Fram_AsientoD.Enabled = True
'        TDBFrameDebeCta.Enabled = True
'        TDBFrameDebe.Enabled = True
'        TDBFrameHaber.Enabled = True
'        TDBFrameHaberCta.Enabled = True
'        Fram_AsientoH.Enabled = True
'        cmoney = "Bs"
'
'    End If
' End If
' If cmodificar = "M" Then
'   Ctipomoneda = CmonedaBs
''   Me.TxtDBs.Enabled = True
' End If
'    lblDMonSus.Visible = True
'    lblHMONSUS.Visible = True
''    Me.txtHsus.Visible = True
''    Me.TxtDSus.Visible = True
'    Label_MontoBs.Visible = True
''    LblHMonBs.Visible = True
'    TxtDBs.Visible = True
'    txtHBs.Visible = True
''    Me.TxtDSus.Enabled = False
''    Me.TxtDBs.Enabled = True
'    Ctipomoneda = CmonedaBs
' Select Case CboTipo
' Case "ANL", "DVL", "RVT"
''    Me.TxtDSus.Enabled = False
''    Me.TxtDBs.Enabled = True
' Case "CAM"
'    lblDMonSus.Visible = False
'    lblHMONSUS.Visible = False
''    Me.txtHsus.Visible = False
''    Me.TxtDSus.Visible = False
'    Label_MontoBs.Visible = True
'    LblHMonBs.Visible = True
'    TxtDBs.Visible = True
'    txtHBs.Visible = True
''    Me.TxtDSus.Enabled = False
''    Me.TxtDBs.Enabled = True
' End Select
End Sub

'Private Sub optCAMNo_Click()
'  Dim rsfechacam As ADODB.Recordset
'  Set rsfechacam = New ADODB.Recordset
'  If rsfechacam.State = 1 Then rsfechacam.Close
'  rsfechacam.CursorLocation = adUseClient
'  aa = Month(Date) - 1
'  rsfechacam.Open "SELECT fecha  From CC_CorrelCAM " & _
'          "WHERE (mes ='" & aa & "' AND ges_gestion ='" & Year(Date) & "')", db, adOpenKeyset, adLockReadOnly
'  If rsfechacam.RecordCount <> 0 Then
'    Me.DTPCAM.Value = rsfechacam!Fecha
'    Me.DTPCAM.Value = CFecha
'    CAMcorrel = "CAM" 'trabajar con correlativos del mes para CAM
'    Me.DTPCAM.Enabled = False
'    frameCAM.Visible = False
'  Else
'    MsgBox "Todavía no puede registrar comprobantes CAM en este mes ", vbInformation + vbDefaultButton1
'    Exit Sub
'  End If
'
'End Sub
'
'Private Sub optCAMSi_Click()
'  Me.DTPCAM.Enabled = True
'  Me.DTPCAM.Value = CFecha
'  frameCAM.Visible = False
'  CAMcorrel = "NOR" 'normal
'End Sub
'
'Private Sub optconjunto_Click()
'    Me.cboaprob_inicio.Enabled = True
'    Me.lblcomprob.Visible = True
'    Me.cbo_aprob_final.Visible = True
'    sw1 = 0
'End Sub

Private Sub optdolares_Click()
' If adiciona = "S" Then
'    If Me.optdolares.Value = True Then
''        Me.TxtDBs.Enabled = False
'        'Me.TxtDBs.BackColor = &HE0E0E0
''        Me.TxtDSus.Enabled = True
'        'Me.TxtDSus.BackColor = &HFFFFFF
'        Ctipomoneda = CmonedaSus
'        TDBFrameDebeCta.Enabled = True
'        TDBFrameDebe.Enabled = True
'        TDBFrameHaber.Enabled = True
'        TDBFrameHaberCta.Enabled = True
'      '  Fram_AsientoD.Enabled = True g--
'      '  Fram_AsientoH.Enabled = True g--
'        cmoney = "Sus"
'    End If
' End If
'  If cmodificar = "M" Then
'      Ctipomoneda = CmonedaSus
''          Me.TxtDSus.Enabled = True
'
'  End If
'  lblDMonSus.Visible = True
'    lblHMONSUS.Visible = True
''    Me.txtHsus.Visible = True
''    Me.TxtDSus.Visible = True
'    Label_MontoBs.Visible = True
'    LblHMonBs.Visible = True
'    TxtDBs.Visible = True
'    txtHBs.Visible = True
''    Me.TxtDBs.Enabled = False
''    Me.TxtDSus.Enabled = True
'    Select Case CboTipo
'      Case "CAM"
'        Label_MontoBs.Visible = False
'        LblHMonBs.Visible = False
'        TxtDBs.Visible = False
'        txtHBs.Visible = False
'        'Me.TxtDBs = "0.0"
'        'Me.txtHBs = "0.0"
'        lblDMonSus.Visible = True
'        lblHMONSUS.Visible = True
''        Me.txtHsus.Visible = True
''        Me.TxtDSus.Visible = True
''        Me.TxtDBs.Enabled = False
''        Me.TxtDSus.Enabled = True
'    End Select
End Sub



'Private Sub OptIndividual_Click()
'    Me.cboaprob_inicio.Enabled = True
'    Me.lblcomprob.Visible = False
'    Me.cbo_aprob_final.Visible = False
'    sw1 = 1
'End Sub

Private Sub OptSinAprobar_Click()
'    If rs_datos.State = 1 Then rs_datos.Close
'        rs_datos.Filter = adFilterNone
'        queryinicial = "Select * from CO_comprobante_M where estado_codigo='REG'"
'        rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'        'rs_datos.Sort = "Cod_Comp ASC"
'    Set Me.dg_datos.DataSource = rs_datos
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    If rs_datos.RecordCount <> 0 Then
'    rs_datos.MoveFirst
'    dg_datos_Click
'    'Me.dg_datos_Click
'    End If
    
    '===== Proceso para filtrado general de datos(registros aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From CO_comprobante_M WHERE estado_codigo = 'REG' "
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "doc_codigo, doc_numero"        'Fecha_transacion,
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub opttodos_Click()
'If rs_datos.State = 1 Then rs_datos.Close
'rs_datos.CursorLocation = adUseClient
'    rs_datos.Filter = adFilterNone
'    queryinicial = "Select * from CO_comprobante_M "
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    If rs_datos.RecordCount <> 0 Then
'      'rs_datos.Sort = "cod_comp ASC"
'      Set Me.dg_datos.DataSource = rs_datos
'      Set Ado_datos.Recordset = rs_datos.DataSource
'      rs_datos.MoveFirst
'      dg_datos_Click
'    End If
    
    '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From CO_comprobante_M "
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "doc_codigo,  doc_numero"       'Fecha_transacion,
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Public Sub limpiar()
    'On Error Resume Next
   ' Me.txt_fecha = Empty
    Me.txt_ges = Empty
    Me.txt_campo1 = Empty
    Me.txtcodsolicitud = Empty
'    CboDCta.Text = Empty
'    CboHcta.Text = Empty
    'Me.CboDCta.ListIndex = -1
    'Me.CboDSubcta1.ListIndex = -1
   ' Me.CboDSubcta2.ListIndex = -1
  '  Me.CboHcta.ListIndex = -1
   ' Me.CbohSubcta1.ListIndex = -1
   ' Me.CbohSubcta2.ListIndex = -1
'    Me.frameDaux00.Visible = True
'    Me.frameHAux00.Visible = True
   ' Me.dtc_codigo4 = -1
    Me.txt_codigo1 = Empty
    Me.Txt_glosa = ""
    Me.txt_campo1 = "0"
    Me.TxtComprobante = ""
'    Me.TxtDBs = ""
'    Me.TxtDSus = ""
'    Me.txtHBs = ""
'    Me.txtHsus = ""
'    Me.lblHBenefaux1 = ""
'    Me.lblHnomBenefaux1 = ""
'    Me.lblDBenefaux1 = ""
'    Me.lblDnomBenefaux1 = ""
End Sub

Public Sub genera_codigo()
    'With dtetraspasos
    'Correlativo Unico
    Set rscorrelativo = New ADODB.Recordset
    rscorrelativo.CursorLocation = adUseClient
    'If rscorrelativo.State = 1 Then rscorrelativo.Close
      rscorrelativo.Open "SELECT numero_correlativo, tipo_tramite FROM fc_correl WHERE (tipo_tramite = 'cmbte')", db, adOpenKeyset, adLockOptimistic
      If rscorrelativo.RecordCount <> 0 Then
        rscorrelativo.MoveFirst
        num_comprobante = rscorrelativo!numero_correlativo + 1
        rscorrelativo!numero_correlativo = rscorrelativo!numero_correlativo + 1
        rscorrelativo.Update
      Else
        num_comprobante = 1
        rscorrelativo!numero_correlativo = 1
        rscorrelativo.Update
      End If
    
    'Correlativo por Mes y Tipo de Comprobante
    Set rs_aux2 = New ADODB.Recordset
    'rscorrelativo.Open "SELECT numero_correlativo, tipo_tramite FROM fc_correl WHERE (tipo_tramite = 'cmbte')", db, adOpenKeyset, adLockOptimistic
    SQL_FOR = "select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & Trim(txt_mes.Text) & "' and cta_codigo2 = '" & txt_codigo1.Text & "' ) "              'R-128' "
    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
    If rs_aux2.RecordCount > 0 Then
          rs_aux2!numero_correlativo = rs_aux2!numero_correlativo + 1
          txt_campo1.Caption = rs_aux2!numero_correlativo
          rs_aux2.Update
    End If
    'End With
End Sub
'
'Private Sub rs_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'   'dg_datos_Click_
'End Sub

'Private Sub SSTabHaber_Click(PreviousTab As Integer)
'Select Case SSTabHaber.Tab
'    Case 0
'      habertab haux1
'    Case 1
'      habertab haux2
'    Case 2
'      habertab haux3
'  End Select
'End Sub

'Private Sub SSTabDebe_Click(PreviousTab As Integer)
'  Select Case SSTabDebe.Tab
'    Case 0
'      debetab daux1
'    Case 1
'      debetab daux2
'    Case 2
'      debetab daux3
'  End Select
'End Sub

'Private Sub Txt_glosa_LostFocus()
'Txt_glosa.Text = UCase(Txt_glosa)
''Me.frame_moneda.Enabled = True
'Me.optbolivianos.Value = True
'End Sub
'
'Private Sub TxtDBs_Change()
'On Error GoTo err1
''If Me.optdolares = False Then
'If optbolivianos.Value = True Then
'    If lblDTC = "" Then
'        Exit Sub
'    Else
'        If cmoney = "Sus" Then
'            Exit Sub
'        Else
'          If Me.CboTipo.Text <> "CAM" Then
'            Me.TxtDSus = Round(Val(IIf(IsNull(Me.TxtDBs.Text), 0, Me.TxtDBs)) / Val(IIf(IsNull(Me.lblDTC), 1, lblDTC)), 2)
'            Me.txtHsus = Me.TxtDSus
'            Me.txtHBs = Me.TxtDBs
'          Else
'            Me.txtHBs = Me.TxtDBs
'          End If
'        End If
'    End If
'End If
'err1:
'If Err.Number = 11 Then
'  MsgBox "Introduzca el tipo de cambio", vbExclamation + vbDefaultButton1, "TIPO DE  CAMBIO"
'  Exit Sub
'End If
'End Sub
'
'Private Sub TxtDBs_GotFocus()
' TxtDBs.SelStart = 0
' TxtDBs.SelLength = Len(TxtDBs.Text)
'End Sub
'
'Private Sub TxtDBs_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'        KeyAscii = 0        'Para que no "pite"
'        SendKeys "{tab}"    'Envia una pulsación TAB
'    ElseIf KeyAscii <> 8 Then   'El 8 es la tecla de borrar (backspace)
'    'Si después de añadirle la tecla actual no es un número...
'        If Not IsNumeric("0" & TxtDBs.Text & Chr(KeyAscii)) Then
'        '... se desecha esa tecla y se avisa de que no es correcta
'            Beep
'            KeyAscii = 0
'        End If
'    End If
'End Sub
'Private Sub TxtDBs_LostFocus()
'Select Case CboTipo
' Case "ANL", "DVL", "RVT"
'  verificamonto rs_aux1!cod_trans, rs_aux1!org_codigo, rs_aux1!ges_gestion
'  If Round(Val(TxtDBs), 2) > Round(MontoAnterior, 2) Then
'    MsgBox "El monto no debe exceder a :  " & MontoAnterior, vbExclamation + vbDefaultButton1, "MONTOS DIFERENTES"
'    Me.TxtDBs.SetFocus
'    Exit Sub
'  End If
'End Select
'End Sub
'
'Private Sub TxtDSus_Change()
'If Me.lblDTC = 0 And CboTipo <> "CAM" Then
'  MsgBox "Introduzca el tipo de cambio", vbExclamation + vbDefaultButton1, "TIPO DE  CAMBIO"
'  Exit Sub
'End If
'  If Me.optdolares.Value = True And CboTipo <> "CAM" Then
'    If cmoney = "Bs" Then
'        Exit Sub
'    Else
'        Me.TxtDBs = Round(Val(Me.TxtDSus.Text) * Val(Me.lblDTC), 2)
'        Me.txtHBs = Me.TxtDBs
'        Me.txtHsus = Me.TxtDSus
'    End If
'  End If
'
'If CboTipo = "CAM" Then
'txtHsus.Text = TxtDSus.Text
'End If
'End Sub
'
'Private Sub TxtDSus_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'        KeyAscii = 0        'Para que no "pite"
'        SendKeys "{tab}"    'Envia una pulsación TAB
'    ElseIf KeyAscii <> 8 Then   'El 8 es la tecla de borrar (backspace)
'    'Si después de añadirle la tecla actual no es un número...
'        If Not IsNumeric("0" & TxtDSus.Text & Chr(KeyAscii)) Then
'        '... se desecha esa tecla y se avisa de que no es correcta
'            Beep
'            KeyAscii = 0
'        End If
'    End If
'End Sub

Private Sub Titulo(cuenta As String, subcta1 As String, subcta2 As String)
    Dim rstitulo As ADODB.Recordset
    Set rstitulo = New ADODB.Recordset
    rstitulo.CursorLocation = adUseClient
    If rstitulo.State = 1 Then rstitulo.Close
    rstitulo.Open "SELECT Mov From CC_Plan_Cuentas WHERE Cuenta = '" & cuenta & "' AND SubCta1 = '" & _
     subcta1 & "' AND SubCta2 = '" & subcta2 & "'", db, adOpenForwardOnly, adLockReadOnly
    'rstitulo.Open "select Mov from cc_plan_cuentas where cuenta='" & cuenta & "' and subcta1=' " & _
     '           subcta1 & "' and subcta2='" & subcta2 & "'", db, adOpenForwardOnly, adLockReadOnly
    If rstitulo.RecordCount = 0 Then
        MsgBox "La cuenta no existe,seleccione otra cuenta", vbExclamation + vbDefaultButton1, "Error en el Manejo de Cuentas"
'        lcta = "N"
    Else
'        lcta = "S"
        Select Case rstitulo!mov
        Case "T"
            MsgBox "La cuenta es de Titulo, seleccione otra cuenta", vbExclamation + vbOKOnly, "Error en el manejo de Cuentas"
            MovCuenta = "T"
        Case "S"
            MsgBox "La cuenta es de Sub Titulo, seleccione otra cuenta", vbExclamation + vbOKOnly, "Error en el manejo de Cuentas"
            MovCuenta = "S"
        Case "D"
            MovCuenta = "D"
    End Select
    End If
End Sub

Private Sub buscabenef(Codigo As String)
'    Dim rsBusca As ADODB.Recordset
'    Set rsBusca = New ADODB.Recordset
'    rsBusca.CursorLocation = adUseClient
'    rsBusca.Open "select beneficiario_denominacion from gc_beneficiario where beneficiario_codigo='" & _
'            Codigo & "'", db, adOpenForwardOnly, adLockReadOnly
'
'    If rsBusca.RecordCount <> 0 Then
'        Cdenominacion = rsBusca!beneficiario_denominacion
'    Else
'        MsgBox "El beneficiario no está registrado", vbExclamation + vbDefaultButton1
'        Cdenominacion = ""
'    End If
End Sub

Private Sub buscactabancaria(ctabancaria As String)
'    Dim rsctabanco As ADODB.Recordset
'    Set rsctabanco = New ADODB.Recordset
'    rsctabanco.CursorLocation = adUseClient
'    rsctabanco.Open "select cta_descripcion from fc_cuenta_bancaria where cta_codigo='" & Trim(ctabancaria) & "'", db, adOpenForwardOnly, adLockReadOnly
'    If rsctabanco.RecordCount <> 0 Then
'        cdenomctabancaria = rsctabanco!cta_descripcion
'    Else
'        MsgBox "La cuenta corriente no existe", vbExclamation + vbDefaultButton1
'        cdenomctabancaria = ""
'    End If
End Sub
'

'Private Sub PCO(Cta As String, Movim As String, Cod_Comp As Integer)
'    Dim rsctapco As ADODB.Recordset
'    Dim rsAuxM As ADODB.Recordset
'    Dim rsAuxdiario As ADODB.Recordset
'    Set rsAuxM = New ADODB.Recordset
'    Set rsAuxdiario = New ADODB.Recordset
'    Set rsctapco = New ADODB.Recordset
'    If rspco.State = 1 Then rspco.Close
'    rspco.Open " Select * from Co_MovimientoPCo where cod_comp=" & Trim(Cod_Comp) & " and  tipo_comp='PCO' and cta_codigo='" & Trim(Cta) & "'", db, adOpenKeyset, adLockOptimistic
'        If rspco.RecordCount <> 0 Then
'           MsgBox "El comprobante ya existe", vbExclamation + vbDefaultButton1
'        Exit Sub
'        '*******modificar el comprobante ya existente
'        Else
'            If rsAuxM.State = 1 Then rsAuxM.Close
'            If rsAuxdiario.State = 1 Then rsAuxdiario.Close
'            rsAuxM.CursorLocation = adUseClient
'            rsAuxdiario.CursorLocation = adUseClient
'            rsAuxM.Open "select * from Co_Comprobante_M  where cod_comp=" & Val(Cod_Comp) & " and tipo_comp='PCO'", db, adOpenKeyset, adLockReadOnly
'            rsAuxdiario.Open "select * from Co_Diario where cod_comp=" & Val(Cod_Comp) & " and tipo_comp='PCO'", db, adOpenKeyset, adLockReadOnly
'            rspco.AddNew
'            rspco!ges_gestion = rsAuxM!ges_gestion
'            rspco!org_codigo = "999"
'            rspco!Cod_Comp = rsAuxM!Cod_Comp
'            rspco!tipo_comp = Trim(rsAuxM!tipo_comp)
'            rspco!codigo_pago_detalle = Trim(rsAuxM!cod_trans_detalle)
'            rspco!beneficiario_codigo = Trim(rsAuxM!beneficiario_codigo)
'            rspco!Concepto = Trim(rsAuxM!glosa)
'            If Movim = "D" Then
'              rspco!Cta_Codigo = rsAuxdiario!D_Cta_Aux1
'              rspco!DebeBs = rsAuxdiario!d_montoBs
'              rspco!DebeDl = rsAuxdiario!d_montoDl
'              rspco!HaberBs = 0
'              rspco!HaberDl = 0
'              If rsctapco.State = 1 Then rsctapco.Close
'              rsctapco.CursorLocation = adUseClient
'              rsctapco.Open "SELECT Cta_codigo, Cta_Pco_Debe, Cta_Pco_Haber From fc_cuenta_bancaria " & _
'                       " where cta_codigo='" & Trim(rsAuxdiario!D_Cta_Aux1) & "'", db, adOpenKeyset, adLockOptimistic
'              If rsctapco.RecordCount <> 0 Then
'                rsctapco!Cta_Pco_Debe = rsctapco!Cta_Pco_Debe + rsAuxdiario!d_montoBs
'                rsctapco.Update
'              End If
'            End If
'            If Movim = "H" Then
'                rspco!Cta_Codigo = rsAuxdiario!H_Cta_Aux1
'                rspco!DebeBs = 0
'                rspco!DebeDl = 0
'                rspco!HaberBs = rsAuxdiario!h_montoBs
'                rspco!HaberDl = rsAuxdiario!h_montoDl
'                If rsctapco.State = 1 Then rsctapco.Close
'                rsctapco.CursorLocation = adUseClient
'                rsctapco.Open "SELECT Cta_codigo, Cta_Pco_Debe, Cta_Pco_Haber From fc_cuenta_bancaria " & _
'                       " where cta_codigo='" & Trim(rsAuxdiario!H_Cta_Aux1) & "'", db, adOpenKeyset, adLockOptimistic
'                If rsctapco.RecordCount <> 0 Then
'                    rsctapco!Cta_Pco_Haber = rsctapco!Cta_Pco_Debe + rsAuxdiario!h_montoBs
'                    rsctapco.Update
'                End If
'            End If
'            rspco!tipo_cambio = rsAuxdiario!d_Cambio
'            rspco!fecha_aprobacion = CDate(Format(Date, "dd/mm/yyyy"))
'            rspco!num_respaldo = rsAuxM!num_respaldo
'            rspco!usr_usuario = GlUsuario
'            rspco!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'            rspco!hora_registro = Format(Time, "hh:mm:ss")
'            rspco!tipo_moneda = cmoney
'            rspco!Status = "S"
'            rspco.Update
'        End If
'End Sub

Private Sub buscaorganismo(orgo As String)
'  Dim rsbuscaorg As ADODB.Recordset
'  Set rsbuscaorg = New ADODB.Recordset
'  If rsbuscaorg.State = 1 Then rsbuscaorg.Close
'  rsbuscaorg.CursorLocation = adUseClient
'  rsbuscaorg.Open "SELECT Org_codigo, Org_descripcion From fc_organismo_financiamiento " & _
'                  "WHERE (Org_codigo = '" & orgo & "')", db, adOpenKeyset, adLockReadOnly
'  If rsbuscaorg.RecordCount <> 0 Then
'    denomorgan = rsbuscaorg!org_descripcion
'  Else
'    denomorgan = ""
'  End If
End Sub

'Public Sub genera_CorrelCAM(Fecha As Date)
'  Dim rscorrCAM As ADODB.Recordset
'  Dim año As String
'  Dim mes As String
'  mes = Month(Fecha)
'  año = Year(Fecha)
'  Set rscorrCAM = New ADODB.Recordset
'  If rscorrCAM.State = 1 Then rscorrCAM.Close
'  rscorrCAM.Open "select * from CC_correlCAM where mes='" & mes & "' and  ges_gestion='" & año & "'", db, adOpenKeyset, adLockOptimistic
'  If rscorrCAM.RecordCount <> 0 Then
'    If Val(rscorrCAM!correl_actual) >= Val(rscorrCAM!correl_superior) Then
'      MsgBox "No existen más correlativos para este mes,se utilizará un correlativo actual", vbInformation + vbDefaultButton1
'      Call genera_codigo
'      numcomprobante = num_comprobante
'    Else
'      num_comprobante = rscorrCAM!correl_actual + 1
'      rscorrCAM!correl_actual = rscorrCAM!correl_actual + 1
'      rscorrCAM.Update
'    End If
'  End If
'End Sub
'Public Sub Status(Codigo As Integer, org As String, Gestion As String)
'  Dim Rsstatus As ADODB.Recordset
'  Set Rsstatus = New ADODB.Recordset
'  Rsstatus.Open "select estado_pagado,estado_contabilidad from pagos where codigo_pago=" & _
'                Codigo & " and org_codigo='" & org & "' and ges_gestion='" & Gestion & "'", db, adOpenKeyset, adLockReadOnly
'  If Rsstatus.RecordCount <> 0 Then
'    estadoconta = Rsstatus!estado_contabilidad
'    estadopago = Rsstatus!estado_pagado
'  End If
'End Sub
'Private Sub modificar()
'      Me.FraGrabarCancelar.Visible = True
'      Me.fraOpciones.Visible = False
'      'Me.fraOpciones.Visible = False
'      'Me.Fram_AsientoD.Enabled = True
'      TDBFrameDebeCta.Enabled = True
'      TDBFrameDebe.Enabled = True
'      TDBFrameHaber.Enabled = True
'      TDBFrameHaberCta.Enabled = True
'      'Me.Fram_AsientoH.Enabled = True
'      Me.FraGlobal.Enabled = True
'      Me.FraNavega.Enabled = False
'      Me.frame_moneda.Visible = True
'      Me.frame_moneda.Enabled = True
'      cmodificar = "M"
'End Sub
'Private Sub DESHABILITA()
'  Me.CboTipo.Enabled = False
'  Me.frameDaux00.Enabled = False
'  Me.FrameDBeneficiario.Enabled = False
'  Me.frameDCtaBancaria.Enabled = False
'  Me.frameDOrganismos.Enabled = False
'  '---
'  Me.frameHAux00.Enabled = False
'  Me.FrameHBeneficiario.Enabled = False
'  Me.frameHCtaBancaria.Enabled = False
'  Me.frameHOrganismos.Enabled = False
'  Me.dtc_codigo4.Enabled = False
'  Me.dtc_desc4.Enabled = False
'  Me.txt_codigo1.Enabled = False
'  Me.dtcbodocumento2.Enabled = False
'  Me.txt_campo1.Enabled = False
'  Me.txtcodsolicitud.Enabled = False
'  Me.frame_moneda.Enabled = False
'  Me.optbolivianos.Value = True
'  optbolivianos_Click
'  '---
'  Me.CboDCta.Enabled = False
'  Me.CboDSubcta1.Enabled = False
'  Me.CboDSubcta2.Enabled = False
'  Me.CboHcta.Enabled = False
'  Me.CbohSubcta1.Enabled = False
'  Me.CbohSubcta2.Enabled = False
'  cmodificar = "M"
'  '---
'   Me.FraGrabarCancelar.Visible = True
'   Me.fraOpciones.Visible = False
'   Me.fraOpciones.Visible = False
'   'Me.Fram_AsientoD.Enabled = True
'   'Me.Fram_AsientoH.Enabled = True
'   TDBFrameDebeCta.Enabled = True
'    TDBFrameDebe.Enabled = True
'    TDBFrameHaber.Enabled = True
'    TDBFrameHaberCta.Enabled = True
'   Me.FraGlobal.Enabled = True
'   Me.FraNavega.Enabled = False
'   'Me.frame_moneda.Visible = True
'   'Me.frame_moneda.Enabled = True
'End Sub
'Private Sub Habilita()
'  Me.CboTipo.Enabled = True
'  Me.frameDaux00.Enabled = True
'  Me.FrameDBeneficiario.Enabled = True
'  Me.frameDCtaBancaria.Enabled = True
'  Me.frameDOrganismos.Enabled = True
'  '---
'  Me.frame_moneda.Enabled = True
'  Me.frameHAux00.Enabled = True
'  Me.FrameHBeneficiario.Enabled = True
'  Me.frameHCtaBancaria.Enabled = True
'  Me.frameHOrganismos.Enabled = True
'  Me.dtc_codigo4.Enabled = True
'  Me.dtc_desc4.Enabled = True
'  Me.txt_codigo1.Enabled = True
'  Me.dtcbodocumento2.Enabled = True
'  Me.txt_campo1.Enabled = True
'  Me.txtcodsolicitud.Enabled = True
'  Me.frame_moneda.Enabled = True
'  Me.CboDCta.Enabled = True
'  Me.CboDSubcta1.Enabled = True
'  Me.CboDSubcta2.Enabled = True
'  Me.CboHcta.Enabled = True
'  Me.CbohSubcta1.Enabled = True
'  Me.CbohSubcta2.Enabled = True
'
'  'Me.optbolivianos.Value = True
'  End Sub
'Private Sub verificamonto(codanterior As Integer, org As String, Gestion As String)
'Dim rsverifica As ADODB.Recordset
'Set rsverifica = New ADODB.Recordset
'If rsverifica.State = 1 Then rsverifica.Close
'rsverifica.CursorLocation = adUseClient
'rsverifica.Open "SELECT CO_Diario.D_MontoBs, CO_Diario.D_MontoDl" & _
'                " FROM Co_Comprobante_M INNER JOIN CO_Diario ON " & _
'                " Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp" & _
'                " WHERE (Co_Comprobante_M.org_codigo = '" & org & "') AND " & _
'                "(Co_Comprobante_M.ges_gestion = '" & Gestion & "') AND " & _
'                " (Co_Comprobante_M.Cod_Comp=" & codanterior & ")", db, adOpenKeyset, adLockReadOnly
'If rsverifica.RecordCount <> 0 Then
'  MontoAnterior = rsverifica!d_montoBs
'End If
'End Sub
'Private Sub ModifAsientos(glosa As String, bolivianos As Double, dolares As Double)
'  Dim sqlactualizaM As String
'  Dim sqlactualizaD As String
'  sqlactualizaM = "update co_comprobante_m set " & _
'                  "glosa ='" & Trim(glosa) & "' where  cod_comp=" & rs_datos!Cod_Comp & "  and org_codigo='" & rs_datos!org_codigo & "'"
'
'  sqlactualizaD = "update co_diario set " & _
'                 "d_montoBs=" & Round(bolivianos, 2) & "," & _
'                 "d_MontoDl=" & Round(dolares, 2) & "," & _
'                 "h_montoBs=" & Round(bolivianos, 2) & "," & _
'                 "h_MontoDl=" & Round(dolares, 2) & " where  cod_comp=" & rs_datos!Cod_Comp
'  db.Execute sqlactualizaM
'  db.Execute sqlactualizaD
'End Sub
'
'Private Sub tipocompadiciona(SW As String, tipo As String)
'    '-----
'    rstipocomp.Filter = adFilterNone
'    rstipocomp.Filter = "contabilidad='CC'"
'    'For i = 0 To CboTipo.ListCount - 1
'    '  If CboTipo.List(i - 1) <> "CAM" And CboTipo.List(i - 1) <> "PCO" And CboTipo.List(i - 1) <> "PCE" Then
'    '    CboTipo.RemoveItem (i)
'    '  End If
'    'Next
'    CboTipo.Clear
'    cboNomTipo.Clear
'        If rstipocomp.RecordCount <> 0 Then
'    Do While Not rstipocomp.EOF
'          CboTipo.AddItem Trim(rstipocomp!Codigo_tipo)
'          cboNomTipo.AddItem Trim(rstipocomp!Denominacion_Tipo)
'          rstipocomp.MoveNext
'      Loop
'    End If
'    If SW = "M" Then
'      CboTipo.Text = tipo
'      CboTipo_Click
'    End If
'    '---
'End Sub
'Private Sub tipocompllena(tipo As String)
'    '-----
'    rstipocomp.Filter = adFilterNone
'    CboTipo.Clear
'    cboNomTipo.Clear
'    If rstipocomp.RecordCount <> 0 Then
'      rstipocomp.MoveFirst
'      Do While Not rstipocomp.EOF
'          CboTipo.AddItem Trim(rstipocomp!Codigo_tipo)
'          cboNomTipo.AddItem Trim(rstipocomp!Denominacion_Tipo)
'          rstipocomp.MoveNext
'      Loop
'    End If
'    '---
'        CboTipo.Text = tipo
'      '  CboTipo_Click
'End Sub
'Public Sub auxDebe(Aux As String)
'  Dim sql1 As String
'  Select Case Aux
'      Case "09"
'        frameDaux00.Visible = False
'        frameDCtaBancaria.Visible = False
'        frameDOrganismos.Visible = False
''        Me.FrameDBeneficiario.Visible = False
'        TDBFrameDConvenio.Visible = True
'        TDBFrameDCaja.Visible = False
'      Case "10"
'        frameDaux00.Visible = False
'        frameDCtaBancaria.Visible = False
'        frameDOrganismos.Visible = False
''        Me.FrameDBeneficiario.Visible = False
'        TDBFrameDConvenio.Visible = False
'        TDBFrameDCaja.Visible = True
'      Case "00" ' no se introduce nada
'          frameDaux00.Visible = True
'          frameDCtaBancaria.Visible = False
''          Me.FrameDBeneficiario.Visible = False
'          frameDOrganismos.Visible = False
'          TDBFrameDConvenio.Visible = False
'          TDBFrameDCaja.Visible = False
'          dauxiliar = ""
'      Case "01" ' se introduce un beneficiario
'          frameDaux00.Visible = False
'          frameDCtaBancaria.Visible = False
'          frameDOrganismos.Visible = False
''          Me.FrameDBeneficiario.Visible = True
''          Me.lblDBenefaux1 = Trim(Me.DtCDcodbenef.Text)
''          Me.lblDnomBenefaux1 = Trim(Me.DtCDDescripbenef.Text)
'          TDBFrameDConvenio.Visible = False
'          TDBFrameDCaja.Visible = False
''          dauxiliar = Trim(Me.DtCDcodbenef.Text)
'      Case "02" 'se introduce una cuenta bancaria
'          auxctacorriente = cboDctaaux1
'          frameDaux00.Visible = False
'          TDBFrameDConvenio.Visible = False
''          Me.FrameDBeneficiario.Visible = False
'          frameDCtaBancaria.Visible = True
'          frameDOrganismos.Visible = False
'          TDBFrameDCaja.Visible = False
'          If (Trim(CboDCta) = "1111" And Trim(CboDSubcta1) = "02") Or (Trim(CboDCtaCAM) = "1111" And Trim(CboDSub1CAM) = "02") Then
'            If Trim(CboDCta) = "1111" Then
''              Select Case Me.CboDSubcta2
''                  Case "01"
''                      sql1 = "SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                          "where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
''                  Case "02"
''                      sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                          "where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
''                  Case "03"
''                      sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                          "where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
''              End Select
'          Else
'            If Trim(CboDCtaCAM) = "1111" Then
''              Select Case Me.CboDSub2CAM.Text
''                  Case "01"
''                      sql1 = "SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                          "where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
''                  Case "02"
''                      sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                          "where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
''                  Case "03"
''                      sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
''                          "where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
''              End Select
'            End If
'          End If
''              Me.cboDctaaux1.Clear
''              Me.cboDctanomaux1.Clear
''              Set rscta_corrienteDebe = New ADODB.Recordset
''              rscta_corrienteDebe.Filter = adFilterNone
''              If rscta_corrienteDebe.State = 1 Then rscta_corrienteDebe.Close
''              rscta_corrienteDebe.CursorLocation = adUseClient
''              rscta_corrienteDebe.Open sql1, db, adOpenForwardOnly, adLockReadOnly
''              If rscta_corrienteDebe.RecordCount <> 0 Then
''                  rscta_corrienteDebe.MoveFirst
''                  Do While Not rscta_corrienteDebe.EOF
''                      cboDctaaux1.AddItem rscta_corrienteDebe!Cta_Codigo
''                      cboDctanomaux1.AddItem rscta_corrienteDebe!cta_descripcion
''                      rscta_corrienteDebe.MoveNext
'                  Loop
'              End If
'          End If
'      Case "08"
'                    frameDaux00.Visible = False
''                    Me.FrameDBeneficiario.Visible = False
'                    frameDCtaBancaria.Visible = False
'                    frameDOrganismos.Enabled = True
'                    frameDOrganismos.Visible = True
'                    TDBFrameDConvenio.Visible = False
'                    TDBFrameDCaja.Visible = False
''                    If rsOrganismo.State = 1 Then rsOrganismo.Close
''                    rsOrganismo.CursorLocation = adUseClient
''                    rsOrganismo.Filter = adFilterNone
''                    rsOrganismo.Open "SELECT Org_codigo,(Org_descripcion) AS descripcion" & _
'                                      " From fc_organismo_financiamiento order by org_codigo", db, adOpenKeyset, adLockReadOnly
'                    cboDCodOrg.Clear
'                    cboDDenomOrg.Clear
''                    If rsOrganismo.RecordCount <> 0 Then
''                      rsOrganismo.MoveFirst
''                      Do While Not rsOrganismo.EOF
''                          cboDCodOrg.AddItem rsOrganismo!org_codigo
''                          cboDDenomOrg.AddItem rsOrganismo!descripcion
''                          rsOrganismo.MoveNext
'                      Loop
'                    End If
'     Case Else ' no se ha definido todavia
'            frameDaux00.Visible = True
'            frameDCtaBancaria.Visible = False
''            Me.FrameDBeneficiario.Visible = False
'            TDBFrameDConvenio.Visible = False
'            TDBFrameDCaja.Visible = False
'            dauxiliar = ""
'   End Select
'          'trabajar con auyxiliar 2
'End Sub

'Public Sub Auxhaber(hauxiliar As String)
'Select Case hauxiliar
'                Case "09" 'auxiliar de convenios}
'                    frameHAux00.Visible = False
'                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
''                    Me.frameHOrganismos.Visible = False
'                    TDBFrameHConvenio.Visible = True
'                    TDBFrameHCaja.Visible = False
'                Case "10" 'AUXILIAR DE CAJA  ' auxiliar municipio
'                    frameHAux00.Visible = False
'                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
''                    Me.frameHOrganismos.Visible = False
'                    'TDBFrameHConvenio.Visible = True
'                    TDBFrameHCaja.Visible = True
'                Case "00" ' no se introduce nada
'                    frameHAux00.Visible = True
'                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
''                    Me.frameHOrganismos.Visible = False
'                    TDBFrameHConvenio.Visible = False
'                    TDBFrameHCaja.Visible = False
''                    'hctalarga = ""
'                Case "01" ' se introduce un beneficiario
'                    frameHAux00.Visible = False
'                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = True
''                    Me.frameHOrganismos.Visible = False
'                    TDBFrameHConvenio.Visible = False
'                    TDBFrameHCaja.Visible = False
''                    Me.lblHBenefaux1 = Trim(Me.DtCHcodbenef.Text)
''                    Me.lblHnomBenefaux1 = Trim(Me.DtCHDescripbenef.Text)
'                    'hctalarga = Trim(Me.DtCHcodbenef.Text)
'                 Case "02" 'se introduce una cuenta bancaria
'                    frameHAux00.Visible = False
'                    frameHCtaBancaria.Visible = True
''                    Me.FrameHBeneficiario.Visible = False
''                    Me.frameHOrganismos.Visible = False
'                    TDBFrameHConvenio.Visible = False
'                    TDBFrameHCaja.Visible = False
'                    If (Trim(CboHcta) = "1111" And Trim(CbohSubcta1) = "02") Or (Trim(CboHCtaCAM) = "1111" And Trim(CboHSub1CAM) = "02") Then
'                      If CboHcta.Text = "1111" Then
''                        Select Case Me.CbohSubcta2
''                            Case "01"
'                                sql1 = "SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
'                                    "where  fc_cuenta_bancaria.Fte_codigo = '20' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
''                            Case "02"
'                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
'                                    "where  fc_cuenta_bancaria.Fte_codigo = '30' order by fc_cuenta_bancaria.Cta_codigo"
''                            Case "03"
'                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
'                                    "where  fc_cuenta_bancaria.Fte_codigo = '40' or fc_cuenta_bancaria.Fte_codigo = '50' order by fc_cuenta_bancaria.Cta_codigo"
''                        End Select
'                      End If
'                      If CboHCtaCAM.Text = "1111" Then
'                        Select Case CboHSub2CAM
'                            Case "01"
'                                sql1 = "SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
'                                    "where  fc_cuenta_bancaria.Fte_codigo = '20' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
'                            Case "02"
'                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
'                                    "where  fc_cuenta_bancaria.Fte_codigo = '30' order by fc_cuenta_bancaria.Cta_codigo"
'                            Case "03"
'                                sql1 = " SELECT Cta_codigo, cta_descripcion,org_codigo FROM fc_cuenta_bancaria " & _
'                                    "where  fc_cuenta_bancaria.Fte_codigo = '40' or fc_cuenta_bancaria.Fte_codigo = '50' order by fc_cuenta_bancaria.Cta_codigo"
'                        End Select
'                      End If
''                        Me.cboHctaaux1.Clear
''                        Me.cboHctanomaux1.Clear
''                        If rscta_corrienteHaber.State = 1 Then rscta_corrienteHaber.Close
''                        Set rscta_corrienteHaber = New ADODB.Recordset
''                        rscta_corrienteHaber.Filter = adFilterNone
''                        rscta_corrienteHaber.CursorLocation = adUseClient
'                        rscta_corrienteHaber.Open sql1, db, adOpenForwardOnly, adLockReadOnly
''                        If rscta_corrienteHaber.RecordCount <> 0 Then
''                            rscta_corrienteHaber.MoveFirst
''                            Do While Not rscta_corrienteHaber.EOF
''                                cboHctaaux1.AddItem rscta_corrienteHaber!Cta_Codigo
''                                cboHctanomaux1.AddItem rscta_corrienteHaber!cta_descripcion
''                                rscta_corrienteHaber.MoveNext
'                            Loop
'                        End If
'                    End If
'                Case "08"
'                    frameHAux00.Visible = False
'                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
'                    TDBFrameHConvenio.Visible = False
''                    Me.frameHOrganismos.Visible = True
''                    Me.frameHOrganismos.Enabled = True
'                    TDBFrameHCaja.Visible = False
'''                    If rsOrganismo.State = 1 Then rsOrganismo.Close
''                    rsOrganismo.CursorLocation = adUseClient
''                    rsOrganismo.Filter = adFilterNone
''                    rsOrganismo.Open "SELECT Org_codigo,(Org_descripcion) AS descripcion" & _
'                                      " From fc_organismo_financiamiento order by org_codigo", db, adOpenKeyset, adLockReadOnly
'                    cboHCodOrg.Clear
'                    cboHDenomOrg.Clear
''                    If rsOrganismo.RecordCount <> 0 Then
''                      rsOrganismo.MoveFirst
''                      Do While Not rsOrganismo.EOF
''                          cboHCodOrg.AddItem rsOrganismo!org_codigo
''                          cboHDenomOrg.AddItem rsOrganismo!descripcion
''                          rsOrganismo.MoveNext
'                      Loop
'                    End If
'                Case Else ' no se ha definido todavia
'                    frameHAux00.Visible = True
''                    Me.frameHOrganismos.Visible = False
'                    frameHCtaBancaria.Visible = False
''                    Me.FrameHBeneficiario.Visible = False
'                    TDBFrameHConvenio.Visible = False
'                    TDBFrameHCaja.Visible = False
'                    'hctalarga = ""
'            End Select
'End Sub

'Public Sub frameactivoDebe()
'    Select Case daux1
'    Case "00"
'      dctalarga = ""
'    Case "01"
'      Select Case CboTipo
'        Case "PCO"
'          dctalarga = Trim(DtCDcodbenef.Text)
'        Case Else
'          dctalarga = lblDBenefaux1
'      End Select
'    Case "02"
'      If cboDctaaux1.Text <> "" Then
'        dctalarga = Trim(cboDctaaux1.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "10"
'      If dtcDIdCaja.Text <> "" Then
'        dctalarga = Trim(dtcDIdCaja.Text)
'        salir = 0
'      Else
''        MsgBox "Seleccione una Caja", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'      MsgBox "Seleccione la Unidad Educativa", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "08"
'      If cboDCodOrg.Text <> "" Then
'        dctalarga = Trim(cboDCodOrg.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "09"
'      If DtCDIdConvenio.Text <> "" Then
'        dctalarga = Trim(DtCDIdConvenio.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'  End Select
'  Select Case daux2
'    Case "00"
'      dctaaux2 = ""
'    Case "01"
'        Select Case CboTipo
'        Case "PCO"
'          dctaaux2 = Trim(DtCDcodbenef.Text)
'        Case Else
'          dctaaux2 = lblDBenefaux1
'        End Select
'      'dctaaux2 = lblDBenefaux1
'    Case "02"
'      If cboDctaaux1.Text <> "" Then
'        dctaaux2 = Trim(cboDctaaux1.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "10"
'      If dtcDIdCaja.Text <> "" Then
'        dctaaux2 = Trim(dtcDIdCaja.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione la Unidad Educativa", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "08"
'      If cboDCodOrg.Text <> "" Then
'        dctaaux2 = Trim(cboDCodOrg.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "09"
'      If DtCDIdConvenio.Text <> "" Then
'        dctaaux2 = Trim(DtCDIdConvenio.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'  End Select
'  Select Case daux3
'    Case "00"
'      dctaaux3 = ""
'    Case "01"
'      Select Case CboTipo
'        Case "PCO"
'          dctaaux3 = Trim(DtCDcodbenef.Text)
'        Case Else
'          dctaaux3 = lblDBenefaux1
'        End Select
'      'dctaaux3 = lblDBenefaux1
'    Case "02"
'      If cboDctaaux1.Text <> "" Then
'        dctaaux3 = Trim(cboDctaaux1.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "10"
'      If dtcDIdCaja.Text <> "" Then
'        dctaaux3 = Trim(dtcDIdCaja.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione la Unidad Educativa", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "08"
'      If cboDCodOrg.Text <> "" Then
'        dctaaux3 = Trim(cboDCodOrg.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "09"
'      If DtCDIdConvenio.Text <> "" Then
'        dctaaux3 = Trim(DtCDIdConvenio.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'  End Select
'End Sub
'Public Sub frameactivoHaber()
'Select Case haux1
'    Case "00"
'      hctalarga = ""
'    Case "01"
'     Select Case CboTipo
'        Case "PCO"
'          hctalarga = Trim(DtCHcodbenef.Text)
'        Case Else
'          hctalarga = lblHBenefaux1
'     End Select
'      'hctalarga = lblHBenefaux1
'    Case "02"
'      If cboHctaaux1.Text <> "" Then
'        hctalarga = Trim(cboHctaaux1.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "10"
'      If DTCHidcaja.Text <> "" Then
'        hctalarga = Trim(DTCHidcaja.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione la Unidad Educativa", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "08"
'      If cboHCodOrg.Text <> "" Then
'        hctalarga = Trim(cboHCodOrg.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "09"
'      If DtCHIdConvenio.Text <> "" Then
'        hctalarga = Trim(DtCHIdConvenio.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Convenio en el Crédito", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'  End Select
'  Select Case haux2
'    Case "00"
'      hctaaux2 = ""
'    Case "01"
'      Select Case CboTipo
'        Case "PCO"
'          hctaaux2 = Trim(DtCHcodbenef.Text)
'        Case Else
'          hctaaux2 = lblHBenefaux1
'     End Select
''      hctaaux2 = lblHBenefaux1
'    Case "02"
'      If cboHctaaux1.Text <> "" Then
'        hctaaux2 = Trim(cboHctaaux1.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "10"
'      If DTCHidcaja.Text <> "" Then
'        hctaaux2 = Trim(DTCHidcaja.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione la Unidad Educativa", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "08"
'      If cboHCodOrg.Text <> "" Then
'        hctaaux2 = Trim(cboHCodOrg.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "09"
'      If DtCHIdConvenio.Text <> "" Then
'        hctaaux2 = Trim(DtCHIdConvenio.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Convenio en el Crédito", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'  End Select
'  Select Case haux3
'    Case "00"
'      hctaaux3 = ""
'    Case "01"
'      Select Case CboTipo
'        Case "PCO"
'          hctaaux3 = Trim(DtCHcodbenef.Text)
'        Case Else
'          hctaaux3 = lblHBenefaux1
'      End Select
'      'hctaaux3 = lblHBenefaux1
'    Case "02"
'      If cboHctaaux1.Text <> "" Then
'        hctaaux3 = Trim(cboHctaaux1.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "10"
'      If DTCHidcaja.Text <> "" Then
'        hctaaux3 = Trim(DTCHidcaja.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione la Unidad Educativa", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "08"
'      If cboHCodOrg.Text <> "" Then
'        hctaaux3 = Trim(cboHCodOrg.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'    Case "09"
'      If DtCHIdConvenio.Text <> "" Then
'        hctaaux3 = Trim(DtCHIdConvenio.Text)
'        salir = 0
'      Else
'        MsgBox "Seleccione un Convenio en el Crédito", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        salir = 1
'        Exit Sub
'      End If
'  End Select
'End Sub

Public Sub debetab(Aux)
  Dim sql1 As String
  Select Case Aux
      Case "00" ' no se introduce nada
          frameDaux00.Visible = True
          frameDCtaBancaria.Visible = False
'          Me.FrameDBeneficiario.Visible = False
          frameDOrganismos.Visible = False
          TDBFrameDConvenio.Visible = False
          TDBFrameDCaja.Visible = False
      Case "01" ' se introduce un beneficiario
          frameDaux00.Visible = False
          frameDCtaBancaria.Visible = False
          frameDOrganismos.Visible = False
'          Me.FrameDBeneficiario.Visible = True
          TDBFrameDConvenio.Visible = False
          TDBFrameDCaja.Visible = False
      Case "02" 'se introduce una cuenta bancaria
          auxctacorriente = cboDctaaux1
          frameDaux00.Visible = False
'          Me.FrameDBeneficiario.Visible = False
          frameDCtaBancaria.Visible = True
          frameDOrganismos.Visible = False
          TDBFrameDConvenio.Visible = False
          TDBFrameDCaja.Visible = False
      Case "10"
          frameDaux00.Visible = False
          frameDCtaBancaria.Visible = False
          frameDOrganismos.Visible = False
'          Me.FrameDBeneficiario.Visible = False
          TDBFrameDConvenio.Visible = False
          TDBFrameDCaja.Visible = True
      Case "08"
          frameDaux00.Visible = False
'          Me.FrameDBeneficiario.Visible = False
          frameDCtaBancaria.Visible = False
          TDBFrameDConvenio.Visible = False
          frameDOrganismos.Enabled = True
          frameDOrganismos.Visible = True
          TDBFrameDCaja.Visible = False
      Case "09"
          frameDaux00.Visible = False
'          Me.FrameDBeneficiario.Visible = False
          frameDCtaBancaria.Visible = False
          frameDOrganismos.Visible = False
          TDBFrameDConvenio.Visible = True
          TDBFrameDConvenio.Enabled = True
          TDBFrameDCaja.Visible = False
     Case Else ' no se ha definido todavia
          frameDaux00.Visible = True
          frameDCtaBancaria.Visible = False
'          Me.FrameDBeneficiario.Visible = False
          TDBFrameDCaja.Visible = False
   End Select
          'trabajar con auyxiliar 2
End Sub

Public Sub habertab(hauxi)
Select Case hauxi
      Case "09" 'auxiliar de convenio
          frameHAux00.Visible = False
          frameHCtaBancaria.Visible = False
'          Me.FrameHBeneficiario.Visible = False
'          Me.frameHOrganismos.Visible = False
          TDBFrameHConvenio.Visible = True
          TDBFrameHCaja.Visible = False
      Case "10" 'AUXILIAR DE CAJA
          frameHAux00.Visible = False
          frameHCtaBancaria.Visible = False
'          Me.FrameHBeneficiario.Visible = False
'          Me.frameHOrganismos.Visible = False
          TDBFrameHConvenio.Visible = False
          TDBFrameHCaja.Visible = True
      Case "00" ' no se introduce nada
          frameHAux00.Visible = True
          frameHCtaBancaria.Visible = False
'          Me.FrameHBeneficiario.Visible = False
'          Me.frameHOrganismos.Visible = False
          TDBFrameHConvenio.Visible = False
          TDBFrameHCaja.Visible = False
      Case "01" ' se introduce un beneficiario
          frameHAux00.Visible = False
          frameHCtaBancaria.Visible = False
'          Me.FrameHBeneficiario.Visible = True
'          Me.frameHOrganismos.Visible = False
          TDBFrameHConvenio.Visible = False
          TDBFrameHCaja.Visible = False
       Case "02" 'se introduce una cuenta bancaria
          frameHAux00.Visible = False
          frameHCtaBancaria.Visible = True
'          Me.FrameHBeneficiario.Visible = False
'          Me.frameHOrganismos.Visible = False
          TDBFrameHConvenio.Visible = False
          TDBFrameHCaja.Visible = False
      Case "08"
          frameHAux00.Visible = False
          frameHCtaBancaria.Visible = False
'          Me.FrameHBeneficiario.Visible = False
'          Me.frameHOrganismos.Visible = True
'          Me.frameHOrganismos.Enabled = True
          TDBFrameHConvenio.Visible = False
          TDBFrameHCaja.Visible = False
      Case Else ' no se ha definido todavia
          frameHAux00.Visible = True
'          Me.frameHOrganismos.Visible = False
          frameHCtaBancaria.Visible = False
'          Me.FrameHBeneficiario.Visible = False
          TDBFrameHConvenio.Visible = False
          TDBFrameHCaja.Visible = False
'          hctalarga = ""
  End Select
End Sub

Public Sub DatosHaber(hauxiliar1 As String, hlarga As String)
'Select Case IIf(IsNull(rs_aux1!h_Aux1), "", rs_aux1!h_Aux1)
Select Case hauxiliar1
        Case "00"
'            Me.FrameHBeneficiario.Visible = False
'            Me.frameHCtaBancaria.Visible = False
'            Me.frameHAux00.Visible = True
'            Me.frameHOrganismos.Visible = False
            TDBFrameHCaja.Visible = False
        Case "01"
'            Me.frameHOrganismos.Visible = False
'            Me.FrameHBeneficiario.Visible = True
'            Me.frameHCtaBancaria.Visible = False
'            Me.frameHAux00.Visible = False
            TDBFrameHCaja.Visible = False
            Select Case CboTipo.Text
              Case "PCO"
'                Me.lblHBenefaux1.Visible = False
'                Me.lblHnomBenefaux1.Visible = False
                DtCHcodbenef.Visible = True
                DtCHDescripbenef.Visible = True
                DtCHcodbenef.Text = hlarga
                DtCHcodbenef_Click (1)
              Case Else
                DtCHcodbenef.Visible = False
                DtCHDescripbenef.Visible = False
'                Me.lblHBenefaux1.Visible = True
'                Me.lblHnomBenefaux1.Visible = True
'                Me.lblHBenefaux1 = hlarga
                Call buscabenef(hlarga)
'                hctalarga = Me.lblHBenefaux1
'                Me.lblHnomBenefaux1 = Trim(Cdenominacion)
            End Select
        '**buscar nombre beneficiario
        Case "02"
'            Me.frameHOrganismos.Visible = False
'            Me.FrameHBeneficiario.Visible = False
'            Me.frameHAux00.Visible = False
'            Me.frameHCtaBancaria.Visible = True
            TDBFrameHCaja.Visible = False
'            Me.cboHctaaux1 = hlarga
            Call buscactabancaria(hlarga)
'            Me.cboHctanomaux1 = cdenomctabancaria
'            hctalarga = Me.cboHctaaux1
        Case "08"
'            Me.FrameHBeneficiario.Visible = False
'            Me.frameHAux00.Visible = False
'            Me.frameHCtaBancaria.Visible = False
            frameHOrganismos.Visible = True
            TDBFrameHCaja.Visible = False
'            Me.cboHCodOrg = hlarga
            ''Call buscactabancaria(Trim(rs_aux1!H_Cta_Aux1))
            Call buscaorganismo(Trim(cboHCodOrg.Text))
'            hctalarga = Me.cboHCodOrg
'            Me.cboHDenomOrg = Me.denomorgan
        '***buscar nombre de la cuenta
        Case "10"
'            Me.FrameHBeneficiario.Visible = False
'            Me.frameHCtaBancaria.Visible = False
'            Me.frameHAux00.Visible = True
'            Me.frameHOrganismos.Visible = False
            TDBFrameHCaja.Visible = True
            DTCHidcaja.Text = hlarga
            hctalarga = hlarga
            'DtCHIdCaja_Click 0
            'buscacaja hlarga
            DTCHDesCaja.Text = DTCHidcaja.BoundText
        Case Else
'            Me.FrameHBeneficiario.Visible = False
'            Me.frameHCtaBancaria.Visible = False
'            Me.frameHAux00.Visible = True
'            Me.frameHOrganismos.Visible = False
            TDBFrameHCaja.Visible = False
'            hctalarga = ""
        End Select
End Sub

Public Sub DatosDebe(Daux As String, dcta As String)
  Select Case Daux
        Case "00"
'            Me.FrameDBeneficiario.Visible = False
'            Me.frameDCtaBancaria.Visible = False
'            Me.frameDOrganismos.Visible = False
'            Me.frameDaux00.Visible = True
'            Me.TDBFrameDCaja.Visible = False
'            dctalarga = ""
        Case "01"
'            Me.frameDOrganismos.Visible = False
'            Me.frameDCtaBancaria.Visible = False
'            Me.frameDaux00.Visible = False
'            Me.FrameDBeneficiario.Visible = True
'            Me.TDBFrameDCaja.Visible = False
            Select Case CboTipo.Text 'rs_aux1!tipo_comp
              Case "PCO"
                lblDBenefaux1.Visible = False
'                Me.lblDnomBenefaux1.Visible = False
                DtCDcodbenef.Visible = True
                DtCDDescripbenef.Visible = True
                DtCDcodbenef.Text = dcta
'                DtCDcodbenef_Click (1)
'                dctalarga = DtCDcodbenef.Text 'dcta
              Case "CAD"
                lblDBenefaux1.Visible = False
'                Me.lblDnomBenefaux1.Visible = False
                DtCDcodbenef.Visible = True
                DtCDDescripbenef.Visible = True
                DtCDcodbenef.Text = dcta
                'DtCDcodbenef_Click (1)
'                dctalarga = DtCDcodbenef.Text 'dcta
              Case Else
                lblDBenefaux1.Visible = True
'                Me.lblDnomBenefaux1.Visible = True
                DtCDcodbenef.Visible = False
                DtCDDescripbenef.Visible = False
'                Me.lblDBenefaux1 = dcta
                Call buscabenef(dcta)
'                Me.lblDnomBenefaux1 = Trim(Cdenominacion)
''                dctalarga = Me.lblDBenefaux1
            End Select
        Case "02"
'            Me.frameDOrganismos.Visible = False
'            Me.frameDaux00.Visible = False
'            Me.FrameDBeneficiario.Visible = False
'            Me.frameDCtaBancaria.Visible = True
'            Me.TDBFrameDCaja.Visible = False
'            Me.cboDctaaux1 = dcta
            Call buscactabancaria(dcta)
'            Me.cboDctanomaux1 = cdenomctabancaria
'            dctalarga = Me.cboDctaaux1
        Case "08"
'            Me.frameDaux00.Visible = False
'            Me.FrameDBeneficiario.Visible = False
'            Me.frameDCtaBancaria.Visible = True
            frameDOrganismos.Visible = True
'            Me.TDBFrameDCaja.Visible = False
'            Me.cboDCodOrg = dcta
            ''Call buscactabancaria(Trim(rs_aux1!H_Cta_Aux1))
            Call buscaorganismo(Trim(cboDCodOrg.Text))
'            Me.cboDDenomOrg = Me.denomorgan
'            dctalarga = Me.cboDCodOrg
        Case "10"
'            Me.FrameDBeneficiario.Visible = False
'            Me.frameDCtaBancaria.Visible = False
'            Me.frameDaux00.Visible = True
'            Me.frameDOrganismos.Visible = False
'            Me.TDBFrameDCaja.Visible = True
            dtcDIdCaja.Text = dcta
            DTCDDesCaja.Text = dtcDIdCaja.BoundText
'            dctalarga = dcta
            'buscacaja dcta
            'DTCDDesCaja.Text = Trim(Gdenomcaja)
            'DTCDDesCaja.Text = dtcDIdCaja.BoundText
            'DtCDIDCaja_Click 0
        Case Else
'            Me.FrameDBeneficiario.Visible = False
'            Me.frameDCtaBancaria.Visible = False
'            Me.frameDaux00.Visible = True
'            Me.frameDOrganismos.Visible = False
'            Me.TDBFrameDCaja.Visible = False
'            dctalarga = ""
        End Select
End Sub

'Public Sub activdatosdebe()
' Select Case daux1
'    Case "00"
''      dctalarga = ""
'    Case "01"
''      dctalarga = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'      cboDctaaux1.Text = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'    Case "02"
'      'If cboDctaaux1.Text <> "" Then
''        dctalarga = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'        cboDctaaux1.Text = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'      'Else
'        'MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        'Exit Sub
'    '  End If
'    Case "08"
'      'If cboDCodOrg.Text <> "" Then
''        dctalarga = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'        cboDCodOrg.Text = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'      'Else
'        MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        Exit Sub
'      'End If
'    Case "09"
''        dctalarga = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'        DtCDIdConvenio = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'        DtCDIdConvenio_Change
'    Case "03"
''        dctalarga = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'        dtcDIdCaja.Text = IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'        buscacaja IIf(IsNull(rs_aux1!D_Cta_Aux1), "", rs_aux1!D_Cta_Aux1)
'        DTCDDesCaja.Text = Trim(Gdenomcaja)
'        'DTCDDesCaja.BoundText = dtcDIdCaja.BoundText
'        'DtCDIDCaja_Click 0
'  End Select
''  Select Case daux2
'    Case "00"
''      dctaaux2 = ""
'    Case "01"
''      dctaaux2 = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'      lblDBenefaux1 = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'    Case "02"
'      'If cboDctaaux1.Text <> "" Then
''        dctaaux2 = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'        cboDctaaux1.Text = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'      'Else
'        'MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        'Exit Sub
'      'End If
'    Case "08"
'      'If cboDCodOrg.Text <> "" Then
''        dctaaux2 = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'        cboDCodOrg.Text = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'      'Else
'        'MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        'Exit Sub
'      'End If
'    Case "03"
''        dctaaux2 = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'        dtcDIdCaja.Text = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'        DtCDIDCaja_Click 0
'    Case "09"
''        dctaaux2 = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'        DtCDIdConvenio.Text = IIf(IsNull(rs_aux1!D_Cta_Aux2), "", rs_aux1!D_Cta_Aux2)
'        DtCDIdConvenio_Change
'  End Select
''  Select Case daux3
'    Case "00"
''      dctaaux3 = ""
'    Case "01"
''      dctaaux3 = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'      lblDBenefaux1 = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'    Case "02"
'      'If cboDctaaux1.Text <> "" Then
''        dctaaux3 = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'        cboDctaaux1.Text = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'      'Else
'        'MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'        'Exit Sub
'      'End If
'    Case "03"
''        dctaaux3 = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'        dtcDIdCaja.Text = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'        DtCDIDCaja_Click 0
'    Case "08"
'      'If cboDCodOrg.Text <> "" Then
''        dctaaux3 = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'        cboDCodOrg.Text = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'      'Else
'       ' MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'       ' Exit Sub
'      'End If
'     Case "09"
''        dctaaux3 = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'        DtCDIdConvenio.Text = IIf(IsNull(rs_aux1!d_CtaAux3), "", rs_aux1!d_CtaAux3)
'        DtCDIdConvenio_Change
'  End Select
'End Sub
'
'Public Sub activdatosHaber()
''Select Case haux1
'    Case "00"
''      hctalarga = ""
'    Case "01"
''      hctalarga = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'      lblHBenefaux1 = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'    Case "02"
'      'If cboHctaaux1.Text <> "" Then
''        hctalarga = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'        cboHctaaux1.Text = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'      'Else
'      '  MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'      '  Exit Sub
'      'End If
'    Case "03"
''        hctalarga = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'        DTCHidcaja.Text = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'        buscacaja IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'        DTCHDesCaja.Text = Gdenomcaja
'       'DTCHidcaja.Text = Str(hctalarga)
'        'DtCHIdCaja_Click 0
'    Case "08"
'      'If cboHCodOrg.Text <> "" Then
''        hctalarga = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'        cboHCodOrg.Text = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'      'Else
'       ' MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'       ' Exit Sub
'      'End If
'    Case "09"
''        hctalarga = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'        DtCHIdConvenio.Text = IIf(IsNull(rs_aux1!H_Cta_Aux1), "", rs_aux1!H_Cta_Aux1)
'        DtCHIdConvenio_Change
'  End Select
''  Select Case haux2
'    Case "00"
''      hctaaux2 = ""
'    Case "01"
''      hctaaux2 = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'      lblHBenefaux1 = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'    Case "02"
'      'If cboHctaaux1.Text <> "" Then
''        hctaaux2 = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'        cboHctaaux1.Text = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'      'Else
'      '  MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'      '  Exit Sub
'      'End If
'    Case "03"
''        hctaaux2 = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'        DTCHidcaja.Text = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'        buscacaja IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'        DTCHDesCaja.Text = Gdenomcaja
'        'DtCHIdCaja_Click 0
'    Case "08"
'      'If cboHCodOrg.Text <> "" Then
''        hctaaux2 = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'        cboHCodOrg.Text = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'     ' Else
'      '  MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'      '  Exit Sub
'      'End If
'     Case "09"
''           hctaaux2 = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
'           DtCHIdConvenio.Text = IIf(IsNull(rs_aux1!H_Cta_Aux2), "", rs_aux1!H_Cta_Aux2)
''           DtCHIdConvenio.Text = LTrim(RTrim(hctaaux2))
'           DtCHIdConvenio_Change
'  End Select
''  Select Case haux3
'    Case "00"
''      hctaaux3 = ""
'    Case "01"
''      hctaaux3 = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'      lblHBenefaux1 = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'    Case "02"
'      'If cboHctaaux1.Text <> "" Then
''        hctaaux3 = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'        cboHctaaux1.Text = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'      'Else
'       ' MsgBox "Seleccione una cuenta bancaria", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'       ' Exit Sub
'      'End If
'    Case "03"
''        hctaaux3 = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'        DTCHidcaja.Text = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'        buscacaja IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'        DTCHDesCaja.Text = Gdenomcaja
'        'DtCHIdCaja_Click 0
'    Case "08"
'      'If cboHCodOrg.Text <> "" Then
''        hctaaux3 = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'        cboHCodOrg.Text = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'      'Else
'       ' MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1, "Introducción de Datos"
'       ' Exit Sub
'      'End If
'    Case "09"
''           hctaaux3 = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'           DtCHIdConvenio.Text = IIf(IsNull(rs_aux1!H_Cta_Aux3), "", rs_aux1!H_Cta_Aux3)
'           DtCHIdConvenio_Change
'  End Select
'End Sub

Private Sub buscacaja(codcaja As String)
'Dim sqlbuscaja As String
'Dim rsbuscaja As ADODB.Recordset
'Set rsbuscaja = New ADODB.Recordset
'rsbuscaja.CursorLocation = adUseClient
'sqlbuscaja = "SELECT denominacion_caja From cc_Cajas" & _
'              " WHERE (codigo_caja = '" & codcaja & "')"
'rsbuscaja.Open sqlbuscaja, db, adOpenKeyset, adLockReadOnly
'If rsbuscaja.RecordCount <> 0 Then
'   Gdenomcaja = Trim(rsbuscaja!denominacion_caja)
'Else
'  Gdenomcaja = ""
'End If
End Sub

Private Sub txt_codigo1_Click(Area As Integer)
    dtc_desc14.BoundText = txt_codigo1.BoundText
End Sub
